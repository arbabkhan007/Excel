# Pet Sitting & Dog Walking Business Spreadsheet — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 27 tab `.xlsx` workbook with about 50,000 formulas: client and pet records, bookings, visit logs, pricing, invoices, payments, expenses, mileage, a dashboard, reports and automatic reminders. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_pet_sitting_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --business TEXT       business name on the cover and invoices
  --currency SYMBOL     currency symbol  (default: $)
  --unit {km,miles}     mileage unit
  --rows N              rows of capacity in the big registers
  --today YYYY-MM-DD    pretend it is this date (for the demo data)
```

---

## Contents

* [0. How it is put together](#0-how-it-is-put-together)
  * [The build pipeline](#the-build-pipeline)
  * [Describing a column](#describing-a-column)
  * [The rules every formula obeys](#the-rules-every-formula-obeys)
  * [The hidden `Calc` tab](#the-hidden-calc-tab)
  * [Every tab at a glance](#every-tab-at-a-glance)
  * [Every file at a glance](#every-file-at-a-glance)
* [1. Run it](#1-run-it)
  * [`build_pet_sitting_workbook.py`](#build_pet_sitting_workbookpy)
* [2. The package](#2-the-package)
  * [`pet_sitting_spreadsheet/__init__.py`](#pet_sitting_spreadsheet__init__py)
  * [`pet_sitting_spreadsheet/__main__.py`](#pet_sitting_spreadsheet__main__py)
  * [`pet_sitting_spreadsheet/cli.py`](#pet_sitting_spreadsheetclipy)
  * [`pet_sitting_spreadsheet/config.py`](#pet_sitting_spreadsheetconfigpy)
  * [`pet_sitting_spreadsheet/layout.py`](#pet_sitting_spreadsheetlayoutpy)
  * [`pet_sitting_spreadsheet/formats.py`](#pet_sitting_spreadsheetformatspy)
  * [`pet_sitting_spreadsheet/writer.py`](#pet_sitting_spreadsheetwriterpy)
  * [`pet_sitting_spreadsheet/names.py`](#pet_sitting_spreadsheetnamespy)
  * [`pet_sitting_spreadsheet/build.py`](#pet_sitting_spreadsheetbuildpy)
  * [`pet_sitting_spreadsheet/sample_data.py`](#pet_sitting_spreadsheetsample_datapy)
* [3. The registers](#3-the-registers)
  * [`pet_sitting_spreadsheet/sheets/__init__.py`](#pet_sitting_spreadsheetsheets__init__py)
  * [`pet_sitting_spreadsheet/sheets/common.py`](#pet_sitting_spreadsheetsheetscommonpy)
  * [`pet_sitting_spreadsheet/sheets/lists.py`](#pet_sitting_spreadsheetsheetslistspy)
  * [`pet_sitting_spreadsheet/sheets/settings.py`](#pet_sitting_spreadsheetsheetssettingspy)
  * [`pet_sitting_spreadsheet/sheets/services.py`](#pet_sitting_spreadsheetsheetsservicespy)
  * [`pet_sitting_spreadsheet/sheets/clients.py`](#pet_sitting_spreadsheetsheetsclientspy)
  * [`pet_sitting_spreadsheet/sheets/pets.py`](#pet_sitting_spreadsheetsheetspetspy)
  * [`pet_sitting_spreadsheet/sheets/bookings.py`](#pet_sitting_spreadsheetsheetsbookingspy)
  * [`pet_sitting_spreadsheet/sheets/visits.py`](#pet_sitting_spreadsheetsheetsvisitspy)
  * [`pet_sitting_spreadsheet/sheets/health.py`](#pet_sitting_spreadsheetsheetshealthpy)
  * [`pet_sitting_spreadsheet/sheets/incidents.py`](#pet_sitting_spreadsheetsheetsincidentspy)
  * [`pet_sitting_spreadsheet/sheets/messages.py`](#pet_sitting_spreadsheetsheetsmessagespy)
  * [`pet_sitting_spreadsheet/sheets/invoices.py`](#pet_sitting_spreadsheetsheetsinvoicespy)
  * [`pet_sitting_spreadsheet/sheets/payments.py`](#pet_sitting_spreadsheetsheetspaymentspy)
  * [`pet_sitting_spreadsheet/sheets/expenses.py`](#pet_sitting_spreadsheetsheetsexpensespy)
  * [`pet_sitting_spreadsheet/sheets/mileage.py`](#pet_sitting_spreadsheetsheetsmileagepy)
  * [`pet_sitting_spreadsheet/sheets/staff.py`](#pet_sitting_spreadsheetsheetsstaffpy)
  * [`pet_sitting_spreadsheet/sheets/inventory.py`](#pet_sitting_spreadsheetsheetsinventorypy)
  * [`pet_sitting_spreadsheet/sheets/reminders.py`](#pet_sitting_spreadsheetsheetsreminderspy)
* [4. The pages](#4-the-pages)
  * [`pet_sitting_spreadsheet/sheets/calc.py`](#pet_sitting_spreadsheetsheetscalcpy)
  * [`pet_sitting_spreadsheet/sheets/dashboard.py`](#pet_sitting_spreadsheetsheetsdashboardpy)
  * [`pet_sitting_spreadsheet/sheets/reports.py`](#pet_sitting_spreadsheetsheetsreportspy)
  * [`pet_sitting_spreadsheet/sheets/budget.py`](#pet_sitting_spreadsheetsheetsbudgetpy)
  * [`pet_sitting_spreadsheet/sheets/start_here.py`](#pet_sitting_spreadsheetsheetsstart_herepy)
  * [`pet_sitting_spreadsheet/sheets/printables.py`](#pet_sitting_spreadsheetsheetsprintablespy)
  * [`pet_sitting_spreadsheet/sheets/help_tab.py`](#pet_sitting_spreadsheetsheetshelp_tabpy)
* [5. Quality checks (optional)](#5-quality-checks-optional)
  * [`tools/validate_workbook.py`](#toolsvalidate_workbookpy)
  * [`tools/check_sample_data.py`](#toolscheck_sample_datapy)
  * [`tools/recalc_check.py`](#toolsrecalc_checkpy)
  * [`tools/export_source_md.py`](#toolsexport_source_mdpy)
  * [`tools/make_download_zip.py`](#toolsmake_download_zippy)
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

`build_pet_sitting_workbook.py` parses the command line into a `Config`, then
`pet_sitting_spreadsheet.build.build_workbook(path, cfg)` does six things in order:

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
6. **Hide the `Calc` engine**, activate the first tab and close the file.

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

### The hidden `Calc` tab

Dashboards do not aggregate the registers directly. A hidden sheet
pre-computes one row per month, per platform/service, per category and
per owner, each with its own ranking key, and the visible pages read
those columns with plain `INDEX`/`MATCH`. It keeps the heavy `SUMIFS`
work in one place, keeps recalculation predictable, and makes every
dashboard number traceable to a single cell.


### Every tab at a glance

27 tabs in the order they appear in the finished workbook. "Rows" is the capacity of a register at the default `--rows` setting.

| # | Tab | Kind | Written by | Columns | Rows |
|---|-----|------|------------|--------:|-----:|
| 1 | **Start Here** | page | `pet_sitting_spreadsheet/build.py` | - | - |
| 2 | **Dashboard** | page | `pet_sitting_spreadsheet/sheets/dashboard.py` | - | - |
| 3 | **Settings** | page | `pet_sitting_spreadsheet/sheets/settings.py` | - | - |
| 4 | **Lists** | page | `pet_sitting_spreadsheet/sheets/lists.py` | - | - |
| 5 | **Services** | register | `pet_sitting_spreadsheet/sheets/services.py` | 20 | 50 |
| 6 | **Clients** | register | `pet_sitting_spreadsheet/sheets/clients.py` | 38 | 250 |
| 7 | **Pets** | register | `pet_sitting_spreadsheet/sheets/pets.py` | 46 | 250 |
| 8 | **Bookings** | register | `pet_sitting_spreadsheet/sheets/bookings.py` | 52 | 400 |
| 9 | **Visit Log** | register | `pet_sitting_spreadsheet/sheets/visits.py` | 36 | 400 |
| 10 | **Health Log** | register | `pet_sitting_spreadsheet/sheets/health.py` | 22 | 200 |
| 11 | **Incidents** | register | `pet_sitting_spreadsheet/sheets/incidents.py` | 23 | 100 |
| 12 | **Messages** | register | `pet_sitting_spreadsheet/sheets/messages.py` | 20 | 200 |
| 13 | **Invoices** | register | `pet_sitting_spreadsheet/sheets/invoices.py` | 30 | 300 |
| 14 | **Payments** | register | `pet_sitting_spreadsheet/sheets/payments.py` | 20 | 300 |
| 15 | **Expenses** | register | `pet_sitting_spreadsheet/sheets/expenses.py` | 21 | 300 |
| 16 | **Mileage** | register | `pet_sitting_spreadsheet/sheets/mileage.py` | 24 | 300 |
| 17 | **Staff** | register | `pet_sitting_spreadsheet/sheets/staff.py` | 28 | 40 |
| 18 | **Inventory** | register | `pet_sitting_spreadsheet/sheets/inventory.py` | 19 | 120 |
| 19 | **Reminders** | register | `pet_sitting_spreadsheet/sheets/reminders.py` | 10 | 40 |
| 20 | **Reports** | page | `pet_sitting_spreadsheet/sheets/reports.py` | - | - |
| 21 | **Budget** | page | `pet_sitting_spreadsheet/sheets/budget.py` | - | - |
| 22 | **Print Invoice** | page | `pet_sitting_spreadsheet/sheets/printables.py` | - | - |
| 23 | **Print Pet Care** | page | `pet_sitting_spreadsheet/sheets/printables.py` | - | - |
| 24 | **Print Visit** | page | `pet_sitting_spreadsheet/sheets/printables.py` | - | - |
| 25 | **Print Forms** | page | `pet_sitting_spreadsheet/sheets/printables.py` | - | - |
| 26 | **Help** | page | `pet_sitting_spreadsheet/sheets/help_tab.py` | - | - |
| 27 | **Calc** | register | `pet_sitting_spreadsheet/sheets/calc.py` | 44 | 250 |

### Every file at a glance

| File | Lines | What it is |
|------|------:|------------|
| [`build_pet_sitting_workbook.py`](#build_pet_sitting_workbookpy) | 32 | Build the Pet Sitting & Dog Walking Business Spreadsheet |
| [`pet_sitting_spreadsheet/__init__.py`](#pet_sitting_spreadsheet__init__py) | 29 | Pet Sitting & Dog Walking Business Spreadsheet ============================================== Generates a... |
| [`pet_sitting_spreadsheet/__main__.py`](#pet_sitting_spreadsheet__main__py) | 6 | ``python -m pet_sitting_spreadsheet`` |
| [`pet_sitting_spreadsheet/cli.py`](#pet_sitting_spreadsheetclipy) | 111 | Command line front end |
| [`pet_sitting_spreadsheet/config.py`](#pet_sitting_spreadsheetconfigpy) | 210 | Configuration for the Pet Sitting & Dog Walking Business Spreadsheet generator |
| [`pet_sitting_spreadsheet/layout.py`](#pet_sitting_spreadsheetlayoutpy) | 243 | Layout engine |
| [`pet_sitting_spreadsheet/formats.py`](#pet_sitting_spreadsheetformatspy) | 190 | Cell format factory |
| [`pet_sitting_spreadsheet/writer.py`](#pet_sitting_spreadsheetwriterpy) | 444 | Sheet writing helpers shared by every tab of the workbook |
| [`pet_sitting_spreadsheet/names.py`](#pet_sitting_spreadsheetnamespy) | 40 | Workbook level defined names that point at the register sheets |
| [`pet_sitting_spreadsheet/build.py`](#pet_sitting_spreadsheetbuildpy) | 148 | Workbook builder: puts every tab together in the right order |
| [`pet_sitting_spreadsheet/sample_data.py`](#pet_sitting_spreadsheetsample_datapy) | 903 | Demo data |
| [`pet_sitting_spreadsheet/sheets/__init__.py`](#pet_sitting_spreadsheetsheets__init__py) | 14 | One module per tab of the workbook |
| [`pet_sitting_spreadsheet/sheets/common.py`](#pet_sitting_spreadsheetsheetscommonpy) | 110 | Small building blocks shared by the register sheets |
| [`pet_sitting_spreadsheet/sheets/lists.py`](#pet_sitting_spreadsheetsheetslistspy) | 190 | The `Lists` tab: every drop down list in the workbook lives here |
| [`pet_sitting_spreadsheet/sheets/settings.py`](#pet_sitting_spreadsheetsheetssettingspy) | 274 | The `Settings` tab |
| [`pet_sitting_spreadsheet/sheets/services.py`](#pet_sitting_spreadsheetsheetsservicespy) | 92 | `Services` - the service catalogue and price list that drives booking costs |
| [`pet_sitting_spreadsheet/sheets/clients.py`](#pet_sitting_spreadsheetsheetsclientspy) | 152 | `Clients` - the client database with lifetime value and follow-up flags |
| [`pet_sitting_spreadsheet/sheets/pets.py`](#pet_sitting_spreadsheetsheetspetspy) | 136 | `Pets` - full pet profiles: care routine, medical, vaccinations, warnings |
| [`pet_sitting_spreadsheet/sheets/bookings.py`](#pet_sitting_spreadsheetsheetsbookingspy) | 225 | `Bookings` - the scheduling engine: prices, surcharges, conflicts, alerts |
| [`pet_sitting_spreadsheet/sheets/visits.py`](#pet_sitting_spreadsheetsheetsvisitspy) | 134 | `Visit Log` - the daily pet care log, one row per visit or walk |
| [`pet_sitting_spreadsheet/sheets/health.py`](#pet_sitting_spreadsheetsheetshealthpy) | 85 | `Health Log` - medication administration, vaccinations and vet events |
| [`pet_sitting_spreadsheet/sheets/incidents.py`](#pet_sitting_spreadsheetsheetsincidentspy) | 137 | `Incidents` - safety, emergency contacts and the incident report register |
| [`pet_sitting_spreadsheet/sheets/messages.py`](#pet_sitting_spreadsheetsheetsmessagespy) | 85 | `Messages` - client communication, feedback, complaints and follow-ups |
| [`pet_sitting_spreadsheet/sheets/invoices.py`](#pet_sitting_spreadsheetsheetsinvoicespy) | 158 | `Invoices` - invoice register with automatic numbering, tax and ageing |
| [`pet_sitting_spreadsheet/sheets/payments.py`](#pet_sitting_spreadsheetsheetspaymentspy) | 89 | `Payments` - money received, deposits, tips, refunds and reconciliation |
| [`pet_sitting_spreadsheet/sheets/expenses.py`](#pet_sitting_spreadsheetsheetsexpensespy) | 87 | `Expenses` - every business cost, with tax, deductibility and approvals |
| [`pet_sitting_spreadsheet/sheets/mileage.py`](#pet_sitting_spreadsheetsheetsmileagepy) | 96 | `Mileage` - trip log with odometer readings and reimbursement maths |
| [`pet_sitting_spreadsheet/sheets/staff.py`](#pet_sitting_spreadsheetsheetsstaffpy) | 120 | `Staff` - team directory, workload, timesheet totals and pay |
| [`pet_sitting_spreadsheet/sheets/inventory.py`](#pet_sitting_spreadsheetsheetsinventorypy) | 99 | `Inventory` - pet supplies, stock levels, reorder alerts and value |
| [`pet_sitting_spreadsheet/sheets/reminders.py`](#pet_sitting_spreadsheetsheetsreminderspy) | 255 | `Reminders` - the alert centre plus a custom task list |
| [`pet_sitting_spreadsheet/sheets/calc.py`](#pet_sitting_spreadsheetsheetscalcpy) | 213 | `Calc` - hidden helper sheet |
| [`pet_sitting_spreadsheet/sheets/dashboard.py`](#pet_sitting_spreadsheetsheetsdashboardpy) | 375 | `Dashboard` - the one page view of the business |
| [`pet_sitting_spreadsheet/sheets/reports.py`](#pet_sitting_spreadsheetsheetsreportspy) | 380 | `Reports` - printable business analytics for any date range |
| [`pet_sitting_spreadsheet/sheets/budget.py`](#pet_sitting_spreadsheetsheetsbudgetpy) | 272 | `Budget` - targets, budget vs actual, break-even, forecast and cash flow |
| [`pet_sitting_spreadsheet/sheets/start_here.py`](#pet_sitting_spreadsheetsheetsstart_herepy) | 219 | `Start Here` - welcome page, navigation menu and setup checklist |
| [`pet_sitting_spreadsheet/sheets/printables.py`](#pet_sitting_spreadsheetsheetsprintablespy) | 751 | The four print-ready tabs: invoice, pet care sheet, visit report, forms |
| [`pet_sitting_spreadsheet/sheets/help_tab.py`](#pet_sitting_spreadsheetsheetshelp_tabpy) | 295 | `Help` - instructions, formula notes, troubleshooting and version log |
| [`tools/validate_workbook.py`](#toolsvalidate_workbookpy) | 287 | Static checks on a generated workbook |
| [`tools/check_sample_data.py`](#toolscheck_sample_datapy) | 171 | Cross-check the example records in a generated workbook |
| [`tools/recalc_check.py`](#toolsrecalc_checkpy) | 216 | Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error... |
| [`tools/export_source_md.py`](#toolsexport_source_mdpy) | 922 | Write every source file of a generator into one markdown document |
| [`tools/make_download_zip.py`](#toolsmake_download_zippy) | 232 | Pack a generator into a single zip people can download and run |

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_pet_sitting_workbook.py
```

The workbook lands in `output/Pet Sitting Business Spreadsheet.xlsx`.

### `build_pet_sitting_workbook.py`

*32 lines*

> Build the Pet Sitting & Dog Walking Business Spreadsheet.
>
> Double click it, or run::
>
>     python build_pet_sitting_workbook.py
>
> The finished .xlsx lands in the ``output`` folder next to this script.
> Author of the generated workbook: Novality Store.
>
> Options (all optional)::
>
>     python build_pet_sitting_workbook.py --blank          # no example data
>     python build_pet_sitting_workbook.py --currency "£" --currency-code GBP
>     python build_pet_sitting_workbook.py --unit miles
>     python build_pet_sitting_workbook.py -o "~/Desktop/Pet Sitting.xlsx"
>
> See ``python build_pet_sitting_workbook.py --help`` for the full list.

````python
#!/usr/bin/env python3
"""
Build the Pet Sitting & Dog Walking Business Spreadsheet.

Double click it, or run::

    python build_pet_sitting_workbook.py

The finished .xlsx lands in the ``output`` folder next to this script.
Author of the generated workbook: Novality Store.

Options (all optional)::

    python build_pet_sitting_workbook.py --blank          # no example data
    python build_pet_sitting_workbook.py --currency "£" --currency-code GBP
    python build_pet_sitting_workbook.py --unit miles
    python build_pet_sitting_workbook.py -o "~/Desktop/Pet Sitting.xlsx"

See ``python build_pet_sitting_workbook.py --help`` for the full list.
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pet_sitting_spreadsheet.cli import main   # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer and the demo data. These are the parts every tab leans on.

### `pet_sitting_spreadsheet/__init__.py`

*29 lines*

> Pet Sitting & Dog Walking Business Spreadsheet
> ==============================================
>
> Generates a ready-to-sell Excel workbook for pet sitters and dog walkers:
> clients, pets, bookings, visit logs, health records, invoices, payments,
> expenses, mileage, staff, inventory, a dashboard, reports, a budget, a
> reminder centre and printable forms.
>
>     python -m pet_sitting_spreadsheet --output "My Workbook.xlsx"
>
> or from Python::
>
>     from pet_sitting_spreadsheet import build_workbook, Config
>     build_workbook("workbook.xlsx", Config(include_sample_data=False))
>
> Every formula cell is locked and every sheet is protected with the password
> in :class:`~pet_sitting_spreadsheet.config.Config` (default ``premium``), so
> the maths cannot be edited by accident while the input cells stay open.

````python
"""
Pet Sitting & Dog Walking Business Spreadsheet
==============================================

Generates a ready-to-sell Excel workbook for pet sitters and dog walkers:
clients, pets, bookings, visit logs, health records, invoices, payments,
expenses, mileage, staff, inventory, a dashboard, reports, a budget, a
reminder centre and printable forms.

    python -m pet_sitting_spreadsheet --output "My Workbook.xlsx"

or from Python::

    from pet_sitting_spreadsheet import build_workbook, Config
    build_workbook("workbook.xlsx", Config(include_sample_data=False))

Every formula cell is locked and every sheet is protected with the password
in :class:`~pet_sitting_spreadsheet.config.Config` (default ``premium``), so
the maths cannot be edited by accident while the input cells stay open.
"""

from .build import build_workbook
from .config import AUTHOR, FORMULA_PASSWORD, Config

__all__ = ["build_workbook", "Config", "AUTHOR", "FORMULA_PASSWORD",
           "__version__"]

__version__ = "1.0.0"
__author__ = AUTHOR
````

### `pet_sitting_spreadsheet/__main__.py`

*6 lines*

> ``python -m pet_sitting_spreadsheet``.

````python
"""``python -m pet_sitting_spreadsheet``."""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
````

### `pet_sitting_spreadsheet/cli.py`

*111 lines · 3 functions*

> Command line front end.
>
> python -m pet_sitting_spreadsheet
> python -m pet_sitting_spreadsheet -o "output/Pet Sitting.xlsx" --blank
> python -m pet_sitting_spreadsheet --currency "£" --unit miles --rows 500

| Defines | What it does |
|---------|--------------|
| `build_parser()` |  |
| `config_from_args(args)` |  |
| `main(argv)` |  |

**Module constants:** `DEFAULT_OUTPUT` (os.path.join(...)).

````python
"""Command line front end.

    python -m pet_sitting_spreadsheet
    python -m pet_sitting_spreadsheet -o "output/Pet Sitting.xlsx" --blank
    python -m pet_sitting_spreadsheet --currency "£" --unit miles --rows 500
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import build_workbook
from .config import AUTHOR, PRODUCT_LISTING_TITLE, PRODUCT_NAME, Config

DEFAULT_OUTPUT = os.path.join(
    "output", "Pet Sitting Business Spreadsheet.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pet_sitting_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty "
                             "(no example records)")
    parser.add_argument("--no-protect", action="store_true",
                        help="do not protect the sheets (for development)")
    parser.add_argument("--password", default=None,
                        help="sheet protection password")
    parser.add_argument("--author", default=AUTHOR,
                        help="document author / creator")
    parser.add_argument("--business", default=None,
                        help="pre-fill the business name on Settings")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--unit", choices=["km", "miles"], default=None,
                        help="distance unit for mileage")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (default row counts are "
                             "multiplied to fit this many booking rows)")
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
    if args.unit:
        kwargs["distance_unit"] = args.unit
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.business:
        cfg.business.name = args.business
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

    print(f"{cfg.product_name}")
    print(f"  listing title : {PRODUCT_LISTING_TITLE}")
    print(f"  author        : {cfg.author}")
    print("  sheets        : 27 (1 hidden calculation engine)")
    print(f"  example data  : {'no' if not cfg.include_sample_data else 'yes'}")
    print(f"  protection    : "
          f"{'off' if not cfg.protect_sheets else 'on, password ' + cfg.password}")
    print(f"  written to    : {os.path.abspath(path)}  ({size:,.0f} KB)")
    return 0


if __name__ == "__main__":        # pragma: no cover
    raise SystemExit(main())
````

### `pet_sitting_spreadsheet/config.py`

*210 lines · 3 classes*

> Configuration for the Pet Sitting & Dog Walking Business Spreadsheet generator.
>
> Everything that a shop owner may want to re-brand or re-price lives here, so the
> workbook can be regenerated for a different store, currency or country without
> touching any of the sheet code.
>
> Author of the generated workbooks: Novality Store.

| Defines | What it does |
|---------|--------------|
| `class Palette` |  |
| `class BusinessDefaults` | Sample values written into the Settings sheet |
| `class Config` | All generator options |

**Module constants:** `AUTHOR` ('Novality Store'), `PRODUCT_NAME` ('Pet Sitting & Dog Walking Business Spreadsheet'), `PRODUCT_SUBTITLE` ('Client Booking, Pet Care Log, Invoices, Payments, Mileag...), `PRODUCT_VERSION` ('1.0.0'), `PRODUCT_LISTING_TITLE` ('Pet Sitting Business Spreadsheet, Dog Walker Client Book...), `FORMULA_PASSWORD` ('premium'), `DEFAULT_ROWS` (21 entries).

````python
"""
Configuration for the Pet Sitting & Dog Walking Business Spreadsheet generator.

Everything that a shop owner may want to re-brand or re-price lives here, so the
workbook can be regenerated for a different store, currency or country without
touching any of the sheet code.

Author of the generated workbooks: Novality Store.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import date


# ---------------------------------------------------------------------------
# Product / branding constants
# ---------------------------------------------------------------------------

AUTHOR = "Novality Store"
PRODUCT_NAME = "Pet Sitting & Dog Walking Business Spreadsheet"
PRODUCT_SUBTITLE = (
    "Client Booking, Pet Care Log, Invoices, Payments, Mileage & Profit Tracker"
)
PRODUCT_VERSION = "1.0.0"

#: Title suggested for the shop listing.
PRODUCT_LISTING_TITLE = (
    "Pet Sitting Business Spreadsheet, Dog Walker Client Booking, "
    "Payments, Mileage and Profit Tracker"
)

#: Password that protects every sheet so formula cells cannot be edited by
#: accident.  Input cells stay unlocked, formula cells are locked.
FORMULA_PASSWORD = "premium"


# ---------------------------------------------------------------------------
# Colour palette (professional teal / amber scheme used across every tab)
# ---------------------------------------------------------------------------

class Palette:
    ink = "#17313B"          # darkest navy-teal - titles
    primary = "#1F4E5F"      # headers
    primary_mid = "#2E7D8F"  # section bands
    primary_soft = "#DCEAEE"  # light band / alternating rows
    accent = "#E2A33C"       # amber accent
    accent_soft = "#FDF3E0"
    input_bg = "#FFFDF4"     # editable cell background (cream)
    input_border = "#E4D8B4"
    calc_bg = "#EEF3F6"      # locked formula cell background (cool grey)
    calc_border = "#D2DCE2"
    white = "#FFFFFF"
    grey_text = "#5A6B72"
    line = "#BFCCD2"
    good_bg = "#D7F2E3"
    good_text = "#116241"
    warn_bg = "#FFF0CC"
    warn_text = "#8A5A00"
    bad_bg = "#FBD9DD"
    bad_text = "#9C1C2E"
    info_bg = "#DDEBF7"
    info_text = "#1F4E79"
    note_bg = "#F7F9FA"


# ---------------------------------------------------------------------------
# Row capacity of every register (how many pre-formatted / pre-formulated rows)
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Settings": 0,
    "Lists": 40,
    "Services": 50,
    "Clients": 250,
    "Pets": 250,
    "Bookings": 400,
    "Visits": 400,
    "Visit Log": 400,
    "Health": 200,
    "Health Log": 200,
    "Incidents": 100,
    "Messages": 200,
    "Invoices": 300,
    "Payments": 300,
    "Expenses": 300,
    "Mileage": 300,
    "Staff": 40,
    "Inventory": 120,
    "Holidays": 30,
    "Tasks": 40,
    "Drawings": 24,
}


@dataclass
class BusinessDefaults:
    """Sample values written into the Settings sheet."""

    name: str = "Happy Paws Pet Care"
    owner: str = "Alex Morgan"
    phone: str = "+1 (555) 014 2200"
    email: str = "hello@happypawspetcare.com"
    website: str = "www.happypawspetcare.com"
    address: str = "28 Brookfield Lane"
    city: str = "Springfield"
    state: str = "IL"
    postcode: str = "62704"
    country: str = "United States"
    tax_id: str = "TAX-884 221 905"
    insurance_provider: str = "PetGuard Business Insurance"
    insurance_policy: str = "PG-7741920"
    licence_name: str = "City Pet Care Operator Licence"
    bank_details: str = "Happy Paws Pet Care - Acct 1023 4455 - Routing 021000021"


@dataclass
class Config:
    """All generator options."""

    output_path: str = "output/Pet Sitting and Dog Walking Business Spreadsheet.xlsx"

    # Branding / document properties
    author: str = AUTHOR
    company: str = AUTHOR
    product_name: str = PRODUCT_NAME
    product_subtitle: str = PRODUCT_SUBTITLE
    version: str = PRODUCT_VERSION
    release_date: str = "2026-10-11"

    # Protection
    password: str = FORMULA_PASSWORD
    protect_sheets: bool = True

    # Formats
    currency_symbol: str = "$"
    currency_code: str = "USD"
    date_format: str = "dd-mmm-yyyy"
    time_format: str = "h:mm AM/PM"
    distance_unit: str = "km"

    # Business / pricing defaults pre-filled on the Settings tab
    business: BusinessDefaults = field(default_factory=BusinessDefaults)
    tax_name: str = "Sales Tax"
    tax_registered: str = "Yes"
    tax_rate: float = 0.08
    prices_include_tax: str = "No"
    payment_terms_days: int = 7
    deposit_percent: float = 0.25
    late_fee_percent: float = 0.05
    invoice_prefix: str = "INV-"
    invoice_start_number: int = 1001
    client_prefix: str = "CL-"
    pet_prefix: str = "PET-"
    booking_prefix: str = "BK-"
    mileage_rate: float = 0.67
    travel_fee_per_unit: float = 0.80
    free_travel_distance: float = 8.0
    holiday_surcharge: float = 0.25
    weekend_surcharge: float = 0.10
    emergency_fee: float = 15.0
    last_minute_days: int = 2
    extra_pet_fee: float = 7.5
    loyalty_discount: float = 0.05
    loyalty_threshold: int = 20
    staff_cost_rate: float = 18.0
    tax_reserve_percent: float = 0.25
    reminder_lead_days: int = 3
    vaccination_alert_days: int = 30
    inactive_client_days: int = 60
    follow_up_days: int = 5
    expiry_alert_days: int = 30
    default_visit_minutes: int = 30
    travel_time_minutes: int = 15
    day_start: str = "07:00"
    day_end: str = "19:00"

    # Content switches
    include_sample_data: bool = True
    rows: dict = field(default_factory=lambda: dict(DEFAULT_ROWS))
    today: date = field(default_factory=date.today)

    # ---------------------------------------------------------------- helpers
    def row_count(self, sheet: str) -> int:
        return int(self.rows.get(sheet, 100))

    def scale_rows(self, booking_rows: int) -> None:
        """Scale every register in place so Bookings holds `booking_rows`."""
        base = DEFAULT_ROWS["Bookings"]
        factor = max(0.05, float(booking_rows) / float(base))
        self.rows = {k: (max(5, int(round(v * factor))) if v else v)
                     for k, v in self.rows.items()}

    def scaled(self, factor: float) -> "Config":
        """Return a copy with every register scaled (used by --rows-scale)."""
        rows = {k: max(5, int(round(v * factor))) if v else v
                for k, v in self.rows.items()}
        return replace(self, rows=rows)

    # Number format strings derived from the chosen currency ------------------
    @property
    def money_format(self) -> str:
        s = self.currency_symbol
        return f'"{s}"#,##0.00;[Red]-"{s}"#,##0.00'

    @property
    def money_format_0(self) -> str:
        s = self.currency_symbol
        return f'"{s}"#,##0;[Red]-"{s}"#,##0'
````

### `pet_sitting_spreadsheet/layout.py`

*243 lines · 3 classes · 10 functions*

> Layout engine.
>
> The workbook is described declaratively: every register sheet (Clients, Pets,
> Bookings, ...) is a :class:`SheetSpec` made of :class:`Col` objects.  Formulas
> are written as templates that never contain hard coded column letters, so
> columns can be reordered or inserted without breaking a single formula:
>
>     "=IF({client}{r}=\"\",\"\",SUMIFS([[Payments.net]],[[Payments.client]],{client}{r}))"
>
>     {r}            -> the current Excel row number (1 based)
>     {key}          -> column letter of the column `key` on the *same* sheet
>     [[Sheet.key]]  -> absolute range of that column's data rows, e.g.
>                       Payments!$K$5:$K$304
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

The workbook is described declaratively: every register sheet (Clients, Pets,
Bookings, ...) is a :class:`SheetSpec` made of :class:`Col` objects.  Formulas
are written as templates that never contain hard coded column letters, so
columns can be reordered or inserted without breaking a single formula:

    "=IF({client}{r}=\\"\\",\\"\\",SUMIFS([[Payments.net]],[[Payments.client]],{client}{r}))"

    {r}            -> the current Excel row number (1 based)
    {key}          -> column letter of the column `key` on the *same* sheet
    [[Sheet.key]]  -> absolute range of that column's data rows, e.g.
                      Payments!$K$5:$K$304
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

### `pet_sitting_spreadsheet/formats.py`

*190 lines · 1 class*

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
        self._styles: Dict[str, Dict[str, Any]] = {
            "text": {"align": "left"},
            "id": {"align": "left", "bold": True, "font_color": P.primary},
            "wrap": {"align": "left", "valign": "top", "text_wrap": True},
            "date": {"num_format": config.date_format, "align": "center"},
            "time": {"num_format": config.time_format, "align": "center"},
            "money": {"num_format": money, "align": "right"},
            "money0": {"num_format": money0, "align": "right"},
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

### `pet_sitting_spreadsheet/writer.py`

*444 lines · 1 class · 7 functions*

> Sheet writing helpers shared by every tab of the workbook.
>
> :class:`SheetWriter` wraps an XlsxWriter worksheet and adds the house style:
> title band, navigation strip, section headers, KPI tiles, label/value settings
> rows and - most importantly - :meth:`SheetWriter.grid`, which turns a
> :class:`~pet_sitting_spreadsheet.layout.SheetSpec` into a finished register
> with headers, drop downs, conditional formats, sample rows, frozen panes and
> locked formula cells.

| Defines | What it does |
|---------|--------------|
| `class SheetWriter` | Convenience wrapper around a worksheet |
| `yes_no(validate_note)` |  |
| `list_from(name)` | Drop down fed by a defined name (see lists.py) |
| `date_check()` |  |
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
:class:`~pet_sitting_spreadsheet.layout.SheetSpec` into a finished register
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
    ("Clients", "Clients"),
    ("Pets", "Pets"),
    ("Bookings", "Bookings"),
    ("Visit Log", "Visit Log"),
    ("Invoices", "Invoices"),
    ("Reminders", "Reminders"),
    ("Help", "Help"),
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
            f"&L&\"Calibri,Bold\"&12{cfg.business.name if False else ''}"
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

### `pet_sitting_spreadsheet/names.py`

*40 lines · 1 function*

> Workbook level defined names that point at the register sheets.
>
> The drop down lists on the `Lists` tab define their own names; these are the
> ones that have to follow live data (client names, pet references, invoice
> numbers, ...), so they are built from the registry once every register has
> been created.  Each one is an ``OFFSET`` range that grows and shrinks with
> the number of filled rows.

| Defines | What it does |
|---------|--------------|
| `define_cross_sheet_names(book, reg)` | Create the live lookup names |

**Module constants:** `CROSS_SHEET_NAMES` (10 entries).

````python
"""Workbook level defined names that point at the register sheets.

The drop down lists on the `Lists` tab define their own names; these are the
ones that have to follow live data (client names, pet references, invoice
numbers, ...), so they are built from the registry once every register has
been created.  Each one is an ``OFFSET`` range that grows and shrinks with
the number of filled rows.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    "ClientNames": ("Clients", "name"),
    "ClientIDs": ("Clients", "id"),
    "PetLabels": ("Pets", "label"),
    "PetNames": ("Pets", "name"),
    "ServiceNames": ("Services", "name"),
    "StaffNames": ("Staff", "name"),
    "BookingIDs": ("Bookings", "id"),
    "InvoiceNos": ("Invoices", "invoice_no"),
    "VisitIDs": ("Visit Log", "id"),
    "IncidentIDs": ("Incidents", "id"),
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

### `pet_sitting_spreadsheet/build.py`

*148 lines · 4 functions*

> Workbook builder: puts every tab together in the right order.
>
> Typical use::
>
>     from pet_sitting_spreadsheet import build_workbook
>     build_workbook("Pet Sitting Business Spreadsheet.xlsx")

| Defines | What it does |
|---------|--------------|
| `_start_here(book, fmts, cfg, reg)` |  |
| `build_register(book, fmts, cfg, reg, module)` | Create one register sheet from its :class:`SheetSpec` |
| `build_workbook(path, cfg)` | Write the whole product to *path* and return the path |
| `_set_properties(book, cfg)` |  |

**Module constants:** `REGISTER_MODULES` (16 items), `SHEET_ORDER` (27 items), `PAGE_BUILDERS` (11 entries), `REGISTER_BY_NAME` ({module.NAME: module for module in REGISTER_MODULES}).

````python
"""Workbook builder: puts every tab together in the right order.

Typical use::

    from pet_sitting_spreadsheet import build_workbook
    build_workbook("Pet Sitting Business Spreadsheet.xlsx")
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
from .sheets import (bookings, budget, calc, clients, dashboard, expenses,
                     health, help_tab, incidents, inventory, invoices, lists,
                     messages, mileage, payments, pets, printables, reminders,
                     reports, services, settings, staff, visits)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [services, clients, pets, bookings, visits, health,
                    incidents, messages, invoices, payments, expenses,
                    mileage, staff, inventory, reminders, calc]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Start Here", "Dashboard", "Settings", "Lists", "Services", "Clients",
    "Pets", "Bookings", "Visit Log", "Health Log", "Incidents", "Messages",
    "Invoices", "Payments", "Expenses", "Mileage", "Staff", "Inventory",
    "Reminders", "Reports", "Budget", "Print Invoice", "Print Pet Care",
    "Print Visit", "Print Forms", "Help", "Calc",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Start Here": lambda b, f, c, r: _start_here(b, f, c, r),
    "Dashboard": dashboard.build,
    "Settings": settings.build,
    "Lists": lists.build,
    "Reports": reports.build,
    "Budget": budget.build,
    "Print Invoice": printables.build_invoice,
    "Print Pet Care": printables.build_care_sheet,
    "Print Visit": printables.build_visit_report,
    "Print Forms": printables.build_forms,
    "Help": help_tab.build,
}

#: register modules by sheet name
REGISTER_BY_NAME = {module.NAME: module for module in REGISTER_MODULES}


def _start_here(book, fmts, cfg, reg):
    from .sheets import start_here
    return start_here.build(book, fmts, cfg, reg)


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


def build_workbook(path: str = "Pet Sitting Business Spreadsheet.xlsx",
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

    # 1. Register every spec first so cross sheet formulas can be resolved
    #    no matter which tab is written first.
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

    # 3. Live lookup names (client names, pet references, invoice numbers).
    define_cross_sheet_names(book, reg)

    # 4. Final touches: hide the engine, open on the welcome page.
    if calc.NAME in sheets:
        sheets[calc.NAME].ws.hide()
    first = sheets[SHEET_ORDER[0]].ws
    first.activate()
    first.set_first_sheet()

    book.close()
    return path


def _set_properties(book, cfg: Config) -> None:
    book.set_properties({
        "title": cfg.product_name,
        "subject": "Pet sitting and dog walking business management",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("pet sitting, dog walking, pet care, client booking, "
                     "invoices, payments, mileage, expenses, profit, "
                     "spreadsheet template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked and "
                     f"every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

### `pet_sitting_spreadsheet/sample_data.py`

*903 lines · 1 class · 4 functions*

> Demo data.
>
> Sample clients, pets, bookings, invoices, payments, expenses and trips are
> written into the input (cream) cells so the dashboard, reports and reminders
> show meaningful numbers the moment the workbook is opened.  Everything is
> dated relative to the build date and can be switched off with
> ``--no-sample-data``.

| Defines | What it does |
|---------|--------------|
| `d(base, offset)` |  |
| `t(hour, minute)` |  |
| `class Sample` | fields: clients, pets, services, staff, bookings, visits, invoices, payments |
| `get(cfg)` | Build (once) the demo dataset for this configuration |
| `_build(cfg)` |  |

````python
"""
Demo data.

Sample clients, pets, bookings, invoices, payments, expenses and trips are
written into the input (cream) cells so the dashboard, reports and reminders
show meaningful numbers the moment the workbook is opened.  Everything is
dated relative to the build date and can be switched off with
``--no-sample-data``.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .config import Config


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def t(hour: int, minute: int = 0) -> _dt.time:
    return _dt.time(hour, minute)


@dataclass
class Sample:
    clients: List[Dict[str, Any]] = field(default_factory=list)
    pets: List[Dict[str, Any]] = field(default_factory=list)
    services: List[Dict[str, Any]] = field(default_factory=list)
    staff: List[Dict[str, Any]] = field(default_factory=list)
    bookings: List[Dict[str, Any]] = field(default_factory=list)
    visits: List[Dict[str, Any]] = field(default_factory=list)
    invoices: List[Dict[str, Any]] = field(default_factory=list)
    payments: List[Dict[str, Any]] = field(default_factory=list)
    expenses: List[Dict[str, Any]] = field(default_factory=list)
    mileage: List[Dict[str, Any]] = field(default_factory=list)
    health: List[Dict[str, Any]] = field(default_factory=list)
    incidents: List[Dict[str, Any]] = field(default_factory=list)
    messages: List[Dict[str, Any]] = field(default_factory=list)
    inventory: List[Dict[str, Any]] = field(default_factory=list)
    tasks: List[Dict[str, Any]] = field(default_factory=list)
    drawings: List[Dict[str, Any]] = field(default_factory=list)


_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (once) the demo dataset for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg)
    return _CACHE[key]


# ---------------------------------------------------------------------------

def _build(cfg: Config) -> Sample:
    rnd = random.Random(20261011)
    today = cfg.today
    s = Sample()

    # ------------------------------------------------------------- services
    s.services = [
        dict(name="Dog Walk - 30 min", category="Dog walking",
             description="Solo neighbourhood walk, water top-up and paw wipe",
             unit="Per walk", duration=30, rate=25, extra_pet=10,
             staff_cost=11, active="Yes", travel="No",
             notes="Most popular weekday service"),
        dict(name="Dog Walk - 60 min", category="Dog walking",
             description="Long walk or park session with recall practice",
             unit="Per walk", duration=60, rate=40, extra_pet=15,
             staff_cost=20, active="Yes", travel="No"),
        dict(name="Puppy Visit - 30 min", category="Puppy care",
             description="Toilet break, short play, crate reset and feed",
             unit="Per visit", duration=30, rate=28, extra_pet=8,
             staff_cost=12, active="Yes", travel="No"),
        dict(name="Drop-in Visit - 30 min", category="Drop-in visit",
             description="Feed, fresh water, litter or potty break, playtime",
             unit="Per visit", duration=30, rate=26, extra_pet=8,
             staff_cost=11, active="Yes", travel="No"),
        dict(name="Cat Feeding Visit", category="Feeding visit",
             description="Feed, water, litter tray clean, cuddles, plant care",
             unit="Per visit", duration=20, rate=22, extra_pet=7,
             staff_cost=9, active="Yes", travel="No"),
        dict(name="Pet Sitting - Half Day", category="Pet sitting",
             description="Four hours of care in the client's home",
             unit="Per hour", duration=240, rate=18, extra_pet=6,
             staff_cost=9, active="Yes", travel="Yes"),
        dict(name="Overnight Pet Sitting", category="Overnight stay",
             description="Evening arrival to morning departure, in-home stay",
             unit="Per night", duration=840, rate=95, extra_pet=20,
             staff_cost=45, active="Yes", travel="Yes"),
        dict(name="House Sitting - Per Night", category="House sitting",
             description="Overnight house and pet care, mail and bins",
             unit="Per night", duration=840, rate=85, extra_pet=18,
             staff_cost=40, active="Yes", travel="Yes"),
        dict(name="Senior Pet Care Visit", category="Senior pet care",
             description="Gentle care, mobility help, medication, comfort",
             unit="Per visit", duration=45, rate=34, extra_pet=10,
             staff_cost=15, active="Yes", travel="No"),
        dict(name="Medication Administration", category="Medication visit",
             description="Scheduled medication with written administration log",
             unit="Per visit", duration=20, rate=18, extra_pet=6,
             staff_cost=8, active="Yes", travel="No"),
        dict(name="Pet Taxi / Vet Transport", category="Pet transport",
             description="Door to door transport including waiting time",
             unit="Per hour", duration=60, rate=38, extra_pet=10,
             staff_cost=17, active="Yes", travel="Yes"),
        dict(name="Extra Pet Add-on", category="Add-on",
             description="Additional pet on any visit",
             unit="Flat fee", duration=0, rate=8, extra_pet=0,
             staff_cost=3, active="Yes", travel="No"),
    ]

    # ---------------------------------------------------------------- staff
    s.staff = [
        dict(name="Alex Morgan", phone="+1 (555) 014 2200",
             email="alex@happypawspetcare.com", role="Owner",
             employment="Employee - full time", start_date=d(today, -980),
             status="Active", availability="Mon-Fri 07:00-18:00",
             days_off="Sunday", pay_type="Salary", pay_rate=26,
             certs="Pet First Aid, Canine Behaviour L3",
             cert_expiry=d(today, 160), check_date=d(today, -220),
             emergency="Jordan Morgan +1 (555) 014 2201",
             leave="None booked", notes="Handles reactive dogs"),
        dict(name="Jamie Lee", phone="+1 (555) 330 8812",
             email="jamie@happypawspetcare.com", role="Dog walker",
             employment="Employee - part time", start_date=d(today, -420),
             status="Active", availability="Mon-Sat 08:00-16:00",
             days_off="Sunday", pay_type="Hourly", pay_rate=19,
             certs="Pet First Aid", cert_expiry=d(today, 24),
             check_date=d(today, -180),
             emergency="Pat Lee +1 (555) 330 8813",
             leave="2 days in December", notes="Great with puppies"),
        dict(name="Priya Shah", phone="+1 (555) 772 6690",
             email="priya@happypawspetcare.com", role="Pet sitter",
             employment="Casual", start_date=d(today, -260),
             status="Active", availability="Tue-Sun 09:00-20:00",
             days_off="Monday", pay_type="Per visit", pay_rate=16,
             certs="Pet First Aid, Cat Care", cert_expiry=d(today, 310),
             check_date=d(today, -150),
             emergency="Ravi Shah +1 (555) 772 6691",
             leave="None booked", notes="Overnight specialist"),
        dict(name="Tom Reed", phone="+1 (555) 448 1120",
             email="tom.reed@example.com", role="Contractor",
             employment="Contractor", start_date=d(today, -120),
             status="Active", availability="Weekends only",
             days_off="Mon-Fri", pay_type="Commission %", pay_rate=55,
             certs="Pet First Aid", cert_expiry=d(today, -8),
             check_date=d(today, -110),
             emergency="Sam Reed +1 (555) 448 1121",
             leave="None booked",
             notes="Invoices monthly - certificate expired, chase renewal"),
    ]

    # -------------------------------------------------------------- clients
    client_rows = [
        ("Sarah Mitchell", "+1 (555) 201 7744", "sarah.mitchell@example.com",
         "14 Juniper Way", "Springfield", "62704", "Tom Mitchell",
         "+1 (555) 201 7745", "SMS", "Google search", -240, "Active",
         "Key held", "KEY-01", "Lockbox by side gate, code 4471",
         "Signed", "Yes", 500,
         "Bella pulls left on the lead. Treats in blue tin.",
         "Prefers evening walks. Pays by card on the 1st."),
        ("James Okafor", "+1 (555) 889 2031", "james.okafor@example.com",
         "7 Ashgrove Court", "Springfield", "62702", "Ada Okafor",
         "+1 (555) 889 2032", "Email", "Client referral", -190, "Active",
         "Key held", "KEY-02", "Front door key, alarm code 8812 then #",
         "Signed", "Yes", 400,
         "Luna hides under the bed - give her 10 minutes.",
         "Travels for work most months."),
        ("Priya Raman", "+1 (555) 663 4417", "priya.raman@example.com",
         "92 Hilltop Drive", "Riverton", "62711", "Vijay Raman",
         "+1 (555) 663 4418", "WhatsApp", "Instagram", -150, "Active",
         "Lockbox code", "KEY-03", "Lockbox on tap, code 2290",
         "Signed", "Yes", 600,
         "Cooper is in training - no treats before walks.",
         "Weekly Tuesday and Thursday walks."),
        ("Daniel Brooks", "+1 (555) 712 9034", "daniel.brooks@example.com",
         "3 Meadow Lane", "Springfield", "62703", "Erin Brooks",
         "+1 (555) 712 9035", "Phone call", "Vet clinic", -120, "Active",
         "Key held", "KEY-04", "Back door, lift handle then turn",
         "Signed", "Yes", 750,
         "Daisy needs 60 minute runs. Teddy stays home.",
         "Referred by Riverton Vet Clinic."),
        ("Megan Foster", "+1 (555) 204 5588", "megan.foster@example.com",
         "18 Willow Street", "Riverton", "62712", "Chris Foster",
         "+1 (555) 204 5589", "SMS", "Facebook", -95, "Active",
         "Smart lock code", "-", "Smart lock code 3391, auto locks",
         "Signed", "Yes", 350,
         "Two cats, feed separately. Milo is on medication.",
         "Books two weeks of holiday cover twice a year."),
        ("Luis Hernandez", "+1 (555) 556 1177", "luis.hernandez@example.com",
         "45 Oakfield Road", "Springfield", "62705", "Maria Hernandez",
         "+1 (555) 556 1178", "Phone call", "Word of mouth", -70, "Active",
         "Not held", "-", "Owner home most days - ring the bell",
         "Signed", "Yes", 300,
         "Ruby is nervous with men in hats.",
         "Pays cash at the end of each week."),
        ("Chloe Bennett", "+1 (555) 991 3320", "chloe.bennett@example.com",
         "11 Garden Terrace", "Riverton", "62713", "Paul Bennett",
         "+1 (555) 991 3321", "Email", "Flyer / poster", -45, "Active",
         "Key held", "KEY-05", "Side gate latch, key under planter",
         "Sent", "Yes", 250,
         "Poppy the rabbit needs fresh greens daily.",
         "New client - send the 6 week check-in."),
        ("Robert Nguyen", "+1 (555) 338 7742", "robert.nguyen@example.com",
         "60 Station Road", "Springfield", "62701", "Linh Nguyen",
         "+1 (555) 338 7743", "SMS", "Google search", -400, "Inactive",
         "Key returned", "-", "Key returned 10 months ago",
         "Expired", "No", 0,
         "Charlie loves the river path.",
         "Moved to a new suburb - win-back candidate."),
        ("Amelia Clarke", "+1 (555) 447 2218", "amelia.clarke@example.com",
         "2 Bridge View", "Riverton", "62714", "Noah Clarke",
         "+1 (555) 447 2219", "Email", "Instagram", -12, "Prospective",
         "Not held", "-", "Meet and greet booked",
         "Not sent", "No", 0,
         "Two rescue dogs - needs a slow introduction.",
         "Quote sent for weekday walks."),
    ]
    for (name, phone, email, addr, city, post, ename, ephone, contact,
         referral, reg_offset, status, key_status, key_tag, access,
         contract, consent, limit, prefs, notes) in client_rows:
        s.clients.append(dict(
            name=name, phone=phone, email=email, address=addr, city=city,
            postcode=post, emergency_name=ename, emergency_phone=ephone,
            contact_method=contact, referral=referral,
            reg_date=d(today, reg_offset), status=status,
            key_status=key_status, key_tag=key_tag, access_notes=access,
            access_code="See access notes",
            contract=contract, agreement_date=d(today, reg_offset + 2),
            vet_consent=consent, vet_limit=limit, preferences=prefs,
            notes=notes))

    # ----------------------------------------------------------------- pets
    pet_rows = [
        ("Bella", "Sarah Mitchell", "Dog", "Labrador Retriever", -1460,
         "Female (spayed)", 28.5, "Golden, white chest patch", "Friendly",
         "Pulls on the lead when excited", "None known",
         "Mild hip stiffness in cold weather", "Riverton Vet Clinic",
         "+1 (555) 100 2000", 40, "Pet Secure", "PS-55012",
         "2 cups dry food am and pm", "Dry kibble - salmon", "2 cups",
         "07:30 and 18:00", "Joint supplement", "1 chew with breakfast",
         "Daily with breakfast", "60 min walks, loves fetch",
         "Garden every 4 hours", "Brush twice a week",
         "Tennis ball, peanut butter kong", "None",
         "Harness on before the front door opens",
         "Do not let off lead near the main road"),
        ("Max", "Sarah Mitchell", "Dog", "Beagle", -980, "Male (neutered)",
         14.2, "Tricolour", "Energetic", "Food obsessed - counter surfs",
         "Chicken", "None", "Riverton Vet Clinic", "+1 (555) 100 2000", 70,
         "Pet Secure", "PS-55013", "1.5 cups twice daily",
         "Dry kibble - lamb", "1.5 cups", "07:30 and 18:00", "None", "-",
         "-", "30 min sniffy walks", "Garden every 4 hours",
         "Monthly nail trim", "Squeaky duck", "Escape artist",
         "Double check the side gate latch",
         "Will bolt through open gates"),
        ("Luna", "James Okafor", "Cat", "Ragdoll", -1100, "Female (spayed)",
         4.8, "Cream and grey point", "Shy", "Hides from strangers",
         "Dust mites", "None", "Springfield Animal Hospital",
         "+1 (555) 100 3000", 150, "CatCare Plus", "CC-9902",
         "Wet food am, dry topped up pm", "Wet - tuna pate", "1 pouch",
         "08:00 and 17:30", "None", "-", "-", "Indoor only - wand toy play",
         "Litter tray in laundry - scoop daily", "Brush every visit",
         "Feather wand, sunny windowsill", "None",
         "Let her come to you, never lift her", "Indoor cat - keep doors shut"),
        ("Cooper", "Priya Raman", "Dog", "Cockapoo", -620, "Male (neutered)",
         9.4, "Apricot curly coat", "Playful", "Jumps up when greeting",
         "Grain", "None", "Hilltop Veterinary", "+1 (555) 100 4000", 25,
         "Pet Secure", "PS-55120", "1 cup twice daily", "Grain free kibble",
         "1 cup", "07:00 and 18:30", "None", "-", "-",
         "45 min walks, training recall", "Garden before and after walks",
         "Groomer every 6 weeks", "Rope tug, training treats", "None",
         "Lead on in the hallway, treats for calm", "Training in progress"),
        ("Daisy", "Daniel Brooks", "Dog", "Border Collie", -1750, "Female",
         19.0, "Black and white", "Energetic", "Herds joggers",
         "None known", "None", "Riverton Vet Clinic", "+1 (555) 100 2000",
         210, "Pet Secure", "PS-55211", "2 cups am, 1 cup pm",
         "Dry kibble - chicken", "2 cups", "06:30 and 17:30", "None", "-",
         "-", "60 min off-lead run at Hill Park", "Garden every 3 hours",
         "Brush weekly", "Frisbee, flirt pole", "None",
         "Long line until recall is solid", "Keep away from cyclists"),
        ("Milo", "Megan Foster", "Cat", "Domestic Short Hair", -2200,
         "Male (neutered)", 5.6, "Black with white bib", "Calm",
         "Dislikes being picked up", "None known",
         "Hyperthyroid - daily medication", "Springfield Animal Hospital",
         "+1 (555) 100 3000", -5, "CatCare Plus", "CC-7781",
         "Half pouch three times daily", "Wet - chicken", "Half pouch",
         "07:00, 13:00, 19:00", "Felimazole 2.5mg",
         "1 tablet in food, morning only", "Daily 07:00",
         "Indoor with catio access", "Two trays - scoop every visit",
         "Brush twice weekly", "Catnip mouse", "Medical risk",
         "Medication must be given before food",
         "Vomiting or refusing food = call owner and vet"),
        ("Ruby", "Luis Hernandez", "Dog", "Staffordshire Bull Terrier",
         -1330, "Female (spayed)", 16.8, "Brindle", "Needs slow introduction",
         "Nervous of men in hats", "Beef", "None", "Oakfield Vets",
         "+1 (555) 100 5000", 95, "None", "-", "1.5 cups twice daily",
         "Dry kibble - fish", "1.5 cups", "08:00 and 18:00", "None", "-",
         "-", "30 min quiet street walks", "Garden every 4 hours",
         "Wipe down after walks", "Stuffed kong", "Not dog friendly",
         "Keep 10m from other dogs, treat scatter to pass",
         "Reactive to other dogs on lead"),
        ("Oscar", "Chloe Bennett", "Dog", "Dachshund", -700,
         "Male (neutered)", 7.2, "Smooth black and tan", "Friendly",
         "Barks at the postman", "None known", "Mild back sensitivity",
         "Hilltop Veterinary", "+1 (555) 100 4000", 300, "Pet Secure",
         "PS-55330", "0.75 cup twice daily", "Dry kibble - turkey",
         "0.75 cup", "08:00 and 18:00", "None", "-", "-",
         "20 min flat walks - no stairs", "Garden every 4 hours",
         "Ear clean weekly", "Soft bear, chew stick", "None",
         "Lift with two hands supporting the back",
         "Never let him jump off furniture"),
        ("Poppy", "Chloe Bennett", "Rabbit", "Mini Lop", -500, "Female",
         1.9, "White with grey ears", "Shy", "Startles easily",
         "None known", "None", "Hilltop Veterinary", "+1 (555) 100 4000",
         20, "None", "-", "Unlimited hay, fresh greens daily",
         "Hay and pellets", "Handful of greens", "Morning", "None", "-",
         "-", "Run time in the pen for 30 minutes",
         "Litter corner - clean daily", "Check nails monthly",
         "Parsley, cardboard tunnel", "None",
         "Scoop from underneath, support the back legs",
         "Keep away from the dogs next door"),
        ("Charlie", "Robert Nguyen", "Dog", "Golden Retriever", -2500,
         "Male (neutered)", 32.0, "Golden", "Calm", "None",
         "Pollen - seasonal", "Arthritis", "Station Road Vets",
         "+1 (555) 100 6000", -120, "None", "-", "3 cups twice daily",
         "Senior kibble", "3 cups", "07:00 and 18:00", "Metacam",
         "0.8ml with food", "Daily morning", "Gentle 30 min river walks",
         "Garden every 3 hours", "Brush every visit", "Old tennis ball",
         "None", "Help him into the car with the ramp",
         "Senior dog - watch for limping"),
        ("Simba", "Megan Foster", "Cat", "Maine Coon", -900, "Male",
         7.9, "Ginger tabby", "Playful", "Door dasher", "None known",
         "None", "Springfield Animal Hospital", "+1 (555) 100 3000", 35,
         "CatCare Plus", "CC-7782", "1 pouch twice daily", "Wet - salmon",
         "1 pouch", "07:00 and 19:00", "None", "-", "-",
         "Indoor play, laser pointer", "Two trays - scoop every visit",
         "Daily brush - knots easily", "Laser pointer", "Door dasher",
         "Shut the hallway door before opening the front door",
         "Will dash out of the front door"),
        ("Teddy", "Daniel Brooks", "Dog", "Cavoodle", -300,
         "Male (neutered)", 8.1, "Chocolate curly", "Friendly",
         "Separation anxiety - whines at first", "None known", "None",
         "Riverton Vet Clinic", "+1 (555) 100 2000", 180, "Pet Secure",
         "PS-55212", "1 cup twice daily", "Puppy kibble", "1 cup",
         "07:30 and 18:00", "None", "-", "-", "30 min walk plus garden play",
         "Garden every 3 hours", "Groomer every 8 weeks",
         "Rope toy, frozen kong", "None",
         "Leave the radio on when you go", "Anxious if left alone too long"),
    ]
    for (name, owner, ptype, breed, dob_off, gender, weight, colour,
         temperament, behaviour, allergies, medical, vet, vet_phone,
         vacc_off, insurer, policy, feeding, food, portion, schedule,
         medication, dosage, med_schedule, walking, potty, grooming,
         favourites, risk, handling, warnings) in pet_rows:
        s.pets.append(dict(
            name=name, owner=owner, type=ptype, breed=breed,
            dob=d(today, dob_off), gender=gender, weight=weight,
            colour=colour, photo="Link to photo folder",
            microchip=f"9560000{rnd.randint(100000, 999999)}",
            temperament=temperament, behaviour=behaviour,
            allergies=allergies, medical=medical, vet_name=vet,
            vet_phone=vet_phone, emerg_clinic="24/7 Springfield Pet ER",
            emerg_phone="+1 (555) 100 9111",
            vacc_date=d(today, vacc_off - 365),
            vacc_expiry=d(today, vacc_off), insurer=insurer, policy=policy,
            insurance_expiry=d(today, vacc_off + 40) if insurer != "None"
            else None,
            feeding=feeding, food=food, portion=portion,
            feed_schedule=schedule, medication=medication, dosage=dosage,
            med_schedule=med_schedule, walking=walking, potty=potty,
            grooming=grooming, favourites=favourites, risk=risk,
            handling=handling, warnings=warnings))

    pet_label = {p["name"]: f'{p["name"]} ({p["owner"]})' for p in s.pets}

    # ------------------------------------------------------------- bookings
    plan = [
        # (day offset, client, pet, service, start hour, staff, status)
        (-28, "Sarah Mitchell", "Bella", "Dog Walk - 60 min", 10, "Jamie Lee"),
        (-27, "Priya Raman", "Cooper", "Dog Walk - 30 min", 9, "Jamie Lee"),
        (-26, "Megan Foster", "Milo", "Cat Feeding Visit", 8, "Priya Shah"),
        (-24, "Daniel Brooks", "Daisy", "Dog Walk - 60 min", 7, "Alex Morgan"),
        (-23, "James Okafor", "Luna", "Drop-in Visit - 30 min", 17,
         "Priya Shah"),
        (-21, "Sarah Mitchell", "Max", "Dog Walk - 30 min", 11, "Jamie Lee"),
        (-20, "Luis Hernandez", "Ruby", "Dog Walk - 30 min", 15, "Alex Morgan"),
        (-18, "Chloe Bennett", "Oscar", "Puppy Visit - 30 min", 12,
         "Priya Shah"),
        (-17, "Megan Foster", "Simba", "Cat Feeding Visit", 8, "Priya Shah"),
        (-15, "Daniel Brooks", "Teddy", "Pet Sitting - Half Day", 9,
         "Alex Morgan"),
        (-14, "Priya Raman", "Cooper", "Dog Walk - 30 min", 9, "Jamie Lee"),
        (-12, "James Okafor", "Luna", "Overnight Pet Sitting", 19,
         "Priya Shah"),
        (-11, "Sarah Mitchell", "Bella", "Dog Walk - 60 min", 10, "Jamie Lee"),
        (-9, "Megan Foster", "Milo", "Medication Administration", 7,
         "Priya Shah"),
        (-8, "Chloe Bennett", "Poppy", "Drop-in Visit - 30 min", 16,
         "Priya Shah"),
        (-7, "Daniel Brooks", "Daisy", "Dog Walk - 60 min", 7, "Alex Morgan"),
        (-6, "Luis Hernandez", "Ruby", "Dog Walk - 30 min", 15, "Tom Reed"),
        (-5, "Priya Raman", "Cooper", "Dog Walk - 30 min", 9, "Jamie Lee"),
        (-4, "Sarah Mitchell", "Max", "Dog Walk - 30 min", 11, "Jamie Lee"),
        (-3, "Megan Foster", "Simba", "Cat Feeding Visit", 8, "Priya Shah"),
        (-2, "James Okafor", "Luna", "Drop-in Visit - 30 min", 17,
         "Priya Shah"),
        (-1, "Daniel Brooks", "Teddy", "Senior Pet Care Visit", 13,
         "Alex Morgan"),
        (0, "Sarah Mitchell", "Bella", "Dog Walk - 60 min", 10, "Jamie Lee"),
        (0, "Priya Raman", "Cooper", "Dog Walk - 30 min", 9, "Jamie Lee"),
        (1, "Chloe Bennett", "Oscar", "Puppy Visit - 30 min", 12,
         "Priya Shah"),
        (2, "Megan Foster", "Milo", "Cat Feeding Visit", 8, "Priya Shah"),
        (3, "Luis Hernandez", "Ruby", "Dog Walk - 30 min", 15, "Tom Reed"),
        (4, "Daniel Brooks", "Daisy", "Dog Walk - 60 min", 7, "Alex Morgan"),
        (6, "James Okafor", "Luna", "Overnight Pet Sitting", 19,
         "Priya Shah"),
        (8, "Sarah Mitchell", "Bella", "Pet Taxi / Vet Transport", 14,
         "Alex Morgan"),
        (10, "Priya Raman", "Cooper", "Dog Walk - 30 min", 9, "Jamie Lee"),
        (12, "Megan Foster", "Simba", "House Sitting - Per Night", 18,
         "Priya Shah"),
    ]
    svc_by_name = {s_["name"]: s_ for s_ in s.services}
    for i, (offset, client, pet, service, hour, staff) in enumerate(plan):
        svc = svc_by_name[service]
        minutes = svc["duration"] or 30
        start = t(hour, 0)
        end_minutes = (hour * 60 + minutes) % (24 * 60)
        end = t(end_minutes // 60, end_minutes % 60)
        if offset < -1:
            status = "Completed"
        elif offset <= 0:
            status = "In progress" if offset == 0 else "Completed"
        elif offset <= 4:
            status = "Confirmed"
        else:
            status = "Pending"
        if i in (7, 19):
            status = "Cancelled"
        booking = dict(
            request_date=d(today, offset - rnd.randint(2, 10)),
            client=client, pet=pet_label[pet], pet_count=1, service=service,
            date=d(today, offset), start=start, end=end, qty=1,
            distance=round(rnd.uniform(2, 14), 1), staff=staff,
            status=status,
            confirm_sent="Yes" if status not in ("Pending", "Enquiry")
            else "No",
            reminder_sent="Yes" if offset < 2 else "No",
            recurring="Weekly" if service.startswith("Dog Walk") else "None",
            discount_type="None",
            notes="")
        if status == "Cancelled":
            booking["cancel_reason"] = rnd.choice(
                ["Client away", "Pet unwell", "Weather"])
            booking["notes"] = "Cancelled by client - offered a reschedule"
        if service == "Overnight Pet Sitting":
            booking["qty"] = 2
            booking["notes"] = "Two nights - keys collected in advance"
        if service == "House Sitting - Per Night":
            booking["qty"] = 4
            booking["notes"] = "Holiday cover - bins out Tuesday"
        if service == "Pet Sitting - Half Day":
            booking["qty"] = 4
            booking["notes"] = "Four hours of in-home care"
        if client == "Sarah Mitchell" and pet == "Bella":
            booking["extra_pets"] = "Max joins the walk"
            booking["pet_count"] = 2
        s.bookings.append(booking)

    # --------------------------------------------------------------- visits
    walk_services = ("Dog Walk - 30 min", "Dog Walk - 60 min")
    visit_index = 0
    for idx, bk in enumerate(s.bookings):
        if bk["status"] not in ("Completed", "In progress"):
            continue
        visit_index += 1
        booking_id = f'BK-{idx + 1:04d}'
        start_hour = bk["start"].hour
        arrive = t(start_hour, rnd.choice([0, 2, 5, 8]))
        dur = svc_by_name[bk["service"]]["duration"] or 30
        depart_minutes = (arrive.hour * 60 + arrive.minute + dur) % (24 * 60)
        depart = t(depart_minutes // 60, depart_minutes % 60)
        is_walk = bk["service"] in walk_services
        visit = dict(
            booking=booking_id, date=bk["date"], staff=bk["staff"],
            arrive=arrive, depart=depart,
            fed="Yes" if not is_walk else "N/A",
            water="Yes",
            meds="Yes" if bk["service"] == "Medication Administration"
            else "N/A",
            potty=rnd.choice(["Both", "Pee only", "Litter cleaned"]),
            play="15 min fetch" if is_walk else "10 min wand toy play",
            grooming="Paws wiped" if is_walk else "Brushed",
            mood=rnd.choice(["Happy", "Calm", "Playful"]),
            appetite=rnd.choice(["Normal", "Normal", "Reduced"]),
            health="All normal", incident="No",
            photos="Photo sent to owner", update_sent="Yes",
            checklist="Yes", status="Completed",
            notes="Visit completed as planned")
        if is_walk:
            visit["walk_start"] = arrive
            walk_end_minutes = (arrive.hour * 60 + arrive.minute + dur - 5)
            visit["walk_end"] = t((walk_end_minutes // 60) % 24,
                                  walk_end_minutes % 60)
            visit["walk_distance"] = round(dur / 20 + rnd.uniform(0.2, 1.4), 1)
        if visit_index == 6:
            visit.update(status="Missed", notes="Client was home - no access "
                                                "needed, visit cancelled on "
                                                "arrival",
                         update_sent="Yes", mood="", appetite="")
        if visit_index == 11:
            visit.update(incident="Yes", incident_ref="IN-002",
                         health="Small cut on left front paw - cleaned",
                         notes="Owner notified immediately, vet not required")
        s.visits.append(visit)

    # ------------------------------------------------------------- invoices
    inv_plan = [
        ("Sarah Mitchell", -26, "Dog walking - 4 visits", 4, 25, "Paid"),
        ("Priya Raman", -25, "Dog walking - weekly package", 4, 25, "Paid"),
        ("Megan Foster", -24, "Cat feeding visits", 6, 22, "Paid"),
        ("Daniel Brooks", -22, "Dog walking - 60 min x 3", 3, 40, "Paid"),
        ("James Okafor", -20, "Drop-in visits x 5", 5, 26, "Paid"),
        ("Luis Hernandez", -18, "Dog walking - 30 min x 4", 4, 25, "Paid"),
        ("Chloe Bennett", -15, "Puppy visits x 3", 3, 28, "Partially paid"),
        ("Megan Foster", -12, "Cat feeding - fortnight cover", 10, 22, "Paid"),
        ("Daniel Brooks", -10, "Half day pet sitting", 4, 18, "Paid"),
        ("James Okafor", -9, "Overnight pet sitting x 2 nights", 2, 95,
         "Sent"),
        ("Sarah Mitchell", -6, "Dog walking - 60 min x 2", 2, 40, "Sent"),
        ("Megan Foster", -4, "Medication visits x 3", 3, 18, "Sent"),
        ("Priya Raman", -2, "Dog walking - 30 min x 2", 2, 25, "Draft"),
        ("Chloe Bennett", -1, "Drop-in visit - rabbit care", 1, 26, "Draft"),
        ("Luis Hernandez", -21, "Dog walking - 30 min x 6", 6, 25, "Paid"),
        ("Daniel Brooks", -16, "Senior pet care visits x 4", 4, 32, "Paid"),
        ("Sarah Mitchell", -13, "Dog walking - 60 min x 4", 4, 40, "Paid"),
        ("Chloe Bennett", -10, "Puppy visits x 5", 5, 28, "Paid"),
        ("James Okafor", -7, "Drop-in visits x 4", 4, 26, "Paid"),
        ("Megan Foster", -5, "Cat feeding - week cover", 7, 22, "Sent"),
    ]
    # A few invoices are raised straight from a booking instead of being
    # typed by hand - that is what the Booking ref column is for.  The
    # client and the wording come from the booking itself, so the two
    # records can never drift apart.  (booking number, days ago, status)
    linked_plan = [(19, -3, "Sent"), (21, -2, "Sent"), (22, -1, "Draft")]
    for i, (client, offset, desc, qty, rate, status) in enumerate(inv_plan):
        s.invoices.append(dict(
            date=d(today, offset), client=client, booking="",
            description=desc, qty=qty, rate=rate, status=status,
            fees=0, discount=0,
            deposit=0,
            sent_date=d(today, offset) if status != "Draft" else None,
            reminder_sent="Yes" if status == "Sent" and offset < -7 else "No",
            notes="Thank you for your business" if status == "Paid" else ""))
    for number, offset, status in linked_plan:
        booked = s.bookings[number - 1]
        ref = f"BK-{number:04d}"
        pet_name = booked["pet"].split(" (")[0]
        # No Qty / Rate: the amount comes from the linked booking.
        s.invoices.append(dict(
            date=d(today, offset), client=booked["client"], booking=ref,
            description=f'{booked["service"]} for {pet_name}', status=status,
            fees=0, discount=0, deposit=0,
            sent_date=d(today, offset) if status != "Draft" else None,
            reminder_sent="No",
            notes=f"Raised from booking {ref} - the amount comes straight "
                  f"from the booking"))

    # ------------------------------------------------------------- payments
    # Payments settle the invoice *including* tax, so the invoice shows as
    # Paid rather than part paid.  `share` < 1 leaves a balance on purpose.
    def _gross(index: int) -> float:
        qty, rate = inv_plan[index][3], inv_plan[index][4]
        gross = qty * rate
        if (cfg.tax_registered == "Yes"
                and cfg.prices_include_tax != "Yes"):
            gross *= 1 + cfg.tax_rate
        return round(gross, 2)

    #        invoice, day, method, share paid, tip
    pay_plan = [
        (0, -24, "Card - online", 1.0, 5.0),
        (1, -23, "Bank transfer", 1.0, 0.0),
        (2, -21, "Card - in person", 1.0, 0.0),
        (3, -19, "Bank transfer", 1.0, 10.0),
        (4, -17, "Digital wallet", 1.0, 0.0),
        (5, -14, "Cash", 1.0, 0.0),
        (6, -11, "Card - online", 0.45, 0.0),
        (7, -8, "Bank transfer", 1.0, 0.0),
        (8, -6, "Card - online", 1.0, 8.0),
        (14, -20, "Bank transfer", 1.0, 0.0),
        (15, -15, "Card - online", 1.0, 12.0),
        (16, -12, "Digital wallet", 1.0, 0.0),
        (17, -9, "Card - in person", 1.0, 0.0),
        (18, -6, "Bank transfer", 1.0, 5.0),
        (9, -3, "Card - online", 1.0, 15.0),
    ]
    for inv_index, offset, method, share, tip in pay_plan:
        client = inv_plan[inv_index][0]
        amount = round(_gross(inv_index) * share, 2)
        s.payments.append(dict(
            date=d(today, offset), client=client,
            invoice_no=f"INV-{1001 + inv_index}", amount=amount,
            method=method, reference=f"REF{rnd.randint(10000, 99999)}",
            tip=tip, late_fee=0, refund=0,
            deposit="No", confirmed="Yes",
            reconciled="Yes" if offset < -2 else "No",
            notes="Part payment - balance to follow" if share < 1 else ""))

    # ------------------------------------------------------------- expenses
    exp_plan = [
        (-29, "Fuel", "Weekly fuel - client visits", "QuickFuel", 62.40,
         "Card - online"),
        (-27, "Pet supplies", "Poop bags, wipes and treats", "PetMart",
         48.75, "Card - in person"),
        (-25, "Software & subscriptions", "Scheduling app monthly",
         "PetDesk", 29.00, "Direct debit"),
        (-24, "Marketing & advertising", "Local Facebook ads", "Meta",
         45.00, "Card - online"),
        (-22, "Insurance", "Business liability - monthly",
         "PetGuard Business Insurance", 48.00, "Direct debit"),
        (-20, "Cleaning supplies", "Enzyme cleaner and towels",
         "CleanCo", 32.10, "Cash"),
        (-18, "Equipment", "Two slip leads and a treat pouch", "PetMart",
         39.90, "Card - in person"),
        (-16, "Fuel", "Weekly fuel - client visits", "QuickFuel", 58.20,
         "Card - online"),
        (-14, "Vehicle maintenance", "Tyre rotation and oil", "AutoFix",
         86.00, "Bank transfer"),
        (-12, "Pet food & treats", "Emergency food supply", "PetMart",
         38.40, "Card - in person"),
        (-11, "Parking & tolls", "City parking - vet transport", "City Park",
         14.00, "Card - in person"),
        (-9, "Staff wages", "Jamie Lee - fortnight", "Payroll", 396.00,
         "Bank transfer"),
        (-8, "Training & certification", "Canine first aid refresher",
         "PetCare Academy", 65.00, "Card - online"),
        (-6, "Fuel", "Weekly fuel - client visits", "QuickFuel", 60.80,
         "Card - online"),
        (-5, "Phone & internet", "Mobile plan", "TelCo", 45.00,
         "Direct debit"),
        (-4, "Office & admin", "Printer paper and ink", "OfficeHub", 27.65,
         "Card - online"),
        (-3, "Contractor payments", "Tom Reed - weekend walks", "Tom Reed",
         145.00, "Bank transfer"),
        (-2, "Vet & emergency supplies", "First aid kit restock",
         "Riverton Vet Clinic", 42.30, "Card - in person"),
        (-1, "Pet supplies", "Waste bags bulk pack", "PetMart", 24.99,
         "Card - online"),
        (0, "Bank & merchant fees", "Card processing fees", "PayGate", 18.45,
         "Direct debit"),
    ]
    recurring_map = {
        "Software & subscriptions": "Monthly",
        "Insurance": "Monthly",
        "Phone & internet": "Monthly",
        "Staff wages": "Fortnightly",
    }
    for offset, category, desc, supplier, amount, method in exp_plan:
        s.expenses.append(dict(
            date=d(today, offset), category=category, description=desc,
            supplier=supplier, amount=amount, method=method, type="Business",
            deductible=1.0, receipt="Receipts folder",
            recurring=recurring_map.get(category, "No"),
            approval="Approved", notes=""))

    # Two special cases so the Business use % and Business / personal
    # columns are shown doing their job.
    s.expenses.append(dict(
        date=d(today, -7), category="Phone & internet",
        description="Home broadband - part business use", supplier="TelCo",
        amount=55.00, method="Direct debit", type="Business",
        deductible=0.6, receipt="Receipts folder", recurring="Monthly",
        approval="Approved",
        notes="Only 60% is claimable - see the Business use % column"))
    s.expenses.append(dict(
        date=d(today, -13), category="Pet supplies",
        description="Toys for my own dog", supplier="PetMart",
        amount=32.00, method="Card - in person", type="Personal",
        deductible=0.0, receipt="", recurring="No", approval="Approved",
        notes="Personal spend - left out of profit and tax figures"))

    # -------------------------------------------------------------- mileage
    for i in range(18):
        offset = -28 + i * 2
        staff = ["Alex Morgan", "Jamie Lee", "Priya Shah"][i % 3]
        client = s.clients[i % 7]["name"]
        start_odo = 48200 + i * 37
        km = round(rnd.uniform(6, 28), 1)
        s.mileage.append(dict(
            date=d(today, offset), staff=staff, client=client,
            purpose="Client visit" if i % 4 else "Supplies run",
            start_loc="Home office", end_loc=f"{client} - client home",
            start_odo=start_odo, end_odo=round(start_odo + km, 1),
            round_trip="Yes" if i % 3 else "No",
            fuel=0, parking=4.0 if i % 5 == 0 else 0,
            tolls=2.5 if i % 6 == 0 else 0,
            travel_mins=int(km * 2.2), notes=""))

    # --------------------------------------------------------------- health
    s.health = [
        dict(date=d(today, -9), time=t(7, 10), pet="Milo (Megan Foster)",
             type="Medication", item="Felimazole 2.5mg",
             dosage="1 tablet in food", sched_time=t(7, 0),
             given_time=t(7, 10), staff="Priya Shah", status="Given",
             supply=18, refill_due=d(today, 12), cost=0, notified="Yes",
             notes="Ate the full portion"),
        dict(date=d(today, -5), time=t(7, 5), pet="Milo (Megan Foster)",
             type="Medication", item="Felimazole 2.5mg",
             dosage="1 tablet in food", sched_time=t(7, 0),
             given_time=t(7, 5), staff="Priya Shah", status="Given",
             supply=16, refill_due=d(today, 12), cost=0, notified="Yes",
             notes=""),
        dict(date=d(today, -2), time=t(7, 0), pet="Milo (Megan Foster)",
             type="Medication", item="Felimazole 2.5mg",
             dosage="1 tablet in food", sched_time=t(7, 0),
             staff="Priya Shah", status="Missed", supply=16,
             refill_due=d(today, 12), cost=0, notified="Yes",
             notes="Refused food - owner called, giving it tonight"),
        dict(date=d(today, -20), time=t(10, 30), pet="Charlie (Robert Nguyen)",
             type="Vet visit", item="Arthritis review",
             dosage="Metacam 0.8ml daily", staff="Alex Morgan",
             status="Not required", vet_appt=d(today, 25), cost=110.0,
             notified="Yes", notes="Owner attended, we transported"),
        dict(date=d(today, -15), time=t(9, 0), pet="Bella (Sarah Mitchell)",
             type="Vaccination", item="Annual booster",
             dosage="-", staff="Alex Morgan", status="Not required",
             cost=86.0, notified="Yes", notes="Certificate photographed"),
        dict(date=d(today, -1), time=t(13, 20), pet="Teddy (Daniel Brooks)",
             type="Health observation", item="Loose stool",
             dosage="-", staff="Alex Morgan", status="Not required",
             cost=0, notified="Yes",
             notes="Monitor - owner changing food gradually"),
        dict(date=d(today, 3), time=t(7, 0), pet="Milo (Megan Foster)",
             type="Medication", item="Felimazole 2.5mg",
             dosage="1 tablet in food", sched_time=t(7, 0),
             staff="Priya Shah", status="Scheduled", supply=14,
             refill_due=d(today, 12), cost=0, notified="No", notes=""),
        dict(date=d(today, 10), time=t(11, 0), pet="Poppy (Chloe Bennett)",
             type="Vet visit", item="Nail trim and check",
             dosage="-", staff="Priya Shah", status="Scheduled",
             vet_appt=d(today, 10), cost=0, notified="Yes", notes=""),
    ]

    # ------------------------------------------------------------ incidents
    s.incidents = [
        dict(date=d(today, -22), time=t(10, 40), client="Luis Hernandez",
             pet="Ruby (Luis Hernandez)", staff="Alex Morgan",
             ref="BK-0007", type="Aggression", severity="Medium",
             description="Ruby lunged at an off-lead dog on Oakfield Road",
             action="Moved to a side street, treat scatter, finished the "
                    "walk on quiet streets",
             vet_contacted="No", clinic="-", owner_notified="Yes",
             notified_at="Same day 11:05", follow_up="Yes",
             follow_due=d(today, -15), status="Resolved", cost=0,
             claim="No", notes="Route changed permanently"),
        dict(date=d(today, -11), time=t(16, 15), client="Chloe Bennett",
             pet="Oscar (Chloe Bennett)", staff="Priya Shah",
             ref="BK-0015", type="Injury", severity="Low",
             description="Small cut on the left front paw from a stick",
             action="Cleaned with saline and covered, photo sent to owner",
             vet_contacted="No", clinic="-", owner_notified="Yes",
             notified_at="Same day 16:30", follow_up="Yes",
             follow_due=d(today, -4), status="Resolved", cost=0,
             claim="No", notes="Healed well, checked next visit"),
        dict(date=d(today, -3), time=t(8, 20), client="Megan Foster",
             pet="Simba (Megan Foster)", staff="Priya Shah",
             ref="BK-0020", type="Escape", severity="High",
             description="Simba dashed into the hallway when the front door "
                         "opened",
             action="Door closed immediately, Simba recovered within 2 "
                    "minutes, no injuries",
             vet_contacted="No", clinic="-", owner_notified="Yes",
             notified_at="Same day 08:35", follow_up="Yes",
             follow_due=d(today, 4), status="In progress", cost=0,
             claim="No",
             notes="Added a door-dash warning to the pet profile"),
    ]

    # ------------------------------------------------------------- messages
    msg_plan = [
        (-26, "Sarah Mitchell", "Booking confirmation", "Email", "Outgoing",
         "Confirmed 4 walks for the month", "Yes", 5, "", "No"),
        (-24, "Priya Raman", "Visit update", "WhatsApp", "Outgoing",
         "Photo and walk summary sent", "Yes", 5, "", "No"),
        (-21, "Megan Foster", "Payment reminder", "Email", "Outgoing",
         "Invoice INV-1003 reminder", "Yes", 4, "", "No"),
        (-18, "Daniel Brooks", "Follow-up", "SMS", "Outgoing",
         "Checked in after Daisy's long run", "Yes", 5,
         "Daisy slept all afternoon - very happy", "No"),
        (-16, "Chloe Bennett", "Complaint", "Phone call", "Incoming",
         "Visit ran 15 minutes late", "Yes", 3,
         "Understood the traffic, asked for earlier slots", "Yes"),
        (-14, "James Okafor", "Booking reminder", "Email", "Outgoing",
         "Overnight stay starts Friday", "Yes", 5, "", "No"),
        (-12, "Luis Hernandez", "Thank you", "SMS", "Outgoing",
         "Thanks for the referral", "No", 5, "", "No"),
        (-9, "Sarah Mitchell", "Review request", "Email", "Outgoing",
         "Asked for a Google review", "Yes", 5,
         "Left a 5 star review - mentioned Jamie", "No"),
        (-7, "Megan Foster", "Visit update", "App message", "Outgoing",
         "Milo medication log sent", "Yes", 5, "", "No"),
        (-5, "Robert Nguyen", "Win-back offer", "Email", "Outgoing",
         "20% off the first walk back", "No", "", "", "No"),
        (-3, "Amelia Clarke", "Booking confirmation", "Email", "Outgoing",
         "Meet and greet booked for next week", "Yes", "", "", "No"),
        (-1, "Chloe Bennett", "Follow-up", "SMS", "Outgoing",
         "Checked Oscar's paw is healing", "Yes", 5,
         "Paw fine, thank you for the care", "No"),
    ]
    for (offset, client, mtype, channel, direction, summary, response,
         rating, feedback, complaint) in msg_plan:
        s.messages.append(dict(
            date=d(today, offset), client=client, type=mtype,
            channel=channel, direction=direction, summary=summary,
            response=response, rating=rating or None, feedback=feedback,
            complaint=complaint,
            issue="Visit arrived late due to traffic"
            if complaint == "Yes" else "",
            resolution="Resolved" if complaint == "Yes" else "Not required",
            resolved_date=d(today, offset + 1) if complaint == "Yes"
            else None,
            follow_due=d(today, offset + 30) if mtype == "Win-back offer"
            else None,
            follow_done="No" if mtype == "Win-back offer" else "Yes",
            notes=""))

    # ------------------------------------------------------------ inventory
    inv_items = [
        ("Waste bags (roll of 300)", "Waste bags", "PetMart", "Roll", 6.50,
         12, 6, 24, -20, None, 9),
        ("Training treats 500g", "Treats", "PetMart", "Bag", 9.20, 4, 5, 10,
         -14, 180, 6),
        ("Enzyme cleaner 1L", "Cleaning", "CleanCo", "Bottle", 11.40, 3, 2, 6,
         -25, 400, 2),
        ("Slip lead", "Leads & harnesses", "PetMart", "Each", 8.75, 6, 3, 6,
         -18, None, 1),
        ("Pet first aid kit", "First aid", "PetCare Academy", "Kit", 34.00,
         2, 2, 2, -60, 500, 0),
        ("Saline wound wash", "First aid", "Riverton Vet Clinic", "Bottle",
         6.80, 1, 2, 4, -30, 120, 1),
        ("Flea treatment (spot on)", "Medication", "Riverton Vet Clinic",
         "Pack", 22.00, 2, 1, 3, -45, 60, 1),
        ("Emergency dry food 2kg", "Pet food", "PetMart", "Bag", 14.60, 2, 1,
         3, -35, 150, 1),
        ("Hand sanitiser 500ml", "Cleaning", "CleanCo", "Bottle", 5.20, 5, 3,
         6, -22, 300, 2),
        ("Reflective vest", "Equipment", "SafetyGear", "Each", 18.00, 3, 2, 2,
         -90, None, 0),
        ("Poo bag dispenser", "Equipment", "PetMart", "Each", 4.20, 8, 4, 8,
         -40, None, 2),
        ("Business cards (250)", "Office", "PrintFast", "Pack", 32.00, 1, 1,
         2, -75, None, 1),
    ]
    for (item, category, supplier, unit, cost, qty, min_qty, reorder,
         purchase_off, expiry_off, used) in inv_items:
        s.inventory.append(dict(
            item=item, category=category, supplier=supplier, unit=unit,
            unit_cost=cost, qty=qty, min_qty=min_qty, reorder=reorder,
            last_purchase=d(today, purchase_off), qty_purchased=reorder,
            expiry=d(today, expiry_off) if expiry_off else None,
            used_month=used, notes=""))

    # ---------------------------------------------------------------- tasks
    s.tasks = [
        dict(task="Send October invoices", category="Invoicing",
             due=d(today, 2), owner="Alex Morgan", priority="High",
             status="In progress", notes="14 invoices to issue"),
        dict(task="Chase Tom Reed first aid certificate",
             category="Staff", due=d(today, -2), owner="Alex Morgan",
             priority="Urgent", status="Not started",
             notes="Certificate expired"),
        dict(task="Order waste bags and treats", category="Inventory",
             due=d(today, 5), owner="Jamie Lee", priority="Medium",
             status="Not started", notes="Two items below minimum"),
        dict(task="Call Robert Nguyen - win-back", category="Marketing",
             due=d(today, 7), owner="Alex Morgan", priority="Medium",
             status="Not started", notes="Inactive 4 months"),
        dict(task="Book insurance renewal review", category="Admin",
             due=d(today, 30), owner="Alex Morgan", priority="Low",
             status="Not started", notes=""),
        dict(task="Update holiday surcharge dates", category="Admin",
             due=d(today, 45), owner="Alex Morgan", priority="Low",
             status="Not started", notes="December peak period"),
    ]

    # ---------------------------------------------------- owner withdrawals
    s.drawings = [
        dict(date=d(today, -60), amount=400, notes="Owner drawings"),
        dict(date=d(today, -30), amount=450, notes="Owner drawings"),
        dict(date=d(today, -5), amount=350, notes="Owner drawings"),
    ]

    return s
````

---

## 3. The registers

One module per data tab. Each returns a `SheetSpec` - a list of columns with their formulas - and the writer turns that into a formatted, validated, protected sheet.

### `pet_sitting_spreadsheet/sheets/__init__.py`

*14 lines*

> One module per tab of the workbook.

````python
"""One module per tab of the workbook."""

from . import (bookings, budget, calc, clients, common, dashboard, expenses,
               health, help_tab, incidents, inventory, invoices, lists,
               messages, mileage, payments, pets, printables, reminders,
               reports, services, settings, staff, start_here, visits)

__all__ = [
    "bookings", "budget", "calc", "clients", "common", "dashboard",
    "expenses", "health", "help_tab", "incidents", "inventory", "invoices",
    "lists", "messages", "mileage", "payments", "pets", "printables",
    "reminders", "reports", "services", "settings", "staff", "start_here",
    "visits",
]
````

### `pet_sitting_spreadsheet/sheets/common.py`

*110 lines · 11 functions*

> Small building blocks shared by the register sheets.

| Defines | What it does |
|---------|--------------|
| `id_col(trigger, prefix, digits, width, header, note)` | Auto numbered ID column (blank until the trigger column is filled) |
| `month_col(date_key, key, header, width)` |  |
| `alert_cf()` | Traffic light colouring for the 'Alert' / 'Check' columns |
| `status_cf()` |  |
| `money_cf()` |  |
| `owing_cf()` |  |
| `dup_cf()` |  |
| `flag_yes_no_cf()` |  |
| `yesno(note)` |  |
| `rating_validation()` |  |
| `blank(trigger, body)` | `=IF(trigger="","",body)` with the standard row placeholders |

**Module constants:** `ROW_OFFSET` (HEADER_ROW).

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
    """Auto numbered ID column (blank until the trigger column is filled)."""
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
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"yyyy-mm"))'))


def alert_cf() -> List[Dict[str, Any]]:
    """Traffic light colouring for the 'Alert' / 'Check' columns."""
    return [
        cf_text("OK", "good", "begins with"),
        cf_text("Overdue", "bad"),
        cf_text("Missed", "bad"),
        cf_text("Conflict", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Duplicate", "bad"),
        cf_text("Action", "warn"),
        cf_text("Due", "warn"),
        cf_text("Check", "warn"),
        cf_text("Today", "info"),
        cf_text("Upcoming", "info"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    return [
        cf_text("Completed", "good"),
        cf_text("Confirmed", "good"),
        cf_text("Paid", "good"),
        cf_text("Active", "good"),
        cf_text("Resolved", "good"),
        cf_text("Given", "good"),
        cf_text("Pending", "warn"),
        cf_text("Draft", "warn"),
        cf_text("Enquiry", "warn"),
        cf_text("Partially", "warn"),
        cf_text("In progress", "info"),
        cf_text("Scheduled", "info"),
        cf_text("Sent", "info"),
        cf_text("Cancelled", "bad"),
        cf_text("No show", "bad"),
        cf_text("Missed", "bad"),
        cf_text("Overdue", "bad"),
        cf_text("Inactive", "muted"),
        cf_text("Lost", "bad"),
    ]


def money_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0, "fmt": "bold_good"}]


def owing_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0.004, "fmt": "bad"}]


def dup_cf() -> List[Dict[str, Any]]:
    return [cf_text("Duplicate", "bad"), cf_text("OK", "good")]


def flag_yes_no_cf() -> List[Dict[str, Any]]:
    return [cf_text("No", "warn"), cf_text("Yes", "good")]


def yesno(note: str = "") -> Dict[str, Any]:
    return yes_no(note)


def rating_validation() -> Dict[str, Any]:
    return {
        "validate": "integer",
        "criteria": "between",
        "minimum": 1,
        "maximum": 5,
        "error_title": "Rating 1-5",
        "error_message": "Enter a whole number from 1 (poor) to 5 "
                         "(excellent).",
    }


def blank(trigger: str, body: str) -> str:
    """`=IF(trigger="","",body)` with the standard row placeholders."""
    return f'=IF({{{trigger}}}{{r}}="","",{body})'
````

### `pet_sitting_spreadsheet/sheets/lists.py`

*190 lines · 3 functions*

> The `Lists` tab: every drop down list in the workbook lives here.
>
> Each column becomes a *dynamic* defined name (OFFSET + COUNTIF) so the user can
> add or remove options and every drop down in the workbook follows along.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |
| `ref(name)` | A1 reference to this list's column on the current row, e.g |
| `options(name)` | Static copy of a list (used for sample data sanity) |

**Module constants:** `SHEET` ('Lists'), `LISTS` (OrderedDict(...)), `ROWS` (40).

````python
"""
The `Lists` tab: every drop down list in the workbook lives here.

Each column becomes a *dynamic* defined name (OFFSET + COUNTIF) so the user can
add or remove options and every drop down in the workbook follows along.
"""

from __future__ import annotations

from collections import OrderedDict
from typing import List, Tuple

from ..config import Config, Palette
from ..layout import (DATA_ROW, HEADER_ROW, Registry, col_letter,
                      dynamic_name, quote_sheet)
from ..writer import SheetWriter

SHEET = "Lists"

#: name -> (column header, options)
LISTS: "OrderedDict[str, Tuple[str, List[str]]]" = OrderedDict([
    ("YesNoList", ("Yes / No", ["Yes", "No"])),
    ("YesNoNAList", ("Yes / No / N/A", ["Yes", "No", "N/A"])),
    ("ClientStatusList", ("Client status", [
        "Active", "Prospective", "On hold", "Inactive", "Lost"])),
    ("ContactMethodList", ("Preferred contact", [
        "Phone call", "SMS", "Email", "WhatsApp", "App message"])),
    ("ReferralSourceList", ("Referral source", [
        "Google search", "Facebook", "Instagram", "Word of mouth",
        "Client referral", "Vet clinic", "Flyer / poster", "Local directory",
        "Repeat client", "Other"])),
    ("ContractStatusList", ("Contract status", [
        "Not sent", "Sent", "Signed", "Expired", "Renewal due"])),
    ("KeyStatusList", ("Key / access status", [
        "Not held", "Key held", "Key returned", "Lockbox code",
        "Smart lock code", "Owner home"])),
    ("PetTypeList", ("Pet type", [
        "Dog", "Cat", "Bird", "Rabbit", "Guinea pig", "Small animal",
        "Reptile", "Fish", "Other"])),
    ("PetGenderList", ("Pet gender", [
        "Male", "Male (neutered)", "Female", "Female (spayed)", "Unknown"])),
    ("TemperamentList", ("Temperament", [
        "Friendly", "Calm", "Playful", "Energetic", "Shy", "Anxious",
        "Reactive", "Protective", "Needs slow introduction"])),
    ("RiskFlagList", ("Risk flag", [
        "None", "Bite risk", "Escape artist", "Door dasher",
        "Pulls on lead", "Not dog friendly", "Not cat friendly",
        "Resource guarding", "Needs muzzle", "Medical risk", "Other"])),
    ("ServiceCategoryList", ("Service category", [
        "Dog walking", "Drop-in visit", "Pet sitting", "Overnight stay",
        "House sitting", "Feeding visit", "Puppy care", "Senior pet care",
        "Medication visit", "Pet transport", "Add-on"])),
    ("BillingUnitList", ("Billing unit", [
        "Per visit", "Per hour", "Per night", "Per walk", "Per day",
        "Flat fee"])),
    ("BookingStatusList", ("Booking status", [
        "Enquiry", "Pending", "Confirmed", "In progress", "Completed",
        "Cancelled", "No show"])),
    ("RecurringList", ("Recurring pattern", [
        "None", "Daily", "Weekdays", "Weekly", "Fortnightly", "Monthly"])),
    ("DiscountTypeList", ("Discount type", [
        "None", "Loyalty", "Promotional", "Package", "Multi-pet",
        "Goodwill", "Custom"])),
    ("CancelReasonList", ("Cancellation reason", [
        "Client cancelled", "Client away", "Pet unwell", "Weather",
        "Staff unavailable", "Double booked", "Payment issue",
        "Rescheduled", "No reason given", "Other"])),
    ("VisitStatusList", ("Visit status", [
        "Scheduled", "Completed", "Partially completed", "Missed",
        "Cancelled"])),
    ("MoodList", ("Pet mood", [
        "Happy", "Calm", "Playful", "Sleepy", "Anxious", "Lethargic",
        "Unwell", "Aggressive"])),
    ("AppetiteList", ("Appetite / intake", [
        "Normal", "Reduced", "None", "Increased", "Not offered"])),
    ("PottyList", ("Potty / litter", [
        "Pee only", "Poo only", "Both", "None", "Litter cleaned",
        "Accident indoors"])),
    ("HealthRecordTypeList", ("Health record type", [
        "Medication", "Vaccination", "Vet visit", "Health observation",
        "Treatment", "Grooming", "Other"])),
    ("MedStatusList", ("Medication status", [
        "Scheduled", "Given", "Missed", "Refused", "Not required"])),
    ("IncidentTypeList", ("Incident type", [
        "Bite / nip", "Escape", "Injury", "Illness", "Property damage",
        "Lost key", "Vehicle incident", "Aggression", "Theft", "Other"])),
    ("SeverityList", ("Severity", ["Low", "Medium", "High", "Critical"])),
    ("IncidentStatusList", ("Incident status", [
        "Open", "In progress", "Resolved", "Closed"])),
    ("MessageTypeList", ("Message type", [
        "Booking confirmation", "Booking reminder", "Visit update",
        "Payment reminder", "Follow-up", "Thank you", "Referral thank-you",
        "Birthday / anniversary", "Complaint", "Review request",
        "Win-back offer", "Other"])),
    ("ChannelList", ("Channel", [
        "Phone call", "SMS", "Email", "WhatsApp", "In person",
        "App message", "Social media"])),
    ("DirectionList", ("Direction", ["Outgoing", "Incoming"])),
    ("ResolutionList", ("Resolution status", [
        "Not required", "Open", "In progress", "Resolved"])),
    ("InvoiceStatusList", ("Invoice status", [
        "Draft", "Sent", "Paid", "Partially paid", "Cancelled",
        "Refunded"])),
    ("PaymentMethodList", ("Payment method", [
        "Cash", "Bank transfer", "Card - in person", "Card - online",
        "Digital wallet", "Direct debit", "Cheque", "Other"])),
    ("ExpenseCategoryList", ("Expense category", [
        "Fuel", "Vehicle maintenance", "Parking & tolls", "Pet supplies",
        "Pet food & treats", "Cleaning supplies", "Equipment",
        "Marketing & advertising", "Software & subscriptions", "Insurance",
        "Licences & permits", "Staff wages", "Contractor payments",
        "Training & certification", "Phone & internet", "Office & admin",
        "Vet & emergency supplies", "Bank & merchant fees", "Other"])),
    ("ExpenseApprovalList", ("Approval status", [
        "Pending", "Approved", "Rejected", "Reimbursed"])),
    ("BusinessPersonalList", ("Business / personal", [
        "Business", "Personal", "Mixed"])),
    ("RecurringExpenseList", ("Recurring expense", [
        "No", "Weekly", "Fortnightly", "Monthly", "Quarterly", "Annually"])),
    ("TripPurposeList", ("Trip purpose", [
        "Client visit", "Dog walk", "Supplies run", "Vet trip",
        "Admin / bank", "Marketing", "Training", "Personal"])),
    ("StaffRoleList", ("Staff role", [
        "Owner", "Dog walker", "Pet sitter", "Assistant", "Admin",
        "Contractor", "Trainee"])),
    ("EmploymentTypeList", ("Employment type", [
        "Employee - full time", "Employee - part time", "Casual",
        "Contractor", "Volunteer"])),
    ("PayTypeList", ("Pay type", [
        "Hourly", "Per visit", "Salary", "Commission %"])),
    ("StaffStatusList", ("Staff status", ["Active", "On leave", "Inactive"])),
    ("InventoryCategoryList", ("Inventory category", [
        "Pet food", "Treats", "Medication", "First aid", "Cleaning",
        "Waste bags", "Leads & harnesses", "Toys", "Equipment", "Office",
        "Other"])),
    ("PriorityList", ("Priority", ["Low", "Medium", "High", "Urgent"])),
    ("TaskStatusList", ("Task status", [
        "Not started", "In progress", "Waiting", "Done", "Cancelled"])),
    ("RatingList", ("Satisfaction rating", ["1", "2", "3", "4", "5"])),
    ("MonthList", ("Dashboard period", [
        "All year", "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"])),
])

ROWS = 40


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    sheet = SheetWriter(book, fmts, cfg, reg, SHEET,
                        tab_color=Palette.grey_text)
    last_col = len(LISTS) - 1
    sheet.banner(
        "Customisable Drop Down Lists",
        "Edit, add or delete options below - every drop down in the workbook "
        "updates automatically. Keep each list in one unbroken block, "
        "starting in row 5.",
        last_col,
        tail_text="TIP: lists grow automatically (up to row 44). Deleting "
                  "an option does not change records that already use it.")

    header_fmt = fmts.chrome("header")
    input_fmt = fmts.input("text")
    for idx, (name, (header, values)) in enumerate(LISTS.items()):
        sheet.set_col(idx, idx, max(16, min(26, len(header) + 8)), input_fmt)
        sheet.write(HEADER_ROW - 1, idx, header, header_fmt)
        for offset, value in enumerate(values):
            sheet.write(DATA_ROW - 1 + offset, idx, value, input_fmt)
        book.define_name(name, dynamic_name(SHEET, col_letter(idx),
                                            DATA_ROW, DATA_ROW + ROWS - 1))

    sheet.ws.set_row(HEADER_ROW - 1, 30)
    sheet.ws.freeze_panes(HEADER_ROW, 0)
    sheet.finish(fit_width=False)
    return sheet


def ref(name: str) -> str:
    """A1 reference to this list's column on the current row, e.g.

        Lists!$AH{r}

    Used by the hidden Calc sheet to mirror a list down a block of rows.
    """
    idx = list(LISTS).index(name)
    return f"{quote_sheet(SHEET)}!${col_letter(idx)}{{r}}"


def options(name: str) -> List[str]:
    """Static copy of a list (used for sample data sanity)."""
    return list(LISTS[name][1])
````

### `pet_sitting_spreadsheet/sheets/settings.py`

*274 lines · 2 functions*

> The `Settings` tab.
>
> Every rate, surcharge, prefix and reminder window used by the rest of the
> workbook is entered once here and published as a defined name (TaxRate,
> MileageRate, PaymentTermsDays ...) so the formulas stay readable.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |
| `_time(text)` |  |

**Module constants:** `SHEET` ('Settings'), `LAST_COL` (8), `HOLIDAY_ROWS` (30).

````python
"""
The `Settings` tab.

Every rate, surcharge, prefix and reminder window used by the rest of the
workbook is entered once here and published as a defined name (TaxRate,
MileageRate, PaymentTermsDays ...) so the formulas stay readable.
"""

from __future__ import annotations

import datetime as _dt

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check, list_from, number_check

SHEET = "Settings"
LAST_COL = 8

HOLIDAY_ROWS = 30


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, SHEET, tab_color=Palette.accent)
    b = cfg.business

    s.banner("Business Setup & Settings",
             "Fill in the cream cells once. Grey cells are locked formulas. "
             "Everything else in the workbook reads these values.",
             LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 24, fmts.chrome("label"))
    s.set_col(2, 2, 16, fmts.chrome("label"))
    s.set_col(3, 3, 26)
    s.set_col(4, 4, 18)
    s.set_col(5, 8, 17, fmts.chrome("note"))

    r = 4  # 0 based row index -> Excel row 5
    r = s.section(r, "1. Business details (used on invoices and printables)",
                  LAST_COL)
    r = s.field(r, "Business name", b.name, name="Biz_Name", value_end=4,
                help_text="Shown on every printable form and invoice.")
    r = s.field(r, "Owner / main contact", b.owner, name="Biz_Owner",
                value_end=4)
    r = s.field(r, "Phone", b.phone, name="Biz_Phone", value_end=4)
    r = s.field(r, "Email", b.email, name="Biz_Email", value_end=4)
    r = s.field(r, "Website / social", b.website, name="Biz_Web", value_end=4)
    r = s.field(r, "Street address", b.address, name="Biz_Address",
                value_end=4)
    r = s.field(r, "City / suburb", b.city, name="Biz_City", value_end=4)
    r = s.field(r, "State / region", b.state, name="Biz_State", value_end=4)
    r = s.field(r, "Post / ZIP code", b.postcode, name="Biz_Post",
                value_end=4)
    r = s.field(r, "Country", b.country, name="Biz_Country", value_end=4)
    r = s.field(r, "Business / tax number", b.tax_id, name="Biz_TaxID",
                value_end=4)
    r = s.field(r, "Payment details for invoices", b.bank_details,
                name="Biz_Bank", value_end=4,
                help_text="Bank account, payment link or app handle.")
    r = s.field(r, "Insurance provider", b.insurance_provider,
                name="Biz_Insurer", value_end=4)
    r = s.field(r, "Insurance policy number", b.insurance_policy,
                name="Biz_Policy", value_end=4)
    r = s.field(r, "Insurance renewal date",
                cfg.today + _dt.timedelta(days=120), style="date",
                name="Biz_InsuranceRenewal", validate=date_check(),
                help_text="Feeds the renewal reminder on the Reminders tab.")
    r = s.field(r, "Licence / permit", b.licence_name, name="Biz_Licence",
                value_end=4)
    r = s.field(r, "Licence renewal date",
                cfg.today + _dt.timedelta(days=210), style="date",
                name="Biz_LicenceRenewal", validate=date_check())
    r += 1

    r = s.section(r, "2. Logo & branding", LAST_COL)
    r = s.note(r, "To add your logo: unprotect the sheet (Review > Unprotect "
                  f"Sheet, password '{cfg.password}'), click Insert > "
                  "Pictures, place the image in the box below, then protect "
                  "the sheet again. The printable forms have their own logo "
                  "placeholders.", LAST_COL, first_col=1, height=38)
    s.merge(r, 1, 3, "LOGO PLACEHOLDER - insert your image here",
            fmts.chrome("legend_calc", align="center"), row2=r + 3)
    r += 5

    r = s.section(r, "3. Currency, date & time formats", LAST_COL)
    r = s.field(r, "Currency symbol", cfg.currency_symbol,
                name="CurrencySymbol",
                help_text="Display only. To change the cell formatting use "
                          "Format Cells > Currency on the money columns.")
    r = s.field(r, "Currency code", cfg.currency_code, name="CurrencyCode")
    r = s.field(r, "Date format in use", cfg.date_format.upper(),
                mode="plain", help_text="Select a date column and use Format "
                                        "Cells to switch to your local style.")
    r = s.field(r, "Time format in use", cfg.time_format,
                mode="plain", help_text="24 hour clock: format as hh:mm.")
    r = s.field(r, "Distance unit", cfg.distance_unit, name="DistanceUnit",
                validate={"validate": "list", "source": ["km", "miles"]})
    r = s.field(r, "Financial year starts", _dt.date(cfg.today.year, 1, 1),
                style="date", name="FY_Start", validate=date_check())
    r += 1

    r = s.section(r, "4. Tax settings", LAST_COL)
    r = s.field(r, "Tax name", cfg.tax_name, name="TaxName",
                help_text="GST, VAT, Sales Tax ...")
    r = s.field(r, "Registered for tax?", cfg.tax_registered,
                name="TaxRegistered", validate=list_from("YesNoList"),
                help_text="No = tax is never added to bookings or invoices.")
    r = s.field(r, "Tax rate", cfg.tax_rate, style="percent", name="TaxRate",
                validate=number_check(0, 1),
                help_text="Enter as a percentage, e.g. 10%.")
    r = s.field(r, "Prices already include tax?", cfg.prices_include_tax,
                name="PricesIncludeTax", validate=list_from("YesNoList"),
                help_text="Yes = tax is extracted from the price. "
                          "No = tax is added on top.")
    r = s.field(r, "Tax reserve to set aside", cfg.tax_reserve_percent,
                style="percent", name="TaxReservePct",
                validate=number_check(0, 1),
                help_text="Used by the Budget tab to estimate money to park "
                          "for the tax bill.")
    r += 1

    r = s.section(r, "5. Invoicing & payments", LAST_COL)
    r = s.field(r, "Invoice prefix", cfg.invoice_prefix, name="InvoicePrefix")
    r = s.field(r, "First invoice number", cfg.invoice_start_number,
                style="int", name="InvoiceStartNo",
                validate=number_check(1, 999999, decimal=False))
    r = s.field(r, "Payment terms (days)", cfg.payment_terms_days,
                style="int", name="PaymentTermsDays",
                validate=number_check(0, 180, decimal=False),
                help_text="Invoice due date = invoice date + this many days.")
    r = s.field(r, "Standard deposit", cfg.deposit_percent, style="percent",
                name="DepositPct", validate=number_check(0, 1))
    r = s.field(r, "Late fee", cfg.late_fee_percent, style="percent",
                name="LateFeePct", validate=number_check(0, 1),
                help_text="Suggested late fee on overdue invoices.")
    r = s.field(r, "Client ID prefix", cfg.client_prefix, name="ClientPrefix")
    r = s.field(r, "Pet ID prefix", cfg.pet_prefix, name="PetPrefix")
    r = s.field(r, "Booking ID prefix", cfg.booking_prefix,
                name="BookingPrefix")
    r += 1

    r = s.section(r, "6. Pricing, surcharges & discounts", LAST_COL)
    r = s.field(r, "Additional pet fee (default)", cfg.extra_pet_fee,
                style="money", name="ExtraPetFeeDefault",
                validate=number_check(0, 10000),
                help_text="Used when a service has no fee of its own.")
    r = s.field(r, "Holiday surcharge", cfg.holiday_surcharge,
                style="percent", name="HolidaySurcharge",
                validate=number_check(0, 2))
    r = s.field(r, "Weekend surcharge", cfg.weekend_surcharge,
                style="percent", name="WeekendSurcharge",
                validate=number_check(0, 2))
    r = s.field(r, "Emergency / last-minute fee", cfg.emergency_fee,
                style="money", name="EmergencyFee",
                validate=number_check(0, 10000))
    r = s.field(r, "Last-minute window (days)", cfg.last_minute_days,
                style="int", name="LastMinuteDays",
                validate=number_check(0, 30, decimal=False),
                help_text="Booked this close to the visit = emergency fee.")
    r = s.field(r, "Loyalty discount", cfg.loyalty_discount, style="percent",
                name="LoyaltyDiscount", validate=number_check(0, 1))
    r = s.field(r, "Loyalty threshold (bookings)", cfg.loyalty_threshold,
                style="int", name="LoyaltyThreshold",
                validate=number_check(0, 1000, decimal=False),
                help_text="Clients above this many bookings are flagged for "
                          "the loyalty discount.")
    r = s.field(r, "Average staff cost per hour", cfg.staff_cost_rate,
                style="money", name="StaffCostRate",
                validate=number_check(0, 1000),
                help_text="Used for service profitability estimates.")
    r += 1

    r = s.section(r, "7. Travel & mileage", LAST_COL)
    r = s.field(r, f"Mileage rate per {cfg.distance_unit}", cfg.mileage_rate,
                style="money", name="MileageRate",
                validate=number_check(0, 100),
                help_text="Your tax office reimbursement rate.")
    r = s.field(r, f"Client travel fee per {cfg.distance_unit}",
                cfg.travel_fee_per_unit, style="money", name="TravelFeeRate",
                validate=number_check(0, 100))
    r = s.field(r, f"Free travel distance ({cfg.distance_unit})",
                cfg.free_travel_distance, style="number1",
                name="FreeTravelDistance", validate=number_check(0, 1000),
                help_text="Travel inside this radius is not charged.")
    r = s.field(r, "Travel time allowance (mins)", cfg.travel_time_minutes,
                style="int", name="TravelTimeMins",
                validate=number_check(0, 240, decimal=False),
                help_text="Added to each booking when planning the day.")
    r += 1

    r = s.section(r, "8. Reminders, alerts & operations", LAST_COL)
    r = s.field(r, "Booking reminder lead (days)", cfg.reminder_lead_days,
                style="int", name="ReminderLeadDays",
                validate=number_check(0, 60, decimal=False))
    r = s.field(r, "Vaccination alert (days before)",
                cfg.vaccination_alert_days, style="int", name="VaccAlertDays",
                validate=number_check(0, 365, decimal=False))
    r = s.field(r, "Inactive client after (days)", cfg.inactive_client_days,
                style="int", name="InactiveDays",
                validate=number_check(0, 999, decimal=False))
    r = s.field(r, "Follow-up after service (days)", cfg.follow_up_days,
                style="int", name="FollowUpDays",
                validate=number_check(0, 365, decimal=False))
    r = s.field(r, "Expiry / renewal alert (days)", cfg.expiry_alert_days,
                style="int", name="ExpiryAlertDays",
                validate=number_check(0, 365, decimal=False))
    r = s.field(r, "Default visit length (mins)", cfg.default_visit_minutes,
                style="int", name="DefaultVisitMins",
                validate=number_check(0, 1440, decimal=False))
    r = s.field(r, "Working day starts", _time(cfg.day_start), style="time",
                name="DayStart")
    r = s.field(r, "Working day ends", _time(cfg.day_end), style="time",
                name="DayEnd")
    r += 1

    # ------------------------------------------------------------ holidays
    r = s.section(r, "9. Public holidays & peak dates (holiday surcharge)",
                  LAST_COL)
    r = s.note(r, "Any booking dated on one of these days automatically picks "
                  "up the holiday surcharge set above.", LAST_COL,
               first_col=1, height=20)
    head_row = r
    s.write(head_row, 1, "Date", fmts.chrome("header_left"))
    s.write(head_row, 2, "Occasion", fmts.chrome("header"))
    s.write(head_row, 3, "Notes", fmts.chrome("header"))
    s.ws.set_row(head_row, 22)

    first_holiday = head_row + 1
    date_fmt = fmts.input("date")
    text_fmt = fmts.input("text")
    for i in range(HOLIDAY_ROWS):
        s.ws.write_blank(first_holiday + i, 1, None, date_fmt)
        s.ws.write_blank(first_holiday + i, 2, None, text_fmt)
        s.ws.write_blank(first_holiday + i, 3, None, text_fmt)

    if cfg.include_sample_data:
        year = cfg.today.year
        samples = [
            (_dt.date(year, 1, 1), "New Year's Day", "Peak demand"),
            (_dt.date(year, 7, 4), "Independence Day", "Peak demand"),
            (_dt.date(year, 11, 26), "Thanksgiving", "Overnight bookings"),
            (_dt.date(year, 12, 24), "Christmas Eve", "Holiday surcharge"),
            (_dt.date(year, 12, 25), "Christmas Day", "Holiday surcharge"),
            (_dt.date(year, 12, 26), "Boxing Day", "Holiday surcharge"),
            (_dt.date(year, 12, 31), "New Year's Eve", "Holiday surcharge"),
            (_dt.date(year + 1, 1, 1), "New Year's Day", "Holiday surcharge"),
        ]
        for i, (day, name, note) in enumerate(samples):
            s.ws.write_datetime(first_holiday + i, 1, day, date_fmt)
            s.write(first_holiday + i, 2, name, text_fmt)
            s.write(first_holiday + i, 3, note, text_fmt)

    s.ws.data_validation(first_holiday, 1, first_holiday + HOLIDAY_ROWS - 1, 1,
                         date_check())
    last_holiday = first_holiday + HOLIDAY_ROWS
    book.define_name(
        "HolidayDates",
        f"={quote_sheet(SHEET)}!$B${first_holiday + 1}:$B${last_holiday}")

    r = last_holiday + 1
    r = s.note(r, "Backup tip: save a dated copy of this workbook at the end "
                  "of every month (File > Save As > add the date to the file "
                  "name). Keep client addresses and access codes private - "
                  "see the Help tab for data protection guidance.",
               LAST_COL, first_col=1, height=34)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s


def _time(text: str) -> _dt.time:
    hour, minute = text.split(":")
    return _dt.time(int(hour), int(minute))
````

### `pet_sitting_spreadsheet/sheets/services.py`

*92 lines · 1 function*

> `Services` - the service catalogue and price list that drives booking costs.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Services').

````python
"""`Services` - the service catalogue and price list that drives booking costs."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check, yes_no
from .common import blank, id_col, status_cf

NAME = "Services"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"SV-"', digits=3, width=9),
        Col("name", "Service name", 26, "text",
            note="This exact name is used by the Bookings drop down - "
                 "rename it here and every price follows."),
        Col("category", "Category", 16, "text",
            validate=list_from("ServiceCategoryList")),
        Col("description", "What is included", 38, "wrap"),
        Col("unit", "Billing unit", 13, "text",
            validate=list_from("BillingUnitList"),
            note="Per hour multiplies the rate by the booking duration. "
                 "Everything else multiplies the rate by the Qty column on "
                 "the booking."),
        Col("duration", "Standard length (mins)", 12, "int",
            validate=number_check(0, 1440, decimal=False)),
        Col("rate", "Base rate", 12, "money",
            validate=number_check(0, 100000)),
        Col("extra_pet", "Additional pet fee", 12, "money",
            validate=number_check(0, 10000),
            note="Leave blank to use the default additional pet fee from "
                 "the Settings tab."),
        Col("hol_pct", "Holiday surcharge", 11, "percent",
            validate=number_check(0, 2),
            note="Leave blank to use the global holiday surcharge."),
        Col("wkd_pct", "Weekend surcharge", 11, "percent",
            validate=number_check(0, 2),
            note="Leave blank to use the global weekend surcharge."),
        Col("travel", "Travel fee applies?", 11, "flag",
            validate=yes_no()),
        Col("staff_cost", "Staff cost per unit", 12, "money",
            validate=number_check(0, 100000),
            note="What it costs you to deliver one unit - used for the "
                 "profitability columns."),
        Col("profit", "Gross profit per unit", 12, "money",
            formula=blank("rate",
                          '{rate}{r}-IF({staff_cost}{r}="",0,'
                          '{staff_cost}{r})')),
        Col("margin", "Margin %", 10, "percent",
            formula='=IF(OR({rate}{r}="",{rate}{r}=0),"",'
                    '{profit}{r}/{rate}{r})',
            cf=[{"type": "3_color_scale", "min_color": "#F8696B",
                 "mid_color": "#FFEB84", "max_color": "#8DD1A4"}]),
        Col("times_booked", "Times booked", 11, "int",
            formula=blank("name",
                          'COUNTIFS([[Bookings.service]],{name}{r},'
                          '[[Bookings.status]],"<>Cancelled")')),
        Col("revenue", "Revenue to date", 13, "money",
            formula=blank("name",
                          'SUMIFS([[Bookings.total]],[[Bookings.service]],'
                          '{name}{r},[[Bookings.status]],"<>Cancelled")'),
            cf=[{"type": "data_bar", "bar_color": "#2E7D8F"}]),
        Col("avg_value", "Average value", 12, "money",
            formula='=IF(OR({times_booked}{r}="",{times_booked}{r}=0),"",'
                    '{revenue}{r}/{times_booked}{r})'),
        Col("est_profit", "Estimated profit", 13, "money",
            formula='=IF(OR({times_booked}{r}="",{profit}{r}=""),"",'
                    '{times_booked}{r}*{profit}{r})',
            cf=[{"type": "data_bar", "bar_color": "#E2A33C"}]),
        Col("active", "Active?", 9, "flag", validate=yes_no(),
            cf=status_cf()),
        Col("notes", "Notes", 30, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Services & Price List",
        subtitle="Add every service you sell. Bookings pull the billing "
                 "unit, rate, additional pet fee and surcharges straight "
                 "from this table.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).services,
        help_text="Cream cells = your prices. Grey cells are locked "
                  "formulas: profit, margin, how often each service sells "
                  "and what it has earned.",
    )
````

### `pet_sitting_spreadsheet/sheets/clients.py`

*152 lines · 1 function*

> `Clients` - the client database with lifetime value and follow-up flags.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Clients').

````python
"""`Clients` - the client database with lifetime value and follow-up flags."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, yes_no
from .common import (alert_cf, blank, dup_cf, id_col, owing_cf,
                     status_cf)

NAME = "Clients"

_LAST_SERVICE = ('SUMPRODUCT(MAX(([[Bookings.client]]={name}{r})'
                 '*([[Bookings.status]]="Completed")*[[Bookings.date]]))')
_NEXT_BOOKING = ('SUMPRODUCT(MIN((([[Bookings.client]]={name}{r})'
                 '*([[Bookings.date]]>=TODAY())'
                 '*([[Bookings.status]]<>"Cancelled"))*[[Bookings.date]]'
                 '+(1-(([[Bookings.client]]={name}{r})'
                 '*([[Bookings.date]]>=TODAY())'
                 '*([[Bookings.status]]<>"Cancelled")))*2958465))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ClientPrefix"),
        Col("name", "Client name", 22, "text",
            note="Used everywhere else in the workbook. If two clients "
                 "share a name add an initial, e.g. 'Sam B.'"),
        Col("phone", "Phone", 17, "text"),
        Col("email", "Email", 26, "text"),
        Col("address", "Street address", 26, "text"),
        Col("city", "City / suburb", 15, "text"),
        Col("postcode", "Post code", 10, "text"),
        Col("emergency_name", "Emergency contact", 20, "text"),
        Col("emergency_phone", "Emergency phone", 17, "text"),
        Col("contact_method", "Preferred contact", 14, "text",
            validate=list_from("ContactMethodList")),
        Col("referral", "Referral source", 16, "text",
            validate=list_from("ReferralSourceList")),
        Col("reg_date", "Registered on", 13, "date", validate=date_check()),
        Col("status", "Status", 12, "status",
            validate=list_from("ClientStatusList"), cf=status_cf()),
        Col("key_status", "Key / access", 14, "text",
            validate=list_from("KeyStatusList")),
        Col("key_tag", "Key tag / ref", 12, "text"),
        Col("key_returned", "Key returned on", 13, "date",
            validate=date_check()),
        Col("access_notes", "Home access instructions", 32, "wrap"),
        Col("access_code", "Alarm / gate code", 16, "text",
            note="Security tip: store full alarm codes in a separate "
                 "password manager and keep only a hint here."),
        Col("contract", "Contract status", 13, "status",
            validate=list_from("ContractStatusList"), cf=status_cf()),
        Col("agreement_date", "Agreement signed", 13, "date",
            validate=date_check()),
        Col("vet_consent", "Emergency vet consent", 11, "flag",
            validate=yes_no(), cf=status_cf()),
        Col("vet_limit", "Vet spend authority", 12, "money",
            validate=number_check(0, 100000)),
        Col("preferences", "Preferences & special instructions", 36, "wrap"),
        Col("notes", "Notes & communication history", 36, "wrap"),
        Col("pets", "Pets", 8, "int",
            formula=blank("name", 'COUNTIF([[Pets.owner]],{name}{r})')),
        Col("bookings", "Bookings", 9, "int",
            formula=blank("name",
                          'COUNTIFS([[Bookings.client]],{name}{r})')),
        Col("visits", "Visits done", 9, "int",
            formula=blank("name",
                          'COUNTIFS([[Visit Log.client]],{name}{r},'
                          '[[Visit Log.status]],"Completed")')),
        Col("last_service", "Last service", 13, "date",
            formula=('=IF({name}{r}="","",'
                     f'IF({_LAST_SERVICE}=0,"",{_LAST_SERVICE}))')),
        Col("days_since", "Days since", 9, "int",
            formula='=IF({last_service}{r}="","",TODAY()-{last_service}{r})',
            cf=[{"type": "cell", "criteria": ">",
                 "value": "=InactiveDays", "fmt": "warn"}]),
        Col("next_booking", "Next booking", 13, "date",
            formula=('=IF({name}{r}="","",IF(COUNTIFS([[Bookings.client]],'
                     '{name}{r},[[Bookings.date]],">="&TODAY(),'
                     '[[Bookings.status]],"<>Cancelled")=0,"",'
                     f'{_NEXT_BOOKING}))')),
        Col("revenue", "Lifetime revenue", 14, "money",
            formula=blank("name",
                          'SUMIFS([[Payments.received]],'
                          '[[Payments.client]],{name}{r})'),
            cf=[{"type": "data_bar", "bar_color": "#2E7D8F"}]),
        Col("invoiced", "Invoiced to date", 14, "money",
            formula=blank("name",
                          'SUMIFS([[Invoices.total]],[[Invoices.client]],'
                          '{name}{r},[[Invoices.status]],"<>Draft")')),
        Col("outstanding", "Outstanding", 13, "money",
            formula=blank("name",
                          'SUMIFS([[Invoices.balance]],[[Invoices.client]],'
                          '{name}{r},[[Invoices.status]],"<>Cancelled")'),
            cf=owing_cf()),
        Col("avg_booking", "Average booking", 13, "money",
            formula='=IF(OR({bookings}{r}="",{bookings}{r}=0),"",'
                    'SUMIFS([[Bookings.total]],[[Bookings.client]],'
                    '{name}{r},[[Bookings.status]],"<>Cancelled")'
                    '/{bookings}{r})'),
        Col("loyalty", "Loyalty status", 14, "status",
            formula=blank("name",
                          'IF({bookings}{r}>=LoyaltyThreshold,'
                          '"Loyalty discount","Standard")'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Loyalty", "fmt": "good"}]),
        Col("dup_check", "Duplicate check", 14, "status",
            formula=blank("name",
                          'IF(COUNTIF([[Clients.name]],{name}{r})>1,'
                          '"Duplicate name",IF(AND({phone}{r}<>"",'
                          'COUNTIF([[Clients.phone]],{phone}{r})>1),'
                          '"Duplicate phone","OK"))'),
            cf=dup_cf()),
        Col("data_check", "Missing information", 20, "status",
            formula=blank("name",
                          'IF(AND({phone}{r}<>"",{email}{r}<>"",'
                          '{reg_date}{r}<>"",{status}{r}<>""),"OK",'
                          '"Add: "&IF({phone}{r}="","phone ","")'
                          '&IF({email}{r}="","email ","")'
                          '&IF({reg_date}{r}="","date ","")'
                          '&IF({status}{r}="","status",""))'),
            cf=alert_cf()),
        Col("alert", "Action needed", 22, "status",
            formula=blank(
                "name",
                'IF({outstanding}{r}>0.5,"Overdue - chase payment",'
                'IF(AND({status}{r}="Active",ISNUMBER({days_since}{r}),'
                '{days_since}{r}>InactiveDays),"Inactive - win back",'
                'IF(AND({status}{r}="Active",{contract}{r}<>"Signed"),'
                '"Action - contract not signed",'
                'IF(AND({status}{r}="Prospective",{next_booking}{r}=""),'
                '"Action - follow up quote","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Client Database",
        subtitle="One row per client. Grey columns calculate pets, visits, "
                 "lifetime revenue, outstanding balances and follow-up "
                 "actions for you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).clients,
        help_text="Type the client name first - the Client ID appears "
                  "automatically. Use the filter arrows in the header row "
                  "to search; avoid re-sorting rows so IDs stay stable.",
    )
````

### `pet_sitting_spreadsheet/sheets/pets.py`

*136 lines · 1 function*

> `Pets` - full pet profiles: care routine, medical, vaccinations, warnings.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Pets').

````python
"""`Pets` - full pet profiles: care routine, medical, vaccinations, warnings."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col

NAME = "Pets"

_LAST_VISIT = ('SUMPRODUCT(MAX(([[Visit Log.pet]]={label}{r})'
               '*([[Visit Log.status]]="Completed")*[[Visit Log.date]]))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "PetPrefix"),
        Col("name", "Pet name", 14, "text"),
        Col("owner", "Owner (client)", 20, "text",
            validate=list_from("ClientNames"),
            note="Pick the client from the drop down so bookings, visits "
                 "and invoices all line up."),
        Col("label", "Pet reference", 24, "text",
            formula='=IF({name}{r}="","",{name}{r}&" ("&{owner}{r}&")")',
            note="This is the value used by the Bookings, Visit Log and "
                 "Health Log drop downs."),
        Col("client_id", "Client ID", 10, "text",
            formula=blank("owner",
                          'IFERROR(INDEX([[Clients.id]],'
                          'MATCH({owner}{r},[[Clients.name]],0)),"")')),
        Col("type", "Pet type", 11, "text",
            validate=list_from("PetTypeList")),
        Col("breed", "Breed", 20, "text"),
        Col("dob", "Date of birth", 13, "date", validate=date_check()),
        Col("age", "Age (yrs)", 9, "number1",
            formula=blank("dob", 'ROUND(YEARFRAC({dob}{r},TODAY()),1)')),
        Col("gender", "Gender", 14, "text",
            validate=list_from("PetGenderList")),
        Col("weight", "Weight (kg)", 10, "number1",
            validate=number_check(0, 200)),
        Col("colour", "Colour & markings", 22, "text"),
        Col("photo", "Photo link", 18, "text",
            note="Paste a link to the photo in your cloud folder, or "
                 "insert a small picture next to the row."),
        Col("microchip", "Microchip number", 16, "text"),
        Col("temperament", "Temperament", 16, "text",
            validate=list_from("TemperamentList")),
        Col("behaviour", "Behaviour issues", 28, "wrap"),
        Col("allergies", "Allergies", 20, "wrap"),
        Col("medical", "Medical conditions / special needs", 30, "wrap"),
        Col("vet_name", "Veterinarian", 22, "text"),
        Col("vet_phone", "Vet phone", 16, "text"),
        Col("emerg_clinic", "Emergency clinic", 22, "text"),
        Col("emerg_phone", "Emergency clinic phone", 16, "text"),
        Col("vacc_date", "Vaccinated on", 13, "date", validate=date_check()),
        Col("vacc_expiry", "Vaccination expires", 13, "date",
            validate=date_check()),
        Col("vacc_status", "Vaccination status", 17, "status",
            formula=blank("vacc_expiry",
                          'IF({vacc_expiry}{r}<TODAY(),"Expired",'
                          'IF({vacc_expiry}{r}-TODAY()<=VaccAlertDays,'
                          '"Due in "&TEXT({vacc_expiry}{r}-TODAY(),"0")'
                          '&" days","Valid"))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Expired", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Due in", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Valid", "fmt": "good"}]),
        Col("insurer", "Pet insurance", 18, "text"),
        Col("policy", "Policy number", 14, "text"),
        Col("insurance_expiry", "Insurance renewal", 13, "date",
            validate=date_check()),
        Col("feeding", "Feeding instructions", 30, "wrap"),
        Col("food", "Food type", 18, "text"),
        Col("portion", "Portion size", 13, "text"),
        Col("feed_schedule", "Feeding schedule", 18, "text"),
        Col("medication", "Medication", 20, "text"),
        Col("dosage", "Dosage & instructions", 26, "wrap"),
        Col("med_schedule", "Medication schedule", 18, "text"),
        Col("walking", "Walking & exercise preferences", 30, "wrap"),
        Col("potty", "Potty / litter routine", 26, "wrap"),
        Col("grooming", "Grooming requirements", 24, "wrap"),
        Col("favourites", "Favourite toys, treats & games", 28, "wrap"),
        Col("risk", "Risk flag", 16, "status",
            validate=list_from("RiskFlagList"),
            cf=[{"type": "text", "criteria": "not containing",
                 "value": "None", "fmt": "warn"}]),
        Col("handling", "Handling instructions", 30, "wrap"),
        Col("warnings", "Safety warnings", 30, "wrap"),
        Col("visits", "Visits", 8, "int",
            formula=blank("label",
                          'COUNTIFS([[Visit Log.pet]],{label}{r},'
                          '[[Visit Log.status]],"Completed")')),
        Col("last_visit", "Last visit", 13, "date",
            formula=('=IF({label}{r}="","",'
                     f'IF({_LAST_VISIT}=0,"",{_LAST_VISIT}))')),
        Col("meds_logged", "Med records", 10, "int",
            formula=blank("label",
                          'COUNTIFS([[Health Log.pet]],{label}{r})')),
        Col("alert", "Alerts", 30, "status",
            formula=blank(
                "name",
                'TRIM(IF({vacc_status}{r}="Expired","VACCINATION EXPIRED. ",'
                'IF(LEFT({vacc_status}{r},6)="Due in","Vaccination due. ",""))'
                '&IF(AND({risk}{r}<>"",{risk}{r}<>"None"),{risk}{r}&". ","")'
                '&IF(AND({medication}{r}<>"",{medication}{r}<>"None",'
                '{medication}{r}<>"-"),"On medication. ","")'
                '&IF({allergies}{r}<>"",IF(OR({allergies}{r}="None",'
                '{allergies}{r}="None known"),"","Allergy: "'
                '&{allergies}{r}),""))'),
            cf=alert_cf() + [{"type": "text", "criteria": "containing",
                              "value": "EXPIRED", "fmt": "bold_bad"},
                             {"type": "text", "criteria": "containing",
                              "value": "Allergy", "fmt": "warn"},
                             {"type": "text", "criteria": "containing",
                              "value": "risk", "fmt": "warn"}]),
    ]

    return SheetSpec(
        name=NAME,
        title="Pet Profiles",
        subtitle="Everything a sitter needs before they walk through the "
                 "door: routine, medical, vet, vaccination status and "
                 "safety warnings.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).pets,
        help_text="Choose the owner from the drop down first. The grey "
                  "'Pet reference' column is what bookings and visit logs "
                  "use, so two pets with the same name never get mixed up.",
    )
````

### `pet_sitting_spreadsheet/sheets/bookings.py`

*225 lines · 2 functions*

> `Bookings` - the scheduling engine: prices, surcharges, conflicts, alerts.

| Defines | What it does |
|---------|--------------|
| `_svc(field, fallback)` | INDEX/MATCH into the Services table for the booked service |
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Bookings').

````python
"""`Bookings` - the scheduling engine: prices, surcharges, conflicts, alerts."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, yes_no
from .common import alert_cf, blank, id_col, month_col, status_cf

NAME = "Bookings"


def _svc(field: str, fallback: str = "0") -> str:
    """INDEX/MATCH into the Services table for the booked service."""
    return (f'IFERROR(INDEX([[Services.{field}]],'
            f'MATCH({{service}}{{r}},[[Services.name]],0)),{fallback})')


def spec(cfg: Config) -> SheetSpec:
    conflict = ('SUMPRODUCT(([[Bookings.staff]]={staff}{r})'
                '*([[Bookings.date]]={date}{r})'
                '*([[Bookings.start]]<{end}{r})'
                '*([[Bookings.end]]>{start}{r})'
                '*([[Bookings.status]]<>"Cancelled"))')

    columns = [
        id_col("client", "BookingPrefix"),
        Col("request_date", "Requested on", 12, "date",
            validate=date_check(),
            note="When the client asked for the booking. Used to spot "
                 "last-minute bookings."),
        Col("client", "Client", 20, "text",
            validate=list_from("ClientNames")),
        Col("client_id", "Client ID", 10, "text",
            formula=blank("client",
                          'IFERROR(INDEX([[Clients.id]],'
                          'MATCH({client}{r},[[Clients.name]],0)),"")')),
        Col("pet", "Pet", 24, "text", validate=list_from("PetLabels")),
        Col("extra_pets", "Additional pets", 20, "text"),
        Col("pet_count", "No. of pets", 9, "int",
            validate=number_check(1, 20, decimal=False),
            note="Blank counts as 1. Extra pets are charged the additional "
                 "pet fee."),
        Col("service", "Service", 24, "text",
            validate=list_from("ServiceNames")),
        Col("date", "Booking date", 12, "date", validate=date_check()),
        Col("day", "Day", 7, "flag",
            formula=blank("date", 'TEXT({date}{r},"ddd")')),
        Col("start", "Start", 9, "time"),
        Col("end", "End", 9, "time"),
        Col("duration", "Hours", 8, "number",
            formula='=IF(OR({start}{r}="",{end}{r}=""),"",'
                    'ROUND(({end}{r}-{start}{r}+IF({end}{r}<{start}{r},1,0))'
                    '*24,2))'),
        Col("qty", "Qty / nights", 9, "number",
            validate=number_check(0, 1000),
            note="Number of visits, nights or hours being billed. Blank "
                 "counts as 1."),
        Col("unit", "Billing unit", 11, "text",
            formula=blank("service", _svc("unit", '""'))),
        Col("rate", "Rate", 11, "money",
            formula=blank("service", _svc("rate"))),
        Col("pet_rate", "Extra pet rate", 10, "money", hidden=True,
            formula=blank("service",
                          f'IF({_svc("extra_pet")}=0,ExtraPetFeeDefault,'
                          f'{_svc("extra_pet")})')),
        Col("hol_pct", "Holiday %", 9, "percent", hidden=True,
            formula=blank("service",
                          f'IF({_svc("hol_pct")}=0,HolidaySurcharge,'
                          f'{_svc("hol_pct")})')),
        Col("wkd_pct", "Weekend %", 9, "percent", hidden=True,
            formula=blank("service",
                          f'IF({_svc("wkd_pct")}=0,WeekendSurcharge,'
                          f'{_svc("wkd_pct")})')),
        Col("travel_flag", "Travel billable", 9, "flag", hidden=True,
            formula=blank("service", _svc("travel", '"No"'))),
        Col("base", "Base charge", 12, "money",
            formula=blank(
                "service",
                'ROUND(IF({unit}{r}="Per hour",{rate}{r}*IF({duration}{r}="",'
                'IF({qty}{r}="",1,{qty}{r}),{duration}{r}),'
                '{rate}{r}*IF({qty}{r}="",1,{qty}{r})),2)')),
        Col("extra_pet_fee", "Extra pet fee", 11, "money",
            formula=blank("service",
                          'ROUND(MAX(0,IF({pet_count}{r}="",1,'
                          '{pet_count}{r})-1)*{pet_rate}{r},2)')),
        Col("holiday", "Holiday?", 9, "flag",
            formula=blank("date",
                          'IF(COUNTIF(HolidayDates,{date}{r})>0,"Yes","No")'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "warn"}]),
        Col("weekend", "Weekend?", 9, "flag",
            formula=blank("date",
                          'IF(WEEKDAY({date}{r},2)>5,"Yes","No")'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "info"}]),
        Col("surcharge", "Surcharges", 11, "money",
            formula=blank("service",
                          'ROUND(({base}{r}+{extra_pet_fee}{r})'
                          '*(IF({holiday}{r}="Yes",{hol_pct}{r},0)'
                          '+IF({weekend}{r}="Yes",{wkd_pct}{r},0)),2)')),
        Col("distance", "Travel distance", 10, "number1",
            validate=number_check(0, 2000)),
        Col("travel_fee", "Travel fee", 10, "money",
            formula=blank("service",
                          'IF({travel_flag}{r}<>"Yes",0,'
                          'ROUND(MAX(0,IF({distance}{r}="",0,'
                          '{distance}{r})-FreeTravelDistance)'
                          '*TravelFeeRate,2))')),
        Col("travel_time", "Travel allowance (min)", 10, "int",
            formula=blank("date", "TravelTimeMins")),
        Col("rush_fee", "Last-minute fee", 10, "money",
            formula='=IF(OR({date}{r}="",{request_date}{r}=""),0,'
                    'IF({date}{r}-{request_date}{r}<=LastMinuteDays,'
                    'EmergencyFee,0))',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("discount_type", "Discount type", 13, "text",
            validate=list_from("DiscountTypeList")),
        Col("discount_pct", "Discount %", 9, "percent",
            validate=number_check(0, 1)),
        Col("discount", "Discount", 11, "money",
            formula=blank("service",
                          'ROUND(({base}{r}+{extra_pet_fee}{r}'
                          '+{surcharge}{r}+{travel_fee}{r}+{rush_fee}{r})'
                          '*IF({discount_pct}{r}="",0,{discount_pct}{r}),2)')),
        Col("override", "Price override", 11, "money",
            validate=number_check(0, 100000),
            note="Only fill this in to force a price. Always add the reason "
                 "in the next column."),
        Col("override_note", "Override reason", 20, "text"),
        Col("subtotal", "Subtotal", 12, "money",
            formula=blank("service",
                          'IF({override}{r}<>"",{override}{r},'
                          'ROUND({base}{r}+{extra_pet_fee}{r}+{surcharge}{r}'
                          '+{travel_fee}{r}+{rush_fee}{r}-{discount}{r},2))')),
        Col("tax", "Tax", 10, "money",
            formula='=IF({subtotal}{r}="","",IF(TaxRegistered<>"Yes",0,'
                    'IF(PricesIncludeTax="Yes",'
                    'ROUND({subtotal}{r}-{subtotal}{r}/(1+TaxRate),2),'
                    'ROUND({subtotal}{r}*TaxRate,2))))'),
        Col("total", "Booking total", 13, "money",
            formula='=IF({subtotal}{r}="","",IF(PricesIncludeTax="Yes",'
                    '{subtotal}{r},ROUND({subtotal}{r}+{tax}{r},2)))',
            cf=[{"type": "data_bar", "bar_color": "#2E7D8F"}]),
        Col("staff", "Assigned to", 16, "text",
            validate=list_from("StaffNames")),
        Col("status", "Status", 13, "status",
            validate=list_from("BookingStatusList"), cf=status_cf()),
        Col("confirm_sent", "Confirmed to client", 10, "flag",
            validate=yes_no(), cf=status_cf()),
        Col("reminder_sent", "Reminder sent", 10, "flag",
            validate=yes_no(), cf=status_cf()),
        Col("recurring", "Repeats", 12, "text",
            validate=list_from("RecurringList")),
        Col("next_due", "Next booking due", 13, "date",
            formula='=IF(OR({date}{r}="",{recurring}{r}="",'
                    '{recurring}{r}="None"),"",'
                    'IF({recurring}{r}="Monthly",EDATE({date}{r},1),'
                    '{date}{r}+IF({recurring}{r}="Daily",1,'
                    'IF({recurring}{r}="Weekdays",'
                    'IF(WEEKDAY({date}{r},2)>=5,8-WEEKDAY({date}{r},2),1),'
                    'IF({recurring}{r}="Weekly",7,'
                    'IF({recurring}{r}="Fortnightly",14,0))))))'),
        Col("conflict", "Schedule check", 18, "status",
            formula=('=IF(OR({date}{r}="",{staff}{r}="",{start}{r}="",'
                     '{end}{r}=""),"",IF(' + conflict +
                     '>1,"Conflict - double booked","OK"))'),
            cf=alert_cf()),
        Col("visit_logged", "Visit logged", 10, "flag",
            formula=blank("id",
                          'IF(COUNTIF([[Visit Log.booking]],{id}{r})>0,'
                          '"Yes","No")'),
            cf=status_cf()),
        Col("invoice_no", "Invoice", 12, "text",
            formula=blank("id",
                          'IFERROR(INDEX([[Invoices.invoice_no]],'
                          'MATCH({id}{r},[[Invoices.booking]],0)),"")')),
        Col("paid_status", "Payment status", 15, "status",
            formula=blank("id",
                          'IF({invoice_no}{r}="","Not invoiced",'
                          'IFERROR(INDEX([[Invoices.pay_status]],'
                          'MATCH({invoice_no}{r},'
                          '[[Invoices.invoice_no]],0)),""))'),
            cf=status_cf() + [{"type": "text", "criteria": "containing",
                               "value": "Not invoiced", "fmt": "warn"}]),
        Col("cancel_reason", "Cancellation reason", 18, "text",
            validate=list_from("CancelReasonList")),
        Col("resched_from", "Rescheduled from", 12, "date",
            validate=date_check()),
        Col("notes", "Notes & special requests", 34, "wrap"),
        month_col("date"),
        Col("alert", "Action needed", 24, "status",
            formula=blank(
                "date",
                'IF({status}{r}="Cancelled","Cancelled",'
                'IF(AND({date}{r}<TODAY(),OR({status}{r}="Pending",'
                '{status}{r}="Confirmed",{status}{r}="Enquiry")),'
                '"Overdue - update status",'
                'IF(AND({status}{r}="Completed",{visit_logged}{r}="No"),'
                '"Action - log the visit",'
                'IF({date}{r}=TODAY(),"Today",'
                'IF(AND({date}{r}>TODAY(),{date}{r}-TODAY()<=ReminderLeadDays,'
                '{confirm_sent}{r}<>"Yes"),"Action - unconfirmed",'
                'IF(AND({date}{r}>TODAY(),'
                '{date}{r}-TODAY()<=ReminderLeadDays),"Upcoming","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Bookings & Schedule",
        subtitle="Every booking, priced automatically from the Services "
                 "tab with holiday, weekend, extra pet, travel and "
                 "last-minute rules applied.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=3,
        samples=sample(cfg).bookings,
        help_text="Fill the cream cells from left to right. Everything "
                  "grey - prices, surcharges, totals, conflicts and alerts "
                  "- is calculated and locked.",
    )
````

### `pet_sitting_spreadsheet/sheets/visits.py`

*134 lines · 2 functions*

> `Visit Log` - the daily pet care log, one row per visit or walk.

| Defines | What it does |
|---------|--------------|
| `_bk(field, fallback)` |  |
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Visit Log').

````python
"""`Visit Log` - the daily pet care log, one row per visit or walk."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, yes_no
from .common import alert_cf, blank, id_col, month_col, status_cf

NAME = "Visit Log"


def _bk(field: str, fallback: str = '""') -> str:
    return (f'IFERROR(INDEX([[Bookings.{field}]],'
            f'MATCH({{booking}}{{r}},[[Bookings.id]],0)),{fallback})')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"VS-"'),
        Col("booking", "Booking ID", 12, "text",
            validate=list_from("BookingIDs"),
            note="Pick the booking and the client, pet, service and "
                 "scheduled time fill themselves in."),
        Col("date", "Visit date", 12, "date", validate=date_check()),
        Col("client", "Client", 20, "text",
            formula=blank("booking", _bk("client"))),
        Col("pet", "Pet", 24, "text",
            formula=blank("booking", _bk("pet"))),
        Col("service", "Service", 22, "text",
            formula=blank("booking", _bk("service"))),
        Col("staff", "Carer", 15, "text", validate=list_from("StaffNames")),
        Col("sched_start", "Scheduled", 9, "time",
            formula=blank("booking", _bk("start", '""'))),
        Col("arrive", "Arrived", 9, "time"),
        Col("depart", "Departed", 9, "time"),
        Col("duration", "Minutes on site", 10, "int",
            formula='=IF(OR({arrive}{r}="",{depart}{r}=""),"",'
                    'ROUND(({depart}{r}-{arrive}{r}'
                    '+IF({depart}{r}<{arrive}{r},1,0))*1440,0))'),
        Col("punctual", "On time?", 13, "status",
            formula='=IF(OR({arrive}{r}="",{sched_start}{r}=""),"",'
                    'IF({arrive}{r}<={sched_start}{r}+TIME(0,10,0),'
                    '"On time","Late "'
                    '&TEXT(({arrive}{r}-{sched_start}{r})*1440,"0")&" min"))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "On time", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Late", "fmt": "warn"}]),
        Col("walk_start", "Walk start", 9, "time"),
        Col("walk_end", "Walk end", 9, "time"),
        Col("walk_mins", "Walk minutes", 9, "int",
            formula='=IF(OR({walk_start}{r}="",{walk_end}{r}=""),"",'
                    'ROUND(({walk_end}{r}-{walk_start}{r}'
                    '+IF({walk_end}{r}<{walk_start}{r},1,0))*1440,0))'),
        Col("walk_distance", "Distance walked", 10, "number1",
            validate=number_check(0, 100)),
        Col("fed", "Fed", 8, "flag", validate=list_from("YesNoNAList"),
            cf=status_cf()),
        Col("water", "Fresh water", 9, "flag",
            validate=list_from("YesNoNAList"), cf=status_cf()),
        Col("meds", "Medication given", 10, "flag",
            validate=list_from("YesNoNAList"), cf=status_cf()),
        Col("potty", "Potty / litter", 14, "text",
            validate=list_from("PottyList")),
        Col("play", "Exercise & play", 22, "text"),
        Col("grooming", "Grooming / cleaning", 20, "text"),
        Col("mood", "Mood & behaviour", 14, "text",
            validate=list_from("MoodList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Unwell", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Anxious", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Aggressive", "fmt": "bad"}]),
        Col("appetite", "Appetite", 12, "text",
            validate=list_from("AppetiteList"),
            cf=[{"type": "text", "criteria": "containing", "value": "None",
                 "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Reduced", "fmt": "warn"}]),
        Col("health", "Health observations", 30, "wrap"),
        Col("incident", "Incident?", 9, "flag", validate=yes_no(),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"}]),
        Col("incident_ref", "Incident ref", 11, "text"),
        Col("photos", "Photos / report link", 20, "text"),
        Col("update_sent", "Client updated", 10, "flag", validate=yes_no(),
            cf=status_cf()),
        Col("checklist", "Care checklist done", 10, "flag",
            validate=yes_no(), cf=status_cf()),
        Col("status", "Visit status", 15, "status",
            validate=list_from("VisitStatusList"), cf=status_cf()),
        Col("notes", "Visit notes", 36, "wrap"),
        month_col("date"),
        # Hidden helper keys used by the printable report sheets: they
        # number each visit within its booking / pet so a one page report
        # can pull "line 1, line 2, ..." without array formulas.
        Col("line_key", "Line key", 10, "text", hidden=True,
            formula='=IF({booking}{r}="","",{booking}{r}&"|"'
                    '&COUNTIF(${booking}${first}:${booking}{r},'
                    '{booking}{r}))'),
        Col("pet_key", "Pet key", 10, "text", hidden=True,
            formula='=IF({pet}{r}="","",{pet}{r}&"|"'
                    '&COUNTIF(${pet}${first}:${pet}{r},{pet}{r}))'),
        Col("alert", "Action needed", 24, "status",
            formula=blank(
                "date",
                'IF({status}{r}="Missed","Missed visit - contact client",'
                'IF({incident}{r}="Yes","Incident logged - follow up",'
                'IF(AND({status}{r}="Completed",{update_sent}{r}<>"Yes"),'
                '"Action - send client update",'
                'IF(AND({status}{r}="Completed",{checklist}{r}<>"Yes"),'
                '"Check care checklist",'
                'IF(AND({date}{r}<TODAY(),{status}{r}="Scheduled"),'
                '"Overdue - log the result","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Daily Visit & Pet Care Log",
        subtitle="Proof of service for every walk, drop-in and overnight: "
                 "times, care given, mood, health notes and client updates.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=3,
        samples=sample(cfg).visits,
        help_text="Start with the Booking ID - the client, pet, service "
                  "and scheduled time are pulled in automatically. "
                  "Minutes on site, punctuality and alerts are calculated.",
    )
````

### `pet_sitting_spreadsheet/sheets/health.py`

*85 lines · 1 function*

> `Health Log` - medication administration, vaccinations and vet events.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Health Log').

````python
"""`Health Log` - medication administration, vaccinations and vet events."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, yes_no
from .common import alert_cf, blank, id_col, month_col, status_cf

NAME = "Health Log"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"HL-"'),
        Col("date", "Date", 12, "date", validate=date_check()),
        Col("time", "Time", 9, "time"),
        Col("pet", "Pet", 24, "text", validate=list_from("PetLabels")),
        Col("client", "Client", 20, "text",
            formula=blank("pet",
                          'IFERROR(INDEX([[Pets.owner]],'
                          'MATCH({pet}{r},[[Pets.label]],0)),"")')),
        Col("type", "Record type", 16, "text",
            validate=list_from("HealthRecordTypeList")),
        Col("item", "Medication / treatment", 24, "text"),
        Col("dosage", "Dosage & instructions", 28, "wrap"),
        Col("sched_time", "Scheduled", 9, "time"),
        Col("given_time", "Given at", 9, "time"),
        Col("staff", "Given by", 15, "text", validate=list_from("StaffNames")),
        Col("status", "Status", 13, "status",
            validate=list_from("MedStatusList"), cf=status_cf()),
        Col("supply", "Supply left", 10, "number1",
            validate=number_check(0, 10000),
            note="Doses, tablets or ml left in the client's supply."),
        Col("refill_due", "Refill due", 12, "date", validate=date_check()),
        Col("refill_alert", "Refill alert", 16, "status",
            formula=blank("refill_due",
                          'IF({refill_due}{r}<TODAY(),"Overdue - refill now",'
                          'IF({refill_due}{r}-TODAY()<=ExpiryAlertDays,'
                          '"Due in "&TEXT({refill_due}{r}-TODAY(),"0")'
                          '&" days","OK"))'),
            cf=alert_cf()),
        Col("vet_appt", "Vet appointment", 12, "date",
            validate=date_check()),
        Col("appt_alert", "Appointment alert", 16, "status",
            formula=blank("vet_appt",
                          'IF({vet_appt}{r}<TODAY(),"Past",'
                          'IF({vet_appt}{r}-TODAY()<=ExpiryAlertDays,'
                          '"Due in "&TEXT({vet_appt}{r}-TODAY(),"0")'
                          '&" days","Booked"))'),
            cf=alert_cf() + [{"type": "text", "criteria": "containing",
                              "value": "Booked", "fmt": "good"}]),
        Col("cost", "Cost", 11, "money", validate=number_check(0, 100000),
            note="Vet bills or treatment costs you paid on behalf of the "
                 "client - recharge them on the invoice."),
        Col("notified", "Owner notified", 10, "flag", validate=yes_no(),
            cf=status_cf()),
        Col("notes", "Notes", 34, "wrap"),
        month_col("date"),
        Col("alert", "Action needed", 24, "status",
            formula=blank(
                "date",
                'IF({status}{r}="Missed","Missed dose - tell the owner",'
                'IF({status}{r}="Refused","Refused - monitor and report",'
                'IF(AND({status}{r}="Scheduled",{date}{r}<TODAY()),'
                '"Overdue - not recorded",'
                'IF(AND({notified}{r}<>"Yes",{type}{r}="Vet visit"),'
                '"Action - update the owner","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Medication & Health Monitoring",
        subtitle="A dated record of every dose given, vaccination, vet "
                 "visit and health observation - your proof of care.",
        columns=columns,
        rows=cfg.row_count("Health"),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).health,
        help_text="Log every dose, even missed ones. The alerts column "
                  "tells you what still needs an owner conversation.",
    )
````

### `pet_sitting_spreadsheet/sheets/incidents.py`

*137 lines · 2 functions*

> `Incidents` - safety, emergency contacts and the incident report register.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg)` | Write the emergency contact directory above the register |

**Module constants:** `NAME` ('Incidents'), `DIRECTORY_HEADER_ROW` (5), `REGISTER_HEADER_ROW` (17), `DIRECTORY` (9 items).

````python
"""`Incidents` - safety, emergency contacts and the incident report register."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, yes_no
from .common import alert_cf, blank, id_col, status_cf

NAME = "Incidents"

#: The emergency directory block sits above the register.
DIRECTORY_HEADER_ROW = 5          # Excel row of the directory header
REGISTER_HEADER_ROW = 17          # Excel row of the register header

DIRECTORY = [
    ("Business emergency line", "=Biz_Phone", "Answered 07:00 - 21:00"),
    ("Owner / manager", "=Biz_Owner", "Call for any High or Critical event"),
    ("24/7 emergency vet", "24/7 Springfield Pet ER",
     "+1 (555) 100 9111 - open all night"),
    ("Backup emergency vet", "Riverton Vet Clinic",
     "+1 (555) 100 2000 - 08:00 to 18:00"),
    ("Animal poison helpline", "Pet Poison Hotline", "+1 (555) 100 7788"),
    ("Animal control / council", "Springfield Animal Services",
     "+1 (555) 100 4422"),
    ("Insurance claims", "=Biz_Insurer", "=Biz_Policy"),
    ("Backup sitter on call", "Priya Shah", "+1 (555) 772 6690"),
    ("Locksmith / key issues", "Springfield Locks", "+1 (555) 100 6655"),
]


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"IN-"', digits=3, width=9),
        Col("date", "Date", 12, "date", validate=date_check()),
        Col("time", "Time", 9, "time"),
        Col("client", "Client", 20, "text",
            validate=list_from("ClientNames")),
        Col("pet", "Pet", 22, "text", validate=list_from("PetLabels")),
        Col("staff", "Reported by", 15, "text",
            validate=list_from("StaffNames")),
        Col("ref", "Booking / visit ref", 13, "text"),
        Col("type", "Incident type", 16, "text",
            validate=list_from("IncidentTypeList")),
        Col("severity", "Severity", 11, "status",
            validate=list_from("SeverityList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Critical", "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing", "value": "High",
                 "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Medium", "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "Low",
                 "fmt": "good"}]),
        Col("description", "What happened", 44, "wrap"),
        Col("action", "Action taken", 40, "wrap"),
        Col("vet_contacted", "Vet contacted", 10, "flag", validate=yes_no(),
            cf=status_cf()),
        Col("clinic", "Clinic attended", 20, "text"),
        Col("owner_notified", "Owner notified", 10, "flag",
            validate=yes_no(), cf=status_cf()),
        Col("notified_at", "Notified when", 16, "text"),
        Col("follow_up", "Follow-up needed", 10, "flag", validate=yes_no(),
            cf=status_cf()),
        Col("follow_due", "Follow-up due", 12, "date",
            validate=date_check()),
        Col("status", "Status", 13, "status",
            validate=list_from("IncidentStatusList"), cf=status_cf()),
        Col("cost", "Cost incurred", 11, "money",
            validate=number_check(0, 100000)),
        Col("claim", "Insurance claim", 10, "flag", validate=yes_no()),
        Col("days_open", "Days open", 9, "text",
            formula=blank("date",
                          'IF(OR({status}{r}="Resolved",'
                          '{status}{r}="Closed"),"Closed",'
                          'TEXT(TODAY()-{date}{r},"0")&" days")')),
        Col("notes", "Notes & outcome", 34, "wrap"),
        Col("alert", "Action needed", 24, "status",
            formula=blank(
                "date",
                'IF({owner_notified}{r}<>"Yes","Action - notify the owner",'
                'IF(AND({follow_up}{r}="Yes",{follow_due}{r}<>"",'
                '{follow_due}{r}<TODAY(),{status}{r}<>"Resolved",'
                '{status}{r}<>"Closed"),"Overdue follow-up",'
                'IF(OR({status}{r}="Open",{status}{r}="In progress"),'
                '"Open - keep working it","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Pet Safety, Emergencies & Incident Reports",
        subtitle="Emergency numbers at the top, a dated incident report "
                 "register below. Print a row straight from this tab if "
                 "an owner or insurer asks for it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        header_row=REGISTER_HEADER_ROW,
        tab_color=Palette.bad_text,
        freeze_cols=2,
        freeze_at=3,
        samples=sample(cfg).incidents,
        help_text="Every incident needs: what happened, what you did, who "
                  "you told and when. The alert column chases anything "
                  "still open.",
    )


def decorate(sheet, fmts, cfg: Config) -> None:
    """Write the emergency contact directory above the register."""
    row = DIRECTORY_HEADER_ROW - 2  # 0 based row for the section band
    sheet.section(row, "Emergency contact directory - keep this current",
                  7)
    head = DIRECTORY_HEADER_ROW - 1
    sheet.write(head, 0, "Contact", fmts.chrome("header_left"))
    sheet.merge(head, 1, 2, "Name", fmts.chrome("header"))
    sheet.merge(head, 3, 6, "Number / details", fmts.chrome("header"))
    sheet.write(head, 7, "Notes", fmts.chrome("header"))
    sheet.ws.set_row(head, 20)

    label_fmt = fmts.chrome("label_bold")
    value_fmt = fmts.input("text")
    for i, (label, name, detail) in enumerate(DIRECTORY):
        r = head + 1 + i
        sheet.write(r, 0, label, label_fmt)
        if isinstance(name, str) and name.startswith("="):
            sheet.ws.merge_range(r, 1, r, 2, "", fmts.calc("text"))
            sheet.ws.write_formula(r, 1, name, fmts.calc("text"))
        else:
            sheet.merge(r, 1, 2, name, value_fmt)
        if isinstance(detail, str) and detail.startswith("="):
            sheet.ws.merge_range(r, 3, r, 6, "", fmts.calc("text"))
            sheet.ws.write_formula(r, 3, detail, fmts.calc("text"))
        else:
            sheet.merge(r, 3, 6, detail, value_fmt)
        sheet.ws.write_blank(r, 7, None, value_fmt)
````

### `pet_sitting_spreadsheet/sheets/messages.py`

*85 lines · 1 function*

> `Messages` - client communication, feedback, complaints and follow-ups.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Messages').

````python
"""`Messages` - client communication, feedback, complaints and follow-ups."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, yes_no
from .common import alert_cf, blank, id_col, month_col, rating_validation, \
    status_cf

NAME = "Messages"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"MSG-"'),
        Col("date", "Date", 12, "date", validate=date_check()),
        Col("client", "Client", 20, "text",
            validate=list_from("ClientNames")),
        Col("type", "Message type", 20, "text",
            validate=list_from("MessageTypeList")),
        Col("channel", "Channel", 14, "text",
            validate=list_from("ChannelList")),
        Col("direction", "Direction", 11, "text",
            validate=list_from("DirectionList")),
        Col("summary", "Summary", 40, "wrap"),
        Col("response", "Reply received", 10, "flag", validate=yes_no(),
            cf=status_cf()),
        Col("rating", "Satisfaction (1-5)", 10, "rating",
            validate=rating_validation(),
            cf=[{"type": "cell", "criteria": ">=", "value": 4,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<=", "value": 2,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "=", "value": 3,
                 "fmt": "warn"}]),
        Col("feedback", "Feedback / review", 40, "wrap"),
        Col("complaint", "Complaint?", 9, "flag", validate=yes_no(),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"}]),
        Col("issue", "Issue raised", 32, "wrap"),
        Col("resolution", "Resolution status", 15, "status",
            validate=list_from("ResolutionList"), cf=status_cf()),
        Col("resolved_date", "Resolved on", 12, "date",
            validate=date_check()),
        Col("follow_due", "Follow-up due", 12, "date",
            validate=date_check()),
        Col("follow_done", "Follow-up done", 10, "flag", validate=yes_no(),
            cf=status_cf()),
        Col("days_open", "Days open", 9, "int",
            formula='=IF(AND({complaint}{r}="Yes",'
                    '{resolution}{r}<>"Resolved",{date}{r}<>""),'
                    'TODAY()-{date}{r},"")',
            cf=[{"type": "cell", "criteria": ">", "value": 7,
                 "fmt": "bad"}]),
        Col("notes", "Notes", 30, "wrap"),
        month_col("date"),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "date",
                'IF(AND({complaint}{r}="Yes",{resolution}{r}<>"Resolved"),'
                '"Action - unresolved complaint",'
                'IF(AND({follow_due}{r}<>"",{follow_due}{r}<TODAY(),'
                '{follow_done}{r}<>"Yes"),"Overdue follow-up",'
                'IF(AND({follow_due}{r}<>"",'
                '{follow_due}{r}-TODAY()<=ReminderLeadDays,'
                '{follow_done}{r}<>"Yes"),"Follow-up due soon","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Client Communication, Feedback & Complaints",
        subtitle="Every confirmation, reminder, update, review and "
                 "complaint in one searchable history.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=3,
        samples=sample(cfg).messages,
        help_text="Log the message the moment you send it. Satisfaction "
                  "ratings feed the customer-experience numbers on the "
                  "Reports tab.",
    )
````

### `pet_sitting_spreadsheet/sheets/invoices.py`

*158 lines · 1 function*

> `Invoices` - invoice register with automatic numbering, tax and ageing.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Invoices').

````python
"""`Invoices` - invoice register with automatic numbering, tax and ageing."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import DATA_ROW, Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, yes_no
from .common import alert_cf, blank, dup_cf, month_col, owing_cf, status_cf

NAME = "Invoices"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("invoice_no", "Invoice no.", 12, "id",
            formula=blank("client",
                          'InvoicePrefix&TEXT(InvoiceStartNo+ROW()-'
                          f'{DATA_ROW},"0000")'),
            note="Numbered automatically from the first invoice number on "
                 "the Settings tab."),
        Col("date", "Invoice date", 12, "date", validate=date_check()),
        Col("due", "Due date", 12, "date",
            formula=blank("date", '{date}{r}+PaymentTermsDays')),
        Col("client", "Client", 20, "text",
            validate=list_from("ClientNames")),
        Col("booking", "Booking ref", 12, "text",
            validate=list_from("BookingIDs"),
            note="Link a booking and the invoice total is taken straight "
                 "from it. Leave blank to bill manually with Qty x Rate."),
        Col("booking_total", "Booking total", 12, "money",
            # The booking subtotal is used (not the booking total) because
            # this invoice adds the tax itself a few columns to the right.
            formula='=IF({booking}{r}="",0,IFERROR('
                    'INDEX([[Bookings.subtotal]],'
                    'MATCH({booking}{r},[[Bookings.id]],0)),0))'),
        Col("description", "Description of service", 40, "wrap"),
        Col("qty", "Qty", 8, "number", validate=number_check(0, 10000)),
        Col("rate", "Rate", 11, "money", validate=number_check(0, 100000)),
        Col("line_total", "Line amount", 12, "money",
            formula=blank("client",
                          'IF({booking}{r}<>"",{booking_total}{r},'
                          'ROUND(IF({qty}{r}="",0,{qty}{r})'
                          '*IF({rate}{r}="",0,{rate}{r}),2))')),
        Col("fees", "Additional fees", 11, "money",
            validate=number_check(0, 100000)),
        Col("discount", "Discount", 11, "money",
            validate=number_check(0, 100000)),
        Col("subtotal", "Subtotal", 12, "money",
            formula=blank("client",
                          'ROUND({line_total}{r}+IF({fees}{r}="",0,'
                          '{fees}{r})-IF({discount}{r}="",0,'
                          '{discount}{r}),2)')),
        Col("tax", "Tax", 11, "money",
            formula='=IF({subtotal}{r}="","",IF(TaxRegistered<>"Yes",0,'
                    'IF(PricesIncludeTax="Yes",'
                    'ROUND({subtotal}{r}-{subtotal}{r}/(1+TaxRate),2),'
                    'ROUND({subtotal}{r}*TaxRate,2))))'),
        Col("total", "Invoice total", 13, "money",
            formula='=IF({subtotal}{r}="","",IF(PricesIncludeTax="Yes",'
                    '{subtotal}{r},ROUND({subtotal}{r}+{tax}{r},2)))',
            cf=[{"type": "data_bar", "bar_color": "#2E7D8F"}]),
        Col("deposit", "Deposit received", 12, "money",
            formula=blank("invoice_no",
                          'SUMIFS([[Payments.amount]],'
                          '[[Payments.invoice_no]],{invoice_no}{r},'
                          '[[Payments.deposit]],"Yes")'),
            note="Log deposits on the Payments tab with Deposit = Yes."),
        Col("paid", "Amount paid", 12, "money",
            formula=blank("invoice_no",
                          'SUMIFS([[Payments.net]],'
                          '[[Payments.invoice_no]],{invoice_no}{r})')),
        Col("refunded", "Refunded", 11, "money",
            formula=blank("invoice_no",
                          'SUMIFS([[Payments.refund]],'
                          '[[Payments.invoice_no]],{invoice_no}{r})')),
        Col("balance", "Balance due", 12, "money",
            formula=blank("client",
                          'ROUND(IF({status}{r}="Cancelled",0,'
                          '{total}{r}-{paid}{r}),2)'),
            cf=owing_cf()),
        Col("status", "Invoice status", 13, "status",
            validate=list_from("InvoiceStatusList"), cf=status_cf(),
            note="You control Draft / Sent / Cancelled. The payment status "
                 "next to it is calculated from the Payments tab."),
        Col("pay_status", "Payment status", 15, "status",
            formula=blank(
                "invoice_no",
                'IF({status}{r}="Cancelled","Cancelled",'
                'IF({balance}{r}<=0.004,"Paid",'
                'IF({paid}{r}>0.004,IF(AND({due}{r}<>"",TODAY()>{due}{r}),'
                '"Overdue - part paid","Partially paid"),'
                'IF(AND({due}{r}<>"",TODAY()>{due}{r}),"Overdue",'
                'IF({status}{r}="Draft","Draft","Unpaid")))))'),
            cf=status_cf()),
        Col("days_overdue", "Days overdue", 10, "int",
            formula='=IF(OR({due}{r}="",{balance}{r}<=0.004),"",'
                    'MAX(0,TODAY()-{due}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}]),
        Col("aging", "Ageing bucket", 13, "status",
            formula=blank("invoice_no",
                          'IF({balance}{r}<=0.004,"Paid",'
                          'IF({days_overdue}{r}="","Not due",'
                          'IF({days_overdue}{r}=0,"Current",'
                          'IF({days_overdue}{r}<=30,"1-30 days",'
                          'IF({days_overdue}{r}<=60,"31-60 days",'
                          'IF({days_overdue}{r}<=90,"61-90 days",'
                          '"90+ days"))))))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Paid",
                 "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "90+", "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "61-90", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "31-60", "fmt": "warn"}]),
        Col("late_fee", "Suggested late fee", 11, "money",
            formula='=IF(OR({days_overdue}{r}="",{days_overdue}{r}=0),0,'
                    'ROUND({balance}{r}*LateFeePct,2))'),
        Col("sent_date", "Sent on", 12, "date", validate=date_check()),
        Col("reminder_sent", "Reminder sent", 10, "flag", validate=yes_no(),
            cf=status_cf()),
        Col("dup_check", "Number check", 13, "status",
            formula=blank("invoice_no",
                          'IF(COUNTIF([[Invoices.invoice_no]],'
                          '{invoice_no}{r})>1,"Duplicate number","OK")'),
            cf=dup_cf()),
        Col("notes", "Invoice notes", 30, "wrap"),
        month_col("date"),
        Col("alert", "Action needed", 24, "status",
            formula=blank(
                "invoice_no",
                'IF({status}{r}="Draft","Action - still a draft",'
                'IF({pay_status}{r}="Overdue","Overdue - chase payment",'
                'IF({pay_status}{r}="Overdue - part paid",'
                '"Overdue - chase the balance",'
                'IF(AND({balance}{r}>0.004,{due}{r}<>"",'
                '{due}{r}-TODAY()<=ReminderLeadDays,{due}{r}>=TODAY()),'
                '"Due soon - send a reminder","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Invoices",
        subtitle="Automatic numbering, due dates, tax, balances, ageing "
                 "buckets and overdue alerts - all from the booking you "
                 "link.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.good_text,
        freeze_cols=2,
        samples=sample(cfg).invoices,
        help_text="Type the client name to create the invoice number. "
                  "Link a Booking ID to bill the calculated booking total, "
                  "or use Qty x Rate for anything custom.",
    )
````

### `pet_sitting_spreadsheet/sheets/payments.py`

*89 lines · 2 functions*

> `Payments` - money received, deposits, tips, refunds and reconciliation.

| Defines | What it does |
|---------|--------------|
| `_inv(field, fallback)` |  |
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Payments').

````python
"""`Payments` - money received, deposits, tips, refunds and reconciliation."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, yes_no
from .common import alert_cf, blank, id_col, month_col, status_cf

NAME = "Payments"


def _inv(field: str, fallback: str = '""') -> str:
    return (f'IFERROR(INDEX([[Invoices.{field}]],'
            f'MATCH({{invoice_no}}{{r}},[[Invoices.invoice_no]],0)),'
            f'{fallback})')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"PM-"'),
        Col("date", "Payment date", 12, "date", validate=date_check()),
        Col("client", "Client", 20, "text",
            validate=list_from("ClientNames")),
        Col("invoice_no", "Invoice no.", 12, "text",
            validate=list_from("InvoiceNos"),
            note="Match the payment to an invoice so balances update "
                 "automatically. Leave blank for an unlinked payment."),
        Col("amount", "Amount applied", 12, "money",
            validate=number_check(0, 1000000),
            note="The part of the payment that pays down the invoice."),
        Col("method", "Payment method", 16, "text",
            validate=list_from("PaymentMethodList")),
        Col("reference", "Reference", 14, "text"),
        Col("tip", "Tip / gratuity", 10, "money",
            validate=number_check(0, 100000)),
        Col("late_fee", "Late fee collected", 10, "money",
            validate=number_check(0, 100000)),
        Col("refund", "Refund given", 11, "money",
            validate=number_check(0, 1000000)),
        Col("net", "Applied to invoice", 12, "money",
            formula=blank("date",
                          'ROUND(IF({amount}{r}="",0,{amount}{r})'
                          '-IF({refund}{r}="",0,{refund}{r}),2)')),
        Col("received", "Total cash in", 12, "money",
            formula=blank("date",
                          'ROUND(IF({amount}{r}="",0,{amount}{r})'
                          '+IF({tip}{r}="",0,{tip}{r})'
                          '+IF({late_fee}{r}="",0,{late_fee}{r})'
                          '-IF({refund}{r}="",0,{refund}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": "#2E7D8F"}]),
        Col("deposit", "Deposit?", 9, "flag", validate=yes_no()),
        Col("confirmed", "Cleared / confirmed", 10, "flag",
            validate=yes_no(), cf=status_cf()),
        Col("reconciled", "Reconciled to bank", 10, "flag",
            validate=yes_no(), cf=status_cf()),
        Col("inv_total", "Invoice total", 12, "money",
            formula=blank("invoice_no", _inv("total", "0"))),
        Col("inv_balance", "Invoice balance", 12, "money",
            formula=blank("invoice_no", _inv("balance", "0")),
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "warn"}]),
        Col("notes", "Notes", 30, "wrap"),
        month_col("date"),
        Col("alert", "Action needed", 22, "status",
            formula=blank(
                "date",
                'IF({confirmed}{r}<>"Yes","Unconfirmed - check it cleared",'
                'IF({reconciled}{r}<>"Yes","To reconcile",'
                'IF(AND({invoice_no}{r}<>"",{inv_balance}{r}>0.004),'
                '"Invoice still has a balance","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Payment Tracking",
        subtitle="Every cash, card, transfer, wallet and deposit payment, "
                 "with tips, refunds and bank reconciliation.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.good_text,
        freeze_cols=2,
        samples=sample(cfg).payments,
        help_text="Record payments the day they land. 'Applied to invoice' "
                  "clears the balance; 'Total cash in' (tips and late fees "
                  "included) is what the dashboard counts as revenue.",
    )
````

### `pet_sitting_spreadsheet/sheets/expenses.py`

*87 lines · 1 function*

> `Expenses` - every business cost, with tax, deductibility and approvals.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Expenses').

````python
"""`Expenses` - every business cost, with tax, deductibility and approvals."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, month_col, status_cf

NAME = "Expenses"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"EX-"'),
        Col("date", "Date", 12, "date", validate=date_check()),
        Col("category", "Category", 22, "text",
            validate=list_from("ExpenseCategoryList")),
        Col("description", "Description", 34, "wrap"),
        Col("supplier", "Supplier / vendor", 20, "text"),
        Col("amount", "Amount paid", 12, "money",
            validate=number_check(0, 1000000),
            note="Enter the full amount you paid, including tax."),
        Col("tax", "Tax included", 11, "money",
            formula=blank("amount",
                          'IF(TaxRegistered<>"Yes",0,'
                          'ROUND({amount}{r}-{amount}{r}/(1+TaxRate),2))')),
        Col("net", "Net cost", 11, "money",
            formula=blank("amount", 'ROUND({amount}{r}-{tax}{r},2)')),
        Col("method", "Paid by", 16, "text",
            validate=list_from("PaymentMethodList")),
        Col("type", "Business / personal", 12, "status",
            validate=list_from("BusinessPersonalList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Personal", "fmt": "muted"}]),
        Col("deductible", "Business use %", 10, "percent",
            validate=number_check(0, 1),
            note="Blank counts as 100%. Use 50% for a phone bill you share "
                 "with personal use."),
        Col("claim", "Claimable amount", 12, "money",
            formula=blank("amount",
                          'IF({type}{r}="Personal",0,'
                          'ROUND({net}{r}*IF({deductible}{r}="",1,'
                          '{deductible}{r}),2))')),
        Col("receipt", "Receipt link / ref", 18, "text"),
        Col("recurring", "Recurring", 12, "text",
            validate=list_from("RecurringExpenseList")),
        Col("next_due", "Next due", 12, "date",
            formula='=IF(OR({date}{r}="",{recurring}{r}="",'
                    '{recurring}{r}="No"),"",'
                    'IF({recurring}{r}="Monthly",EDATE({date}{r},1),'
                    'IF({recurring}{r}="Quarterly",EDATE({date}{r},3),'
                    'IF({recurring}{r}="Annually",EDATE({date}{r},12),'
                    '{date}{r}+IF({recurring}{r}="Weekly",7,'
                    'IF({recurring}{r}="Fortnightly",14,0))))))'),
        Col("approval", "Approval", 12, "status",
            validate=list_from("ExpenseApprovalList"), cf=status_cf()),
        Col("related", "Related client / booking", 18, "text"),
        Col("staff", "Staff member", 15, "text",
            validate=list_from("StaffNames")),
        Col("notes", "Notes", 26, "wrap"),
        month_col("date"),
        Col("alert", "Action needed", 22, "status",
            formula=blank(
                "date",
                'IF({approval}{r}="Pending","Awaiting approval",'
                'IF({receipt}{r}="","Missing receipt",'
                'IF(AND({next_due}{r}<>"",{next_due}{r}<=TODAY()),'
                '"Recurring cost due","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Expense Management",
        subtitle="Fuel, supplies, wages, insurance, software and every "
                 "other cost - split into tax, net and claimable amounts.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.bad_text,
        freeze_cols=2,
        samples=sample(cfg).expenses,
        help_text="Keep a photo of every receipt. The Reports tab groups "
                  "these by category and the Budget tab compares them to "
                  "your monthly budget.",
    )
````

### `pet_sitting_spreadsheet/sheets/mileage.py`

*96 lines · 1 function*

> `Mileage` - trip log with odometer readings and reimbursement maths.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Mileage').

````python
"""`Mileage` - trip log with odometer readings and reimbursement maths."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, yes_no
from .common import alert_cf, blank, id_col, month_col

NAME = "Mileage"


def spec(cfg: Config) -> SheetSpec:
    unit = cfg.distance_unit
    columns = [
        id_col("date", '"TR-"'),
        Col("date", "Trip date", 12, "date", validate=date_check()),
        Col("staff", "Driver", 15, "text", validate=list_from("StaffNames")),
        Col("client", "Client / booking", 20, "text",
            validate=list_from("ClientNames")),
        Col("purpose", "Purpose", 16, "text",
            validate=list_from("TripPurposeList")),
        Col("start_loc", "From", 20, "text"),
        Col("end_loc", "To", 20, "text"),
        Col("start_odo", "Odometer start", 12, "number1",
            validate=number_check(0, 10000000)),
        Col("end_odo", "Odometer end", 12, "number1",
            validate=number_check(0, 10000000)),
        Col("manual_km", f"Manual {unit}", 10, "number1",
            validate=number_check(0, 10000),
            note="Only needed when you did not record odometer readings."),
        Col("trip_km", f"Trip {unit}", 10, "number1",
            formula='=IF(AND({start_odo}{r}<>"",{end_odo}{r}<>""),'
                    'ROUND({end_odo}{r}-{start_odo}{r},1),'
                    'IF({manual_km}{r}="","",{manual_km}{r}))'),
        Col("round_trip", "Round trip?", 9, "flag", validate=yes_no()),
        Col("total_km", f"Total {unit}", 10, "number1",
            formula='=IF({trip_km}{r}="","",IF({round_trip}{r}="Yes",'
                    'ROUND({trip_km}{r}*2,1),{trip_km}{r}))',
            cf=[{"type": "data_bar", "bar_color": "#2E7D8F"}]),
        Col("business", "Classification", 12, "status",
            formula=blank("purpose",
                          'IF({purpose}{r}="Personal","Personal",'
                          '"Business")'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Business", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Personal", "fmt": "muted"}]),
        Col("rate", "Rate", 10, "money",
            formula=blank("date", "MileageRate")),
        Col("reimbursement", "Mileage claim", 12, "money",
            formula='=IF({total_km}{r}="","",IF({business}{r}="Business",'
                    'ROUND({total_km}{r}*MileageRate,2),0))'),
        Col("fuel", "Fuel cost", 10, "money",
            validate=number_check(0, 100000),
            note="For reference only - the mileage rate already covers "
                 "fuel. Log fuel receipts on the Expenses tab if you claim "
                 "actual costs instead."),
        Col("parking", "Parking", 10, "money",
            validate=number_check(0, 100000)),
        Col("tolls", "Tolls", 10, "money", validate=number_check(0, 100000)),
        Col("travel_mins", "Travel time (min)", 10, "int",
            validate=number_check(0, 1440, decimal=False)),
        Col("claimable", "Total claimable", 12, "money",
            formula=blank("date",
                          'ROUND(IF({reimbursement}{r}="",0,'
                          '{reimbursement}{r})+IF({business}{r}="Business",'
                          'IF({parking}{r}="",0,{parking}{r})'
                          '+IF({tolls}{r}="",0,{tolls}{r}),0),2)')),
        Col("notes", "Notes", 26, "wrap"),
        month_col("date"),
        Col("alert", "Check", 20, "status",
            formula=blank(
                "date",
                'IF(AND({start_odo}{r}<>"",{end_odo}{r}<>"",'
                '{end_odo}{r}<{start_odo}{r}),"Check odometer readings",'
                'IF({trip_km}{r}="","Action - add distance",'
                'IF({purpose}{r}="","Action - add a purpose","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Mileage & Travel Log",
        subtitle=f"Odometer in, odometer out: {unit} travelled, round "
                 "trips, reimbursement, parking and tolls - ready for your "
                 "tax return.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).mileage,
        help_text="Record trips as you do them. Business trips are valued "
                  "at the mileage rate from the Settings tab.",
    )
````

### `pet_sitting_spreadsheet/sheets/staff.py`

*120 lines · 1 function*

> `Staff` - team directory, workload, timesheet totals and pay.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Staff').

````python
"""`Staff` - team directory, workload, timesheet totals and pay."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, status_cf

NAME = "Staff"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"ST-"', digits=3, width=9),
        Col("name", "Name", 18, "text",
            note="This name is used by the Bookings, Visit Log and Mileage "
                 "drop downs."),
        Col("phone", "Phone", 17, "text"),
        Col("email", "Email", 26, "text"),
        Col("role", "Role", 14, "text", validate=list_from("StaffRoleList")),
        Col("employment", "Employment type", 18, "text",
            validate=list_from("EmploymentTypeList")),
        Col("start_date", "Started on", 12, "date", validate=date_check()),
        Col("status", "Status", 11, "status",
            validate=list_from("StaffStatusList"), cf=status_cf()),
        Col("availability", "Availability", 22, "text"),
        Col("days_off", "Regular days off", 14, "text"),
        Col("pay_type", "Pay type", 13, "text",
            validate=list_from("PayTypeList")),
        Col("pay_rate", "Pay rate", 11, "money",
            validate=number_check(0, 100000),
            note="Hourly rate, per-visit fee, or the commission "
                 "percentage as a number (55 = 55%)."),
        Col("certs", "Certifications", 26, "wrap"),
        Col("cert_expiry", "Certificate expires", 12, "date",
            validate=date_check()),
        Col("cert_alert", "Certificate status", 16, "status",
            formula=blank("cert_expiry",
                          'IF({cert_expiry}{r}<TODAY(),"EXPIRED",'
                          'IF({cert_expiry}{r}-TODAY()<=ExpiryAlertDays,'
                          '"Due in "&TEXT({cert_expiry}{r}-TODAY(),"0")'
                          '&" days","Valid"))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "EXPIRED", "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Due in", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Valid", "fmt": "good"}]),
        Col("check_date", "Background check", 12, "date",
            validate=date_check()),
        Col("emergency", "Emergency contact", 24, "text"),
        Col("leave", "Leave & absence", 22, "wrap"),
        Col("bookings", "Bookings assigned", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Bookings.staff]],{name}{r},'
                          '[[Bookings.status]],"<>Cancelled")')),
        Col("visits", "Visits completed", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Visit Log.staff]],{name}{r},'
                          '[[Visit Log.status]],"Completed")')),
        Col("hours", "Hours worked", 10, "number",
            formula=blank("name",
                          'ROUND(SUMIFS([[Visit Log.duration]],'
                          '[[Visit Log.staff]],{name}{r},'
                          '[[Visit Log.status]],"Completed")/60,2)')),
        Col("distance", f"{cfg.distance_unit.title()} driven", 10, "number1",
            formula=blank("name",
                          'SUMIFS([[Mileage.total_km]],[[Mileage.staff]],'
                          '{name}{r},[[Mileage.business]],"Business")')),
        Col("revenue", "Revenue delivered", 13, "money",
            formula=blank("name",
                          'SUMIFS([[Bookings.total]],[[Bookings.staff]],'
                          '{name}{r},[[Bookings.status]],"<>Cancelled")'),
            cf=[{"type": "data_bar", "bar_color": "#2E7D8F"}]),
        Col("wages", "Pay due", 12, "money",
            formula=blank("name",
                          'ROUND(IF({pay_type}{r}="Hourly",{hours}{r}'
                          '*IF({pay_rate}{r}="",0,{pay_rate}{r}),'
                          'IF({pay_type}{r}="Per visit",{visits}{r}'
                          '*IF({pay_rate}{r}="",0,{pay_rate}{r}),'
                          'IF({pay_type}{r}="Commission %",{revenue}{r}'
                          '*IF({pay_rate}{r}="",0,{pay_rate}{r})/100,0))),2)'),
            note="Salaried team members show 0 - pay them from payroll."),
        Col("rev_per_hour", "Revenue per hour", 11, "money",
            formula='=IF(OR({hours}{r}="",{hours}{r}=0),"",'
                    'ROUND({revenue}{r}/{hours}{r},2))'),
        Col("completion", "Completion rate", 10, "percent",
            formula='=IF(OR({bookings}{r}="",{bookings}{r}=0),"",'
                    '{visits}{r}/{bookings}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": 0.8,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": 0.95,
                 "fmt": "good"}]),
        Col("notes", "Notes", 28, "wrap"),
        Col("alert", "Action needed", 22, "status",
            formula=blank(
                "name",
                'IF({cert_alert}{r}="EXPIRED","Certificate expired",'
                'IF(LEFT({cert_alert}{r},6)="Due in","Certificate due",'
                'IF({check_date}{r}="","Action - background check",'
                'IF(AND({status}{r}="Active",{bookings}{r}=0),'
                '"No bookings assigned","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff & Contractors",
        subtitle="Who works for you, when they are available, what they "
                 "have delivered and what they are owed.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).staff,
        help_text="Hours and pay come from completed visits in the Visit "
                  "Log, so log every visit to keep the timesheet honest.",
    )
````

### `pet_sitting_spreadsheet/sheets/inventory.py`

*99 lines · 1 function*

> `Inventory` - pet supplies, stock levels, reorder alerts and value.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Inventory').

````python
"""`Inventory` - pet supplies, stock levels, reorder alerts and value."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col

NAME = "Inventory"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("item", '"IT-"', digits=3, width=9),
        Col("item", "Item", 28, "text"),
        Col("category", "Category", 18, "text",
            validate=list_from("InventoryCategoryList")),
        Col("supplier", "Supplier", 20, "text"),
        Col("unit", "Unit", 10, "text"),
        Col("unit_cost", "Unit cost", 11, "money",
            validate=number_check(0, 100000)),
        Col("qty", "In stock", 10, "number1",
            validate=number_check(0, 100000)),
        Col("min_qty", "Minimum level", 10, "number1",
            validate=number_check(0, 100000),
            note="When stock drops to this level the item is flagged for "
                 "reorder on the Reminders tab."),
        Col("reorder", "Reorder qty", 10, "number1",
            validate=number_check(0, 100000)),
        Col("value", "Stock value", 12, "money",
            formula=blank("item",
                          'ROUND(IF({qty}{r}="",0,{qty}{r})'
                          '*IF({unit_cost}{r}="",0,{unit_cost}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": "#2E7D8F"}]),
        Col("stock_status", "Stock status", 14, "status",
            formula=blank("item",
                          'IF({qty}{r}="","Set stock level",'
                          'IF({qty}{r}<=0,"Out of stock",'
                          'IF({qty}{r}<=IF({min_qty}{r}="",0,'
                          '{min_qty}{r}),"Low - reorder","OK")))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Out of stock", "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing", "value": "Low",
                 "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "OK",
                 "fmt": "good"}]),
        Col("last_purchase", "Last purchased", 12, "date",
            validate=date_check()),
        Col("qty_purchased", "Qty last bought", 10, "number1",
            validate=number_check(0, 100000)),
        Col("purchase_cost", "Last purchase cost", 12, "money",
            formula=blank("item",
                          'ROUND(IF({qty_purchased}{r}="",0,'
                          '{qty_purchased}{r})*IF({unit_cost}{r}="",0,'
                          '{unit_cost}{r}),2)')),
        Col("expiry", "Expiry date", 12, "date", validate=date_check()),
        Col("expiry_alert", "Expiry status", 14, "status",
            formula=blank("expiry",
                          'IF({expiry}{r}<TODAY(),"EXPIRED",'
                          'IF({expiry}{r}-TODAY()<=ExpiryAlertDays,'
                          '"Use within "&TEXT({expiry}{r}-TODAY(),"0")'
                          '&" days","OK"))'),
            cf=alert_cf() + [{"type": "text", "criteria": "containing",
                              "value": "EXPIRED", "fmt": "bold_bad"},
                             {"type": "text", "criteria": "containing",
                              "value": "Use within", "fmt": "warn"}]),
        Col("used_month", "Used this month", 10, "number1",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 26, "wrap"),
        Col("alert", "Action needed", 22, "status",
            formula=blank(
                "item",
                'IF({stock_status}{r}="Out of stock",'
                '"Order now - out of stock",'
                'IF({stock_status}{r}="Low - reorder",'
                '"Reorder "&TEXT(IF({reorder}{r}="",0,{reorder}{r}),"0")'
                '&" "&{unit}{r},'
                'IF({expiry_alert}{r}="EXPIRED","Dispose - expired","OK")))'),
            cf=alert_cf() + [{"type": "text", "criteria": "containing",
                              "value": "Order now", "fmt": "bold_bad"},
                             {"type": "text", "criteria": "containing",
                              "value": "Reorder", "fmt": "warn"}]),
    ]

    return SheetSpec(
        name=NAME,
        title="Inventory & Pet Supplies",
        subtitle="Treats, waste bags, cleaning kit, first aid and "
                 "equipment - with low stock and expiry alerts.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).inventory,
        help_text="Update the 'In stock' column after each supply run. "
                  "Anything at or below the minimum level is flagged.",
    )
````

### `pet_sitting_spreadsheet/sheets/reminders.py`

*255 lines · 3 functions*

> `Reminders` - the alert centre plus a custom task list.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` | The custom task list that sits underneath the alert panel |
| `decorate(sheet, fmts, cfg, reg)` | Write the automatic alert panel above the task list |
| `_cell(col, row)` |  |

**Module constants:** `NAME` ('Reminders'), `ALERT_HEADER_ROW` (6), `TASK_HEADER_ROW` (38), `LAST_COL` (9), `ALERTS` (24 items), `RENEWALS` (2 items).

````python
"""`Reminders` - the alert centre plus a custom task list."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, Registry, SheetSpec, quote_sheet
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import alert_cf, blank, id_col, status_cf

NAME = "Reminders"

ALERT_HEADER_ROW = 6      # Excel row of the alert panel header
TASK_HEADER_ROW = 38      # Excel row of the task list header
LAST_COL = 9

#: (alert, count formula, value formula, what to do, jump-to sheet)
ALERTS = [
    ("Bookings in the next 7 days",
     '=COUNTIFS([[Bookings.date]],">="&TODAY(),[[Bookings.date]],'
     '"<="&TODAY()+7,[[Bookings.status]],"<>Cancelled")',
     '=SUMIFS([[Bookings.total]],[[Bookings.date]],">="&TODAY(),'
     '[[Bookings.date]],"<="&TODAY()+7,[[Bookings.status]],"<>Cancelled")',
     "Confirm the schedule and send reminders", "Bookings"),
    ("Unconfirmed bookings",
     '=COUNTIFS([[Bookings.date]],">="&TODAY(),[[Bookings.confirm_sent]],'
     '"<>Yes",[[Bookings.status]],"<>Cancelled")', "",
     "Send a confirmation message today", "Bookings"),
    ("Booking conflicts / double bookings",
     '=COUNTIF([[Bookings.conflict]],"Conflict*")', "",
     "Two jobs overlap for the same carer", "Bookings"),
    ("Bookings waiting on a visit log",
     '=COUNTIFS([[Bookings.status]],"Completed",'
     '[[Bookings.visit_logged]],"No")', "",
     "Log the visit so the client has a record", "Visit Log"),
    ("Missed visits",
     '=COUNTIF([[Visit Log.status]],"Missed")', "",
     "Call the client and agree what happens next", "Visit Log"),
    ("Visits with no client update sent",
     '=COUNTIFS([[Visit Log.status]],"Completed",'
     '[[Visit Log.update_sent]],"<>Yes")', "",
     "Send the photo and summary", "Visit Log"),
    ("Overdue invoices",
     '=COUNTIF([[Invoices.pay_status]],"Overdue")'
     '+COUNTIF([[Invoices.pay_status]],"Overdue - part paid")',
     '=SUMIFS([[Invoices.balance]],[[Invoices.pay_status]],"Overdue")'
     '+SUMIFS([[Invoices.balance]],[[Invoices.pay_status]],'
     '"Overdue - part paid")',
     "Send a payment reminder with the late fee", "Invoices"),
    ("Invoices due within the reminder window",
     '=COUNTIFS([[Invoices.balance]],">0.004",[[Invoices.due]],'
     '">="&TODAY(),[[Invoices.due]],"<="&TODAY()+ReminderLeadDays)',
     '=SUMIFS([[Invoices.balance]],[[Invoices.balance]],">0.004",'
     '[[Invoices.due]],">="&TODAY(),[[Invoices.due]],'
     '"<="&TODAY()+ReminderLeadDays)',
     "A friendly nudge before the due date", "Invoices"),
    ("Draft invoices not sent",
     '=COUNTIF([[Invoices.status]],"Draft")',
     '=SUMIFS([[Invoices.total]],[[Invoices.status]],"Draft")',
     "Finish and send them - this is unbilled work", "Invoices"),
    ("Payments not reconciled",
     '=COUNTIF([[Payments.reconciled]],"<>Yes")', "",
     "Match them against your bank statement", "Payments"),
    ("Vaccinations expired",
     '=COUNTIF([[Pets.vacc_status]],"Expired")', "",
     "Ask the owner for an updated certificate", "Pets"),
    ("Vaccinations due soon",
     '=COUNTIF([[Pets.vacc_status]],"Due in*")', "",
     "Remind the owner before the booking", "Pets"),
    ("Medication doses missed",
     '=COUNTIF([[Health Log.status]],"Missed")', "",
     "Tell the owner and note the reason", "Health Log"),
    ("Medication refills due",
     '=COUNTIF([[Health Log.refill_alert]],"Overdue*")'
     '+COUNTIF([[Health Log.refill_alert]],"Due in*")', "",
     "Ask the owner to restock the supply", "Health Log"),
    ("Vet appointments coming up",
     '=COUNTIF([[Health Log.appt_alert]],"Due in*")', "",
     "Confirm transport and who attends", "Health Log"),
    ("Open incidents",
     '=COUNTIF([[Incidents.status]],"Open")'
     '+COUNTIF([[Incidents.status]],"In progress")', "",
     "Close the loop with the owner", "Incidents"),
    ("Clients going inactive",
     '=COUNTIF([[Clients.alert]],"Inactive*")', "",
     "Send a win-back offer", "Clients"),
    ("Client follow-ups overdue",
     '=COUNTIF([[Messages.alert]],"Overdue*")', "",
     "Check the communication log", "Messages"),
    ("Unresolved complaints",
     '=COUNTIFS([[Messages.complaint]],"Yes",[[Messages.resolution]],'
     '"<>Resolved")', "",
     "Resolve it before it becomes a review", "Messages"),
    ("Low or out of stock items",
     '=COUNTIF([[Inventory.stock_status]],"Low*")'
     '+COUNTIF([[Inventory.stock_status]],"Out of stock")', "",
     "Add them to the next supply run", "Inventory"),
    ("Supplies past their expiry date",
     '=COUNTIF([[Inventory.expiry_alert]],"EXPIRED")', "",
     "Dispose of them and reorder", "Inventory"),
    ("Staff certificates expired or expiring",
     '=COUNTIF([[Staff.cert_alert]],"EXPIRED")'
     '+COUNTIF([[Staff.cert_alert]],"Due in*")', "",
     "Book the refresher course", "Staff"),
    ("Client contracts not signed",
     '=COUNTIFS([[Clients.status]],"Active",[[Clients.contract]],'
     '"<>Signed")', "",
     "No service without a signed agreement", "Clients"),
    ("Expenses waiting for a receipt",
     '=COUNTIF([[Expenses.alert]],"Missing receipt")', "",
     "Attach the receipt before tax time", "Expenses"),
]

RENEWALS = [
    ("Business insurance renewal", "=Biz_InsuranceRenewal"),
    ("Licence / permit renewal", "=Biz_LicenceRenewal"),
]


def spec(cfg: Config) -> SheetSpec:
    """The custom task list that sits underneath the alert panel."""
    columns = [
        id_col("task", '"TK-"', digits=3, width=9),
        Col("task", "Task", 38, "text"),
        Col("category", "Category", 16, "text"),
        Col("due", "Due date", 12, "date", validate=date_check()),
        Col("owner", "Owner", 16, "text", validate=list_from("StaffNames")),
        Col("priority", "Priority", 11, "status",
            validate=list_from("PriorityList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Urgent", "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing", "value": "High",
                 "fmt": "warn"}]),
        Col("status", "Status", 13, "status",
            validate=list_from("TaskStatusList"), cf=status_cf()),
        Col("days_left", "Days left", 9, "int",
            formula='=IF(OR({due}{r}="",{status}{r}="Done",'
                    '{status}{r}="Cancelled"),"",{due}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 3, "fmt": "warn"}]),
        Col("flag", "Status check", 18, "status",
            formula=blank("task",
                          'IF(OR({status}{r}="Done",'
                          '{status}{r}="Cancelled"),"Closed",'
                          'IF({due}{r}="","No due date",'
                          'IF({due}{r}<TODAY(),"Overdue",'
                          'IF({due}{r}-TODAY()<=ReminderLeadDays,'
                          '"Due soon","On track"))))'),
            cf=alert_cf() + [{"type": "text", "criteria": "containing",
                              "value": "On track", "fmt": "good"},
                             {"type": "text", "criteria": "containing",
                              "value": "Closed", "fmt": "muted"}]),
        Col("notes", "Notes", 34, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Reminders, Alerts & Task List",
        subtitle="Everything the workbook thinks you should do today, plus "
                 "room for your own reminders.",
        columns=columns,
        rows=cfg.row_count("Tasks"),
        header_row=TASK_HEADER_ROW,
        tab_color=Palette.warn_text,
        freeze_cols=2,
        freeze_at=3,
        samples=sample(cfg).tasks,
        help_text="Add your own reminders here - renewals, marketing "
                  "ideas, supply runs, anything with a due date.",
    )


def decorate(sheet, fmts, cfg: Config, reg: Registry) -> None:
    """Write the automatic alert panel above the task list."""
    ws = sheet.ws
    row = ALERT_HEADER_ROW - 3
    sheet.section(row, "Automatic alerts - refreshed every time the file "
                       "opens", LAST_COL)
    row += 1
    sheet.merge(row, 0, LAST_COL,
                "Anything with a count above zero needs a decision today. "
                "Click the sheet name on the right to jump straight to it.",
                fmts.chrome("note"))
    row += 1

    sheet.write(row, 0, "Alert", fmts.chrome("header_left"))
    sheet.merge(row, 1, 1, "Count", fmts.chrome("header"))
    sheet.merge(row, 2, 3, "Value", fmts.chrome("header"))
    sheet.merge(row, 4, 7, "What to do", fmts.chrome("header"))
    sheet.merge(row, 8, LAST_COL, "Go to", fmts.chrome("header"))
    ws.set_row(row, 22)
    row += 1

    label_fmt = fmts.calc("text")
    count_fmt = fmts.calc("int")
    money_fmt = fmts.calc("money")
    note_fmt = fmts.chrome("note")
    link_fmt = fmts.chrome("nav")

    for label, count, value, todo, target in ALERTS:
        sheet.write(row, 0, label, label_fmt)
        ws.write_formula(row, 1, reg.render(count), count_fmt)
        ws.conditional_format(row, 1, row, 1, {
            "type": "cell", "criteria": ">", "value": 0,
            "format": fmts.cf("bad")})
        if value:
            ws.merge_range(row, 2, row, 3, "", money_fmt)
            ws.write_formula(row, 2, reg.render(value), money_fmt)
        else:
            sheet.merge(row, 2, 3, "", money_fmt)
        sheet.merge(row, 4, 7, todo, note_fmt)
        ws.merge_range(row, 8, row, LAST_COL, "", link_fmt)
        ws.write_url(row, 8, f"internal:{quote_sheet(target)}!A1", link_fmt,
                     target)
        row += 1

    # Renewal dates pulled from the Settings tab.
    row += 1
    sheet.subsection(row, "Renewal dates", LAST_COL)
    row += 1
    sheet.write(row, 0, "Renewal", fmts.chrome("header_left"))
    sheet.merge(row, 1, 2, "Date", fmts.chrome("header"))
    sheet.merge(row, 3, 4, "Days away", fmts.chrome("header"))
    sheet.merge(row, 5, LAST_COL, "Status", fmts.chrome("header"))
    row += 1
    for label, formula in RENEWALS:
        sheet.write(row, 0, label, label_fmt)
        ws.merge_range(row, 1, row, 2, "", fmts.calc("date"))
        ws.write_formula(row, 1, formula, fmts.calc("date"))
        ws.merge_range(row, 3, row, 4, "", fmts.calc("int"))
        ws.write_formula(
            row, 3,
            f'=IF({_cell(1, row)}="","",{_cell(1, row)}-TODAY())',
            fmts.calc("int"))
        ws.merge_range(row, 5, row, LAST_COL, "", fmts.calc("status"))
        ws.write_formula(
            row, 5,
            f'=IF({_cell(1, row)}="","Add the date on the Settings tab",'
            f'IF({_cell(1, row)}<TODAY(),"EXPIRED - renew now",'
            f'IF({_cell(1, row)}-TODAY()<=ExpiryAlertDays,'
            f'"Renew soon","Current")))', fmts.calc("status"))
        ws.conditional_format(row, 5, row, 5, {
            "type": "text", "criteria": "containing", "value": "EXPIRED",
            "format": fmts.cf("bad")})
        ws.conditional_format(row, 5, row, 5, {
            "type": "text", "criteria": "containing", "value": "Renew soon",
            "format": fmts.cf("warn")})
        row += 1


def _cell(col: int, row: int) -> str:
    from ..layout import col_letter
    return f"${col_letter(col)}${row + 1}"
````

---

## 4. The pages

The tabs that read the registers: the hidden calculation engine, the dashboard, the reports, the planner, the welcome tab, the printable pages and the help tab.

### `pet_sitting_spreadsheet/sheets/calc.py`

*213 lines · 1 function*

> `Calc` - hidden helper sheet.
>
> Excel has no SORT/FILTER in the 2007-2016 formula set, so the "top 10"
> tables, the twelve month trend series and the category breakdowns are
> pre-computed here with plain SUMIFS/COUNTIFS and then pulled into the
> Dashboard and Reports tabs with LARGE + INDEX/MATCH.
>
> Nothing here needs to be touched by the user; the tab is hidden and every
> cell is a locked formula.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Calc').

````python
"""
`Calc` - hidden helper sheet.

Excel has no SORT/FILTER in the 2007-2016 formula set, so the "top 10"
tables, the twelve month trend series and the category breakdowns are
pre-computed here with plain SUMIFS/COUNTIFS and then pulled into the
Dashboard and Reports tabs with LARGE + INDEX/MATCH.

Nothing here needs to be touched by the user; the tab is hidden and every
cell is a locked formula.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from . import lists

NAME = "Calc"

_PERIOD = ('[[Payments.date]],">="&PeriodStart,[[Payments.date]],'
           '"<="&PeriodEnd')


def spec(cfg: Config) -> SheetSpec:
    rows = max(cfg.row_count("Clients"), cfg.row_count("Services"),
               cfg.row_count("Staff"), 60)

    def mirror(sheet: str, key: str) -> str:
        return f'=IF([[{sheet}.{key}#]]="","",[[{sheet}.{key}#]])'

    columns = [
        # ---------------------------------------------------- clients block
        Col("c_name", "Client", 20, "text", formula=mirror("Clients", "name")),
        Col("c_rev_p", "Revenue (dashboard period)", 14, "money",
            formula='=IF({c_name}{r}="","",SUMIFS([[Payments.received]],'
                    '[[Payments.client]],{c_name}{r},' + _PERIOD + '))'),
        Col("c_key_p", "Rank key (period)", 14, "number",
            formula='=IF({c_name}{r}="","",IF({c_rev_p}{r}<=0,"",'
                    '{c_rev_p}{r}+ROW()/1000000))'),
        Col("c_rev_r", "Revenue (report range)", 14, "money",
            formula='=IF({c_name}{r}="","",SUMIFS([[Payments.received]],'
                    '[[Payments.client]],{c_name}{r},[[Payments.date]],'
                    '">="&RepFrom,[[Payments.date]],"<="&RepTo))'),
        Col("c_key_r", "Rank key (range)", 14, "number",
            formula='=IF({c_name}{r}="","",IF({c_rev_r}{r}<=0,"",'
                    '{c_rev_r}{r}+ROW()/1000000))'),
        Col("c_out", "Outstanding", 13, "money",
            formula='=IF({c_name}{r}="","",SUMIFS([[Invoices.balance]],'
                    '[[Invoices.client]],{c_name}{r},[[Invoices.status]],'
                    '"<>Cancelled"))'),
        Col("c_out_key", "Outstanding key", 14, "number",
            formula='=IF({c_name}{r}="","",IF({c_out}{r}<=0.004,"",'
                    '{c_out}{r}+ROW()/1000000))'),
        Col("c_bookings_r", "Bookings (range)", 12, "int",
            formula='=IF({c_name}{r}="","",COUNTIFS([[Bookings.client]],'
                    '{c_name}{r},[[Bookings.date]],">="&RepFrom,'
                    '[[Bookings.date]],"<="&RepTo,[[Bookings.status]],'
                    '"<>Cancelled"))'),
        # --------------------------------------------------- services block
        Col("s_name", "Service", 22, "text",
            formula=mirror("Services", "name")),
        Col("s_rev_p", "Service revenue (period)", 14, "money",
            formula='=IF({s_name}{r}="","",SUMIFS([[Bookings.total]],'
                    '[[Bookings.service]],{s_name}{r},[[Bookings.date]],'
                    '">="&PeriodStart,[[Bookings.date]],"<="&PeriodEnd,'
                    '[[Bookings.status]],"<>Cancelled"))'),
        Col("s_count_p", "Service bookings (period)", 12, "int",
            formula='=IF({s_name}{r}="","",COUNTIFS([[Bookings.service]],'
                    '{s_name}{r},[[Bookings.date]],">="&PeriodStart,'
                    '[[Bookings.date]],"<="&PeriodEnd,[[Bookings.status]],'
                    '"<>Cancelled"))'),
        Col("s_rev_r", "Service revenue (range)", 14, "money",
            formula='=IF({s_name}{r}="","",SUMIFS([[Bookings.total]],'
                    '[[Bookings.service]],{s_name}{r},[[Bookings.date]],'
                    '">="&RepFrom,[[Bookings.date]],"<="&RepTo,'
                    '[[Bookings.status]],"<>Cancelled"))'),
        Col("s_count_r", "Service bookings (range)", 12, "int",
            formula='=IF({s_name}{r}="","",COUNTIFS([[Bookings.service]],'
                    '{s_name}{r},[[Bookings.date]],">="&RepFrom,'
                    '[[Bookings.date]],"<="&RepTo,[[Bookings.status]],'
                    '"<>Cancelled"))'),
        Col("s_cost_r", "Service cost (range)", 13, "money",
            formula='=IF({s_name}{r}="","",{s_count_r}{r}'
                    '*IF([[Services.staff_cost#]]="",0,'
                    '[[Services.staff_cost#]]))'),
        Col("s_profit_r", "Service profit (range)", 13, "money",
            formula='=IF({s_name}{r}="","",{s_rev_r}{r}-{s_cost_r}{r})'),
        # ------------------------------------------- expense category block
        Col("e_cat", "Expense category", 20, "text",
            formula='=IF(' + lists.ref("ExpenseCategoryList") + '="","",'
                    + lists.ref("ExpenseCategoryList") + ')'),
        Col("e_total_p", "Spend (period)", 13, "money",
            formula='=IF({e_cat}{r}="","",SUMIFS([[Expenses.amount]],'
                    '[[Expenses.category]],{e_cat}{r},[[Expenses.date]],'
                    '">="&PeriodStart,[[Expenses.date]],"<="&PeriodEnd,'
                    '[[Expenses.type]],"<>Personal"))'),
        Col("e_total_r", "Spend (range)", 13, "money",
            formula='=IF({e_cat}{r}="","",SUMIFS([[Expenses.amount]],'
                    '[[Expenses.category]],{e_cat}{r},[[Expenses.date]],'
                    '">="&RepFrom,[[Expenses.date]],"<="&RepTo,'
                    '[[Expenses.type]],"<>Personal"))'),
        # ------------------------------------------------ 12 month series --
        Col("m_start", "Month start", 12, "date",
            formula='=IF(ROW()-4>12,"",DATE(Dash_Year,ROW()-4,1))'),
        Col("m_end", "Month end", 12, "date",
            formula='=IF({m_start}{r}="","",EOMONTH({m_start}{r},0))'),
        Col("m_label", "Month", 10, "text",
            formula='=IF({m_start}{r}="","",TEXT({m_start}{r},"mmm"))'),
        Col("m_rev", "Revenue", 13, "money",
            formula='=IF({m_start}{r}="","",SUMIFS([[Payments.received]],'
                    '[[Payments.date]],">="&{m_start}{r},'
                    '[[Payments.date]],"<="&{m_end}{r}))'),
        Col("m_inv", "Invoiced", 13, "money",
            formula='=IF({m_start}{r}="","",SUMIFS([[Invoices.total]],'
                    '[[Invoices.date]],">="&{m_start}{r},[[Invoices.date]],'
                    '"<="&{m_end}{r},[[Invoices.status]],"<>Cancelled"))'),
        Col("m_exp", "Expenses", 13, "money",
            formula='=IF({m_start}{r}="","",SUMIFS([[Expenses.amount]],'
                    '[[Expenses.date]],">="&{m_start}{r},[[Expenses.date]],'
                    '"<="&{m_end}{r},[[Expenses.type]],"<>Personal"))'),
        Col("m_profit", "Profit", 13, "money",
            formula='=IF({m_start}{r}="","",{m_rev}{r}-{m_exp}{r})'),
        Col("m_bookings", "Bookings", 10, "int",
            formula='=IF({m_start}{r}="","",COUNTIFS([[Bookings.date]],'
                    '">="&{m_start}{r},[[Bookings.date]],"<="&{m_end}{r},'
                    '[[Bookings.status]],"<>Cancelled"))'),
        Col("m_completed", "Completed", 10, "int",
            formula='=IF({m_start}{r}="","",COUNTIFS([[Bookings.date]],'
                    '">="&{m_start}{r},[[Bookings.date]],"<="&{m_end}{r},'
                    '[[Bookings.status]],"Completed"))'),
        Col("m_cancelled", "Cancelled", 10, "int",
            formula='=IF({m_start}{r}="","",COUNTIFS([[Bookings.date]],'
                    '">="&{m_start}{r},[[Bookings.date]],"<="&{m_end}{r},'
                    '[[Bookings.status]],"Cancelled"))'),
        Col("m_visits", "Visits", 10, "int",
            formula='=IF({m_start}{r}="","",COUNTIFS([[Visit Log.date]],'
                    '">="&{m_start}{r},[[Visit Log.date]],"<="&{m_end}{r},'
                    '[[Visit Log.status]],"Completed"))'),
        Col("m_new_clients", "New clients", 10, "int",
            formula='=IF({m_start}{r}="","",COUNTIFS([[Clients.reg_date]],'
                    '">="&{m_start}{r},[[Clients.reg_date]],'
                    '"<="&{m_end}{r}))'),
        Col("m_km", "Business distance", 11, "number1",
            formula='=IF({m_start}{r}="","",SUMIFS([[Mileage.total_km]],'
                    '[[Mileage.date]],">="&{m_start}{r},[[Mileage.date]],'
                    '"<="&{m_end}{r},[[Mileage.business]],"Business"))'),
        Col("m_prev_rev", "Revenue last year", 13, "money",
            formula='=IF({m_start}{r}="","",SUMIFS([[Payments.received]],'
                    '[[Payments.date]],">="&DATE(Dash_Year-1,ROW()-4,1),'
                    '[[Payments.date]],"<="&EOMONTH(DATE(Dash_Year-1,'
                    'ROW()-4,1),0)))'),
        # ---------------------------------------------------- staff block --
        Col("st_name", "Staff", 18, "text", formula=mirror("Staff", "name")),
        Col("st_visits_r", "Visits (range)", 11, "int",
            formula='=IF({st_name}{r}="","",COUNTIFS([[Visit Log.staff]],'
                    '{st_name}{r},[[Visit Log.date]],">="&RepFrom,'
                    '[[Visit Log.date]],"<="&RepTo,[[Visit Log.status]],'
                    '"Completed"))'),
        Col("st_hours_r", "Hours (range)", 11, "number",
            formula='=IF({st_name}{r}="","",ROUND(SUMIFS('
                    '[[Visit Log.duration]],[[Visit Log.staff]],'
                    '{st_name}{r},[[Visit Log.date]],">="&RepFrom,'
                    '[[Visit Log.date]],"<="&RepTo,[[Visit Log.status]],'
                    '"Completed")/60,2))'),
        Col("st_rev_r", "Revenue (range)", 13, "money",
            formula='=IF({st_name}{r}="","",SUMIFS([[Bookings.total]],'
                    '[[Bookings.staff]],{st_name}{r},[[Bookings.date]],'
                    '">="&RepFrom,[[Bookings.date]],"<="&RepTo,'
                    '[[Bookings.status]],"<>Cancelled"))'),
        Col("st_km_r", "Distance (range)", 11, "number1",
            formula='=IF({st_name}{r}="","",SUMIFS([[Mileage.total_km]],'
                    '[[Mileage.staff]],{st_name}{r},[[Mileage.date]],'
                    '">="&RepFrom,[[Mileage.date]],"<="&RepTo))'),
        # ------------------------------------------- cancellations/methods -
        Col("cr_reason", "Cancellation reason", 20, "text",
            formula='=IF(' + lists.ref("CancelReasonList") + '="","",'
                    + lists.ref("CancelReasonList") + ')'),
        Col("cr_count", "Cancellations (range)", 12, "int",
            formula='=IF({cr_reason}{r}="","",COUNTIFS('
                    '[[Bookings.cancel_reason]],{cr_reason}{r},'
                    '[[Bookings.date]],">="&RepFrom,[[Bookings.date]],'
                    '"<="&RepTo))'),
        Col("pm_method", "Payment method", 18, "text",
            formula='=IF(' + lists.ref("PaymentMethodList") + '="","",'
                    + lists.ref("PaymentMethodList") + ')'),
        Col("pm_total", "Collected (range)", 13, "money",
            formula='=IF({pm_method}{r}="","",SUMIFS([[Payments.received]],'
                    '[[Payments.method]],{pm_method}{r},[[Payments.date]],'
                    '">="&RepFrom,[[Payments.date]],"<="&RepTo))'),
        Col("ref_source", "Referral source", 18, "text",
            formula='=IF(' + lists.ref("ReferralSourceList") + '="","",'
                    + lists.ref("ReferralSourceList") + ')'),
        Col("ref_clients", "Clients from source", 12, "int",
            formula='=IF({ref_source}{r}="","",COUNTIFS('
                    '[[Clients.referral]],{ref_source}{r}))'),
        Col("ref_revenue", "Revenue from source", 13, "money",
            formula='=IF({ref_source}{r}="","",SUMIFS([[Clients.revenue]],'
                    '[[Clients.referral]],{ref_source}{r}))'),
    ]

    return SheetSpec(
        name=NAME,
        title="Calculation engine (do not edit)",
        subtitle="Helper tables for the dashboard and reports. This tab is "
                 "hidden by default - unhide it only if you want to audit "
                 "the maths.",
        columns=columns,
        rows=rows,
        tab_color="#9AA7AD",
        freeze_cols=1,
        help_text="Every cell on this tab is a locked formula.",
    )
````

### `pet_sitting_spreadsheet/sheets/dashboard.py`

*375 lines · 2 functions*

> `Dashboard` - the one page view of the business.

| Defines | What it does |
|---------|--------------|
| `_period(field_sheet, field)` |  |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dashboard'), `LAST_COL` (17), `MONTH_ROWS` (12), `FIRST_CALC_ROW` (5).

````python
"""`Dashboard` - the one page view of the business."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Dashboard"
LAST_COL = 17            # column R - six KPI tiles of three columns each
MONTH_ROWS = 12
FIRST_CALC_ROW = 5       # Calc sheet data starts on row 5


def _period(field_sheet: str, field: str) -> str:
    return (f'[[{field_sheet}.{field}]],">="&PeriodStart,'
            f'[[{field_sheet}.{field}]],"<="&PeriodEnd')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    ws = s.ws

    s.banner("Business Dashboard",
             "Pick a period, then read the business at a glance. Every "
             "number is live - it updates as you fill in the other tabs.",
             LAST_COL)

    for c in range(0, LAST_COL + 1):
        s.set_col(c, c, 12.5)

    # ------------------------------------------------------ period picker
    row = 4
    row = s.section(row, "Reporting period", LAST_COL)
    s.merge(row, 0, 1, "Year", fmts.chrome("label_bold"))
    s.write(row, 2, cfg.today.year, fmts.input("int"))
    ws.data_validation(row, 2, row, 2, number_check(2000, 2100, False))
    book.define_name("Dash_Year",
                     f"={quote_sheet(NAME)}!$C${row + 1}")

    s.merge(row, 3, 4, "Month", fmts.chrome("label_bold"))
    # Opens on the whole year so the first screen is full of numbers;
    # pick a month from the drop down to narrow everything down.
    s.merge(row, 5, 6, "All year", fmts.input("text"))
    ws.data_validation(row, 5, row, 5, list_from("MonthList"))
    book.define_name("Dash_MonthName",
                     f"={quote_sheet(NAME)}!$F${row + 1}")

    s.merge(row, 7, 8, "Period start", fmts.chrome("label_bold"))
    s.merge(row, 9, 10, None, fmts.calc("date"))
    ws.write_formula(
        row, 9,
        '=IF(Dash_MonthName="All year",DATE(Dash_Year,1,1),'
        'DATE(Dash_Year,MAX(1,MATCH(Dash_MonthName,MonthList,0)-1),1))',
        fmts.calc("date"))
    book.define_name("PeriodStart", f"={quote_sheet(NAME)}!$J${row + 1}")

    s.merge(row, 11, 12, "Period end", fmts.chrome("label_bold"))
    s.merge(row, 13, 14, None, fmts.calc("date"))
    ws.write_formula(
        row, 13,
        '=IF(Dash_MonthName="All year",DATE(Dash_Year,12,31),'
        'EOMONTH(PeriodStart,0))', fmts.calc("date"))
    book.define_name("PeriodEnd", f"={quote_sheet(NAME)}!$N${row + 1}")

    s.merge(row, 15, 16, "Today", fmts.chrome("label_bold"))
    s.write(row, 17, "=TODAY()", fmts.calc("date"))
    row += 2

    # ------------------------------------------------------- money tiles --
    row = s.section(row, "Money", LAST_COL)
    revenue = ('=SUMIFS([[Payments.received]],'
               + _period("Payments", "date") + ')')
    invoiced = ('=SUMIFS([[Invoices.total]],' + _period("Invoices", "date")
                + ',[[Invoices.status]],"<>Cancelled")')
    expenses = ('=SUMIFS([[Expenses.amount]],' + _period("Expenses", "date")
                + ',[[Expenses.type]],"<>Personal")')

    s.kpi(row, 0, "Revenue collected", revenue, "money", 3,
          "Payments received in the period")
    book.define_name("Dash_Revenue", f"={quote_sheet(NAME)}!$A${row + 2}")
    s.kpi(row, 3, "Invoiced", invoiced, "money", 3,
          "Value of invoices raised")
    s.kpi(row, 6, "Outstanding", '=SUMIFS([[Invoices.balance]],'
                                 '[[Invoices.status]],"<>Cancelled")',
          "money", 3, "Unpaid balance, all invoices")
    s.kpi(row, 9, "Overdue", '=SUMIFS([[Invoices.balance]],'
                             '[[Invoices.pay_status]],"Overdue")'
                             '+SUMIFS([[Invoices.balance]],'
                             '[[Invoices.pay_status]],"Overdue - part paid")',
          "money", 3, "Past the due date - chase today")
    s.kpi(row, 12, "Expenses", expenses, "money", 3,
          "Business costs in the period")
    book.define_name("Dash_Expenses", f"={quote_sheet(NAME)}!$M${row + 2}")
    s.kpi(row, 15, "Net profit", "=Dash_Revenue-Dash_Expenses", "money", 3,
          "Revenue collected less expenses")
    book.define_name("Dash_Profit", f"={quote_sheet(NAME)}!$P${row + 2}")
    row += 3

    s.kpi(row, 0, "Profit margin",
          '=IFERROR(Dash_Profit/Dash_Revenue,0)', "percent", 3,
          "Net profit as a share of revenue")
    s.kpi(row, 3, "Average booking value",
          '=IFERROR(SUMIFS([[Bookings.total]],'
          + _period("Bookings", "date")
          + ',[[Bookings.status]],"<>Cancelled")/COUNTIFS('
          + _period("Bookings", "date")
          + ',[[Bookings.status]],"<>Cancelled"),0)', "money", 3,
          "Booked value per job")
    s.kpi(row, 6, "Revenue per active client",
          '=IFERROR(Dash_Revenue/COUNTIF([[Clients.status]],"Active"),0)',
          "money", 3, "Period revenue / active clients")
    s.kpi(row, 9, "Mileage claim",
          '=SUMIFS([[Mileage.claimable]],' + _period("Mileage", "date") + ')',
          "money", 3, "Travel you can claim back")
    s.kpi(row, 12, "Tax to set aside",
          '=MAX(0,Dash_Profit)*TaxReservePct', "money", 3,
          "Profit x tax reserve % from Settings")
    s.kpi(row, 15, "Cash collected per visit",
          '=IFERROR(Dash_Revenue/COUNTIFS(' + _period("Visit Log", "date")
          + ',[[Visit Log.status]],"Completed"),0)', "money", 3,
          "Revenue / completed visits")
    row += 4

    # -------------------------------------------------- operations tiles --
    row = s.section(row, "Bookings, visits & clients", LAST_COL)
    s.kpi(row, 0, "Bookings in period",
          '=COUNTIFS(' + _period("Bookings", "date")
          + ',[[Bookings.status]],"<>Cancelled")', "int", 3,
          "Excludes cancellations")
    s.kpi(row, 3, "Completed",
          '=COUNTIFS(' + _period("Bookings", "date")
          + ',[[Bookings.status]],"Completed")', "int", 3,
          "Jobs delivered")
    s.kpi(row, 6, "Cancelled",
          '=COUNTIFS(' + _period("Bookings", "date")
          + ',[[Bookings.status]],"Cancelled")', "int", 3,
          "Watch the cancellation rate")
    s.kpi(row, 9, "Upcoming (7 days)",
          '=COUNTIFS([[Bookings.date]],">="&TODAY(),[[Bookings.date]],'
          '"<="&TODAY()+7,[[Bookings.status]],"<>Cancelled")', "int", 3,
          "Next seven days")
    s.kpi(row, 12, "Visits completed",
          '=COUNTIFS(' + _period("Visit Log", "date")
          + ',[[Visit Log.status]],"Completed")', "int", 3,
          "Logged in the Visit Log")
    s.kpi(row, 15, "Missed visits",
          '=COUNTIFS(' + _period("Visit Log", "date")
          + ',[[Visit Log.status]],"Missed")', "int", 3,
          "Investigate every one")
    row += 3

    s.kpi(row, 0, "Active clients",
          '=COUNTIF([[Clients.status]],"Active")', "int", 3,
          "Clients marked Active")
    s.kpi(row, 3, "Pets on the books",
          '=COUNTIF([[Pets.name]],"?*")', "int", 3,
          "All registered pets")
    s.kpi(row, 6, "New clients",
          '=COUNTIFS(' + _period("Clients", "reg_date") + ')', "int", 3,
          "Registered in the period")
    s.kpi(row, 9, "Repeat clients",
          '=COUNTIF([[Clients.bookings]],">1")', "int", 3,
          "More than one booking")
    s.kpi(row, 12, "Unconfirmed bookings",
          '=COUNTIFS([[Bookings.date]],">="&TODAY(),'
          '[[Bookings.confirm_sent]],"<>Yes",[[Bookings.status]],'
          '"<>Cancelled")', "int", 3, "Send a confirmation")
    s.kpi(row, 15, "Average rating",
          '=IFERROR(AVERAGEIFS([[Messages.rating]],'
          + _period("Messages", "date") + '),"-")', "number", 3,
          "Client satisfaction, 1 to 5")
    row += 4

    # ------------------------------------------------- 12 month trend ----
    row = s.section(row,
                    "Monthly performance - trends for the selected year",
                    LAST_COL)
    headers = ["Month", "Revenue", "Invoiced", "Expenses", "Profit",
               "Margin", "Bookings", "Completed", "Cancelled", "Visits",
               "New clients", "Distance", "Last year", "Change"]
    s.table_head(row, 0, headers)
    first_table_row = row + 1
    styles = ["text", "money", "money", "money", "money", "percent", "int",
              "int", "int", "int", "int", "number1", "money", "percent"]
    for i in range(MONTH_ROWS):
        r = first_table_row + i
        calc_row = FIRST_CALC_ROW + i
        values = [
            f'=[[Calc.m_label#{calc_row}]]',
            f'=[[Calc.m_rev#{calc_row}]]',
            f'=[[Calc.m_inv#{calc_row}]]',
            f'=[[Calc.m_exp#{calc_row}]]',
            f'=[[Calc.m_profit#{calc_row}]]',
            f'=IFERROR([[Calc.m_profit#{calc_row}]]'
            f'/[[Calc.m_rev#{calc_row}]],"")',
            f'=[[Calc.m_bookings#{calc_row}]]',
            f'=[[Calc.m_completed#{calc_row}]]',
            f'=[[Calc.m_cancelled#{calc_row}]]',
            f'=[[Calc.m_visits#{calc_row}]]',
            f'=[[Calc.m_new_clients#{calc_row}]]',
            f'=[[Calc.m_km#{calc_row}]]',
            f'=[[Calc.m_prev_rev#{calc_row}]]',
            f'=IFERROR([[Calc.m_rev#{calc_row}]]'
            f'/[[Calc.m_prev_rev#{calc_row}]]-1,"")',
        ]
        s.calc_row(r, 0, values, styles)
    last_table_row = first_table_row + MONTH_ROWS - 1

    total_row = last_table_row + 1
    s.write(total_row, 0, "Year total", fmts.chrome("total_label"))
    for i, style in enumerate(styles[1:], start=1):
        letter = col_letter(i)
        fmt = fmts.chrome("total_value",
                          num_format=s.num_format(style))
        if style == "percent":
            formula = (f'=IFERROR(E{total_row + 1}/B{total_row + 1},"")'
                       if i == 5 else
                       f'=IFERROR(B{total_row + 1}/M{total_row + 1}-1,"")')
            ws.write_formula(total_row, i, formula, fmt)
        else:
            ws.write_formula(
                total_row, i,
                f"=SUM({letter}{first_table_row + 1}:"
                f"{letter}{last_table_row + 1})", fmt)

    # Sparklines (supported by XlsxWriter 0.3.0 and later).
    spark_row = total_row + 1
    s.write(spark_row, 0, "Trend", fmts.chrome("label_bold"))
    spark_specs = [(1, "#2E7D8F"), (3, "#C0504D"), (4, "#1F4E5F"),
                   (6, "#E2A33C")]
    for col_idx, colour in spark_specs:
        letter = col_letter(col_idx)
        try:
            ws.add_sparkline(spark_row, col_idx, {
                "range": (f"{quote_sheet(NAME)}!{letter}"
                          f"{first_table_row + 1}:{letter}"
                          f"{last_table_row + 1}"),
                "type": "column",
                "series_color": colour,
            })
        except Exception:        # pragma: no cover - older/newer libraries
            pass
    for col_idx in range(1, len(headers)):
        if col_idx not in [c for c, _ in spark_specs]:
            ws.write_blank(spark_row, col_idx, None,
                           fmts.chrome("legend_calc"))
    row = spark_row + 2

    # ------------------------------------------ service / category tables
    row = s.section(row, "Where the money comes from - and where it goes",
                    LAST_COL)
    s.merge(row, 0, 5, "Revenue by service (selected period)",
            fmts.chrome("subsection"))
    s.merge(row, 7, 12, "Expenses by category (selected period)",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["Service", "Bookings", "Revenue", "Share"])
    s.table_head(row, 7, ["Category", "Spend", "Share", ""])
    first_split = row + 1
    service_rows = min(16, cfg.row_count("Services"))
    for i in range(service_rows):
        r = first_split + i
        calc_row = FIRST_CALC_ROW + i
        s.calc_row(r, 0, [
            f'=IF([[Calc.s_name#{calc_row}]]="","",'
            f'[[Calc.s_name#{calc_row}]])',
            f'=IF([[Calc.s_name#{calc_row}]]="","",'
            f'[[Calc.s_count_p#{calc_row}]])',
            f'=IF([[Calc.s_name#{calc_row}]]="","",'
            f'[[Calc.s_rev_p#{calc_row}]])',
            f'=IFERROR([[Calc.s_rev_p#{calc_row}]]'
            f'/SUM([[Calc.s_rev_p]]),"")',
        ], ["text", "int", "money", "percent"])
    ws.conditional_format(first_split, 2, first_split + service_rows - 1, 2,
                          {"type": "data_bar", "bar_color": "#2E7D8F"})

    category_rows = 19
    for i in range(category_rows):
        r = first_split + i
        calc_row = FIRST_CALC_ROW + i
        s.calc_row(r, 7, [
            f'=IF([[Calc.e_cat#{calc_row}]]="","",'
            f'[[Calc.e_cat#{calc_row}]])',
            f'=IF([[Calc.e_cat#{calc_row}]]="","",'
            f'[[Calc.e_total_p#{calc_row}]])',
            f'=IFERROR([[Calc.e_total_p#{calc_row}]]'
            f'/SUM([[Calc.e_total_p]]),"")',
            "",
        ], ["text", "money", "percent", "text"])
    ws.conditional_format(first_split, 8, first_split + category_rows - 1, 8,
                          {"type": "data_bar", "bar_color": "#C0504D"})
    row = first_split + max(service_rows, category_rows) + 1

    # ------------------------------------------------ top clients + alerts
    row = s.section(row, "Top clients & action list", LAST_COL)
    s.merge(row, 0, 5, "Top 10 clients this period",
            fmts.chrome("subsection"))
    s.merge(row, 7, 12, "What needs attention right now",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["#", "Client", "Revenue", "Share"])
    s.table_head(row, 7, ["Alert", "Count", "Value", ""])
    first_top = row + 1
    for i in range(10):
        r = first_top + i
        n = i + 1
        s.calc_row(r, 0, [
            n,
            f'=IFERROR(INDEX([[Calc.c_name]],MATCH(LARGE([[Calc.c_key_p]],'
            f'{n}),[[Calc.c_key_p]],0)),"")',
            f'=IFERROR(INDEX([[Calc.c_rev_p]],MATCH(LARGE([[Calc.c_key_p]],'
            f'{n}),[[Calc.c_key_p]],0)),"")',
            f'=IFERROR(INDEX([[Calc.c_rev_p]],MATCH(LARGE([[Calc.c_key_p]],'
            f'{n}),[[Calc.c_key_p]],0))/SUM([[Calc.c_rev_p]]),"")',
        ], ["int", "text", "money", "percent"])
    ws.conditional_format(first_top, 2, first_top + 9, 2,
                          {"type": "data_bar", "bar_color": "#E2A33C"})

    alerts = [
        ("Overdue invoices",
         '=COUNTIFS([[Invoices.pay_status]],"Overdue")'
         '+COUNTIFS([[Invoices.pay_status]],"Overdue - part paid")',
         '=SUMIFS([[Invoices.balance]],[[Invoices.pay_status]],"Overdue")'
         '+SUMIFS([[Invoices.balance]],[[Invoices.pay_status]],'
         '"Overdue - part paid")'),
        ("Unconfirmed bookings",
         '=COUNTIFS([[Bookings.date]],">="&TODAY(),'
         '[[Bookings.confirm_sent]],"<>Yes",[[Bookings.status]],'
         '"<>Cancelled")',
         '=SUMIFS([[Bookings.total]],[[Bookings.date]],">="&TODAY(),'
         '[[Bookings.confirm_sent]],"<>Yes",[[Bookings.status]],'
         '"<>Cancelled")'),
        ("Booking conflicts",
         '=COUNTIF([[Bookings.conflict]],"Conflict*")', ""),
        ("Visits to log",
         '=COUNTIFS([[Bookings.status]],"Completed",'
         '[[Bookings.visit_logged]],"No")', ""),
        ("Vaccinations expired or due",
         '=COUNTIF([[Pets.vacc_status]],"Expired")'
         '+COUNTIF([[Pets.vacc_status]],"Due in*")', ""),
        ("Medication follow-ups",
         '=COUNTIF([[Health Log.status]],"Missed")'
         '+COUNTIF([[Health Log.refill_alert]],"Overdue*")', ""),
        ("Open incidents",
         '=COUNTIF([[Incidents.status]],"Open")'
         '+COUNTIF([[Incidents.status]],"In progress")', ""),
        ("Low or out of stock",
         '=COUNTIF([[Inventory.stock_status]],"Low*")'
         '+COUNTIF([[Inventory.stock_status]],"Out of stock")',
         '=SUMIFS([[Inventory.value]],[[Inventory.stock_status]],'
         '"Out of stock")'),
        ("Inactive clients to win back",
         '=COUNTIFS([[Clients.alert]],"Inactive*")', ""),
        ("Unresolved complaints",
         '=COUNTIFS([[Messages.complaint]],"Yes",[[Messages.resolution]],'
         '"<>Resolved")', ""),
    ]
    for i, (label, count, value) in enumerate(alerts):
        r = first_top + i
        s.calc_row(r, 7, [label, count, value or "", ""],
                   ["text", "int", "money", "text"])
        ws.conditional_format(r, 8, r, 8, {
            "type": "cell", "criteria": ">", "value": 0,
            "format": fmts.cf("bad")})
    row = first_top + 11

    s.note(row,
           "Tip: the Reminders tab lists every one of these alerts row by "
           "row, and the Reports tab turns the same data into printable "
           "monthly, quarterly and year-end reports.",
           LAST_COL, height=24)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `pet_sitting_spreadsheet/sheets/reports.py`

*380 lines · 3 functions*

> `Reports` - printable business analytics for any date range.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |
| `_start_of_year(cfg)` |  |
| `_end_of_year(cfg)` |  |

**Module constants:** `NAME` ('Reports'), `LAST_COL` (11), `FIRST_CALC_ROW` (5).

````python
"""`Reports` - printable business analytics for any date range."""

from __future__ import annotations

import datetime as _dt

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Reports"
LAST_COL = 11            # column L
FIRST_CALC_ROW = 5


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    ws = s.ws

    s.banner("Reports & Business Analytics",
             "Set a date range once at the top - every report below "
             "recalculates. Print the whole tab for your accountant or "
             "year-end review.",
             LAST_COL)

    s.set_col(0, 0, 30)
    s.set_col(1, 1, 16)
    s.set_col(2, 2, 16)
    s.set_col(3, 3, 16)
    s.set_col(4, 4, 3)
    s.set_col(5, 5, 30)
    s.set_col(6, 8, 16)
    s.set_col(9, LAST_COL, 14)

    def metric(row, col, label, formula, style="money", label_span=1):
        s.merge(row, col, col + label_span - 1, label,
                fmts.chrome("label"))
        fmt = fmts.calc(style)
        ws.write_formula(row, col + label_span,
                         reg.render(formula), fmt)
        return row + 1

    def table(row, col, headers, widths=None):
        return s.table_head(row, col, headers, widths)

    # ------------------------------------------------------- range picker
    row = 4
    row = s.section(row, "Report period", LAST_COL)
    s.merge(row, 0, 0, "From", fmts.chrome("label_bold"))
    s.write(row, 1, _start_of_year(cfg), fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("RepFrom", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "To", fmts.chrome("label_bold"))
    s.write(row, 3, _end_of_year(cfg), fmts.input("date"))
    ws.data_validation(row, 3, row, 3, date_check())
    book.define_name("RepTo", f"={quote_sheet(NAME)}!$D${row + 1}")
    s.merge(row, 5, LAST_COL,
            "Try 1st to 31st of a month for a monthly report, or 1 Jan to "
            "31 Dec for the year-end pack.", fmts.chrome("note"))
    row += 2

    # ------------------------------------------------------ booking report
    row = s.section(row, "1. Booking & service delivery report", LAST_COL)
    bk_range = ('[[Bookings.date]],">="&RepFrom,[[Bookings.date]],'
                '"<="&RepTo')
    vs_range = ('[[Visit Log.date]],">="&RepFrom,[[Visit Log.date]],'
                '"<="&RepTo')
    left = row
    left = metric(left, 0, "Bookings taken",
                  f'=COUNTIFS({bk_range})', "int")
    left = metric(left, 0, "Completed bookings",
                  f'=COUNTIFS({bk_range},[[Bookings.status]],"Completed")',
                  "int")
    left = metric(left, 0, "Cancelled bookings",
                  f'=COUNTIFS({bk_range},[[Bookings.status]],"Cancelled")',
                  "int")
    left = metric(left, 0, "No shows",
                  f'=COUNTIFS({bk_range},[[Bookings.status]],"No show")',
                  "int")
    left = metric(left, 0, "Still to deliver",
                  f'=COUNTIFS({bk_range},[[Bookings.status]],"Confirmed")'
                  f'+COUNTIFS({bk_range},[[Bookings.status]],"Pending")',
                  "int")
    left = metric(left, 0, "Completion rate",
                  f'=IFERROR(COUNTIFS({bk_range},[[Bookings.status]],'
                  f'"Completed")/COUNTIFS({bk_range}),0)', "percent")
    left = metric(left, 0, "Cancellation rate",
                  f'=IFERROR(COUNTIFS({bk_range},[[Bookings.status]],'
                  f'"Cancelled")/COUNTIFS({bk_range}),0)', "percent")

    right = row
    right = metric(right, 5, "Visits completed",
                   f'=COUNTIFS({vs_range},[[Visit Log.status]],'
                   f'"Completed")', "int")
    right = metric(right, 5, "Visits missed",
                   f'=COUNTIFS({vs_range},[[Visit Log.status]],"Missed")',
                   "int")
    right = metric(right, 5, "Hours of care delivered",
                   f'=ROUND(SUMIFS([[Visit Log.duration]],{vs_range},'
                   f'[[Visit Log.status]],"Completed")/60,1)', "number")
    right = metric(right, 5, "Walk distance logged",
                   f'=SUMIFS([[Visit Log.walk_distance]],{vs_range})',
                   "number1")
    right = metric(right, 5, "On-time arrival rate",
                   f'=IFERROR(COUNTIFS({vs_range},[[Visit Log.punctual]],'
                   f'"On time")/COUNTIFS({vs_range},'
                   f'[[Visit Log.status]],"Completed"),0)', "percent")
    right = metric(right, 5, "Medication doses given",
                   '=COUNTIFS([[Health Log.date]],">="&RepFrom,'
                   '[[Health Log.date]],"<="&RepTo,[[Health Log.status]],'
                   '"Given")', "int")
    right = metric(right, 5, "Medication doses missed",
                   '=COUNTIFS([[Health Log.date]],">="&RepFrom,'
                   '[[Health Log.date]],"<="&RepTo,[[Health Log.status]],'
                   '"Missed")', "int")
    row = max(left, right) + 1

    # ----------------------------------------------------- financial report
    row = s.section(row, "2. Profit & loss for the period", LAST_COL)
    inc = ('[[Payments.date]],">="&RepFrom,[[Payments.date]],"<="&RepTo')
    exp = ('[[Expenses.date]],">="&RepFrom,[[Expenses.date]],"<="&RepTo,'
           '[[Expenses.type]],"<>Personal"')
    pl_first = row
    row = metric(row, 0, "Service income collected",
                 f'=SUMIFS([[Payments.amount]],{inc})'
                 f'-SUMIFS([[Payments.refund]],{inc})')
    row = metric(row, 0, "Tips & gratuities",
                 f'=SUMIFS([[Payments.tip]],{inc})')
    row = metric(row, 0, "Late fees collected",
                 f'=SUMIFS([[Payments.late_fee]],{inc})')
    income_row = row
    row = metric(row, 0, "TOTAL INCOME",
                 f'=SUM(B{pl_first + 1}:B{row})')
    row = metric(row, 0, "Wages & contractor payments",
                 f'=SUMIFS([[Expenses.amount]],{exp},'
                 f'[[Expenses.category]],"Staff wages")'
                 f'+SUMIFS([[Expenses.amount]],{exp},'
                 f'[[Expenses.category]],"Contractor payments")')
    wages_row = row
    row = metric(row, 0, "Gross profit",
                 f'=B{income_row + 1}-B{wages_row}')
    gross_row = row
    row = metric(row, 0, "Other operating expenses",
                 f'=SUMIFS([[Expenses.amount]],{exp})'
                 f'-B{wages_row}')
    other_row = row
    row = metric(row, 0, "NET PROFIT",
                 f'=B{gross_row}-B{other_row}')
    profit_row = row
    row = metric(row, 0, "Profit margin",
                 f'=IFERROR(B{profit_row}/B{income_row + 1},0)', "percent")
    row = metric(row, 0, "Tax to set aside",
                 f'=MAX(0,B{profit_row})*TaxReservePct')

    money_right = pl_first
    money_right = metric(money_right, 5, "Invoiced in period",
                         '=SUMIFS([[Invoices.total]],[[Invoices.date]],'
                         '">="&RepFrom,[[Invoices.date]],"<="&RepTo,'
                         '[[Invoices.status]],"<>Cancelled")')
    money_right = metric(money_right, 5, "Collected in period",
                         f'=SUMIFS([[Payments.received]],{inc})')
    money_right = metric(money_right, 5, "Refunds issued",
                         f'=SUMIFS([[Payments.refund]],{inc})')
    money_right = metric(money_right, 5, "Outstanding right now",
                         '=SUMIFS([[Invoices.balance]],[[Invoices.status]],'
                         '"<>Cancelled")')
    money_right = metric(money_right, 5, "Overdue right now",
                         '=SUMIFS([[Invoices.balance]],'
                         '[[Invoices.pay_status]],"Overdue")'
                         '+SUMIFS([[Invoices.balance]],'
                         '[[Invoices.pay_status]],"Overdue - part paid")')
    money_right = metric(money_right, 5, "Tax charged on invoices",
                         '=SUMIFS([[Invoices.tax]],[[Invoices.date]],'
                         '">="&RepFrom,[[Invoices.date]],"<="&RepTo,'
                         '[[Invoices.status]],"<>Cancelled")')
    money_right = metric(money_right, 5, "Tax paid inside expenses",
                         f'=SUMIFS([[Expenses.tax]],{exp})')
    money_right = metric(money_right, 5, "Mileage claim",
                         '=SUMIFS([[Mileage.claimable]],[[Mileage.date]],'
                         '">="&RepFrom,[[Mileage.date]],"<="&RepTo)')
    money_right = metric(money_right, 5, "Business distance travelled",
                         '=SUMIFS([[Mileage.total_km]],[[Mileage.date]],'
                         '">="&RepFrom,[[Mileage.date]],"<="&RepTo,'
                         '[[Mileage.business]],"Business")', "number1")
    row = max(row, money_right) + 1

    # ------------------------------------------------- revenue by service
    row = s.section(row, "3. Revenue & profit by service", LAST_COL)
    row = table(row, 0, ["Service", "Bookings", "Revenue", "Est. cost",
                         "Est. profit", "Margin"])
    first = row
    service_rows = min(16, cfg.row_count("Services"))
    for i in range(service_rows):
        c = FIRST_CALC_ROW + i
        s.calc_row(first + i, 0, [
            f'=IF([[Calc.s_name#{c}]]="","",[[Calc.s_name#{c}]])',
            f'=IF([[Calc.s_name#{c}]]="","",[[Calc.s_count_r#{c}]])',
            f'=IF([[Calc.s_name#{c}]]="","",[[Calc.s_rev_r#{c}]])',
            f'=IF([[Calc.s_name#{c}]]="","",[[Calc.s_cost_r#{c}]])',
            f'=IF([[Calc.s_name#{c}]]="","",[[Calc.s_profit_r#{c}]])',
            f'=IFERROR([[Calc.s_profit_r#{c}]]/[[Calc.s_rev_r#{c}]],"")',
        ], ["text", "int", "money", "money", "money", "percent"])
    ws.conditional_format(first, 2, first + service_rows - 1, 2,
                          {"type": "data_bar", "bar_color": "#2E7D8F"})
    row = first + service_rows + 1

    # --------------------------------------------------- clients & ageing
    row = s.section(row, "4. Clients, outstanding money & invoice ageing",
                    LAST_COL)
    s.merge(row, 0, 3, "Top 10 clients in the period",
            fmts.chrome("subsection"))
    s.merge(row, 5, 8, "Largest outstanding balances",
            fmts.chrome("subsection"))
    row += 1
    table(row, 0, ["#", "Client", "Revenue", "Bookings"])
    table(row, 5, ["#", "Client", "Owed", "Unpaid invoices"])
    first = row + 1
    for i in range(10):
        n = i + 1
        s.calc_row(first + i, 0, [
            n,
            f'=IFERROR(INDEX([[Calc.c_name]],MATCH(LARGE([[Calc.c_key_r]],'
            f'{n}),[[Calc.c_key_r]],0)),"")',
            f'=IFERROR(INDEX([[Calc.c_rev_r]],MATCH(LARGE([[Calc.c_key_r]],'
            f'{n}),[[Calc.c_key_r]],0)),"")',
            f'=IFERROR(INDEX([[Calc.c_bookings_r]],'
            f'MATCH(LARGE([[Calc.c_key_r]],{n}),[[Calc.c_key_r]],0)),"")',
        ], ["int", "text", "money", "int"])
        client_lookup = (f'INDEX([[Calc.c_name]],MATCH(LARGE('
                         f'[[Calc.c_out_key]],{n}),[[Calc.c_out_key]],0))')
        s.calc_row(first + i, 5, [
            n,
            f'=IFERROR({client_lookup},"")',
            f'=IFERROR(INDEX([[Calc.c_out]],MATCH(LARGE('
            f'[[Calc.c_out_key]],{n}),[[Calc.c_out_key]],0)),"")',
            f'=IFERROR(COUNTIFS([[Invoices.client]],{client_lookup},'
            f'[[Invoices.balance]],">0.004"),"")',
        ], ["int", "text", "money", "int"])
    row = first + 11

    s.merge(row, 0, 3, "Invoice ageing", fmts.chrome("subsection"))
    s.merge(row, 5, 8, "How clients pay", fmts.chrome("subsection"))
    row += 1
    table(row, 0, ["Bucket", "Invoices", "Amount", "Share"])
    table(row, 5, ["Method", "Collected", "Share", ""])
    first = row + 1
    buckets = ["Not due", "Current", "1-30 days", "31-60 days",
               "61-90 days", "90+ days"]
    for i, bucket in enumerate(buckets):
        s.calc_row(first + i, 0, [
            bucket,
            f'=COUNTIFS([[Invoices.aging]],"{bucket}")',
            f'=SUMIFS([[Invoices.balance]],[[Invoices.aging]],"{bucket}")',
            f'=IFERROR(SUMIFS([[Invoices.balance]],[[Invoices.aging]],'
            f'"{bucket}")/SUMIFS([[Invoices.balance]],'
            f'[[Invoices.status]],"<>Cancelled"),"")',
        ], ["text", "int", "money", "percent"])
    for i in range(8):
        c = FIRST_CALC_ROW + i
        s.calc_row(first + i, 5, [
            f'=IF([[Calc.pm_method#{c}]]="","",[[Calc.pm_method#{c}]])',
            f'=IF([[Calc.pm_method#{c}]]="","",[[Calc.pm_total#{c}]])',
            f'=IFERROR([[Calc.pm_total#{c}]]/SUM([[Calc.pm_total]]),"")',
            "",
        ], ["text", "money", "percent", "text"])
    row = first + max(len(buckets), 8) + 1

    # ------------------------------------------------ expenses & mileage
    row = s.section(row, "5. Expense report by category", LAST_COL)
    row = table(row, 0, ["Category", "Spend", "Share", "", "", ""])
    first = row
    for i in range(19):
        c = FIRST_CALC_ROW + i
        s.calc_row(first + i, 0, [
            f'=IF([[Calc.e_cat#{c}]]="","",[[Calc.e_cat#{c}]])',
            f'=IF([[Calc.e_cat#{c}]]="","",[[Calc.e_total_r#{c}]])',
            f'=IFERROR([[Calc.e_total_r#{c}]]/SUM([[Calc.e_total_r]]),"")',
            "", "", "",
        ], ["text", "money", "percent", "text", "text", "text"])
    ws.conditional_format(first, 1, first + 18, 1,
                          {"type": "data_bar", "bar_color": "#C0504D"})
    s.calc_row(first + 19, 0, ["Total business expenses",
                               f'=SUM(B{first + 1}:B{first + 19})', "",
                               "", "", ""],
               ["text", "money", "text", "text", "text", "text"])
    row = first + 21

    # --------------------------------------------------- staff & retention
    row = s.section(row, "6. Staff performance", LAST_COL)
    row = table(row, 0, ["Team member", "Visits", "Hours", "Revenue",
                         "Revenue / hour", "Distance"])
    first = row
    staff_rows = min(12, cfg.row_count("Staff"))
    for i in range(staff_rows):
        c = FIRST_CALC_ROW + i
        s.calc_row(first + i, 0, [
            f'=IF([[Calc.st_name#{c}]]="","",[[Calc.st_name#{c}]])',
            f'=IF([[Calc.st_name#{c}]]="","",[[Calc.st_visits_r#{c}]])',
            f'=IF([[Calc.st_name#{c}]]="","",[[Calc.st_hours_r#{c}]])',
            f'=IF([[Calc.st_name#{c}]]="","",[[Calc.st_rev_r#{c}]])',
            f'=IFERROR([[Calc.st_rev_r#{c}]]/[[Calc.st_hours_r#{c}]],"")',
            f'=IF([[Calc.st_name#{c}]]="","",[[Calc.st_km_r#{c}]])',
        ], ["text", "int", "number", "money", "money", "number1"])
    row = first + staff_rows + 1

    row = s.section(row, "7. Client retention & marketing", LAST_COL)
    ret = row
    ret = metric(ret, 0, "Clients served in the period",
                 '=COUNTIF([[Calc.c_bookings_r]],">0")', "int")
    ret = metric(ret, 0, "Clients with repeat bookings",
                 '=COUNTIF([[Calc.c_bookings_r]],">1")', "int")
    ret = metric(ret, 0, "Repeat booking rate",
                 '=IFERROR(COUNTIF([[Calc.c_bookings_r]],">1")'
                 '/COUNTIF([[Calc.c_bookings_r]],">0"),0)', "percent")
    ret = metric(ret, 0, "New clients registered",
                 '=COUNTIFS([[Clients.reg_date]],">="&RepFrom,'
                 '[[Clients.reg_date]],"<="&RepTo)', "int")
    ret = metric(ret, 0, "Returning client rate",
                 '=IFERROR((COUNTIF([[Calc.c_bookings_r]],">0")'
                 '-COUNTIFS([[Clients.reg_date]],">="&RepFrom,'
                 '[[Clients.reg_date]],"<="&RepTo))'
                 '/COUNTIF([[Calc.c_bookings_r]],">0"),0)', "percent")
    ret = metric(ret, 0, "Inactive clients to win back",
                 '=COUNTIF([[Clients.alert]],"Inactive*")', "int")
    ret = metric(ret, 0, "Average client lifetime revenue",
                 '=IFERROR(AVERAGEIF([[Clients.revenue]],">0"),0)')
    ret = metric(ret, 0, "Average satisfaction rating",
                 '=IFERROR(AVERAGEIFS([[Messages.rating]],'
                 '[[Messages.date]],">="&RepFrom,[[Messages.date]],'
                 '"<="&RepTo),"-")', "number")
    ret = metric(ret, 0, "Complaints logged",
                 '=COUNTIFS([[Messages.complaint]],"Yes",'
                 '[[Messages.date]],">="&RepFrom,[[Messages.date]],'
                 '"<="&RepTo)', "int")
    ret = metric(ret, 0, "Complaints still open",
                 '=COUNTIFS([[Messages.complaint]],"Yes",'
                 '[[Messages.resolution]],"<>Resolved")', "int")

    s.merge(row, 5, 8, "Where clients come from", fmts.chrome("subsection"))
    s.table_head(row + 1, 5, ["Referral source", "Clients", "Revenue", ""])
    for i in range(10):
        c = FIRST_CALC_ROW + i
        s.calc_row(row + 2 + i, 5, [
            f'=IF([[Calc.ref_source#{c}]]="","",[[Calc.ref_source#{c}]])',
            f'=IF([[Calc.ref_source#{c}]]="","",[[Calc.ref_clients#{c}]])',
            f'=IF([[Calc.ref_source#{c}]]="","",[[Calc.ref_revenue#{c}]])',
            "",
        ], ["text", "int", "money", "text"])
    row = max(ret, row + 12) + 1

    # ------------------------------------------------ cancellation reasons
    row = s.section(row, "8. Cancellation analysis", LAST_COL)
    row = table(row, 0, ["Reason", "Cancellations", "Share", "", "", ""])
    first = row
    for i in range(10):
        c = FIRST_CALC_ROW + i
        s.calc_row(first + i, 0, [
            f'=IF([[Calc.cr_reason#{c}]]="","",[[Calc.cr_reason#{c}]])',
            f'=IF([[Calc.cr_reason#{c}]]="","",[[Calc.cr_count#{c}]])',
            f'=IFERROR([[Calc.cr_count#{c}]]/SUM([[Calc.cr_count]]),"")',
            "", "", "",
        ], ["text", "int", "percent", "text", "text", "text"])
    row = first + 11

    s.note(row,
           "Printing: this tab is set up to print on A4 at one page wide. "
           "Use File > Print and choose 'Print active sheet'. For a single "
           "report, select the block you need and print the selection.",
           LAST_COL, height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s


def _start_of_year(cfg: Config) -> _dt.date:
    return _dt.date(cfg.today.year, 1, 1)


def _end_of_year(cfg: Config) -> _dt.date:
    return _dt.date(cfg.today.year, 12, 31)
````

### `pet_sitting_spreadsheet/sheets/budget.py`

*272 lines · 1 function*

> `Budget` - targets, budget vs actual, break-even, forecast and cash flow.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Budget'), `LAST_COL` (12), `FIRST_CALC_ROW` (5), `DRAW_ROWS` (24).

````python
"""`Budget` - targets, budget vs actual, break-even, forecast and cash flow."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..sample_data import get as sample
from ..writer import SheetWriter, date_check, number_check

NAME = "Budget"
LAST_COL = 12
FIRST_CALC_ROW = 5
DRAW_ROWS = 24


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.good_text)
    ws = s.ws

    s.banner("Financial Planning & Profitability",
             "Set the targets, then watch budget against actual month by "
             "month. Break-even, forecast, cash flow and tax reserve are "
             "all worked out for you.",
             LAST_COL)

    s.set_col(0, 0, 26)
    s.set_col(1, 12, 13)

    def metric(row, col, label, value=None, formula=None, style="money",
               mode="calc", name=None, validate=None, span=1):
        s.merge(row, col, col + span - 1, label, fmts.chrome("label"))
        target_col = col + span
        fmt = fmts.calc(style) if mode == "calc" else fmts.input(style)
        if formula:
            ws.write_formula(row, target_col, reg.render(formula), fmt)
        else:
            s.write(row, target_col, value, fmt)
        if validate:
            ws.data_validation(row, target_col, row, target_col, validate)
        if name:
            book.define_name(
                name, f"={quote_sheet(NAME)}!"
                      f"${col_letter(target_col)}${row + 1}")
        return row + 1

    # ----------------------------------------------------------- targets
    row = 4
    row = s.section(row, "1. Targets for the dashboard year", LAST_COL)
    row = metric(row, 0, "Annual revenue target", 48000, style="money",
                 mode="input", name="Target_Revenue",
                 validate=number_check(0, 100000000))
    row = metric(row, 0, "Annual expense budget", 19000, style="money",
                 mode="input", name="Target_Expenses",
                 validate=number_check(0, 100000000))
    row = metric(row, 0, "Target profit",
                 formula="=Target_Revenue-Target_Expenses")
    row = metric(row, 0, "Target profit margin",
                 formula="=IFERROR((Target_Revenue-Target_Expenses)"
                         "/Target_Revenue,0)", style="percent")
    row = metric(row, 0, "Fixed costs per month", 650, style="money",
                 mode="input", name="FixedCosts",
                 validate=number_check(0, 1000000))
    row = metric(row, 0, "Average gross margin",
                 formula='=IFERROR(1-SUMIFS([[Services.staff_cost]],'
                         '[[Services.active]],"Yes")'
                         '/SUMIFS([[Services.rate]],[[Services.active]],'
                         '"Yes"),0.5)', style="percent")
    margin_row = row
    row = metric(row, 0, "Break-even revenue per month",
                 formula=f"=IFERROR(FixedCosts/B{margin_row},0)")
    row = metric(row, 0, "Break-even bookings per month",
                 formula=f'=IFERROR(B{margin_row + 1}'
                         f'/IFERROR(AVERAGEIF([[Bookings.total]],">0"),1),0)',
                 style="int")
    row = metric(row, 0, "Growth assumption for forecast", 0.05,
                 style="percent", mode="input", name="GrowthRate",
                 validate=number_check(-1, 5))
    row += 1

    # --------------------------------------------- budget versus actual --
    row = s.section(row, "2. Monthly budget vs actual", LAST_COL)
    headers = ["Month", "Revenue target", "Revenue actual", "Variance",
               "Achieved", "Expense budget", "Expense actual", "Variance",
               "Profit", "Margin", "Cumulative profit", "Forecast revenue"]
    row = s.table_head(row, 0, headers)
    first = row
    for i in range(12):
        r = first + i
        c = FIRST_CALC_ROW + i
        excel_row = r + 1
        prev = excel_row - 1
        ws.write_formula(r, 0, reg.render(f'=[[Calc.m_label#{c}]]'),
                         fmts.calc("text"))
        ws.write_formula(r, 1, "=ROUND(Target_Revenue/12,2)",
                         fmts.calc("money"))
        ws.write_formula(r, 2, reg.render(f'=[[Calc.m_rev#{c}]]'),
                         fmts.calc("money"))
        ws.write_formula(r, 3, f"=C{excel_row}-B{excel_row}",
                         fmts.calc("money"))
        ws.write_formula(r, 4, f'=IFERROR(C{excel_row}/B{excel_row},"")',
                         fmts.calc("percent"))
        ws.write_formula(r, 5, "=ROUND(Target_Expenses/12,2)",
                         fmts.calc("money"))
        ws.write_formula(r, 6, reg.render(f'=[[Calc.m_exp#{c}]]'),
                         fmts.calc("money"))
        ws.write_formula(r, 7, f"=F{excel_row}-G{excel_row}",
                         fmts.calc("money"))
        ws.write_formula(r, 8, f"=C{excel_row}-G{excel_row}",
                         fmts.calc("money"))
        ws.write_formula(r, 9, f'=IFERROR(I{excel_row}/C{excel_row},"")',
                         fmts.calc("percent"))
        ws.write_formula(
            r, 10,
            f"=I{excel_row}" if i == 0 else f"=K{prev}+I{excel_row}",
            fmts.calc("money"))
        ws.write_formula(
            r, 11,
            f'=IFERROR(ROUND(AVERAGE(C{first + 1}:C{excel_row})'
            f'*(1+GrowthRate),2),0)', fmts.calc("money"))

    last = first + 11
    total_row = last + 1
    s.write(total_row, 0, "Year", fmts.chrome("total_label"))
    for idx, style in enumerate(
            ["money", "money", "money", "percent", "money", "money",
             "money", "money", "percent", "money", "money"], start=1):
        letter = col_letter(idx)
        fmt = fmts.chrome("total_value", num_format=s.num_format(style))
        if idx == 4:
            ws.write_formula(total_row, idx,
                             f'=IFERROR(C{total_row + 1}/B{total_row + 1},"")',
                             fmt)
        elif idx == 9:
            ws.write_formula(total_row, idx,
                             f'=IFERROR(I{total_row + 1}/C{total_row + 1},"")',
                             fmt)
        elif idx == 10:
            ws.write_formula(total_row, idx, f"=K{last + 1}", fmt)
        else:
            ws.write_formula(total_row, idx,
                             f"=SUM({letter}{first + 1}:{letter}{last + 1})",
                             fmt)
    for col_idx, colour in ((2, "#2E7D8F"), (6, "#C0504D"), (8, "#1F4E5F")):
        letter = col_letter(col_idx)
        try:
            ws.add_sparkline(total_row + 1, col_idx, {
                "range": (f"{quote_sheet(NAME)}!{letter}{first + 1}:"
                          f"{letter}{last + 1}"),
                "type": "column", "series_color": colour})
        except Exception:        # pragma: no cover
            pass
    s.write(total_row + 1, 0, "Trend", fmts.chrome("label_bold"))
    row = total_row + 3

    # ------------------------------------------------------- cash summary
    row = s.section(row, "3. Cash flow & owner money", LAST_COL)
    left = row
    left = metric(left, 0, "Opening cash balance", 1200, style="money",
                  mode="input", name="OpeningCash",
                  validate=number_check(-1000000, 10000000))
    left = metric(left, 0, "Cash in (payments received)",
                  formula='=SUMIFS([[Payments.received]],[[Payments.date]],'
                          '">="&DATE(Dash_Year,1,1),[[Payments.date]],'
                          '"<="&DATE(Dash_Year,12,31))')
    left = metric(left, 0, "Cash out (expenses paid)",
                  formula='=SUMIFS([[Expenses.amount]],[[Expenses.date]],'
                          '">="&DATE(Dash_Year,1,1),[[Expenses.date]],'
                          '"<="&DATE(Dash_Year,12,31))')
    left = metric(left, 0, "Owner withdrawals",
                  formula=f"=SUM(K{row + 3}:K{row + 2 + DRAW_ROWS})")
    left = metric(left, 0, "Closing cash balance",
                  formula=f"=B{row + 1}+B{row + 2}-B{row + 3}-B{row + 4}")
    left = metric(left, 0, "Money to set aside for tax",
                  formula='=MAX(0,SUMIFS([[Payments.received]],'
                          '[[Payments.date]],">="&DATE(Dash_Year,1,1),'
                          '[[Payments.date]],"<="&DATE(Dash_Year,12,31))'
                          '-SUMIFS([[Expenses.amount]],[[Expenses.date]],'
                          '">="&DATE(Dash_Year,1,1),[[Expenses.date]],'
                          '"<="&DATE(Dash_Year,12,31)))*TaxReservePct')
    left = metric(left, 0, "Business savings target", 2500, style="money",
                  mode="input", name="SavingsTarget",
                  validate=number_check(0, 10000000))
    left = metric(left, 0, "Savings set aside so far", 900, style="money",
                  mode="input", name="SavingsActual",
                  validate=number_check(0, 10000000))
    left = metric(left, 0, "Savings progress",
                  formula="=IFERROR(SavingsActual/SavingsTarget,0)",
                  style="percent")

    # Owner drawings register (right hand side).
    s.merge(row, 9, 12, "Owner withdrawals register",
            fmts.chrome("subsection"))
    s.table_head(row + 1, 9, ["Date", "Amount", "Notes", ""])
    draws = sample(cfg).drawings if cfg.include_sample_data else []
    for i in range(DRAW_ROWS):
        r = row + 2 + i
        date_fmt = fmts.input("date")
        money_fmt = fmts.input("money")
        text_fmt = fmts.input("text")
        if i < len(draws):
            s.write(r, 9, draws[i]["date"], date_fmt)
            s.write(r, 10, draws[i]["amount"], money_fmt)
            s.write(r, 11, draws[i]["notes"], text_fmt)
        else:
            ws.write_blank(r, 9, None, date_fmt)
            ws.write_blank(r, 10, None, money_fmt)
            ws.write_blank(r, 11, None, text_fmt)
        ws.write_blank(r, 12, None, text_fmt)
    ws.data_validation(row + 2, 9, row + 1 + DRAW_ROWS, 9, date_check())
    ws.data_validation(row + 2, 10, row + 1 + DRAW_ROWS, 10,
                       number_check(0, 1000000))
    row = max(left, row + 2 + DRAW_ROWS) + 1

    # -------------------------------------------------- year on year block
    row = s.section(row, "4. Year on year comparison", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "This year", "Last year",
                                "Change", "Change %", "", "", "", "", "",
                                "", "", ""])
    yoy = [
        ("Revenue collected",
         '=SUMIFS([[Payments.received]],[[Payments.date]],'
         '">="&DATE(Dash_Year,1,1),[[Payments.date]],'
         '"<="&DATE(Dash_Year,12,31))',
         '=SUMIFS([[Payments.received]],[[Payments.date]],'
         '">="&DATE(Dash_Year-1,1,1),[[Payments.date]],'
         '"<="&DATE(Dash_Year-1,12,31))', "money"),
        ("Expenses",
         '=SUMIFS([[Expenses.amount]],[[Expenses.date]],'
         '">="&DATE(Dash_Year,1,1),[[Expenses.date]],'
         '"<="&DATE(Dash_Year,12,31),[[Expenses.type]],"<>Personal")',
         '=SUMIFS([[Expenses.amount]],[[Expenses.date]],'
         '">="&DATE(Dash_Year-1,1,1),[[Expenses.date]],'
         '"<="&DATE(Dash_Year-1,12,31),[[Expenses.type]],"<>Personal")',
         "money"),
        ("Bookings",
         '=COUNTIFS([[Bookings.date]],">="&DATE(Dash_Year,1,1),'
         '[[Bookings.date]],"<="&DATE(Dash_Year,12,31),'
         '[[Bookings.status]],"<>Cancelled")',
         '=COUNTIFS([[Bookings.date]],">="&DATE(Dash_Year-1,1,1),'
         '[[Bookings.date]],"<="&DATE(Dash_Year-1,12,31),'
         '[[Bookings.status]],"<>Cancelled")', "int"),
        ("New clients",
         '=COUNTIFS([[Clients.reg_date]],">="&DATE(Dash_Year,1,1),'
         '[[Clients.reg_date]],"<="&DATE(Dash_Year,12,31))',
         '=COUNTIFS([[Clients.reg_date]],">="&DATE(Dash_Year-1,1,1),'
         '[[Clients.reg_date]],"<="&DATE(Dash_Year-1,12,31))', "int"),
        ("Business distance",
         '=SUMIFS([[Mileage.total_km]],[[Mileage.date]],'
         '">="&DATE(Dash_Year,1,1),[[Mileage.date]],'
         '"<="&DATE(Dash_Year,12,31),[[Mileage.business]],"Business")',
         '=SUMIFS([[Mileage.total_km]],[[Mileage.date]],'
         '">="&DATE(Dash_Year-1,1,1),[[Mileage.date]],'
         '"<="&DATE(Dash_Year-1,12,31),[[Mileage.business]],"Business")',
         "number1"),
    ]
    for i, (label, this_year, last_year, style) in enumerate(yoy):
        r = row + i
        excel_row = r + 1
        s.calc_row(r, 0, [label, this_year, last_year,
                          f"=B{excel_row}-C{excel_row}",
                          f'=IFERROR(B{excel_row}/C{excel_row}-1,"")'],
                   ["text", style, style, style, "percent"])
    row += len(yoy) + 1

    s.note(row,
           "How the forecast works: it averages the actual revenue "
           "recorded so far this year and adds your growth assumption. "
           "Update the growth rate as the year unfolds.",
           LAST_COL, height=24)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `pet_sitting_spreadsheet/sheets/start_here.py`

*219 lines · 1 function*

> `Start Here` - welcome page, navigation menu and setup checklist.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Start Here'), `LAST_COL` (9), `TOUR` (25 items), `QUICK_START` (10 items), `CHECKLIST` (12 items).

````python
"""`Start Here` - welcome page, navigation menu and setup checklist."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from

NAME = "Start Here"
LAST_COL = 9

TOUR = [
    ("Settings", "Your business details, rates, tax, surcharges and "
                 "reminder windows - set this up first."),
    ("Lists", "Every drop down list in the workbook. Rename anything to "
              "match the way you talk about your services."),
    ("Services", "Your price list. Billing unit, rate, extra pet fee and "
                 "profit per service."),
    ("Clients", "Client database with keys, contracts, lifetime revenue "
                "and follow-up flags."),
    ("Pets", "Pet profiles: feeding, medication, vet, vaccinations, "
             "temperament and safety warnings."),
    ("Bookings", "The schedule. Prices itself from the Services tab and "
                 "warns you about double bookings."),
    ("Visit Log", "Proof of every walk and visit: times, care given, mood, "
                  "health notes and client updates."),
    ("Health Log", "Medication administration, vaccinations and vet "
                   "appointments."),
    ("Incidents", "Emergency contacts and incident reports."),
    ("Messages", "Client communication, feedback, ratings and complaints."),
    ("Invoices", "Automatic invoice numbers, due dates, tax, balances and "
                 "ageing."),
    ("Payments", "Cash in: payments, deposits, tips, refunds and bank "
                 "reconciliation."),
    ("Expenses", "Every business cost split into tax, net and claimable."),
    ("Mileage", "Trip log with odometer readings and reimbursement."),
    ("Staff", "Team directory, availability, hours worked and pay due."),
    ("Inventory", "Supplies, stock levels, reorder and expiry alerts."),
    ("Dashboard", "The whole business on one page for any month or year."),
    ("Reports", "Printable analytics for any date range - including a "
                "profit and loss statement."),
    ("Budget", "Targets, budget vs actual, break-even, forecast and cash "
               "flow."),
    ("Reminders", "Everything that needs doing today, plus your own task "
                  "list."),
    ("Print Invoice", "A client-ready invoice - pick the invoice number."),
    ("Print Pet Care", "A one page care sheet to leave with a sitter."),
    ("Print Visit", "A single visit report for the client."),
    ("Print Forms", "Blank forms: registration, consent, incident, "
                    "timesheet, mileage, feedback and checklists."),
    ("Help", "Instructions, formula notes, troubleshooting and version "
             "history."),
]

QUICK_START = [
    "1. Open the Settings tab and enter your business details, currency, "
    "tax rate and mileage rate.",
    "2. Edit the Services tab so the names and prices match what you "
    "actually sell.",
    "3. Add your clients, then their pets (choose the owner from the drop "
    "down so they link up).",
    "4. Add your team on the Staff tab - even if that is only you.",
    "5. Take a booking: pick the client, pet and service and the price "
    "calculates itself.",
    "6. After each visit, add a row to the Visit Log and tick off the care "
    "you gave.",
    "7. Raise an invoice, link the booking, and record the payment when it "
    "lands.",
    "8. Log expenses and trips as you go - future you will thank you at "
    "tax time.",
    "9. Check the Reminders tab every morning and the Dashboard every "
    "week.",
    "10. Print the reports you need at month end and save a dated backup.",
]

CHECKLIST = [
    "Business details added on the Settings tab",
    "Currency, tax rate and payment terms checked",
    "Mileage rate and travel fees set",
    "Service names and prices updated",
    "Drop down lists adjusted to suit my business",
    "Public holidays added for this year",
    "First clients entered",
    "Pets added and linked to their owners",
    "Staff or contractors added",
    "Sample data deleted (rows 5 onwards on each tab)",
    "Invoice number and prefix set the way I want",
    "A backup copy saved somewhere safe",
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent,
                    landscape=False)
    ws = s.ws

    s.banner(cfg.product_name, cfg.product_subtitle, LAST_COL, nav=False,
             badge=f"{cfg.author} - v{cfg.version}")

    s.set_col(0, 0, 3)
    s.set_col(1, 1, 22)
    s.set_col(2, 7, 13)
    s.set_col(8, 9, 14)

    row = 3
    ws.merge_range(row, 0, row, LAST_COL, "", fmts.chrome("note_box"))
    ws.write_formula(
        row, 0,
        '="Welcome "&Biz_Owner&" - this workbook is set up for "&Biz_Name'
        '&". Cream cells are yours to type in, grey cells are locked '
        'formulas that do the maths for you."',
        fmts.chrome("note_box"))
    ws.set_row(row, 24)
    row += 2

    # ------------------------------------------------------- quick start
    row = s.section(row, "Quick start - ten minutes to a working workbook",
                    LAST_COL)
    row = s.bullets(row, QUICK_START, LAST_COL, first_col=1, height=16)
    row += 1

    # ---------------------------------------------------------- navigation
    row = s.section(row, "Where everything lives", LAST_COL)
    header = row
    s.write(header, 1, "Tab", fmts.chrome("header_left"))
    s.merge(header, 2, LAST_COL, "What it does", fmts.chrome("header"))
    ws.set_row(header, 20)
    row = header + 1
    link_fmt = fmts.chrome("nav")
    note_fmt = fmts.chrome("note")
    for sheet_name, description in TOUR:
        ws.write_url(row, 1, f"internal:{quote_sheet(sheet_name)}!A1",
                     link_fmt, sheet_name)
        s.merge(row, 2, LAST_COL, description, note_fmt)
        ws.set_row(row, 15)
        row += 1
    row += 1

    # ----------------------------------------------------------- checklist
    row = s.section(row, "Getting started checklist", LAST_COL)
    head = row
    s.write(head, 1, "Step", fmts.chrome("header_left"))
    s.merge(head, 2, 6, "What to do", fmts.chrome("header"))
    s.merge(head, 7, 8, "Done?", fmts.chrome("header"))
    ws.set_row(head, 20)
    first = head + 1
    for i, item in enumerate(CHECKLIST):
        r = first + i
        s.write(r, 1, f"Step {i + 1}", fmts.chrome("label"))
        s.merge(r, 2, 6, item, fmts.chrome("label"))
        ws.merge_range(r, 7, r, 8, "No", fmts.input("flag"))
        ws.data_validation(r, 7, r, 8, list_from("YesNoList"))
        ws.conditional_format(r, 7, r, 8, {
            "type": "text", "criteria": "containing", "value": "Yes",
            "format": fmts.cf("good")})
    last = first + len(CHECKLIST) - 1

    progress_row = last + 1
    s.merge(progress_row, 1, 6, "Setup progress",
            fmts.chrome("total_label"))
    fmt = fmts.chrome("total_value", num_format="0%")
    ws.merge_range(progress_row, 7, progress_row, 8, "", fmt)
    ws.write_formula(
        progress_row, 7,
        f'=COUNTIF($H${first + 1}:$H${last + 1},"Yes")/{len(CHECKLIST)}',
        fmt)
    row = progress_row + 2

    # -------------------------------------------------------- colour key
    row = s.section(row, "How to read the colours", LAST_COL)
    legend = [
        ("legend_input", "Cream cell",
         "Type here. These cells stay unlocked while the sheet is "
         "protected."),
        ("legend_calc", "Grey cell",
         "A locked formula. Protected with the password so it cannot be "
         "overwritten by accident."),
    ]
    for key, label, text in legend:
        s.write(row, 1, label, fmts.chrome(key))
        s.merge(row, 2, LAST_COL, text, note_fmt)
        row += 1
    s.write(row, 1, "Red / amber", fmts.chrome("legend_calc",
                                               bg_color=Palette.bad_bg))
    s.merge(row, 2, LAST_COL,
            "Conditional formatting: red means act today, amber means "
            "act soon, green means you are on top of it.", note_fmt)
    row += 2

    # ------------------------------------------------------------- facts
    row = s.section(row, "Product information", LAST_COL)
    facts = [
        ("Product", cfg.product_name),
        ("Edition", f"Version {cfg.version} - released {cfg.release_date}"),
        ("Author", cfg.author),
        ("Formula protection",
         f"Every sheet is protected. Password: {cfg.password} "
         "(Review > Unprotect Sheet if you need to change a formula)."),
        ("Works with",
         "Microsoft Excel 2010 or newer on Windows and Mac, Excel 365, "
         "and Google Sheets after import."),
        ("Support",
         "See the Help tab for troubleshooting, formula notes and backup "
         "guidance."),
    ]
    for label, text in facts:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, note_fmt)
        ws.set_row(row, 15)
        row += 1

    ws.set_row(row + 1, 18)
    s.merge(row + 1, 0, LAST_COL,
            f"(c) {cfg.author} - thank you for your purchase. "
            "Personal and single-business use.",
            fmts.chrome("subtitle"))

    s.finish(freeze_row=3, gridlines=False, fit_width=True)
    return s
````

### `pet_sitting_spreadsheet/sheets/printables.py`

*751 lines · 10 functions*

> The four print-ready tabs: invoice, pet care sheet, visit report, forms.
>
> Each one is driven by a single drop down at the top (which sits outside the
> print area), so the user picks a record and presses print.

| Defines | What it does |
|---------|--------------|
| `_pick(sheet, fmts, label, list_name, define_as, hint, default)` | Drop down selector on row 1 |
| `_row(sheet, fmts, row, label, formula, style, label_cols, value_cols, value, bold)` | `Label: value` line used all over the print sheets |
| `_first_invoice(cfg)` |  |
| `_first_pet(cfg)` |  |
| `_first_visit(cfg)` |  |
| `_index(sheet, field, key, key_field, fallback)` | INDEX/MATCH against a register, written in template syntax |
| `build_invoice(book, fmts, cfg, reg)` |  |
| `build_care_sheet(book, fmts, cfg, reg)` |  |
| `build_visit_report(book, fmts, cfg, reg)` |  |
| `build_forms(book, fmts, cfg, reg)` |  |

**Module constants:** `INVOICE_SHEET` ('Print Invoice'), `CARE_SHEET` ('Print Pet Care'), `VISIT_SHEET` ('Print Visit'), `FORMS_SHEET` ('Print Forms'), `LAST_COL` (7), `FORMS` (10 items), `MEET_ITEMS` (14 items).

````python
"""The four print-ready tabs: invoice, pet care sheet, visit report, forms.

Each one is driven by a single drop down at the top (which sits outside the
print area), so the user picks a record and presses print.
"""

from __future__ import annotations

from typing import List, Optional, Sequence, Tuple

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..sample_data import get as sample
from ..writer import SheetWriter

INVOICE_SHEET = "Print Invoice"
CARE_SHEET = "Print Pet Care"
VISIT_SHEET = "Print Visit"
FORMS_SHEET = "Print Forms"

LAST_COL = 7          # column H - a comfortable portrait page


# ---------------------------------------------------------------------------
# shared helpers
# ---------------------------------------------------------------------------

def _pick(sheet: SheetWriter, fmts, label: str, list_name: str,
          define_as: str, hint: str, default: str = "") -> str:
    """Drop down selector on row 1. Returns the absolute cell reference."""
    ws = sheet.ws
    sheet.merge(0, 0, 1, label, fmts.chrome("label_bold"))
    fmt = fmts.input("text", bold=True, font_size=11)
    ws.merge_range(0, 2, 0, 3, default, fmt)
    ws.data_validation(0, 2, 0, 3, {
        "validate": "list", "source": f"={list_name}",
        "input_title": "Choose a record",
        "input_message": hint[:255]})
    sheet.merge(0, 4, LAST_COL, hint, fmts.chrome("note"))
    sheet.row_height(0, 22)
    sheet.book.define_name(
        define_as, f"={quote_sheet(sheet.name)}!$C$1")
    return define_as


def _row(sheet, fmts, row: int, label: str, formula: Optional[str] = None,
         style: str = "text", label_cols: Tuple[int, int] = (0, 1),
         value_cols: Tuple[int, int] = (2, 3), value: str = "",
         bold: bool = False):
    """`Label: value` line used all over the print sheets."""
    sheet.merge(row, label_cols[0], label_cols[1], label,
                fmts.chrome("label_bold" if bold else "label"))
    fmt = fmts.calc(style) if formula else fmts.chrome("label")
    if formula:
        if value_cols[1] > value_cols[0]:
            sheet.ws.merge_range(row, value_cols[0], row, value_cols[1], "",
                                 fmt)
        sheet.ws.write_formula(row, value_cols[0],
                               sheet.reg.render(formula), fmt)
    else:
        sheet.merge(row, value_cols[0], value_cols[1], value, fmt)
    return row + 1


def _first_invoice(cfg: Config) -> str:
    if not cfg.include_sample_data:
        return ""
    return f"{cfg.invoice_prefix}{cfg.invoice_start_number:04d}"


def _first_pet(cfg: Config) -> str:
    pets = sample(cfg).pets if cfg.include_sample_data else []
    if not pets:
        return ""
    return f"{pets[0]['name']} ({pets[0]['owner']})"


def _first_visit(cfg: Config) -> str:
    return "VS-0001" if cfg.include_sample_data else ""


def _index(sheet: str, field: str, key: str, key_field: str,
           fallback: str = '""') -> str:
    """INDEX/MATCH against a register, written in template syntax."""
    return (f'IFERROR(INDEX([[{sheet}.{field}]],'
            f'MATCH({key},[[{sheet}.{key_field}]],0)),{fallback})')


# ---------------------------------------------------------------------------
# 1. Print Invoice
# ---------------------------------------------------------------------------

def build_invoice(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, INVOICE_SHEET,
                    tab_color=Palette.primary_soft, landscape=False)
    ws = s.ws
    s.set_col(0, 0, 14)
    s.set_col(1, 1, 16)
    s.set_col(2, 2, 14)
    s.set_col(3, 3, 12)
    s.set_col(4, 4, 8)
    s.set_col(5, 5, 10)
    s.set_col(6, 6, 12)
    s.set_col(7, 7, 14)

    _pick(s, fmts, "Invoice to print", "InvoiceNos", "PrintInvoiceNo",
          "Pick an invoice number - everything below fills in "
          "automatically. This row does not print.",
          default=_first_invoice(cfg))

    def inv(field: str, fallback: str = '""') -> str:
        return _index("Invoices", field, "PrintInvoiceNo", "invoice_no",
                      fallback)

    client = inv("client")
    booking = inv("booking")

    def cli(field: str) -> str:
        return _index("Clients", field, f"({client})", "name")

    def bkg(field: str) -> str:
        return _index("Bookings", field, f"({booking})", "id")

    row = 2
    s.merge(row, 0, 3, "", fmts.chrome("print_title"))
    ws.write_formula(row, 0, "=Biz_Name", fmts.chrome("print_title"))
    s.merge(row, 5, LAST_COL, "INVOICE",
            fmts.chrome("print_title", align="right",
                        font_color=Palette.accent, font_size=22))
    s.row_height(row, 30)

    left = row + 1
    for formula in ('=Biz_Address', '=Biz_City&" "&Biz_State&" "&Biz_Post',
                    '=Biz_Phone', '=Biz_Email',
                    '=IF(Biz_TaxID="","",TaxName&" no: "&Biz_TaxID)'):
        s.merge(left, 0, 3, "", fmts.chrome("print_sub"))
        ws.write_formula(left, 0, formula, fmts.chrome("print_sub"))
        left += 1

    right = row + 1
    meta = [("Invoice no.", f"={inv('invoice_no')}", "text"),
            ("Invoice date", f"={inv('date')}", "date"),
            ("Due date", f"={inv('due')}", "date"),
            ("Status", f"={inv('pay_status')}", "text"),
            ("Amount due", f"={inv('balance',0)}", "money")]
    for label, formula, style in meta:
        s.merge(right, 4, 5, label, fmts.chrome("label_bold"))
        fmt = fmts.calc(style, align="right")
        ws.merge_range(right, 6, right, LAST_COL, "", fmt)
        ws.write_formula(right, 6, reg.render(formula), fmt)
        right += 1

    row = max(left, right) + 1

    s.merge(row, 0, 3, "BILL TO", fmts.chrome("subsection"))
    s.merge(row, 4, LAST_COL, "SERVICE", fmts.chrome("subsection"))
    row += 1
    block = row
    for formula in (f"={client}", f"={cli('address')}",
                    f'={cli("city")}&" "&{cli("postcode")}',
                    f"={cli('phone')}", f"={cli('email')}"):
        s.merge(block, 0, 3, "", fmts.calc("text"))
        ws.write_formula(block, 0, reg.render(formula), fmts.calc("text"))
        block += 1

    right = row
    service_meta = [
        ("Booking ref", f"={booking}", "text"),
        ("Service", f"={bkg('service')}", "text"),
        ("Pet", f"={bkg('pet')}", "text"),
        ("Service date", f"={bkg('date')}", "date"),
        ("Carer", f"={bkg('staff')}", "text"),
    ]
    for label, formula, style in service_meta:
        s.merge(right, 4, 5, label, fmts.chrome("label_bold"))
        fmt = fmts.calc(style, align="right")
        ws.merge_range(right, 6, right, LAST_COL, "", fmt)
        ws.write_formula(right, 6, reg.render(formula), fmt)
        right += 1
    row = max(block, right) + 1

    # ------------------------------------------------------- line items
    s.merge(row, 0, 3, "Description", fmts.chrome("header_left"))
    s.write(row, 4, "Qty", fmts.chrome("header"))
    s.write(row, 5, "Rate", fmts.chrome("header"))
    s.merge(row, 6, LAST_COL, "Amount", fmts.chrome("header"))
    s.row_height(row, 20)
    row += 1

    s.merge(row, 0, 3, "", fmts.calc("wrap"))
    ws.write_formula(
        row, 0,
        reg.render(f'=IF({inv("description")}="",'
                   f'{bkg("service")},{inv("description")})'),
        fmts.calc("wrap"))
    ws.write_formula(row, 4, reg.render(f'=IF({inv("qty",0)}=0,1,'
                                        f'{inv("qty",0)})'),
                     fmts.calc("number"))
    ws.write_formula(row, 5, reg.render(f'=IF({inv("rate",0)}=0,'
                                        f'{inv("line_total",0)},'
                                        f'{inv("rate",0)})'),
                     fmts.calc("money"))
    ws.merge_range(row, 6, row, LAST_COL, "", fmts.calc("money"))
    ws.write_formula(row, 6, reg.render(f'={inv("line_total",0)}'),
                     fmts.calc("money"))
    row += 1

    s.merge(row, 0, LAST_COL, "Visits covered by this invoice",
            fmts.chrome("note"))
    row += 1
    for n in range(1, 9):
        key = f'({booking})&"|"&{n}'
        vis = (lambda field, key=key:
               f'IFERROR(INDEX([[Visit Log.{field}]],'
               f'MATCH({key},[[Visit Log.line_key]],0)),"")')
        s.merge(row, 0, 3, "", fmts.calc("text"))
        ws.write_formula(
            row, 0,
            reg.render(f'=IF({vis("date")}="","","Visit on "'
                       f'&TEXT({vis("date")},"{cfg.date_format}")'
                       f'&" - "&{vis("service")}'
                       f'&IF({vis("staff")}=""," "," with "&{vis("staff")}))'),
            fmts.calc("text"))
        ws.write_formula(
            row, 4,
            reg.render(f'=IF({vis("duration")}="","",'
                       f'{vis("duration")}&" min")'),
            fmts.calc("text", align="right"))
        s.merge(row, 5, LAST_COL, "", fmts.calc("text"))
        row += 1

    row += 1
    totals = [
        ("Subtotal", f'={inv("subtotal",0)}', "money"),
        ("Additional fees", f'={inv("fees",0)}', "money"),
        ("Discount", f'=-{inv("discount",0)}', "money"),
        ("Tax", f'={inv("tax",0)}', "money"),
        ("TOTAL", f'={inv("total",0)}', "money"),
        ("Deposit / paid", f'=-{inv("paid",0)}', "money"),
    ]
    for label, formula, style in totals:
        s.merge(row, 4, 5, label, fmts.chrome("label_bold"))
        fmt = fmts.calc(style, align="right")
        ws.merge_range(row, 6, row, LAST_COL, "", fmt)
        ws.write_formula(row, 6, reg.render(formula), fmt)
        row += 1
    s.merge(row, 4, 5, "BALANCE DUE", fmts.chrome("total_label"))
    fmt = fmts.chrome("total_value", num_format=s.num_format("money"),
                      font_size=12)
    ws.merge_range(row, 6, row, LAST_COL, "", fmt)
    ws.write_formula(row, 6, reg.render(f'={inv("balance",0)}'), fmt)
    s.row_height(row, 22)
    row += 2

    s.merge(row, 0, LAST_COL, "HOW TO PAY", fmts.chrome("subsection"))
    row += 1
    row = _row(s, fmts, row, "Bank details", "=Biz_Bank",
               value_cols=(2, LAST_COL))
    row = _row(s, fmts, row, "Payment terms",
               '="Payment is due within "&PaymentTermsDays&" days '
               'of the invoice date."', value_cols=(2, LAST_COL))
    row = _row(s, fmts, row, "Late payment",
               '=IF(LateFeePct=0,"",TEXT(LateFeePct,"0%")&'
               '" late fee may be applied to overdue balances.")',
               value_cols=(2, LAST_COL))
    row = _row(s, fmts, row, "Notes", f'={inv("notes")}',
               value_cols=(2, LAST_COL))
    row += 1
    s.merge(row, 0, LAST_COL, "", fmts.chrome("note_box"))
    ws.write_formula(
        row, 0,
        '="Thank you for trusting "&Biz_Name&" with your pets. '
        'Questions about this invoice? Call "&Biz_Phone&" or email "'
        '&Biz_Email&"."', fmts.chrome("note_box"))
    s.row_height(row, 26)

    ws.print_area(2, 0, row, LAST_COL)
    ws.fit_to_pages(1, 1)
    ws.center_horizontally()
    ws.hide_gridlines(2)
    s.protect()
    return s


# ---------------------------------------------------------------------------
# 2. Print Pet Care sheet
# ---------------------------------------------------------------------------

def build_care_sheet(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, CARE_SHEET,
                    tab_color=Palette.primary_soft, landscape=False)
    ws = s.ws
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 16)
    s.set_col(2, 2, 14)
    s.set_col(3, 3, 12)
    s.set_col(4, 4, 14)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 12)
    s.set_col(7, 7, 14)

    _pick(s, fmts, "Pet to print", "PetLabels", "PrintPetLabel",
          "Pick a pet reference - the care sheet fills itself in. "
          "This row does not print.", default=_first_pet(cfg))

    def pet(field: str, fallback: str = '""') -> str:
        return _index("Pets", field, "PrintPetLabel", "label", fallback)

    owner = pet("owner")

    def cli(field: str) -> str:
        return _index("Clients", field, f"({owner})", "name")

    row = 2
    s.merge(row, 0, LAST_COL, "", fmts.chrome("print_title"))
    ws.write_formula(
        row, 0,
        reg.render(f'="Pet Care Sheet - "&{pet("name")}'),
        fmts.chrome("print_title"))
    s.row_height(row, 28)
    row += 1
    s.merge(row, 0, LAST_COL, "", fmts.chrome("print_sub"))
    ws.write_formula(
        row, 0,
        '="Prepared by "&Biz_Name&" - "&Biz_Phone&" - printed "'
        '&TEXT(TODAY(),"' + cfg.date_format + '")',
        fmts.chrome("print_sub"))
    row += 2

    def block(start: int, title: str, fields: Sequence[Tuple[str, str, str]],
              columns: int = 2) -> int:
        s.merge(start, 0, LAST_COL, title, fmts.chrome("subsection"))
        start += 1
        half = (len(fields) + columns - 1) // columns
        for i, (label, formula, style) in enumerate(fields):
            col_block = i // half if columns > 1 else 0
            r = start + (i % half)
            base = 0 if col_block == 0 else 4
            s.merge(r, base, base + 1, label, fmts.chrome("label_bold"))
            fmt = fmts.calc(style)
            end = base + 3 if base == 0 else LAST_COL
            ws.merge_range(r, base + 2, r, end, "", fmt)
            ws.write_formula(r, base + 2, reg.render(formula), fmt)
        return start + half + 1

    row = block(row, "PET DETAILS", [
        ("Pet name", f'={pet("name")}', "text"),
        ("Type / breed", f'={pet("type")}&" - "&{pet("breed")}', "text"),
        ("Age", f'=IF({pet("age")}="","",{pet("age")}&" years")', "text"),
        ("Gender", f'={pet("gender")}', "text"),
        ("Weight", f'=IF({pet("weight")}="","",{pet("weight")}&" kg")',
         "text"),
        ("Colour & markings", f'={pet("colour")}', "text"),
        ("Microchip", f'={pet("microchip")}', "text"),
        ("Temperament", f'={pet("temperament")}', "text"),
    ])

    row = block(row, "OWNER & EMERGENCY CONTACTS", [
        ("Owner", f'={owner}', "text"),
        ("Phone", f'={cli("phone")}', "text"),
        ("Email", f'={cli("email")}', "text"),
        ("Address", f'={cli("address")}&" "&{cli("city")}', "text"),
        ("Emergency contact", f'={cli("emergency_name")}', "text"),
        ("Emergency phone", f'={cli("emergency_phone")}', "text"),
        ("Home access", f'={cli("key_status")}', "text"),
        ("Access notes", f'={cli("access_notes")}', "text"),
    ])

    row = block(row, "VET & MEDICAL", [
        ("Veterinarian", f'={pet("vet_name")}', "text"),
        ("Vet phone", f'={pet("vet_phone")}', "text"),
        ("Emergency clinic", f'={pet("emerg_clinic")}', "text"),
        ("Clinic phone", f'={pet("emerg_phone")}', "text"),
        ("Vaccination status", f'={pet("vacc_status")}', "text"),
        ("Vaccination expires", f'={pet("vacc_expiry")}', "date"),
        ("Insurance", f'={pet("insurer")}&" "&{pet("policy")}', "text"),
        ("Vet spend authority", f'={cli("vet_limit")}', "money"),
    ])

    row = block(row, "HEALTH, ALLERGIES & MEDICATION", [
        ("Allergies", f'={pet("allergies")}', "text"),
        ("Medical conditions", f'={pet("medical")}', "text"),
        ("Medication", f'={pet("medication")}', "text"),
        ("Dosage", f'={pet("dosage")}', "text"),
        ("Medication times", f'={pet("med_schedule")}', "text"),
        ("Emergency vet consent", f'={cli("vet_consent")}', "text"),
    ])

    row = block(row, "DAILY ROUTINE", [
        ("Food", f'={pet("food")}', "text"),
        ("Portion", f'={pet("portion")}', "text"),
        ("Feeding times", f'={pet("feed_schedule")}', "text"),
        ("Feeding notes", f'={pet("feeding")}', "text"),
        ("Walking & exercise", f'={pet("walking")}', "text"),
        ("Potty / litter", f'={pet("potty")}', "text"),
        ("Grooming", f'={pet("grooming")}', "text"),
        ("Favourites", f'={pet("favourites")}', "text"),
    ])

    s.merge(row, 0, LAST_COL, "SAFETY & HANDLING", fmts.chrome("subsection"))
    row += 1
    for label, formula in (("Risk flag", f'={pet("risk")}'),
                           ("Behaviour issues", f'={pet("behaviour")}'),
                           ("Handling instructions", f'={pet("handling")}'),
                           ("Safety warnings", f'={pet("warnings")}')):
        s.merge(row, 0, 1, label, fmts.chrome("label_bold"))
        fmt = fmts.calc("wrap")
        ws.merge_range(row, 2, row, LAST_COL, "", fmt)
        ws.write_formula(row, 2, reg.render(formula), fmt)
        s.row_height(row, 20)
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL, "RECENT VISITS", fmts.chrome("subsection"))
    row += 1
    s.merge(row, 0, 1, "Date", fmts.chrome("header_left"))
    s.merge(row, 2, 3, "Service", fmts.chrome("header"))
    s.merge(row, 4, 5, "Carer", fmts.chrome("header"))
    s.merge(row, 6, LAST_COL, "Notes", fmts.chrome("header"))
    row += 1
    count = 'COUNTIF([[Visit Log.pet]],PrintPetLabel)'
    for i in range(6):
        key = f'PrintPetLabel&"|"&({count}-{i})'
        vis = (lambda field, key=key:
               f'IFERROR(INDEX([[Visit Log.{field}]],'
               f'MATCH({key},[[Visit Log.pet_key]],0)),"")')
        for label_cols, field, style in (((0, 1), "date", "date"),
                                         ((2, 3), "service", "text"),
                                         ((4, 5), "staff", "text"),
                                         ((6, LAST_COL), "notes", "text")):
            fmt = fmts.calc(style)
            ws.merge_range(row, label_cols[0], row, label_cols[1], "", fmt)
            ws.write_formula(row, label_cols[0],
                             reg.render(f'={vis(field)}'), fmt)
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL,
            "Sitter notes (write on the printed sheet):",
            fmts.chrome("label_bold"))
    row += 1
    for _ in range(4):
        s.merge(row, 0, LAST_COL, "", fmts.chrome("form_line"))
        s.row_height(row, 18)
        row += 1

    ws.print_area(2, 0, row, LAST_COL)
    ws.fit_to_pages(1, 0)
    ws.hide_gridlines(2)
    s.protect()
    return s


# ---------------------------------------------------------------------------
# 3. Print Visit report
# ---------------------------------------------------------------------------

def build_visit_report(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, VISIT_SHEET,
                    tab_color=Palette.primary_soft, landscape=False)
    ws = s.ws
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 14)
    s.set_col(2, 2, 14)
    s.set_col(3, 3, 12)
    s.set_col(4, 4, 14)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 12)
    s.set_col(7, 7, 14)

    _pick(s, fmts, "Visit to print", "VisitIDs", "PrintVisitID",
          "Pick a visit ID from the Visit Log. This row does not print.",
          default=_first_visit(cfg))

    def vis(field: str, fallback: str = '""') -> str:
        return _index("Visit Log", field, "PrintVisitID", "id", fallback)

    row = 2
    s.merge(row, 0, LAST_COL, "", fmts.chrome("print_title"))
    ws.write_formula(row, 0, '="Visit Report"', fmts.chrome("print_title"))
    s.row_height(row, 28)
    row += 1
    s.merge(row, 0, LAST_COL, "", fmts.chrome("print_sub"))
    ws.write_formula(
        row, 0,
        '="Prepared by "&Biz_Name&" - "&Biz_Phone&" - "&Biz_Email',
        fmts.chrome("print_sub"))
    row += 2

    pairs: List[Tuple[str, str, str, str, str, str]] = [
        ("Visit ID", f'={vis("id")}', "text",
         "Date", f'={vis("date")}', "date"),
        ("Client", f'={vis("client")}', "text",
         "Pet", f'={vis("pet")}', "text"),
        ("Service", f'={vis("service")}', "text",
         "Carer", f'={vis("staff")}', "text"),
        ("Scheduled", f'={vis("sched_start")}', "time",
         "Status", f'={vis("status")}', "text"),
        ("Arrived", f'={vis("arrive")}', "time",
         "Departed", f'={vis("depart")}', "time"),
        ("Minutes on site", f'={vis("duration")}', "int",
         "Punctuality", f'={vis("punctual")}', "text"),
    ]
    s.merge(row, 0, LAST_COL, "VISIT DETAILS", fmts.chrome("subsection"))
    row += 1
    for l1, f1, s1, l2, f2, s2 in pairs:
        for base, label, formula, style in ((0, l1, f1, s1),
                                            (4, l2, f2, s2)):
            s.merge(row, base, base + 1, label, fmts.chrome("label_bold"))
            fmt = fmts.calc(style)
            end = base + 3 if base == 0 else LAST_COL
            ws.merge_range(row, base + 2, row, end, "", fmt)
            ws.write_formula(row, base + 2, reg.render(formula), fmt)
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL, "CARE GIVEN", fmts.chrome("subsection"))
    row += 1
    care = [("Fed", "fed", "text"), ("Fresh water", "water", "text"),
            ("Medication given", "meds", "text"),
            ("Potty / litter", "potty", "text"),
            ("Walk minutes", "walk_mins", "int"),
            ("Distance walked", "walk_distance", "number1"),
            ("Exercise & play", "play", "text"),
            ("Grooming", "grooming", "text")]
    half = (len(care) + 1) // 2
    for i, (label, field, style) in enumerate(care):
        base = 0 if i < half else 4
        r = row + (i % half)
        s.merge(r, base, base + 1, label, fmts.chrome("label_bold"))
        fmt = fmts.calc(style)
        end = base + 3 if base == 0 else LAST_COL
        ws.merge_range(r, base + 2, r, end, "", fmt)
        ws.write_formula(r, base + 2, reg.render(f'={vis(field)}'), fmt)
    row += half + 1

    s.merge(row, 0, LAST_COL, "HOW YOUR PET WAS", fmts.chrome("subsection"))
    row += 1
    for label, field in (("Mood & behaviour", "mood"),
                         ("Appetite", "appetite"),
                         ("Health observations", "health"),
                         ("Visit notes", "notes"),
                         ("Photos / report link", "photos")):
        s.merge(row, 0, 1, label, fmts.chrome("label_bold"))
        fmt = fmts.calc("wrap")
        ws.merge_range(row, 2, row, LAST_COL, "", fmt)
        ws.write_formula(row, 2, reg.render(f'={vis(field)}'), fmt)
        s.row_height(row, 22)
        row += 1
    row += 1

    s.merge(row, 0, 1, "Incident reported", fmts.chrome("label_bold"))
    fmt = fmts.calc("text")
    ws.merge_range(row, 2, row, LAST_COL, "", fmt)
    ws.write_formula(
        row, 2,
        reg.render(f'=IF({vis("incident")}="Yes","Yes - reference "'
                   f'&{vis("incident_ref")},"No")'), fmt)
    row += 2

    s.merge(row, 0, LAST_COL, "", fmts.chrome("note_box"))
    ws.write_formula(
        row, 0,
        '="Thank you for choosing "&Biz_Name&". Any questions about this '
        'visit? Call "&Biz_Phone&"."', fmts.chrome("note_box"))
    s.row_height(row, 24)
    row += 2
    s.merge(row, 0, 2, "Carer signature", fmts.chrome("label_bold"))
    s.merge(row, 3, LAST_COL, "", fmts.chrome("form_line"))

    ws.print_area(2, 0, row, LAST_COL)
    ws.fit_to_pages(1, 1)
    ws.hide_gridlines(2)
    s.protect()
    return s


# ---------------------------------------------------------------------------
# 4. Print Forms - blank paperwork
# ---------------------------------------------------------------------------

#: (title, intro, [field labels], (grid header, grid rows))
FORMS: List[Tuple[str, str, List[str], Optional[Tuple[List[str], int]]]] = [
    ("New Client Registration",
     "Please complete and return before your first booking.",
     ["Client name", "Phone", "Email", "Street address", "City / suburb",
      "Post code", "Emergency contact name", "Emergency contact phone",
      "Preferred contact method", "How did you hear about us?",
      "Home access (key / lockbox / code)", "Alarm instructions",
      "Anyone else with access to the property",
      "Special instructions"],
     None),
    ("Pet Profile & Care Instructions",
     "One form per pet. Attach a photo if you can.",
     ["Pet name", "Type / breed", "Date of birth", "Gender", "Weight",
      "Colour & markings", "Microchip number", "Temperament",
      "Food & portion", "Feeding times", "Medication & dosage",
      "Medication times", "Walking / exercise routine",
      "Potty / litter routine", "Allergies", "Medical conditions",
      "Veterinarian & phone", "Emergency clinic & phone",
      "Pet insurance & policy", "Favourite toys, treats & games",
      "Behaviour warnings", "Handling instructions"],
     None),
    ("Service Agreement & Veterinary Consent",
     "This authorises emergency treatment and confirms the agreed terms.",
     ["Client name", "Pet(s) covered", "Services agreed",
      "Start date", "Rates agreed", "Payment terms",
      "Cancellation policy understood (initial)",
      "Maximum veterinary spend authorised",
      "Preferred veterinary practice",
      "Permission to transport pet by vehicle (yes / no)",
      "Permission to use photos on social media (yes / no)",
      "Client signature", "Date", "Business representative", "Date"],
     None),
    ("Key Handover Receipt",
     "Complete both halves - one for you, one for the client.",
     ["Client name", "Address", "Key tag / reference",
      "Number of keys received", "Date received", "Received by",
      "Client signature", "Date returned", "Returned by",
      "Client signature on return"],
     None),
    ("Incident / Accident Report",
     "Complete on the day. Keep the original with your records.",
     ["Incident reference", "Date & time", "Location", "Pet(s) involved",
      "Client", "Staff present", "Incident type", "Severity",
      "What happened", "Immediate action taken",
      "Veterinary treatment required (yes / no)",
      "Vet / clinic attended", "Cost incurred", "Owner informed (time)",
      "Follow up required", "Prevention / lessons learned",
      "Completed by", "Signature & date"],
     None),
    ("Daily Visit Checklist",
     "Leave this with the client or keep for your records.",
     ["Client", "Pet(s)", "Week commencing", "Carer"],
     (["Date", "In", "Out", "Fed", "Water", "Meds", "Walk", "Potty",
       "Notes"], 14)),
    ("Staff Timesheet",
     "One sheet per person per week.",
     ["Team member", "Week commencing", "Pay rate", "Approved by"],
     (["Date", "Client / job", "Start", "Finish", "Break",
       "Hours", "Distance", "Notes"], 14)),
    ("Mileage Log",
     "Keep this in the vehicle and copy it into the Mileage tab weekly.",
     ["Driver", "Vehicle", "Registration", "Month"],
     (["Date", "From", "To", "Purpose", "Start odo", "End odo",
       "Distance", "Business?"], 16)),
    ("Client Feedback Form",
     "We would love to know how we did.",
     ["Client name", "Date of service", "Carer",
      "Rating out of 5 (circle): 1  2  3  4  5",
      "What went well?", "What could we do better?",
      "Would you recommend us? (yes / no)",
      "May we use your comments as a review? (yes / no)",
      "Signature"],
     None),
    ("Meet & Greet Checklist",
     "Run through this at the first visit.",
     ["Date & time", "Client", "Pet(s)", "Carer"],
     (["Checked", "Item", "Notes", "", "", "", "", ""], 14)),
]

MEET_ITEMS = [
    "Introductions - pet comfortable with the carer",
    "Keys / access tested (door, gate, alarm, lockbox)",
    "Feeding routine and food storage shown",
    "Medication, dosage and storage confirmed",
    "Lead, harness, carrier and towels located",
    "Walking route, pick-up bags and rules agreed",
    "Potty / litter routine and cleaning supplies shown",
    "Emergency contacts and vet details confirmed",
    "Vaccination records sighted",
    "House rules (furniture, rooms, visitors) agreed",
    "Waste, bins and recycling day explained",
    "Preferred updates: photos, messages, frequency",
    "Service agreement signed and consent given",
    "Booking schedule, rates and payment confirmed",
]


def build_forms(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, FORMS_SHEET,
                    tab_color=Palette.primary_soft, landscape=False)
    ws = s.ws
    s.set_col(0, 0, 26)
    s.set_col(1, 7, 11)

    s.merge(0, 0, LAST_COL,
            "Blank forms - print the page you need. Each form starts on a "
            "new page. These are for printing, so nothing here links to "
            "the rest of the workbook.",
            fmts.chrome("note_box"))
    s.row_height(0, 26)

    breaks: List[int] = []
    row = 2
    for title, intro, fields, grid in FORMS:
        if row > 2:
            breaks.append(row)
        s.merge(row, 0, LAST_COL, "", fmts.chrome("print_title"))
        ws.write_formula(row, 0, f'=Biz_Name&" - {title}"',
                         fmts.chrome("print_title"))
        s.row_height(row, 26)
        row += 1
        s.merge(row, 0, LAST_COL, intro, fmts.chrome("print_sub"))
        row += 2

        for label in fields:
            s.merge(row, 0, 1, label, fmts.chrome("form_label"))
            s.merge(row, 2, LAST_COL, "", fmts.chrome("form_line"))
            s.row_height(row, 20)
            row += 1

        if grid:
            headers, lines = grid
            row += 1
            for i, head in enumerate(headers[:LAST_COL + 1]):
                s.write(row, i, head,
                        fmts.chrome("header_left" if i == 0
                                    else "header"))
            s.row_height(row, 18)
            row += 1
            seed = MEET_ITEMS if title.startswith("Meet") else []
            for line in range(lines):
                for i in range(len(headers[:LAST_COL + 1])):
                    if seed and i == 1 and line < len(seed):
                        s.merge(row, 1, LAST_COL, seed[line],
                                fmts.chrome("form_box"))
                        break
                    s.write(row, i, None, fmts.chrome("form_box"))
                s.row_height(row, 18)
                row += 1

        row += 1
        s.merge(row, 0, LAST_COL,
                "", fmts.chrome("print_sub"))
        ws.write_formula(
            row, 0,
            '=Biz_Name&"  |  "&Biz_Phone&"  |  "&Biz_Email',
            fmts.chrome("print_sub"))
        row += 3

    try:
        ws.set_h_pagebreaks(breaks)
    except Exception:          # pragma: no cover - older API guard
        pass

    ws.print_area(2, 0, row, LAST_COL)
    ws.fit_to_pages(1, 0)
    ws.hide_gridlines(2)
    s.protect()
    return s
````

### `pet_sitting_spreadsheet/sheets/help_tab.py`

*295 lines · 1 function*

> `Help` - instructions, formula notes, troubleshooting and version log.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Help'), `LAST_COL` (9), `HOW_IT_WORKS` (5 items), `DAILY` (6 items), `TASKS` (14 items), `FORMULA_NOTES` (11 items), `TROUBLESHOOTING` (9 items), `PRIVACY` (5 items), `VERSIONS` (1 items).

````python
"""`Help` - instructions, formula notes, troubleshooting and version log."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Help"
LAST_COL = 9

HOW_IT_WORKS = [
    "Cream cells are for you. Grey cells are locked formulas - they fill "
    "themselves in as soon as the row has enough information.",
    "Every sheet is protected so a stray keystroke cannot wipe out a "
    "formula. The password is below if you ever need to change one.",
    "Records join up by name, not by hidden ID: pick the client, pet, "
    "service or staff member from the drop down and the rest follows.",
    "The drop downs all live on the Lists tab. Edit a list there and every "
    "sheet picks up the change.",
    "Numbers on the Dashboard, Reports and Budget tabs come from a hidden "
    "sheet called Calc. You never need to open it - it does the heavy "
    "lifting so the visible formulas stay short.",
]

DAILY = [
    "Morning: open the Reminders tab and work through anything with a "
    "count above zero.",
    "Before each job: check the pet's care sheet (Print Pet Care) if "
    "someone else is covering.",
    "After each visit: add a row to the Visit Log and send the client "
    "their update.",
    "Weekly: raise invoices, record payments, log expenses and trips.",
    "Monthly: set the Dashboard period, check the Budget tab and print "
    "the Reports pack.",
    "Yearly: print the full-year report, save a dated copy of the file, "
    "then clear last year's rows from a fresh copy.",
]

TASKS = [
    ("Add a client",
     "Clients tab > first empty row > type the name. The ID, counters and "
     "revenue columns fill themselves in."),
    ("Add a pet",
     "Pets tab > type the pet name, then choose the owner from the drop "
     "down. The pet reference column builds the 'Pet (Owner)' label that "
     "the other sheets use."),
    ("Take a booking",
     "Bookings tab > date, client, pet, service. The price comes from the "
     "Services tab and adds extra pet, weekend, holiday, travel and "
     "last-minute fees automatically."),
    ("Price something differently",
     "Type your figure in the 'Price override' column - it beats the "
     "calculated subtotal. Add a reason so you remember why."),
    ("Log a visit",
     "Visit Log tab > choose the booking ID; the client, pet and service "
     "copy across. Fill in arrival and departure times and the care you "
     "gave."),
    ("Invoice a booking",
     "Invoices tab > invoice date, client, booking ref. The total, due "
     "date, tax and ageing all calculate. Print it from the Print Invoice "
     "tab."),
    ("Invoice several visits at once",
     "Leave the booking ref blank, write your own description and use "
     "Qty x Rate."),
    ("Record a payment",
     "Payments tab > date, client, invoice number, amount. The invoice "
     "balance and status update instantly."),
    ("Take a deposit",
     "Record it on the Payments tab with Deposit = Yes. It counts towards "
     "the invoice and shows in the deposit column."),
    ("Claim mileage",
     "Mileage tab > start and end odometer readings. The claim uses the "
     "mileage rate on the Settings tab."),
    ("Pay your team",
     "Staff tab > set the pay type and rate. Hours and visits come from "
     "the Visit Log, and the wages column shows what is owed."),
    ("Check profit",
     "Dashboard for the headline numbers, Reports for a full profit and "
     "loss over any date range."),
    ("Print a care sheet",
     "Print Pet Care tab > pick the pet reference at the top > print."),
    ("Add more rows",
     "Select the last row of the grid, right click > Insert copied cells, "
     "or copy the last row down. The formulas copy with it."),
]

FORMULA_NOTES = [
    ("Record IDs", "CL-0001, PET-0001, BK-0001 ...",
     "Built from the row number, so they stay unique. Filter the list "
     "instead of re-sorting it and the IDs keep lining up with your other "
     "tabs."),
    ("Booking price", "rate x quantity (or hours) + extras - discount",
     "Looks the service up on the Services tab, then applies the extra "
     "pet fee, weekend and holiday surcharges, travel fee and last-minute "
     "fee from the Settings tab."),
    ("Travel fee", "MAX(0, distance - free distance) x travel rate",
     "Only charges for the distance beyond the free radius."),
    ("Tax", "subtotal x tax rate (or extracted when prices include tax)",
     "Controlled by 'Tax registered' and 'Prices include tax' on the "
     "Settings tab."),
    ("Invoice balance", "total - payments applied",
     "Payments apply as amount minus refund. Tips and late fees are "
     "counted as income but not against the invoice."),
    ("Ageing bucket", "today - due date",
     "Paid, Not due, Current, 1-30, 31-60, 61-90 and 90+ days."),
    ("Double booking check",
     "SUMPRODUCT over same carer, same date, overlapping times",
     "Shows 'Conflict' when two jobs overlap for one person."),
    ("Client revenue", "SUMIFS of payments received for that client",
     "Lifetime revenue, invoiced and outstanding all follow the client "
     "name, so keep names consistent."),
    ("Dashboard period",
     "PeriodStart / PeriodEnd from the year and month pickers",
     "Choose 'All year' for a full year view. Every tile and the service, "
     "expense and client tables follow the picker."),
    ("Mileage claim", "business distance x mileage rate",
     "Personal trips are excluded from the claim and the dashboard."),
    ("Break-even", "fixed costs / gross margin",
     "Gross margin comes from your service rates and staff cost per "
     "service on the Services tab."),
]

TROUBLESHOOTING = [
    ("A cell shows #N/A or a blank where I expected a number",
     "The name in that row does not match a record. Re-pick the client, "
     "pet or service from the drop down rather than typing it."),
    ("I cannot type in a cell",
     "It is a formula cell. Review > Unprotect Sheet, enter the password, "
     "make the change, then protect the sheet again."),
    ("My new rows have no formulas",
     "You typed below the prepared grid. Copy the last prepared row down "
     "as far as you need."),
    ("The totals look wrong",
     "Check the Dashboard year and month pickers, and the From / To dates "
     "on the Reports tab - every figure obeys them."),
    ("Dates look like numbers",
     "Format the cells as a date (Home > Number format > Short date). "
     "This can happen after importing or exporting."),
    ("Nothing recalculates",
     "Press F9, or set Formulas > Calculation Options > Automatic."),
    ("A drop down is missing an option",
     "Add it to the matching column on the Lists tab - the drop downs "
     "grow automatically."),
    ("I want a different currency",
     "Change the currency symbol on the Settings tab, then select the "
     "money columns and apply your own number format if you want the "
     "symbol to change on screen too."),
    ("The file opens in Google Sheets with warnings",
     "Upload to Google Drive and open with Google Sheets. Sheet "
     "protection becomes a warning rather than a password, and "
     "sparklines are redrawn by Sheets - all formulas work."),
]

PRIVACY = [
    "This workbook holds personal information about your clients: names, "
    "addresses, phone numbers and door codes. Treat it like a key safe.",
    "Keep the file on an encrypted drive or in a password protected cloud "
    "folder, and use a file open password (File > Info > Protect "
    "Workbook) if the device is shared.",
    "Only record what you need, and delete alarm codes and key details "
    "once a client leaves.",
    "Back up weekly - keep one copy off the machine you work on.",
    "If a client asks for their data or asks to be deleted, you can "
    "filter the register by their name and export or remove those rows.",
]

VERSIONS = [
    ("1.0", "First release: 27 tabs, 13 registers, dashboard, reports, "
            "budget, reminders and printable forms."),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.grey_text,
                    landscape=False)

    s.banner("Help & Instructions",
             "Everything you need to run the workbook - and what to do "
             "when something looks wrong.", LAST_COL)

    s.set_col(0, 0, 3)
    s.set_col(1, 1, 30)
    s.set_col(2, 8, 12)
    s.set_col(9, 9, 14)

    row = 4
    row = s.section(row, "How this workbook thinks", LAST_COL)
    row = s.bullets(row, HOW_IT_WORKS, LAST_COL, first_col=1, height=16)
    row += 1

    row = s.section(row, "A simple routine", LAST_COL)
    row = s.bullets(row, DAILY, LAST_COL, first_col=1, height=16)
    row += 1

    row = s.section(row, "How do I ...", LAST_COL)
    s.write(row, 1, "Task", fmts.chrome("header_left"))
    s.merge(row, 2, LAST_COL, "What to do", fmts.chrome("header"))
    row += 1
    for task, answer in TASKS:
        s.write(row, 1, task, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, answer, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    row = s.section(row, "What the formulas do", LAST_COL)
    s.write(row, 1, "Calculation", fmts.chrome("header_left"))
    s.merge(row, 2, 4, "In plain English", fmts.chrome("header"))
    s.merge(row, 5, LAST_COL, "Notes", fmts.chrome("header"))
    row += 1
    for label, plain, note in FORMULA_NOTES:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, 4, plain, fmts.chrome("label"))
        s.merge(row, 5, LAST_COL, note, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    row = s.section(row, "Troubleshooting", LAST_COL)
    s.write(row, 1, "Problem", fmts.chrome("header_left"))
    s.merge(row, 2, LAST_COL, "Fix", fmts.chrome("header"))
    row += 1
    for problem, fix in TROUBLESHOOTING:
        s.write(row, 1, problem, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, fix, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    row = s.section(row, "Protection & passwords", LAST_COL)
    facts = [
        ("Sheet password", cfg.password),
        ("To edit a formula",
         "Review > Unprotect Sheet > enter the password > make your change "
         "> Review > Protect Sheet (tick the same boxes) to lock it again."),
        ("What is locked",
         "Only the grey formula cells, plus the headings and layout. "
         "Every cream cell stays editable while the sheet is protected."),
        ("Rows and columns",
         "You can insert and delete rows and sort with the filter "
         "buttons. Inserting columns is blocked because the formulas "
         "point at specific columns."),
    ]
    for label, text in facts:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    row = s.section(row, "Looking after client data", LAST_COL)
    row = s.bullets(row, PRIVACY, LAST_COL, first_col=1, height=16)
    row += 1

    row = s.section(row, "Compatibility", LAST_COL)
    compat = [
        ("Microsoft Excel", "2010, 2013, 2016, 2019, 2021 and Microsoft "
                            "365 - Windows and Mac. Everything works."),
        ("Google Sheets", "Upload and open with Google Sheets. Formulas, "
                          "drop downs and conditional formatting carry "
                          "over; sheet protection becomes a Google Sheets "
                          "warning."),
        ("Apple Numbers / LibreOffice",
         "Formulas and layout open fine. Some conditional formats and "
         "sparklines may look different."),
        ("Mobile", "Read and light editing work well in the Excel and "
                   "Google Sheets apps. Do the data entry on a computer "
                   "where you can see the whole grid."),
    ]
    for label, text in compat:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    row = s.section(row, "Version history", LAST_COL)
    s.write(row, 1, "Version", fmts.chrome("header_left"))
    s.merge(row, 2, LAST_COL, "What changed", fmts.chrome("header"))
    row += 1
    for version, text in VERSIONS:
        s.write(row, 1, version, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        row += 1
    row += 1

    s.note(row,
           f"{cfg.product_name} - (c) {cfg.author}. Licensed for personal "
           "use and for a single business. Please do not resell or "
           "redistribute the file.",
           LAST_COL, height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

---

## 5. Quality checks (optional)

Not needed to build the workbook - these are the checks that keep it honest. The first two need `openpyxl`, the third needs `formulas`.

### `tools/validate_workbook.py`

*287 lines · 3 functions*

> Static checks on a generated workbook.
>
> Run it with any Python that has openpyxl installed::
>
>     python tools/validate_workbook.py output/test.xlsx
>
> It never opens Excel - it parses the XML and reports anything that would
> break for a buyer: unresolved template tokens, formulas pointing at the
> wrong sheet, defined names that do not exist, missing sheet protection,
> formula cells that are not locked, and input cells that are.

| Defines | What it does |
|---------|--------------|
| `_colnum(ref)` |  |
| `check_raw_xml(path)` | Overlapping merges and stray hyperlinks - Excel calls these corrupt |
| `check(path)` |  |

**Module constants:** `EXPECTED_SHEETS` (27 items), `FUNCTION_RE` (re.compile(...)), `NAME_RE` (re.compile(...)), `KNOWN_FUNCTIONS` ({'IF', 'IFERROR', 'SUM', 'SUMIF', 'SUMIFS', 'COUNT', 'COU...), `MODERN_ONLY` ({'XLOOKUP', 'IFS', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCA...).

````python
"""Static checks on a generated workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_workbook.py output/test.xlsx

It never opens Excel - it parses the XML and reports anything that would
break for a buyer: unresolved template tokens, formulas pointing at the
wrong sheet, defined names that do not exist, missing sheet protection,
formula cells that are not locked, and input cells that are.
"""

from __future__ import annotations

import re
import sys
import zipfile
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:                                   # pragma: no cover
    print("openpyxl is required: pip install openpyxl")
    raise SystemExit(2)

EXPECTED_SHEETS = [
    "Start Here", "Dashboard", "Settings", "Lists", "Services", "Clients",
    "Pets", "Bookings", "Visit Log", "Health Log", "Incidents", "Messages",
    "Invoices", "Payments", "Expenses", "Mileage", "Staff", "Inventory",
    "Reminders", "Reports", "Budget", "Print Invoice", "Print Pet Care",
    "Print Visit", "Print Forms", "Help", "Calc",
]

FUNCTION_RE = re.compile(r"\b([A-Z][A-Z0-9\.]{1,})\s*\(")
NAME_RE = re.compile(r"(?<![A-Za-z0-9_.!$\"])([A-Za-z_][A-Za-z0-9_.]{2,})"
                     r"(?![A-Za-z0-9_.(!])")

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
    if hidden != ["Calc"]:
        problems.append(f"hidden sheets should be ['Calc'], got {hidden}")

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
                    if cell.alignment and cell.protection and \
                            not cell.protection.locked:
                        locked_formula_problem += 1
                        if locked_formula_problem < 25:
                            problems.append(
                                f"{ws.title}!{cell.coordinate}: formula cell "
                                f"is NOT locked")
                    stripped = re.sub(r'"[^"]*"', '""', value)
                    stripped = re.sub(r"'[^']*'!", "SHEET!", stripped)
                    if value.count('"') % 2:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: odd number of "
                            f"quotes -> {value[:80]}")
                    if stripped.count("(") != stripped.count(")"):
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unbalanced "
                            f"brackets -> {value[:80]}")
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
    target = sys.argv[1] if len(sys.argv) > 1 else "output/test.xlsx"
    raise SystemExit(check(target))
````

### `tools/check_sample_data.py`

*171 lines · 3 functions*

> Cross-check the example records in a generated workbook.
>
> The demo data has to join up, otherwise the dashboard opens full of zeros.
> This reads the literal values out of the finished file and makes sure every
> client, pet, service, staff member, booking and invoice reference points at
> a record that really exists.
>
>     python tools/check_sample_data.py output/test.xlsx

| Defines | What it does |
|---------|--------------|
| `column_values(ws, header, header_row)` | Every non-empty literal value under the given header |
| `column_rows(ws, header, header_row)` | `{row number: value}` for every non-empty literal under a header |
| `main(path)` |  |

**Module constants:** `HEADER_ROW` (4), `INCIDENT_HEADER_ROW` (17), `TASK_HEADER_ROW` (38).

````python
"""Cross-check the example records in a generated workbook.

The demo data has to join up, otherwise the dashboard opens full of zeros.
This reads the literal values out of the finished file and makes sure every
client, pet, service, staff member, booking and invoice reference points at
a record that really exists.

    python tools/check_sample_data.py output/test.xlsx
"""

from __future__ import annotations

import sys

import openpyxl

HEADER_ROW = 4
INCIDENT_HEADER_ROW = 17
TASK_HEADER_ROW = 38


def column_values(ws, header: str, header_row: int = HEADER_ROW) -> list:
    """Every non-empty literal value under the given header."""
    headers = {}
    for cell in ws[header_row]:
        if isinstance(cell.value, str):
            headers[cell.value.strip()] = cell.column
    if header not in headers:
        raise KeyError(f"{ws.title}: no column called {header!r} "
                       f"(have {sorted(headers)[:8]} ...)")
    col = headers[header]
    out = []
    for row in range(header_row + 1, ws.max_row + 1):
        value = ws.cell(row=row, column=col).value
        if value is None:
            continue
        if isinstance(value, str) and (value.startswith("=") or not value):
            continue
        out.append(value)
    return out


def column_rows(ws, header: str, header_row: int = HEADER_ROW) -> dict:
    """`{row number: value}` for every non-empty literal under a header.

    Same as `column_values` but it keeps the row, which matters for
    columns that are mostly blank.
    """
    headers = {}
    for cell in ws[header_row]:
        if isinstance(cell.value, str):
            headers[cell.value.strip()] = cell.column
    if header not in headers:
        raise KeyError(f"{ws.title}: no column called {header!r}")
    col = headers[header]
    out = {}
    for row in range(header_row + 1, ws.max_row + 1):
        value = ws.cell(row=row, column=col).value
        if value is None or (isinstance(value, str)
                             and (not value or value.startswith("="))):
            continue
        out[row] = value
    return out


def main(path: str) -> int:
    wb = openpyxl.load_workbook(path)
    problems = []

    clients = set(column_values(wb["Clients"], "Client name"))
    services = set(column_values(wb["Services"], "Service name"))
    staff = set(column_values(wb["Staff"], "Name"))
    pet_names = column_values(wb["Pets"], "Pet name")
    pet_owners = column_values(wb["Pets"], "Owner (client)")
    pets = {f"{n} ({o})" for n, o in zip(pet_names, pet_owners)}

    # IDs are formulas, so rebuild them the same way the sheet does.
    booking_rows = len(column_values(wb["Bookings"], "Client"))
    bookings = {f"BK-{i:04d}" for i in range(1, booking_rows + 1)}
    invoice_rows = len(column_values(wb["Invoices"], "Client"))
    invoices = {f"INV-{1001 + i}" for i in range(invoice_rows)}

    def check(sheet, header, allowed, label, header_row=HEADER_ROW,
              optional=True):
        values = column_values(wb[sheet], header, header_row)
        if not optional and not values:
            problems.append(f"{sheet}.{header}: no example data")
        for value in values:
            if value not in allowed:
                problems.append(f"{sheet}.{header}: {value!r} is not a "
                                f"known {label}")

    print(f"clients={len(clients)} pets={len(pets)} services={len(services)} "
          f"staff={len(staff)} bookings={len(bookings)} "
          f"invoices={len(invoices)}")

    check("Pets", "Owner (client)", clients, "client", optional=False)
    check("Bookings", "Client", clients, "client", optional=False)
    check("Bookings", "Pet", pets, "pet reference", optional=False)
    check("Bookings", "Service", services, "service", optional=False)
    check("Bookings", "Assigned to", staff, "staff member")
    check("Visit Log", "Booking ID", bookings, "booking", optional=False)
    check("Visit Log", "Carer", staff, "staff member")
    check("Health Log", "Pet", pets, "pet reference")
    check("Health Log", "Given by", staff, "staff member")
    check("Messages", "Client", clients, "client")
    check("Invoices", "Client", clients, "client", optional=False)
    check("Invoices", "Booking ref", bookings, "booking")
    check("Payments", "Client", clients, "client", optional=False)
    check("Payments", "Invoice no.", invoices, "invoice", optional=False)
    check("Mileage", "Client / booking", clients | bookings,
          "client or booking")
    check("Mileage", "Driver", staff, "staff member")
    check("Expenses", "Staff member", staff, "staff member")
    check("Expenses", "Related client / booking", clients | bookings,
          "client or booking")
    check("Incidents", "Client", clients, "client", INCIDENT_HEADER_ROW)
    check("Incidents", "Pet", pets, "pet reference", INCIDENT_HEADER_ROW)
    check("Reminders", "Owner", staff, "staff member", TASK_HEADER_ROW)

    # payments must be posted against their own client's invoice
    inv_client = {f"INV-{1001 + i}": value for i, value in
                  enumerate(column_values(wb["Invoices"], "Client"))}
    pay_clients = column_values(wb["Payments"], "Client")
    pay_invoices = column_values(wb["Payments"], "Invoice no.")
    for client, invoice in zip(pay_clients, pay_invoices):
        if inv_client.get(invoice) != client:
            problems.append(f"Payments: {invoice} belongs to "
                            f"{inv_client.get(invoice)!r} but the payment "
                            f"is logged against {client!r}")

    # visits must match the client and pet on their booking
    bk_client = {f"BK-{i + 1:04d}": value for i, value in
                 enumerate(column_values(wb["Bookings"], "Client"))}
    bk_pet = {f"BK-{i + 1:04d}": value for i, value in
              enumerate(column_values(wb["Bookings"], "Pet"))}
    vs_booking = column_values(wb["Visit Log"], "Booking ID")
    vs_client = column_values(wb["Visit Log"], "Client")
    vs_pet = column_values(wb["Visit Log"], "Pet")
    for booking, client, pet in zip(vs_booking, vs_client, vs_pet):
        if bk_client.get(booking) != client:
            problems.append(f"Visit Log: {booking} is for "
                            f"{bk_client.get(booking)!r}, visit says "
                            f"{client!r}")
        if bk_pet.get(booking) != pet:
            problems.append(f"Visit Log: {booking} is for "
                            f"{bk_pet.get(booking)!r}, visit says {pet!r}")

    # an invoice raised from a booking must be addressed to that booking's
    # client, otherwise the invoice total is pulled from somebody else's job
    inv_booking = column_rows(wb["Invoices"], "Booking ref")
    inv_clients = column_rows(wb["Invoices"], "Client")
    for row, booking in inv_booking.items():
        client = inv_clients.get(row)
        if bk_client.get(booking) != client:
            problems.append(f"Invoices: INV-{1001 + row - 5} is addressed "
                            f"to {client!r} but {booking} belongs to "
                            f"{bk_client.get(booking)!r}")

    if problems:
        print(f"! {len(problems)} problem(s)")
        for problem in problems[:40]:
            print(f"  - {problem}")
        return 1
    print("OK - every example record joins up")
    return 0


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "output/test.xlsx"
    raise SystemExit(main(target))
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

### `tools/export_source_md.py`

*922 lines · 10 functions*

> Write every source file of a generator into one markdown document.
>
>     python tools/export_source_md.py                   # -> SOURCE_CODE.md
>     python tools/export_source_md.py --product social  # -> SOURCE_CODE_SOCIAL.md
>     python tools/export_source_md.py -o /tmp/src.md
>
> Handy when the code has to travel somewhere that is not a git checkout:
> one file, every module in reading order, each in a fenced block with its
> real path above it, plus a short description taken from the module's own
> docstring so the two can never disagree.

| Defines | What it does |
|---------|--------------|
| `outline(path)` | `[(kind, name, note)]` for the classes, functions and constants |
| `describe_value(node)` | A short, human description of a module level constant |
| `first_sentence(text)` |  |
| `blurb(path)` | First sentence of the module docstring, as a one line description |
| `full_docstring(path)` |  |
| `anchor(heading)` |  |
| `tab_table(product)` | A row per tab: order, name, kind, module, columns, rows |
| `data_dictionary(product)` | One table per register: every column, its type and where it comes from |
| `file_table(sections)` |  |
| `main(argv)` |  |

**Module constants:** `ROOT` (os.path.dirname(...)), `SHARED_TOOLS` (2 items), `PET_SECTIONS` (5 items), `PET` (9 entries), `SOCIAL_SECTIONS` (7 items), `SOCIAL` (9 entries), `MEAL_SECTIONS` (6 items), `MEAL` (11 entries), `PRODUCTS` (3 entries), `ARCHITECTURE` ('Everything here is plain Python writing a `.xlsx` file t...), `DEFAULT_CLOSING` ('**Hide the `Calc` engine**, activate the first tab and c...), `CALC_ENGINE` ('### The hidden `Calc` tab\n\nDashboards do not aggregate...), `RANK_ENGINE` ('### Ranking columns instead of a hidden engine\n\nThere ...).

````python
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

PRODUCTS = {"pet": PET, "social": SOCIAL, "meal": MEAL}


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
````

### `tools/make_download_zip.py`

*232 lines · 2 functions*

> Pack a generator into a single zip people can download and run.
>
>     python tools/make_download_zip.py
>     python tools/make_download_zip.py --product social
>     python tools/make_download_zip.py --product meal
>
> Writes `download/<product>-source.zip` containing the generator, the
> quality tools, the documentation and a `requirements.txt`, all under one
> folder so it unpacks tidily.
>
> The vendored copy of XlsxWriter in this repository is deliberately left
> out: the zip is meant to be unpacked anywhere and run against a normal
> `pip install XlsxWriter`.  Timestamps are fixed so rebuilding the zip
> without changing any source produces exactly the same bytes.

| Defines | What it does |
|---------|--------------|
| `gather(product)` | `[(archive name, text)]` for everything that goes in the zip |
| `main(argv)` |  |

**Module constants:** `ROOT` (os.path.dirname(...)), `STAMP` (6 items), `SHARED_TOOLS` (3 items), `REQUIREMENTS` ('# The generator needs one package.\nXlsxWriter>=3.0\n\n#...), `PET_HOW_TO_RUN` ('Pet Sitting & Dog Walking Business Spreadsheet\n========...), `SOCIAL_HOW_TO_RUN` ('Social Media Content Calendar & Analytics Planner\n=====...), `MEAL_HOW_TO_RUN` ('Automated Meal Planner & Grocery List Spreadsheet\n=====...), `PRODUCTS` (3 entries).

````python
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
````

---

## Appendix A. Data dictionary

Every column of every register tab, straight from the specs that build them: the Excel column letter, the header a buyer sees, the key formulas use, the number format and whether the cell is typed in, chosen from a drop down or calculated. Pages (dashboards, reports, printables) are not listed - they have no fixed grid.

#### Services

Add every service you sell. Bookings pull the billing unit, rate, additional pet fee and surcharges straight from this table.

20 columns, 50 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Service name | `name` | text | you type it |
| C | Category | `category` | text | drop down `ServiceCategoryList` |
| D | What is included | `description` | wrap | you type it |
| E | Billing unit | `unit` | text | drop down `BillingUnitList` |
| F | Standard length (mins) | `duration` | int | you type it (checked) |
| G | Base rate | `rate` | money | you type it (checked) |
| H | Additional pet fee | `extra_pet` | money | you type it (checked) |
| I | Holiday surcharge | `hol_pct` | percent | you type it (checked) |
| J | Weekend surcharge | `wkd_pct` | percent | you type it (checked) |
| K | Travel fee applies? | `travel` | flag | you type it (checked) |
| L | Staff cost per unit | `staff_cost` | money | you type it (checked) |
| M | Gross profit per unit | `profit` | money | formula |
| N | Margin % | `margin` | percent | formula |
| O | Times booked | `times_booked` | int | formula |
| P | Revenue to date | `revenue` | money | formula |
| Q | Average value | `avg_value` | money | formula |
| R | Estimated profit | `est_profit` | money | formula |
| S | Active? | `active` | flag | you type it (checked) |
| T | Notes | `notes` | wrap | you type it |

#### Clients

One row per client. Grey columns calculate pets, visits, lifetime revenue, outstanding balances and follow-up actions for you.

38 columns, 250 rows of capacity, 23 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Client name | `name` | text | you type it |
| C | Phone | `phone` | text | you type it |
| D | Email | `email` | text | you type it |
| E | Street address | `address` | text | you type it |
| F | City / suburb | `city` | text | you type it |
| G | Post code | `postcode` | text | you type it |
| H | Emergency contact | `emergency_name` | text | you type it |
| I | Emergency phone | `emergency_phone` | text | you type it |
| J | Preferred contact | `contact_method` | text | drop down `ContactMethodList` |
| K | Referral source | `referral` | text | drop down `ReferralSourceList` |
| L | Registered on | `reg_date` | date | you type it (checked) |
| M | Status | `status` | status | drop down `ClientStatusList` |
| N | Key / access | `key_status` | text | drop down `KeyStatusList` |
| O | Key tag / ref | `key_tag` | text | you type it |
| P | Key returned on | `key_returned` | date | you type it (checked) |
| Q | Home access instructions | `access_notes` | wrap | you type it |
| R | Alarm / gate code | `access_code` | text | you type it |
| S | Contract status | `contract` | status | drop down `ContractStatusList` |
| T | Agreement signed | `agreement_date` | date | you type it (checked) |
| U | Emergency vet consent | `vet_consent` | flag | you type it (checked) |
| V | Vet spend authority | `vet_limit` | money | you type it (checked) |
| W | Preferences & special instructions | `preferences` | wrap | you type it |
| X | Notes & communication history | `notes` | wrap | you type it |
| Y | Pets | `pets` | int | formula |
| Z | Bookings | `bookings` | int | formula |
| AA | Visits done | `visits` | int | formula |
| AB | Last service | `last_service` | date | formula |
| AC | Days since | `days_since` | int | formula |
| AD | Next booking | `next_booking` | date | formula |
| AE | Lifetime revenue | `revenue` | money | formula |
| AF | Invoiced to date | `invoiced` | money | formula |
| AG | Outstanding | `outstanding` | money | formula |
| AH | Average booking | `avg_booking` | money | formula |
| AI | Loyalty status | `loyalty` | status | formula |
| AJ | Duplicate check | `dup_check` | status | formula |
| AK | Missing information | `data_check` | status | formula |
| AL | Action needed | `alert` | status | formula |

#### Pets

Everything a sitter needs before they walk through the door: routine, medical, vet, vaccination status and safety warnings.

46 columns, 250 rows of capacity, 37 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Pet name | `name` | text | you type it |
| C | Owner (client) | `owner` | text | drop down `ClientNames` |
| D | Pet reference | `label` | text | formula |
| E | Client ID | `client_id` | text | formula |
| F | Pet type | `type` | text | drop down `PetTypeList` |
| G | Breed | `breed` | text | you type it |
| H | Date of birth | `dob` | date | you type it (checked) |
| I | Age (yrs) | `age` | number1 | formula |
| J | Gender | `gender` | text | drop down `PetGenderList` |
| K | Weight (kg) | `weight` | number1 | you type it (checked) |
| L | Colour & markings | `colour` | text | you type it |
| M | Photo link | `photo` | text | you type it |
| N | Microchip number | `microchip` | text | you type it |
| O | Temperament | `temperament` | text | drop down `TemperamentList` |
| P | Behaviour issues | `behaviour` | wrap | you type it |
| Q | Allergies | `allergies` | wrap | you type it |
| R | Medical conditions / special needs | `medical` | wrap | you type it |
| S | Veterinarian | `vet_name` | text | you type it |
| T | Vet phone | `vet_phone` | text | you type it |
| U | Emergency clinic | `emerg_clinic` | text | you type it |
| V | Emergency clinic phone | `emerg_phone` | text | you type it |
| W | Vaccinated on | `vacc_date` | date | you type it (checked) |
| X | Vaccination expires | `vacc_expiry` | date | you type it (checked) |
| Y | Vaccination status | `vacc_status` | status | formula |
| Z | Pet insurance | `insurer` | text | you type it |
| AA | Policy number | `policy` | text | you type it |
| AB | Insurance renewal | `insurance_expiry` | date | you type it (checked) |
| AC | Feeding instructions | `feeding` | wrap | you type it |
| AD | Food type | `food` | text | you type it |
| AE | Portion size | `portion` | text | you type it |
| AF | Feeding schedule | `feed_schedule` | text | you type it |
| AG | Medication | `medication` | text | you type it |
| AH | Dosage & instructions | `dosage` | wrap | you type it |
| AI | Medication schedule | `med_schedule` | text | you type it |
| AJ | Walking & exercise preferences | `walking` | wrap | you type it |
| AK | Potty / litter routine | `potty` | wrap | you type it |
| AL | Grooming requirements | `grooming` | wrap | you type it |
| AM | Favourite toys, treats & games | `favourites` | wrap | you type it |
| AN | Risk flag | `risk` | status | drop down `RiskFlagList` |
| AO | Handling instructions | `handling` | wrap | you type it |
| AP | Safety warnings | `warnings` | wrap | you type it |
| AQ | Visits | `visits` | int | formula |
| AR | Last visit | `last_visit` | date | formula |
| AS | Med records | `meds_logged` | int | formula |
| AT | Alerts | `alert` | status | formula |

#### Bookings

Every booking, priced automatically from the Services tab with holiday, weekend, extra pet, travel and last-minute rules applied.

52 columns, 400 rows of capacity, 23 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Requested on | `request_date` | date | you type it (checked) |
| C | Client | `client` | text | drop down `ClientNames` |
| D | Client ID | `client_id` | text | formula |
| E | Pet | `pet` | text | drop down `PetLabels` |
| F | Additional pets | `extra_pets` | text | you type it |
| G | No. of pets | `pet_count` | int | you type it (checked) |
| H | Service | `service` | text | drop down `ServiceNames` |
| I | Booking date | `date` | date | you type it (checked) |
| J | Day | `day` | flag | formula |
| K | Start | `start` | time | you type it |
| L | End | `end` | time | you type it |
| M | Hours | `duration` | number | formula |
| N | Qty / nights | `qty` | number | you type it (checked) |
| O | Billing unit | `unit` | text | formula |
| P | Rate | `rate` | money | formula |
| Q | Extra pet rate | `pet_rate` | money | formula, hidden |
| R | Holiday % | `hol_pct` | percent | formula, hidden |
| S | Weekend % | `wkd_pct` | percent | formula, hidden |
| T | Travel billable | `travel_flag` | flag | formula, hidden |
| U | Base charge | `base` | money | formula |
| V | Extra pet fee | `extra_pet_fee` | money | formula |
| W | Holiday? | `holiday` | flag | formula |
| X | Weekend? | `weekend` | flag | formula |
| Y | Surcharges | `surcharge` | money | formula |
| Z | Travel distance | `distance` | number1 | you type it (checked) |
| AA | Travel fee | `travel_fee` | money | formula |
| AB | Travel allowance (min) | `travel_time` | int | formula |
| AC | Last-minute fee | `rush_fee` | money | formula |
| AD | Discount type | `discount_type` | text | drop down `DiscountTypeList` |
| AE | Discount % | `discount_pct` | percent | you type it (checked) |
| AF | Discount | `discount` | money | formula |
| AG | Price override | `override` | money | you type it (checked) |
| AH | Override reason | `override_note` | text | you type it |
| AI | Subtotal | `subtotal` | money | formula |
| AJ | Tax | `tax` | money | formula |
| AK | Booking total | `total` | money | formula |
| AL | Assigned to | `staff` | text | drop down `StaffNames` |
| AM | Status | `status` | status | drop down `BookingStatusList` |
| AN | Confirmed to client | `confirm_sent` | flag | you type it (checked) |
| AO | Reminder sent | `reminder_sent` | flag | you type it (checked) |
| AP | Repeats | `recurring` | text | drop down `RecurringList` |
| AQ | Next booking due | `next_due` | date | formula |
| AR | Schedule check | `conflict` | status | formula |
| AS | Visit logged | `visit_logged` | flag | formula |
| AT | Invoice | `invoice_no` | text | formula |
| AU | Payment status | `paid_status` | status | formula |
| AV | Cancellation reason | `cancel_reason` | text | drop down `CancelReasonList` |
| AW | Rescheduled from | `resched_from` | date | you type it (checked) |
| AX | Notes & special requests | `notes` | wrap | you type it |
| AY | Month | `month` | text | formula |
| AZ | Action needed | `alert` | status | formula |

#### Visit Log

Proof of service for every walk, drop-in and overnight: times, care given, mood, health notes and client updates.

36 columns, 400 rows of capacity, 24 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Booking ID | `booking` | text | drop down `BookingIDs` |
| C | Visit date | `date` | date | you type it (checked) |
| D | Client | `client` | text | formula |
| E | Pet | `pet` | text | formula |
| F | Service | `service` | text | formula |
| G | Carer | `staff` | text | drop down `StaffNames` |
| H | Scheduled | `sched_start` | time | formula |
| I | Arrived | `arrive` | time | you type it |
| J | Departed | `depart` | time | you type it |
| K | Minutes on site | `duration` | int | formula |
| L | On time? | `punctual` | status | formula |
| M | Walk start | `walk_start` | time | you type it |
| N | Walk end | `walk_end` | time | you type it |
| O | Walk minutes | `walk_mins` | int | formula |
| P | Distance walked | `walk_distance` | number1 | you type it (checked) |
| Q | Fed | `fed` | flag | drop down `YesNoNAList` |
| R | Fresh water | `water` | flag | drop down `YesNoNAList` |
| S | Medication given | `meds` | flag | drop down `YesNoNAList` |
| T | Potty / litter | `potty` | text | drop down `PottyList` |
| U | Exercise & play | `play` | text | you type it |
| V | Grooming / cleaning | `grooming` | text | you type it |
| W | Mood & behaviour | `mood` | text | drop down `MoodList` |
| X | Appetite | `appetite` | text | drop down `AppetiteList` |
| Y | Health observations | `health` | wrap | you type it |
| Z | Incident? | `incident` | flag | you type it (checked) |
| AA | Incident ref | `incident_ref` | text | you type it |
| AB | Photos / report link | `photos` | text | you type it |
| AC | Client updated | `update_sent` | flag | you type it (checked) |
| AD | Care checklist done | `checklist` | flag | you type it (checked) |
| AE | Visit status | `status` | status | drop down `VisitStatusList` |
| AF | Visit notes | `notes` | wrap | you type it |
| AG | Month | `month` | text | formula |
| AH | Line key | `line_key` | text | formula, hidden |
| AI | Pet key | `pet_key` | text | formula, hidden |
| AJ | Action needed | `alert` | status | formula |

#### Health Log

A dated record of every dose given, vaccination, vet visit and health observation - your proof of care.

22 columns, 200 rows of capacity, 16 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Time | `time` | time | you type it |
| D | Pet | `pet` | text | drop down `PetLabels` |
| E | Client | `client` | text | formula |
| F | Record type | `type` | text | drop down `HealthRecordTypeList` |
| G | Medication / treatment | `item` | text | you type it |
| H | Dosage & instructions | `dosage` | wrap | you type it |
| I | Scheduled | `sched_time` | time | you type it |
| J | Given at | `given_time` | time | you type it |
| K | Given by | `staff` | text | drop down `StaffNames` |
| L | Status | `status` | status | drop down `MedStatusList` |
| M | Supply left | `supply` | number1 | you type it (checked) |
| N | Refill due | `refill_due` | date | you type it (checked) |
| O | Refill alert | `refill_alert` | status | formula |
| P | Vet appointment | `vet_appt` | date | you type it (checked) |
| Q | Appointment alert | `appt_alert` | status | formula |
| R | Cost | `cost` | money | you type it (checked) |
| S | Owner notified | `notified` | flag | you type it (checked) |
| T | Notes | `notes` | wrap | you type it |
| U | Month | `month` | text | formula |
| V | Action needed | `alert` | status | formula |

#### Incidents

Emergency numbers at the top, a dated incident report register below. Print a row straight from this tab if an owner or insurer asks for it.

23 columns, 100 rows of capacity, 20 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Time | `time` | time | you type it |
| D | Client | `client` | text | drop down `ClientNames` |
| E | Pet | `pet` | text | drop down `PetLabels` |
| F | Reported by | `staff` | text | drop down `StaffNames` |
| G | Booking / visit ref | `ref` | text | you type it |
| H | Incident type | `type` | text | drop down `IncidentTypeList` |
| I | Severity | `severity` | status | drop down `SeverityList` |
| J | What happened | `description` | wrap | you type it |
| K | Action taken | `action` | wrap | you type it |
| L | Vet contacted | `vet_contacted` | flag | you type it (checked) |
| M | Clinic attended | `clinic` | text | you type it |
| N | Owner notified | `owner_notified` | flag | you type it (checked) |
| O | Notified when | `notified_at` | text | you type it |
| P | Follow-up needed | `follow_up` | flag | you type it (checked) |
| Q | Follow-up due | `follow_due` | date | you type it (checked) |
| R | Status | `status` | status | drop down `IncidentStatusList` |
| S | Cost incurred | `cost` | money | you type it (checked) |
| T | Insurance claim | `claim` | flag | you type it (checked) |
| U | Days open | `days_open` | text | formula |
| V | Notes & outcome | `notes` | wrap | you type it |
| W | Action needed | `alert` | status | formula |

#### Messages

Every confirmation, reminder, update, review and complaint in one searchable history.

20 columns, 200 rows of capacity, 16 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Client | `client` | text | drop down `ClientNames` |
| D | Message type | `type` | text | drop down `MessageTypeList` |
| E | Channel | `channel` | text | drop down `ChannelList` |
| F | Direction | `direction` | text | drop down `DirectionList` |
| G | Summary | `summary` | wrap | you type it |
| H | Reply received | `response` | flag | you type it (checked) |
| I | Satisfaction (1-5) | `rating` | rating | you type it (checked) |
| J | Feedback / review | `feedback` | wrap | you type it |
| K | Complaint? | `complaint` | flag | you type it (checked) |
| L | Issue raised | `issue` | wrap | you type it |
| M | Resolution status | `resolution` | status | drop down `ResolutionList` |
| N | Resolved on | `resolved_date` | date | you type it (checked) |
| O | Follow-up due | `follow_due` | date | you type it (checked) |
| P | Follow-up done | `follow_done` | flag | you type it (checked) |
| Q | Days open | `days_open` | int | formula |
| R | Notes | `notes` | wrap | you type it |
| S | Month | `month` | text | formula |
| T | Action needed | `alert` | status | formula |

#### Invoices

Automatic numbering, due dates, tax, balances, ageing buckets and overdue alerts - all from the booking you link.

30 columns, 300 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Invoice no. | `invoice_no` | id | formula |
| B | Invoice date | `date` | date | you type it (checked) |
| C | Due date | `due` | date | formula |
| D | Client | `client` | text | drop down `ClientNames` |
| E | Booking ref | `booking` | text | drop down `BookingIDs` |
| F | Booking total | `booking_total` | money | formula |
| G | Description of service | `description` | wrap | you type it |
| H | Qty | `qty` | number | you type it (checked) |
| I | Rate | `rate` | money | you type it (checked) |
| J | Line amount | `line_total` | money | formula |
| K | Additional fees | `fees` | money | you type it (checked) |
| L | Discount | `discount` | money | you type it (checked) |
| M | Subtotal | `subtotal` | money | formula |
| N | Tax | `tax` | money | formula |
| O | Invoice total | `total` | money | formula |
| P | Deposit received | `deposit` | money | formula |
| Q | Amount paid | `paid` | money | formula |
| R | Refunded | `refunded` | money | formula |
| S | Balance due | `balance` | money | formula |
| T | Invoice status | `status` | status | drop down `InvoiceStatusList` |
| U | Payment status | `pay_status` | status | formula |
| V | Days overdue | `days_overdue` | int | formula |
| W | Ageing bucket | `aging` | status | formula |
| X | Suggested late fee | `late_fee` | money | formula |
| Y | Sent on | `sent_date` | date | you type it (checked) |
| Z | Reminder sent | `reminder_sent` | flag | you type it (checked) |
| AA | Number check | `dup_check` | status | formula |
| AB | Invoice notes | `notes` | wrap | you type it |
| AC | Month | `month` | text | formula |
| AD | Action needed | `alert` | status | formula |

#### Payments

Every cash, card, transfer, wallet and deposit payment, with tips, refunds and bank reconciliation.

20 columns, 300 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Payment date | `date` | date | you type it (checked) |
| C | Client | `client` | text | drop down `ClientNames` |
| D | Invoice no. | `invoice_no` | text | drop down `InvoiceNos` |
| E | Amount applied | `amount` | money | you type it (checked) |
| F | Payment method | `method` | text | drop down `PaymentMethodList` |
| G | Reference | `reference` | text | you type it |
| H | Tip / gratuity | `tip` | money | you type it (checked) |
| I | Late fee collected | `late_fee` | money | you type it (checked) |
| J | Refund given | `refund` | money | you type it (checked) |
| K | Applied to invoice | `net` | money | formula |
| L | Total cash in | `received` | money | formula |
| M | Deposit? | `deposit` | flag | you type it (checked) |
| N | Cleared / confirmed | `confirmed` | flag | you type it (checked) |
| O | Reconciled to bank | `reconciled` | flag | you type it (checked) |
| P | Invoice total | `inv_total` | money | formula |
| Q | Invoice balance | `inv_balance` | money | formula |
| R | Notes | `notes` | wrap | you type it |
| S | Month | `month` | text | formula |
| T | Action needed | `alert` | status | formula |

#### Expenses

Fuel, supplies, wages, insurance, software and every other cost - split into tax, net and claimable amounts.

21 columns, 300 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Category | `category` | text | drop down `ExpenseCategoryList` |
| D | Description | `description` | wrap | you type it |
| E | Supplier / vendor | `supplier` | text | you type it |
| F | Amount paid | `amount` | money | you type it (checked) |
| G | Tax included | `tax` | money | formula |
| H | Net cost | `net` | money | formula |
| I | Paid by | `method` | text | drop down `PaymentMethodList` |
| J | Business / personal | `type` | status | drop down `BusinessPersonalList` |
| K | Business use % | `deductible` | percent | you type it (checked) |
| L | Claimable amount | `claim` | money | formula |
| M | Receipt link / ref | `receipt` | text | you type it |
| N | Recurring | `recurring` | text | drop down `RecurringExpenseList` |
| O | Next due | `next_due` | date | formula |
| P | Approval | `approval` | status | drop down `ExpenseApprovalList` |
| Q | Related client / booking | `related` | text | you type it |
| R | Staff member | `staff` | text | drop down `StaffNames` |
| S | Notes | `notes` | wrap | you type it |
| T | Month | `month` | text | formula |
| U | Action needed | `alert` | status | formula |

#### Mileage

Odometer in, odometer out: km travelled, round trips, reimbursement, parking and tolls - ready for your tax return.

24 columns, 300 rows of capacity, 15 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Trip date | `date` | date | you type it (checked) |
| C | Driver | `staff` | text | drop down `StaffNames` |
| D | Client / booking | `client` | text | drop down `ClientNames` |
| E | Purpose | `purpose` | text | drop down `TripPurposeList` |
| F | From | `start_loc` | text | you type it |
| G | To | `end_loc` | text | you type it |
| H | Odometer start | `start_odo` | number1 | you type it (checked) |
| I | Odometer end | `end_odo` | number1 | you type it (checked) |
| J | Manual km | `manual_km` | number1 | you type it (checked) |
| K | Trip km | `trip_km` | number1 | formula |
| L | Round trip? | `round_trip` | flag | you type it (checked) |
| M | Total km | `total_km` | number1 | formula |
| N | Classification | `business` | status | formula |
| O | Rate | `rate` | money | formula |
| P | Mileage claim | `reimbursement` | money | formula |
| Q | Fuel cost | `fuel` | money | you type it (checked) |
| R | Parking | `parking` | money | you type it (checked) |
| S | Tolls | `tolls` | money | you type it (checked) |
| T | Travel time (min) | `travel_mins` | int | you type it (checked) |
| U | Total claimable | `claimable` | money | formula |
| V | Notes | `notes` | wrap | you type it |
| W | Month | `month` | text | formula |
| X | Check | `alert` | status | formula |

#### Staff

Who works for you, when they are available, what they have delivered and what they are owed.

28 columns, 40 rows of capacity, 17 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Name | `name` | text | you type it |
| C | Phone | `phone` | text | you type it |
| D | Email | `email` | text | you type it |
| E | Role | `role` | text | drop down `StaffRoleList` |
| F | Employment type | `employment` | text | drop down `EmploymentTypeList` |
| G | Started on | `start_date` | date | you type it (checked) |
| H | Status | `status` | status | drop down `StaffStatusList` |
| I | Availability | `availability` | text | you type it |
| J | Regular days off | `days_off` | text | you type it |
| K | Pay type | `pay_type` | text | drop down `PayTypeList` |
| L | Pay rate | `pay_rate` | money | you type it (checked) |
| M | Certifications | `certs` | wrap | you type it |
| N | Certificate expires | `cert_expiry` | date | you type it (checked) |
| O | Certificate status | `cert_alert` | status | formula |
| P | Background check | `check_date` | date | you type it (checked) |
| Q | Emergency contact | `emergency` | text | you type it |
| R | Leave & absence | `leave` | wrap | you type it |
| S | Bookings assigned | `bookings` | int | formula |
| T | Visits completed | `visits` | int | formula |
| U | Hours worked | `hours` | number | formula |
| V | Km driven | `distance` | number1 | formula |
| W | Revenue delivered | `revenue` | money | formula |
| X | Pay due | `wages` | money | formula |
| Y | Revenue per hour | `rev_per_hour` | money | formula |
| Z | Completion rate | `completion` | percent | formula |
| AA | Notes | `notes` | wrap | you type it |
| AB | Action needed | `alert` | status | formula |

#### Inventory

Treats, waste bags, cleaning kit, first aid and equipment - with low stock and expiry alerts.

19 columns, 120 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Item | `item` | text | you type it |
| C | Category | `category` | text | drop down `InventoryCategoryList` |
| D | Supplier | `supplier` | text | you type it |
| E | Unit | `unit` | text | you type it |
| F | Unit cost | `unit_cost` | money | you type it (checked) |
| G | In stock | `qty` | number1 | you type it (checked) |
| H | Minimum level | `min_qty` | number1 | you type it (checked) |
| I | Reorder qty | `reorder` | number1 | you type it (checked) |
| J | Stock value | `value` | money | formula |
| K | Stock status | `stock_status` | status | formula |
| L | Last purchased | `last_purchase` | date | you type it (checked) |
| M | Qty last bought | `qty_purchased` | number1 | you type it (checked) |
| N | Last purchase cost | `purchase_cost` | money | formula |
| O | Expiry date | `expiry` | date | you type it (checked) |
| P | Expiry status | `expiry_alert` | status | formula |
| Q | Used this month | `used_month` | number1 | you type it (checked) |
| R | Notes | `notes` | wrap | you type it |
| S | Action needed | `alert` | status | formula |

#### Reminders

Everything the workbook thinks you should do today, plus room for your own reminders.

10 columns, 40 rows of capacity, 7 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Task | `task` | text | you type it |
| C | Category | `category` | text | you type it |
| D | Due date | `due` | date | you type it (checked) |
| E | Owner | `owner` | text | drop down `StaffNames` |
| F | Priority | `priority` | status | drop down `PriorityList` |
| G | Status | `status` | status | drop down `TaskStatusList` |
| H | Days left | `days_left` | int | formula |
| I | Status check | `flag` | status | formula |
| J | Notes | `notes` | wrap | you type it |

#### Calc

Helper tables for the dashboard and reports. This tab is hidden by default - unhide it only if you want to audit the maths.

44 columns, 250 rows of capacity, 0 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Client | `c_name` | text | formula |
| B | Revenue (dashboard period) | `c_rev_p` | money | formula |
| C | Rank key (period) | `c_key_p` | number | formula |
| D | Revenue (report range) | `c_rev_r` | money | formula |
| E | Rank key (range) | `c_key_r` | number | formula |
| F | Outstanding | `c_out` | money | formula |
| G | Outstanding key | `c_out_key` | number | formula |
| H | Bookings (range) | `c_bookings_r` | int | formula |
| I | Service | `s_name` | text | formula |
| J | Service revenue (period) | `s_rev_p` | money | formula |
| K | Service bookings (period) | `s_count_p` | int | formula |
| L | Service revenue (range) | `s_rev_r` | money | formula |
| M | Service bookings (range) | `s_count_r` | int | formula |
| N | Service cost (range) | `s_cost_r` | money | formula |
| O | Service profit (range) | `s_profit_r` | money | formula |
| P | Expense category | `e_cat` | text | formula |
| Q | Spend (period) | `e_total_p` | money | formula |
| R | Spend (range) | `e_total_r` | money | formula |
| S | Month start | `m_start` | date | formula |
| T | Month end | `m_end` | date | formula |
| U | Month | `m_label` | text | formula |
| V | Revenue | `m_rev` | money | formula |
| W | Invoiced | `m_inv` | money | formula |
| X | Expenses | `m_exp` | money | formula |
| Y | Profit | `m_profit` | money | formula |
| Z | Bookings | `m_bookings` | int | formula |
| AA | Completed | `m_completed` | int | formula |
| AB | Cancelled | `m_cancelled` | int | formula |
| AC | Visits | `m_visits` | int | formula |
| AD | New clients | `m_new_clients` | int | formula |
| AE | Business distance | `m_km` | number1 | formula |
| AF | Revenue last year | `m_prev_rev` | money | formula |
| AG | Staff | `st_name` | text | formula |
| AH | Visits (range) | `st_visits_r` | int | formula |
| AI | Hours (range) | `st_hours_r` | number | formula |
| AJ | Revenue (range) | `st_rev_r` | money | formula |
| AK | Distance (range) | `st_km_r` | number1 | formula |
| AL | Cancellation reason | `cr_reason` | text | formula |
| AM | Cancellations (range) | `cr_count` | int | formula |
| AN | Payment method | `pm_method` | text | formula |
| AO | Collected (range) | `pm_total` | money | formula |
| AP | Referral source | `ref_source` | text | formula |
| AQ | Clients from source | `ref_clients` | int | formula |
| AR | Revenue from source | `ref_revenue` | money | formula |

---

*9,227 lines of Python in 42 files.*
