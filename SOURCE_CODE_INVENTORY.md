# Ultimate Inventory & Wholesale Order Manager — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 23 tab `.xlsx` workbook with over 136,000 formulas for anybody who buys stock and sells it: a product database where every separately stocked variant has its own SKU; a stock tracker with nothing typed on it at all; a movement log for damage, loss, samples and transfers; suppliers scored on the lead time they actually deliver; purchase orders with landed costs, part deliveries and a duplicate-receiving guard; a wholesale CRM with terms and credit limits; price breaks you set yourself, checked against cost and against your own floor; wholesale orders that reserve stock the moment they are confirmed; retail sales across ten channels with the platform fees taken out; a reorder planner that works out the date you have to order by; weighted average valuation and COGS; returns that decide whether a unit goes back on the shelf; bin locations; stock counts that raise a variance instead of overwriting the number; profit per product; and printable purchase orders, invoices and picking lists. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

**Stock is derived, never typed.** The Inventory Tracker has no input columns. What you have is opening stock, plus what was received against a purchase order, minus what was sold retail and wholesale, plus what came back, minus what was damaged, plus or minus what you adjusted - and every one of those terms is a row somebody entered on another tab. A stock figure you can type over is a stock figure that has stopped meaning anything, which is how most inventory spreadsheets quietly fail in their second month.

**Reserved stock is separate from available stock.** Confirming a wholesale order takes those units out of availability without pretending they have shipped, so the same ninety candles cannot be promised to two shops. The order register tells you what you are short before you promise it, not after.

**One costing method, used everywhere.** Weighted average, from the opening value plus everything received, applied identically in the valuation, in the retail margin, in the wholesale line profit and in the profit-by-product ranking. Mixing methods between tabs is the most common way these workbooks produce two different answers to the same question.

**Receiving twice cannot double-count.** A purchase order line that appears twice for the same SKU is flagged before the delivery is counted, because the easiest way to invent stock that does not exist is to receive the same pallet on two rows.

Nothing is hard coded. The markup a new product starts at, the margin you refuse to go under, the lead time, the days of safety stock, the window the sales rate averages over, when stock counts as slow or dead, the marketplace and payment fees, the tax rate and whether tax applies at all, the three price breaks, the five customer tiers and their multipliers, every target, every alert threshold and all 40 drop-down lists are settings, not assumptions.

One thing it refuses to do, deliberately. It never claims to synchronise with Etsy, Shopify, Amazon or anything else - you enter or paste your orders and the formulas do the arithmetic. The Getting Started tab says so in plain words and the validator fails the build if any cell claims otherwise.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_inventory_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --business TEXT       your business name on the cover
  --owner TEXT          the owner's name
  --currency SYMBOL     currency symbol  (default: $)
  --currency-code CODE  three letter code (default: USD)
  --markup NUMBER       the markup a new product starts at
  --min-margin PERCENT  the margin you refuse to go under
  --lead-time DAYS      lead time when a supplier has none set
  --safety-days N       days of safety stock in the reorder point
  --sales-window N      days the sales rate averages over
  --marketplace-fee P   marketplace commission percentage
  --payment-fee P       payment processing percentage
  --tax PERCENT         sales tax or VAT on sales
  --revenue-target AMT  revenue target for the year
  --margin-target P     the gross margin you are aiming at
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
  * [`build_inventory_workbook.py`](#build_inventory_workbookpy)
* [2. The package](#2-the-package)
  * [`inventory_wholesale_spreadsheet/__init__.py`](#inventory_wholesale_spreadsheet__init__py)
  * [`inventory_wholesale_spreadsheet/__main__.py`](#inventory_wholesale_spreadsheet__main__py)
  * [`inventory_wholesale_spreadsheet/cli.py`](#inventory_wholesale_spreadsheetclipy)
  * [`inventory_wholesale_spreadsheet/config.py`](#inventory_wholesale_spreadsheetconfigpy)
  * [`inventory_wholesale_spreadsheet/layout.py`](#inventory_wholesale_spreadsheetlayoutpy)
  * [`inventory_wholesale_spreadsheet/formats.py`](#inventory_wholesale_spreadsheetformatspy)
  * [`inventory_wholesale_spreadsheet/writer.py`](#inventory_wholesale_spreadsheetwriterpy)
  * [`inventory_wholesale_spreadsheet/names.py`](#inventory_wholesale_spreadsheetnamespy)
  * [`inventory_wholesale_spreadsheet/sample_data.py`](#inventory_wholesale_spreadsheetsample_datapy)
  * [`inventory_wholesale_spreadsheet/build.py`](#inventory_wholesale_spreadsheetbuildpy)
* [3. Shared column factories](#3-shared-column-factories)
  * [`inventory_wholesale_spreadsheet/sheets/__init__.py`](#inventory_wholesale_spreadsheetsheets__init__py)
  * [`inventory_wholesale_spreadsheet/sheets/common.py`](#inventory_wholesale_spreadsheetsheetscommonpy)
* [4. Settings and the two front pages](#4-settings-and-the-two-front-pages)
  * [`inventory_wholesale_spreadsheet/sheets/settings.py`](#inventory_wholesale_spreadsheetsheetssettingspy)
  * [`inventory_wholesale_spreadsheet/sheets/dashboard.py`](#inventory_wholesale_spreadsheetsheetsdashboardpy)
  * [`inventory_wholesale_spreadsheet/sheets/getting_started.py`](#inventory_wholesale_spreadsheetsheetsgetting_startedpy)
* [5. Product, SKU and stock](#5-product-sku-and-stock)
  * [`inventory_wholesale_spreadsheet/sheets/products.py`](#inventory_wholesale_spreadsheetsheetsproductspy)
  * [`inventory_wholesale_spreadsheet/sheets/inventory.py`](#inventory_wholesale_spreadsheetsheetsinventorypy)
  * [`inventory_wholesale_spreadsheet/sheets/movements.py`](#inventory_wholesale_spreadsheetsheetsmovementspy)
* [6. Supplier, purchase order, receiving](#6-supplier-purchase-order-receiving)
  * [`inventory_wholesale_spreadsheet/sheets/suppliers.py`](#inventory_wholesale_spreadsheetsheetssupplierspy)
  * [`inventory_wholesale_spreadsheet/sheets/purchases.py`](#inventory_wholesale_spreadsheetsheetspurchasespy)
* [7. Wholesale customer, tiered price, fulfilment](#7-wholesale-customer-tiered-price-fulfilment)
  * [`inventory_wholesale_spreadsheet/sheets/customers.py`](#inventory_wholesale_spreadsheetsheetscustomerspy)
  * [`inventory_wholesale_spreadsheet/sheets/pricing.py`](#inventory_wholesale_spreadsheetsheetspricingpy)
  * [`inventory_wholesale_spreadsheet/sheets/orders.py`](#inventory_wholesale_spreadsheetsheetsorderspy)
* [8. Sales, reordering and what the stock is worth](#8-sales-reordering-and-what-the-stock-is-worth)
  * [`inventory_wholesale_spreadsheet/sheets/retail.py`](#inventory_wholesale_spreadsheetsheetsretailpy)
  * [`inventory_wholesale_spreadsheet/sheets/reorder.py`](#inventory_wholesale_spreadsheetsheetsreorderpy)
  * [`inventory_wholesale_spreadsheet/sheets/valuation.py`](#inventory_wholesale_spreadsheetsheetsvaluationpy)
* [9. Returns, locations and counting the shelf](#9-returns-locations-and-counting-the-shelf)
  * [`inventory_wholesale_spreadsheet/sheets/returns.py`](#inventory_wholesale_spreadsheetsheetsreturnspy)
  * [`inventory_wholesale_spreadsheet/sheets/locations.py`](#inventory_wholesale_spreadsheetsheetslocationspy)
  * [`inventory_wholesale_spreadsheet/sheets/audit.py`](#inventory_wholesale_spreadsheetsheetsauditpy)
* [10. What each product actually made](#10-what-each-product-actually-made)
  * [`inventory_wholesale_spreadsheet/sheets/performance.py`](#inventory_wholesale_spreadsheetsheetsperformancepy)
  * [`inventory_wholesale_spreadsheet/sheets/expenses.py`](#inventory_wholesale_spreadsheetsheetsexpensespy)
  * [`inventory_wholesale_spreadsheet/sheets/profit.py`](#inventory_wholesale_spreadsheetsheetsprofitpy)
  * [`inventory_wholesale_spreadsheet/sheets/monthly.py`](#inventory_wholesale_spreadsheetsheetsmonthlypy)
  * [`inventory_wholesale_spreadsheet/sheets/annual.py`](#inventory_wholesale_spreadsheetsheetsannualpy)
  * [`inventory_wholesale_spreadsheet/sheets/printables.py`](#inventory_wholesale_spreadsheetsheetsprintablespy)
* [11. The quality tools](#11-the-quality-tools)
  * [`tools/validate_inventory_workbook.py`](#toolsvalidate_inventory_workbookpy)
  * [`tools/check_parens.py`](#toolscheck_parenspy)
  * [`tools/check_inventory_specs.py`](#toolscheck_inventory_specspy)
  * [`tools/check_inventory_sample.py`](#toolscheck_inventory_samplepy)
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

`build_inventory_workbook.py` parses the command line into a `Config`, then
`inventory_wholesale_spreadsheet.build.build_workbook(path, cfg)` does six things in order:

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
| 1 | **Dashboard** | page | `inventory_wholesale_spreadsheet/sheets/dashboard.py` | - | - |
| 2 | **Getting Started** | page | `inventory_wholesale_spreadsheet/sheets/getting_started.py` | - | - |
| 3 | **Settings & Lists** | page | `inventory_wholesale_spreadsheet/sheets/settings.py` | - | - |
| 4 | **Product & SKU Database** | register | `inventory_wholesale_spreadsheet/sheets/products.py` | 41 | 300 |
| 5 | **Inventory Tracker** | register | `inventory_wholesale_spreadsheet/sheets/inventory.py` | 33 | 300 |
| 6 | **Stock Movement Log** | register | `inventory_wholesale_spreadsheet/sheets/movements.py` | 25 | 900 |
| 7 | **Supplier Database** | register | `inventory_wholesale_spreadsheet/sheets/suppliers.py` | 39 | 60 |
| 8 | **Purchase Orders** | register | `inventory_wholesale_spreadsheet/sheets/purchases.py` | 48 | 500 |
| 9 | **Wholesale Customer CRM** | register | `inventory_wholesale_spreadsheet/sheets/customers.py` | 42 | 120 |
| 10 | **Wholesale Pricing** | register | `inventory_wholesale_spreadsheet/sheets/pricing.py` | 31 | 300 |
| 11 | **Wholesale Orders** | register | `inventory_wholesale_spreadsheet/sheets/orders.py` | 50 | 600 |
| 12 | **Retail Sales Tracker** | register | `inventory_wholesale_spreadsheet/sheets/retail.py` | 34 | 900 |
| 13 | **Reorder Planner** | register | `inventory_wholesale_spreadsheet/sheets/reorder.py` | 33 | 300 |
| 14 | **Inventory Valuation & COGS** | register | `inventory_wholesale_spreadsheet/sheets/valuation.py` | 28 | 300 |
| 15 | **Returns & Damaged Goods** | register | `inventory_wholesale_spreadsheet/sheets/returns.py` | 32 | 200 |
| 16 | **Warehouse Locations** | register | `inventory_wholesale_spreadsheet/sheets/locations.py` | 21 | 200 |
| 17 | **Stock Audit** | register | `inventory_wholesale_spreadsheet/sheets/audit.py` | 27 | 300 |
| 18 | **Product Performance** | register | `inventory_wholesale_spreadsheet/sheets/performance.py` | 35 | 300 |
| 19 | **Expense Tracker** | register | `inventory_wholesale_spreadsheet/sheets/expenses.py` | 24 | 400 |
| 20 | **Profit Analysis** | register | `inventory_wholesale_spreadsheet/sheets/profit.py` | 26 | 14 |
| 21 | **Monthly Reports** | page | `inventory_wholesale_spreadsheet/sheets/monthly.py` | - | - |
| 22 | **Annual Reports** | page | `inventory_wholesale_spreadsheet/sheets/annual.py` | - | - |
| 23 | **Printable Orders & Invoices** | page | `inventory_wholesale_spreadsheet/sheets/printables.py` | - | - |

### Every file at a glance

| File | Lines | What it is |
|------|------:|------------|
| [`build_inventory_workbook.py`](#build_inventory_workbookpy) | 12 | Build the Ultimate Inventory & Wholesale Order Manager workbook |
| [`inventory_wholesale_spreadsheet/__init__.py`](#inventory_wholesale_spreadsheet__init__py) | 19 | Ultimate Inventory & Wholesale Order Manager |
| [`inventory_wholesale_spreadsheet/__main__.py`](#inventory_wholesale_spreadsheet__main__py) | 5 | ``python -m inventory_wholesale_spreadsheet`` |
| [`inventory_wholesale_spreadsheet/cli.py`](#inventory_wholesale_spreadsheetclipy) | 201 | Command line front end |
| [`inventory_wholesale_spreadsheet/config.py`](#inventory_wholesale_spreadsheetconfigpy) | 520 | Every number, name, list and disclaimer the workbook uses |
| [`inventory_wholesale_spreadsheet/layout.py`](#inventory_wholesale_spreadsheetlayoutpy) | 253 | Layout engine |
| [`inventory_wholesale_spreadsheet/formats.py`](#inventory_wholesale_spreadsheetformatspy) | 202 | Cell format factory |
| [`inventory_wholesale_spreadsheet/writer.py`](#inventory_wholesale_spreadsheetwriterpy) | 497 | Sheet writing helpers shared by every tab of the workbook |
| [`inventory_wholesale_spreadsheet/names.py`](#inventory_wholesale_spreadsheetnamespy) | 314 | Workbook level defined names that point at the register sheets |
| [`inventory_wholesale_spreadsheet/sample_data.py`](#inventory_wholesale_spreadsheetsample_datapy) | 1,322 | A worked example: one small supply business, mid-year, mid-mess |
| [`inventory_wholesale_spreadsheet/build.py`](#inventory_wholesale_spreadsheetbuildpy) | 168 | Workbook builder: puts the twenty-three tabs together in order |
| [`inventory_wholesale_spreadsheet/sheets/__init__.py`](#inventory_wholesale_spreadsheetsheets__init__py) | 0 |  |
| [`inventory_wholesale_spreadsheet/sheets/common.py`](#inventory_wholesale_spreadsheetsheetscommonpy) | 691 | Small building blocks shared by the register sheets |
| [`inventory_wholesale_spreadsheet/sheets/settings.py`](#inventory_wholesale_spreadsheetsheetssettingspy) | 776 | `Settings & Lists` - every number and every list the workbook runs on |
| [`inventory_wholesale_spreadsheet/sheets/dashboard.py`](#inventory_wholesale_spreadsheetsheetsdashboardpy) | 402 | `Dashboard` - the page that answers "what do I need to do today?" Everything here is worked out from the... |
| [`inventory_wholesale_spreadsheet/sheets/getting_started.py`](#inventory_wholesale_spreadsheetsheetsgetting_startedpy) | 297 | `Getting Started` - twenty minutes, in the right order |
| [`inventory_wholesale_spreadsheet/sheets/products.py`](#inventory_wholesale_spreadsheetsheetsproductspy) | 259 | 4 |
| [`inventory_wholesale_spreadsheet/sheets/inventory.py`](#inventory_wholesale_spreadsheetsheetsinventorypy) | 285 | 5 |
| [`inventory_wholesale_spreadsheet/sheets/movements.py`](#inventory_wholesale_spreadsheetsheetsmovementspy) | 195 | 6 |
| [`inventory_wholesale_spreadsheet/sheets/suppliers.py`](#inventory_wholesale_spreadsheetsheetssupplierspy) | 239 | 7 |
| [`inventory_wholesale_spreadsheet/sheets/purchases.py`](#inventory_wholesale_spreadsheetsheetspurchasespy) | 339 | 8 |
| [`inventory_wholesale_spreadsheet/sheets/customers.py`](#inventory_wholesale_spreadsheetsheetscustomerspy) | 271 | 9 |
| [`inventory_wholesale_spreadsheet/sheets/pricing.py`](#inventory_wholesale_spreadsheetsheetspricingpy) | 227 | 10 |
| [`inventory_wholesale_spreadsheet/sheets/orders.py`](#inventory_wholesale_spreadsheetsheetsorderspy) | 338 | 11 |
| [`inventory_wholesale_spreadsheet/sheets/retail.py`](#inventory_wholesale_spreadsheetsheetsretailpy) | 230 | 12 |
| [`inventory_wholesale_spreadsheet/sheets/reorder.py`](#inventory_wholesale_spreadsheetsheetsreorderpy) | 272 | 13 |
| [`inventory_wholesale_spreadsheet/sheets/valuation.py`](#inventory_wholesale_spreadsheetsheetsvaluationpy) | 237 | 14 |
| [`inventory_wholesale_spreadsheet/sheets/returns.py`](#inventory_wholesale_spreadsheetsheetsreturnspy) | 205 | 15 |
| [`inventory_wholesale_spreadsheet/sheets/locations.py`](#inventory_wholesale_spreadsheetsheetslocationspy) | 161 | 16 |
| [`inventory_wholesale_spreadsheet/sheets/audit.py`](#inventory_wholesale_spreadsheetsheetsauditpy) | 198 | 17 |
| [`inventory_wholesale_spreadsheet/sheets/performance.py`](#inventory_wholesale_spreadsheetsheetsperformancepy) | 302 | 18 |
| [`inventory_wholesale_spreadsheet/sheets/expenses.py`](#inventory_wholesale_spreadsheetsheetsexpensespy) | 173 | 19 |
| [`inventory_wholesale_spreadsheet/sheets/profit.py`](#inventory_wholesale_spreadsheetsheetsprofitpy) | 318 | 20 |
| [`inventory_wholesale_spreadsheet/sheets/monthly.py`](#inventory_wholesale_spreadsheetsheetsmonthlypy) | 276 | `Monthly Reports` - one month, properly examined |
| [`inventory_wholesale_spreadsheet/sheets/annual.py`](#inventory_wholesale_spreadsheetsheetsannualpy) | 250 | `Annual Reports` - the year, and whether it was a good one |
| [`inventory_wholesale_spreadsheet/sheets/printables.py`](#inventory_wholesale_spreadsheetsheetsprintablespy) | 354 | `Printable Orders & Invoices` - three documents, ready to send |
| [`tools/validate_inventory_workbook.py`](#toolsvalidate_inventory_workbookpy) | 1,238 | Static checks on a generated Ultimate Inventory & Wholesale Order Manager |
| [`tools/check_parens.py`](#toolscheck_parenspy) | 121 | Parenthesis and quote balance check for every formula in a workbook |
| [`tools/check_inventory_specs.py`](#toolscheck_inventory_specspy) | 206 | Fast structural sweep over the inventory product's register specs |
| [`tools/check_inventory_sample.py`](#toolscheck_inventory_samplepy) | 170 | Validate the inventory product's example data before it is built |
| [`tools/recalc_check.py`](#toolsrecalc_checkpy) | 216 | Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error... |

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_inventory_workbook.py
```

The workbook lands in `output/Ultimate Inventory & Wholesale Order Manager.xlsx`.

### `build_inventory_workbook.py`

*12 lines*

> Build the Ultimate Inventory & Wholesale Order Manager workbook.
>
> python3 build_inventory_workbook.py
> python3 build_inventory_workbook.py -o "out/Stock.xlsx" --blank
> python3 build_inventory_workbook.py --currency "£" --rows 120

````python
#!/usr/bin/env python3
"""Build the Ultimate Inventory & Wholesale Order Manager workbook.

    python3 build_inventory_workbook.py
    python3 build_inventory_workbook.py -o "out/Stock.xlsx" --blank
    python3 build_inventory_workbook.py --currency "£" --rows 120
"""

from inventory_wholesale_spreadsheet.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer, the cross sheet names and the worked example. These are the parts every tab leans on. `config.py` holds every price break, every status set, every alert threshold and every disclaimer, so a candle maker, a boutique and a wholesale distributor can all be matched without touching a formula.

### `inventory_wholesale_spreadsheet/__init__.py`

*19 lines*

> Ultimate Inventory & Wholesale Order Manager.
>
> Build the workbook with::
>
>     python -m inventory_wholesale_spreadsheet -o "output/Inventory.xlsx"
>
> or from Python::
>
>     from inventory_wholesale_spreadsheet import build_workbook
>     build_workbook("Ultimate Inventory & Wholesale Order Manager.xlsx")

````python
"""Ultimate Inventory & Wholesale Order Manager.

Build the workbook with::

    python -m inventory_wholesale_spreadsheet -o "output/Inventory.xlsx"

or from Python::

    from inventory_wholesale_spreadsheet import build_workbook
    build_workbook("Ultimate Inventory & Wholesale Order Manager.xlsx")
"""

from .build import build_workbook
from .config import (AUTHOR, LISTING_TITLE, PRODUCT_NAME,
                     PRODUCT_SUBTITLE, VERSION, Config)

__all__ = ["build_workbook", "Config", "AUTHOR", "PRODUCT_NAME",
           "PRODUCT_SUBTITLE", "LISTING_TITLE", "VERSION"]
__version__ = VERSION
````

### `inventory_wholesale_spreadsheet/__main__.py`

*5 lines*

> ``python -m inventory_wholesale_spreadsheet``.

````python
"""``python -m inventory_wholesale_spreadsheet``."""

from .cli import main

raise SystemExit(main())
````

### `inventory_wholesale_spreadsheet/cli.py`

*201 lines · 3 functions*

> Command line front end.
>
> python -m inventory_wholesale_spreadsheet
> python -m inventory_wholesale_spreadsheet -o "out/Stock.xlsx" --blank
> python -m inventory_wholesale_spreadsheet --currency "£" --rows 120

| Defines | What it does |
|---------|--------------|
| `build_parser()` |  |
| `config_from_args(args)` |  |
| `main(argv)` |  |

**Module constants:** `DEFAULT_OUTPUT` (os.path.join(...)).

````python
"""Command line front end.

    python -m inventory_wholesale_spreadsheet
    python -m inventory_wholesale_spreadsheet -o "out/Stock.xlsx" --blank
    python -m inventory_wholesale_spreadsheet --currency "£" --rows 120
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import (AUTHOR, DEFAULT_ROWS, LISTING_TITLE,
                     PRICE_TIERS, PRODUCT_NAME, SCALE_ANCHOR, Config)

DEFAULT_OUTPUT = os.path.join(
    "output", "Ultimate Inventory & Wholesale Order Manager.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="inventory_wholesale_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty (no example "
                             "business)")
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
                        help="pre-fill the owner's name")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="the year the reports open on")
    parser.add_argument("--markup", type=float, default=None,
                        help="default retail markup on cost, "
                             "e.g. 2.5")
    parser.add_argument("--min-margin", type=float, default=None,
                        help="minimum acceptable gross margin as a "
                             "percentage, e.g. 25 for 25%%")
    parser.add_argument("--lead-time", type=int, default=None,
                        help="default supplier lead time in days")
    parser.add_argument("--safety-days", type=int, default=None,
                        help="safety stock expressed in days of "
                             "cover")
    parser.add_argument("--sales-window", type=int, default=None,
                        help="days of history used for average "
                             "daily sales")
    parser.add_argument("--marketplace-fee", type=float, default=None,
                        help="default marketplace commission as a "
                             "percentage, e.g. 6.5")
    parser.add_argument("--payment-fee", type=float, default=None,
                        help="payment processing as a percentage, "
                             "e.g. 2.9")
    parser.add_argument("--tax", type=float, default=None,
                        help="sales tax or VAT as a percentage, "
                             "e.g. 20 for 20%% (zero by default)")
    parser.add_argument("--revenue-target", type=float, default=None,
                        help="revenue target for the year")
    parser.add_argument("--margin-target", type=float, default=None,
                        help="gross margin target as a percentage, "
                             "e.g. 55 for 55%%")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the defaults are "
                             "scaled to fit this many product rows)")
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
    if args.markup is not None:
        kwargs["default_markup"] = args.markup
    if args.min_margin is not None:
        kwargs["minimum_margin"] = round(args.min_margin / 100.0, 4)
    if args.lead_time is not None:
        kwargs["default_lead_time_days"] = args.lead_time
    if args.safety_days is not None:
        kwargs["safety_stock_days"] = args.safety_days
    if args.sales_window is not None:
        kwargs["sales_history_days"] = args.sales_window
    if args.marketplace_fee is not None:
        kwargs["marketplace_fee_pct"] = round(
            args.marketplace_fee / 100.0, 5)
    if args.payment_fee is not None:
        kwargs["payment_fee_pct"] = round(args.payment_fee / 100.0, 5)
    if args.tax is not None:
        kwargs["tax_rate"] = round(args.tax / 100.0, 5)
        kwargs["tax_on_sales"] = args.tax > 0
    if args.revenue_target is not None:
        kwargs["target_revenue_year"] = args.revenue_target
        kwargs["target_revenue_month"] = round(
            args.revenue_target / 12.0, 2)
    if args.margin_target is not None:
        kwargs["target_gross_margin"] = round(
            args.margin_target / 100.0, 4)
    if args.today:
        kwargs["_today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.business:
        cfg.business.name = args.business
    if args.owner:
        cfg.business.owner = args.owner
    if args.rows:
        # every register scales against the Product & SKU Database,
        # which is what `row_count()` and `scale_rows()` both read
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
    sign = cfg.currency
    breaks = ", ".join(str(qty) for qty, _label in PRICE_TIERS)

    print(f"{PRODUCT_NAME}")
    print(f"  listing title : {LISTING_TITLE}")
    print(f"  author        : {cfg.author}")
    print(f"  sheets        : {len(SHEET_ORDER)}")
    print(f"  business      : {cfg.business.name} "
          f"({cfg.business.owner})")
    print(f"  today         : {cfg.today.isoformat()}")
    print(f"  register rows : {rows:,} across "
          f"{len(DEFAULT_ROWS)} registers "
          f"(scaled on {SCALE_ANCHOR})")
    print("  costing       : weighted average, everywhere")
    print(f"  price breaks  : {breaks}")
    print(f"  margin floor  : {cfg.minimum_margin:.0%} "
          f"(target {cfg.target_gross_margin:.0%})")
    print(f"  reordering    : (daily sales x lead time) + "
          f"{cfg.safety_stock_days} days safety, "
          f"{cfg.sales_history_days}-day window")
    print(f"  lead time     : {cfg.default_lead_time_days} days "
          f"default, overridden per supplier")
    print(f"  selling costs : marketplace "
          f"{cfg.marketplace_fee_pct:.1%}, payment "
          f"{cfg.payment_fee_pct:.1%} + {sign}"
          f"{cfg.payment_fee_fixed:.2f}")
    print(f"  tax           : {cfg.tax_rate:.2%} "
          f"({'on' if cfg.tax_on_sales else 'off'})")
    print(f"  targets       : {sign}{cfg.target_revenue_year:,.0f} "
          f"revenue, {cfg.target_turnover:.1f} stock turns")
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

### `inventory_wholesale_spreadsheet/config.py`

*520 lines · 3 classes*

> Every number, name, list and disclaimer the workbook uses.
>
> Nothing downstream invents a figure. If a price, a lead time, a status
> or a drop-down value appears on a sheet, it came from here, which is
> why a seller in another country on another costing basis can be matched
> without touching a formula.
>
> Two conventions worth knowing before you read the rest:
>
> * **SKU is the join key.** Stock, purchases, sales, returns, audits,
>   valuation and profit all match on it. It is a drop-down everywhere
>   except the Product Database, where it is created.
> * **Stock is derived, never typed.** The only way to change a stock
>   figure is to add a row that says why - a purchase received, a sale
>   shipped, a return restocked, or an explicit adjustment in the Stock
>   Movement Log.

| Defines | What it does |
|---------|--------------|
| `class Palette` | A cool, slightly industrial scheme |
| `class Business` | fields: name, owner, email, phone, website, shop, address_1, address_2 |
| `class Config` | fields: business, author, password, protect_sheets, include_sample_data, currency, currency_code, _today |

**Module constants:** `PRODUCT_NAME` ('Ultimate Inventory & Wholesale Order Manager'), `PRODUCT_SHORT` ('Inventory & Wholesale Manager'), `PRODUCT_SUBTITLE` ('Products, stock, suppliers, purchase orders, wholesale c...), `LISTING_TITLE` ('Inventory Management Spreadsheet | SKU Stock Tracker, Su...), `VERSION` ('1.0.0'), `AUTHOR` ('Novality Store'), `PASSWORD` ('premium'), `MAX_ROWS_SHOWN` (10), `COVER_WARNING_DAYS` (21), `HONESTY_NOTE` ('This workbook does not connect to Etsy, Shopify, Amazon,...), `STOCK_NOTE` ('Stock is never typed. There is no cell anywhere in this ...), `COSTING_NOTE` ('This workbook uses WEIGHTED AVERAGE COST throughout. Not...), `RESERVED_NOTE` ('Reserved stock is stock you still physically have but ha...), `PRIVACY_NOTE` ("Hold the business details you need to trade: a company n...) and 27 more.

````python
"""Every number, name, list and disclaimer the workbook uses.

Nothing downstream invents a figure. If a price, a lead time, a status
or a drop-down value appears on a sheet, it came from here, which is
why a seller in another country on another costing basis can be matched
without touching a formula.

Two conventions worth knowing before you read the rest:

* **SKU is the join key.** Stock, purchases, sales, returns, audits,
  valuation and profit all match on it. It is a drop-down everywhere
  except the Product Database, where it is created.
* **Stock is derived, never typed.** The only way to change a stock
  figure is to add a row that says why - a purchase received, a sale
  shipped, a return restocked, or an explicit adjustment in the Stock
  Movement Log.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# the product
# ---------------------------------------------------------------------------

PRODUCT_NAME = "Ultimate Inventory & Wholesale Order Manager"
PRODUCT_SHORT = "Inventory & Wholesale Manager"
PRODUCT_SUBTITLE = ("Products, stock, suppliers, purchase orders, "
                    "wholesale customers, tiered pricing and profit by "
                    "product - in one workbook")
LISTING_TITLE = ("Inventory Management Spreadsheet | SKU Stock Tracker, "
                 "Supplier Database, Purchase Orders, Wholesale Orders "
                 "& Profit Dashboard")
VERSION = "1.0.0"
AUTHOR = "Novality Store"
PASSWORD = "premium"

#: how many rows the dashboard lists before it stops
MAX_ROWS_SHOWN = 10
#: the "needs reordering soon" window, in days of cover
COVER_WARNING_DAYS = 21


# ---------------------------------------------------------------------------
# honesty
# ---------------------------------------------------------------------------

HONESTY_NOTE = (
    "This workbook does not connect to Etsy, Shopify, Amazon, eBay or "
    "any other marketplace. There is no sync, no API and no import "
    "robot. You type or paste your orders in, and every stock figure, "
    "cost and profit number on the other twenty-two tabs follows from "
    "what you typed. Anything that claimed otherwise would be lying to "
    "you.")

STOCK_NOTE = (
    "Stock is never typed. There is no cell anywhere in this workbook "
    "where you can simply overwrite a quantity, and that is "
    "deliberate. Current stock is calculated from opening stock plus "
    "what you received, minus what you sold, minus what was damaged, "
    "plus or minus any adjustment you logged. If a number looks wrong, "
    "there is a row that explains it.")

COSTING_NOTE = (
    "This workbook uses WEIGHTED AVERAGE COST throughout. Not FIFO, "
    "not LIFO, not specific identification - weighted average, "
    "everywhere, with no exceptions. Opening stock at its stated cost "
    "plus everything received at what it actually cost, divided by the "
    "units those two represent. Mixing costing methods is the single "
    "most common way a stock spreadsheet stops reconciling, so this "
    "one does not offer you the choice. If your accountant requires "
    "FIFO, tell them which method this is before you hand it over.")

RESERVED_NOTE = (
    "Reserved stock is stock you still physically have but have "
    "already promised to somebody. A confirmed wholesale order "
    "reserves its units; it does not remove them. They leave stock "
    "when the order is marked Shipped. That is why Available Stock "
    "and Current Stock are two different columns, and why selling "
    "against Available is the only safe thing to do.")

PRIVACY_NOTE = (
    "Hold the business details you need to trade: a company name, a "
    "contact, an address, payment terms. You do not need a customer's "
    "card number or a copy of anybody's passport in a spreadsheet, so "
    "there is nowhere here to put one. A resale certificate number is "
    "optional and only there because some jurisdictions require you to "
    "keep it.")

DISCLAIMER = (
    "This is a business organisation tool, not accounting software and "
    "not tax advice. Inventory valuation, cost of goods sold and the "
    "profit figures here are management numbers for making decisions. "
    "How stock and COGS must be reported, and when, depends on your "
    "jurisdiction, your business structure and your accounting basis - "
    "ask an accountant, and tell them this workbook uses weighted "
    "average cost.")


# ---------------------------------------------------------------------------
# colour
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Palette:
    """A cool, slightly industrial scheme. Stock work is not romantic."""

    ink: str = "#1C2B33"
    deep: str = "#24404D"
    primary: str = "#2E6E7E"
    primary_mid: str = "#59929F"
    primary_soft: str = "#E8F1F3"
    accent: str = "#C2703D"
    good: str = "#2F7D55"
    warn: str = "#B5802A"
    bad: str = "#B33A3A"
    muted: str = "#6D7F87"
    paper: str = "#FCFEFE"
    rule: str = "#C9DBDF"

    # ---- derived tones the formats and the writer ask for ----------
    white: str = "#FFFFFF"
    input_bg: str = "#FFFFFF"
    input_border: str = "#CFE1E5"
    calc_bg: str = "#F5FAFB"
    calc_border: str = "#E2EFF1"
    line: str = "#C9DBDF"
    note_bg: str = "#F8FCFC"
    grey_text: str = "#55676F"
    good_bg: str = "#DCEDE3"
    good_text: str = "#1F5739"
    warn_bg: str = "#F9EBD2"
    warn_text: str = "#7C5712"
    bad_bg: str = "#F7DCDC"
    bad_text: str = "#8C2626"
    info_bg: str = "#E2EAF2"
    info_text: str = "#2E4A6B"
    accent_soft: str = "#FBEADF"


PALETTE = Palette()


# ---------------------------------------------------------------------------
# the business
# ---------------------------------------------------------------------------

@dataclass
class Business:
    name: str = "Thornfield Supply Co."
    owner: str = "Rosa Thornfield"
    email: str = "orders@thornfieldsupply.example"
    phone: str = "(503) 555-0172"
    website: str = "thornfieldsupply.example"
    shop: str = "ThornfieldSupplyCo"
    address_1: str = "Unit 7, Alder Works"
    address_2: str = "118 Kestrel Road"
    city: str = "Portland"
    region: str = "Oregon"
    postcode: str = "97202"
    country: str = "United States"
    tax_id: str = "US-EIN 88-4471203"


# ---------------------------------------------------------------------------
# identifiers - every one of them is a join key somewhere
# ---------------------------------------------------------------------------

ID_PREFIXES: Dict[str, Tuple[str, int]] = {
    "product": ("PRD-", 4),
    "supplier": ("SUP-", 3),
    "purchase": ("PO-", 4),
    "customer": ("WC-", 3),
    "order": ("WO-", 4),
    "sale": ("SO-", 4),
    "movement": ("MOV-", 5),
    "return": ("RET-", 4),
    "audit": ("AUD-", 4),
    "location": ("LOC-", 3),
    "expense": ("EXP-", 4),
    "invoice": ("INV-", 4),
}

#: (label, default, defined name, digits) - rendered as a block on
#: Settings so a seller can use their own code scheme
ID_PREFIX_FIELDS: List[Tuple[str, str, str, int]] = [
    ("Product ID prefix", "PRD-", "ProductPrefix", 4),
    ("Supplier ID prefix", "SUP-", "SupplierPrefix", 3),
    ("Purchase order prefix", "PO-", "PurchasePrefix", 4),
    ("Wholesale customer prefix", "WC-", "CustomerPrefix", 3),
    ("Wholesale order prefix", "WO-", "OrderPrefix", 4),
    ("Retail sale prefix", "SO-", "SalePrefix", 4),
    ("Stock movement prefix", "MOV-", "MovementPrefix", 5),
    ("Return prefix", "RET-", "ReturnPrefix", 4),
    ("Stock audit prefix", "AUD-", "AuditPrefix", 4),
    ("Warehouse location prefix", "LOC-", "LocationPrefix", 3),
    ("Expense prefix", "EXP-", "ExpensePrefix", 4),
    ("Invoice prefix", "INV-", "InvoicePrefix", 4),
]


# ---------------------------------------------------------------------------
# status vocabularies
# ---------------------------------------------------------------------------
# The sets below are what the stock arithmetic actually keys on, so
# they are defined once here and imported by sheets/common.py rather
# than retyped in nine registers.

#: a wholesale order that has committed stock but not shipped it
RESERVING_STATUSES = ["Confirmed", "Processing", "Picking", "Packed"]
#: a wholesale order whose stock has physically left
SHIPPED_STATUSES = ["Shipped", "Delivered"]
#: a retail sale that has taken stock off the shelf
SOLD_RETAIL_STATUSES = ["Shipped", "Delivered", "Collected"]
#: a purchase order still expected to arrive
OPEN_PO_STATUSES = ["Sent", "Confirmed", "Partially Received"]
#: a purchase order the supplier has agreed to - stock on order
COMMITTED_PO_STATUSES = ["Confirmed", "Partially Received",
                         "Fully Received"]
#: a returned item that can go back on the shelf
SELLABLE_CONDITIONS = ["As new", "Opened - resellable"]

PO_STATUSES = ["Draft", "Pending Approval", "Sent", "Confirmed",
               "Partially Received", "Fully Received", "Cancelled"]

ORDER_STATUSES = ["New", "Confirmed", "Processing", "Picking",
                  "Packed", "Shipped", "Delivered", "Cancelled",
                  "Returned"]

SALES_CHANNELS = ["Etsy", "Shopify", "Amazon", "eBay", "WooCommerce",
                  "Physical Store", "Direct Sales", "Craft Fair",
                  "Wholesale", "Other"]


# ---------------------------------------------------------------------------
# the worked example's catalogue shape
# ---------------------------------------------------------------------------

#: (category, [subcategories])
CATEGORIES: List[Tuple[str, List[str]]] = [
    ("Candles", ["Jar candles", "Pillar candles", "Tealights",
                 "Wax melts"]),
    ("Ceramics", ["Mugs", "Bowls", "Planters", "Vases"]),
    ("Textiles", ["Tea towels", "Cushion covers", "Aprons",
                  "Tote bags"]),
    ("Stationery", ["Notebooks", "Greeting cards", "Print sets",
                    "Stickers"]),
    ("Bath & Body", ["Soap bars", "Bath salts", "Balms"]),
    ("Packaging", ["Boxes", "Tissue", "Labels", "Mailers"]),
]

UNITS = ["Each", "Pair", "Pack of 2", "Pack of 4", "Pack of 6",
         "Pack of 10", "Pack of 12", "Box of 24", "Case of 48",
         "Metre", "Kilogram", "Litre"]

#: quantity break, label - the default tier ladder on new products
PRICE_TIERS: List[Tuple[int, str]] = [
    (10, "Tier 1 - 10+"),
    (50, "Tier 2 - 50+"),
    (100, "Tier 3 - 100+"),
]

#: how much below retail each customer tier buys at, as a multiplier
CUSTOMER_TIERS: List[Tuple[str, float, str]] = [
    ("Retail", 1.00, "Walk-in and marketplace buyers"),
    ("Trade", 0.60, "Small shops buying a dozen at a time"),
    ("Wholesale", 0.50, "Regular stockists on standard terms"),
    ("Distributor", 0.42, "Volume buyers who resell to shops"),
    ("Key account", 0.38, "Negotiated - check the margin first"),
]

PAYMENT_TERMS: List[Tuple[str, int]] = [
    ("Prepaid", 0),
    ("On delivery", 0),
    ("Net 7", 7),
    ("Net 14", 14),
    ("Net 30", 30),
    ("Net 45", 45),
    ("Net 60", 60),
    ("End of month", 30),
]

#: movement types the Stock Movement Log offers, and the sign each one
#: has on stock: +1 adds, -1 removes, 0 is a correction that carries
#: its own sign
MOVEMENT_TYPES: List[Tuple[str, int, str]] = [
    ("Opening balance", 1, "Stock you already had on day one"),
    ("Purchase received", 1, "Booked in against a purchase order"),
    ("Customer return", 1, "Came back and went on the shelf"),
    ("Found in count", 1, "A stock count found more than expected"),
    ("Transfer in", 1, "Arrived from another location"),
    ("Sale", -1, "Sold and despatched"),
    ("Supplier return", -1, "Sent back to the supplier"),
    ("Damaged", -1, "Broken, spoiled or otherwise unsellable"),
    ("Lost or stolen", -1, "Gone, with no paperwork"),
    ("Sample or gift", -1, "Given away deliberately"),
    ("Used in production", -1, "Consumed making something else"),
    ("Transfer out", -1, "Sent to another location"),
    ("Adjustment", 0, "A correction - say why in the notes"),
]

RETURN_REASONS = ["Damaged in transit", "Defective", "Wrong item sent",
                  "Incorrect size", "Customer changed mind",
                  "Late delivery", "Not as described", "Other"]

RETURN_CONDITIONS = ["As new", "Opened - resellable",
                     "Damaged - not sellable", "Faulty - return to "
                     "supplier", "Destroyed"]

EXPENSE_CATEGORIES = [
    "Stock purchases", "Raw materials", "Packaging & mailers",
    "Shipping & postage", "Courier & freight", "Import duty & customs",
    "Warehouse rent", "Storage & fulfilment", "Utilities",
    "Marketplace fees",
    "Payment processing fees", "Advertising & marketing",
    "Website & hosting", "Software subscriptions", "Photography",
    "Printing & labels", "Equipment", "Equipment repair",
    "Wages & contractors", "Insurance", "Accounting & legal",
    "Bank charges", "Travel", "Trade shows & markets",
    "Samples & giveaways", "Office expenses", "Training",
    "Tax & national insurance", "Miscellaneous",
]

#: expenses that are really the cost of buying stock, not running the
#: business - counting them twice is how a margin ends up negative
STOCK_CATEGORIES = ["Stock purchases", "Raw materials"]

TAX_CATEGORIES = ["Cost of sale", "Operating expense",
                  "Capital / asset", "Owner drawing",
                  "Not deductible", "Review with accountant"]

MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]


# ---------------------------------------------------------------------------
# how many rows each register gets
# ---------------------------------------------------------------------------

SCALE_ANCHOR = "Product & SKU Database"

DEFAULT_ROWS: Dict[str, int] = {
    "Product & SKU Database": 300,
    "Inventory Tracker": 300,
    "Stock Movement Log": 900,
    "Supplier Database": 60,
    "Purchase Orders": 500,
    "Wholesale Customer CRM": 120,
    "Wholesale Pricing": 300,
    "Wholesale Orders": 600,
    "Retail Sales Tracker": 900,
    "Reorder Planner": 300,
    "Inventory Valuation & COGS": 300,
    "Returns & Damaged Goods": 200,
    "Warehouse Locations": 200,
    "Stock Audit": 300,
    "Product Performance": 300,
    "Expense Tracker": 400,
    "Profit Analysis": 300,
}


# ---------------------------------------------------------------------------
# configuration
# ---------------------------------------------------------------------------

@dataclass
class Config:
    business: Business = field(default_factory=Business)
    author: str = AUTHOR
    password: str = PASSWORD
    protect_sheets: bool = True
    include_sample_data: bool = True

    currency: str = "$"
    currency_code: str = "USD"

    _today: Optional[_dt.date] = None
    year: Optional[int] = None

    # ---- pricing -------------------------------------------------------
    default_markup: float = 2.5          # retail = cost x this
    wholesale_of_retail: float = 0.50    # wholesale = retail x this
    distributor_of_retail: float = 0.42
    minimum_margin: float = 0.25         # refuse to sell below this
    default_moq: int = 10                # minimum wholesale order

    # ---- stock ---------------------------------------------------------
    default_lead_time_days: int = 14
    safety_stock_days: int = 10
    default_reorder_point: int = 12
    default_max_stock: int = 200
    sales_history_days: int = 90         # window for average daily sales
    dead_stock_days: int = 120           # no sales for this long = dead
    slow_stock_days: int = 60
    stocktake_months: int = 3            # how often to count

    # ---- orders --------------------------------------------------------
    default_payment_terms: str = "Net 30"
    default_credit_limit: float = 2000.0
    overdue_chase_days: int = 7
    po_chase_days: int = 3               # chase a late delivery after
    pick_warning_days: int = 2           # unshipped after confirmation

    # ---- money ---------------------------------------------------------
    tax_rate: float = 0.0
    tax_name: str = "Sales tax"
    tax_on_sales: bool = False
    default_shipping_charge: float = 12.50
    free_shipping_over: float = 350.0
    marketplace_fee_pct: float = 0.065
    payment_fee_pct: float = 0.029
    payment_fee_fixed: float = 0.30
    packaging_cost_per_order: float = 1.40

    # ---- targets -------------------------------------------------------
    target_gross_margin: float = 0.55
    target_revenue_month: float = 18000.0
    target_revenue_year: float = 216000.0
    target_turnover: float = 4.0         # inventory turns a year
    target_stock_value: float = 42000.0
    target_return_rate: float = 0.03
    target_fill_rate: float = 0.97

    rows_override: Optional[int] = None

    # ----------------------------------------------------------- dates --
    @property
    def today(self) -> _dt.date:
        return self._today or _dt.date.today()

    @today.setter
    def today(self, value: _dt.date) -> None:
        self._today = value

    @property
    def report_year(self) -> int:
        return self.year or self.today.year

    # --------------------------------------------------------- formats --
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
        return 'hh:mm'

    # ---------------------------------------------------------- naming --
    @property
    def version(self) -> str:
        return VERSION

    @property
    def product_name(self) -> str:
        return PRODUCT_NAME

    @property
    def listing_title(self) -> str:
        return LISTING_TITLE

    # ------------------------------------------------------------- ids --
    def prefix(self, key: str) -> str:
        return ID_PREFIXES[key][0]

    def digits(self, key: str) -> int:
        return ID_PREFIXES[key][1]

    # ------------------------------------------------------------ rows --
    def row_count(self, sheet: str) -> int:
        """Rows for a register, scaled if --rows was given."""
        base = DEFAULT_ROWS.get(sheet, 200)
        if not self.rows_override:
            return base
        anchor = DEFAULT_ROWS[SCALE_ANCHOR]
        scaled = int(round(base * self.rows_override / anchor))
        return max(6, scaled)

    def scale_rows(self, rows: int) -> None:
        self.rows_override = rows

    # ----------------------------------------------------------- lists --
    @property
    def category_names(self) -> List[str]:
        return [name for name, _subs in CATEGORIES]

    @property
    def subcategory_names(self) -> List[str]:
        out: List[str] = []
        for _name, subs in CATEGORIES:
            out.extend(subs)
        return out

    @property
    def customer_tier_names(self) -> List[str]:
        return [name for name, _m, _d in CUSTOMER_TIERS]

    @property
    def payment_term_names(self) -> List[str]:
        return [name for name, _d in PAYMENT_TERMS]

    @property
    def movement_type_names(self) -> List[str]:
        return [name for name, _s, _d in MOVEMENT_TYPES]
````

### `inventory_wholesale_spreadsheet/layout.py`

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

### `inventory_wholesale_spreadsheet/formats.py`

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

### `inventory_wholesale_spreadsheet/writer.py`

*497 lines · 1 class · 8 functions*

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

**Module constants:** `P` (Palette), `NAV_LINKS` (13 items), `PROTECT_OPTIONS` (15 entries).

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
    ("Products", "Product & SKU Database"),
    ("Stock", "Inventory Tracker"),
    ("Movements", "Stock Movement Log"),
    ("Suppliers", "Supplier Database"),
    ("Purchases", "Purchase Orders"),
    ("Customers", "Wholesale Customer CRM"),
    ("Wholesale", "Wholesale Orders"),
    ("Retail", "Retail Sales Tracker"),
    ("Reorder", "Reorder Planner"),
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

### `inventory_wholesale_spreadsheet/names.py`

*314 lines · 1 function*

> Workbook level defined names that point at the register sheets.
>
> The fixed drop-downs - categories, payment terms, movement types,
> expense categories - define their own names on `Settings & Lists`.
> The names here are the ones that have to follow live data: the SKUs
> you actually stock, the suppliers you actually buy from, the stockists
> who actually order.  Each is an ``OFFSET`` range that grows and
> shrinks with the filled rows, so a product added this morning is in
> every drop-down in the workbook this afternoon.
>
> `SKUList` is the important one.  It is the SKU column from the Product
> & SKU Database, it feeds the drop-down on nine other tabs, and every
> join in this workbook runs through it.  Get a SKU wrong and the row is
> orphaned; that is why the SKU cell is a list rather than free text
> everywhere except the one sheet where SKUs are created.

| Defines | What it does |
|---------|--------------|
| `define_cross_sheet_names(book, reg)` | Create the live lookup names |

**Module constants:** `CROSS_SHEET_NAMES` (227 entries).

````python
"""Workbook level defined names that point at the register sheets.

The fixed drop-downs - categories, payment terms, movement types,
expense categories - define their own names on `Settings & Lists`.
The names here are the ones that have to follow live data: the SKUs
you actually stock, the suppliers you actually buy from, the stockists
who actually order.  Each is an ``OFFSET`` range that grows and
shrinks with the filled rows, so a product added this morning is in
every drop-down in the workbook this afternoon.

`SKUList` is the important one.  It is the SKU column from the Product
& SKU Database, it feeds the drop-down on nine other tabs, and every
join in this workbook runs through it.  Get a SKU wrong and the row is
orphaned; that is why the SKU cell is a list rather than free text
everywhere except the one sheet where SKUs are created.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # ---- the three join keys ------------------------------------------
    "SKUList": ("Product & SKU Database", "sku"),
    "SupplierList": ("Supplier Database", "id"),
    "CustomerList": ("Wholesale Customer CRM", "id"),

    # ---- product columns the lookups read ------------------------------
    "ProductNames": ("Product & SKU Database", "name"),
    "ProductCategoryCol": ("Product & SKU Database", "category"),
    "ProductCostCol": ("Product & SKU Database", "cost"),
    "ProductRetailCol": ("Product & SKU Database", "retail_price"),
    "ProductWholesaleCol": ("Product & SKU Database",
                            "wholesale_price"),
    "ProductFloorCol": ("Product & SKU Database", "min_price"),
    "ProductOpeningQty": ("Product & SKU Database", "opening_qty"),
    "ProductOpeningValue": ("Product & SKU Database",
                            "opening_value"),
    "ProductStatusCol": ("Product & SKU Database", "status"),
    "ProductSupplierCol": ("Product & SKU Database", "supplier"),
    "ProductMinStock": ("Product & SKU Database", "min_stock"),
    "ProductMaxStock": ("Product & SKU Database", "max_stock"),
    "ProductActive": ("Product & SKU Database", "is_active"),
    "ProductDupCount": ("Product & SKU Database", "dup_count"),
    "ProductAlert": ("Product & SKU Database", "alert"),

    # ---- the stock position - the busiest ranges in the workbook -------
    "StockSKU": ("Inventory Tracker", "sku"),
    "StockProduct": ("Inventory Tracker", "product"),
    "StockCategory": ("Inventory Tracker", "category"),
    "StockCurrent": ("Inventory Tracker", "current"),
    "StockReserved": ("Inventory Tracker", "reserved"),
    "StockAvailable": ("Inventory Tracker", "available"),
    "StockOnOrder": ("Inventory Tracker", "on_order"),
    "StockAvgCost": ("Inventory Tracker", "avg_cost"),
    "StockValueCol": ("Inventory Tracker", "stock_value"),
    "StockRetailValue": ("Inventory Tracker", "retail_value"),
    "StockDailySales": ("Inventory Tracker", "daily_sales"),
    "StockDaysCover": ("Inventory Tracker", "days_cover"),
    "StockState": ("Inventory Tracker", "stock_state"),
    "StockIsOut": ("Inventory Tracker", "is_out"),
    "StockIsLow": ("Inventory Tracker", "is_low"),
    "StockIsOver": ("Inventory Tracker", "is_over"),
    "StockIsDead": ("Inventory Tracker", "is_dead"),
    "StockAlert": ("Inventory Tracker", "alert"),

    # ---- movements -----------------------------------------------------
    "MovementSKU": ("Stock Movement Log", "sku"),
    "MovementDates": ("Stock Movement Log", "date_num"),
    "MovementMonths": ("Stock Movement Log", "month"),
    "MovementTypeCol": ("Stock Movement Log", "type"),
    "MovementQty": ("Stock Movement Log", "signed_qty"),
    "MovementValue": ("Stock Movement Log", "value"),
    "MovementDamaged": ("Stock Movement Log", "damaged"),
    "MovementAdjustment": ("Stock Movement Log", "adjustment"),
    "MovementIsWriteoff": ("Stock Movement Log", "is_writeoff"),
    "MovementIsManual": ("Stock Movement Log", "is_manual"),
    "MovementAlert": ("Stock Movement Log", "alert"),

    # ---- suppliers -----------------------------------------------------
    "SupplierNames": ("Supplier Database", "name"),
    "SupplierLeadTime": ("Supplier Database", "lead_time"),
    "SupplierActualLead": ("Supplier Database", "actual_lead"),
    "SupplierLeadGap": ("Supplier Database", "lead_gap"),
    "SupplierSpend": ("Supplier Database", "spend"),
    "SupplierOpenValue": ("Supplier Database", "open_value"),
    "SupplierOnTime": ("Supplier Database", "on_time_rate"),
    "SupplierStatusCol": ("Supplier Database", "status"),
    "SupplierRating": ("Supplier Database", "rating"),
    "SupplierAlert": ("Supplier Database", "alert"),

    # ---- purchase orders -----------------------------------------------
    "POId": ("Purchase Orders", "po_id"),
    "POSupplier": ("Purchase Orders", "supplier"),
    "POSKU": ("Purchase Orders", "sku"),
    "PODates": ("Purchase Orders", "order_num"),
    "POReceivedDates": ("Purchase Orders", "received_num"),
    "POMonths": ("Purchase Orders", "month"),
    "POQty": ("Purchase Orders", "qty"),
    "POReceived": ("Purchase Orders", "received"),
    "POReceivedValue": ("Purchase Orders", "received_value"),
    "POTotal": ("Purchase Orders", "total"),
    "POOutstanding": ("Purchase Orders", "outstanding"),
    "POOutstandingValue": ("Purchase Orders", "outstanding_value"),
    "POStatusCol": ("Purchase Orders", "status"),
    "POIsOpen": ("Purchase Orders", "is_open"),
    "POIsLate": ("Purchase Orders", "is_late"),
    "POIsOver": ("Purchase Orders", "is_over"),
    "PODupLine": ("Purchase Orders", "dup_line"),
    "PODaysLate": ("Purchase Orders", "days_late"),
    "POPaid": ("Purchase Orders", "paid"),
    "POExpected": ("Purchase Orders", "expected"),
    "POAlert": ("Purchase Orders", "alert"),

    # ---- wholesale customers -------------------------------------------
    "CustomerNames": ("Wholesale Customer CRM", "business"),
    "CustomerTierCol": ("Wholesale Customer CRM", "tier"),
    "CustomerTermsCol": ("Wholesale Customer CRM", "terms"),
    "CustomerDiscount": ("Wholesale Customer CRM", "discount"),
    "CustomerMinOrder": ("Wholesale Customer CRM", "min_order"),
    "CustomerCreditLimit": ("Wholesale Customer CRM",
                            "credit_limit"),
    "CustomerRevenue": ("Wholesale Customer CRM", "revenue"),
    "CustomerProfit": ("Wholesale Customer CRM", "profit"),
    "CustomerOwing": ("Wholesale Customer CRM", "owing"),
    "CustomerOverdue": ("Wholesale Customer CRM", "overdue"),
    "CustomerOrders": ("Wholesale Customer CRM", "orders"),
    "CustomerDaysSince": ("Wholesale Customer CRM", "days_since"),
    "CustomerStatusCol": ("Wholesale Customer CRM", "status"),
    "CustomerRank": ("Wholesale Customer CRM", "revenue_rank"),
    "CustomerAlert": ("Wholesale Customer CRM", "alert"),

    # ---- pricing --------------------------------------------------------
    "PricingSKU": ("Wholesale Pricing", "sku"),
    "PricingBase": ("Wholesale Pricing", "wholesale"),
    "PricingBest": ("Wholesale Pricing", "best_price"),
    "PricingBestGP": ("Wholesale Pricing", "best_gp"),
    "PricingBelowFloor": ("Wholesale Pricing", "below_floor"),
    "PricingThin": ("Wholesale Pricing", "thin_margin"),
    "PricingIsPriced": ("Wholesale Pricing", "is_priced"),
    "PricingCaseQty": ("Wholesale Pricing", "case_qty"),
    "PricingAlert": ("Wholesale Pricing", "alert"),

    # ---- wholesale orders ------------------------------------------------
    "OrderId": ("Wholesale Orders", "order_id"),
    "OrderCustomer": ("Wholesale Orders", "customer"),
    "OrderSKU": ("Wholesale Orders", "sku"),
    "OrderDates": ("Wholesale Orders", "date_num"),
    "OrderShipDates": ("Wholesale Orders", "ship_num"),
    "OrderMonths": ("Wholesale Orders", "month"),
    "OrderQty": ("Wholesale Orders", "qty"),
    "OrderTotal": ("Wholesale Orders", "line_total"),
    "OrderCost": ("Wholesale Orders", "line_cost"),
    "OrderProfit": ("Wholesale Orders", "line_profit"),
    "OrderStatusCol": ("Wholesale Orders", "status"),
    "OrderIsReserved": ("Wholesale Orders", "is_reserved"),
    "OrderIsShipped": ("Wholesale Orders", "is_shipped"),
    "OrderIsOpen": ("Wholesale Orders", "is_open"),
    "OrderIsOverdue": ("Wholesale Orders", "is_overdue"),
    "OrderFirstLine": ("Wholesale Orders", "is_first_line"),
    "OrderOutstanding": ("Wholesale Orders", "outstanding"),
    "OrderShortfall": ("Wholesale Orders", "shortfall"),
    "OrderAge": ("Wholesale Orders", "age_days"),
    "OrderDaysToShip": ("Wholesale Orders", "days_to_ship"),
    "OrderAlert": ("Wholesale Orders", "alert"),

    # ---- retail sales -----------------------------------------------------
    "SaleId": ("Retail Sales Tracker", "sale_id"),
    "SaleSKU": ("Retail Sales Tracker", "sku"),
    "SaleDates": ("Retail Sales Tracker", "date_num"),
    "SaleMonths": ("Retail Sales Tracker", "month"),
    "SaleWeekdays": ("Retail Sales Tracker", "weekday"),
    "SaleChannel": ("Retail Sales Tracker", "channel"),
    "SaleQty": ("Retail Sales Tracker", "qty"),
    "SaleGross": ("Retail Sales Tracker", "gross"),
    "SaleNet": ("Retail Sales Tracker", "net_revenue"),
    "SaleFees": ("Retail Sales Tracker", "platform_fee"),
    "SalePaymentFees": ("Retail Sales Tracker", "payment_fee"),
    "SaleCOGS": ("Retail Sales Tracker", "cogs"),
    "SaleProfit": ("Retail Sales Tracker", "profit"),
    "SaleIsSold": ("Retail Sales Tracker", "is_sold"),
    "SaleIsLoss": ("Retail Sales Tracker", "is_loss"),
    "SaleShipCharged": ("Retail Sales Tracker", "shipping_charged"),
    "SaleShipCost": ("Retail Sales Tracker", "shipping_cost"),
    "SaleAlert": ("Retail Sales Tracker", "alert"),

    # ---- reorder planning ---------------------------------------------------
    "ReorderSKU": ("Reorder Planner", "sku"),
    "ReorderSupplier": ("Reorder Planner", "supplier"),
    "ReorderQty": ("Reorder Planner", "order_rounded"),
    "ReorderValue": ("Reorder Planner", "order_value"),
    "ReorderUrgency": ("Reorder Planner", "urgency"),
    "ReorderPriority": ("Reorder Planner", "priority"),
    "ReorderNeeded": ("Reorder Planner", "needs_order"),
    "ReorderBy": ("Reorder Planner", "order_by"),
    "ReorderPoint": ("Reorder Planner", "reorder_point"),
    "ReorderAlert": ("Reorder Planner", "alert"),

    # ---- valuation ----------------------------------------------------------
    "ValuationSKU": ("Inventory Valuation & COGS", "sku"),
    "ValuationAvgCost": ("Inventory Valuation & COGS", "avg_cost"),
    "ValuationStockValue": ("Inventory Valuation & COGS",
                            "stock_value"),
    "ValuationCOGS": ("Inventory Valuation & COGS", "cogs"),
    "ValuationRevenue": ("Inventory Valuation & COGS", "revenue"),
    "ValuationProfit": ("Inventory Valuation & COGS",
                        "gross_profit"),
    "ValuationWriteoff": ("Inventory Valuation & COGS",
                          "writeoff_value"),
    "ValuationTurns": ("Inventory Valuation & COGS", "turns"),
    "ValuationABC": ("Inventory Valuation & COGS", "abc_class"),
    "ValuationDead": ("Inventory Valuation & COGS", "dead_value"),
    "ValuationAlert": ("Inventory Valuation & COGS", "alert"),

    # ---- returns -------------------------------------------------------------
    "ReturnSKU": ("Returns & Damaged Goods", "sku"),
    "ReturnDates": ("Returns & Damaged Goods", "date_num"),
    "ReturnMonths": ("Returns & Damaged Goods", "month"),
    "ReturnQty": ("Returns & Damaged Goods", "qty"),
    "ReturnReasonCol": ("Returns & Damaged Goods", "reason"),
    "ReturnRestocked": ("Returns & Damaged Goods", "restocked"),
    "ReturnLoss": ("Returns & Damaged Goods", "net_loss"),
    "ReturnTotalCost": ("Returns & Damaged Goods", "total_cost"),
    "ReturnUnclaimed": ("Returns & Damaged Goods", "unclaimed"),
    "ReturnIsWriteoff": ("Returns & Damaged Goods", "is_writeoff"),
    "ReturnIsOpen": ("Returns & Damaged Goods", "is_open"),
    "ReturnAlert": ("Returns & Damaged Goods", "alert"),

    # ---- locations -----------------------------------------------------------
    "LocationSKU": ("Warehouse Locations", "sku"),
    "LocationCol": ("Warehouse Locations", "location"),
    "LocationQty": ("Warehouse Locations", "qty_here"),
    "LocationGap": ("Warehouse Locations", "placed_gap"),
    "LocationAlert": ("Warehouse Locations", "alert"),

    # ---- audit ----------------------------------------------------------------
    "AuditSKU": ("Stock Audit", "sku"),
    "AuditDates": ("Stock Audit", "date_num"),
    "AuditVariance": ("Stock Audit", "variance"),
    "AuditVarianceValue": ("Stock Audit", "variance_value"),
    "AuditAbsValue": ("Stock Audit", "abs_value"),
    "AuditHasVariance": ("Stock Audit", "has_variance"),
    "AuditUnposted": ("Stock Audit", "unposted"),
    "AuditAccuracy": ("Stock Audit", "accuracy"),
    "AuditAlert": ("Stock Audit", "alert"),

    # ---- performance ------------------------------------------------------------
    "PerfSKU": ("Product Performance", "sku"),
    "PerfProduct": ("Product Performance", "product"),
    "PerfCategory": ("Product Performance", "category"),
    "PerfUnits": ("Product Performance", "units_sold"),
    "PerfRevenue": ("Product Performance", "revenue"),
    "PerfProfit": ("Product Performance", "profit"),
    "PerfGP": ("Product Performance", "gp_pct"),
    "PerfPerUnit": ("Product Performance", "profit_per_unit"),
    "PerfReturnOnStock": ("Product Performance", "return_on_stock"),
    "PerfStockValue": ("Product Performance", "stock_value"),
    "PerfBand": ("Product Performance", "band"),
    "PerfIsDead": ("Product Performance", "is_dead"),
    "PerfIsStar": ("Product Performance", "is_star"),
    "PerfRevenueRank": ("Product Performance", "revenue_rank"),
    "PerfProfitRank": ("Product Performance", "profit_rank"),
    "PerfDaysSince": ("Product Performance", "days_since_sale"),
    "PerfAlert": ("Product Performance", "alert"),

    # ---- expenses -----------------------------------------------------------------
    "ExpenseDates": ("Expense Tracker", "date_num"),
    "ExpenseMonths": ("Expense Tracker", "month"),
    "ExpenseCategoryCol": ("Expense Tracker", "category"),
    "ExpenseAmount": ("Expense Tracker", "amount"),
    "ExpenseOperating": ("Expense Tracker", "operating"),
    "ExpenseInventory": ("Expense Tracker", "inventory_spend"),
    "ExpenseMonthly": ("Expense Tracker", "monthly"),
    "ExpenseReceipt": ("Expense Tracker", "receipt"),
    "ExpenseDuplicate": ("Expense Tracker", "duplicate"),
    "ExpenseAlert": ("Expense Tracker", "alert"),

    # ---- monthly profit --------------------------------------------------------------
    "ProfitMonths": ("Profit Analysis", "month"),
    "ProfitRevenue": ("Profit Analysis", "revenue"),
    "ProfitCOGS": ("Profit Analysis", "cogs"),
    "ProfitGross": ("Profit Analysis", "gross_profit"),
    "ProfitExpenses": ("Profit Analysis", "expenses"),
    "ProfitNet": ("Profit Analysis", "net_profit"),
    "ProfitMargin": ("Profit Analysis", "net_margin"),
    "ProfitStockSpend": ("Profit Analysis", "stock_spend"),
    "ProfitCash": ("Profit Analysis", "cash_effect"),
    "ProfitRetail": ("Profit Analysis", "retail_revenue"),
    "ProfitWholesale": ("Profit Analysis", "wholesale_revenue"),
    "ProfitAlert": ("Profit Analysis", "alert"),
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

### `inventory_wholesale_spreadsheet/sample_data.py`

*1,322 lines · 22 functions*

> A worked example: one small supply business, mid-year, mid-mess.
>
> Thornfield Supply Co. buys and makes homeware and sells it three
> ways - a marketplace shop, a handful of wholesale stockists, and the
> occasional market stall. The year it is having is deliberately not a
> tidy one, because a sample company where everything reconciles teaches
> you nothing about what the alert columns are for.
>
> In this example you will find: a bestseller that has been out of stock
> for nine days while the purchase order sits unsent; a supplier whose
> deliveries are reliably a week late and whose lead time in the
> database still says seven days; a wholesale order confirmed three
> weeks ago and never picked, quietly reserving ninety units; a pallet
> of mugs received against a purchase order where the invoice quantity
> and the delivered quantity disagree; a product being sold at a tier
> price below its own floor; and a stock count that found eleven fewer
> candles than the system believed.
>
> Every row here is checked twice before it ships. The keys must all be
> *input* columns on their register - a value written into a formula
> column is silently dropped and takes the formula with it - and every
> value that feeds a drop-down must exist in that drop-down's list.
> ``tools/check_inventory_sample.py`` does both checks.
>
> Dates are relative to ``cfg.today``, so the example stays current
> whenever the workbook is built.

| Defines | What it does |
|---------|--------------|
| `_day(cfg, offset)` | A date `offset` days from the build date |
| `_sup(n)` |  |
| `_po(n)` |  |
| `_cust(n)` |  |
| `_ord(n)` |  |
| `_loc(n)` |  |
| `_sku_at(index)` |  |
| `_cost_at(index)` |  |
| `products(cfg)` |  |
| `suppliers(cfg)` |  |
| `purchases(cfg)` |  |
| `customers(cfg)` |  |
| `pricing(cfg)` | Mirrors the catalogue - the sheet generates its own SKU column |
| `orders(cfg)` |  |
| `_stock_timeline(cfg)` | Inbound and outbound units per SKU per day, retail aside |
| `retail(cfg)` | Four months of marketplace and market-stall trade |
| `movements(cfg)` |  |
| `returns(cfg)` |  |
| `locations(cfg)` |  |
| `audit(cfg)` |  |
| `expenses(cfg)` |  |
| `get(module, cfg)` | Sample rows for a register module, or an empty list |

````python
"""A worked example: one small supply business, mid-year, mid-mess.

Thornfield Supply Co. buys and makes homeware and sells it three
ways - a marketplace shop, a handful of wholesale stockists, and the
occasional market stall. The year it is having is deliberately not a
tidy one, because a sample company where everything reconciles teaches
you nothing about what the alert columns are for.

In this example you will find: a bestseller that has been out of stock
for nine days while the purchase order sits unsent; a supplier whose
deliveries are reliably a week late and whose lead time in the
database still says seven days; a wholesale order confirmed three
weeks ago and never picked, quietly reserving ninety units; a pallet
of mugs received against a purchase order where the invoice quantity
and the delivered quantity disagree; a product being sold at a tier
price below its own floor; and a stock count that found eleven fewer
candles than the system believed.

Every row here is checked twice before it ships. The keys must all be
*input* columns on their register - a value written into a formula
column is silently dropped and takes the formula with it - and every
value that feeds a drop-down must exist in that drop-down's list.
``tools/check_inventory_sample.py`` does both checks.

Dates are relative to ``cfg.today``, so the example stays current
whenever the workbook is built.
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Callable, Dict, List

from .config import Config, MOVEMENT_TYPES

Rows = List[Dict[str, Any]]


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _day(cfg: Config, offset: int) -> _dt.date:
    """A date `offset` days from the build date."""
    return cfg.today + _dt.timedelta(days=offset)


def _sup(n: int) -> str:
    return f"SUP-{n:03d}"


def _po(n: int) -> str:
    return f"PO-{n:04d}"


def _cust(n: int) -> str:
    return f"WC-{n:03d}"


def _ord(n: int) -> str:
    return f"WO-{n:04d}"


def _loc(n: int) -> str:
    return f"LOC-{n:03d}"


# ---------------------------------------------------------------------------
# the catalogue
# ---------------------------------------------------------------------------
# (sku, name, category, subcategory, variant, size, colour, material,
#  unit, supplier, cost, retail, wholesale, floor, opening_qty,
#  min, max, status)
_CATALOGUE = [
    ("CND-JAR-SAGE", "Sage & Sea Salt Jar Candle", "Candles",
     "Jar candles", "Standard", "220ml", "Sage", "Soy wax", "Each",
     1, 4.80, 18.00, 9.00, 7.50, 180, 60, 400, "Active"),
    ("CND-JAR-FIG", "Black Fig Jar Candle", "Candles", "Jar candles",
     "Standard", "220ml", "Charcoal", "Soy wax", "Each", 1, 4.80,
     18.00, 9.00, 7.50, 42, 60, 400, "Active"),
    ("CND-JAR-LIN", "Fresh Linen Jar Candle", "Candles",
     "Jar candles", "Standard", "220ml", "White", "Soy wax", "Each",
     1, 4.80, 18.00, 9.00, 7.50, 96, 48, 300, "Active"),
    ("CND-PIL-IVY", "Ivory Pillar Candle", "Candles",
     "Pillar candles", "Tall", "20cm", "Ivory", "Soy wax", "Each", 1,
     3.20, 12.00, 6.00, 5.00, 210, 40, 300, "Active"),
    ("CND-TEA-12", "Tealights, Box of 12", "Candles", "Tealights",
     "Boxed", "12 pack", "Natural", "Soy wax", "Box of 24", 1, 2.40,
     9.50, 4.75, 4.00, 64, 30, 240, "Active"),
    ("CND-MLT-CIT", "Citrus Grove Wax Melts", "Candles", "Wax melts",
     "Clamshell", "6 cube", "Yellow", "Soy wax", "Pack of 6", 1,
     1.60, 7.00, 3.50, 2.90, 150, 40, 260, "Active"),

    ("CER-MUG-350", "Stoneware Mug 350ml", "Ceramics", "Mugs",
     "Glazed", "350ml", "Oat", "Stoneware", "Each", 2, 6.10, 24.00,
     12.00, 10.00, 120, 48, 320, "Active"),
    ("CER-MUG-250", "Stoneware Mug 250ml", "Ceramics", "Mugs",
     "Glazed", "250ml", "Oat", "Stoneware", "Each", 2, 5.40, 21.00,
     10.50, 8.80, 88, 36, 240, "Active"),
    ("CER-MUG-SEC", "Stoneware Mug - Seconds", "Ceramics", "Mugs",
     "Seconds", "350ml", "Mixed", "Stoneware", "Each", 2, 6.10,
     12.00, 7.00, 6.50, 34, 0, 60, "Clearance"),
    ("CER-BWL-SML", "Breakfast Bowl, Small", "Ceramics", "Bowls",
     "Glazed", "14cm", "Oat", "Stoneware", "Each", 2, 7.20, 26.00,
     13.00, 11.00, 60, 24, 180, "Active"),
    ("CER-PLT-TER", "Terracotta Planter", "Ceramics", "Planters",
     "Unglazed", "16cm", "Terracotta", "Clay", "Each", 2, 5.90,
     22.00, 11.00, 9.50, 57, 24, 160, "Active"),
    ("CER-VAS-BUD", "Bud Vase", "Ceramics", "Vases", "Glazed",
     "12cm", "Sage", "Stoneware", "Each", 2, 4.40, 17.00, 8.50,
     7.20, 100, 24, 150, "Active"),

    ("TEX-TWL-2PK", "Linen Tea Towels, Pair", "Textiles",
     "Tea towels", "Pair", "50x70cm", "Natural", "Linen", "Pair", 3,
     5.60, 20.00, 10.00, 8.50, 114, 30, 200, "Active"),
    ("TEX-CSH-45", "Cushion Cover 45cm", "Textiles",
     "Cushion covers", "Zipped", "45cm", "Clay", "Cotton", "Each", 3,
     8.30, 32.00, 16.00, 13.50, 36, 18, 120, "Active"),
    ("TEX-APR-NAT", "Cross-back Apron", "Textiles", "Aprons",
     "Adult", "One size", "Natural", "Linen", "Each", 3, 11.40,
     42.00, 21.00, 18.00, 22, 12, 80, "Active"),
    ("TEX-TOT-CAN", "Canvas Tote Bag", "Textiles", "Tote bags",
     "Printed", "38x42cm", "Natural", "Cotton", "Each", 3, 3.90,
     16.00, 8.00, 6.50, 140, 40, 300, "Active"),

    ("STA-NTB-A5", "A5 Lined Notebook", "Stationery", "Notebooks",
     "Softcover", "A5", "Sage", "Paper", "Each", 4, 3.10, 13.00,
     6.50, 5.40, 160, 48, 320, "Active"),
    ("STA-CRD-6PK", "Greeting Cards, Pack of 6", "Stationery",
     "Greeting cards", "Assorted", "A6", "Mixed", "Paper",
     "Pack of 6", 4, 2.70, 11.00, 5.50, 4.60, 95, 36, 240, "Active"),
    ("STA-PRN-A4", "Botanical Print Set A4", "Stationery",
     "Print sets", "Set of 3", "A4", "Mixed", "Paper", "Pack of 4",
     4, 4.20, 18.00, 9.00, 7.60, 93, 18, 120, "Active"),
    ("STA-STK-SHT", "Sticker Sheet", "Stationery", "Stickers",
     "Single", "A6", "Mixed", "Vinyl", "Each", 4, 0.70, 4.00, 2.00,
     1.60, 320, 100, 600, "Active"),

    ("BTH-SOP-OAT", "Oatmeal Soap Bar", "Bath & Body", "Soap bars",
     "Wrapped", "100g", "Cream", "Soap", "Each", 5, 2.30, 9.00,
     4.50, 3.80, 110, 40, 260, "Active"),
    ("BTH-SLT-LAV", "Lavender Bath Salts", "Bath & Body",
     "Bath salts", "Pouch", "500g", "Purple", "Mineral salt",
     "Each", 5, 3.40, 14.00, 7.00, 5.90, 65, 24, 180, "Active"),
    ("BTH-BLM-LIP", "Lip Balm Trio", "Bath & Body", "Balms",
     "Boxed", "3x10ml", "Mixed", "Beeswax", "Pack of 2", 5, 4.10,
     16.00, 8.00, 6.80, 12, 24, 140, "Active"),

    ("PKG-BOX-SML", "Kraft Gift Box, Small", "Packaging", "Boxes",
     "Flatpack", "12cm", "Kraft", "Card", "Pack of 10", 6, 0.42,
     0.00, 0.00, 0.00, 480, 200, 1200, "Active"),
    ("PKG-TIS-NAT", "Tissue Paper, Natural", "Packaging", "Tissue",
     "Sheet", "50x75cm", "Natural", "Paper", "Pack of 10", 6, 0.09,
     0.00, 0.00, 0.00, 1400, 500, 4000, "Active"),
    ("PKG-LBL-RND", "Round Branded Labels", "Packaging", "Labels",
     "Roll", "40mm", "Sage", "Paper", "Pack of 10", 6, 0.06, 0.00,
     0.00, 0.00, 220, 300, 2000, "Active"),
    ("PKG-MLR-MED", "Padded Mailer, Medium", "Packaging", "Mailers",
     "Padded", "C4", "Kraft", "Paper", "Pack of 10", 6, 0.38, 0.00,
     0.00, 0.00, 36, 150, 900, "Active"),

    ("CND-JAR-ROS", "Wild Rose Jar Candle", "Candles", "Jar candles",
     "Discontinued", "220ml", "Pink", "Soy wax", "Each", 1, 4.80,
     18.00, 9.00, 7.50, 9, 0, 0, "Discontinued"),
    ("TEX-SCF-WOL", "Lambswool Scarf", "Textiles", "Cushion covers",
     "Winter", "180cm", "Charcoal", "Wool", "Each", 3, 14.20, 52.00,
     26.00, 22.00, 18, 10, 60, "Seasonal"),
]


def _sku_at(index: int) -> str:
    return _CATALOGUE[index][0]


def _cost_at(index: int) -> float:
    return _CATALOGUE[index][10]


# ---------------------------------------------------------------------------
# 4. Product & SKU Database
# ---------------------------------------------------------------------------

_DESCRIPTIONS = {
    "CND-JAR-SAGE": "Hand poured, cotton wick, 45 hour burn. The one "
                    "everybody reorders.",
    "CER-MUG-350": "Thrown and glazed in house. Dishwasher safe. "
                   "Sold singly and in fours.",
    "CER-MUG-SEC": "Glaze pooling or a soft rim. Perfectly usable, "
                   "sold cheap, never to stockists.",
    "PKG-MLR-MED": "Despatch packaging. Not for sale - kept here so "
                   "its cost lands in the right place.",
    "CND-JAR-ROS": "Withdrawn after the fragrance house changed the "
                   "oil. Sell the remaining nine and let it go.",
}

_LAUNCH_OFFSETS = [-820, -640, -512, -790, -455, -288, -735, -700,
                   -410, -512, -366, -120, -690, -520, -300, -610,
                   -580, -540, -330, -250, -660, -470, -140, -900,
                   -900, -900, -900, -760, -210]


def products(cfg: Config) -> Rows:
    rows: Rows = []
    for i, item in enumerate(_CATALOGUE):
        (sku, name, cat, sub, variant, size, colour, material, unit,
         sup, cost, retail, wholesale, floor, opening, lo, hi,
         status) = item
        rows.append({
            "sku": sku,
            "name": name,
            "description": _DESCRIPTIONS.get(sku, ""),
            "category": cat,
            "subcategory": sub,
            "brand": "Thornfield",
            "variant": variant,
            "size": size,
            "colour": colour,
            "material": material,
            "barcode": f"50{12000 + i * 37:010d}"[:13],
            "unit": unit,
            "image": "",
            "supplier": _sup(sup),
            "supplier_sku": f"T-{sku.split('-')[-1]}{i:02d}",
            "cost": cost,
            "retail_price": retail,
            "wholesale_price": wholesale,
            "min_price": floor,
            "opening_qty": opening,
            "opening_value": round(opening * cost, 2),
            "min_stock": lo,
            "max_stock": hi,
            "reorder_point": "",
            "status": status,
            "launched": _day(cfg, _LAUNCH_OFFSETS[i]),
            "notes": "",
        })

    # the deliberate duplicate-SKU teaching moment lives in the
    # Stock Audit sheet instead - a duplicate here would corrupt
    # every other tab's sample numbers, which is exactly the point
    # the alert column is making.
    return rows


# ---------------------------------------------------------------------------
# 7. Supplier Database
# ---------------------------------------------------------------------------

#: (name, contact, type, city, country, terms, lead, rating, status,
#:  preferred, supplies, notes)
_SUPPLIERS = [
    ("Kestrel Wax & Wick", "Dan Mehta", "Wholesaler", "Portland",
     "United States", "Net 30", 7, 4, "Preferred", "Yes",
     "Soy wax, wicks, jars, lids, fragrance oils",
     "Reliable on the wax, slow on anything they have to order in. "
     "Lead time says seven days and the scorecard disagrees."),
    ("Ridgeway Ceramics", "Ana Sol", "Manufacturer", "Stoke",
     "United Kingdom", "Net 45", 35, 5, "Preferred", "Yes",
     "Stoneware blanks, glazes, firing",
     "Excellent quality, long lead time, worth planning around. "
     "Ship by sea, so never order in a hurry."),
    ("Fenwick Textiles", "Marta Oyelaran", "Importer", "Leicester",
     "United Kingdom", "Net 30", 21, 3, "Active", "No",
     "Linen, cotton canvas, finished tea towels",
     "Price went up twice this year without warning."),
    ("Marsh & Co Paper", "Theo Lindqvist", "Printer", "Seattle",
     "United States", "Net 14", 10, 4, "Active", "Yes",
     "Notebooks, cards, print sets, stickers", ""),
    ("Lavender Lane Botanics", "Priya Raman", "Maker / artisan",
     "Eugene", "United States", "Prepaid", 14, 4, "Active", "No",
     "Soap bases, bath salts, balms, essential oils", ""),
    ("Northgate Packaging", "Sam Okafor", "Packaging supplier",
     "Tacoma", "United States", "Net 30", 5, 3, "Active", "Yes",
     "Boxes, tissue, labels, mailers, tape",
     "Cheap and quick. Quality is adequate, not lovely."),
    ("Harbour Glass Supply", "Lena Vogt", "Distributor", "Bristol",
     "United Kingdom", "Net 30", 28, 2, "On hold", "No",
     "Glass jars and bottles",
     "Two short deliveries and a damaged pallet. On hold until "
     "they explain themselves."),
    ("Beacon Trade Goods", "Owen Pryce", "Wholesaler", "Denver",
     "United States", "Net 30", 12, 3, "Trial", "No",
     "Mixed homeware for resale", "First order placed, nothing "
     "received yet."),
]


def suppliers(cfg: Config) -> Rows:
    rows: Rows = []
    for i, (name, contact, kind, city, country, terms, lead, rating,
            status, preferred, supplies, note) in enumerate(
                _SUPPLIERS, start=1):
        handle = name.lower().replace(" & ", "").replace(" ", "")[:16]
        rows.append({
            "name": name,
            "contact": contact,
            "email": f"{contact.split()[0].lower()}@{handle}.example",
            "phone": f"(503) 555-0{100 + i * 7:03d}",
            "website": f"{handle}.example",
            "address": f"{i * 13} Trade Street",
            "city": city,
            "country": country,
            "type": kind,
            "supplies": supplies,
            "terms": terms,
            "currency": "GBP" if country == "United Kingdom"
                        else "USD",
            "lead_time": lead,
            "moq": [250, 500, 300, 150, 120, 200, 400, 250][i - 1],
            "free_ship_over": [500, 1500, 800, 400, 0, 350, 0,
                               600][i - 1],
            "discount": [0.05, 0.0, 0.0, 0.03, 0.0, 0.08, 0.0,
                         0.0][i - 1],
            "account_no": f"TH-{4100 + i * 17}",
            "tax_id": "",
            "status": status,
            "preferred": preferred,
            "rating": rating,
            "first_order": _day(cfg, -[700, 640, 520, 480, 400, 610,
                                       300, 24][i - 1]),
            "notes": note,
        })
    return rows


# ---------------------------------------------------------------------------
# 8. Purchase Orders
# ---------------------------------------------------------------------------
# Deliberate teaching cases in here:
#   PO-0007 - sent weeks ago, nothing received, and it is the SKU
#             that is out of stock. The chase list exists for this.
#   PO-0009 - part delivery: 30 of 50 arrived, 20 still outstanding.
#   PO-0011 - invoice disagrees with the order value.
#   PO-0012 - received more than was ordered.
#   PO-0013 - still a draft, dated, never sent.

#: (po, supplier_n, days_ago, lead, [(sku_index, qty, cost, freight,
#:  received)], status, condition, invoice_gap, paid)
_PO_PLAN = [
    ("PO-0001", 1, -96, 7,
     [(0, 240, 4.60, 38.00, 240), (1, 120, 4.60, 19.00, 120),
      (2, 120, 4.60, 19.00, 120)],
     "Fully Received", "Good", 0.0, "Yes"),
    ("PO-0002", 2, -88, 35,
     [(6, 144, 5.90, 120.00, 144), (7, 96, 5.20, 80.00, 96),
      (9, 72, 7.00, 60.00, 72)],
     "Fully Received", "Good", 0.0, "Yes"),
    ("PO-0003", 4, -74, 10,
     [(16, 200, 3.00, 24.00, 200), (17, 120, 2.60, 15.00, 120),
      (19, 400, 0.66, 18.00, 400)],
     "Fully Received", "Good", 0.0, "Yes"),
    ("PO-0004", 6, -66, 5,
     [(23, 600, 0.40, 22.00, 600), (24, 1600, 0.085, 14.00, 1600),
      (26, 300, 0.36, 12.00, 300)],
     "Fully Received", "Good", 0.0, "Yes"),
    ("PO-0005", 3, -58, 21,
     [(12, 96, 5.40, 72.00, 96), (13, 48, 8.00, 44.00, 48),
      (15, 180, 3.75, 54.00, 180)],
     "Fully Received", "Minor damage", 0.0, "Yes"),
    ("PO-0006", 5, -44, 14,
     [(20, 144, 2.20, 26.00, 144), (21, 72, 3.30, 18.00, 72)],
     "Fully Received", "Good", 0.0, "Yes"),
    ("PO-0007", 1, -31, 7,
     [(1, 180, 4.80, 30.00, 0), (11, 96, 4.40, 22.00, 0)],
     "Sent", "", 0.0, ""),
    ("PO-0008", 2, -26, 35,
     [(6, 120, 6.10, 110.00, 0), (10, 60, 5.90, 55.00, 0)],
     "Confirmed", "", 0.0, ""),
    ("PO-0009", 4, -22, 10,
     [(18, 50, 4.10, 16.00, 30), (16, 150, 3.10, 20.00, 150)],
     "Partially Received", "Short delivery", 0.0, "No"),
    ("PO-0010", 6, -16, 5,
     [(27, 500, 0.37, 18.00, 500), (25, 800, 0.058, 9.00, 800)],
     "Fully Received", "Good", 0.0, "No"),
    ("PO-0011", 3, -14, 21,
     [(14, 36, 11.90, 40.00, 36), (12, 72, 5.90, 48.00, 72)],
     "Fully Received", "Good", 86.40, "No"),
    ("PO-0012", 5, -11, 14,
     [(22, 60, 4.10, 16.00, 72)],
     "Fully Received", "Over delivery", 0.0, "No"),
    ("PO-0013", 8, -6, 12,
     [(3, 150, 3.10, 28.00, 0), (4, 90, 2.35, 14.00, 0)],
     "Draft", "", 0.0, ""),
    ("PO-0014", 1, -4, 7,
     [(5, 180, 1.55, 20.00, 0)],
     "Confirmed", "", 0.0, ""),
    ("PO-0015", 6, -2, 5,
     [(23, 400, 0.41, 18.00, 0)],
     "Sent", "", 0.0, ""),
]


def purchases(cfg: Config) -> Rows:
    rows: Rows = []
    checkers = ["Owner", "Warehouse", "Packer", "Assistant"]
    for index, (po, sup, ago, lead, lines, status, condition,
                gap, paid) in enumerate(_PO_PLAN):
        ordered = _day(cfg, ago)
        # the lead time that actually happened, not the promised one
        slip = {1: 6, 2: 3, 3: 9, 4: 1, 5: 2, 6: 0, 7: 0,
                8: 0}.get(sup, 0)
        expected = _day(cfg, ago + lead)
        received_on = _day(cfg, ago + lead + slip)
        got_anything = any(line[4] > 0 for line in lines)
        for pos, (sku_i, qty, cost, freight, received) in enumerate(
                lines):
            row = {
                "po_id": po,
                "supplier": _sup(sup),
                "order_date": ordered if status != "Draft" else "",
                "expected": expected,
                "sku": _sku_at(sku_i),
                "qty": qty,
                "unit_cost": cost,
                "discount": 0.0,
                "freight": freight,
                "received": received,
                "status": status,
                "notes": "",
            }
            if received > 0:
                row["received_date"] = received_on
                row["received_by"] = checkers[(index + pos) % 4]
                row["condition"] = condition or "Good"
            if got_anything and pos == 0:
                row["invoice_no"] = f"INV-{9100 + index * 13}"
                total = sum(q * c + f
                            for _s, q, c, f, _r in lines)
                row["invoice_amount"] = round(total + gap, 2)
                row["paid"] = paid or "No"
            if status == "Draft":
                row["notes"] = ("Drafted after the last stock count "
                                "and never sent. This is what the "
                                "'Send -' alert is for.")
            if po == "PO-0009" and pos == 0:
                row["notes"] = ("Thirty of fifty arrived. The line "
                                "stays open and twenty are still "
                                "counted as on order.")
            if po == "PO-0012":
                row["notes"] = ("Twelve more arrived than were "
                                "ordered. Check before you pay for "
                                "them.")
            if po == "PO-0011" and pos == 0:
                row["notes"] = ("Invoice is higher than the order. "
                                "Price went up and nobody told us.")
            rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 9. Wholesale Customer CRM
# ---------------------------------------------------------------------------

#: (business, buyer, type, city, tier, terms, limit, min, status,
#:  rating, source, extra discount, notes)
_CUSTOMERS = [
    ("Marlowe & Hart", "Freya Marlowe", "Independent boutique",
     "Portland", "Wholesale", "Net 30", 4000, 250, "Active", 5,
     "Trade show", 0.0,
     "Best account by a distance. Reorders candles monthly without "
     "being asked."),
    ("The Fern Room", "Jonah Beck", "Gift shop", "Seattle",
     "Wholesale", "Net 30", 2500, 200, "Active", 4, "Instagram",
     0.0, ""),
    ("Alder & Oak Living", "Nadia Qureshi", "Garden centre",
     "Eugene", "Trade", "Net 14", 1500, 150, "Active", 4,
     "Referral", 0.0, "Seasonal - heavy in spring, quiet in "
     "winter."),
    ("Coastline Home", "Greg Alvarez", "Online retailer",
     "San Diego", "Distributor", "Net 45", 8000, 600, "Active", 3,
     "Faire or marketplace", 0.02,
     "Biggest volume, thinnest margin. Watch the GP% column before "
     "agreeing to anything else."),
    ("Bramble Cafe", "Ivy Donnelly", "Cafe or restaurant",
     "Portland", "Trade", "Net 14", 800, 100, "Active", 4,
     "Walk-in", 0.0, "Buys mugs and tea towels only."),
    ("Pinewood Museum Shop", "Hal Osei", "Museum shop", "Tacoma",
     "Wholesale", "Net 30", 2000, 200, "Occasional", 4,
     "Cold outreach", 0.0, ""),
    ("Sundry Goods Co.", "Remi Takahashi", "Marketplace seller",
     "Boise", "Trade", "Prepaid", 0, 150, "Active", 3,
     "Website enquiry", 0.0, "Prepay only - they were late twice "
     "and we moved them."),
    ("Harbourview Hotel", "Claire Dunne", "Hotel or B&B",
     "Astoria", "Wholesale", "Net 30", 3000, 300, "Active", 4,
     "Referral", 0.05,
     "Overdue. Lovely people, terrible at paying on time."),
    ("Juniper Box", "Teo Ferreira", "Subscription box", "Denver",
     "Distributor", "Net 30", 6000, 500, "Lapsed", 3,
     "Trade show", 0.0,
     "Ordered twice last year and went quiet. Worth a call."),
    ("Greenhill Interiors", "Suki Lam", "Independent boutique",
     "Vancouver", "Wholesale", "Net 30", 2500, 200, "Prospect", 3,
     "Trade show", 0.0, "Samples sent, no order yet."),
    ("Westbrook Stores", "Martin Ayers", "Department store",
     "Chicago", "Key account", "Net 60", 12000, 1500, "On hold", 2,
     "Sales agent", 0.0,
     "On hold. Wanted 38% of retail and sixty day terms, which "
     "does not work at our costs."),
]


def customers(cfg: Config) -> Rows:
    rows: Rows = []
    first_order_ago = [-430, -380, -300, -250, -210, -180, -150,
                       -340, -400, 0, -120]
    for i, (business, buyer, kind, city, tier, terms, limit, minimum,
            status, rating, source, extra, note) in enumerate(
                _CUSTOMERS):
        handle = business.lower().replace(" & ", "").replace(
            " ", "").replace(".", "")[:18]
        row = {
            "business": business,
            "contact": buyer,
            "email": f"{buyer.split()[0].lower()}@{handle}.example",
            "phone": f"(206) 555-0{200 + i * 9:03d}",
            "website": f"{handle}.example",
            "address": f"{12 + i * 7} High Street",
            "city": city,
            "region": "Oregon" if city in ("Portland", "Eugene",
                                           "Astoria") else "",
            "postcode": f"9{7000 + i * 31:04d}",
            "country": "Canada" if city == "Vancouver"
                       else "United States",
            "type": kind,
            "tier": tier,
            "terms": terms,
            "credit_limit": limit,
            "min_order": minimum,
            "discount": extra,
            "tax_exempt": "Yes" if tier in ("Wholesale",
                                            "Distributor",
                                            "Key account") else "No",
            "resale_no": f"RS-{20450 + i * 13}" if tier != "Trade"
                         else "",
            "shipping_pref": "Courier" if i % 3 else "Pallet",
            "status": status,
            "rating": rating,
            "source": source,
            "notes": note,
        }
        if first_order_ago[i]:
            row["first_order"] = _day(cfg, first_order_ago[i])
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 10. Wholesale Pricing
# ---------------------------------------------------------------------------


def pricing(cfg: Config) -> Rows:
    """Mirrors the catalogue - the sheet generates its own SKU column.

    Two teaching cases: the seconds mugs are priced below their own
    floor at the deepest break, and the lambswool scarf has not been
    reviewed since the wool price moved.
    """
    rows: Rows = []
    for i, item in enumerate(_CATALOGUE):
        sku = item[0]
        wholesale = item[12]
        floor = item[13]
        trade = "No" if sku.startswith("PKG-") else "Yes"
        if wholesale <= 0:
            rows.append({
                "wholesale": "", "tier1_qty": "", "tier1_price": "",
                "tier2_qty": "", "tier2_price": "", "tier3_qty": "",
                "tier3_price": "", "moq": "", "case_qty": 10,
                "available_trade": "No",
                "notes": "Packaging - not sold to trade.",
            })
            continue
        t1 = round(wholesale * 0.94, 2)
        t2 = round(wholesale * 0.88, 2)
        t3 = round(wholesale * 0.82, 2)
        note = ""
        if sku == "CER-MUG-SEC":
            t3 = round(floor - 0.60, 2)
            note = ("The hundred-plus price slipped below the floor "
                    "price when the glaze cost went up. This is what "
                    "the red alert is for.")
        if sku == "TEX-SCF-WOL":
            note = ("Wool went up 20% in the spring and this price "
                    "has not been looked at since.")
        rows.append({
            "wholesale": wholesale,
            "tier1_qty": 10, "tier1_price": t1,
            "tier2_qty": 50, "tier2_price": t2,
            "tier3_qty": 100, "tier3_price": t3,
            "moq": 6 if i % 3 else 12,
            "case_qty": [6, 12, 4, 24][i % 4],
            "available_trade": trade,
            "reviewed": _day(cfg, -(40 + (i * 23) % 300))
                        if sku != "TEX-SCF-WOL"
                        else _day(cfg, -430),
            "notes": note,
        })
    return rows


# ---------------------------------------------------------------------------
# 11. Wholesale Orders
# ---------------------------------------------------------------------------
# Teaching cases:
#   WO-0009 - confirmed three weeks ago, never picked, ninety units
#             reserved the whole time.
#   WO-0012 - confirmed for more units than are available.
#   WO-0006 - shipped, invoiced, overdue, and the customer is the
#             hotel with the reputation for paying late.

#: (order, customer_n, days_ago, [(sku_index, qty)], status,
#:  ship_offset, paid_share, discount)
_ORDER_PLAN = [
    ("WO-0001", 1, -78, [(0, 48), (2, 36), (6, 24)], "Delivered",
     3, 1.0, 0.0),
    ("WO-0002", 2, -70, [(16, 60), (17, 48), (19, 120)], "Delivered",
     2, 1.0, 0.0),
    ("WO-0003", 4, -63, [(0, 120), (3, 96), (12, 72)], "Delivered",
     4, 1.0, 0.0),
    ("WO-0004", 3, -54, [(10, 24), (11, 18), (9, 24)], "Delivered",
     3, 1.0, 0.0),
    ("WO-0005", 5, -45, [(6, 36), (12, 24)], "Delivered", 2, 1.0,
     0.0),
    ("WO-0006", 8, -38, [(0, 60), (20, 48), (21, 36)], "Delivered",
     3, 0.0, 0.0),
    ("WO-0007", 1, -29, [(0, 72), (2, 48), (5, 60)], "Delivered",
     2, 1.0, 0.0),
    ("WO-0008", 6, -24, [(16, 48), (18, 36)], "Shipped", 4, 0.5,
     0.0),
    ("WO-0009", 4, -21, [(0, 60), (6, 30)], "Confirmed", None, 0.0,
     0.0),
    ("WO-0010", 2, -12, [(13, 18), (15, 36)], "Packed", None, 0.0,
     0.0),
    # the order that cleared the bud vase shelf - and the reason
    # PO-0007 sitting unsent is the most expensive row in the book
    ("WO-0011", 4, -12, [(11, 52), (16, 72)], "Delivered", 3, 1.0,
     0.0),
    ("WO-0012", 7, -8, [(19, 96), (17, 36)], "Shipped", 2, 1.0,
     0.0),
    ("WO-0013", 1, -5, [(1, 90), (11, 48)], "Confirmed", None, 0.0,
     0.0),
    ("WO-0014", 5, -3, [(6, 24), (7, 24)], "Processing", None, 0.0,
     0.0),
    ("WO-0015", 3, -2, [(10, 30), (22, 24)], "New", None, 0.0, 0.0),
    ("WO-0016", 2, -1, [(0, 36), (16, 48), (23, 200)], "Confirmed",
     None, 0.0, 0.05),
]


def orders(cfg: Config) -> Rows:
    rows: Rows = []
    methods = ["Courier", "Pallet", "Tracked post", "Own delivery"]
    for index, (order, cust, ago, lines, status, ship_off, paid,
                discount) in enumerate(_ORDER_PLAN):
        placed = _day(cfg, ago)
        shipped = _day(cfg, ago + ship_off) if ship_off is not None \
            else ""
        invoiced = status in ("Shipped", "Delivered")
        for pos, (sku_i, qty) in enumerate(lines):
            item = _CATALOGUE[sku_i]
            row = {
                "order_id": order,
                "date": placed,
                "customer": _cust(cust),
                "po_ref": f"PO{4400 + index * 7}",
                "sku": item[0],
                "qty": qty,
                "line_discount": discount if pos == 0 else 0.0,
                "status": status,
                "notes": "",
            }
            if shipped:
                row["ship_date"] = shipped
                row["ship_method"] = methods[index % 4]
                row["tracking"] = f"TRK{7200100 + index * 431}"
            if invoiced:
                row["invoice_no"] = f"INV-{5200 + index * 11}"
                if pos == 0:
                    row["shipping_charge"] = 0.0 if qty > 80 else 14.50
                if paid > 0:
                    unit = round(item[12] * 0.9, 2)
                    row["paid_amount"] = round(qty * unit * paid, 2)
            if order == "WO-0009":
                row["notes"] = ("Confirmed three weeks ago and never "
                                "picked. Those units have been "
                                "reserved - unavailable to anyone "
                                "else - the whole time.")
            if order == "WO-0013" and pos == 0:
                row["notes"] = ("More than we have. The Short by "
                                "column says exactly how many we "
                                "cannot pick.")
            if order == "WO-0006" and pos == 0:
                row["notes"] = "Invoiced, delivered, still not paid."
            rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 12. Retail Sales Tracker
# ---------------------------------------------------------------------------


def _stock_timeline(cfg: Config) -> Dict[str, Dict[int, int]]:
    """Inbound and outbound units per SKU per day, retail aside.

    The retail generator below has to know what was actually on the
    shelf on the day it wants to sell something. Without this it
    cheerfully sells forty vases that were never bought, the stock
    tracker shows a negative balance, and the first thing a buyer
    sees in the example data is a number that cannot happen.
    """
    moves: Dict[str, Dict[int, int]] = {}

    def add(sku: str, day: int, qty: int) -> None:
        moves.setdefault(sku, {})
        moves[sku][day] = moves[sku].get(day, 0) + qty

    # opening stock, treated as present from before the window
    for item in _CATALOGUE:
        add(item[0], -9999, item[14])
    # deliveries land on the day they were received
    for index, (_po, _sup, ago, lead, lines, status, _cond, _gap,
                _paid) in enumerate(_PO_PLAN):
        slip = {1: 6, 2: 3, 3: 9, 4: 1, 5: 2, 6: 0, 7: 0,
                8: 0}.get(_sup, 0)
        for sku_i, _qty, _cost, _freight, received in lines:
            if received > 0:
                add(_sku_at(sku_i), ago + lead + slip, received)
    # wholesale orders take stock out when they ship, and reserve it
    # from the day they are confirmed
    for _order, _cust, ago, lines, status, ship_off, _paid, _disc in \
            _ORDER_PLAN:
        for sku_i, qty in lines:
            if status in ("Shipped", "Delivered"):
                add(_sku_at(sku_i), ago + (ship_off or 0), -qty)
            elif status in ("Confirmed", "Processing", "Picking",
                            "Packed"):
                add(_sku_at(sku_i), ago, -qty)      # reserved
    # damage, samples, production use, transfers and corrections
    signs = {name: sign for name, sign, _blurb in MOVEMENT_TYPES}
    for ago, sku_i, kind, qty, *_rest in _MOVEMENTS:
        sign = signs.get(kind, 0)
        if kind == "Adjustment":
            add(_sku_at(sku_i), ago, qty)
        else:
            add(_sku_at(sku_i), ago, sign * qty)
    # returns that went back on the shelf
    for (ago, _src, _ref, _who, sku_i, qty, _reason, condition,
         *_rest) in _RETURNS:
        if condition in ("As new", "Opened - resellable"):
            add(_sku_at(sku_i), ago, qty)
    return moves


def retail(cfg: Config) -> Rows:
    """Four months of marketplace and market-stall trade.

    Generated rather than hand written, because the point of this
    sheet is volume: ninety-day averages, channel splits and run
    rates only mean something when there is a run of real-looking
    days behind them.

    The generator will not sell stock that was not there on the day.
    It walks the window forwards carrying a running balance per SKU,
    and skips a sale it cannot cover - which is what produces the
    nine-day stockout on the bud vase rather than a negative number.
    """
    rows: Rows = []
    channels = ["Etsy", "Etsy", "Etsy", "Shopify", "Shopify",
                "Amazon", "eBay", "Physical Store", "Craft Fair",
                "Direct Sales"]
    sellable = [i for i, item in enumerate(_CATALOGUE)
                if item[12] > 0 and item[17] in ("Active",
                                                 "Clearance",
                                                 "Seasonal")]
    timeline = _stock_timeline(cfg)
    balance = {item[0]: 0 for item in _CATALOGUE}
    for sku, days in timeline.items():
        balance[sku] += sum(q for day, q in days.items() if day < -118)

    seed = 7
    counter = 0
    for day_offset in range(-118, 1):
        for sku, days in timeline.items():
            balance[sku] += days.get(day_offset, 0)
        seed = (seed * 1103515245 + 12345) % 2147483648
        # busier at weekends, quieter in the middle of the week
        weekday = (cfg.today + _dt.timedelta(days=day_offset)).weekday()
        base = 4 if weekday >= 5 else 2
        count = base + seed % 3
        for _ in range(count):
            seed = (seed * 1103515245 + 12345) % 2147483648
            sku_i = sellable[seed % len(sellable)]
            item = _CATALOGUE[sku_i]
            seed = (seed * 1103515245 + 12345) % 2147483648
            channel = channels[seed % len(channels)]
            qty = 1 + (seed // 7) % 3
            if balance.get(item[0], 0) < qty:
                continue            # it was not on the shelf that day
            balance[item[0]] -= qty
            counter += 1
            retail_price = item[11]
            discounted = (counter % 23 == 0)
            price = (round(retail_price * 0.75, 2) if discounted
                     else retail_price)
            heavy = item[8] in ("Box of 24", "Case of 48")
            ship_charged = 0.0 if channel in ("Physical Store",
                                              "Craft Fair",
                                              "Direct Sales") \
                else (0.0 if qty >= 3 else 4.95)
            ship_cost = 0.0 if channel in ("Physical Store",
                                           "Craft Fair") \
                else round((7.80 if heavy else 4.20) +
                           (1.60 if qty > 2 else 0.0), 2)
            row = {
                "sale_id": f"SO-{1000 + counter:04d}",
                "date": _day(cfg, day_offset),
                "channel": channel,
                "sku": item[0],
                "qty": qty,
                "unit_price": price,
                "shipping_charged": ship_charged,
                "shipping_cost": ship_cost,
                "discount": 0.0,
                "status": ("Delivered" if day_offset < -6
                           else "Shipped" if day_offset < -1
                           else "Paid"),
                "customer": f"buyer{(counter * 37) % 900:03d}",
                "location": "Main store",
                "notes": "",
            }
            if channel in ("Physical Store", "Craft Fair"):
                row["status"] = "Collected"
                row["location"] = ("Shop floor"
                                   if channel == "Physical Store"
                                   else "Market stall")
            if counter % 47 == 0:
                row["status"] = "Refunded"
                balance[item[0]] += qty      # it came back
            if counter % 31 == 0:
                row["fee_override"] = round(
                    price * qty * 0.11 + 0.25, 2)
                row["notes"] = ("Real fee from the platform "
                                "statement, which beats the "
                                "estimate.")
            rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 6. Stock Movement Log
# ---------------------------------------------------------------------------
# Only the things that are NOT purchases, sales or returns: damage,
# loss, samples, production use, transfers and corrections.

_MOVEMENTS = [
    (-104, 6, "Damaged", 4, "Box crushed in transit from Ridgeway. "
     "Claimed and credited.", "PO-0002", "Main store", "Warehouse"),
    (-96, 0, "Sample or gift", 6, "Sent to three prospective "
     "stockists after the Portland trade show.", "", "Main store",
     "Owner"),
    (-88, 23, "Used in production", 120, "Gift boxes used packing "
     "the autumn wholesale orders.", "", "Workshop", "Packer"),
    (-81, 12, "Damaged", 3, "Water damage on the bottom shelf after "
     "the gutter overflowed.", "", "Overflow unit", "Warehouse"),
    (-74, 16, "Found in count", 8, "Found a sealed box behind the "
     "printer stock during the quarterly count.", "AUD-0002",
     "Main store", "Warehouse"),
    (-66, 3, "Transfer out", 40, "Moved to the market stall for the "
     "harvest fair.", "", "Main store", "Assistant"),
    (-66, 3, "Transfer in", 40, "Arrived at the market stall.", "",
     "Market stall", "Assistant"),
    (-61, 19, "Damaged", 22, "Sticker sheets curled in the damp. "
     "Unsellable.", "", "Workshop", "Owner"),
    (-54, 24, "Used in production", 400, "Tissue paper used across "
     "the September despatches.", "", "Workshop", "Packer"),
    (-48, 8, "Adjustment", -5, "Count found five fewer seconds mugs "
     "than the system believed. Most likely sold at the market "
     "stall without being logged.", "AUD-0004", "Main store",
     "Owner"),
    (-41, 20, "Sample or gift", 12, "Soap bars into the gift bags "
     "for the makers' market.", "", "Main store", "Assistant"),
    (-36, 9, "Damaged", 2, "Dropped while restocking the shop "
     "floor.", "", "Shop floor", "Packer"),
    (-29, 26, "Used in production", 180, "Labels used on the "
     "autumn candle run.", "", "Workshop", "Packer"),
    (-24, 0, "Damaged", 3, "Three jars cracked - the wax was poured "
     "too hot.", "", "Workshop", "Owner"),
    (-19, 13, "Lost or stolen", 2, "Two cushion covers missing from "
     "the market stall. No paperwork, no explanation.", "",
     "Market stall", "Assistant"),
    (-16, 2, "Supplier return", 12, "Twelve linen candles returned "
     "to Kestrel - the fragrance was wrong.", "SUP-001",
     "Main store", "Owner"),
    (-12, 7, "Adjustment", -11, "Stocktake found eleven fewer 250ml "
     "mugs. Reason not established. Posted so the system and the "
     "shelf agree again.", "AUD-0007", "Main store", "Owner"),
    (-9, 21, "Damaged", 4, "Bath salt pouches split on the shelf.",
     "", "Main store", "Warehouse"),
    (-6, 18, "Sample or gift", 8, "Print sets sent to a magazine "
     "for a feature.", "", "Main store", "Owner"),
    (-4, 23, "Used in production", 90, "Gift boxes for the "
     "Marlowe & Hart order.", "WO-0015", "Workshop", "Packer"),
    (-2, 4, "Found in count", 6, "Six boxes of tealights found in "
     "the overflow unit.", "AUD-0009", "Overflow unit",
     "Warehouse"),
    (-1, 10, "Damaged", 1, "One planter chipped unloading the "
     "pallet.", "", "Main store", "Warehouse"),
]


def movements(cfg: Config) -> Rows:
    rows: Rows = []
    for ago, sku_i, kind, qty, reason, ref, location, who in \
            _MOVEMENTS:
        row = {
            "date": _day(cfg, ago),
            "sku": _sku_at(sku_i),
            "type": kind,
            "qty": qty,
            "reason": reason,
            "reference": ref,
            "location": location,
            "recorded_by": who,
        }
        if kind == "Transfer out":
            row["to_location"] = "Market stall"
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 15. Returns & Damaged Goods
# ---------------------------------------------------------------------------

_RETURNS = [
    (-92, "Retail customer", "SO-1041", "buyer214", 6, 1,
     "Damaged in transit", "Damaged - not sellable", 18.00, "No",
     0.0, "Yes", "Yes", 6.10,
     "Mug arrived in three pieces. Supplier packed it badly; they "
     "credited us."),
    (-84, "Retail customer", "SO-1063", "buyer377", 0, 1,
     "Customer changed mind", "As new", 18.00, "No", 3.95, "No",
     "No", 0.0, "Back on the shelf, no harm done."),
    (-71, "Wholesale customer", "WO-0002", "The Fern Room", 17, 12,
     "Not as described", "Opened - resellable", 0.0, "Yes", 0.0,
     "No", "No", 0.0,
     "Card designs were not what they expected. Replaced with a "
     "different assortment; these went back into stock."),
    (-58, "Courier damage", "SO-1092", "buyer481", 9, 2,
     "Damaged in transit", "Destroyed", 52.00, "No", 0.0, "No", "No",
     0.0, "Claim with the courier, not the supplier."),
    (-44, "Retail customer", "SO-1130", "buyer092", 13, 1,
     "Incorrect size", "As new", 32.00, "No", 4.20, "No", "No",
     0.0, ""),
    (-37, "Wholesale customer", "WO-0004", "Alder & Oak Living", 11,
     6, "Defective", "Faulty - return to supplier", 0.0, "Yes",
     0.0, "Yes", "Yes", 26.40,
     "Glaze crazing on all six. Ridgeway replaced them."),
    (-29, "Retail customer", "SO-1178", "buyer655", 22, 1,
     "Defective", "Damaged - not sellable", 16.00, "Yes", 0.0,
     "Yes", "No", 0.0,
     "Balm had separated. Supplier's fault and we have still not "
     "claimed for it."),
    (-21, "Marketplace", "SO-1204", "buyer318", 0, 2,
     "Customer changed mind", "Opened - resellable", 36.00, "No",
     5.60, "No", "No", 0.0, ""),
    (-14, "Retail customer", "SO-1241", "buyer730", 20, 3,
     "Wrong item sent", "As new", 27.00, "Yes", 4.20, "No", "No",
     0.0, "Our picking error. Replacement sent same day."),
    (-8, "Own warehouse", "", "", 1, 4, "Damaged in transit",
     "Destroyed", 0.0, "No", 0.0, "Yes", "No", 0.0,
     "Found broken in the box from Kestrel. Their fault, not "
     "claimed yet."),
    (-3, "Wholesale customer", "WO-0011", "Sundry Goods Co.", 19, 24,
     "Not as described", "Opened - resellable", 0.0, "No", 0.0,
     "No", "No", 0.0, "Restocked, no refund - they took a credit "
     "against the next order."),
]


def returns(cfg: Config) -> Rows:
    rows: Rows = []
    for (ago, source, ref, who, sku_i, qty, reason, condition,
         refund, replacement, postage, fault, claimed, credit,
         note) in _RETURNS:
        rows.append({
            "date": _day(cfg, ago),
            "source": source,
            "reference": ref,
            "customer": who,
            "sku": _sku_at(sku_i),
            "qty": qty,
            "reason": reason,
            "condition": condition,
            "location": "Main store" if "resellable" in condition
                        or condition == "As new" else "",
            "refund": refund,
            "replacement": replacement,
            "return_shipping": postage,
            "supplier_fault": fault,
            "claimed": claimed,
            "credit_received": credit,
            "status": "Closed" if ago < -30 else
                      ("Claim raised" if claimed == "Yes"
                       else "Logged"),
            "notes": note,
        })
    return rows


# ---------------------------------------------------------------------------
# 16. Warehouse Locations
# ---------------------------------------------------------------------------

_ZONES = {
    "Main store": ("Zone A", "Shelf"),
    "Workshop": ("Zone B", "Bin"),
    "Overflow unit": ("Zone C", "Pallet"),
    "Shop floor": ("Front", "Shelf"),
    "Market stall": ("Mobile", "Bin"),
}

#: sku index -> [(location, share of stock, primary, capacity)]
_LOCATION_PLAN = {
    0: [("Main store", 0.7, "Yes", 400), ("Shop floor", 0.2, "No", 60),
        ("Market stall", 0.1, "No", 48)],
    2: [("Main store", 0.8, "Yes", 400), ("Shop floor", 0.2, "No", 60)],
    3: [("Main store", 0.6, "Yes", 300),
        ("Overflow unit", 0.4, "No", 600)],
    5: [("Main store", 1.0, "Yes", 300)],
    6: [("Main store", 0.75, "Yes", 240), ("Shop floor", 0.25, "No", 48)],
    7: [("Main store", 1.0, "Yes", 240)],
    9: [("Main store", 0.65, "Yes", 200), ("Shop floor", 0.35, "No", 40)],
    10: [("Overflow unit", 1.0, "Yes", 300)],
    12: [("Main store", 0.7, "Yes", 200), ("Market stall", 0.3, "No", 36)],
    13: [("Main store", 1.0, "Yes", 150)],
    15: [("Main store", 0.8, "Yes", 400), ("Shop floor", 0.2, "No", 60)],
    16: [("Main store", 0.6, "Yes", 500),
         ("Overflow unit", 0.4, "No", 800)],
    17: [("Main store", 1.0, "Yes", 400)],
    19: [("Main store", 1.0, "Yes", 900)],
    20: [("Main store", 0.7, "Yes", 300), ("Shop floor", 0.3, "No", 48)],
    21: [("Main store", 1.0, "Yes", 240)],
    23: [("Workshop", 0.8, "Yes", 1200),
         ("Overflow unit", 0.2, "No", 1500)],
    24: [("Workshop", 1.0, "Yes", 3000)],
    25: [("Workshop", 1.0, "Yes", 2000)],
    26: [("Workshop", 1.0, "Yes", 900)],
    27: [("Workshop", 0.6, "Yes", 700),
         ("Overflow unit", 0.4, "No", 900)],
}


def locations(cfg: Config) -> Rows:
    rows: Rows = []
    bins = {}
    for sku_i in sorted(_LOCATION_PLAN):
        item = _CATALOGUE[sku_i]
        opening = item[14]
        plan = _LOCATION_PLAN[sku_i]
        placed = 0
        for pos, (where, share, primary, capacity) in enumerate(plan):
            zone, kind = _ZONES[where]
            count = bins.get(where, 0) + 1
            bins[where] = count
            qty = (opening - placed if pos == len(plan) - 1
                   else int(opening * share))
            placed += qty
            note = ""
            if where == "Overflow unit":
                note = ("Off-site. Allow an extra day when picking "
                        "from here.")
            rows.append({
                "sku": item[0],
                "location": where,
                "zone": zone,
                "bin": f"{zone.split()[-1][0].upper()}"
                       f"{count:02d}-{pos + 1}",
                "qty_here": max(qty, 0),
                "is_primary": primary,
                "capacity": capacity,
                "type": kind,
                "checked": _day(cfg, -(7 + (sku_i * 11) % 90)),
                "notes": note,
            })
    return rows


# ---------------------------------------------------------------------------
# 17. Stock Audit
# ---------------------------------------------------------------------------
# A count never overwrites stock. It produces a variance, and the
# user posts an adjustment on the Movement Log. The unposted ones
# are what the dashboard counts.

#: (days ago, type, sku index, location, counter, counted, reason,
#:  recount, posted, movement ref, status)
_AUDITS = [
    (-104, "Full stocktake", 0, "Main store", "Owner", None, "", "No",
     "Yes", "", "Closed", ""),
    (-104, "Full stocktake", 6, "Main store", "Owner", None, "", "No",
     "Yes", "", "Closed", ""),
    (-74, "Cycle count", 16, "Main store", "Warehouse", 8,
     "Found elsewhere", "Yes", "Yes", "MOV-0005", "Closed",
     "A sealed box had been shelved behind the printer stock."),
    (-61, "Spot check", 12, "Overflow unit", "Warehouse", None, "",
     "No", "Yes", "", "Closed", ""),
    (-48, "Cycle count", 8, "Main store", "Owner", -5,
     "Sale not recorded", "Yes", "Yes", "MOV-0010", "Closed",
     "Five seconds mugs short. Most likely sold at the market "
     "stall without being logged."),
    (-33, "Cycle count", 19, "Main store", "Packer", None, "", "No",
     "Yes", "", "Closed", ""),
    (-26, "Spot check", 23, "Workshop", "Packer", None, "", "No",
     "Yes", "", "Closed", ""),
    (-12, "Cycle count", 7, "Main store", "Owner", -11, "Unknown",
     "Yes", "Yes", "MOV-0017", "Closed",
     "Eleven 250ml mugs short and no explanation found. Posted as "
     "an adjustment so the shelf and the system agree again."),
    (-9, "Cycle count", 20, "Main store", "Assistant", None, "",
     "No", "Yes", "", "Closed", ""),
    (-5, "Cycle count", 2, "Main store", "Warehouse", -3, "Unknown",
     "Yes", "No", "", "Investigating",
     "Three short. The recount agreed with the first count, so the "
     "stock genuinely is not there. Still not posted."),
    (-4, "Spot check", 13, "Main store", "Assistant", 2,
     "Receipt not recorded", "No", "No", "", "Recount needed",
     "Two more cushion covers than expected - a return that went "
     "back on the shelf without paperwork."),
    (-2, "Cycle count", 17, "Main store", "Packer", None, "", "No",
     "No", "", "Counted", ""),
    (-1, "Cycle count", 21, "Main store", "Warehouse", -2,
     "Damage not logged", "No", "No", "", "Recount needed",
     "Two pouches missing - most likely the split ones that were "
     "binned without being logged."),
]


def audit(cfg: Config) -> Rows:
    rows: Rows = []
    for (ago, kind, sku_i, where, who, delta, reason, recount,
         posted, ref, status, note) in _AUDITS:
        item = _CATALOGUE[sku_i]
        # a clean count matches the opening quantity; a variance is
        # that quantity plus the (signed) difference found
        counted = item[14] + (delta or 0)
        rows.append({
            "date": _day(cfg, ago),
            "audit_type": kind,
            "sku": item[0],
            "location": where,
            "counted_by": who,
            "counted_qty": max(counted, 0),
            "reason": reason,
            "recount": recount,
            "posted": posted,
            "movement_ref": ref,
            "status": status,
            "notes": note,
        })
    return rows


# ---------------------------------------------------------------------------
# 19. Expense Tracker
# ---------------------------------------------------------------------------
# Stock purchases are kept out of the operating total on purpose.
# Buying stock is not a cost until the stock sells; mixing the two
# is how people convince themselves a good month was a bad one.

#: (category, description, amount, method, recurring, tax category)
_MONTHLY_COSTS = [
    ("Warehouse rent", "Unit 4, Thornfield Yard - monthly rent",
     650.00, "Bank transfer", "Monthly", "Operating expense"),
    ("Utilities", "Electricity and water", 118.40, "Direct debit",
     "Monthly", "Operating expense"),
    ("Software subscriptions", "Shop subscription, accounting "
     "software and email", 96.00, "Card", "Monthly", "Operating expense"),
    ("Insurance", "Stock and public liability cover", 74.00,
     "Direct debit", "Monthly", "Operating expense"),
    ("Advertising & marketing", "Marketplace ads", 180.00, "Card",
     "Monthly", "Operating expense"),
    ("Wages & contractors", "Weekend packing help", 420.00,
     "Bank transfer", "Monthly", "Operating expense"),
]

#: (days ago, category, description, amount, method, tax category)
_ONE_OFF_COSTS = [
    (-112, "Packaging & mailers", "Branded tissue and ribbon - "
     "not for resale", 142.60, "Card", "Operating expense"),
    (-104, "Trade shows & markets", "Portland trade show stand",
     480.00, "Bank transfer", "Operating expense"),
    (-96, "Shipping & postage", "Courier account top-up", 300.00,
     "Card", "Operating expense"),
    (-88, "Equipment", "Heat sealer and label printer", 318.00,
     "Card", "Capital / asset"),
    (-77, "Accounting & legal", "Bookkeeper - quarterly", 225.00,
     "Bank transfer", "Operating expense"),
    (-66, "Shipping & postage", "Courier account top-up", 300.00,
     "Card", "Operating expense"),
    (-58, "Travel", "Van hire and fuel for the Ridgeway "
     "collection", 164.30, "Card", "Operating expense"),
    (-51, "Equipment repair", "Shelving repair after the "
     "gutter leak", 88.00, "Cash", "Operating expense"),
    (-44, "Trade shows & markets", "Harvest fair pitch fee",
     95.00, "Cash", "Operating expense"),
    (-38, "Shipping & postage", "Courier account top-up", 300.00,
     "Card", "Operating expense"),
    (-31, "Bank charges", "Quarterly account charges",
     42.00, "Direct debit", "Operating expense"),
    (-24, "Accounting & legal", "Accountant - year-end advice",
     180.00, "Bank transfer", "Operating expense"),
    (-18, "Packaging & mailers", "Void fill and tape", 96.50,
     "Card", "Operating expense"),
    (-11, "Shipping & postage", "Courier account top-up", 300.00,
     "Card", "Operating expense"),
    (-7, "Training", "Small-business pricing course",
     149.00, "Card", "Operating expense"),
    (-3, "Advertising & marketing", "Autumn catalogue print run",
     264.00, "Card", "Operating expense"),
]


def expenses(cfg: Config) -> Rows:
    rows: Rows = []
    # four months of fixed costs, billed on the 1st
    for months_back in range(3, -1, -1):
        when = cfg.today.replace(day=1)
        month = when.month - months_back
        year = when.year
        while month < 1:
            month += 12
            year -= 1
        billed = _dt.date(year, month, 1)
        if billed > cfg.today:
            continue
        for (category, description, amount, method, recurring,
             tax_cat) in _MONTHLY_COSTS:
            rows.append({
                "date": billed,
                "description": description,
                "category": category,
                "supplier": "",
                "amount": amount,
                "tax_amount": 0.0,
                "method": method,
                "recurring": recurring,
                "tax_category": tax_cat,
                "receipt": "Yes",
                "reference": "",
                "notes": "",
            })
    for (ago, category, description, amount, method, tax_cat) in \
            _ONE_OFF_COSTS:
        rows.append({
            "date": _day(cfg, ago),
            "description": description,
            "category": category,
            "supplier": "",
            "amount": amount,
            "tax_amount": 0.0,
            "method": method,
            "recurring": "One off",
            "tax_category": tax_cat,
            "receipt": "Yes" if ago < -10 else "No",
            "reference": "",
            "notes": ("No receipt kept. Find it before the "
                      "year end." if ago >= -10 else ""),
        })
    # and one that proves the point about stock purchases
    rows.append({
        "date": _day(cfg, -14),
        "description": "Ridgeway Ceramics invoice INV-9243 - stock "
                       "for resale",
        "category": "Stock purchases",
        "supplier": _sup(2),
        "amount": 1342.80,
        "tax_amount": 0.0,
        "method": "Bank transfer",
        "recurring": "One off",
        "tax_category": "Cost of sale",
        "receipt": "Yes",
        "reference": "PO-0011",
        "notes": "Logged here for the bank reconciliation only. It "
                 "is deliberately kept out of the operating total - "
                 "stock is a cost when it sells, not when it is "
                 "bought.",
    })
    return rows


# ---------------------------------------------------------------------------
# registry
# ---------------------------------------------------------------------------
# Keyed by register MODULE name, not sheet name. A mismatch here
# produces zero sample rows, silently, and the build still succeeds.

_BUILDERS: Dict[str, Callable[[Config], Rows]] = {
    "products": products,
    "movements": movements,
    "suppliers": suppliers,
    "purchases": purchases,
    "customers": customers,
    "pricing": pricing,
    "orders": orders,
    "retail": retail,
    "returns": returns,
    "locations": locations,
    "audit": audit,
    "expenses": expenses,
}


def get(module: str, cfg: Config) -> Rows:
    """Sample rows for a register module, or an empty list."""
    builder = _BUILDERS.get(module)
    if builder is None:
        return []
    return builder(cfg)
````

### `inventory_wholesale_spreadsheet/build.py`

*168 lines · 3 functions*

> Workbook builder: puts the twenty-three tabs together in order.
>
> Typical use::
>
>     from inventory_wholesale_spreadsheet import build_workbook
>     build_workbook("Ultimate Inventory & Wholesale Order Manager.xlsx")

| Defines | What it does |
|---------|--------------|
| `build_register(book, fmts, cfg, reg, module)` | Create one register sheet from its :class:`SheetSpec` |
| `build_workbook(path, cfg)` | Write the whole product to *path* and return the path |
| `_set_properties(book, cfg)` |  |

**Module constants:** `REGISTER_MODULES` (17 items), `SHEET_ORDER` (23 items), `PAGE_BUILDERS` (6 entries), `REGISTER_BY_NAME` ({module.NAME: module for module in REGISTER_MODULES}).

````python
"""Workbook builder: puts the twenty-three tabs together in order.

Typical use::

    from inventory_wholesale_spreadsheet import build_workbook
    build_workbook("Ultimate Inventory & Wholesale Order Manager.xlsx")
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
from .sheets import (annual, audit, customers, dashboard, expenses,
                     getting_started, inventory, locations, monthly,
                     movements, orders, performance, pricing,
                     printables, products, profit, purchases, reorder,
                     retail, returns, settings, suppliers, valuation)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    products, inventory, movements, suppliers, purchases, customers,
    pricing, orders, retail, reorder, valuation, returns, locations,
    audit, performance, expenses, profit,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Dashboard",
    "Getting Started",
    "Settings & Lists",
    "Product & SKU Database",
    "Inventory Tracker",
    "Stock Movement Log",
    "Supplier Database",
    "Purchase Orders",
    "Wholesale Customer CRM",
    "Wholesale Pricing",
    "Wholesale Orders",
    "Retail Sales Tracker",
    "Reorder Planner",
    "Inventory Valuation & COGS",
    "Returns & Damaged Goods",
    "Warehouse Locations",
    "Stock Audit",
    "Product Performance",
    "Expense Tracker",
    "Profit Analysis",
    "Monthly Reports",
    "Annual Reports",
    "Printable Orders & Invoices",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Dashboard": dashboard.build,
    "Getting Started": getting_started.build,
    "Settings & Lists": settings.build,
    "Monthly Reports": monthly.build,
    "Annual Reports": annual.build,
    "Printable Orders & Invoices": printables.build,
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


def build_workbook(path: str = "Inventory & Wholesale Manager.xlsx",
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

    # 3. Live lookup names - SKUs above all, because nine tabs take
    #    their drop-down from that one column.
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
        "subject": "Products, SKUs, stock levels, suppliers, purchase "
                   "orders, wholesale customers, tiered pricing, "
                   "orders, reordering, valuation and profit by "
                   "product",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("inventory management spreadsheet, stock "
                     "tracker, SKU tracker, wholesale order "
                     "manager, purchase order template, supplier "
                     "database, reorder point calculator, inventory "
                     "valuation, COGS tracker, profit by product, "
                     "small business inventory template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {VERSION}. Formula cells are locked "
                     f"and every sheet is protected with the "
                     f"password '{cfg.password}'. Stock is derived "
                     f"from recorded transactions and is never typed "
                     f"over. Costing is weighted average throughout. "
                     f"This workbook does not connect to or "
                     f"synchronise with any marketplace."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

---

## 3. Shared column factories

The helpers every register uses: automatic ID columns, date and month keys, ranking columns, alert conditional formats, the status-set tests that decide what counts as reserved or shipped, and the weighted average cost lookup that every margin in the workbook is measured against.

### `inventory_wholesale_spreadsheet/sheets/__init__.py`

*0 lines*

````python

````

### `inventory_wholesale_spreadsheet/sheets/common.py`

*691 lines · 62 functions*

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
| `rank_desc(key, guard, header, width, note)` | Hidden dense rank, 1 = the biggest value in `key` |
| `rank_asc(key, guard, header, width, note)` | Hidden dense rank, 1 = the smallest value in `key` |
| `seq_key(parent, trigger)` | Hidden `1, 2, 3 ...` within a parent, plus a `parent|n` key |
| `days_until(date_cell)` | Whole days from today to a date - negative once it has passed |
| `countdown_band(days_cell)` | Turn a day count into the band the dashboard groups on |
| `lookup_product(target, key, missing)` | A field from the product record behind a SKU |
| `lookup_product_num(target, key)` |  |
| `lookup_product_num_zero(target, key)` |  |
| `lookup_supplier(target, key, missing)` |  |
| `lookup_supplier_num(target, key)` |  |
| `lookup_customer(target, key, missing)` |  |
| `lookup_customer_num(target, key)` |  |
| `lookup_tier(target, key)` | A price break from the Wholesale Pricing table for this SKU |
| `product_name_col(key)` | Standard read-only product name beside a SKU column |
| `sku_col(key, header, note)` | The join key, as a drop-down off the live product list |
| `in_list(cell, values)` | `TRUE` when `cell` matches any of `values` |
| `is_reserving(cell)` | A wholesale order that has committed stock but not shipped it |
| `is_shipped(cell)` | A wholesale order whose stock has physically left the building |
| `is_sold_retail(cell)` | A retail sale that has actually taken stock off the shelf |
| `is_open_po(cell)` | A purchase order still expected to arrive |
| `is_committed_po(cell)` | A purchase order the supplier has agreed to - stock on order |
| `is_sellable(cell)` | A returned item that can go back on the shelf |
| `sum_by_sku(sheet, value_col, sku_cell, *extra)` | `SUMIFS` of a column on another sheet, matched on SKU |
| `count_by_sku(sheet, sku_cell, *extra)` |  |
| `received_qty(sku_cell)` | Units actually booked in against purchase orders |
| `sold_qty(sku_cell)` | Units gone: shipped wholesale plus completed retail |
| `reserved_qty(sku_cell)` | Units promised to a confirmed order that has not shipped |
| `returned_sellable_qty(sku_cell)` | Returned units that went back on the shelf |
| `damaged_qty(sku_cell)` | Units written off: returned unsellable, plus damage logged |
| `adjustment_qty(sku_cell)` | Manual corrections, which only ever come from the movement log |
| `weighted_average_cost(sku_cell, key)` | The costing method this workbook uses, in one place |
| `tier_price(qty_cell, sku_cell, key)` | The right price break for this quantity, off Wholesale Pricing |
| `sales_tax(amount)` | Sales tax / VAT on an order total |
| `monthly_equivalent(amount, frequency_cell)` | Turn any recurring cost into what it really costs each month |
| `margin(profit_cell, revenue_cell)` | Margin, guarded |
| `alert_chain(trigger, tests, default)` | Build the nested IF that drives a sheet's "what needs doing" |

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

from ..config import (COMMITTED_PO_STATUSES, OPEN_PO_STATUSES,
                      RESERVING_STATUSES, SELLABLE_CONDITIONS,
                      SHIPPED_STATUSES, SOLD_RETAIL_STATUSES)
from ..layout import HEADER_ROW, Col
from ..writer import cf_text, list_from, yes_no

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


# ---------------------------------------------------------------------------
# Cross sheet lookups - the five join keys
# ---------------------------------------------------------------------------
# SKU is the join key for everything that moves. Supplier ID, Purchase
# Order ID, Customer ID and Wholesale Order ID join the paperwork. Every
# one of them is a drop-down somewhere, never a typed string.

def lookup_product(target: str, key: str = "sku",
                   missing: str = "Not in Product Database") -> str:
    """A field from the product record behind a SKU."""
    return lookup("Product & SKU Database", target, "sku", key, missing)


def lookup_product_num(target: str, key: str = "sku") -> str:
    return lookup_num("Product & SKU Database", target, "sku", key)


def lookup_product_num_zero(target: str, key: str = "sku") -> str:
    return lookup_num_zero("Product & SKU Database", target, "sku", key)


def lookup_supplier(target: str, key: str = "supplier",
                    missing: str = "Not in Supplier Database") -> str:
    return lookup("Supplier Database", target, "id", key, missing)


def lookup_supplier_num(target: str, key: str = "supplier") -> str:
    return lookup_num("Supplier Database", target, "id", key)


def lookup_customer(target: str, key: str = "customer",
                    missing: str = "Not in Wholesale Customer CRM") -> str:
    return lookup("Wholesale Customer CRM", target, "id", key, missing)


def lookup_customer_num(target: str, key: str = "customer") -> str:
    return lookup_num("Wholesale Customer CRM", target, "id", key)


def lookup_tier(target: str, key: str = "sku") -> str:
    """A price break from the Wholesale Pricing table for this SKU."""
    return lookup_num("Wholesale Pricing", target, "sku", key)


def product_name_col(key: str = "sku") -> Col:
    """Standard read-only product name beside a SKU column.

    Every register that touches stock shows the product name, and not
    one of them lets you type it: a name typed twice is a name spelled
    two ways, and then the SUMIFS stop agreeing with each other.
    """
    return Col("product", "Product name", 26, "text",
               formula=lookup_product("name", key, ""),
               note="Looked up from the Product & SKU Database. "
                    "Change the name there and it changes here.")


def sku_col(key: str = "sku", header: str = "SKU",
            note: str = "") -> Col:
    """The join key, as a drop-down off the live product list."""
    return Col(key, header, 14, "id",
               validate=list_from("SKUList"),
               note=note or ("Pick the SKU from the list. Never type "
                             "one by hand - a SKU that is not in the "
                             "Product Database silently disconnects "
                             "this row from stock, costing and "
                             "profit."))


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


def is_reserving(cell: str) -> str:
    """A wholesale order that has committed stock but not shipped it.

    This is the whole reason reserved stock exists as a separate
    figure. The goods are still on the shelf, so they are still in
    stock - but they are spoken for, so they are not available to sell
    to anybody else.
    """
    return in_list(cell, RESERVING_STATUSES)


def is_shipped(cell: str) -> str:
    """A wholesale order whose stock has physically left the building."""
    return in_list(cell, SHIPPED_STATUSES)


def is_sold_retail(cell: str) -> str:
    """A retail sale that has actually taken stock off the shelf."""
    return in_list(cell, SOLD_RETAIL_STATUSES)


def is_open_po(cell: str) -> str:
    """A purchase order still expected to arrive."""
    return in_list(cell, OPEN_PO_STATUSES)


def is_committed_po(cell: str) -> str:
    """A purchase order the supplier has agreed to - stock on order."""
    return in_list(cell, COMMITTED_PO_STATUSES)


def is_sellable(cell: str) -> str:
    """A returned item that can go back on the shelf."""
    return in_list(cell, SELLABLE_CONDITIONS)


# ---------------------------------------------------------------------------
# Stock arithmetic
# ---------------------------------------------------------------------------
# Rule one of this workbook: stock is never typed, it is always derived.
# The only way to change a stock figure is to add a row somewhere that
# says why it changed. These helpers are the sums that do it.

def sum_by_sku(sheet: str, value_col: str, sku_cell: str,
               *extra: str) -> str:
    """`SUMIFS` of a column on another sheet, matched on SKU.

    `extra` arrives as already-paired criteria range and criteria
    strings - a criterion is always *two* arguments, and collapsing
    them into one is the quickest way to a `#VALUE!`.
    """
    parts = [f"[[{sheet}.{value_col}]]", f"[[{sheet}.sku]]", sku_cell]
    parts.extend(extra)
    return f'SUMIFS({",".join(parts)})'


def count_by_sku(sheet: str, sku_cell: str, *extra: str) -> str:
    parts = [f"[[{sheet}.sku]]", sku_cell]
    parts.extend(extra)
    return f'COUNTIFS({",".join(parts)})'


def received_qty(sku_cell: str) -> str:
    """Units actually booked in against purchase orders.

    Read straight off the purchase order line, which is the single
    place a receipt is recorded. One PO line is one cell, so the same
    delivery physically cannot be counted twice.
    """
    return sum_by_sku("Purchase Orders", "received", sku_cell)


def sold_qty(sku_cell: str) -> str:
    """Units gone: shipped wholesale plus completed retail."""
    wholesale = sum_by_sku("Wholesale Orders", "qty", sku_cell,
                           "[[Wholesale Orders.is_shipped]]", "1")
    retail = sum_by_sku("Retail Sales Tracker", "qty", sku_cell,
                        "[[Retail Sales Tracker.is_sold]]", "1")
    return f"({wholesale}+{retail})"


def reserved_qty(sku_cell: str) -> str:
    """Units promised to a confirmed order that has not shipped."""
    return sum_by_sku("Wholesale Orders", "qty", sku_cell,
                      "[[Wholesale Orders.is_reserved]]", "1")


def returned_sellable_qty(sku_cell: str) -> str:
    """Returned units that went back on the shelf."""
    return sum_by_sku("Returns & Damaged Goods", "qty", sku_cell,
                      "[[Returns & Damaged Goods.restocked]]", "1")


def damaged_qty(sku_cell: str) -> str:
    """Units written off: returned unsellable, plus damage logged."""
    returns = sum_by_sku("Returns & Damaged Goods", "qty", sku_cell,
                         "[[Returns & Damaged Goods.restocked]]", "0")
    movements = sum_by_sku("Stock Movement Log", "damaged", sku_cell)
    return f"({returns}+{movements})"


def adjustment_qty(sku_cell: str) -> str:
    """Manual corrections, which only ever come from the movement log."""
    return sum_by_sku("Stock Movement Log", "adjustment", sku_cell)


# ---------------------------------------------------------------------------
# Money rules used in more than one place
# ---------------------------------------------------------------------------

def weighted_average_cost(sku_cell: str, key: str = "sku") -> str:
    """The costing method this workbook uses, in one place.

    Weighted average, not FIFO. Opening stock at its stated cost plus
    everything received at what it actually cost, divided by the units
    those two represent. It is the only method that survives a
    spreadsheet where rows get edited out of order, and mixing it with
    FIFO halfway through is how inventory numbers stop reconciling.
    """
    open_qty = lookup_product_num_zero("opening_qty", key)[1:]
    open_val = lookup_product_num_zero("opening_value", key)[1:]
    recv_qty = received_qty(sku_cell)
    recv_val = sum_by_sku("Purchase Orders", "received_value", sku_cell)
    return (f'IF(({open_qty})+({recv_qty})<=0,'
            f'{lookup_product_num_zero("cost", key)[1:]},'
            f'ROUND((({open_val})+({recv_val}))/'
            f'(({open_qty})+({recv_qty})),4))')


def tier_price(qty_cell: str, sku_cell: str,
               key: str = "sku") -> str:
    """The right price break for this quantity, off Wholesale Pricing.

    Checked from the largest break down, so a 100 unit order gets the
    100+ price rather than stopping at the first band it qualifies for.
    An unknown SKU falls back to the product's own wholesale price
    rather than to zero, because a price of nothing on an order line is
    a mistake that ships.
    """
    def band(n: int) -> str:
        return (f'IFERROR(INDEX([[Wholesale Pricing.tier{n}_qty]],'
                f'MATCH({sku_cell},[[Wholesale Pricing.sku]],0))*1,0)')

    def price(n: int) -> str:
        return (f'IFERROR(INDEX([[Wholesale Pricing.tier{n}_price]],'
                f'MATCH({sku_cell},[[Wholesale Pricing.sku]],0))*1,0)')

    base = (f'IFERROR(INDEX([[Wholesale Pricing.wholesale]],'
            f'MATCH({sku_cell},[[Wholesale Pricing.sku]],0))*1,'
            f'{lookup_product_num_zero("wholesale_price", key)[1:]})')
    out = base
    for n in (1, 2, 3):
        out = (f'IF(AND({band(n)}>0,{qty_cell}>={band(n)},'
               f'{price(n)}>0),{price(n)},{out})')
    return out


def sales_tax(amount: str) -> str:
    """Sales tax / VAT on an order total.

    Whether you charge it at all, and at what rate, is a local
    question - and a wholesale order to a reseller usually is not
    taxed. Both the rate and the switch live on Settings and the rate
    ships at zero.
    """
    return f'IF(TaxOnSales<>"Yes",0,ROUND({amount}*TaxRate,2))'


def monthly_equivalent(amount: str, frequency_cell: str) -> str:
    """Turn any recurring cost into what it really costs each month.

    A yearly software licence and a monthly warehouse rent are not
    comparable until both are expressed per month, and anybody looking
    at an expense breakdown is asking the per-month question.
    """
    return (f'IF({frequency_cell}="","",'
            f'IFERROR(INDEX(FrequencyPerYear,MATCH({frequency_cell},'
            f'FrequencyList,0))*{amount}/12,0))')


def margin(profit_cell: str, revenue_cell: str) -> str:
    """Margin, guarded. Dividing by a blank revenue cell is #DIV/0!."""
    return (f'IF({revenue_cell}<=0,"",'
            f'ROUND({profit_cell}/{revenue_cell},4))')


# ---------------------------------------------------------------------------
# Alert chains
# ---------------------------------------------------------------------------

def alert_chain(trigger: str, tests: List[Any],
                default: str = "OK") -> str:
    """Build the nested IF that drives a sheet's "what needs doing".

    Pass the tests in priority order as (condition, message) pairs and
    get back a balanced formula.  Hand nesting a dozen IFs is how a
    workbook ends up with a stray bracket on row 300 of one sheet, so
    nobody writes these by hand.

    A message may itself be a nested chain: pass a list of pairs as the
    message and it is expanded in place, which is how "this product is
    active, so now ask the stock questions" stays readable.
    """
    def build(pairs: List[Any], fallback: str) -> str:
        if not pairs:
            return fallback
        head = pairs[0]
        cond, msg = head[0], head[1]
        inner_default = head[2] if len(head) > 2 else ""
        rest = build(pairs[1:], fallback)
        if isinstance(msg, list):
            inner = build(msg, f'"{inner_default}"')
            return f'IF({cond},{inner},{rest})'
        return f'IF({cond},"{msg}",{rest})'

    body = build(list(tests), f'"{default}"')
    return f'=IF({{{trigger}}}{{r}}="","",{body})'
````

---

## 4. Settings and the two front pages

Settings & Lists is the only tab a buyer has to fill in before anything else works: the business, the markup and margin floor, the lead time and safety stock, the selling fees, the targets and the 40 drop-down lists. The Dashboard is the centrepiece - stock value, what is out of stock, what needs ordering, what is reserved and unpicked, what is late, and a matrix that counts every alert on seventeen registers. Getting Started explains the SKU and is honest about what the workbook is not.

### `inventory_wholesale_spreadsheet/sheets/settings.py`

*776 lines · 2 functions*

> `Settings & Lists` - every number and every list the workbook runs on.
>
> Nothing in this file is hard coded to one country, one currency or one
> way of selling. Your price tiers, your lead times, your safety stock,
> your marketplace fee rates, your minimum acceptable margin and all
> forty-odd drop-downs are set here, once, and every other tab follows.
>
> Two groups of settings deserve a second look before you start.
>
> The ID prefixes decide what your Product IDs, Purchase Order numbers
> and Order IDs look like. Change them before you have typed three
> hundred products, not after.
>
> The fee rates decide whether your profit figures are true. A
> marketplace taking 6.5% and a payment processor taking 2.9% plus a
> fixed charge are the difference between a healthy margin and a
> hobby, and the defaults here are only a starting point - put your
> own in.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |
| `_month_keys(today, back, ahead)` | A window of yyyy-mm keys around today |

**Module constants:** `NAME` ('Settings & Lists'), `LAST_COL` (11), `YES_NO` (2 items), `PRODUCT_STATUSES` (7 items), `SUPPLIER_TYPES` (10 items), `SUPPLIER_STATUSES` (6 items), `DELIVERY_CONDITIONS` (7 items), `CUSTOMER_TYPES` (13 items), `CUSTOMER_STATUSES` (6 items), `LEAD_SOURCES` (11 items), `SHIPPING_METHODS` (9 items), `SALE_STATUSES` (8 items), `RETURN_SOURCES` (6 items), `RETURN_STATUSES` (6 items) and 11 more.

````python
"""`Settings & Lists` - every number and every list the workbook runs on.

Nothing in this file is hard coded to one country, one currency or one
way of selling. Your price tiers, your lead times, your safety stock,
your marketplace fee rates, your minimum acceptable margin and all
forty-odd drop-downs are set here, once, and every other tab follows.

Two groups of settings deserve a second look before you start.

The ID prefixes decide what your Product IDs, Purchase Order numbers
and Order IDs look like. Change them before you have typed three
hundred products, not after.

The fee rates decide whether your profit figures are true. A
marketplace taking 6.5% and a payment processor taking 2.9% plus a
fixed charge are the difference between a healthy margin and a
hobby, and the defaults here are only a starting point - put your
own in.
"""

from __future__ import annotations

import datetime as _dt
from typing import List, Tuple

from ..config import (CATEGORIES, COSTING_NOTE, CUSTOMER_TIERS,
                      DISCLAIMER, EXPENSE_CATEGORIES, HONESTY_NOTE,
                      ID_PREFIX_FIELDS, MONTH_NAMES, MOVEMENT_TYPES,
                      ORDER_STATUSES, PAYMENT_TERMS, PO_STATUSES,
                      PRICE_TIERS, PRIVACY_NOTE, RESERVED_NOTE,
                      RETURN_CONDITIONS, RETURN_REASONS,
                      SALES_CHANNELS, STOCK_CATEGORIES, STOCK_NOTE,
                      TAX_CATEGORIES, UNITS, Config, Palette)
from ..layout import Registry, col_letter, dynamic_name, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Settings & Lists"
LAST_COL = 11

# ---------------------------------------------------------------------------
# lists the config does not already carry, defined where they are used
# ---------------------------------------------------------------------------

YES_NO = ["Yes", "No"]

PRODUCT_STATUSES = ["Active", "Seasonal", "Clearance", "Pre-order",
                    "Out of production", "Discontinued", "Draft"]

SUPPLIER_TYPES = ["Manufacturer", "Wholesaler", "Distributor",
                  "Importer", "Maker / artisan", "Printer",
                  "Packaging supplier", "Dropshipper",
                  "Local retailer", "Other"]

SUPPLIER_STATUSES = ["Active", "Preferred", "Occasional", "Trial",
                     "On hold", "Inactive"]

DELIVERY_CONDITIONS = ["Good", "Minor damage", "Damaged",
                       "Short delivery", "Over delivery",
                       "Wrong items", "Rejected"]

CUSTOMER_TYPES = ["Independent boutique", "Gift shop", "Garden centre",
                  "Department store", "Online retailer",
                  "Marketplace seller", "Subscription box",
                  "Cafe or restaurant", "Hotel or B&B", "Museum shop",
                  "Distributor", "Corporate buyer", "Other"]

CUSTOMER_STATUSES = ["Prospect", "Active", "Occasional", "Lapsed",
                     "On hold", "Closed"]

LEAD_SOURCES = ["Trade show", "Instagram", "Faire or marketplace",
                "Website enquiry", "Cold outreach", "Referral",
                "Existing retail customer", "Sales agent",
                "Trade directory", "Walk-in", "Other"]

SHIPPING_METHODS = ["Collection", "Standard post", "Tracked post",
                    "Courier", "Next day courier", "Pallet",
                    "Freight", "Own delivery", "Dropship"]

SALE_STATUSES = ["Pending", "Paid", "Packing", "Shipped",
                 "Delivered", "Collected", "Refunded", "Cancelled"]

RETURN_SOURCES = ["Retail customer", "Wholesale customer",
                  "Marketplace", "Own warehouse", "Courier damage",
                  "Supplier"]

RETURN_STATUSES = ["Logged", "Inspecting", "Awaiting customer",
                   "Claim raised", "Resolved", "Closed"]

AUDIT_TYPES = ["Full stocktake", "Cycle count", "Spot check",
               "Investigation", "Year end"]

AUDIT_STATUSES = ["Counted", "Recount needed", "Investigating",
                  "Adjusted", "Closed"]

VARIANCE_REASONS = ["Miscount", "Damage not logged",
                    "Sale not recorded", "Receipt not recorded",
                    "Theft or loss", "Wrong SKU picked",
                    "Returned to supplier", "Sample taken",
                    "Used in production", "Found elsewhere",
                    "Unknown"]

STORAGE_TYPES = ["Shelf", "Bin", "Pallet", "Rack", "Drawer",
                 "Cold store", "Fragile", "Secure", "Overflow",
                 "Third party warehouse"]

PAYMENT_METHODS = ["Bank transfer", "Card", "Direct debit",
                   "Standing order", "PayPal", "Cash", "Cheque",
                   "Credit account", "Other"]

RECURRING_FREQUENCIES = ["One off", "Weekly", "Fortnightly",
                         "Monthly", "Quarterly", "Twice a year",
                         "Yearly"]

FREQUENCY_PER_YEAR = {"One off": 0.0, "Weekly": 52.0,
                      "Fortnightly": 26.0, "Monthly": 12.0,
                      "Quarterly": 4.0, "Twice a year": 2.0,
                      "Yearly": 1.0}

CURRENCIES = ["USD", "GBP", "EUR", "CAD", "AUD", "NZD", "JPY", "INR",
              "PKR", "ZAR", "SEK", "CHF", "Other"]

STAFF = ["Owner", "Warehouse", "Packer", "Assistant", "Bookkeeper",
         "Casual help"]

LOCATIONS = ["Main store", "Workshop", "Garage", "Shop floor",
             "Overflow unit", "Market stall", "Third party"]

#: name -> values. Everything the workbook offers you in a drop-down.
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", list(YES_NO)),
    ("ProductStatusList", list(PRODUCT_STATUSES)),
    ("CategoryList", [name for name, _subs in CATEGORIES]),
    ("SubcategoryList", [sub for _n, subs in CATEGORIES
                         for sub in subs]),
    ("UnitList", list(UNITS)),
    ("SupplierTypeList", list(SUPPLIER_TYPES)),
    ("SupplierStatusList", list(SUPPLIER_STATUSES)),
    ("POStatusList", list(PO_STATUSES)),
    ("DeliveryConditionList", list(DELIVERY_CONDITIONS)),
    ("CustomerTypeList", list(CUSTOMER_TYPES)),
    ("CustomerStatusList", list(CUSTOMER_STATUSES)),
    ("CustomerTierList", [name for name, _m, _d in CUSTOMER_TIERS]),
    ("LeadSourceList", list(LEAD_SOURCES)),
    ("OrderStatusList", list(ORDER_STATUSES)),
    ("ShippingMethodList", list(SHIPPING_METHODS)),
    ("ChannelList", list(SALES_CHANNELS)),
    ("SaleStatusList", list(SALE_STATUSES)),
    ("MovementTypeList", [name for name, _s, _d in MOVEMENT_TYPES]),
    ("ReturnSourceList", list(RETURN_SOURCES)),
    ("ReturnReasonList", list(RETURN_REASONS)),
    ("ReturnConditionList", list(RETURN_CONDITIONS)),
    ("ReturnStatusList", list(RETURN_STATUSES)),
    ("AuditTypeList", list(AUDIT_TYPES)),
    ("AuditStatusList", list(AUDIT_STATUSES)),
    ("VarianceReasonList", list(VARIANCE_REASONS)),
    ("StorageTypeList", list(STORAGE_TYPES)),
    ("LocationList", list(LOCATIONS)),
    ("StaffList", list(STAFF)),
    ("ExpenseCategoryList", list(EXPENSE_CATEGORIES)),
    ("StockCategoryList", list(STOCK_CATEGORIES)),
    ("TaxCategoryList", list(TAX_CATEGORIES)),
    ("PaymentMethodList", list(PAYMENT_METHODS)),
    ("FrequencyList", list(RECURRING_FREQUENCIES)),
    ("CurrencyList", list(CURRENCIES)),
    ("MonthNameList", list(MONTH_NAMES)),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    sheet_ref = quote_sheet(NAME)

    s.banner("Settings & Lists",
             "Your business, your prices, your rules, your lists. "
             "Set these once and every tab in the workbook follows.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 30)
    s.set_col(3, 3, 18)
    s.set_col(4, LAST_COL, 17)

    row = 4
    row = s.note(
        row,
        "White cells are yours to change; everything else is locked "
        "and calculated. Nothing here assumes a country, a currency "
        "or a way of selling - set your own price breaks, your own "
        "lead times, your own fee rates and your own targets, and "
        "the rest of the workbook follows.",
        LAST_COL, height=34)
    row += 1

    # ------------------------------------------------------ business --
    row = s.section(row, "Your business", LAST_COL)
    row = s.field(row, "Business name", cfg.business.name,
                  name="BusinessName", value_end=5, help_col=7,
                  help_text="Printed at the top of every report, "
                            "purchase order and invoice.")
    row = s.field(row, "Owner", cfg.business.owner, name="OwnerName",
                  value_end=5, help_col=7)
    row = s.field(row, "Email", cfg.business.email, value_end=5,
                  help_col=7)
    row = s.field(row, "Phone", cfg.business.phone, value_end=5,
                  help_col=7)
    row = s.field(row, "Website", cfg.business.website, value_end=5,
                  help_col=7)
    row = s.field(row, "Shop name", cfg.business.shop, value_end=5,
                  help_col=7,
                  help_text="Your marketplace shop, if you have one.")
    row = s.field(row, "Address line 1", cfg.business.address_1,
                  value_end=5, help_col=7)
    row = s.field(row, "Address line 2", cfg.business.address_2,
                  value_end=5, help_col=7)
    row = s.field(row, "City", cfg.business.city, value_end=5,
                  help_col=7)
    row = s.field(row, "Region", cfg.business.region, value_end=5,
                  help_col=7)
    row = s.field(row, "Postcode", cfg.business.postcode)
    row = s.field(row, "Country", cfg.business.country, value_end=5,
                  help_col=7)
    row = s.field(row, "Tax / VAT number", cfg.business.tax_id,
                  value_end=5, help_col=7)
    row = s.field(row, "Currency symbol", cfg.currency,
                  name="CurrencySign",
                  help_text="Shown on screen. The cell formats "
                            "follow the symbol you build the file "
                            "with.")
    row = s.field(row, "Currency code", cfg.currency_code,
                  validate=list_from("CurrencyList"))
    row = s.field(row, "Today", None, style="date", name="TodayDate",
                  mode="calc", formula="=TODAY()",
                  help_text="Live. Everything that says 'overdue', "
                            "'due soon' or 'this month' reads this "
                            "one cell.")
    row = s.field(row, "Reporting year", cfg.report_year, style="int",
                  name="ReportYear", validate=number_check(2000, 2100),
                  help_text="The year the Profit Analysis, Monthly "
                            "and Annual reports cover.")
    row = s.field(row, "Year starts", None, style="date",
                  name="YearStart", mode="calc",
                  formula="=DATE(ReportYear,1,1)")
    row = s.field(row, "Year ends", None, style="date",
                  name="YearEnd", mode="calc",
                  formula="=DATE(ReportYear,12,31)")
    row = s.field(row, "This month", None, style="text",
                  name="ThisMonthKey", mode="calc",
                  formula='=TEXT(TODAY(),"yyyy-mm")',
                  help_text="The month key every monthly total "
                            "groups on.")
    row += 1

    # ------------------------------------------------- stock settings --
    row = s.section(row, "How you manage stock", LAST_COL)
    row = s.note(
        row,
        "These five numbers drive the Reorder Planner. Safety stock "
        "is expressed in days of cover rather than a flat number of "
        "units, so a line selling ten a day automatically gets a "
        "bigger cushion than one selling ten a month.",
        LAST_COL, height=34)
    row += 1
    row = s.field(row, "Sales window for averages",
                  cfg.sales_history_days, style="int",
                  name="SalesWindow", validate=number_check(7, 730),
                  help_text="Days of sales history used to work out "
                            "average daily sales. Ninety days is a "
                            "reasonable default; shorten it if your "
                            "trade is seasonal and you want the "
                            "planner to react faster.")
    row = s.field(row, "Safety stock, in days", cfg.safety_stock_days,
                  style="int", name="SafetyDays",
                  validate=number_check(0, 365),
                  help_text="The cushion on top of the lead time. "
                            "Ten days is comfortable for most "
                            "sellers.")
    row = s.field(row, "Default lead time",
                  cfg.default_lead_time_days, style="int",
                  name="DefaultLeadDays", validate=number_check(0, 365),
                  help_text="Used when a supplier has no lead time "
                            "of their own on the Supplier Database.")
    row = s.field(row, "Slow moving after", cfg.slow_stock_days,
                  style="int", name="SlowStockDays",
                  validate=number_check(7, 999),
                  help_text="Days without a sale before a product is "
                            "called slow.")
    row = s.field(row, "Dead stock after", cfg.dead_stock_days,
                  style="int", name="DeadStockDays",
                  validate=number_check(7, 999),
                  help_text="Days without a sale before it is dead "
                            "money on a shelf.")
    row = s.field(row, "Count A lines every", cfg.stocktake_months,
                  style="int", name="StocktakeMonths",
                  validate=number_check(1, 24),
                  help_text="Months between cycle counts of your "
                            "highest value lines.")
    row += 1

    # ------------------------------------------------------- the money --
    row = s.section(row, "Pricing and margin", LAST_COL)
    row = s.note(
        row,
        "Minimum margin is the one to set carefully: it is the line "
        "below which the workbook starts calling a price a problem, "
        "on wholesale tiers, on order lines and on individual "
        "products. Everything else here is a default for new rows.",
        LAST_COL, height=34)
    row += 1
    row = s.field(row, "Minimum acceptable margin",
                  cfg.minimum_margin, style="percent",
                  name="MinGP", validate=number_check(0, 1),
                  help_text="Gross margin, not markup. Anything "
                            "below this is flagged wherever it "
                            "appears.")
    row = s.field(row, "Target gross margin", cfg.target_gross_margin,
                  style="percent", name="TargetMargin",
                  validate=number_check(0, 1))
    row = s.field(row, "Default retail markup", cfg.default_markup,
                  style="number1", name="DefaultMarkup",
                  validate=number_check(1, 20),
                  help_text="Cost multiplied by this gives a "
                            "starting retail price. A multiple of "
                            "2.5 is common for handmade goods sold "
                            "both ways.")
    row = s.field(row, "Wholesale as a share of retail",
                  cfg.wholesale_of_retail, style="percent",
                  name="WholesaleRatio", validate=number_check(0, 1),
                  help_text="Half of retail is the convention. A "
                            "stockist needs to double your price "
                            "and still be competitive.")
    row = s.field(row, "Default minimum order quantity",
                  cfg.default_moq, style="int", name="DefaultMOQ",
                  validate=number_check(0, 100000))
    row = s.field(row, "Default credit limit",
                  cfg.default_credit_limit, style="money",
                  name="DefaultCreditLimit",
                  validate=number_check(0, 10000000))
    row += 1

    # ------------------------------------------------- fees and extras --
    row = s.section(row, "What selling actually costs you", LAST_COL)
    row = s.note(
        row,
        "Marketplace commission, payment processing and packaging "
        "come off every retail sale before you see a penny of it. "
        "The rates below are starting points - find your real ones "
        "on a statement and put those in, because this is where "
        "small sellers most often think they are making money and "
        "are not.",
        LAST_COL, height=42)
    row += 1
    row = s.field(row, "Default marketplace fee",
                  cfg.marketplace_fee_pct, style="percent",
                  name="MarketplaceFeePct",
                  validate=number_check(0, 1),
                  help_text="Used for any channel without its own "
                            "rate in the table below.")
    row = s.field(row, "Payment processing", cfg.payment_fee_pct,
                  style="percent", name="PaymentFeePct",
                  validate=number_check(0, 1))
    row = s.field(row, "Payment fixed fee", cfg.payment_fee_fixed,
                  style="money", name="PaymentFeeFixed",
                  validate=number_check(0, 100),
                  help_text="The flat charge per transaction on top "
                            "of the percentage.")
    row = s.field(row, "Packaging per order",
                  cfg.packaging_cost_per_order, style="money",
                  name="PackagingPerOrder",
                  validate=number_check(0, 1000),
                  help_text="Box, tissue, label, filler. Small, "
                            "constant, and worth counting.")
    row = s.field(row, "Default shipping charge",
                  cfg.default_shipping_charge, style="money",
                  name="DefaultShipping",
                  validate=number_check(0, 100000))
    row = s.field(row, "Free shipping over", cfg.free_shipping_over,
                  style="money", name="FreeShippingOver",
                  validate=number_check(0, 1000000))
    row += 1

    # ------------------------------------------------------------ tax --
    row = s.section(row, "Tax", LAST_COL)
    row = s.note(
        row,
        "Whether you charge sales tax or VAT, at what rate, and "
        "whether a wholesale customer is exempt, are local "
        "questions with local answers. The switch ships OFF and the "
        "rate ships at zero, because a wrong tax number is worse "
        "than no tax number.",
        LAST_COL, height=38)
    row += 1
    row = s.field(row, "Charge tax on sales?",
                  "Yes" if cfg.tax_on_sales else "No", style="flag",
                  name="TaxOnSales", validate=list_from("YesNoList"))
    row = s.field(row, "Tax rate", cfg.tax_rate, style="percent",
                  name="TaxRate", validate=number_check(0, 1))
    row = s.field(row, "What it is called", cfg.tax_name,
                  name="TaxName", value_end=5, help_col=7,
                  help_text="Sales tax, VAT, GST - whatever appears "
                            "on your invoices.")
    row += 1

    # --------------------------------------------------------- alerts --
    row = s.section(row, "When to start worrying", LAST_COL)
    row = s.note(
        row,
        "The thresholds the alert columns use. Set them to suit the "
        "size of your business - a fifty pound write-off is a "
        "shrug for some sellers and a bad week for others.",
        LAST_COL, height=30)
    row += 1
    row = s.field(row, "Chase a late delivery after",
                  cfg.po_chase_days, style="int", name="POChaseDays",
                  validate=number_check(0, 90),
                  help_text="Days past the expected date.")
    row = s.field(row, "Flag an unshipped order after",
                  cfg.pick_warning_days, style="int",
                  name="PickWarningDays", validate=number_check(0, 90),
                  help_text="Days a confirmed wholesale order has "
                            "been sitting unpicked.")
    row = s.field(row, "Chase an unpaid invoice after",
                  cfg.overdue_chase_days, style="int",
                  name="ChaseDays", validate=number_check(0, 365))
    row = s.field(row, "A customer is dormant after", 120,
                  style="int", name="DormantDays",
                  validate=number_check(30, 999),
                  help_text="Days since a stockist last ordered.")
    row = s.field(row, "Query an invoice that is out by", 1.00,
                  style="money", name="PriceGapTolerance",
                  validate=number_check(0, 10000),
                  help_text="Difference between the supplier's "
                            "invoice and what you expected to pay.")
    row = s.field(row, "Query postage that overshoots by", 2.00,
                  style="money", name="ShipGapTolerance",
                  validate=number_check(0, 10000))
    row = s.field(row, "A large expense is over", 500.00,
                  style="money", name="LargeExpenseLevel",
                  validate=number_check(0, 1000000))
    row = s.field(row, "A large loss is over", 100.00, style="money",
                  name="LargeLossLevel",
                  validate=number_check(0, 1000000),
                  help_text="Used on returns and stock audits.")
    row = s.field(row, "Review a write-off over", 50.00,
                  style="money", name="WriteoffReview",
                  validate=number_check(0, 1000000))
    row += 1

    # --------------------------------------------------------- targets --
    row = s.section(row, "Targets", LAST_COL)
    row = s.field(row, "Monthly revenue target",
                  cfg.target_revenue_month, style="money",
                  name="MonthlyTarget",
                  validate=number_check(0, 100000000))
    row = s.field(row, "Yearly revenue target",
                  cfg.target_revenue_year, style="money",
                  name="YearlyTarget",
                  validate=number_check(0, 100000000))
    row = s.field(row, "Target stock value", cfg.target_stock_value,
                  style="money", name="TargetStockValue",
                  validate=number_check(0, 100000000),
                  help_text="How much you want tied up in stock. "
                            "More is not better.")
    row = s.field(row, "Target stock turns a year",
                  cfg.target_turnover, style="number1",
                  name="TargetTurns", validate=number_check(0, 100))
    row = s.field(row, "Target return rate", cfg.target_return_rate,
                  style="percent", name="TargetReturnRate",
                  validate=number_check(0, 1))
    row = s.field(row, "Target fill rate", cfg.target_fill_rate,
                  style="percent", name="TargetFillRate",
                  validate=number_check(0, 1),
                  help_text="The share of order lines you can ship "
                            "complete, first time.")
    row += 1

    # --------------------------------------------- business-wide totals --
    row = s.section(row, "Totals the workbook refers to", LAST_COL)
    row = s.note(
        row,
        "Calculated, not typed. Several sheets express a figure as "
        "a share of one of these, and keeping them in one place "
        "means the shares always add up to a hundred per cent.",
        LAST_COL, height=30)
    row += 1
    row = s.field(row, "Total stock value", None, style="money",
                  name="TotalStockValue", mode="calc",
                  formula="=SUM(ValuationStockValue)")
    row = s.field(row, "Total revenue", None, style="money",
                  name="TotalRevenue", mode="calc",
                  formula="=SUM(PerfRevenue)")
    row = s.field(row, "Wholesale revenue", None, style="money",
                  name="WholesaleRevenue", mode="calc",
                  formula="=SUMIF(OrderIsShipped,1,OrderTotal)")
    row = s.field(row, "Retail revenue", None, style="money",
                  name="RetailRevenue", mode="calc",
                  formula="=SUMIF(SaleIsSold,1,SaleGross)")
    row = s.field(row, "Total purchase spend", None, style="money",
                  name="TotalPurchaseSpend", mode="calc",
                  formula="=SUM(POTotal)")
    row += 1

    # ----------------------------------------------- payment terms table --
    row = s.section(row, "Payment terms", LAST_COL)
    row = s.note(
        row,
        "What each term means in days. Purchase orders and wholesale "
        "invoices work out their own due dates from this table, so "
        "if your 'Net 30' really means thirty days from the end of "
        "the month, change the number here.",
        LAST_COL, height=34)
    row += 1
    terms_head = row
    s.table_head(terms_head, 1, ["Term", "Days"], widths=[26, 14])
    first_term = terms_head + 1
    for index, (label, days) in enumerate(PAYMENT_TERMS):
        line = first_term + index
        s.write(line, 1, label, fmts.input("text"))
        s.write(line, 2, days, fmts.input("int"))
    last_term = first_term + len(PAYMENT_TERMS) - 1
    book.define_name(
        "PaymentTermsList",
        f"={sheet_ref}!$B${first_term + 1}:$B${last_term + 1}")
    book.define_name(
        "TermDays",
        f"={sheet_ref}!$C${first_term + 1}:$C${last_term + 1}")
    row = last_term + 2

    # ----------------------------------------------- movement sign table --
    row = s.section(row, "What each movement type does to stock",
                    LAST_COL)
    row = s.note(
        row,
        "This is the table that gives a stock movement its sign. "
        "Plus one adds to stock, minus one takes away, and zero "
        "means the row carries its own sign - which is why an "
        "Adjustment is the only type where a negative quantity "
        "makes sense.",
        LAST_COL, height=38)
    row += 1
    move_head = row
    s.table_head(move_head, 1, ["Movement type", "Sign",
                                "What it means"],
                 widths=[24, 10, 46])
    first_move = move_head + 1
    for index, (label, sign, meaning) in enumerate(MOVEMENT_TYPES):
        line = first_move + index
        s.write(line, 1, label, fmts.input("text"))
        s.write(line, 2, sign, fmts.input("int"))
        s.write(line, 3, meaning, fmts.chrome("note"))
    last_move = first_move + len(MOVEMENT_TYPES) - 1
    book.define_name(
        "MovementSign",
        f"={sheet_ref}!$C${first_move + 1}:$C${last_move + 1}")
    row = last_move + 2

    # ------------------------------------------------- channel fee table --
    row = s.section(row, "What each sales channel charges", LAST_COL)
    row = s.note(
        row,
        "Every retail sale is costed with the rate for its channel. "
        "These are rough defaults and they go out of date the "
        "moment a platform changes its pricing - put your own in "
        "from a real statement, and your profit figures become "
        "trustworthy overnight.",
        LAST_COL, height=38)
    row += 1
    fee_head = row
    s.table_head(fee_head, 1, ["Channel", "Fee rate"],
                 widths=[24, 14])
    default_fees = {
        "Etsy": 0.095, "Shopify": 0.02, "Amazon": 0.15,
        "eBay": 0.129, "WooCommerce": 0.0,
        "Physical Store": 0.0, "Direct Sales": 0.0,
        "Craft Fair": 0.0, "Wholesale": 0.0, "Other": 0.0,
    }
    first_fee = fee_head + 1
    for index, channel in enumerate(SALES_CHANNELS):
        line = first_fee + index
        s.write(line, 1, channel, fmts.input("text"))
        s.write(line, 2, default_fees.get(channel, 0.0),
                fmts.input("percent"))
    last_fee = first_fee + len(SALES_CHANNELS) - 1
    book.define_name(
        "ChannelFeeRates",
        f"={sheet_ref}!$C${first_fee + 1}:$C${last_fee + 1}")
    row = last_fee + 2

    # ------------------------------------------------ price tier ladder --
    row = s.section(row, "Default quantity breaks", LAST_COL)
    row = s.note(
        row,
        "The quantity breaks new products start with on the "
        "Wholesale Pricing tab. Ten, fifty and a hundred suit a lot "
        "of small makers; change them to whatever your trade "
        "actually orders in.",
        LAST_COL, height=32)
    row += 1
    tier_head = row
    s.table_head(tier_head, 1, ["Break", "From this quantity"],
                 widths=[24, 20])
    first_tier = tier_head + 1
    for index, (qty, label) in enumerate(PRICE_TIERS):
        line = first_tier + index
        s.write(line, 1, label, fmts.input("text"))
        s.write(line, 2, qty, fmts.input("int"))
    last_tier = first_tier + len(PRICE_TIERS) - 1
    book.define_name(
        "TierBreaks",
        f"={sheet_ref}!$C${first_tier + 1}:$C${last_tier + 1}")
    row = last_tier + 2

    # ------------------------------------------ customer tier discounts --
    row = s.section(row, "Customer tiers", LAST_COL)
    row = s.note(
        row,
        "What each kind of account pays, as a share of your retail "
        "price. These are a guide for setting your price list - the "
        "actual prices live on the Wholesale Pricing tab, where you "
        "can see the margin next to each one.",
        LAST_COL, height=34)
    row += 1
    ctier_head = row
    s.table_head(ctier_head, 1, ["Tier", "Share of retail",
                                 "Who they are"],
                 widths=[18, 16, 44])
    first_ctier = ctier_head + 1
    for index, (label, mult, who) in enumerate(CUSTOMER_TIERS):
        line = first_ctier + index
        s.write(line, 1, label, fmts.input("text"))
        s.write(line, 2, mult, fmts.input("percent0"))
        s.write(line, 3, who, fmts.chrome("note"))
    last_ctier = first_ctier + len(CUSTOMER_TIERS) - 1
    book.define_name(
        "TierMultipliers",
        f"={sheet_ref}!$C${first_ctier + 1}:$C${last_ctier + 1}")
    row = last_ctier + 2

    # ------------------------------------------- recurring frequencies --
    row = s.section(row, "How often a recurring cost repeats",
                    LAST_COL)
    row = s.note(
        row,
        "This is what turns any expense into a monthly equivalent, "
        "so a yearly insurance premium and a monthly subscription "
        "can sit in the same breakdown. 'One off' is zero because "
        "it is not a standing cost.",
        LAST_COL, height=32)
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
        "The letters in front of every automatic ID. These are the "
        "keys the workbook joins on - Product IDs, SKUs, Supplier "
        "IDs, Purchase Order numbers and Order IDs. Change the "
        "prefixes here if you want to, but do it before you have "
        "typed three hundred products, not after.",
        LAST_COL, height=38)
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
        "not delete a whole column: a drop-down pointing at a "
        "missing list stops working.",
        LAST_COL, height=32)
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
            s.write(top + 1 + offset, column, value,
                    fmts.input("text"))
        for offset in range(len(values), len(values) + spare):
            s.write(top + 1 + offset, column, None,
                    fmts.input("text"))
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
        "The months the reports pick from, in plain yyyy-mm text so "
        "that every monthly total in the workbook groups on exactly "
        "the same thing.",
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

    row = s.note(row, STOCK_NOTE, LAST_COL, height=56)
    row += 1
    row = s.note(row, COSTING_NOTE, LAST_COL, height=76)
    row += 1
    row = s.note(row, RESERVED_NOTE, LAST_COL, height=56)
    row += 1
    row = s.note(row, HONESTY_NOTE, LAST_COL, height=62)
    row += 1
    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=48)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=76)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s


def _month_keys(today: _dt.date, back: int = 18,
                ahead: int = 18) -> List[str]:
    """A window of yyyy-mm keys around today."""
    keys: List[str] = []
    year = today.year
    start = today.month - back
    for step in range(back + ahead + 1):
        m = start + step
        y = year + (m - 1) // 12
        m = (m - 1) % 12 + 1
        keys.append(f"{y:04d}-{m:02d}")
    return keys
````

### `inventory_wholesale_spreadsheet/sheets/dashboard.py`

*402 lines · 5 functions*

> `Dashboard` - the page that answers "what do I need to do today?"
>
> Everything here is worked out from the registers, so nothing is ever
> out of date and there is nothing to refresh. The questions are in the
> order somebody running stock actually asks them: what am I about to
> run out of, what is somebody waiting for, where is my money, what is
> my stock worth, and which products are actually earning their shelf.
>
> The stock panel at the top is the one that matters. Out of stock,
> critically low and reserved-but-not-shipped are three different kinds
> of problem and they need three different responses, so they get three
> different tiles rather than one vague "low stock" count.

| Defines | What it does |
|---------|--------------|
| `_month_key(n)` | The `yyyy-mm` key for n months before this one |
| `_by_rank(sheet, rank_key, field, n)` | The nth row of a register by one of its hidden rank columns |
| `_alert_count(sheet, prefix)` | How many rows on a register are shouting a given prefix |
| `_all_alerts(prefix)` | Total rows across the workbook whose alert starts with prefix |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dashboard'), `LAST_COL` (12), `ACTION_ROWS` (10), `LIST_ROWS` (8), `MONTHS` (6), `ALERT_SHEETS` (17 items).

````python
"""`Dashboard` - the page that answers "what do I need to do today?"

Everything here is worked out from the registers, so nothing is ever
out of date and there is nothing to refresh. The questions are in the
order somebody running stock actually asks them: what am I about to
run out of, what is somebody waiting for, where is my money, what is
my stock worth, and which products are actually earning their shelf.

The stock panel at the top is the one that matters. Out of stock,
critically low and reserved-but-not-shipped are three different kinds
of problem and they need three different responses, so they get three
different tiles rather than one vague "low stock" count.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 12
ACTION_ROWS = 10
LIST_ROWS = 8
MONTHS = 6


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


def _alert_count(sheet: str, prefix: str) -> str:
    """How many rows on a register are shouting a given prefix."""
    return f'COUNTIF([[{sheet}.alert]],"{prefix}*")'


#: every register that carries an alert column
ALERT_SHEETS = [
    "Product & SKU Database", "Inventory Tracker",
    "Stock Movement Log", "Supplier Database", "Purchase Orders",
    "Wholesale Customer CRM", "Wholesale Pricing", "Wholesale Orders",
    "Retail Sales Tracker", "Reorder Planner",
    "Inventory Valuation & COGS", "Returns & Damaged Goods",
    "Warehouse Locations", "Stock Audit", "Product Performance",
    "Expense Tracker", "Profit Analysis",
]


def _all_alerts(prefix: str) -> str:
    """Total rows across the workbook whose alert starts with prefix."""
    return "+".join(_alert_count(sheet, prefix)
                    for sheet in ALERT_SHEETS)


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary)

    s.banner("Inventory & Wholesale Dashboard",
             "Your stock, your orders and your money - worked out "
             "from the tabs you have already filled in.",
             LAST_COL)
    s.set_col(0, 0, 26)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 13)

    # ------------------------------------------------------------ head --
    row = 4
    s.merge(row, 0, 0, "Business", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    s.formula(row, 1, "=BusinessName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    s.formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")',
              fmts.calc("text"))
    s.merge(row, 6, 6, "SKUs", fmts.chrome("label_right"))
    s.formula(row, 7, "=SUM(ProductActive)", fmts.calc("int"))
    s.merge(row, 8, 8, "Today", fmts.chrome("label_right"))
    s.formula(row, 9, "=TodayDate", fmts.calc("date"))
    s.merge(row, 10, LAST_COL, None, fmts.chrome("note"))
    s.formula(
        row, 10,
        '="Figures update the moment you type. Nothing to refresh."',
        fmts.chrome("note"))
    row += 2

    # ----------------------------------------------------- stock panel --
    row = s.section(row, "Stock right now", LAST_COL)
    s.kpi(row, 0, "Out of stock", "=SUM(StockIsOut)", "int",
          note="Nothing available to sell")
    s.kpi(row, 2, "Below minimum", "=SUM(StockIsLow)", "int",
          note="Still sellable, but not for long")
    s.kpi(row, 4, "Needs reordering", "=SUM(ReorderNeeded)", "int",
          note="Lines the planner says to buy")
    s.kpi(row, 6, "Reserved units", "=SUM(StockReserved)", "int",
          note="Promised to confirmed orders, not yet shipped")
    s.kpi(row, 8, "Units on order", "=SUM(StockOnOrder)", "int",
          note="Ordered from suppliers, not arrived")
    row = s.kpi(row, 10, "Dead lines", "=SUM(StockIsDead)", "int",
                span=3, note="Holding stock with no recent sales")
    row += 1

    # ----------------------------------------------------- money panel --
    row = s.section(row, "The money", LAST_COL)
    s.kpi(row, 0, "Stock at cost", "=TotalStockValue", "money",
          note="Weighted average cost")
    s.kpi(row, 2, "Stock at retail", "=SUM(StockRetailValue)",
          "money", note="What it would fetch if it all sold")
    s.kpi(row, 4, "Owed to you", "=SUM(CustomerOwing)", "money",
          note="Invoiced wholesale, unpaid")
    s.kpi(row, 6, "Overdue", "=SUM(CustomerOverdue)", "money",
          note="Past its due date - chase these")
    s.kpi(row, 8, "Committed to suppliers",
          "=SUM(POOutstandingValue)", "money",
          note="Ordered and not yet delivered")
    row = s.kpi(row, 10, "To spend on reorders",
                "=SUM(ReorderValue)", "money", span=3,
                note="What the planner is suggesting")
    row += 1

    # --------------------------------------------------- trading panel --
    row = s.section(row, "Trading this month", LAST_COL)
    month_sale = ('SUMIFS(SaleGross,SaleMonths,ThisMonthKey,'
                  'SaleIsSold,1)')
    month_trade = ('SUMIFS(OrderTotal,OrderMonths,ThisMonthKey,'
                   'OrderIsShipped,1)')
    s.kpi(row, 0, "Revenue this month",
          f"={month_sale}+{month_trade}", "money",
          note="Retail and shipped wholesale")
    s.kpi(row, 2, "Retail", f"={month_sale}", "money")
    s.kpi(row, 4, "Wholesale", f"={month_trade}", "money")
    s.kpi(row, 6, "vs target",
          f'=IF(MonthlyTarget<=0,"",({month_sale}+{month_trade})/'
          f'MonthlyTarget)', "percent0",
          note="Against the Settings target")
    s.kpi(row, 8, "Gross margin",
          '=IFERROR(SUMIFS(ProfitGross,ProfitMonths,ThisMonthKey)/'
          'SUMIFS(ProfitRevenue,ProfitMonths,ThisMonthKey),"")',
          "percent0", note="This month, from Profit Analysis")
    row = s.kpi(row, 10, "Orders to pick",
                "=SUMPRODUCT(OrderIsReserved*OrderFirstLine)", "int",
                span=3, note="Confirmed and waiting to ship")
    row += 1

    # ------------------------------------------------------- integrity --
    row = s.section(row, "Things that would break the numbers",
                    LAST_COL)
    row = s.note(
        row,
        "These four counts should all be zero. A duplicate SKU means "
        "two products share one stock figure; a duplicate purchase "
        "order line means a delivery can be booked in twice; an "
        "unposted stock count means your shelf and your spreadsheet "
        "disagree and nobody has said why; negative stock means "
        "something was sold that was never received.",
        LAST_COL, height=44)
    row += 1
    s.kpi(row, 0, "Duplicate SKUs",
          "=SUMPRODUCT((ProductDupCount>1)*1)", "int",
          note="Two rows sharing one code")
    s.kpi(row, 2, "Duplicate PO lines",
          "=SUMPRODUCT((PODupLine>1)*1)", "int",
          note="Same PO, same SKU, twice")
    s.kpi(row, 4, "Counts not posted", "=SUM(AuditUnposted)", "int",
          note="Variance found, adjustment never written")
    s.kpi(row, 6, "Negative stock",
          "=SUMPRODUCT((StockCurrent<0)*1)", "int",
          note="Sold more than was ever received")
    s.kpi(row, 8, "Lines short of stock",
          "=SUMPRODUCT((OrderShortfall>0)*1)", "int",
          note="Promised more than you hold")
    row = s.kpi(row, 10, "Unclaimed supplier faults",
                "=SUM(ReturnUnclaimed)", "money", span=3,
                note="Their fault, no claim raised")
    row += 1

    # ------------------------------------------------- what needs doing --
    row = s.section(row, "What needs doing", LAST_COL)
    row = s.note(
        row,
        "Every register has a 'What needs doing' column. This is all "
        "of them added up, by how loudly they are shouting. Work "
        "down from the left.",
        LAST_COL, height=30)
    row += 1
    s.kpi(row, 0, "Urgent", f"={_all_alerts('Urgent -')}", "int",
          note="Something is wrong, not just late")
    s.kpi(row, 2, "Overdue", f"={_all_alerts('Overdue -')}", "int",
          note="Past a date that mattered")
    s.kpi(row, 4, "To order",
          f"={_all_alerts('Order -')}+{_all_alerts('Due -')}", "int",
          note="Stock decisions waiting")
    s.kpi(row, 6, "To chase or send",
          f"={_all_alerts('Chase -')}+{_all_alerts('Send -')}+"
          f"{_all_alerts('Follow up -')}", "int",
          note="Somebody is waiting on you")
    s.kpi(row, 8, "To check or add",
          f"={_all_alerts('Check -')}+{_all_alerts('Add -')}+"
          f"{_all_alerts('Action -')}+{_all_alerts('Missing -')}+"
          f"{_all_alerts('Duplicate -')}", "int",
          note="Housekeeping and gaps")
    row = s.kpi(row, 10, "All clear",
                f"={_all_alerts('OK')}", "int", span=3,
                note="Rows with nothing outstanding")
    row += 1

    # --------------------------------------------- reorder list (left) --
    top = row
    row = s.subsection(row, "Buy these first", 5)
    s.table_head(row, 0, ["SKU", "Product", "Available", "Order",
                          "Value", "By"],
                 widths=[15, 24, 11, 9, 11, 11])
    row += 1
    reorder_first = row
    for n in range(1, ACTION_ROWS + 1):
        line = reorder_first + n - 1
        for col, field, style in ((0, "sku", "id"),
                                  (1, "product", "text"),
                                  (2, "available", "int"),
                                  (3, "order_rounded", "int"),
                                  (4, "order_value", "money"),
                                  (5, "order_by", "date")):
            s.formula(
                line, col,
                f'=IFERROR(INDEX([[Reorder Planner.{field}]],'
                f'MATCH({n},[[Reorder Planner.buy_rank]],0)),"")',
                fmts.calc(style))

    # ------------------------------------------- orders to pick (right) --
    right = top
    right = s.subsection(right, "Orders waiting to go", LAST_COL,
                         first_col=7)
    s.table_head(right, 7, ["Order", "Customer", "SKU", "Qty",
                            "Short", "Age"],
                 widths=[12, 20, 15, 8, 9, 8])
    right += 1
    for n in range(1, ACTION_ROWS + 1):
        line = right + n - 1
        for col, field, style in ((7, "order_id", "id"),
                                  (8, "customer_name", "text"),
                                  (9, "sku", "id"),
                                  (10, "qty", "int"),
                                  (11, "shortfall", "int"),
                                  (12, "age_days", "int")):
            s.formula(
                line, col,
                f'=IFERROR(INDEX([[Wholesale Orders.{field}]],'
                f'MATCH({n},[[Wholesale Orders.pick_rank]],0)),"")',
                fmts.calc(style))
    row = max(reorder_first + ACTION_ROWS, right + ACTION_ROWS) + 1

    # ------------------------------------------------- best and worst --
    top = row
    row = s.subsection(row, "Earning their shelf", 5)
    s.table_head(row, 0, ["SKU", "Product", "Units", "Profit",
                          "GP%", "Return"],
                 widths=[15, 24, 9, 11, 9, 10])
    row += 1
    best_first = row
    for n in range(1, LIST_ROWS + 1):
        line = best_first + n - 1
        for col, field, style in ((0, "sku", "id"),
                                  (1, "product", "text"),
                                  (2, "units_sold", "int"),
                                  (3, "profit", "money"),
                                  (4, "gp_pct", "percent0"),
                                  (5, "return_on_stock", "percent0")):
            s.formula(line, col,
                      f"={_by_rank('Product Performance', 'profit_rank', field, n)}",
                      fmts.calc(style))

    right = top
    right = s.subsection(right, "Not earning it", LAST_COL,
                         first_col=7)
    s.table_head(right, 7, ["SKU", "Product", "Stock value",
                            "Days since", "State"],
                 widths=[15, 22, 12, 11, 12])
    right += 1
    for n in range(1, LIST_ROWS + 1):
        line = right + n - 1
        for col, field, style in ((7, "sku", "id"),
                                  (8, "product", "text"),
                                  (9, "stock_value", "money"),
                                  (10, "days_since_sale", "int"),
                                  (11, "band", "status")):
            s.formula(
                line, col,
                f'=IFERROR(INDEX([[Product Performance.{field}]],'
                f'MATCH({n},[[Product Performance.dead_rank]],0)),"")',
                fmts.calc(style))
    row = max(best_first + LIST_ROWS, right + LIST_ROWS) + 1

    # ----------------------------------------------- customers & money --
    top = row
    row = s.subsection(row, "Your best stockists", 5)
    s.table_head(row, 0, ["Customer", "Revenue", "Profit", "GP%",
                          "Owing", "Last order"],
                 widths=[22, 12, 11, 8, 11, 12])
    row += 1
    cust_first = row
    for n in range(1, LIST_ROWS + 1):
        line = cust_first + n - 1
        for col, field, style in ((0, "business", "text"),
                                  (1, "revenue", "money"),
                                  (2, "profit", "money"),
                                  (3, "gp_pct", "percent0"),
                                  (4, "owing", "money"),
                                  (5, "last_order", "date")):
            s.formula(
                line, col,
                f"={_by_rank('Wholesale Customer CRM', 'revenue_rank', field, n)}",
                fmts.calc(style))

    right = top
    right = s.subsection(right, "Deliveries you are waiting for",
                         LAST_COL, first_col=7)
    s.table_head(right, 7, ["PO", "Supplier", "SKU", "Due", "Late",
                            "Value"],
                 widths=[12, 18, 15, 11, 8, 11])
    right += 1
    for n in range(1, LIST_ROWS + 1):
        line = right + n - 1
        for col, field, style in ((7, "po_id", "id"),
                                  (8, "supplier_name", "text"),
                                  (9, "sku", "id"),
                                  (10, "expected", "date"),
                                  (11, "days_late", "int"),
                                  (12, "outstanding_value", "money")):
            s.formula(
                line, col,
                f'=IFERROR(INDEX([[Purchase Orders.{field}]],'
                f'MATCH({n},[[Purchase Orders.chase_rank]],0)),"")',
                fmts.calc(style))
    row = max(cust_first + LIST_ROWS, right + LIST_ROWS) + 1

    # ---------------------------------------------------- recent months --
    row = s.section(row, "The last six months", LAST_COL)
    s.table_head(row, 0, ["Month", "Revenue", "COGS", "Gross profit",
                          "Expenses", "Net profit", "Margin",
                          "Stock bought", "Cash effect"],
                 widths=[12, 12, 11, 12, 11, 12, 9, 12, 12])
    row += 1
    for n in range(MONTHS - 1, -1, -1):
        key = _month_key(n)
        s.formula(row, 0, f'={key}', fmts.calc("text"))
        for col, field, style in (
                (1, "ProfitRevenue", "money"),
                (2, "ProfitCOGS", "money"),
                (3, "ProfitGross", "money"),
                (4, "ProfitExpenses", "money"),
                (5, "ProfitNet", "money"),
        ):
            s.formula(row, col,
                      f'=SUMIFS({field},ProfitMonths,{key})',
                      fmts.calc(style))
        s.formula(
            row, 6,
            f'=IFERROR(SUMIFS(ProfitGross,ProfitMonths,{key})/'
            f'SUMIFS(ProfitRevenue,ProfitMonths,{key}),"")',
            fmts.calc("percent0"))
        s.formula(row, 7,
                  f'=SUMIFS(ProfitStockSpend,ProfitMonths,{key})',
                  fmts.calc("money"))
        s.formula(row, 8,
                  f'=SUMIFS(ProfitCash,ProfitMonths,{key})',
                  fmts.calc("money"))
        s.cf(f"F{row + 1}:F{row + 1}",
             {"type": "cell", "criteria": "<", "value": 0,
              "fmt": "bad"})
        row += 1
    row += 1

    # ------------------------------------------------------------ notes --
    row = s.note(
        row,
        "Everything on this page is calculated from the registers. "
        "There is no refresh button because there is nothing to "
        "refresh - type a sale on the Retail Sales Tracker and the "
        "stock, the margin and the reorder suggestion have already "
        "changed by the time you look back here.",
        LAST_COL, height=44)
    row += 1
    row = s.note(
        row,
        "Reserved is not the same as sold. Units on a confirmed "
        "wholesale order are still on your shelf and still counted "
        "in stock at cost - they simply are not available to promise "
        "to anybody else. They leave stock when the order ships.",
        LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `inventory_wholesale_spreadsheet/sheets/getting_started.py`

*297 lines · 1 function*

> `Getting Started` - twenty minutes, in the right order.
>
> Most inventory spreadsheets fail in the first hour, and almost always
> for the same reason: somebody starts typing products before deciding
> what a SKU means to them, and three weeks later there are four rows
> for the same mug and no stock figure anybody believes.
>
> So this page does not begin with "add your products". It begins with
> the two decisions that cannot be undone cheaply - what counts as a
> separate SKU, and what your opening stock really is - and only then
> lets you loose on the tabs.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Getting Started'), `LAST_COL` (9), `STEPS` (10 items), `DAILY` (4 items), `WEEKLY` (4 items), `MONTHLY` (4 items), `MISTAKES` (6 items).

````python
"""`Getting Started` - twenty minutes, in the right order.

Most inventory spreadsheets fail in the first hour, and almost always
for the same reason: somebody starts typing products before deciding
what a SKU means to them, and three weeks later there are four rows
for the same mug and no stock figure anybody believes.

So this page does not begin with "add your products". It begins with
the two decisions that cannot be undone cheaply - what counts as a
separate SKU, and what your opening stock really is - and only then
lets you loose on the tabs.
"""

from __future__ import annotations

from ..config import (COSTING_NOTE, DISCLAIMER, HONESTY_NOTE,
                      PRIVACY_NOTE, RESERVED_NOTE, STOCK_NOTE, Config,
                      Palette)
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Getting Started"
LAST_COL = 9

STEPS = [
    ("1. Decide what a SKU is, before you type anything",
     "A SKU is one thing you count separately. If you hold the "
     "350ml sage mug and the 250ml sage mug in different boxes and "
     "count them separately, they are two SKUs, not one product "
     "with a size option. The same goes for a six pack against a "
     "single, or a gift boxed version against a plain one. Get this "
     "wrong and no stock figure in the workbook can be right - and "
     "it is painful to fix after three hundred rows.",
     "Settings & Lists, then Product & SKU Database"),
    ("2. Set up your business and your rules",
     "Open Settings & Lists and work down it. Your name and "
     "address, your currency, your minimum acceptable margin, your "
     "lead times, and - this one matters more than it looks - what "
     "your sales channels actually charge you. The fee rates there "
     "decide whether the profit figures in this workbook are true "
     "or comforting.",
     "Settings & Lists"),
    ("3. Add your products and your opening stock",
     "One row per SKU on the Product & SKU Database. Fill in the "
     "cost, the retail price and the trade price. Then put in "
     "Opening stock and Opening value: how many you have right now "
     "and what they cost you in total. This is the only time you "
     "will ever type a stock quantity - from here on stock is "
     "calculated.",
     "Product & SKU Database"),
    ("4. Add your suppliers",
     "Who you buy from, their terms, and their quoted lead time. "
     "The lead time seeds the reorder planner until the workbook "
     "has enough of your own delivery history to know what they "
     "really take, at which point it quietly starts using that "
     "instead.",
     "Supplier Database"),
    ("5. Check the Inventory Tracker",
     "Go and look. Every product should have a line with your "
     "opening stock in it, calculated rather than typed. If a "
     "number is wrong, the fix is never to overwrite it - the fix "
     "is to correct the opening figure on the product record, "
     "because that is the only typed stock number in the workbook.",
     "Inventory Tracker"),
    ("6. Set your wholesale prices",
     "Wholesale Pricing already has a row per SKU. Put in your base "
     "trade price and up to three quantity breaks. Watch the GP% "
     "column next to each price rather than the price itself - that "
     "is the whole reason it is there.",
     "Wholesale Pricing"),
    ("7. Add your stockists",
     "Wholesale Customer CRM. The important cell is Price tier: "
     "every order they place prices itself off it, so you never "
     "quote by hand again. Set their payment terms and a credit "
     "limit while you are there.",
     "Wholesale Customer CRM"),
    ("8. Start trading",
     "Raise purchase orders when you buy, enter wholesale orders "
     "when a shop orders, and type or paste your retail sales in as "
     "they come. Move an order's status along as it progresses. "
     "That is the whole daily routine - everything else in the "
     "workbook is calculated from those three tabs.",
     "Purchase Orders, Wholesale Orders, Retail Sales Tracker"),
    ("9. Work off the Dashboard each morning",
     "Out of stock, below minimum, orders waiting to pick, "
     "deliveries running late, money overdue. Five minutes on the "
     "Dashboard replaces an hour of scrolling, and the four "
     "integrity counts in the middle should all read zero.",
     "Dashboard"),
    ("10. Count the shelf now and again",
     "Use Stock Audit. Count your highest value lines monthly and "
     "everything else once or twice a year. When a count disagrees "
     "with the system, write the difference on the Stock Movement "
     "Log as an adjustment with a reason - that is the step that "
     "actually moves the figure.",
     "Stock Audit, Stock Movement Log"),
]

DAILY = [
    "Enter yesterday's retail sales, or paste them from your "
    "marketplace export.",
    "Move any wholesale order that progressed to its new status - "
    "Confirmed, Packed, Shipped.",
    "Book in anything that arrived, by putting the quantity in the "
    "Qty received cell on its purchase order line.",
    "Glance at the Dashboard: anything Urgent or Overdue.",
]

WEEKLY = [
    "Work the Reorder Planner top down and raise the purchase "
    "orders it suggests.",
    "Chase late deliveries and overdue invoices from the Dashboard "
    "lists.",
    "Log any damage or loss on the Stock Movement Log while you "
    "still remember what happened.",
    "Check the four integrity counts on the Dashboard are all zero.",
]

MONTHLY = [
    "Cycle count your A lines - the few SKUs holding most of your "
    "money, listed on Inventory Valuation & COGS.",
    "Read the Monthly Report and compare the margin to last month.",
    "Look at Product Performance sorted by Return on stock, and "
    "decide what to clear.",
    "Review any price whose margin has drifted, on Wholesale "
    "Pricing.",
]

MISTAKES = [
    ("Typing over a stock figure",
     "There is nowhere to do it, and that is the single best thing "
     "about this workbook. Every quantity is derived from recorded "
     "events. If a figure is wrong, something is missing - a "
     "receipt, a sale, a breakage - and the fix is to record it."),
    ("Logging a purchase receipt on the Stock Movement Log",
     "Do not. Receipts belong in the Qty received cell on the "
     "purchase order line. The movement log is for damage, loss, "
     "gifts, production use, transfers and corrections only - "
     "anything else gets counted twice."),
    ("Two lines on one purchase order for the same SKU",
     "That is the one way to book the same delivery in twice. The "
     "Duplicate line column flags it in red. If a second shipment "
     "arrives against an existing line, update the received "
     "quantity to the running total instead."),
    ("Selling against In stock instead of Available",
     "In stock includes units already promised to a confirmed "
     "order. Available is what you can still sell. Quote from "
     "Available or you will sell the same pallet twice."),
    ("Typing stock purchases into the Expense Tracker",
     "They are already counted as cost of goods through the "
     "purchase orders. The workbook flags them and keeps them out "
     "of the operating cost total, but it is cleaner not to type "
     "them at all."),
    ("Leaving freight off a purchase order line",
     "Landed cost drives the weighted average, which drives every "
     "margin in the workbook. Carriage left off is margin you think "
     "you have and do not."),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.deep)

    s.banner("Getting Started",
             "Twenty minutes, in an order that will save you a "
             "fortnight.", LAST_COL)
    s.set_col(0, 0, 4)
    s.set_col(1, 1, 38)
    s.set_col(2, 7, 15)
    s.set_col(8, LAST_COL, 18)

    row = 4
    row = s.note(
        row,
        "This workbook has twenty-three tabs and you do not need "
        "most of them today. Do the ten steps below in order and "
        "you will have a working stock system this afternoon; the "
        "reporting tabs will have filled themselves in by the time "
        "you get round to looking at them.",
        LAST_COL, height=42)
    row += 1

    # ------------------------------------------------------ the rules --
    row = s.section(row, "Three things to know first", LAST_COL)
    row = s.note(row, STOCK_NOTE, LAST_COL, height=60)
    row += 1
    row = s.note(row, RESERVED_NOTE, LAST_COL, height=56)
    row += 1
    row = s.note(row, COSTING_NOTE, LAST_COL, height=80)
    row += 1

    # ---------------------------------------------------------- steps --
    row = s.section(row, "The ten steps", LAST_COL)
    for title, body, where in STEPS:
        s.merge(row, 0, 1, title, fmts.chrome("label_bold"))
        s.merge(row, 2, 7, body, fmts.chrome("note"))
        s.merge(row, 8, LAST_COL, where, fmts.chrome("legend_calc"))
        s.row_height(row, 54)
        row += 1
    row += 1

    # ------------------------------------------------------- routines --
    row = s.section(row, "The routine, once you are running", LAST_COL)
    s.merge(row, 0, 2, "Every day", fmts.chrome("header"))
    s.merge(row, 3, 6, "Every week", fmts.chrome("header"))
    s.merge(row, 7, LAST_COL, "Every month", fmts.chrome("header"))
    row += 1
    tallest = max(len(DAILY), len(WEEKLY), len(MONTHLY))
    for index in range(tallest):
        if index < len(DAILY):
            s.merge(row, 0, 2, f"- {DAILY[index]}",
                    fmts.chrome("note"))
        if index < len(WEEKLY):
            s.merge(row, 3, 6, f"- {WEEKLY[index]}",
                    fmts.chrome("note"))
        if index < len(MONTHLY):
            s.merge(row, 7, LAST_COL, f"- {MONTHLY[index]}",
                    fmts.chrome("note"))
        s.row_height(row, 34)
        row += 1
    row += 1

    # ------------------------------------------------------- mistakes --
    row = s.section(row, "Six ways to get it wrong", LAST_COL)
    row = s.note(
        row,
        "Every one of these has been designed against, so the "
        "workbook will usually stop you or shout at you. They are "
        "listed anyway, because understanding why a rule exists is "
        "how you remember it.",
        LAST_COL, height=32)
    row += 1
    for title, body in MISTAKES:
        s.merge(row, 0, 1, title, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, body, fmts.chrome("note"))
        s.row_height(row, 46)
        row += 1
    row += 1

    # ---------------------------------------------------- how it works --
    row = s.section(row, "How the tabs fit together", LAST_COL)
    chains = [
        ("Product -> SKU -> stock",
         "A product on the Product Database creates a SKU. The "
         "Inventory Tracker gives it a stock line automatically. "
         "Every other tab joins on that SKU."),
        ("Supplier -> purchase order -> receiving",
         "A supplier gets an ID. A purchase order line points at "
         "it. Putting a quantity in Qty received adds stock, "
         "updates the weighted average cost, and records how long "
         "that supplier really took."),
        ("Customer -> tier -> price -> order",
         "A stockist has a price tier. An order line reads the "
         "tier, finds the right quantity break on Wholesale "
         "Pricing, applies any account discount and shows you the "
         "margin before you confirm."),
        ("Movement -> reorder alert -> replenishment",
         "Sales and receipts move the stock figure. The Reorder "
         "Planner compares it against (daily sales x lead time) + "
         "safety stock and tells you what to buy and by when."),
        ("Sale -> COGS -> product profit",
         "A sale takes units out at weighted average cost. That "
         "cost becomes COGS on the valuation tab, which becomes "
         "profit per product on Product Performance, which becomes "
         "the business profit on Profit Analysis."),
    ]
    for title, body in chains:
        s.merge(row, 0, 1, title, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, body, fmts.chrome("note"))
        s.row_height(row, 40)
        row += 1
    row += 1

    # ---------------------------------------------------------- honest --
    row = s.section(row, "What this workbook is not", LAST_COL)
    row = s.note(row, HONESTY_NOTE, LAST_COL, height=62)
    row += 1
    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=48)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=76)
    row += 1

    row = s.section(row, "Protection", LAST_COL)
    row = s.note(
        row,
        f"Formula cells are locked so a stray paste cannot wipe out "
        f"a calculation, and every sheet is protected with the "
        f"password '{cfg.password}'. Input cells are left unlocked - "
        f"you can type in all of them without unprotecting anything. "
        f"If you do want to change a formula, unprotect the sheet in "
        f"Excel under Review, make the change, and protect it again.",
        LAST_COL, height=48)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

---

## 5. Product, SKU and stock

Workflow one, and the decision the rest of the product rests on. The catalogue mints a Product ID and flags a SKU entered twice. The Inventory Tracker holds no typed numbers at all: every figure on it is derived from the transaction registers, and reserved stock is kept separate from available stock so a confirmed order cannot be sold twice. The Stock Movement Log records the things that are not purchases, sales or returns - damage, loss, samples, production use, transfers - and keeps manual corrections in their own audited column.

### `inventory_wholesale_spreadsheet/sheets/products.py`

*259 lines · 2 functions*

> 4. Product & SKU Database - where a SKU is born.
>
> This is the foundation. Every other tab in the workbook matches on the
> SKU created here, so two rules apply and the sheet enforces both:
>
> * a SKU is unique, and a duplicate is flagged in red the moment it is
>   typed - a repeated SKU means two products share one stock figure,
>   which is the worst failure this workbook can have;
> * a variant that is stocked separately is a separate SKU. "Mug" is not
>   a product; "Mug, 350ml, Sage" is.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` | Priority order: broken data, then money, then housekeeping |

**Module constants:** `NAME` ('Product & SKU Database').

````python
"""4. Product & SKU Database - where a SKU is born.

This is the foundation. Every other tab in the workbook matches on the
SKU created here, so two rules apply and the sheet enforces both:

* a SKU is unique, and a duplicate is flagged in red the moment it is
  typed - a repeated SKU means two products share one stock figure,
  which is the worst failure this workbook can have;
* a variant that is stocked separately is a separate SKU. "Mug" is not
  a product; "Mug, 350ml, Sage" is.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, bar_cf, blank, id_col,
                     margin, money_cf, rank_desc, scale_cf, status_cf,
                     zero)
from ..writer import date_check, list_from, number_check

NAME = "Product & SKU Database"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        id_col("name", "ProductPrefix", digits=cfg.digits("product"),
               header="Product ID",
               note="Generated for you. The SKU is what the rest of "
                    "the workbook joins on, but a stable internal ID "
                    "is useful when a SKU has to be renamed."),
        Col("sku", "SKU", 15, "id",
            note="Your unique code for this exact product. One "
                 "stocked variant, one SKU. This is the join key for "
                 "the whole workbook - every other tab picks it from "
                 "a list."),
        Col("name", "Product name", 28, "text",
            note="What you call it on the shelf and on the invoice."),
        Col("description", "Description", 32, "wrap",
            note="The detail a picker needs to grab the right one."),
        Col("category", "Category", 14, "text",
            validate=list_from("CategoryList")),
        Col("subcategory", "Subcategory", 15, "text",
            validate=list_from("SubcategoryList")),
        Col("brand", "Brand", 14, "text"),
        Col("variant", "Variant", 14, "text",
            note="The thing that makes this different from its "
                 "siblings: 'Gift boxed', '6 pack', 'Seconds'."),
        Col("size", "Size", 10, "text"),
        Col("colour", "Colour", 12, "text"),
        Col("material", "Material", 14, "text"),
        Col("barcode", "Barcode", 15, "text",
            note="EAN, UPC or your own. Stored as text so a leading "
                 "zero survives."),
        Col("unit", "Unit", 11, "text",
            validate=list_from("UnitList"),
            note="What one unit of stock is. Sell tealights in packs "
                 "of six and the unit is 'Pack of 6', not 'Each'."),
        Col("image", "Image link", 18, "link",
            note="A link to the photo. The workbook stores the "
                 "address, not the picture - a hundred embedded "
                 "images would make this file unusable."),
        Col("supplier", "Supplier ID", 11, "id",
            validate=list_from("SupplierList"),
            note="Who you normally buy it from. Alternatives live on "
                 "the Supplier Database."),
        Col("supplier_sku", "Their code", 13, "text",
            note="The supplier's own code, so a purchase order makes "
                 "sense at their end."),

        # ---- money -----------------------------------------------
        Col("cost", "Unit cost", 11, "money",
            validate=number_check(0, 1000000),
            note="What one unit costs you to buy or make, landed. "
                 "This seeds the valuation; once you start receiving "
                 "purchase orders the weighted average takes over."),
        Col("retail_price", "Retail price", 12, "money",
            validate=number_check(0, 1000000)),
        Col("wholesale_price", "Wholesale", 11, "money",
            validate=number_check(0, 1000000),
            note="Your standard trade price. Quantity breaks live on "
                 "the Wholesale Pricing tab."),
        Col("min_price", "Floor price", 11, "money",
            validate=number_check(0, 1000000),
            note="The lowest you will go, ever. Any price below this "
                 "is flagged, wherever it appears."),
        Col("retail_margin", "Retail GP%", 10, "percent",
            formula=blank("sku", margin(
                "({retail_price}{r}-{cost}{r})",
                "{retail_price}{r}")),
            cf=scale_cf(),
            note="Gross margin at the retail price."),
        Col("wholesale_margin", "Trade GP%", 10, "percent",
            formula=blank("sku", margin(
                "({wholesale_price}{r}-{cost}{r})",
                "{wholesale_price}{r}")),
            cf=scale_cf()),
        Col("markup", "Markup", 9, "number1",
            formula=blank("sku", 'IF({cost}{r}<=0,"",'
                                 'ROUND({retail_price}{r}/'
                                 '{cost}{r},2))'),
            note="Retail price divided by unit cost. A different way "
                 "of saying the same thing as margin, and the one "
                 "most sellers price with."),

        # ---- stock levels ----------------------------------------
        Col("opening_qty", "Opening stock", 12, "int",
            validate=number_check(0, 1000000),
            note="What you had the day you started using this "
                 "workbook. Type it once and never again - every "
                 "change after this is a row on the Stock Movement "
                 "Log, a receipt, a sale or a return."),
        Col("opening_value", "Opening value", 12, "money",
            validate=number_check(0, 10000000),
            note="What that opening stock cost you in total. The "
                 "weighted average cost starts from this."),
        Col("min_stock", "Min level", 10, "int",
            validate=number_check(0, 1000000)),
        Col("max_stock", "Max level", 10, "int",
            validate=number_check(0, 1000000),
            note="The most you want on the shelf. The reorder "
                 "planner buys up to this, not beyond it."),
        Col("reorder_point", "Reorder at", 10, "int",
            validate=number_check(0, 1000000),
            note="Leave this blank and the Reorder Planner will work "
                 "one out from your actual sales and the supplier's "
                 "lead time. Fill it in and yours wins."),
        Col("status", "Status", 12, "status",
            validate=list_from("ProductStatusList"), cf=status_cf()),
        Col("launched", "Launched", 11, "date",
            validate=date_check(),
            note="When you started selling it. Used only for "
                 "reference."),
        Col("notes", "Notes", 30, "wrap"),

        # ---- derived ---------------------------------------------
        Col("variant_label", "Full name", 26, "text",
            formula=blank("sku", 'TRIM({name}{r}&'
                                 'IF({variant}{r}="",""," - "&'
                                 '{variant}{r})&'
                                 'IF({size}{r}="",""," "&{size}{r})&'
                                 'IF({colour}{r}="",""," "&'
                                 '{colour}{r}))'),
            note="Name, variant, size and colour in one string - what "
                 "a picking list should actually say.",
            hidden=True),
        Col("dup_count", "Dups", 7, "int",
            formula=blank("sku",
                          'COUNTIF({sku_first}:{sku_last},{sku}{r})'),
            hidden=True),
        Col("is_active", "Active", 7, "int",
            formula=zero("sku", 'IF({status}{r}="Active",1,0)'),
            hidden=True),

        # stock, pulled from the tracker so the catalogue can show it
        Col("on_hand", "In stock", 10, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.current]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)'),
            cf=bar_cf(),
            note="Live from the Inventory Tracker. You cannot type "
                 "here - stock is always calculated."),
        Col("available", "Available", 10, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.available]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)'),
            note="In stock, less what is already promised to a "
                 "confirmed order. This is the number to sell "
                 "against."),
        Col("stock_value", "Stock value", 12, "money",
            formula=zero("sku", 'ROUND({on_hand}{r}*'
                                '{avg_cost}{r},2)'),
            cf=money_cf()),
        Col("avg_cost", "Avg cost", 10, "money4",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Valuation & COGS.'
                                'avg_cost]],MATCH({sku}{r},'
                                '[[Inventory Valuation & COGS.sku]],'
                                '0))*1,{cost}{r})'),
            note="Weighted average cost. Falls back to the unit cost "
                 "you typed until the first purchase order is "
                 "received."),
        Col("units_sold", "Sold", 9, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Product Performance.units_sold]],'
                                'MATCH({sku}{r},'
                                '[[Product Performance.sku]],0))*1,0)'),
            hidden=True),
        Col("value_rank", "Rank", 8, "int",
            formula=rank_desc("stock_value", guard='{sku}{r}=""')
            .formula, hidden=True),

        Col("alert", "What needs doing", 30, "text", cf=alert_cf(),
            note="The one column to scan. Everything else on this "
                 "sheet is reference; this is the bit that asks you "
                 "for something."),
    ]

    alert = _alert()
    for col in cols:
        if col.key == "alert":
            col.formula = alert

    return SheetSpec(
        name=NAME,
        title="Product & SKU Database",
        subtitle=("One row per stocked thing. A variant you count "
                  "separately is a separate SKU - that is the whole "
                  "rule."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("products", cfg),
        help_text=("Create the SKU here first, then everything else "
                   "in the workbook can point at it. Stock, cost and "
                   "value on this sheet are calculated - the only "
                   "stock figure you ever type is the opening one."),
        extra_notes=[
            "A duplicate SKU is the one mistake this workbook cannot "
            "work around: two rows sharing a code means two products "
            "sharing one stock figure, one cost and one profit line. "
            "The Dups column counts them and the alert turns red.",
            "Set a floor price on anything you discount. Tiered "
            "wholesale pricing, customer discounts and a tired "
            "Friday afternoon can combine into a sale below cost "
            "without anybody noticing.",
        ],
    )


def _alert() -> str:
    """Priority order: broken data, then money, then housekeeping."""
    return alert_chain("sku", [
        ('{dup_count}{r}>1',
         "Urgent - duplicate SKU, two rows share this code"),
        ('{name}{r}=""', "Add - product name"),
        ('{cost}{r}<=0', "Add - unit cost, nothing can be costed "
                         "without it"),
        ('{retail_price}{r}<=0', "Add - retail price"),
        ('{retail_price}{r}<{cost}{r}',
         "Urgent - retail price is below cost"),
        ('AND({min_price}{r}>0,{wholesale_price}{r}>0,'
         '{wholesale_price}{r}<{min_price}{r})',
         "Urgent - trade price is below your own floor price"),
        ('{status}{r}=""', "Add - product status"),
        ('{status}{r}="Active"', [
            ('{available}{r}<=0',
             "Action - active product, nothing available to sell"),
            ('{min_stock}{r}<=0',
             "Check - no minimum stock level set"),
            ('{available}{r}<={min_stock}{r}',
             "Order - below minimum stock level"),
        ], "OK"),
        ('{on_hand}{r}>0',
         "Check - not active but still holding stock"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/inventory.py`

*285 lines · 2 functions*

> 5. Inventory Tracker - the answer to "how many have I got?".
>
> There is not one typeable cell on this sheet, and that is the point.
> Every row mirrors a row of the Product Database, and every quantity is
> the arithmetic of things that happened:
>
>     current  = opening + received - sold + returned - damaged
>                +/- adjustments
>     reserved = units on confirmed wholesale orders not yet shipped
>     available = current - reserved
>
> Available is the number that matters. Current stock is what is on the
> shelf; available is what you are still allowed to sell. A seller who
> promises against current stock sells the same pallet twice.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Inventory Tracker').

````python
"""5. Inventory Tracker - the answer to "how many have I got?".

There is not one typeable cell on this sheet, and that is the point.
Every row mirrors a row of the Product Database, and every quantity is
the arithmetic of things that happened:

    current  = opening + received - sold + returned - damaged
               +/- adjustments
    reserved = units on confirmed wholesale orders not yet shipped
    available = current - reserved

Available is the number that matters. Current stock is what is on the
shelf; available is what you are still allowed to sell. A seller who
promises against current stock sells the same pallet twice.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from .common import (adjustment_qty, alert_cf, alert_chain, bar_cf,
                     damaged_qty, lookup_product, lookup_product_num_zero,
                     received_qty, reserved_qty, returned_sellable_qty,
                     sold_qty, status_cf, weighted_average_cost, zero)

NAME = "Inventory Tracker"


def spec(cfg: Config) -> SheetSpec:
    sku = "{sku}{r}"

    cols = [
        Col("sku", "SKU", 15, "id",
            formula='=IF(INDEX([[Product & SKU Database.sku]],'
                    'ROW()-4)="","",'
                    'INDEX([[Product & SKU Database.sku]],ROW()-4))',
            note="Mirrors the Product Database row for row. Add a "
                 "product there and its stock line appears here by "
                 "itself - there is nothing to maintain on this tab."),
        Col("product", "Product", 26, "text",
            formula=lookup_product("name")),
        Col("category", "Category", 13, "text",
            formula=lookup_product("category")),
        Col("unit", "Unit", 10, "text",
            formula=lookup_product("unit")),
        Col("status", "Product status", 12, "status",
            formula=lookup_product("status"), cf=status_cf()),
        Col("location", "Location", 11, "text",
            formula='=IF({sku}{r}="","",IFERROR(INDEX('
                    '[[Warehouse Locations.location]],MATCH({sku}{r},'
                    '[[Warehouse Locations.sku]],0))&"",""))',
            note="The first bin this SKU is stored in. Split across "
                 "two shelves? The Warehouse Locations tab holds all "
                 "of them."),

        # ---- the arithmetic, one term per column ------------------
        Col("opening", "Opening", 9, "int",
            formula=lookup_product_num_zero("opening_qty"),
            note="What you declared on day one, straight from the "
                 "Product Database."),
        Col("received", "Received", 9, "int",
            formula=zero("sku", received_qty(sku)),
            note="Booked in against purchase orders. Receiving is "
                 "recorded once per order line, so the same delivery "
                 "cannot be counted twice."),
        Col("sold", "Sold", 9, "int",
            formula=zero("sku", sold_qty(sku)),
            note="Shipped wholesale orders plus completed retail "
                 "sales. A confirmed order that has not shipped is "
                 "not in here - it is in Reserved."),
        Col("returned", "Returned in", 10, "int",
            formula=zero("sku", returned_sellable_qty(sku)),
            note="Came back and went back on the shelf."),
        Col("damaged", "Damaged", 9, "int",
            formula=zero("sku", damaged_qty(sku)),
            note="Written off: returns that could not be resold, plus "
                 "breakage logged on the Stock Movement Log."),
        Col("adjusted", "Adjusted", 9, "int",
            formula=zero("sku", adjustment_qty(sku)),
            note="Manual corrections from the movement log, signed. A "
                 "stock count that found eleven fewer candles shows "
                 "as -11 here, and the reason is on that row."),

        Col("current", "In stock", 10, "int",
            formula=zero("sku", '{opening}{r}+{received}{r}-'
                                '{sold}{r}+{returned}{r}-'
                                '{damaged}{r}+{adjusted}{r}'),
            cf=bar_cf(),
            note="Opening + received - sold + returned - damaged "
                 "+/- adjustments. Physically on the shelf."),
        Col("reserved", "Reserved", 9, "int",
            formula=zero("sku", reserved_qty(sku)),
            note="Promised to a wholesale order that is confirmed, "
                 "being picked or packed, but not yet shipped. Still "
                 "in the building; no longer yours to sell."),
        Col("available", "Available", 10, "int",
            formula=zero("sku", '{current}{r}-{reserved}{r}'),
            cf=[{"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "<", "value": 5,
                 "fmt": "warn"}],
            note="In stock less reserved. THIS is the number to quote "
                 "a customer, and the number the reorder planner "
                 "works from."),
        Col("on_order", "On order", 9, "int",
            formula=zero("sku", 'SUMIFS([[Purchase Orders.qty]],'
                                '[[Purchase Orders.sku]],{sku}{r},'
                                '[[Purchase Orders.is_open]],1)-'
                                'SUMIFS([[Purchase Orders.received]],'
                                '[[Purchase Orders.sku]],{sku}{r},'
                                '[[Purchase Orders.is_open]],1)'),
            note="Ordered from a supplier, not arrived. Counted as "
                 "the outstanding balance of open purchase order "
                 "lines, so a part delivery leaves only the rest."),
        Col("next_po_row", "Next PO row", 9, "int",
            formula=zero("sku", 'SUMPRODUCT(MAX(('
                                '[[Purchase Orders.sku]]={sku}{r})*('
                                '[[Purchase Orders.is_open]]=1)*'
                                '[[Purchase Orders.seq_num]]))'),
            hidden=True),
        Col("incoming", "Next delivery", 12, "date",
            formula='=IF({sku}{r}="","",IF({next_po_row}{r}<=0,"",'
                    'IF(INDEX([[Purchase Orders.expected]],'
                    '{next_po_row}{r})="","",'
                    'INDEX([[Purchase Orders.expected]],'
                    '{next_po_row}{r}))))',
            note="The expected date on the most recent open purchase "
                 "order line for this SKU."),

        # ---- money ------------------------------------------------
        Col("avg_cost", "Avg cost", 10, "money4",
            formula=zero("sku", weighted_average_cost(sku)),
            note="Weighted average, the one costing method this "
                 "workbook uses. See the note at the bottom of "
                 "Inventory Valuation & COGS."),
        Col("stock_value", "Stock value", 12, "money",
            formula=zero("sku", 'ROUND({current}{r}*{avg_cost}{r},2)'),
            note="What is sitting on the shelf, at cost. Add this "
                 "column up and you have your inventory valuation."),
        Col("retail_value", "Retail value", 12, "money",
            formula=zero("sku", 'ROUND({current}{r}*'
                                'IFERROR(INDEX('
                                '[[Product & SKU Database.'
                                'retail_price]],MATCH({sku}{r},'
                                '[[Product & SKU Database.sku]],0))'
                                '*1,0),2)'),
            hidden=True),

        # ---- levels and cover -------------------------------------
        Col("min_stock", "Min", 7, "int",
            formula=lookup_product_num_zero("min_stock"), hidden=True),
        Col("max_stock", "Max", 7, "int",
            formula=lookup_product_num_zero("max_stock"), hidden=True),
        Col("daily_sales", "Sold/day", 9, "number1",
            formula=zero("sku", 'ROUND(SUMIFS('
                                '[[Retail Sales Tracker.qty]],'
                                '[[Retail Sales Tracker.sku]],'
                                '{sku}{r},'
                                '[[Retail Sales Tracker.is_sold]],1,'
                                '[[Retail Sales Tracker.date_num]],'
                                '">="&TodayDate-SalesWindow)/'
                                'SalesWindow+SUMIFS('
                                '[[Wholesale Orders.qty]],'
                                '[[Wholesale Orders.sku]],{sku}{r},'
                                '[[Wholesale Orders.is_shipped]],1,'
                                '[[Wholesale Orders.date_num]],'
                                '">="&TodayDate-SalesWindow)/'
                                'SalesWindow,2)'),
            note="Average units a day over the sales window set on "
                 "Settings - retail and shipped wholesale together."),
        Col("days_cover", "Days cover", 10, "int",
            formula='=IF({sku}{r}="","",IF({daily_sales}{r}<=0,"",'
                    'ROUND({available}{r}/{daily_sales}{r},0)))',
            cf=[{"type": "cell", "criteria": "<=", "value": 7,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 21,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">", "value": 120,
                 "fmt": "info"}],
            note="How long the available stock lasts at the current "
                 "rate of sale. Blank means nothing has sold lately, "
                 "which is its own kind of answer."),
        Col("stock_state", "Stock level", 13, "status",
            formula='=IF({sku}{r}="","",'
                    'IF({current}{r}<=0,"Out of stock",'
                    'IF({available}{r}<=0,"All reserved",'
                    'IF({min_stock}{r}>0,'
                    'IF({available}{r}<={min_stock}{r}*0.5,"Critical",'
                    'IF({available}{r}<={min_stock}{r},"Low",'
                    'IF({max_stock}{r}>0,'
                    'IF({current}{r}>{max_stock}{r},"Overstocked",'
                    '"OK"),"OK"))),'
                    'IF({available}{r}<5,"Low","OK")))))',
            cf=status_cf()),
        Col("last_move_num", "Last moved (n)", 9, "number",
            formula=zero("sku", 'SUMPRODUCT(MAX(('
                                '[[Stock Movement Log.sku]]='
                                '{sku}{r})*'
                                '[[Stock Movement Log.date_num]]))'),
            hidden=True),
        Col("last_movement", "Last moved", 11, "date",
            formula='=IF({sku}{r}="","",IF({last_move_num}{r}<=0,"",'
                    '{last_move_num}{r}))',
            note="The most recent dated row on the movement log for "
                 "this SKU."),

        # ---- hidden flags the dashboard counts --------------------
        Col("is_out", "Out", 6, "int",
            formula=zero("sku", 'IF({available}{r}<=0,1,0)'),
            hidden=True),
        Col("is_low", "Low", 6, "int",
            formula=zero("sku",
                         'IF({available}{r}<=0,0,'
                         'IF({min_stock}{r}<=0,0,'
                         'IF({available}{r}<={min_stock}{r},1,0)))'),
            hidden=True),
        Col("is_over", "Over", 6, "int",
            formula=zero("sku",
                         'IF({max_stock}{r}<=0,0,'
                         'IF({current}{r}>{max_stock}{r},1,0))'),
            hidden=True),
        Col("is_dead", "Dead", 6, "int",
            formula=zero("sku",
                         'IF({current}{r}<=0,0,'
                         'IF({daily_sales}{r}>0,0,1))'),
            hidden=True),

        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Inventory Tracker",
        subtitle=("Every quantity here is calculated. There is nothing "
                  "to type on this sheet - and no way to quietly "
                  "overwrite a stock figure."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        help_text=("One line per product, mirroring the Product "
                   "Database. To change a number here, add the row "
                   "that explains it: receive a purchase order, log a "
                   "sale, record a return, or write an adjustment on "
                   "the Stock Movement Log."),
        extra_notes=[
            "In stock and Available are different on purpose. A "
            "confirmed wholesale order reserves its units - they are "
            "still in the building, so In stock does not move, but "
            "they are spoken for, so Available does. Quote customers "
            "from Available.",
            "On order is the balance still owed on open purchase "
            "order lines, not the whole order. Receive half a pallet "
            "and only the other half is still counted as coming.",
        ],
    )


def _alert() -> str:
    return alert_chain("sku", [
        ('{status}{r}="Discontinued"', [
            ('{current}{r}>0',
             "Retired - discontinued, still holding stock, sell it "
             "through"),
        ], "OK"),
        ('{current}{r}<0',
         "Urgent - negative stock, something was sold that was never "
         "received"),
        ('{current}{r}<{reserved}{r}',
         "Urgent - more reserved than you physically hold"),
        ('AND({available}{r}<=0,{on_order}{r}>0)',
         "Due - nothing available, delivery already on order"),
        ('{available}{r}<=0',
         "Order - out of stock and nothing on order"),
        ('AND({min_stock}{r}>0,{available}{r}<={min_stock}{r}*0.5,'
         '{on_order}{r}<=0)',
         "Urgent - critically low, nothing on order"),
        ('AND({min_stock}{r}>0,{available}{r}<={min_stock}{r},'
         '{on_order}{r}<=0)',
         "Order - below minimum level"),
        ('{is_dead}{r}=1',
         "Check - holding stock with no sales in the window"),
        ('{is_over}{r}=1',
         "Check - above your maximum stock level, cash is tied up"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/movements.py`

*195 lines · 2 functions*

> 6. Stock Movement Log - the audit trail, and the only way to adjust.
>
> Every change to stock that is not a purchase receipt, a sale or a
> return gets written here, with a reason. This is deliberately the only
> place a human can push a stock figure up or down by hand, and even
> here you cannot do it anonymously: the row carries a type, a signed
> quantity, a reason and your name.
>
> The log is also where damage and loss are recorded, so "where did
> eleven candles go" always has an answer.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Stock Movement Log').

````python
"""6. Stock Movement Log - the audit trail, and the only way to adjust.

Every change to stock that is not a purchase receipt, a sale or a
return gets written here, with a reason. This is deliberately the only
place a human can push a stock figure up or down by hand, and even
here you cannot do it anonymously: the row carries a type, a signed
quantity, a reason and your name.

The log is also where damage and loss are recorded, so "where did
eleven candles go" always has an answer.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, blank, date_num_col,
                     id_col, lookup_product_num_zero,
                     month_col, product_name_col, seq_num_col,
                     sku_col, status_cf, zero)
from ..writer import date_check, list_from, number_check

NAME = "Stock Movement Log"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        id_col("sku", "MovementPrefix", digits=cfg.digits("movement"),
               header="Movement ID"),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="When the stock actually moved, not when you got "
                 "round to typing it."),
        sku_col(header="SKU",
                note="Pick the product. Every movement is against one "
                     "SKU - split a mixed delivery across rows."),
        product_name_col(),
        Col("type", "Movement type", 16, "text",
            validate=list_from("MovementTypeList"),
            note="What happened. The type decides the sign: a "
                 "purchase receipt adds, damage removes, an "
                 "adjustment carries whatever sign you type."),
        Col("qty", "Quantity", 10, "int",
            validate=number_check(-1000000, 1000000),
            note="Always a positive number, except on an Adjustment "
                 "row where a negative is how you say 'we have fewer "
                 "than the system thinks'."),
        Col("direction", "In / out", 9, "int",
            formula=zero("sku", 'IFERROR(INDEX(MovementSign,'
                                'MATCH({type}{r},MovementTypeList,'
                                '0)),0)'),
            hidden=True),
        Col("signed_qty", "Change", 9, "int",
            formula=zero("sku", 'IF({direction}{r}=0,{qty}{r},'
                                '{direction}{r}*ABS({qty}{r}))'),
            note="The quantity with its sign applied. An adjustment "
                 "keeps the sign you typed; everything else gets the "
                 "sign its movement type implies.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}]),
        Col("reason", "Reason", 30, "wrap",
            note="Required on an adjustment or a write-off. In six "
                 "months this sentence is the only thing standing "
                 "between you and a mystery."),
        Col("reference", "Reference", 13, "text",
            note="The PO, order number, audit ID or invoice this "
                 "relates to, if any."),
        Col("location", "Location", 11, "text",
            validate=list_from("LocationList")),
        Col("to_location", "Moved to", 11, "text",
            validate=list_from("LocationList"),
            note="Only on a transfer. Leave blank otherwise."),
        Col("recorded_by", "Recorded by", 13, "text",
            validate=list_from("StaffList"),
            note="Who made the change. Not bureaucracy - the person "
                 "who wrote the row is usually the only one who "
                 "remembers why."),
        Col("cost_each", "Cost each", 10, "money",
            formula=blank("sku", 'IFERROR(INDEX('
                                 '[[Inventory Tracker.avg_cost]],'
                                 'MATCH({sku}{r},'
                                 '[[Inventory Tracker.sku]],0))*1,'
                                 f'{lookup_product_num_zero("cost")[1:]})'),
            note="Weighted average cost at the time you look, used "
                 "to value what moved."),
        Col("value", "Value", 11, "money",
            formula=zero("sku",
                         'ROUND(ABS({signed_qty}{r})*'
                         '{cost_each}{r},2)'),
            note="What this movement was worth at cost. A write-off "
                 "row is a real loss and this is its size."),

        # ---- the three buckets the Inventory Tracker reads --------
        Col("damaged", "Damaged qty", 10, "int",
            formula=zero("sku",
                         'IF({type}{r}="Damaged",ABS({qty}{r}),'
                         'IF({type}{r}="Lost or stolen",'
                         'ABS({qty}{r}),0))'),
            hidden=True,
            note="Damage and loss, summed by the Inventory Tracker."),
        Col("adjustment", "Adjustment qty", 11, "int",
            formula=zero("sku",
                         'IF({type}{r}="Adjustment",{qty}{r},'
                         'IF({type}{r}="Found in count",'
                         'ABS({qty}{r}),'
                         'IF({type}{r}="Sample or gift",'
                         '-ABS({qty}{r}),'
                         'IF({type}{r}="Used in production",'
                         '-ABS({qty}{r}),'
                         'IF({type}{r}="Supplier return",'
                         '-ABS({qty}{r}),0)))))'),
            hidden=True,
            note="Everything that is a correction rather than a "
                 "trade. Receipts, sales and returns are NOT in "
                 "here - they are counted from their own tabs, and "
                 "counting them twice is the classic way a stock "
                 "spreadsheet drifts."),
        Col("is_writeoff", "Write-off", 8, "int",
            formula=zero("sku", 'IF({damaged}{r}>0,1,0)'),
            hidden=True),
        Col("is_manual", "Manual", 7, "int",
            formula=zero("sku",
                         'IF({type}{r}="Adjustment",1,'
                         'IF({type}{r}="Found in count",1,0))'),
            hidden=True),

        month_col("date"),
        date_num_col("date", "date_num"),
        seq_num_col("sku"),

        Col("stock_after", "Stock after", 11, "int",
            formula=blank("sku", 'IFERROR(INDEX('
                                 '[[Inventory Tracker.current]],'
                                 'MATCH({sku}{r},'
                                 '[[Inventory Tracker.sku]],0))*1,"")'),
            note="Current stock for this SKU as things stand now - "
                 "useful when you are reading the log to work out "
                 "where a number came from."),
        Col("flag", "Check", 14, "status",
            formula='=IF({sku}{r}="","",'
                    'IF({is_manual}{r}=1,"Manual",'
                    'IF({is_writeoff}{r}=1,"Write-off","Routine")))',
            cf=status_cf()),
        Col("alert", "What needs doing", 30, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Stock Movement Log",
        subtitle=("Every stock change that is not a purchase, a sale "
                  "or a return - with a reason attached to it."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("movements", cfg),
        help_text=("This is the only place in the workbook where you "
                   "can move stock by hand, and the row will not let "
                   "you do it silently: pick a type, sign the "
                   "quantity, say why."),
        extra_notes=[
            "Do not log purchase receipts, wholesale shipments or "
            "retail sales here. Those are counted from the Purchase "
            "Orders, Wholesale Orders and Retail Sales tabs, and a "
            "duplicate row here would double count them. Use this "
            "log for damage, loss, gifts, production use, transfers "
            "and corrections.",
            "Adjustment is the one type that takes the sign you "
            "type. If a count says you have eleven fewer, enter -11 "
            "and write the reason. Everything else takes its sign "
            "from the movement type.",
        ],
    )


def _alert() -> str:
    return alert_chain("sku", [
        ('{date}{r}=""', "Add - movement date"),
        ('{type}{r}=""', "Add - movement type"),
        ('{qty}{r}=0', "Add - quantity"),
        ('AND({is_manual}{r}=1,{reason}{r}="")',
         "Urgent - manual adjustment with no reason given"),
        ('AND({is_writeoff}{r}=1,{reason}{r}="")',
         "Action - write-off needs a reason"),
        ('{recorded_by}{r}=""', "Add - who recorded this"),
        ('AND({type}{r}="Transfer out",{to_location}{r}="")',
         "Add - where it was transferred to"),
        ('AND({direction}{r}=1,{qty}{r}<0)',
         "Check - an incoming movement with a negative quantity"),
        ('{date}{r}>TodayDate', "Check - dated in the future"),
        ('AND({is_writeoff}{r}=1,{value}{r}>WriteoffReview)',
         "Action - large write-off, worth a second look"),
    ])
````

---

## 6. Supplier, purchase order, receiving

Workflow two. Suppliers are scored on the lead time they actually deliver against the one they quote. Purchase orders carry freight into a landed cost per unit, track part deliveries line by line, compare the invoice to the order, and refuse to let the same PO line be received twice without saying so.

### `inventory_wholesale_spreadsheet/sheets/suppliers.py`

*239 lines · 2 functions*

> 7. Supplier Database - who you buy from, and how well it goes.
>
> Contact details are the boring half. The useful half is the scorecard
> at the right hand end, which is calculated from your own purchase
> order history: what you have actually spent with them, how many
> deliveries arrived late, and what their real lead time has been as
> opposed to the one they quoted you.
>
> A supplier who promises seven days and takes fourteen is not a
> reliability problem you can fix, but it is a number you can order
> around.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Supplier Database').

````python
"""7. Supplier Database - who you buy from, and how well it goes.

Contact details are the boring half. The useful half is the scorecard
at the right hand end, which is calculated from your own purchase
order history: what you have actually spent with them, how many
deliveries arrived late, and what their real lead time has been as
opposed to the one they quoted you.

A supplier who promises seven days and takes fourteen is not a
reliability problem you can fix, but it is a number you can order
around.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, blank, id_col, margin,
                     money_cf, rating_cf, scale_cf, score_validation,
                     status_cf, yesno, zero)
from ..writer import date_check, list_from, number_check

NAME = "Supplier Database"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        id_col("name", "SupplierPrefix", digits=cfg.digits("supplier"),
               header="Supplier ID",
               note="Generated. This is what products and purchase "
                    "orders point at."),
        Col("name", "Supplier name", 24, "text"),
        Col("contact", "Contact", 18, "text"),
        Col("email", "Email", 24, "text"),
        Col("phone", "Phone", 15, "text"),
        Col("website", "Website", 20, "link"),
        Col("address", "Address", 28, "wrap"),
        Col("city", "City", 14, "text"),
        Col("country", "Country", 13, "text"),
        Col("type", "Supplier type", 14, "text",
            validate=list_from("SupplierTypeList")),
        Col("supplies", "What they supply", 22, "wrap",
            note="Plain words. 'Wax, wicks, jars' beats a category "
                 "code when you are trying to remember who sells "
                 "lids."),

        # ---- commercial terms -------------------------------------
        Col("terms", "Payment terms", 13, "text",
            validate=list_from("PaymentTermsList")),
        Col("currency", "Currency", 9, "text",
            validate=list_from("CurrencyList"),
            note="What they invoice in. Costs elsewhere in the "
                 "workbook are in your own currency - convert when "
                 "you enter the purchase order."),
        Col("lead_time", "Lead time (days)", 13, "int",
            validate=number_check(0, 365),
            note="What they promise. The scorecard works out what "
                 "they actually do."),
        Col("moq", "Min order", 10, "money",
            validate=number_check(0, 1000000),
            note="Minimum order value, if they have one."),
        Col("free_ship_over", "Free ship over", 12, "money",
            validate=number_check(0, 1000000)),
        Col("discount", "Standard discount", 13, "percent0",
            validate=number_check(0, 1),
            note="Any trade discount they give you as standard."),
        Col("account_no", "Account no.", 13, "text"),
        Col("tax_id", "Tax / VAT no.", 14, "text"),

        # ---- relationship -----------------------------------------
        Col("status", "Status", 12, "status",
            validate=list_from("SupplierStatusList"), cf=status_cf()),
        Col("preferred", "Preferred", 10, "flag",
            validate=yesno("Your first choice for these goods."),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("rating", "Rating", 9, "rating",
            validate=score_validation(5), cf=rating_cf(),
            note="Your own score out of five. Quality, "
                 "communication, how they handle a problem."),
        Col("first_order", "First order", 11, "date",
            validate=date_check()),
        Col("notes", "Notes", 30, "wrap"),

        # ---- scorecard, all calculated ----------------------------
        Col("orders", "Orders", 8, "int",
            formula=zero("name", 'COUNTIFS('
                                 '[[Purchase Orders.supplier]],'
                                 '{id}{r})'),
            note="Purchase order lines placed with them."),
        Col("spend", "Total spend", 12, "money",
            formula=zero("name", 'SUMIFS('
                                 '[[Purchase Orders.total]],'
                                 '[[Purchase Orders.supplier]],'
                                 '{id}{r})'),
            cf=money_cf()),
        Col("spend_year", "Spend this year", 13, "money",
            formula=zero("name", 'SUMIFS('
                                 '[[Purchase Orders.total]],'
                                 '[[Purchase Orders.supplier]],'
                                 '{id}{r},'
                                 '[[Purchase Orders.order_num]],'
                                 '">="&YearStart,'
                                 '[[Purchase Orders.order_num]],'
                                 '"<="&YearEnd)'),
            note="Calendar year to date, on the order date."),
        Col("open_orders", "Open", 7, "int",
            formula=zero("name", 'COUNTIFS('
                                 '[[Purchase Orders.supplier]],'
                                 '{id}{r},'
                                 '[[Purchase Orders.is_open]],1)'),
            note="Ordered and not yet fully received."),
        Col("open_value", "Open value", 11, "money",
            formula=zero("name", 'SUMIFS('
                                 '[[Purchase Orders.outstanding]],'
                                 '[[Purchase Orders.supplier]],'
                                 '{id}{r},'
                                 '[[Purchase Orders.is_open]],1)'),
            note="Money committed to goods that have not arrived."),
        Col("late_orders", "Late", 7, "int",
            formula=zero("name", 'COUNTIFS('
                                 '[[Purchase Orders.supplier]],'
                                 '{id}{r},'
                                 '[[Purchase Orders.is_late]],1)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}]),
        Col("on_time_rate", "On time", 9, "percent0",
            formula='=IF({name}{r}="","",IF({orders}{r}<=0,"",'
                    'ROUND(1-{late_orders}{r}/{orders}{r},4)))',
            cf=scale_cf(),
            note="Share of their order lines that were not late. "
                 "Below 80% and you are the one absorbing it."),
        Col("actual_lead", "Actual lead", 11, "number1",
            formula='=IF({name}{r}="","",'
                    'IF(COUNTIFS([[Purchase Orders.supplier]],'
                    '{id}{r},[[Purchase Orders.lead_days]],">0")<=0,'
                    '"",ROUND(SUMIFS('
                    '[[Purchase Orders.lead_days]],'
                    '[[Purchase Orders.supplier]],{id}{r},'
                    '[[Purchase Orders.lead_days]],">0")/'
                    'COUNTIFS([[Purchase Orders.supplier]],{id}{r},'
                    '[[Purchase Orders.lead_days]],">0"),1)))',
            note="Average days from order to delivery, measured off "
                 "your own received purchase orders."),
        Col("lead_gap", "Lead gap", 9, "number1",
            formula='=IF({name}{r}="","",IF({actual_lead}{r}="","",'
                    'ROUND({actual_lead}{r}-{lead_time}{r},1)))',
            cf=[{"type": "cell", "criteria": ">", "value": 3,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}],
            note="How many days longer they really take than they "
                 "said. Positive is bad. Add it to the quoted lead "
                 "time when you set reorder points."),
        Col("last_order", "Last order", 11, "date",
            formula='=IF({name}{r}="","",'
                    'IF({last_order_num}{r}<=0,"",'
                    '{last_order_num}{r}))'),
        Col("last_order_num", "Last order (n)", 9, "number",
            formula=zero("name", 'SUMPRODUCT(MAX(('
                                 '[[Purchase Orders.supplier]]='
                                 '{id}{r})*'
                                 '[[Purchase Orders.order_num]]))'),
            hidden=True),
        Col("days_since", "Days since", 10, "int",
            formula='=IF({name}{r}="","",'
                    'IF({last_order_num}{r}<=0,"",'
                    'TodayDate-{last_order_num}{r}))'),
        Col("products", "Products", 9, "int",
            formula=zero("name", 'COUNTIFS('
                                 '[[Product & SKU Database.'
                                 'supplier]],{id}{r})'),
            note="How many SKUs you buy from them. One supplier "
                 "holding thirty lines is a concentration risk."),
        Col("share", "Share of spend", 12, "percent0",
            formula=blank("name", margin("{spend}{r}",
                                         "TotalPurchaseSpend")),
            cf=scale_cf(), hidden=True),

        Col("alert", "What needs doing", 30, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Supplier Database",
        subtitle=("Contacts and terms on the left; how they actually "
                  "perform, measured from your own orders, on the "
                  "right."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("suppliers", cfg),
        help_text=("Add a supplier here before you raise a purchase "
                   "order - the Supplier ID is what the order points "
                   "at. The scorecard columns fill themselves in as "
                   "orders are received."),
        extra_notes=[
            "Lead gap is the column worth watching. A supplier who "
            "quotes seven days and takes fourteen is not necessarily "
            "a supplier to drop, but your reorder points should be "
            "built on fourteen.",
            "Keep retired suppliers rather than deleting them. Their "
            "purchase orders are part of last year's cost of goods, "
            "and a deleted Supplier ID leaves those orders pointing "
            "at nothing.",
        ],
    )


def _alert() -> str:
    return alert_chain("name", [
        ('{email}{r}=""', "Add - email address"),
        ('{terms}{r}=""', "Add - payment terms"),
        ('{lead_time}{r}<=0', "Add - lead time, the reorder planner "
                              "needs it"),
        ('{status}{r}=""', "Add - supplier status"),
        ('{status}{r}="Inactive"', [
            ('{open_orders}{r}>0',
             "Check - marked inactive with orders still open"),
        ], "OK"),
        ('AND({lead_gap}{r}<>"",{lead_gap}{r}>7)',
         "Urgent - deliveries run over a week late, fix the lead "
         "time"),
        ('AND({orders}{r}>=3,{on_time_rate}{r}<>"",'
         '{on_time_rate}{r}<0.8)',
         "Action - late more often than not, worth a conversation"),
        ('AND({lead_gap}{r}<>"",{lead_gap}{r}>2)',
         "Check - slower than quoted, raise the lead time"),
        ('{open_value}{r}>0', [
            ('{late_orders}{r}>0',
             "Chase - money committed to a late delivery"),
        ], "OK"),
        ('AND({products}{r}>0,{orders}{r}=0)',
         "Check - products assigned but nothing ever ordered"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/purchases.py`

*339 lines · 2 functions*

> 8. Purchase Orders - ordering, and booking the delivery in once.
>
> One row is one line of one purchase order: a PO number, a SKU, a
> quantity, a cost. A three-product order is three rows sharing a PO
> number.
>
> Receiving works the way it does here for one reason. There is exactly
> ONE cell per line where a received quantity can live, so booking the
> same delivery in twice is not something the workbook has to detect -
> it is something the shape of the sheet makes impossible. Type 40 into
> that cell twice and it still says 40. The only way to double count a
> delivery is to create a second line for the same PO and SKU, and the
> Duplicate line column flags that in red the moment you do.
>
> Part deliveries are normal: received 30 of 50 leaves 20 outstanding,
> the line stays open, and the Inventory Tracker counts 30 in stock and
> 20 still on order.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Purchase Orders').

````python
"""8. Purchase Orders - ordering, and booking the delivery in once.

One row is one line of one purchase order: a PO number, a SKU, a
quantity, a cost. A three-product order is three rows sharing a PO
number.

Receiving works the way it does here for one reason. There is exactly
ONE cell per line where a received quantity can live, so booking the
same delivery in twice is not something the workbook has to detect -
it is something the shape of the sheet makes impossible. Type 40 into
that cell twice and it still says 40. The only way to double count a
delivery is to create a second line for the same PO and SKU, and the
Duplicate line column flags that in red the moment you do.

Part deliveries are normal: received 30 of 50 leaves 20 outstanding,
the line stays open, and the Inventory Tracker counts 30 in stock and
20 still on order.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, blank, date_num_col,
                     is_committed_po, is_open_po, lookup_product,
                     lookup_product_num_zero, lookup_supplier,
                     month_col, seq_num_col,
                     sku_col, status_cf, yesno, zero)
from ..writer import date_check, list_from, number_check

NAME = "Purchase Orders"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        Col("po_id", "PO number", 12, "id",
            note="Your purchase order number. Repeat it on every "
                 "line of the same order - the printable PO on the "
                 "last tab gathers all the lines that share it."),
        Col("line_no", "Line", 6, "int",
            formula=blank("po_id", 'COUNTIFS({po_id_first}:'
                                   '{po_id}{r},{po_id}{r})'),
            note="Line number within this PO, worked out for you."),
        Col("supplier", "Supplier ID", 11, "id",
            validate=list_from("SupplierList")),
        Col("supplier_name", "Supplier", 20, "text",
            formula=lookup_supplier("name")),
        Col("order_date", "Ordered", 11, "date", validate=date_check(),
            note="The date you sent it, not the date you drafted it."),
        Col("expected", "Expected", 11, "date", validate=date_check(),
            note="Leave blank and the workbook fills in the "
                 "supplier's lead time for you."),
        sku_col(),
        Col("product", "Product", 24, "text",
            formula=lookup_product("name")),
        Col("supplier_sku", "Their code", 12, "text",
            formula=lookup_product("supplier_sku", missing='""'),
            note="The supplier's own code for it, pulled from the "
                 "product record so the order makes sense to them."),

        # ---- ordering ---------------------------------------------
        Col("qty", "Qty ordered", 11, "int",
            validate=number_check(0, 1000000)),
        Col("unit_cost", "Unit cost", 10, "money",
            validate=number_check(0, 1000000),
            note="What they are charging you per unit on this order. "
                 "Not the standing cost on the product record - what "
                 "this delivery actually cost."),
        Col("discount", "Discount", 9, "percent0",
            validate=number_check(0, 1)),
        Col("freight", "Freight", 10, "money",
            validate=number_check(0, 1000000),
            note="Shipping, duty and handling for THIS line. Spread "
                 "a single carriage charge across the lines and it "
                 "lands in the stock value where it belongs."),
        Col("goods_value", "Goods", 11, "money",
            formula=zero("po_id", 'ROUND({qty}{r}*{unit_cost}{r}*'
                                  '(1-{discount}{r}),2)')),
        Col("total", "Line total", 11, "money",
            formula=zero("po_id", 'ROUND({goods_value}{r}+'
                                  '{freight}{r},2)'),
            note="Goods plus freight - the landed cost of the line, "
                 "which is what inventory should be valued at."),
        Col("landed_each", "Landed each", 11, "money4",
            formula=zero("po_id", 'IF({qty}{r}<=0,{unit_cost}{r},'
                                  'ROUND({total}{r}/{qty}{r},4))'),
            note="Landed cost per unit. This is what feeds the "
                 "weighted average cost, which is why freight "
                 "belongs on the line."),

        # ---- receiving - one cell, once ---------------------------
        Col("received", "Qty received", 12, "int",
            validate=number_check(0, 1000000),
            note="THE receiving cell. One per line, so a delivery "
                 "cannot be booked in twice. Part delivery? Put what "
                 "arrived; the rest stays outstanding. Second "
                 "shipment? Update this number to the running total "
                 "received against this line."),
        Col("received_date", "Received", 11, "date",
            validate=date_check()),
        Col("received_by", "Checked by", 12, "text",
            validate=list_from("StaffList")),
        Col("condition", "Condition", 12, "text",
            validate=list_from("DeliveryConditionList"),
            note="How it turned up. Anything other than Good wants a "
                 "note and probably a claim."),
        Col("received_value", "Received value", 12, "money",
            formula=zero("po_id", 'ROUND({received}{r}*'
                                  '{landed_each}{r},2)'),
            note="What arrived, at landed cost. This is the number "
                 "the weighted average cost is built from."),
        Col("outstanding", "Still due", 10, "int",
            formula=zero("po_id", 'MAX(0,{qty}{r}-{received}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("outstanding_value", "Due value", 11, "money",
            formula=zero("po_id", 'ROUND({outstanding}{r}*'
                                  '{landed_each}{r},2)'),
            hidden=True),
        Col("status", "Status", 15, "status",
            validate=list_from("POStatusList"), cf=status_cf(),
            note="Set it yourself, or leave it and read the "
                 "Suggested status column, which works it out from "
                 "the quantities."),
        Col("suggested_status", "Suggested", 15, "text",
            formula='=IF({po_id}{r}="","",'
                    'IF({status}{r}="Cancelled","Cancelled",'
                    'IF({received}{r}<=0,'
                    'IF({order_date}{r}="","Draft","Sent"),'
                    'IF({received}{r}>={qty}{r},"Fully Received",'
                    '"Partially Received"))))',
            note="What the quantities say the status should be. "
                 "Where this disagrees with the Status column the "
                 "alert will tell you."),

        # ---- invoicing --------------------------------------------
        Col("invoice_no", "Invoice no.", 12, "text"),
        Col("invoice_amount", "Invoiced", 11, "money",
            validate=number_check(0, 1000000),
            note="What they actually billed. Where this disagrees "
                 "with the line total, the alert says so - price "
                 "creep on a reorder is the commonest way a margin "
                 "quietly disappears."),
        Col("paid", "Paid", 7, "flag", validate=yesno(),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "No", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("due_date", "Payment due", 11, "date",
            formula='=IF({po_id}{r}="","",'
                    'IF({received_date}{r}="","",'
                    '{received_date}{r}+IFERROR(INDEX(TermDays,'
                    'MATCH(IFERROR(INDEX('
                    '[[Supplier Database.terms]],MATCH({supplier}{r},'
                    '[[Supplier Database.id]],0)),""),'
                    'PaymentTermsList,0)),0)))',
            note="Received date plus the supplier's payment terms."),
        Col("notes", "Notes", 26, "wrap"),

        # ---- flags and timing -------------------------------------
        Col("dup_line", "Duplicate line", 12, "int",
            formula=zero("po_id", 'IF({sku}{r}="",0,'
                                  'COUNTIFS({po_id_first}:'
                                  '{po_id_last},{po_id}{r},'
                                  '{sku_first}:{sku_last},'
                                  '{sku}{r}))'),
            hidden=True,
            note="How many lines on this PO are for this SKU. More "
                 "than one and the same goods can be received twice."),
        Col("is_open", "Open", 6, "int",
            formula=zero("po_id",
                         f'IF({is_open_po("{status}{r}")},'
                         'IF({outstanding}{r}>0,1,0),0)'),
            hidden=True),
        Col("is_committed", "Committed", 9, "int",
            formula=zero("po_id",
                         f'IF({is_committed_po("{status}{r}")},1,0)'),
            hidden=True),
        Col("is_received", "Received?", 9, "int",
            formula=zero("po_id", 'IF({received}{r}>0,1,0)'),
            hidden=True),
        Col("is_over", "Over", 6, "int",
            formula=zero("po_id",
                         'IF({received}{r}>{qty}{r},1,0)'),
            hidden=True),
        Col("days_late", "Days late", 9, "int",
            formula='=IF({po_id}{r}="","",IF({expected}{r}="","",'
                    'IF({received_date}{r}<>"",'
                    '{received_date}{r}-{expected}{r},'
                    'IF({is_open}{r}=1,TodayDate-{expected}{r},""))))',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}],
            note="Against the expected date: how late the delivery "
                 "was, or how late it is running right now."),
        Col("is_late", "Late", 6, "int",
            formula=zero("po_id", 'IF({days_late}{r}="",0,'
                                  'IF({days_late}{r}>0,1,0))'),
            hidden=True),
        Col("lead_days", "Lead days", 9, "int",
            formula=zero("po_id", 'IF({received_date}{r}="",0,'
                                  'IF({order_date}{r}="",0,'
                                  'MAX(0,{received_date}{r}-'
                                  '{order_date}{r})))'),
            hidden=True,
            note="Order to delivery, measured. This is what builds "
                 "the supplier's real lead time."),
        Col("price_gap", "Price gap", 10, "money",
            formula='=IF({po_id}{r}="","",'
                    'IF({invoice_amount}{r}<=0,"",'
                    'ROUND({invoice_amount}{r}-{total}{r},2)))',
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": -0.004,
                 "fmt": "warn"}],
            note="Invoice against what you expected to pay."),
        Col("cost_drift", "Cost drift", 10, "percent0",
            formula='=IF({po_id}{r}="","",'
                    f'IF({lookup_product_num_zero("cost")[1:]}<=0,"",'
                    'IF({unit_cost}{r}<=0,"",'
                    'ROUND({unit_cost}{r}/'
                    f'{lookup_product_num_zero("cost")[1:]}-1,4))))',
            cf=[{"type": "cell", "criteria": ">", "value": 0.1,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.02,
                 "fmt": "warn"}],
            note="This order's unit cost against the standing cost "
                 "on the product record. A supplier putting prices "
                 "up 12% shows here before it shows in your margin."),

        month_col("order_date"),
        date_num_col("order_date", "order_num"),
        date_num_col("received_date", "received_num"),
        seq_num_col("po_id"),

        Col("po_print_rank", "Print line", 9, "int",
            formula=zero("po_id",
                         'IF({po_id}{r}<>PrintPO,0,{line_no}{r})'),
            hidden=True,
            note="Line number within the PO currently typed on the "
                 "Printable Orders & Invoices tab, and zero on every "
                 "other row. That is how the printable gathers the "
                 "lines of one order without an array formula."),
        Col("chase_key", "Chase key", 9, "number",
            formula=zero("po_id",
                         'IF({is_open}{r}=0,0,'
                         'IF({days_late}{r}="",1000000,'
                         'IF({days_late}{r}>0,'
                         '3000000+{days_late}{r}*1000,1000000))+'
                         'ROUND({outstanding_value}{r},0))'),
            hidden=True),
        Col("chase_rank", "Chase rank", 9, "int",
            formula=zero("po_id",
                         'IF({is_open}{r}=0,0,'
                         'COUNTIFS({chase_key_first}:'
                         '{chase_key_last},">"&{chase_key}{r})+'
                         'COUNTIFS({chase_key_first}:{chase_key}{r},'
                         '{chase_key}{r}))'),
            hidden=True,
            note="Open deliveries, late ones first, then by how "
                 "much money is tied up in them."),
        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    # expected date defaults to the supplier's quoted lead time
    for col in cols:
        if col.key == "expected":
            col.note += (" Shown blank here; the Suggested status and "
                         "late columns only work once it is set.")

    return SheetSpec(
        name=NAME,
        title="Purchase Orders",
        subtitle=("One row per order line. One receiving cell per "
                  "line - which is why the same delivery cannot be "
                  "booked in twice."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("purchases", cfg),
        help_text=("Raise the order, send it, and when the boxes "
                   "arrive put the quantity into Qty received. Stock "
                   "moves by itself, the weighted average cost "
                   "updates, and the supplier's lead time scorecard "
                   "takes note of how long they took."),
        extra_notes=[
            "If a second shipment arrives against the same line, "
            "update Qty received to the running total - 30 arrived, "
            "then 20 more, so the cell reads 50. Do not add a second "
            "line for the same PO and SKU: that is the one way to "
            "double count a delivery, and the Duplicate line check "
            "exists to catch it.",
            "Freight on the line is not optional bookkeeping. Landed "
            "cost per unit is what the weighted average is built "
            "from, so carriage left off the order is margin you "
            "think you have and do not.",
        ],
    )


def _alert() -> str:
    return alert_chain("po_id", [
        ('{dup_line}{r}>1',
         "Urgent - this PO already has a line for this SKU, the "
         "delivery can be counted twice"),
        ('{sku}{r}=""', "Add - which product this line is for"),
        ('{supplier}{r}=""', "Add - supplier"),
        ('{qty}{r}<=0', "Add - quantity"),
        ('{unit_cost}{r}<=0', "Add - unit cost"),
        ('{status}{r}="Cancelled"', [
            ('{received}{r}>0',
             "Check - cancelled but stock was received against it"),
        ], "OK"),
        ('{is_over}{r}=1',
         "Urgent - received more than was ordered, check the count"),
        ('AND({received}{r}>0,{received_date}{r}="")',
         "Add - date the delivery was received"),
        ('AND({received}{r}>0,{condition}{r}="")',
         "Check - how did the delivery arrive?"),
        # AND() evaluates every argument, so ABS() still meets the
        # blank price_gap of a continuation line even though the
        # invoice test in front of it is already false. Coerce it.
        ('AND({invoice_amount}{r}>0,ABS(IFERROR({price_gap}{r}*1,0))>'
         'PriceGapTolerance)',
         "Action - invoice does not match the order value"),
        ('AND({status}{r}<>"",{status}{r}<>{suggested_status}{r})',
         "Check - status does not match what has been received"),
        ('AND({is_open}{r}=1,{days_late}{r}<>"",{days_late}{r}>0)',
         "Chase - delivery is overdue"),
        ('AND({status}{r}="Draft",{order_date}{r}<>"")',
         "Send - drafted and dated but never sent"),
        ('AND({received}{r}>=({qty}{r}),{qty}{r}>0,{paid}{r}="No",'
         '{due_date}{r}<>"",{due_date}{r}<TodayDate)',
         "Overdue - supplier invoice is past its due date"),
        ('AND({received}{r}>0,{paid}{r}="")',
         "Check - received, but is it paid?"),
        ('{expected}{r}=""', "Add - expected delivery date"),
    ])
````

---

## 7. Wholesale customer, tiered price, fulfilment

Workflow three, and the half of this product its competitors do not have. A trade CRM with terms, credit limits and what each account is worth; price breaks you set yourself, checked against the weighted average cost and against your own floor price; and an order register that reserves stock on confirmation, tells you what you are short before you promise it, and ages the invoice afterwards.

### `inventory_wholesale_spreadsheet/sheets/customers.py`

*271 lines · 2 functions*

> 9. Wholesale Customer CRM - the shops that buy from you.
>
> A wholesale customer is not a retail buyer with a bigger basket. They
> have a tier, terms, a credit limit, a minimum order and a buyer who
> expects you to remember what they ordered last time. All of that sits
> here, and the Wholesale Orders tab reads the tier off this sheet to
> price their lines automatically.
>
> The columns on the right are the ones that decide whether a stockist
> is worth keeping: what they have spent, what they still owe, whether
> they pay on time, and how long it has been since they last ordered.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Wholesale Customer CRM').

````python
"""9. Wholesale Customer CRM - the shops that buy from you.

A wholesale customer is not a retail buyer with a bigger basket. They
have a tier, terms, a credit limit, a minimum order and a buyer who
expects you to remember what they ordered last time. All of that sits
here, and the Wholesale Orders tab reads the tier off this sheet to
price their lines automatically.

The columns on the right are the ones that decide whether a stockist
is worth keeping: what they have spent, what they still owe, whether
they pay on time, and how long it has been since they last ordered.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, blank, id_col, margin,
                     money_cf, owing_cf, rating_cf, scale_cf,
                     score_validation, status_cf, yesno, zero)
from ..writer import date_check, list_from, number_check

NAME = "Wholesale Customer CRM"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        id_col("business", "CustomerPrefix",
               digits=cfg.digits("customer"), header="Customer ID",
               note="Generated. Wholesale orders point at this."),
        Col("business", "Business name", 24, "text"),
        Col("contact", "Buyer", 18, "text",
            note="The person who actually places the order."),
        Col("email", "Email", 24, "text"),
        Col("phone", "Phone", 14, "text"),
        Col("website", "Website / shop", 20, "link"),
        Col("address", "Address", 26, "wrap"),
        Col("city", "City", 13, "text"),
        Col("region", "Region", 12, "text"),
        Col("postcode", "Postcode", 10, "text"),
        Col("country", "Country", 12, "text"),
        Col("type", "Business type", 14, "text",
            validate=list_from("CustomerTypeList"),
            note="Boutique, gift shop, online retailer, distributor. "
                 "Useful when you are deciding who to show a new "
                 "range to first."),

        # ---- trading terms ----------------------------------------
        Col("tier", "Price tier", 12, "text",
            validate=list_from("CustomerTierList"),
            note="Which price list they buy on. The Wholesale Orders "
                 "tab reads this to price their lines, so it is the "
                 "single most important cell on the row."),
        Col("terms", "Payment terms", 13, "text",
            validate=list_from("PaymentTermsList")),
        Col("credit_limit", "Credit limit", 12, "money",
            validate=number_check(0, 1000000),
            note="The most you are willing to be owed by them at "
                 "once. Leave at zero for prepay-only accounts."),
        Col("min_order", "Min order", 10, "money",
            validate=number_check(0, 1000000),
            note="Their minimum order value. Orders below it are "
                 "flagged on the order line."),
        Col("discount", "Extra discount", 12, "percent0",
            validate=number_check(0, 1),
            note="Anything on top of their tier price - an opening "
                 "order deal, a loyalty arrangement. Applied after "
                 "the tier price on every line."),
        Col("tax_exempt", "Tax exempt", 10, "flag",
            validate=yesno("Resellers often are. Check locally."),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "info"}]),
        Col("resale_no", "Resale cert. no.", 14, "text",
            note="Only if your jurisdiction requires you to hold it. "
                 "Nothing else about their identity belongs in a "
                 "spreadsheet."),
        Col("shipping_pref", "Shipping", 13, "text",
            validate=list_from("ShippingMethodList")),

        # ---- relationship -----------------------------------------
        Col("status", "Status", 12, "status",
            validate=list_from("CustomerStatusList"), cf=status_cf()),
        Col("rating", "Rating", 8, "rating",
            validate=score_validation(5), cf=rating_cf(),
            note="How good an account are they, all things "
                 "considered - volume, payment, how much hand "
                 "holding."),
        Col("first_order", "First order", 11, "date",
            validate=date_check()),
        Col("source", "How they found you", 16, "text",
            validate=list_from("LeadSourceList")),
        Col("notes", "Notes", 28, "wrap",
            note="What they stock, what sells for them, who else "
                 "they buy from. This is the column that makes the "
                 "next call easy."),

        # ---- calculated account health ----------------------------
        Col("orders", "Orders", 8, "int",
            formula=zero("business",
                         'IFERROR(SUMPRODUCT((('
                         '[[Wholesale Orders.customer]]={id}{r})*'
                         '[[Wholesale Orders.is_first_line]])),0)'),
            note="Counted as distinct orders, not order lines."),
        Col("units", "Units", 8, "int",
            formula=zero("business", 'SUMIFS('
                                     '[[Wholesale Orders.qty]],'
                                     '[[Wholesale Orders.customer]],'
                                     '{id}{r},'
                                     '[[Wholesale Orders.is_shipped]]'
                                     ',1)'),
            hidden=True),
        Col("revenue", "Total revenue", 13, "money",
            formula=zero("business", 'SUMIFS('
                                     '[[Wholesale Orders.'
                                     'line_total]],'
                                     '[[Wholesale Orders.customer]],'
                                     '{id}{r})'),
            cf=money_cf()),
        Col("revenue_year", "This year", 12, "money",
            formula=zero("business", 'SUMIFS('
                                     '[[Wholesale Orders.'
                                     'line_total]],'
                                     '[[Wholesale Orders.customer]],'
                                     '{id}{r},'
                                     '[[Wholesale Orders.date_num]],'
                                     '">="&YearStart,'
                                     '[[Wholesale Orders.date_num]],'
                                     '"<="&YearEnd)')),
        Col("profit", "Gross profit", 12, "money",
            formula=zero("business", 'SUMIFS('
                                     '[[Wholesale Orders.'
                                     'line_profit]],'
                                     '[[Wholesale Orders.customer]],'
                                     '{id}{r})'),
            cf=money_cf()),
        Col("gp_pct", "GP%", 8, "percent0",
            formula=blank("business", margin("{profit}{r}",
                                             "{revenue}{r}")),
            cf=scale_cf(),
            note="Their margin to you. A big account at 18% is "
                 "worth less than a small one at 55%."),
        Col("avg_order", "Avg order", 11, "money",
            formula='=IF({business}{r}="","",IF({orders}{r}<=0,"",'
                    'ROUND({revenue}{r}/{orders}{r},2)))'),
        Col("owing", "Owing now", 11, "money",
            formula=zero("business", 'SUMIFS('
                                     '[[Wholesale Orders.'
                                     'outstanding]],'
                                     '[[Wholesale Orders.customer]],'
                                     '{id}{r})'),
            cf=owing_cf(),
            note="Invoiced and unpaid, across every open order."),
        Col("overdue", "Overdue", 11, "money",
            formula=zero("business", 'SUMIFS('
                                     '[[Wholesale Orders.'
                                     'outstanding]],'
                                     '[[Wholesale Orders.customer]],'
                                     '{id}{r},'
                                     '[[Wholesale Orders.is_overdue]]'
                                     ',1)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "bold_bad"}]),
        Col("credit_used", "Credit used", 11, "percent0",
            formula='=IF({business}{r}="","",'
                    'IF({credit_limit}{r}<=0,"",'
                    'ROUND({owing}{r}/{credit_limit}{r},4)))',
            cf=[{"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">=", "value": 0.8,
                 "fmt": "warn"}]),
        Col("reserved_value", "Reserved for them", 14, "money",
            formula=zero("business", 'SUMIFS('
                                     '[[Wholesale Orders.'
                                     'line_total]],'
                                     '[[Wholesale Orders.customer]],'
                                     '{id}{r},'
                                     '[[Wholesale Orders.'
                                     'is_reserved]],1)'),
            hidden=True,
            note="Stock sitting on your shelf with their name on it."),
        Col("last_order_num", "Last order (n)", 9, "number",
            formula=zero("business",
                         'SUMPRODUCT(MAX(('
                         '[[Wholesale Orders.customer]]={id}{r})*'
                         '[[Wholesale Orders.date_num]]))'),
            hidden=True),
        Col("last_order", "Last order", 11, "date",
            formula='=IF({business}{r}="","",'
                    'IF({last_order_num}{r}<=0,"",'
                    '{last_order_num}{r}))'),
        Col("days_since", "Days since", 10, "int",
            formula='=IF({business}{r}="","",'
                    'IF({last_order_num}{r}<=0,"",'
                    'TodayDate-{last_order_num}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 180,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 90,
                 "fmt": "warn"}]),
        Col("revenue_rank", "Rank", 7, "int",
            formula=blank("business",
                          'COUNTIFS({revenue_first}:{revenue_last},'
                          '">"&{revenue}{r})+'
                          'COUNTIFS({revenue_first}:{revenue}{r},'
                          '{revenue}{r})'),
            hidden=True),
        Col("share", "Share", 8, "percent0",
            formula=blank("business", margin("{revenue}{r}",
                                             "WholesaleRevenue")),
            cf=scale_cf(), hidden=True,
            note="What share of your wholesale revenue they are. "
                 "Anything over a third is a concentration risk."),

        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Wholesale Customer CRM",
        subtitle=("Stockists, their terms and their tier - and, on "
                  "the right, whether the account is actually worth "
                  "having."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("customers", cfg),
        help_text=("Set up the account here first. The price tier on "
                   "this row is what the Wholesale Orders tab uses "
                   "to price every line they buy, so get it right "
                   "once and never price an order by hand again."),
        extra_notes=[
            "Credit limit and Owing now work together. When an "
            "account goes past its limit the alert says so before "
            "you pick the next order, which is the only moment the "
            "information is any use.",
            "A stockist who has not ordered in six months has "
            "usually not fallen out with you - they have just run "
            "out and forgotten. Days since is the cheapest sales "
            "lead in the workbook.",
        ],
    )


def _alert() -> str:
    return alert_chain("business", [
        ('{email}{r}=""', "Add - buyer's email"),
        ('{tier}{r}=""',
         "Add - price tier, orders cannot be priced without it"),
        ('{terms}{r}=""', "Add - payment terms"),
        ('{status}{r}=""', "Add - account status"),
        ('{overdue}{r}>0.004',
         "Overdue - money past its due date, chase before shipping "
         "again"),
        ('AND({credit_limit}{r}>0,{owing}{r}>{credit_limit}{r})',
         "Urgent - over their credit limit, do not ship"),
        ('AND({credit_limit}{r}>0,{credit_used}{r}<>"",'
         '{credit_used}{r}>=0.8)',
         "Check - close to their credit limit"),
        ('{status}{r}="On hold"',
         "Action - account on hold, nothing should ship"),
        ('AND({orders}{r}>0,{gp_pct}{r}<>"",{gp_pct}{r}<MinGP)',
         "Check - margin on this account is below your floor"),
        ('AND({orders}{r}>0,{days_since}{r}<>"",'
         '{days_since}{r}>DormantDays)',
         "Follow up - gone quiet, worth a call"),
        ('AND({status}{r}="Prospect",{orders}{r}=0)',
         "Follow up - prospect who has never ordered"),
        ('AND({orders}{r}=0,{first_order}{r}="")',
         "Check - set up but never ordered"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/pricing.py`

*227 lines · 3 functions*

> 10. Wholesale Pricing - price breaks, and whether they survive.
>
> One row per SKU, mirroring the product list, with three quantity
> breaks you set yourself. Buy ten, pay this; buy fifty, pay that; buy a
> hundred, pay less again.
>
> Next to every price is its margin at the current weighted average
> cost, which is the part most price lists leave out. A tier price set
> eighteen months ago against a cost that has gone up 20% since is the
> single commonest way a wholesale business sells hard all year and
> makes nothing, and the only defence is to show the margin next to the
> price where it cannot be ignored.

| Defines | What it does |
|---------|--------------|
| `_tier(n, default_qty)` | The three columns that make up one price break |
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Wholesale Pricing').

````python
"""10. Wholesale Pricing - price breaks, and whether they survive.

One row per SKU, mirroring the product list, with three quantity
breaks you set yourself. Buy ten, pay this; buy fifty, pay that; buy a
hundred, pay less again.

Next to every price is its margin at the current weighted average
cost, which is the part most price lists leave out. A tier price set
eighteen months ago against a cost that has gone up 20% since is the
single commonest way a wholesale business sells hard all year and
makes nothing, and the only defence is to show the margin next to the
price where it cannot be ignored.
"""

from __future__ import annotations

from ..config import PRICE_TIERS, Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, blank, lookup_product,
                     lookup_product_num_zero, margin, scale_cf,
                     status_cf, yesno, zero)
from ..writer import date_check, number_check

NAME = "Wholesale Pricing"


def _tier(n: int, default_qty: int) -> list:
    """The three columns that make up one price break."""
    return [
        Col(f"tier{n}_qty", f"Tier {n} from", 11, "int",
            validate=number_check(0, 1000000),
            note=f"Order this many or more and tier {n} applies. "
                 f"Leave the whole tier blank to skip it."),
        Col(f"tier{n}_price", f"Tier {n} price", 11, "money",
            validate=number_check(0, 1000000)),
        Col(f"tier{n}_gp", f"Tier {n} GP%", 10, "percent0",
            formula=blank("sku", margin(
                f'({{tier{n}_price}}{{r}}-{{avg_cost}}{{r}})',
                f'{{tier{n}_price}}{{r}}')),
            cf=scale_cf(),
            note="Margin at this break, against weighted average "
                 "cost. Watch this column, not the price."),
    ]


def spec(cfg: Config) -> SheetSpec:
    cols = [
        Col("sku", "SKU", 15, "id",
            formula='=IF(INDEX([[Product & SKU Database.sku]],'
                    'ROW()-4)="","",'
                    'INDEX([[Product & SKU Database.sku]],ROW()-4))',
            note="Mirrors the Product Database, so every product has "
                 "a price list row whether you have filled it in or "
                 "not."),
        Col("product", "Product", 24, "text",
            formula=lookup_product("name")),
        Col("category", "Category", 13, "text",
            formula=lookup_product("category"), hidden=True),
        Col("status", "Product status", 12, "status",
            formula=lookup_product("status"), cf=status_cf(),
            hidden=True),
        Col("avg_cost", "Avg cost", 10, "money4",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.avg_cost]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,'
                                f'{lookup_product_num_zero("cost")[1:]})'),
            note="Weighted average cost, live. Every margin on this "
                 "sheet is against this number, which moves when "
                 "your suppliers move their prices."),
        Col("retail", "Retail price", 12, "money",
            formula=lookup_product_num_zero("retail_price"),
            note="From the product record, for comparison. A trade "
                 "price above half of retail usually means the "
                 "stockist cannot make it work."),
        Col("floor", "Floor price", 11, "money",
            formula=lookup_product_num_zero("min_price"),
            note="The product's own floor. Nothing on this sheet "
                 "should go below it."),

        Col("wholesale", "Base wholesale", 13, "money",
            validate=number_check(0, 1000000),
            note="Your standard trade price - what a stockist pays "
                 "below the first quantity break. Leave blank and "
                 "the product record's wholesale price is used."),
        Col("wholesale_gp", "Base GP%", 9, "percent0",
            formula=blank("sku", margin(
                '({wholesale}{r}-{avg_cost}{r})', "{wholesale}{r}")),
            cf=scale_cf()),
    ]

    for n, (qty, _label) in enumerate(PRICE_TIERS, start=1):
        cols.extend(_tier(n, qty))

    cols.extend([
        Col("best_price", "Best price", 11, "money",
            formula=zero("sku",
                         'IF({tier3_price}{r}>0,{tier3_price}{r},'
                         'IF({tier2_price}{r}>0,{tier2_price}{r},'
                         'IF({tier1_price}{r}>0,{tier1_price}{r},'
                         '{wholesale}{r})))'),
            note="The lowest price anybody can reach on this SKU by "
                 "buying volume. If this is below your floor, it is "
                 "not a discount, it is a donation."),
        Col("best_gp", "Best GP%", 9, "percent0",
            formula=blank("sku", margin(
                '({best_price}{r}-{avg_cost}{r})', "{best_price}{r}")),
            cf=scale_cf()),
        Col("discount_depth", "Discount", 9, "percent0",
            formula='=IF({sku}{r}="","",'
                    'IF({wholesale}{r}<=0,"",'
                    'ROUND(1-{best_price}{r}/{wholesale}{r},4)))',
            note="How far the deepest break is below your base "
                 "trade price.",
            hidden=True),
        Col("retail_ratio", "Of retail", 9, "percent0",
            formula='=IF({sku}{r}="","",IF({retail}{r}<=0,"",'
                    'ROUND({wholesale}{r}/{retail}{r},4)))',
            note="Base trade price as a share of retail. Half is the "
                 "convention; much over and a shop cannot double it "
                 "and still be competitive.",
            hidden=True),
        Col("moq", "MOQ", 7, "int",
            validate=number_check(0, 100000),
            note="Minimum they can buy of this SKU - a case size, "
                 "usually."),
        Col("case_qty", "Case qty", 9, "int",
            validate=number_check(0, 100000),
            note="How many come in a case. Orders that are not a "
                 "multiple of it are flagged, because splitting "
                 "cases costs you time you never charge for."),
        Col("available_trade", "Sell to trade", 12, "flag",
            validate=yesno("Some lines are retail only - seconds, "
                           "samples, packaging."),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "No", "fmt": "muted"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("reviewed", "Last reviewed", 12, "date",
            validate=date_check(),
            note="When you last looked at this price. Anything over "
                 "a year old is flagged."),
        Col("notes", "Notes", 24, "wrap"),

        Col("below_floor", "Below floor", 10, "int",
            formula=zero("sku", 'IF({floor}{r}<=0,0,'
                                'IF({best_price}{r}<=0,0,'
                                'IF({best_price}{r}<{floor}{r},1,'
                                '0)))'),
            hidden=True),
        Col("thin_margin", "Thin", 6, "int",
            formula=zero("sku", 'IF({best_price}{r}<=0,0,'
                                'IF({best_gp}{r}="",0,'
                                'IF({best_gp}{r}<MinGP,1,0)))'),
            hidden=True),
        Col("is_priced", "Priced", 7, "int",
            formula=zero("sku", 'IF({wholesale}{r}>0,1,0)'),
            hidden=True),

        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ])

    return SheetSpec(
        name=NAME,
        title="Wholesale Pricing",
        subtitle=("Three quantity breaks per SKU, with the margin at "
                  "each one shown next to it - because a price "
                  "without its margin is just a hope."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("pricing", cfg),
        help_text=("Set the base trade price and up to three "
                   "quantity breaks. Wholesale order lines price "
                   "themselves off this sheet: the workbook checks "
                   "the largest break first, so a hundred-unit order "
                   "gets the hundred-unit price."),
        extra_notes=[
            "The breaks are yours to set. The quantities shipped in "
            "the example - ten, fifty, a hundred - are a starting "
            "point, not a rule, and you can change every number on "
            "this sheet including the currency on Settings.",
            "Review prices when costs move, not on a schedule. The "
            "Avg cost column on this sheet updates every time a "
            "purchase order is received, so a supplier's price rise "
            "shows up here as a shrinking margin within days.",
        ],
    )


def _alert() -> str:
    return alert_chain("sku", [
        ('{status}{r}="Discontinued"', "Retired - discontinued line"),
        ('{available_trade}{r}="No"', "OK"),
        ('{wholesale}{r}<=0',
         "Add - base wholesale price for this SKU"),
        ('{below_floor}{r}=1',
         "Urgent - your deepest break is below the product's floor "
         "price"),
        ('AND({avg_cost}{r}>0,{best_price}{r}>0,'
         '{best_price}{r}<{avg_cost}{r})',
         "Urgent - selling below what it costs you"),
        ('{thin_margin}{r}=1',
         "Action - margin at the deepest break is under your "
         "minimum"),
        ('AND({tier2_qty}{r}>0,{tier1_qty}{r}>0,'
         '{tier2_qty}{r}<={tier1_qty}{r})',
         "Check - tier 2 does not start above tier 1"),
        ('AND({tier3_qty}{r}>0,{tier2_qty}{r}>0,'
         '{tier3_qty}{r}<={tier2_qty}{r})',
         "Check - tier 3 does not start above tier 2"),
        ('AND({tier1_price}{r}>0,{tier1_price}{r}>={wholesale}{r})',
         "Check - the first break is not cheaper than the base "
         "price"),
        ('AND({tier2_price}{r}>0,{tier1_price}{r}>0,'
         '{tier2_price}{r}>={tier1_price}{r})',
         "Check - tier 2 is not cheaper than tier 1"),
        ('AND({retail}{r}>0,{wholesale}{r}>{retail}{r}*0.7)',
         "Check - trade price is too close to retail for a shop to "
         "make it work"),
        ('AND({reviewed}{r}<>"",TodayDate-{reviewed}{r}>365)',
         "Check - this price has not been reviewed in a year"),
        ('{reviewed}{r}=""', "Add - date you last reviewed this "
                             "price"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/orders.py`

*338 lines · 2 functions*

> 11. Wholesale Orders - and the difference between promised and gone.
>
> One row per order line, sharing an Order ID the way purchase order
> lines share a PO number. Each line prices itself: the workbook reads
> the customer's tier, finds the right quantity break on the Wholesale
> Pricing tab, applies any account discount, and tells you the margin
> before you confirm.
>
> The status column is doing more work than it looks. It decides whether
> the stock on this line is RESERVED or GONE:
>
> * New - nothing has happened to stock yet.
> * Confirmed, Processing, Picking, Packed - the units are reserved.
>   Still on your shelf, no longer available to sell to anybody else.
> * Shipped, Delivered - the units have left. Stock drops.
> * Cancelled, Returned - the line releases its stock again.
>
> That distinction is the whole reason the Inventory Tracker has both an
> In stock and an Available column, and it is what stops you selling the
> same forty mugs to two shops in the same week.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Wholesale Orders').

````python
"""11. Wholesale Orders - and the difference between promised and gone.

One row per order line, sharing an Order ID the way purchase order
lines share a PO number. Each line prices itself: the workbook reads
the customer's tier, finds the right quantity break on the Wholesale
Pricing tab, applies any account discount, and tells you the margin
before you confirm.

The status column is doing more work than it looks. It decides whether
the stock on this line is RESERVED or GONE:

* New - nothing has happened to stock yet.
* Confirmed, Processing, Picking, Packed - the units are reserved.
  Still on your shelf, no longer available to sell to anybody else.
* Shipped, Delivered - the units have left. Stock drops.
* Cancelled, Returned - the line releases its stock again.

That distinction is the whole reason the Inventory Tracker has both an
In stock and an Available column, and it is what stops you selling the
same forty mugs to two shops in the same week.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, blank, date_num_col,
                     is_reserving, is_shipped, lookup_customer,
                     lookup_customer_num, lookup_product,
                     lookup_product_num_zero, margin, month_col,
                     owing_cf, scale_cf, seq_num_col, sku_col,
                     status_cf, tier_price, zero)
from ..writer import date_check, list_from, number_check

NAME = "Wholesale Orders"


def spec(cfg: Config) -> SheetSpec:
    sku = "{sku}{r}"

    cols = [
        Col("order_id", "Order ID", 12, "id",
            note="Repeat it across every line of the same order. "
                 "The printable invoice gathers the lines that "
                 "share it."),
        Col("line_no", "Line", 6, "int",
            formula=blank("order_id", 'COUNTIFS({order_id_first}:'
                                      '{order_id}{r},'
                                      '{order_id}{r})')),
        Col("is_first_line", "First line", 8, "int",
            formula=zero("order_id", 'IF({line_no}{r}=1,1,0)'),
            hidden=True,
            note="Marks one line per order, so order counts "
                 "elsewhere count orders rather than lines."),
        Col("date", "Order date", 11, "date", validate=date_check()),
        Col("customer", "Customer ID", 11, "id",
            validate=list_from("CustomerList")),
        Col("customer_name", "Customer", 20, "text",
            formula=lookup_customer("business")),
        Col("tier", "Tier", 11, "text",
            formula=lookup_customer("tier"),
            note="Read off the customer record. Change their tier "
                 "there and every future order follows."),
        Col("po_ref", "Their PO ref", 12, "text",
            note="The reference the shop gave you. Put it on the "
                 "invoice and you get paid faster."),

        sku_col(),
        Col("product", "Product", 24, "text",
            formula=lookup_product("name")),
        Col("qty", "Qty", 8, "int",
            validate=number_check(0, 1000000)),

        # ---- pricing, automatic -----------------------------------
        Col("tier_price", "Tier price", 11, "money",
            formula=zero("order_id", tier_price("{qty}{r}", sku)),
            note="The right price break for this quantity, found by "
                 "checking the largest break first so a hundred "
                 "units gets the hundred price, not the ten price."),
        Col("cust_discount", "Acct disc.", 9, "percent0",
            formula=zero("order_id",
                         lookup_customer_num("discount")[1:]),
            hidden=True),
        Col("line_discount", "Line disc.", 9, "percent0",
            validate=number_check(0, 1),
            note="Anything extra on this line only. Leave blank "
                 "unless you mean it."),
        Col("unit_price", "Unit price", 11, "money",
            formula=zero("order_id",
                         'ROUND({tier_price}{r}*'
                         '(1-{cust_discount}{r})*'
                         '(1-{line_discount}{r}),2)'),
            note="Tier price after the account discount and any "
                 "line discount. This is what they pay per unit."),
        Col("line_total", "Line total", 11, "money",
            formula=zero("order_id", 'ROUND({qty}{r}*'
                                     '{unit_price}{r},2)')),
        Col("unit_cost", "Unit cost", 10, "money4",
            formula=zero("order_id", 'IFERROR(INDEX('
                                     '[[Inventory Tracker.'
                                     'avg_cost]],MATCH({sku}{r},'
                                     '[[Inventory Tracker.sku]],0))'
                                     f'*1,{lookup_product_num_zero("cost")[1:]})'),
            hidden=True),
        Col("line_cost", "Line cost", 11, "money",
            formula=zero("order_id", 'ROUND({qty}{r}*'
                                     '{unit_cost}{r},2)'),
            hidden=True),
        Col("line_profit", "Line profit", 11, "money",
            formula=zero("order_id", 'ROUND({line_total}{r}-'
                                     '{line_cost}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"}]),
        Col("line_gp", "GP%", 8, "percent0",
            formula=blank("order_id", margin("{line_profit}{r}",
                                             "{line_total}{r}")),
            cf=scale_cf()),

        # ---- fulfilment -------------------------------------------
        Col("status", "Status", 13, "status",
            validate=list_from("OrderStatusList"), cf=status_cf(),
            note="This is the cell that moves stock. Confirmed "
                 "through Packed reserves it; Shipped and Delivered "
                 "take it out of stock for good."),
        Col("available", "Available", 10, "int",
            formula=blank("order_id", 'IFERROR(INDEX('
                                      '[[Inventory Tracker.'
                                      'available]],MATCH({sku}{r},'
                                      '[[Inventory Tracker.sku]],0))'
                                      '*1,0)'),
            note="What is free to sell right now, including this "
                 "line's own reservation. Shown here so you can see "
                 "the problem before you promise a date."),
        Col("shortfall", "Short by", 9, "int",
            formula=zero("order_id",
                         'IF({is_shipped}{r}=1,0,'
                         'MAX(0,{qty}{r}-({available}{r}+'
                         'IF({is_reserved}{r}=1,{qty}{r},0))))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bold_bad"}],
            note="How many you are short if you tried to pick this "
                 "line today. Its own reservation is added back "
                 "first, otherwise every confirmed line would look "
                 "short against itself."),
        Col("ship_date", "Shipped", 11, "date", validate=date_check()),
        Col("tracking", "Tracking", 16, "text"),
        Col("ship_method", "Method", 12, "text",
            validate=list_from("ShippingMethodList")),
        Col("shipping_charge", "Shipping", 10, "money",
            validate=number_check(0, 100000),
            note="What you charged them for carriage, on the first "
                 "line of the order only."),

        # ---- money in ---------------------------------------------
        Col("invoice_no", "Invoice no.", 12, "text"),
        Col("terms", "Terms", 11, "text",
            formula=lookup_customer("terms"), hidden=True),
        Col("due_date", "Payment due", 11, "date",
            formula='=IF({order_id}{r}="","",'
                    'IF({invoice_no}{r}="","",'
                    'IF({ship_date}{r}="","",'
                    '{ship_date}{r}+IFERROR(INDEX(TermDays,'
                    'MATCH({terms}{r},PaymentTermsList,0)),0))))'),
        Col("paid_amount", "Paid", 11, "money",
            validate=number_check(0, 1000000),
            note="What has actually landed against this line. Part "
                 "payments are fine - put the running total."),
        Col("outstanding", "Outstanding", 11, "money",
            formula=zero("order_id",
                         'IF({status}{r}="Cancelled",0,'
                         'IF({invoice_no}{r}="",0,'
                         'ROUND(MAX(0,{line_total}{r}+'
                         '{shipping_charge}{r}-'
                         '{paid_amount}{r}),2)))'),
            cf=owing_cf(),
            note="Invoiced and not yet paid. Nothing counts as owed "
                 "until there is an invoice number on the line."),
        Col("is_overdue", "Overdue", 7, "int",
            formula=zero("order_id",
                         'IF({outstanding}{r}<=0,0,'
                         'IF({due_date}{r}="",0,'
                         'IF({due_date}{r}<TodayDate,1,0)))'),
            hidden=True),
        Col("days_overdue", "Days over", 9, "int",
            formula='=IF({order_id}{r}="","",'
                    'IF({is_overdue}{r}=0,"",'
                    'TodayDate-{due_date}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 30,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("notes", "Notes", 24, "wrap"),

        # ---- the two flags the whole stock model turns on ---------
        Col("is_reserved", "Reserved", 8, "int",
            formula=zero("order_id",
                         f'IF({is_reserving("{status}{r}")},1,0)'),
            hidden=True,
            note="Confirmed, Processing, Picking or Packed: the "
                 "units are committed but still on the shelf. The "
                 "Inventory Tracker subtracts these from Available "
                 "without touching In stock."),
        Col("is_shipped", "Shipped?", 8, "int",
            formula=zero("order_id",
                         f'IF({is_shipped("{status}{r}")},1,0)'),
            hidden=True,
            note="Shipped or Delivered: the units have physically "
                 "gone and come out of stock."),
        Col("is_open", "Open", 6, "int",
            formula=zero("order_id",
                         'IF({status}{r}="Cancelled",0,'
                         'IF({is_shipped}{r}=1,0,1))'),
            hidden=True),
        Col("units_shipped", "Units out", 9, "int",
            formula=zero("order_id", 'IF({is_shipped}{r}=1,'
                                     '{qty}{r},0)'),
            hidden=True),
        Col("days_to_ship", "Days to ship", 11, "int",
            formula='=IF({order_id}{r}="","",'
                    'IF({ship_date}{r}="","",'
                    'IF({date}{r}="","",'
                    '{ship_date}{r}-{date}{r})))',
            note="Order date to despatch. Wholesale buyers forgive "
                 "a lot, but not a fortnight of silence."),
        Col("age_days", "Age", 7, "int",
            formula='=IF({order_id}{r}="","",IF({date}{r}="","",'
                    'IF({is_shipped}{r}=1,"",TodayDate-{date}{r})))',
            cf=[{"type": "cell", "criteria": ">", "value": 14,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 7,
                 "fmt": "warn"}],
            note="How long an unshipped order has been sitting."),
        Col("below_min", "Below MOQ", 9, "int",
            formula=zero("order_id",
                         'IF({is_first_line}{r}=0,0,'
                         f'IF({lookup_customer_num("min_order")[1:]}'
                         '<=0,0,'
                         'IF(SUMIFS({line_total_first}:'
                         '{line_total_last},{order_id_first}:'
                         '{order_id_last},{order_id}{r})<'
                         f'{lookup_customer_num("min_order")[1:]}'
                         ',1,0)))'),
            hidden=True),

        month_col("date"),
        date_num_col("date", "date_num"),
        date_num_col("ship_date", "ship_num"),
        seq_num_col("order_id"),

        Col("order_print_rank", "Print line", 9, "int",
            formula=zero("order_id",
                         'IF({order_id}{r}<>PrintOrder,0,'
                         '{line_no}{r})'),
            hidden=True,
            note="Line number within the order currently typed on "
                 "the Printable Orders & Invoices tab. Zero "
                 "everywhere else, which is what lets the invoice "
                 "and picking list gather their lines."),
        Col("pick_key", "Pick key", 9, "number",
            formula=zero("order_id",
                         'IF({is_reserved}{r}=0,0,'
                         'IF({shortfall}{r}>0,2000000,1000000)+'
                         'IF({age_days}{r}="",0,{age_days}{r}*100))'),
            hidden=True),
        Col("pick_rank", "Pick rank", 9, "int",
            formula=zero("order_id",
                         'IF({is_reserved}{r}=0,0,'
                         'COUNTIFS({pick_key_first}:{pick_key_last},'
                         '">"&{pick_key}{r})+'
                         'COUNTIFS({pick_key_first}:{pick_key}{r},'
                         '{pick_key}{r}))'),
            hidden=True,
            note="Orders waiting to ship, worst first: anything "
                 "short of stock, then whatever has been sitting "
                 "longest."),
        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Wholesale Orders",
        subtitle=("Lines price themselves off the customer's tier. "
                  "The status column decides whether stock is "
                  "reserved or gone."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("orders", cfg),
        help_text=("Enter the Order ID, the customer and the lines. "
                   "Price, margin and availability appear by "
                   "themselves. Move the status along as the order "
                   "progresses and the stock figures follow."),
        extra_notes=[
            "Confirming an order reserves stock; shipping it removes "
            "stock. Until you mark a line Shipped the units are "
            "still counted as in stock - because they are still in "
            "the building - but they stop being available to sell.",
            "Short by is calculated with this line's own "
            "reservation added back, so it answers the question you "
            "actually have: can I pick this order today?",
        ],
    )


def _alert() -> str:
    return alert_chain("order_id", [
        ('{status}{r}="Cancelled"', "OK"),
        ('{customer}{r}=""', "Add - which customer"),
        ('{sku}{r}=""', "Add - which product"),
        ('{qty}{r}<=0', "Add - quantity"),
        ('{unit_price}{r}<=0',
         "Urgent - no price found, check the Wholesale Pricing tab"),
        ('{line_profit}{r}<0',
         "Urgent - this line sells below cost"),
        ('AND({line_gp}{r}<>"",{line_gp}{r}<MinGP)',
         "Check - margin below your minimum"),
        ('AND({is_shipped}{r}=0,{shortfall}{r}>0)',
         "Urgent - not enough stock to pick this line"),
        ('{status}{r}=""', "Add - order status"),
        ('AND({is_shipped}{r}=1,{ship_date}{r}="")',
         "Add - date it shipped"),
        ('AND({is_shipped}{r}=1,{invoice_no}{r}="")',
         "Send - shipped but never invoiced"),
        ('{is_overdue}{r}=1',
         "Overdue - payment is past its due date, chase it"),
        ('AND({status}{r}="Packed",{ship_date}{r}="")',
         "Send - packed and waiting to go"),
        ('AND({age_days}{r}<>"",{age_days}{r}>PickWarningDays,'
         '{is_reserved}{r}=1)',
         "Action - confirmed a while ago and still not shipped"),
        ('{below_min}{r}=1',
         "Check - order is below this customer's minimum"),
        ('AND({is_shipped}{r}=1,{paid_amount}{r}<=0,'
         '{invoice_no}{r}<>"")',
         "Unpaid - invoiced and nothing received yet"),
    ])
````

---

## 8. Sales, reordering and what the stock is worth

Workflow four. Retail sales across ten channels with the marketplace and payment fees taken out, so the profit column is the money that actually arrived. The reorder planner turns the sales rate and the real lead time into a reorder point, a suggested quantity rounded to the case, and the date you have to order by. The valuation tab is the single costing method - weighted average - that every other margin in the workbook reads.

### `inventory_wholesale_spreadsheet/sheets/retail.py`

*230 lines · 2 functions*

> 12. Retail Sales Tracker - everything you sell one at a time.
>
> Marketplace orders, your own website, the shop counter, a craft fair.
> One row per item sold, because a two-item order that needs its stock
> split across two SKUs is two rows.
>
> Marketplace fees are taken seriously here. A twenty pound candle sold
> on a platform charging 6.5% plus payment processing at 2.9% and 30p
> is not a twenty pound sale, and the net profit column is the only one
> telling you the truth. Fees are calculated from the rates on Settings
> and can be overridden per row when a platform changes its mind.
>
> Nothing on this tab syncs with anything. You type your sales in or
> paste them from a marketplace export; the workbook does the rest.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Retail Sales Tracker').

````python
"""12. Retail Sales Tracker - everything you sell one at a time.

Marketplace orders, your own website, the shop counter, a craft fair.
One row per item sold, because a two-item order that needs its stock
split across two SKUs is two rows.

Marketplace fees are taken seriously here. A twenty pound candle sold
on a platform charging 6.5% plus payment processing at 2.9% and 30p
is not a twenty pound sale, and the net profit column is the only one
telling you the truth. Fees are calculated from the rates on Settings
and can be overridden per row when a platform changes its mind.

Nothing on this tab syncs with anything. You type your sales in or
paste them from a marketplace export; the workbook does the rest.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, blank, date_num_col,
                     is_sold_retail, lookup_product,
                     lookup_product_num_zero, margin, month_col,
                     scale_cf, seq_num_col, sku_col, status_cf,
                     weekday_col, zero)
from ..writer import date_check, list_from, number_check

NAME = "Retail Sales Tracker"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        Col("sale_id", "Sale ID", 11, "id",
            note="Your own reference, or the marketplace order "
                 "number. Repeat it if one order covers two items."),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("channel", "Channel", 13, "text",
            validate=list_from("ChannelList"),
            note="Where it sold. Drives the fee rate and the channel "
                 "breakdown on the reports."),
        sku_col(),
        Col("product", "Product", 24, "text",
            formula=lookup_product("name")),
        Col("qty", "Qty", 7, "int",
            validate=number_check(0, 100000)),
        Col("unit_price", "Unit price", 11, "money",
            validate=number_check(0, 1000000),
            note="What they actually paid, after any sale price or "
                 "coupon. Leave blank to use the product's retail "
                 "price."),
        Col("price_used", "Price used", 11, "money",
            formula=zero("sale_id",
                         'IF({unit_price}{r}>0,{unit_price}{r},'
                         f'{lookup_product_num_zero("retail_price")[1:]})'),
            hidden=True),
        Col("gross", "Gross sale", 11, "money",
            formula=zero("sale_id", 'ROUND({qty}{r}*'
                                    '{price_used}{r},2)')),
        Col("shipping_charged", "Ship charged", 12, "money",
            validate=number_check(0, 100000),
            note="What the customer paid for postage."),
        Col("shipping_cost", "Ship cost", 10, "money",
            validate=number_check(0, 100000),
            note="What the postage actually cost you. The gap "
                 "between these two columns is where small sellers "
                 "quietly lose money."),
        Col("discount", "Discount", 10, "money",
            validate=number_check(0, 1000000),
            note="Coupons and sale reductions, in money."),

        # ---- what the platform keeps ------------------------------
        Col("fee_rate", "Fee rate", 9, "percent0",
            formula=zero("sale_id",
                         'IFERROR(INDEX(ChannelFeeRates,'
                         'MATCH({channel}{r},ChannelList,0)),'
                         'MarketplaceFeePct)'),
            hidden=True),
        Col("platform_fee", "Platform fee", 11, "money",
            formula=zero("sale_id",
                         'IF({fee_override}{r}>0,{fee_override}{r},'
                         'ROUND(({gross}{r}+'
                         '{shipping_charged}{r}-{discount}{r})*'
                         '{fee_rate}{r},2))'),
            note="Taken from the channel's rate on Settings. Put a "
                 "number in Fee override to use the real figure off "
                 "a statement instead."),
        Col("fee_override", "Fee override", 11, "money",
            validate=number_check(0, 100000),
            note="The actual fee from the platform's statement, "
                 "when you have it. Beats the estimate."),
        Col("payment_fee", "Payment fee", 11, "money",
            formula=zero("sale_id",
                         'ROUND(({gross}{r}+'
                         '{shipping_charged}{r}-{discount}{r})*'
                         'PaymentFeePct+PaymentFeeFixed,2)'),
            note="Card and payment processing, from the rates on "
                 "Settings."),
        Col("packaging", "Packaging", 10, "money",
            formula=zero("sale_id", 'PackagingPerOrder'),
            hidden=True),
        Col("net_revenue", "Net revenue", 12, "money",
            formula=zero("sale_id",
                         'ROUND({gross}{r}+{shipping_charged}{r}-'
                         '{discount}{r}-{platform_fee}{r}-'
                         '{payment_fee}{r},2)'),
            note="What reaches your bank: the sale and the postage "
                 "charged, less discounts and everything the "
                 "platform takes."),

        # ---- what it cost you -------------------------------------
        Col("unit_cost", "Unit cost", 10, "money4",
            formula=zero("sale_id", 'IFERROR(INDEX('
                                    '[[Inventory Tracker.avg_cost]],'
                                    'MATCH({sku}{r},'
                                    '[[Inventory Tracker.sku]],0))*1,'
                                    f'{lookup_product_num_zero("cost")[1:]})'),
            hidden=True),
        Col("cogs", "COGS", 10, "money",
            formula=zero("sale_id", 'ROUND({qty}{r}*'
                                    '{unit_cost}{r},2)'),
            note="Cost of goods sold at weighted average cost - the "
                 "same method used everywhere else in the "
                 "workbook."),
        Col("profit", "Net profit", 11, "money",
            formula=zero("sale_id",
                         'ROUND({net_revenue}{r}-{cogs}{r}-'
                         '{shipping_cost}{r}-{packaging}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}],
            note="After the platform, the postage, the packaging "
                 "and the stock. This is the number that is "
                 "actually yours."),
        Col("margin", "Margin", 8, "percent0",
            formula=blank("sale_id", margin("{profit}{r}",
                                            "{gross}{r}")),
            cf=scale_cf()),

        # ---- fulfilment -------------------------------------------
        Col("status", "Status", 12, "status",
            validate=list_from("SaleStatusList"), cf=status_cf(),
            note="Stock comes off when this reaches Shipped, "
                 "Delivered or Collected."),
        Col("customer", "Customer", 18, "text",
            note="A name or a marketplace handle. Nothing more is "
                 "needed here, and nothing more should be kept."),
        Col("location", "Sold from", 11, "text",
            validate=list_from("LocationList")),
        Col("notes", "Notes", 22, "wrap"),

        Col("is_sold", "Sold?", 7, "int",
            formula=zero("sale_id",
                         f'IF({is_sold_retail("{status}{r}")},1,0)'),
            hidden=True,
            note="The flag the Inventory Tracker reads. A pending "
                 "or cancelled sale does not move stock."),
        Col("is_loss", "Loss", 6, "int",
            formula=zero("sale_id", 'IF({is_sold}{r}=0,0,'
                                    'IF({profit}{r}<0,1,0))'),
            hidden=True),
        Col("ship_gap", "Ship gap", 9, "money",
            formula=blank("sale_id",
                          'ROUND({shipping_charged}{r}-'
                          '{shipping_cost}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": -0.004,
                 "fmt": "bad"}],
            hidden=True),

        month_col("date"),
        weekday_col("date"),
        date_num_col("date", "date_num"),
        seq_num_col("sale_id"),

        Col("alert", "What needs doing", 30, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Retail Sales Tracker",
        subtitle=("Every single sale, with what the platform took "
                  "out of it. Typed or pasted in - nothing here "
                  "syncs with a marketplace."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("retail", cfg),
        help_text=("One row per item sold. Fees come from the rates "
                   "you set on Settings, so a channel that puts its "
                   "commission up costs you exactly one cell to "
                   "correct."),
        extra_notes=[
            "There is no marketplace connection in this workbook "
            "and there is not going to be one. Export your orders "
            "and paste the columns in, or type them as they come. "
            "Everything downstream - stock, COGS, product profit - "
            "runs off what is on this sheet.",
            "Ship charged against Ship cost is worth a monthly "
            "glance. Free postage over a threshold is a marketing "
            "decision; free postage by accident, on a heavy item, "
            "is just an expensive habit.",
        ],
    )


def _alert() -> str:
    return alert_chain("sale_id", [
        ('{status}{r}="Cancelled"', "OK"),
        ('{sku}{r}=""', "Add - which product sold"),
        ('{date}{r}=""', "Add - sale date"),
        ('{qty}{r}<=0', "Add - quantity"),
        ('{channel}{r}=""', "Add - which channel it sold on"),
        ('{price_used}{r}<=0',
         "Add - price, the product has no retail price either"),
        ('{status}{r}=""', "Add - status"),
        ('AND({is_sold}{r}=1,{profit}{r}<0)',
         "Loss - this sale lost money once fees and postage came "
         "off"),
        ('AND({is_sold}{r}=1,{cogs}{r}<=0)',
         "Check - no cost recorded for this product"),
        ('AND({shipping_charged}{r}>0,{shipping_cost}{r}<=0)',
         "Add - what the postage actually cost you"),
        ('{ship_gap}{r}<-ShipGapTolerance',
         "Check - postage cost well over what you charged"),
        ('AND({margin}{r}<>"",{margin}{r}<MinGP,{is_sold}{r}=1)',
         "Check - margin under your minimum"),
        ('{date}{r}>TodayDate', "Check - dated in the future"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/reorder.py`

*272 lines · 2 functions*

> 13. Reorder Planner - what to buy, how much, and how urgently.
>
> The arithmetic is the standard one and it is worth knowing by heart:
>
>     reorder point = (average daily sales x lead time) + safety stock
>     order quantity = maximum stock level - available - on order
>
> Two refinements make it honest. The lead time used is the supplier's
> REAL one, measured from your own received purchase orders, not the
> one they quoted. And safety stock is expressed in days of cover
> rather than a fixed number of units, so a line that sells ten a day
> gets a bigger cushion than one that sells ten a month.
>
> Rows sort themselves by urgency, so the top of this sheet is your
> purchasing list for the morning.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Reorder Planner').

````python
"""13. Reorder Planner - what to buy, how much, and how urgently.

The arithmetic is the standard one and it is worth knowing by heart:

    reorder point = (average daily sales x lead time) + safety stock
    order quantity = maximum stock level - available - on order

Two refinements make it honest. The lead time used is the supplier's
REAL one, measured from your own received purchase orders, not the
one they quoted. And safety stock is expressed in days of cover
rather than a fixed number of units, so a line that sells ten a day
gets a bigger cushion than one that sells ten a month.

Rows sort themselves by urgency, so the top of this sheet is your
purchasing list for the morning.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from .common import (alert_cf, alert_chain, lookup_product,
                     lookup_product_num_zero, priority_cf,
                     status_cf, zero)

NAME = "Reorder Planner"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        Col("sku", "SKU", 15, "id",
            formula='=IF(INDEX([[Product & SKU Database.sku]],'
                    'ROW()-4)="","",'
                    'INDEX([[Product & SKU Database.sku]],ROW()-4))',
            note="Every product gets a planning line automatically."),
        Col("product", "Product", 24, "text",
            formula=lookup_product("name")),
        Col("category", "Category", 13, "text",
            formula=lookup_product("category"), hidden=True),
        Col("status", "Product status", 12, "status",
            formula=lookup_product("status"), cf=status_cf(),
            hidden=True),
        Col("supplier", "Supplier ID", 11, "id",
            formula=lookup_product("supplier", missing='""'),
            hidden=True),
        Col("supplier_name", "Supplier", 18, "text",
            formula='=IF({sku}{r}="","",IFERROR(INDEX('
                    '[[Supplier Database.name]],MATCH('
                    '{supplier}{r},[[Supplier Database.id]],0))&"",'
                    '"No supplier set"))'),

        # ---- where you stand --------------------------------------
        Col("available", "Available", 10, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.available]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)')),
        Col("on_order", "On order", 9, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.on_order]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)')),
        Col("position", "Stock position", 13, "int",
            formula=zero("sku", '{available}{r}+{on_order}{r}'),
            note="Available plus what is already coming. Reordering "
                 "against available alone is how you end up with "
                 "two pallets of the same thing."),
        Col("daily_sales", "Sold/day", 9, "number1",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.daily_sales]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)'),
            note="Average daily sales over the window on Settings, "
                 "retail and wholesale together."),

        # ---- the calculation --------------------------------------
        Col("quoted_lead", "Quoted lead", 11, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Supplier Database.lead_time]],'
                                'MATCH({supplier}{r},'
                                '[[Supplier Database.id]],0))*1,'
                                'DefaultLeadDays)'),
            hidden=True),
        Col("actual_lead", "Real lead", 10, "number1",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Supplier Database.actual_lead]],'
                                'MATCH({supplier}{r},'
                                '[[Supplier Database.id]],0))*1,0)'),
            hidden=True),
        Col("lead_days", "Lead time", 10, "number1",
            formula=zero("sku",
                         'IF({actual_lead}{r}>0,'
                         'MAX({actual_lead}{r},{quoted_lead}{r}),'
                         'IF({quoted_lead}{r}>0,{quoted_lead}{r},'
                         'DefaultLeadDays))'),
            note="The longer of what they quote and what they "
                 "actually manage. Optimism about lead times is "
                 "paid for in stockouts."),
        Col("safety_stock", "Safety stock", 12, "int",
            formula=zero("sku", 'ROUNDUP({daily_sales}{r}*'
                                'SafetyDays,0)'),
            note="A cushion sized in days of cover rather than a "
                 "flat number of units, set on Settings."),
        Col("calc_point", "Calculated ROP", 13, "int",
            formula=zero("sku", 'ROUNDUP({daily_sales}{r}*'
                                '{lead_days}{r}+'
                                '{safety_stock}{r},0)'),
            note="(average daily sales x lead time) + safety stock."),
        Col("set_point", "Your ROP", 10, "int",
            formula=lookup_product_num_zero("reorder_point"),
            hidden=True),
        Col("reorder_point", "Reorder at", 11, "int",
            formula=zero("sku",
                         'IF({set_point}{r}>0,{set_point}{r},'
                         '{calc_point}{r})'),
            note="Yours if you set one on the product record, "
                 "otherwise the calculated figure."),
        Col("max_stock", "Max level", 10, "int",
            formula=lookup_product_num_zero("max_stock"),
            hidden=True),
        Col("target", "Order up to", 11, "int",
            formula=zero("sku",
                         'IF({max_stock}{r}>0,{max_stock}{r},'
                         'ROUNDUP({reorder_point}{r}*2,0))'),
            hidden=True),
        Col("suggested", "Order qty", 10, "int",
            formula=zero("sku",
                         'IF({position}{r}>{reorder_point}{r},0,'
                         'MAX(0,ROUNDUP({target}{r}-'
                         '{position}{r},0)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bold_bad"}],
            note="How many to buy to get back up to the maximum "
                 "level. Zero means you are fine."),
        Col("case_qty", "Case", 7, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Wholesale Pricing.case_qty]],'
                                'MATCH({sku}{r},'
                                '[[Wholesale Pricing.sku]],0))*1,0)'),
            hidden=True),
        Col("order_rounded", "Rounded", 9, "int",
            formula=zero("sku",
                         'IF({suggested}{r}<=0,0,'
                         'IF({case_qty}{r}<=1,{suggested}{r},'
                         'ROUNDUP({suggested}{r}/{case_qty}{r},0)*'
                         '{case_qty}{r}))'),
            note="Rounded up to whole cases where the product has a "
                 "case size."),
        Col("unit_cost", "Unit cost", 10, "money",
            formula=lookup_product_num_zero("cost"), hidden=True),
        Col("order_value", "Order value", 11, "money",
            formula=zero("sku", 'ROUND({order_rounded}{r}*'
                                '{unit_cost}{r},2)'),
            note="What this line would cost to buy. The Dashboard "
                 "adds them up, which is usually a sobering "
                 "number."),

        # ---- urgency ----------------------------------------------
        Col("days_cover", "Days cover", 10, "int",
            formula='=IF({sku}{r}="","",IF({daily_sales}{r}<=0,"",'
                    'ROUND({available}{r}/{daily_sales}{r},0)))',
            cf=[{"type": "cell", "criteria": "<=", "value": 7,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 21,
                 "fmt": "warn"}]),
        Col("stockout_date", "Runs out", 11, "date",
            formula='=IF({sku}{r}="","",IF({days_cover}{r}="","",'
                    'TodayDate+{days_cover}{r}))',
            note="When the shelf empties at the current rate of "
                 "sale, if nothing arrives."),
        Col("order_by", "Order by", 11, "date",
            formula='=IF({sku}{r}="","",IF({days_cover}{r}="","",'
                    'TodayDate+MAX(0,{days_cover}{r}-'
                    '{lead_days}{r})))',
            cf=[{"type": "cell", "criteria": "<=",
                 "value": "TodayDate", "fmt": "bad"}],
            note="The last day you can place the order and still "
                 "have it arrive before you run out. This is the "
                 "date that matters, not the stockout date."),
        Col("urgency", "Urgency", 12, "status",
            formula='=IF({sku}{r}="","",'
                    'IF({available}{r}<=0,"Out now",'
                    'IF({suggested}{r}<=0,"Fine",'
                    'IF({days_cover}{r}="","Review",'
                    'IF({days_cover}{r}<={lead_days}{r},"Urgent",'
                    'IF({days_cover}{r}<={lead_days}{r}*2,"Soon",'
                    '"Plan"))))))',
            cf=priority_cf()),
        Col("priority", "Rank", 7, "int",
            formula=zero("sku",
                         'IF({suggested}{r}<=0,0,'
                         'IF({available}{r}<=0,1,'
                         'IF({days_cover}{r}="",4,'
                         'IF({days_cover}{r}<={lead_days}{r},2,'
                         'IF({days_cover}{r}<={lead_days}{r}*2,3,'
                         '4)))))'),
            hidden=True),
        Col("needs_order", "To order", 8, "int",
            formula=zero("sku", 'IF({order_rounded}{r}>0,1,0)'),
            hidden=True),

        Col("buy_key", "Buy key", 9, "number",
            formula=zero("sku",
                         'IF({needs_order}{r}=0,0,'
                         '(10-{priority}{r})*1000000+'
                         'ROUND({order_value}{r},0))'),
            hidden=True,
            note="Urgency first, then size of the order. Ties are "
                 "broken below so no two lines ever share a rank."),
        Col("buy_rank", "Buy rank", 9, "int",
            formula=zero("sku",
                         'IF({needs_order}{r}=0,0,'
                         'COUNTIFS({buy_key_first}:{buy_key_last},'
                         '">"&{buy_key}{r})+'
                         'COUNTIFS({buy_key_first}:{buy_key}{r},'
                         '{buy_key}{r}))'),
            hidden=True,
            note="What the Dashboard reads to list the morning's "
                 "purchasing in order."),
        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Reorder Planner",
        subtitle=("(average daily sales x real lead time) + safety "
                  "stock. Everything on this sheet is calculated."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        help_text=("Sort by Rank or filter Order qty to more than "
                   "zero and you have this week's purchasing list. "
                   "Nothing here needs typing - set the lead times "
                   "on the Supplier Database and the levels on the "
                   "Product Database and this sheet keeps itself."),
        extra_notes=[
            "Order by is the date to act on. Running out on the "
            "14th is only a problem if the lead time is ten days "
            "and you order on the 12th - the planner works "
            "backwards from the stockout so you are buying against "
            "the right deadline.",
            "The lead time used is the longer of what the supplier "
            "quotes and what they actually achieve, measured from "
            "your own received orders. A supplier cannot talk their "
            "way out of their own delivery history.",
        ],
    )


def _alert() -> str:
    return alert_chain("sku", [
        ('{status}{r}="Discontinued"', "OK"),
        ('{supplier}{r}=""',
         "Add - no supplier on this product, nothing can be "
         "ordered"),
        ('AND({available}{r}<=0,{on_order}{r}>0)',
         "Due - out of stock, delivery already placed"),
        ('{available}{r}<=0',
         "Urgent - out of stock and nothing on order"),
        ('AND({order_by}{r}<>"",{order_by}{r}<TodayDate,'
         '{suggested}{r}>0)',
         "Overdue - past the date you needed to order by"),
        ('AND({urgency}{r}="Urgent",{suggested}{r}>0)',
         "Order - will run out inside the lead time"),
        ('AND({urgency}{r}="Soon",{suggested}{r}>0)',
         "Order - getting low, put it on the next order"),
        ('AND({suggested}{r}>0,{order_value}{r}>0)',
         "Plan - below the reorder point, not urgent yet"),
        ('AND({daily_sales}{r}<=0,{available}{r}>0)',
         "Check - no recent sales, do not reorder on autopilot"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/valuation.py`

*237 lines · 2 functions*

> 14. Inventory Valuation & COGS - what the stock is worth, one way.
>
> This workbook uses WEIGHTED AVERAGE COST. Not FIFO, not LIFO, not
> specific identification, and never two of them in the same sum. The
> calculation is on this sheet in full, one column per term, so you can
> follow it rather than trust it:
>
>     average cost = (opening value + value received)
>                    / (opening units + units received)
>
> Everything else in the workbook - stock value, cost of goods sold,
> product profit, the margin on a wholesale line - reads that one
> number. That is the point of picking a method and sticking to it.
> Mixing methods is the single most reliable way to produce a set of
> inventory figures that cannot be reconciled with anything, including
> themselves.
>
> If your accountant needs FIFO, tell them which method this is before
> you hand the file over.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Inventory Valuation & COGS').

````python
"""14. Inventory Valuation & COGS - what the stock is worth, one way.

This workbook uses WEIGHTED AVERAGE COST. Not FIFO, not LIFO, not
specific identification, and never two of them in the same sum. The
calculation is on this sheet in full, one column per term, so you can
follow it rather than trust it:

    average cost = (opening value + value received)
                   / (opening units + units received)

Everything else in the workbook - stock value, cost of goods sold,
product profit, the margin on a wholesale line - reads that one
number. That is the point of picking a method and sticking to it.
Mixing methods is the single most reliable way to produce a set of
inventory figures that cannot be reconciled with anything, including
themselves.

If your accountant needs FIFO, tell them which method this is before
you hand the file over.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from .common import (alert_cf, alert_chain, blank, lookup_product,
                     lookup_product_num_zero, margin, money_cf,
                     received_qty, scale_cf, sold_qty, status_cf,
                     sum_by_sku, weighted_average_cost, zero)

NAME = "Inventory Valuation & COGS"


def spec(cfg: Config) -> SheetSpec:
    sku = "{sku}{r}"

    cols = [
        Col("sku", "SKU", 15, "id",
            formula='=IF(INDEX([[Product & SKU Database.sku]],'
                    'ROW()-4)="","",'
                    'INDEX([[Product & SKU Database.sku]],ROW()-4))'),
        Col("product", "Product", 24, "text",
            formula=lookup_product("name")),
        Col("category", "Category", 13, "text",
            formula=lookup_product("category")),
        Col("status", "Status", 11, "status",
            formula=lookup_product("status"), cf=status_cf(),
            hidden=True),

        # ---- the weighted average, term by term -------------------
        Col("opening_qty", "Opening units", 12, "int",
            formula=lookup_product_num_zero("opening_qty")),
        Col("opening_value", "Opening value", 13, "money",
            formula=lookup_product_num_zero("opening_value")),
        Col("received_qty", "Units received", 13, "int",
            formula=zero("sku", received_qty(sku))),
        Col("received_value", "Value received", 13, "money",
            formula=zero("sku", sum_by_sku("Purchase Orders",
                                           "received_value", sku)),
            note="At landed cost - goods plus freight - which is "
                 "why carriage belongs on the purchase order line."),
        Col("total_qty", "Units in", 10, "int",
            formula=zero("sku", '{opening_qty}{r}+'
                                '{received_qty}{r}'),
            hidden=True),
        Col("total_value", "Value in", 12, "money",
            formula=zero("sku", 'ROUND({opening_value}{r}+'
                                '{received_value}{r},2)'),
            hidden=True),
        Col("avg_cost", "Average cost", 12, "money4",
            formula=zero("sku", weighted_average_cost(sku)),
            note="Value in divided by units in. Falls back to the "
                 "cost on the product record until the first "
                 "purchase order is received. THIS is the number "
                 "the rest of the workbook costs everything at."),
        Col("book_cost", "Standard cost", 12, "money",
            formula=lookup_product_num_zero("cost"), hidden=True),
        Col("cost_drift", "Drift", 9, "percent0",
            formula='=IF({sku}{r}="","",IF({book_cost}{r}<=0,"",'
                    'ROUND({avg_cost}{r}/{book_cost}{r}-1,4)))',
            cf=[{"type": "cell", "criteria": ">", "value": 0.1,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.03,
                 "fmt": "warn"}],
            note="Real average cost against the standard cost you "
                 "typed. Drifting up means your prices are built on "
                 "a cost that no longer exists."),

        # ---- what is left -----------------------------------------
        Col("units_sold", "Units sold", 10, "int",
            formula=zero("sku", sold_qty(sku))),
        Col("on_hand", "Units held", 10, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.current]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)')),
        Col("stock_value", "Stock at cost", 13, "money",
            formula=zero("sku", 'ROUND({on_hand}{r}*'
                                '{avg_cost}{r},2)'),
            cf=money_cf(),
            note="Your inventory valuation for this line. Add the "
                 "column up for the balance sheet figure."),
        Col("cogs", "COGS", 12, "money",
            formula=zero("sku", 'ROUND({units_sold}{r}*'
                                '{avg_cost}{r},2)'),
            note="Units sold at weighted average cost. The same "
                 "number the Profit Analysis tab uses."),
        Col("revenue", "Revenue", 12, "money",
            formula=zero("sku",
                         'SUMIFS([[Wholesale Orders.line_total]],'
                         '[[Wholesale Orders.sku]],{sku}{r},'
                         '[[Wholesale Orders.is_shipped]],1)+'
                         'SUMIFS([[Retail Sales Tracker.gross]],'
                         '[[Retail Sales Tracker.sku]],{sku}{r},'
                         '[[Retail Sales Tracker.is_sold]],1)')),
        Col("gross_profit", "Gross profit", 12, "money",
            formula=zero("sku", 'ROUND({revenue}{r}-{cogs}{r},2)'),
            cf=money_cf()),
        Col("gp_pct", "GP%", 8, "percent0",
            formula=blank("sku", margin("{gross_profit}{r}",
                                        "{revenue}{r}")),
            cf=scale_cf()),

        # ---- losses and tied-up cash ------------------------------
        Col("write_offs", "Written off", 11, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.damaged]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)')),
        Col("writeoff_value", "Write-off value", 13, "money",
            formula=zero("sku", 'ROUND({write_offs}{r}*'
                                '{avg_cost}{r},2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}],
            note="Damage, loss and unsellable returns, at cost. "
                 "This is a real expense and most spreadsheets "
                 "never show it."),
        Col("shrinkage", "Shrinkage", 9, "percent0",
            formula='=IF({sku}{r}="","",'
                    'IF({total_value}{r}<=0,"",'
                    'ROUND({writeoff_value}{r}/'
                    '{total_value}{r},4)))',
            cf=[{"type": "cell", "criteria": ">", "value": 0.03,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.01,
                 "fmt": "warn"}],
            note="Written off as a share of everything that came "
                 "in. Over 2% on a non-fragile line is worth "
                 "investigating."),
        Col("turns", "Stock turns", 11, "number1",
            formula='=IF({sku}{r}="","",IF({stock_value}{r}<=0,"",'
                    'ROUND({cogs}{r}/{stock_value}{r},1)))',
            cf=[{"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "warn"}],
            note="How many times the stock has turned over. Low "
                 "means cash sitting on a shelf; very high means "
                 "you keep running out."),
        Col("value_share", "Share of value", 12, "percent0",
            formula=blank("sku", margin("{stock_value}{r}",
                                        "TotalStockValue")),
            cf=scale_cf(), hidden=True),
        Col("abc_class", "ABC", 6, "text",
            formula='=IF({sku}{r}="","",'
                    'IF({stock_value}{r}<=0,"C",'
                    'IF({value_share}{r}="","C",'
                    'IF({value_share}{r}>=0.05,"A",'
                    'IF({value_share}{r}>=0.01,"B","C")))))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "A", "fmt": "info"}],
            note="A lines hold most of your money and deserve "
                 "counting often. C lines can be counted once a "
                 "year without anybody noticing."),
        Col("dead_value", "Dead stock", 11, "money",
            formula=zero("sku",
                         'IF(IFERROR(INDEX('
                         '[[Inventory Tracker.is_dead]],'
                         'MATCH({sku}{r},'
                         '[[Inventory Tracker.sku]],0))*1,0)=1,'
                         '{stock_value}{r},0)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            hidden=True),

        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Inventory Valuation & COGS",
        subtitle=("Weighted average cost, shown term by term. One "
                  "method, used everywhere, with no exceptions."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        help_text=("Nothing to fill in. This sheet exists so that "
                   "the number every other tab costs things at can "
                   "be checked by hand, which is what an accountant "
                   "will want to do."),
        extra_notes=[
            "Weighted average cost: (opening value + value "
            "received) divided by (opening units + units received). "
            "Received value is at landed cost, so freight and duty "
            "are in the stock valuation where they belong.",
            "Stock turns and the ABC class together tell you where "
            "to spend your counting time. A lines - the top few per "
            "cent of value - should be counted monthly; C lines "
            "once a year is plenty.",
        ],
    )


def _alert() -> str:
    return alert_chain("sku", [
        ('{total_qty}{r}<=0', [
            ('{book_cost}{r}<=0',
             "Add - no cost anywhere for this product"),
        ], "OK"),
        ('{avg_cost}{r}<=0',
         "Urgent - no average cost, every profit figure for this "
         "line is wrong"),
        ('{on_hand}{r}<0',
         "Urgent - negative stock, something was sold that was "
         "never received"),
        ('AND({revenue}{r}>0,{gp_pct}{r}<>"",{gp_pct}{r}<0)',
         "Loss - sold below cost over the period"),
        ('AND({shrinkage}{r}<>"",{shrinkage}{r}>0.03)',
         "Action - heavy write-offs on this line"),
        ('AND({cost_drift}{r}<>"",{cost_drift}{r}>0.1)',
         "Action - real cost is well above standard, reprice it"),
        ('{dead_value}{r}>0',
         "Check - cash tied up in stock that is not selling"),
        ('AND({stock_value}{r}>0,{turns}{r}<>"",{turns}{r}<1)',
         "Check - slow turning stock"),
        ('AND({abc_class}{r}="A",{on_hand}{r}>0)',
         "OK"),
    ])
````

---

## 9. Returns, locations and counting the shelf

The registers that keep the stock figure honest. Returns decide for themselves whether the unit goes back on the shelf or is written off. Locations split one SKU across bins without losing the total. The stock audit never overwrites stock: a count produces a variance, and you post the adjustment yourself, which is the only way the number stays explainable.

### `inventory_wholesale_spreadsheet/sheets/returns.py`

*205 lines · 2 functions*

> 15. Returns & Damaged Goods - and the one question that matters.
>
> Did it go back on the shelf, or is it gone?
>
> That single yes-or-no decides everything downstream. A restocked
> return adds a unit back to stock. A write-off does not - it becomes a
> loss at cost, counted against the product's profit and the business's
> shrinkage. The Restocked column is the flag the Inventory Tracker
> reads, and the condition you choose sets it for you.
>
> Returns are also where supplier quality shows itself. Three defective
> units from the same batch is a conversation with the supplier, not an
> act of God, and the Claimed column is there so you remember to have
> it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Returns & Damaged Goods').

````python
"""15. Returns & Damaged Goods - and the one question that matters.

Did it go back on the shelf, or is it gone?

That single yes-or-no decides everything downstream. A restocked
return adds a unit back to stock. A write-off does not - it becomes a
loss at cost, counted against the product's profit and the business's
shrinkage. The Restocked column is the flag the Inventory Tracker
reads, and the condition you choose sets it for you.

Returns are also where supplier quality shows itself. Three defective
units from the same batch is a conversation with the supplier, not an
act of God, and the Claimed column is there so you remember to have
it.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, date_num_col,
                     id_col, is_sellable, lookup_product,
                     lookup_product_num_zero, month_col,
                     product_name_col, seq_num_col, sku_col,
                     status_cf, yesno, zero)
from ..writer import date_check, list_from, number_check

NAME = "Returns & Damaged Goods"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        id_col("sku", "ReturnPrefix", digits=cfg.digits("return"),
               header="Return ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("source", "Came from", 13, "text",
            validate=list_from("ReturnSourceList"),
            note="A retail customer, a stockist, or your own "
                 "warehouse finding breakage."),
        Col("reference", "Order ref", 12, "text",
            note="The sale or wholesale order it came back from, if "
                 "there was one."),
        Col("customer", "Customer", 18, "text"),
        sku_col(),
        product_name_col(),
        Col("qty", "Qty", 7, "int",
            validate=number_check(0, 100000)),
        Col("reason", "Reason", 18, "text",
            validate=list_from("ReturnReasonList")),
        Col("condition", "Condition", 20, "text",
            validate=list_from("ReturnConditionList"),
            note="This is the cell that decides whether the units "
                 "come back into stock. As new and Opened - "
                 "resellable go back; everything else is a loss."),
        Col("restocked", "Restocked", 9, "int",
            formula=zero("sku",
                         f'IF({is_sellable("{condition}{r}")},1,0)'),
            note="Set by the condition. 1 means the units went back "
                 "on the shelf and the Inventory Tracker has "
                 "already counted them."),
        Col("location", "Back to", 11, "text",
            validate=list_from("LocationList"),
            note="Where the restocked units went."),

        # ---- money -------------------------------------------------
        Col("unit_cost", "Unit cost", 10, "money4",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.avg_cost]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,'
                                f'{lookup_product_num_zero("cost")[1:]})'),
            hidden=True),
        Col("stock_loss", "Stock loss", 11, "money",
            formula=zero("sku", 'IF({restocked}{r}=1,0,'
                                'ROUND({qty}{r}*{unit_cost}{r},2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}],
            note="What the units cost you, where they are not going "
                 "back on the shelf."),
        Col("refund", "Refund given", 12, "money",
            validate=number_check(0, 1000000),
            note="What you gave back to the customer."),
        Col("replacement", "Replacement sent", 14, "flag",
            validate=yesno("A replacement is a second unit of stock "
                           "going out - log it as a sale so the "
                           "stock figures stay right."),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "info"}]),
        Col("return_shipping", "Return postage", 13, "money",
            validate=number_check(0, 100000)),
        Col("total_cost", "Total cost", 11, "money",
            formula=zero("sku", 'ROUND({stock_loss}{r}+'
                                '{refund}{r}+'
                                '{return_shipping}{r},2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bold_bad"}],
            note="The whole cost of this return: the stock, the "
                 "refund and the postage. Returns are rarely as "
                 "cheap as they look."),

        # ---- supplier fault ----------------------------------------
        Col("supplier", "Supplier ID", 11, "id",
            formula=lookup_product("supplier", missing='""'),
            hidden=True),
        Col("supplier_fault", "Supplier fault", 12, "flag",
            validate=yesno("Was this a manufacturing or packing "
                           "fault at their end?"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "warn"}]),
        Col("claimed", "Claimed back", 11, "flag",
            validate=yesno("Have you actually asked them for a "
                           "credit?"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "No", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("credit_received", "Credit received", 13, "money",
            validate=number_check(0, 1000000)),
        Col("net_loss", "Net loss", 11, "money",
            formula=zero("sku", 'ROUND(MAX(0,{total_cost}{r}-'
                                '{credit_received}{r}),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}],
            note="What it cost you after anything the supplier gave "
                 "back."),
        Col("status", "Status", 12, "status",
            validate=list_from("ReturnStatusList"), cf=status_cf()),
        Col("notes", "Notes", 26, "wrap",
            note="What was actually wrong with it. Three of these "
                 "saying the same thing is a batch problem."),

        Col("is_writeoff", "Write-off", 8, "int",
            formula=zero("sku", 'IF({restocked}{r}=1,0,1)'),
            hidden=True),
        Col("is_open", "Open", 6, "int",
            formula=zero("sku",
                         'IF({status}{r}="Resolved",0,'
                         'IF({status}{r}="Closed",0,1))'),
            hidden=True),
        Col("unclaimed", "Unclaimed", 9, "money",
            formula=zero("sku",
                         'IF({supplier_fault}{r}<>"Yes",0,'
                         'IF({claimed}{r}="Yes",0,'
                         '{total_cost}{r}))'),
            hidden=True,
            note="Supplier's fault, and you have not asked for the "
                 "money back. The Dashboard adds these up."),

        month_col("date"),
        date_num_col("date", "date_num"),
        seq_num_col("sku"),

        Col("alert", "What needs doing", 30, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Returns & Damaged Goods",
        subtitle=("Restocked or written off - the condition decides, "
                  "and the Inventory Tracker follows."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("returns", cfg),
        help_text=("Log what came back, what state it was in, and "
                   "what it cost you. Pick the condition and the "
                   "restocking takes care of itself - there is no "
                   "need to write a separate stock movement."),
        extra_notes=[
            "Do not also log a return on the Stock Movement Log. "
            "This sheet already feeds the Inventory Tracker: "
            "restocked units come back in, written-off units count "
            "as damage. A second entry would count it twice.",
            "Supplier fault plus Claimed back is a small habit that "
            "pays for the workbook. Faulty stock you never claimed "
            "for is money you have decided to donate to your "
            "supplier, and the Dashboard will keep reminding you.",
        ],
    )


def _alert() -> str:
    return alert_chain("sku", [
        ('{date}{r}=""', "Add - return date"),
        ('{qty}{r}<=0', "Add - how many came back"),
        ('{condition}{r}=""',
         "Add - condition, nothing can be restocked without it"),
        ('{reason}{r}=""', "Add - reason for the return"),
        ('AND({restocked}{r}=1,{location}{r}="")',
         "Add - where the restocked units went"),
        ('{unclaimed}{r}>0',
         "Action - supplier's fault and no claim made yet"),
        ('AND({supplier_fault}{r}="Yes",{claimed}{r}="Yes",'
         '{credit_received}{r}<=0,{status}{r}<>"Closed")',
         "Chase - claim raised, credit never arrived"),
        ('AND({net_loss}{r}>LargeLossLevel,{status}{r}<>"Closed")',
         "Action - expensive return, worth understanding"),
        ('{status}{r}=""', "Add - status"),
        ('AND({is_open}{r}=1,TodayDate-{date}{r}>30)',
         "Follow up - open for over a month"),
        ('AND({refund}{r}>0,{replacement}{r}="Yes")',
         "Check - refunded and replaced, was that intended?"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/locations.py`

*161 lines · 2 functions*

> 16. Warehouse Locations - where the thing actually is.
>
> A stock figure that is right and a picker who cannot find the box are
> the same problem. One row per SKU per location, so a line split
> across two shelves gets two rows and both of them say how many.
>
> The sheet also answers the question nobody asks until it is too late:
> does the sum of the bin quantities match what the Inventory Tracker
> thinks you have? When those two disagree, one of them is wrong, and
> it is usually the bin.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Warehouse Locations').

````python
"""16. Warehouse Locations - where the thing actually is.

A stock figure that is right and a picker who cannot find the box are
the same problem. One row per SKU per location, so a line split
across two shelves gets two rows and both of them say how many.

The sheet also answers the question nobody asks until it is too late:
does the sum of the bin quantities match what the Inventory Tracker
thinks you have? When those two disagree, one of them is wrong, and
it is usually the bin.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, date_num_col,
                     lookup_product, product_name_col, sku_col,
                     yesno, zero)
from ..writer import date_check, list_from, number_check

NAME = "Warehouse Locations"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        sku_col(),
        product_name_col(),
        Col("category", "Category", 13, "text",
            formula=lookup_product("category"), hidden=True),
        Col("location", "Location", 12, "text",
            validate=list_from("LocationList"),
            note="The site or room: a unit, a garage, a shop floor, "
                 "a third party warehouse. Set the list up on "
                 "Settings."),
        Col("zone", "Zone", 9, "text",
            note="Aisle, rack or room within the location."),
        Col("bin", "Bin / shelf", 12, "text",
            note="The exact spot. 'B-14-3' beats 'the back' every "
                 "time somebody else is picking."),
        Col("qty_here", "Qty here", 10, "int",
            validate=number_check(0, 1000000),
            note="How many of this SKU are in this bin. This is a "
                 "placement record, not the stock figure - the "
                 "Inventory Tracker owns that."),
        Col("is_primary", "Primary", 8, "flag",
            validate=yesno("The first place a picker should look."),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("capacity", "Capacity", 10, "int",
            validate=number_check(0, 1000000),
            note="How many fit in this bin, if it matters."),
        Col("fill", "Full", 7, "percent0",
            formula='=IF({sku}{r}="","",IF({capacity}{r}<=0,"",'
                    'ROUND({qty_here}{r}/{capacity}{r},4)))',
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.9,
                 "fmt": "warn"}]),
        Col("type", "Storage type", 13, "text",
            validate=list_from("StorageTypeList"),
            note="Shelf, pallet, cold store, fragile. Worth "
                 "recording for anything with a handling "
                 "requirement."),
        Col("checked", "Last checked", 12, "date",
            validate=date_check()),
        Col("notes", "Notes", 24, "wrap"),

        # ---- reconciliation ---------------------------------------
        Col("total_here", "All bins", 9, "int",
            formula=zero("sku", 'SUMIFS({qty_here_first}:'
                                '{qty_here_last},{sku_first}:'
                                '{sku_last},{sku}{r})'),
            hidden=True,
            note="Everything recorded for this SKU across every "
                 "bin."),
        Col("system_qty", "System says", 11, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.current]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)'),
            note="What the Inventory Tracker calculates you hold."),
        Col("placed_gap", "Unplaced", 10, "int",
            formula=zero("sku", '{system_qty}{r}-{total_here}{r}'),
            cf=[{"type": "cell", "criteria": "!=", "value": 0,
                 "fmt": "warn"}],
            note="Stock the system believes in that no bin claims. "
                 "Not necessarily missing - often just not put away "
                 "yet - but worth knowing about."),
        Col("bins_used", "Bins", 6, "int",
            formula=zero("sku", 'COUNTIFS({sku_first}:{sku_last},'
                                '{sku}{r})'),
            hidden=True),
        Col("dup_bin", "Dup", 6, "int",
            formula=zero("sku", 'IF({location}{r}="",0,'
                                'COUNTIFS({sku_first}:{sku_last},'
                                '{sku}{r},{location_first}:'
                                '{location_last},{location}{r},'
                                '{bin_first}:{bin_last},'
                                '{bin}{r}))'),
            hidden=True),
        Col("primary_count", "Primaries", 9, "int",
            formula=zero("sku", 'COUNTIFS({sku_first}:{sku_last},'
                                '{sku}{r},{is_primary_first}:'
                                '{is_primary_last},"Yes")'),
            hidden=True),

        date_num_col("checked", "checked_num"),

        Col("alert", "What needs doing", 30, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Warehouse Locations",
        subtitle=("Which shelf, which bin, how many. One row per SKU "
                  "per location."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("locations", cfg),
        help_text=("Record where things live. A SKU split across "
                   "three shelves gets three rows; mark one of them "
                   "Primary so a picker knows where to start."),
        extra_notes=[
            "The quantities here are a placement record, not the "
            "stock figure. The Inventory Tracker owns stock and "
            "always will. Unplaced simply shows the gap between the "
            "two, which is how you find the pallet nobody put away.",
            "If you only have one storage place, fill in the "
            "location and leave the zone and bin blank. The sheet "
            "is as detailed as you need it to be and no more.",
        ],
    )


def _alert() -> str:
    return alert_chain("sku", [
        ('{location}{r}=""', "Add - which location"),
        ('{dup_bin}{r}>1',
         "Duplicate - this SKU is already recorded in this exact "
         "bin"),
        ('{primary_count}{r}>1',
         "Check - more than one bin marked primary for this SKU"),
        ('AND({bins_used}{r}>1,{primary_count}{r}=0)',
         "Add - stored in several places and none marked primary"),
        ('{qty_here}{r}<=0',
         "Check - bin recorded with nothing in it"),
        ('AND({capacity}{r}>0,{qty_here}{r}>{capacity}{r})',
         "Action - more in the bin than it holds"),
        ('{placed_gap}{r}<0',
         "Check - bins hold more than the system thinks exists"),
        ('{placed_gap}{r}>0',
         "Check - stock on the system that no bin claims"),
        ('AND({checked}{r}<>"",TodayDate-{checked}{r}>180)',
         "Check - this location has not been looked at in six "
         "months"),
        ('{checked}{r}=""', "Add - date you last checked this bin"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/audit.py`

*198 lines · 2 functions*

> 17. Stock Audit - counting, and what to do with the difference.
>
> You count the shelf, you type what you found, and the sheet tells you
> the variance against what the system believed. That part is easy.
>
> The part most stock spreadsheets get wrong is what happens next. A
> count does not silently overwrite your stock figure here. It produces
> a number and a prompt: post the difference as an adjustment on the
> Stock Movement Log, with a reason. That way the stock figure still
> reconciles to a list of events rather than to somebody's memory of a
> Tuesday afternoon, and six months later the variance is still
> explainable.
>
> The Posted column is how you mark that you have done it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Stock Audit').

````python
"""17. Stock Audit - counting, and what to do with the difference.

You count the shelf, you type what you found, and the sheet tells you
the variance against what the system believed. That part is easy.

The part most stock spreadsheets get wrong is what happens next. A
count does not silently overwrite your stock figure here. It produces
a number and a prompt: post the difference as an adjustment on the
Stock Movement Log, with a reason. That way the stock figure still
reconciles to a list of events rather than to somebody's memory of a
Tuesday afternoon, and six months later the variance is still
explainable.

The Posted column is how you mark that you have done it.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, date_num_col,
                     id_col, lookup_product_num_zero,
                     month_col, product_name_col, seq_num_col,
                     sku_col, status_cf, yesno, zero)
from ..writer import date_check, list_from, number_check

NAME = "Stock Audit"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        id_col("sku", "AuditPrefix", digits=cfg.digits("audit"),
               header="Audit ID"),
        Col("date", "Count date", 11, "date", validate=date_check()),
        Col("audit_type", "Type", 13, "text",
            validate=list_from("AuditTypeList"),
            note="A full stocktake, a cycle count of the A lines, "
                 "or a spot check because something looked wrong."),
        sku_col(),
        product_name_col(),
        Col("location", "Location", 11, "text",
            validate=list_from("LocationList")),
        Col("counted_by", "Counted by", 13, "text",
            validate=list_from("StaffList")),

        Col("system_qty", "System says", 11, "int",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.current]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)'),
            note="What the workbook calculates you hold, taken at "
                 "the moment you look at this sheet."),
        Col("counted_qty", "Counted", 10, "int",
            validate=number_check(0, 1000000),
            note="What is actually on the shelf. Count first, then "
                 "look at the system figure - knowing the answer "
                 "first is how counts come out right and stock "
                 "stays wrong."),
        Col("variance", "Variance", 10, "int",
            formula='=IF({sku}{r}="","",IF({counted_qty}{r}="","",'
                    '{counted_qty}{r}-{system_qty}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Counted minus system. Negative means there is "
                 "less on the shelf than there should be."),
        Col("variance_pct", "Variance %", 10, "percent0",
            formula='=IF({sku}{r}="","",IF({counted_qty}{r}="","",'
                    'IF({system_qty}{r}<=0,"",'
                    'ROUND({variance}{r}/{system_qty}{r},4))))',
            cf=[{"type": "cell", "criteria": "<", "value": -0.05,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.05,
                 "fmt": "warn"}]),
        Col("unit_cost", "Unit cost", 10, "money4",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.avg_cost]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,'
                                f'{lookup_product_num_zero("cost")[1:]})'),
            hidden=True),
        Col("variance_value", "Value", 11, "money",
            formula='=IF({sku}{r}="","",IF({counted_qty}{r}="","",'
                    'ROUND({variance}{r}*{unit_cost}{r},2)))',
            cf=[{"type": "cell", "criteria": "<", "value": -0.004,
                 "fmt": "bold_bad"}],
            note="What the difference is worth. This is the number "
                 "to take seriously - twenty missing stickers is "
                 "noise, two missing aprons is not."),

        Col("reason", "Likely reason", 22, "text",
            validate=list_from("VarianceReasonList"),
            note="Your best guess. Patterns show up fast: one SKU "
                 "that is always short has a process problem, not "
                 "a counting problem."),
        Col("recount", "Recounted", 10, "flag",
            validate=yesno("Did somebody count it a second time?"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("posted", "Adjustment posted", 15, "flag",
            validate=yesno("Have you written the adjustment on the "
                           "Stock Movement Log?"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "No", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}],
            note="The count on its own changes nothing. Until the "
                 "adjustment is on the movement log, the stock "
                 "figure has not moved - and that is deliberate."),
        Col("movement_ref", "Movement ID", 12, "text",
            note="The MOV- reference of the adjustment you posted, "
                 "so the two records point at each other."),
        Col("status", "Status", 12, "status",
            validate=list_from("AuditStatusList"), cf=status_cf()),
        Col("notes", "Notes", 26, "wrap"),

        Col("has_variance", "Has variance", 11, "int",
            formula=zero("sku", 'IF({counted_qty}{r}="",0,'
                                'IF({variance}{r}=0,0,1))'),
            hidden=True),
        Col("unposted", "Unposted", 9, "int",
            formula=zero("sku", 'IF({has_variance}{r}=0,0,'
                                'IF({posted}{r}="Yes",0,1))'),
            hidden=True,
            note="A counted difference nobody has acted on. These "
                 "are what the Dashboard counts."),
        Col("abs_value", "Abs value", 10, "money",
            formula=zero("sku", 'IF({counted_qty}{r}="",0,'
                                'ABS({variance_value}{r}))'),
            hidden=True),
        Col("accuracy", "Accuracy", 9, "percent0",
            formula='=IF({sku}{r}="","",IF({counted_qty}{r}="","",'
                    'IF({system_qty}{r}<=0,"",'
                    'ROUND(1-ABS({variance}{r})/'
                    '{system_qty}{r},4))))',
            cf=[{"type": "cell", "criteria": "<", "value": 0.95,
                 "fmt": "warn"}],
            hidden=True),

        month_col("date"),
        date_num_col("date", "date_num"),
        seq_num_col("sku"),

        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Stock Audit",
        subtitle=("Count it, compare it, then post the difference as "
                  "an adjustment. A count never silently overwrites "
                  "your stock."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("audit", cfg),
        help_text=("Count the shelf and type what you found. The "
                   "variance appears with what it is worth. Then "
                   "write the adjustment on the Stock Movement Log "
                   "and tick Adjustment posted - that is the step "
                   "that actually moves the stock figure."),
        extra_notes=[
            "Why not just overwrite the stock number? Because a "
            "stock figure that can be typed over is a figure "
            "nobody can explain. Posting the difference as a dated, "
            "signed adjustment means every unit you hold traces "
            "back to an event, including the ones that went "
            "missing.",
            "Count your A lines - the few SKUs holding most of your "
            "money, shown on the Inventory Valuation tab - monthly, "
            "and everything else once or twice a year. A full "
            "stocktake every quarter is a lot of lost Sundays for "
            "very little extra accuracy.",
        ],
    )


def _alert() -> str:
    return alert_chain("sku", [
        ('{date}{r}=""', "Add - count date"),
        ('{counted_qty}{r}=""', "Add - what you counted"),
        ('{counted_by}{r}=""', "Add - who counted it"),
        ('AND({has_variance}{r}=1,{abs_value}{r}>LargeLossLevel,'
         '{recount}{r}<>"Yes")',
         "Urgent - big variance, count it again before you post it"),
        ('{unposted}{r}=1',
         "Action - variance found, adjustment not posted yet"),
        ('AND({posted}{r}="Yes",{movement_ref}{r}="")',
         "Add - the movement ID of the adjustment you posted"),
        ('AND({has_variance}{r}=1,{reason}{r}="")',
         "Add - a likely reason for the difference"),
        ('AND({accuracy}{r}<>"",{accuracy}{r}<0.9)',
         "Check - this line is badly out, look at the process"),
        ('{status}{r}=""', "Add - status"),
    ])
````

---

## 10. What each product actually made

Workflow five. Product performance ranks the catalogue by profit, by return on stock and by how long it has been sitting there. The expense tracker keeps buying stock out of the operating total on purpose. The profit analysis puts the month-by-month picture together, and the reports and printable documents are the pages you hand to somebody else.

### `inventory_wholesale_spreadsheet/sheets/performance.py`

*302 lines · 2 functions*

> 18. Product Performance - which products deserve the shelf.
>
> One row per SKU, mirroring the catalogue, ranking everything you sell
> by the measures that actually decide what to reorder, what to
> discount and what to stop making.
>
> Revenue rank flatters big cheap sellers. Profit rank is better.
> Profit per unit of stock held - the last column group on this sheet -
> is better still, because it asks the only question a small business
> with finite cash actually faces: of the money I have tied up in
> stock, which lines are earning it back fastest?
>
> Dead stock gets its own flag. A line that has not sold in four months
> is not waiting for its moment; it is a shelf you are renting to
> yourself.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Product Performance').

````python
"""18. Product Performance - which products deserve the shelf.

One row per SKU, mirroring the catalogue, ranking everything you sell
by the measures that actually decide what to reorder, what to
discount and what to stop making.

Revenue rank flatters big cheap sellers. Profit rank is better.
Profit per unit of stock held - the last column group on this sheet -
is better still, because it asks the only question a small business
with finite cash actually faces: of the money I have tied up in
stock, which lines are earning it back fastest?

Dead stock gets its own flag. A line that has not sold in four months
is not waiting for its moment; it is a shelf you are renting to
yourself.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from .common import (alert_cf, alert_chain, bar_cf, blank,
                     lookup_product, margin,
                     money_cf, rank_desc, scale_cf, status_cf, zero)

NAME = "Product Performance"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        Col("sku", "SKU", 15, "id",
            formula='=IF(INDEX([[Product & SKU Database.sku]],'
                    'ROW()-4)="","",'
                    'INDEX([[Product & SKU Database.sku]],ROW()-4))'),
        Col("product", "Product", 24, "text",
            formula=lookup_product("name")),
        Col("category", "Category", 13, "text",
            formula=lookup_product("category")),
        Col("status", "Status", 11, "status",
            formula=lookup_product("status"), cf=status_cf()),

        # ---- volume -----------------------------------------------
        Col("units_retail", "Retail units", 11, "int",
            formula=zero("sku", 'SUMIFS('
                                '[[Retail Sales Tracker.qty]],'
                                '[[Retail Sales Tracker.sku]],'
                                '{sku}{r},'
                                '[[Retail Sales Tracker.is_sold]],'
                                '1)'),
            hidden=True),
        Col("units_wholesale", "Trade units", 11, "int",
            formula=zero("sku", 'SUMIFS([[Wholesale Orders.qty]],'
                                '[[Wholesale Orders.sku]],{sku}{r},'
                                '[[Wholesale Orders.is_shipped]],'
                                '1)'),
            hidden=True),
        Col("units_sold", "Units sold", 10, "int",
            formula=zero("sku", '{units_retail}{r}+'
                                '{units_wholesale}{r}'),
            cf=bar_cf()),
        Col("orders", "Orders", 8, "int",
            formula=zero("sku", 'COUNTIFS('
                                '[[Retail Sales Tracker.sku]],'
                                '{sku}{r},'
                                '[[Retail Sales Tracker.is_sold]],1)'
                                '+COUNTIFS([[Wholesale Orders.sku]],'
                                '{sku}{r},'
                                '[[Wholesale Orders.is_shipped]],1)'),
            hidden=True),

        # ---- money ------------------------------------------------
        Col("revenue_retail", "Retail revenue", 13, "money",
            formula=zero("sku", 'SUMIFS('
                                '[[Retail Sales Tracker.gross]],'
                                '[[Retail Sales Tracker.sku]],'
                                '{sku}{r},'
                                '[[Retail Sales Tracker.is_sold]],'
                                '1)'),
            hidden=True),
        Col("revenue_wholesale", "Trade revenue", 13, "money",
            formula=zero("sku", 'SUMIFS('
                                '[[Wholesale Orders.line_total]],'
                                '[[Wholesale Orders.sku]],{sku}{r},'
                                '[[Wholesale Orders.is_shipped]],'
                                '1)'),
            hidden=True),
        Col("revenue", "Revenue", 12, "money",
            formula=zero("sku", 'ROUND({revenue_retail}{r}+'
                                '{revenue_wholesale}{r},2)'),
            cf=money_cf()),
        Col("cogs", "COGS", 11, "money",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Valuation & COGS.'
                                'cogs]],MATCH({sku}{r},'
                                '[[Inventory Valuation & COGS.sku]],'
                                '0))*1,0)'),
            note="At weighted average cost, from the valuation "
                 "tab - the same method as everywhere else."),
        Col("fees", "Platform fees", 12, "money",
            formula=zero("sku", 'SUMIFS('
                                '[[Retail Sales Tracker.'
                                'platform_fee]],'
                                '[[Retail Sales Tracker.sku]],'
                                '{sku}{r},'
                                '[[Retail Sales Tracker.is_sold]],1)'
                                '+SUMIFS([[Retail Sales Tracker.'
                                'payment_fee]],'
                                '[[Retail Sales Tracker.sku]],'
                                '{sku}{r},'
                                '[[Retail Sales Tracker.is_sold]],'
                                '1)'),
            hidden=True,
            note="What the marketplaces and payment processors "
                 "took on this product's retail sales."),
        Col("returns_cost", "Returns cost", 12, "money",
            formula=zero("sku", 'SUMIFS('
                                '[[Returns & Damaged Goods.'
                                'net_loss]],'
                                '[[Returns & Damaged Goods.sku]],'
                                '{sku}{r})'),
            hidden=True),
        Col("profit", "Profit", 12, "money",
            formula=zero("sku", 'ROUND({revenue}{r}-{cogs}{r}-'
                                '{fees}{r}-{returns_cost}{r},2)'),
            cf=money_cf(),
            note="Revenue less cost of goods, less what the "
                 "platforms took, less what the returns cost. "
                 "Operating expenses are not in here - they live on "
                 "the Profit Analysis tab, where they belong."),
        Col("gp_pct", "GP%", 8, "percent0",
            formula=blank("sku", margin("{profit}{r}",
                                        "{revenue}{r}")),
            cf=scale_cf()),
        Col("profit_per_unit", "Profit/unit", 11, "money",
            formula='=IF({sku}{r}="","",IF({units_sold}{r}<=0,"",'
                    'ROUND({profit}{r}/{units_sold}{r},2)))',
            note="What one more sale is worth to you. The honest "
                 "comparison between a cheap fast seller and an "
                 "expensive slow one."),

        # ---- rate and return on stock -----------------------------
        Col("daily_sales", "Sold/day", 9, "number1",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Tracker.daily_sales]],'
                                'MATCH({sku}{r},'
                                '[[Inventory Tracker.sku]],0))*1,0)'),
            hidden=True),
        Col("monthly_units", "Units/month", 11, "number1",
            formula=zero("sku", 'ROUND({daily_sales}{r}*30,1)'),
            note="Run rate over the sales window, not a count of "
                 "last month."),
        Col("stock_value", "Stock held", 12, "money",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Valuation & COGS.'
                                'stock_value]],MATCH({sku}{r},'
                                '[[Inventory Valuation & COGS.sku]],'
                                '0))*1,0)')),
        Col("return_on_stock", "Return on stock", 14, "percent0",
            formula='=IF({sku}{r}="","",IF({stock_value}{r}<=0,"",'
                    'ROUND({profit}{r}/{stock_value}{r},4)))',
            cf=scale_cf(),
            note="Profit earned per pound of stock held. The "
                 "sharpest single measure on this sheet, and the "
                 "one to sort by when cash is tight."),
        Col("turns", "Turns", 7, "number1",
            formula=zero("sku", 'IFERROR(INDEX('
                                '[[Inventory Valuation & COGS.'
                                'turns]],MATCH({sku}{r},'
                                '[[Inventory Valuation & COGS.sku]],'
                                '0))*1,0)'),
            hidden=True),
        Col("last_sold_num", "Last sold (n)", 9, "number",
            formula=zero("sku",
                         'MAX(SUMPRODUCT(MAX(('
                         '[[Retail Sales Tracker.sku]]={sku}{r})*'
                         '[[Retail Sales Tracker.is_sold]]*'
                         '[[Retail Sales Tracker.date_num]])),'
                         'SUMPRODUCT(MAX(('
                         '[[Wholesale Orders.sku]]={sku}{r})*'
                         '[[Wholesale Orders.is_shipped]]*'
                         '[[Wholesale Orders.date_num]])))'),
            hidden=True),
        Col("last_sold", "Last sold", 11, "date",
            formula='=IF({sku}{r}="","",'
                    'IF({last_sold_num}{r}<=0,"",'
                    '{last_sold_num}{r}))'),
        Col("days_since_sale", "Days since", 10, "int",
            formula='=IF({sku}{r}="","",'
                    'IF({last_sold_num}{r}<=0,"",'
                    'TodayDate-{last_sold_num}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 120,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 60,
                 "fmt": "warn"}]),

        # ---- ranking ----------------------------------------------
        rank_desc("revenue", guard='{sku}{r}=""', header="Rev rank"),
        rank_desc("profit", guard='{sku}{r}=""',
                  header="Profit rank"),
        Col("units_rank", "Unit rank", 9, "int",
            formula=blank("sku",
                          'COUNTIFS({units_sold_first}:'
                          '{units_sold_last},">"&'
                          '{units_sold}{r})+'
                          'COUNTIFS({units_sold_first}:'
                          '{units_sold}{r},{units_sold}{r})'),
            hidden=True),
        Col("revenue_share", "Share", 8, "percent0",
            formula=blank("sku", margin("{revenue}{r}",
                                        "TotalRevenue")),
            hidden=True),
        Col("band", "Performance", 13, "status",
            formula='=IF({sku}{r}="","",'
                    'IF({units_sold}{r}<=0,"No sales",'
                    'IF({profit}{r}<0,"Losing money",'
                    'IF({days_since_sale}{r}="","Review",'
                    'IF({days_since_sale}{r}>DeadStockDays,"Dead",'
                    'IF({profit_rank}{r}<=5,"Star",'
                    'IF({days_since_sale}{r}>SlowStockDays,"Slow",'
                    'IF({gp_pct}{r}="","Review",'
                    'IF({gp_pct}{r}<MinGP,"Thin margin",'
                    '"Steady")))))))))',
            cf=status_cf()),
        Col("is_dead", "Dead", 6, "int",
            formula=zero("sku", 'IF({band}{r}="Dead",1,0)'),
            hidden=True),
        Col("is_star", "Star", 6, "int",
            formula=zero("sku", 'IF({band}{r}="Star",1,0)'),
            hidden=True),

        Col("dead_key", "Dead key", 9, "number",
            formula=zero("sku",
                         'IF({stock_value}{r}<=0,0,'
                         'IF({units_sold}{r}=0,3000000,'
                         'IF({is_dead}{r}=1,2000000,'
                         'IF({band}{r}="Slow",1000000,0)))+'
                         'ROUND({stock_value}{r},0))'),
            hidden=True),
        Col("dead_rank", "Dead rank", 9, "int",
            formula=zero("sku",
                         'IF({dead_key}{r}<=0,0,'
                         'COUNTIFS({dead_key_first}:{dead_key_last},'
                         '">"&{dead_key}{r})+'
                         'COUNTIFS({dead_key_first}:{dead_key}{r},'
                         '{dead_key}{r}))'),
            hidden=True,
            note="Stock that is not selling, most expensive first. "
                 "What the Dashboard lists as not earning its "
                 "shelf."),
        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Product Performance",
        subtitle=("Units, revenue, profit and return on the stock "
                  "held - ranked, so the decisions make themselves."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        help_text=("Everything here is calculated from your sales. "
                   "Sort by Return on stock when cash is tight, by "
                   "Profit rank when you are deciding what to "
                   "promote, and by Days since when you are "
                   "deciding what to clear."),
        extra_notes=[
            "Profit on this sheet is after cost of goods, "
            "marketplace fees and returns, but before your "
            "operating expenses. That is deliberate: rent and "
            "software are not caused by any one SKU, and spreading "
            "them across products invents a precision that does not "
            "exist. Business-level profit is on the Profit Analysis "
            "tab.",
            "A Dead line with stock still in it is worth clearing "
            "even at cost. The money is not in the product, it is "
            "in the shelf space and the cash - and both can be "
            "spent on something that sells.",
        ],
    )


def _alert() -> str:
    return alert_chain("sku", [
        ('{status}{r}="Discontinued"', "OK"),
        ('AND({units_sold}{r}=0,{stock_value}{r}>0)',
         "Action - holding stock that has never sold"),
        ('{units_sold}{r}=0', "OK"),
        ('{profit}{r}<0',
         "Urgent - this product loses money on every pass"),
        ('AND({gp_pct}{r}<>"",{gp_pct}{r}<MinGP)',
         "Action - margin below your minimum, reprice or drop it"),
        ('{is_dead}{r}=1',
         "Action - no sales in months, clear the stock"),
        ('AND({band}{r}="Slow",{stock_value}{r}>0)',
         "Check - slowing down, think about a promotion"),
        ('AND({is_star}{r}=1,{days_since_sale}{r}<>"")',
         "OK"),
        ('AND({return_on_stock}{r}<>"",{return_on_stock}{r}<0.2,'
         '{stock_value}{r}>0)',
         "Check - poor return on the cash tied up here"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/expenses.py`

*173 lines · 2 functions*

> 19. Expense Tracker - and the separation that keeps margins honest.
>
> There are two kinds of money going out of an inventory business and
> mixing them produces nonsense in both directions.
>
> Buying stock is not an expense in the way rent is. It is cash
> converted into an asset that sits on a shelf until it sells, at which
> point it becomes cost of goods sold. The workbook already counts it,
> line by line, on the Purchase Orders tab - every product's margin is
> built from it.
>
> So a stock purchase typed in here as well would be counted twice:
> once as COGS against the product and again as an operating cost
> against the business. The Stock purchase? column flags exactly that,
> and anything it flags is excluded from the operating expense totals
> rather than quietly inflating them.
>
> Rent, software, packaging, advertising, your accountant - those
> belong here, and nowhere else.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Expense Tracker').

````python
"""19. Expense Tracker - and the separation that keeps margins honest.

There are two kinds of money going out of an inventory business and
mixing them produces nonsense in both directions.

Buying stock is not an expense in the way rent is. It is cash
converted into an asset that sits on a shelf until it sells, at which
point it becomes cost of goods sold. The workbook already counts it,
line by line, on the Purchase Orders tab - every product's margin is
built from it.

So a stock purchase typed in here as well would be counted twice:
once as COGS against the product and again as an operating cost
against the business. The Stock purchase? column flags exactly that,
and anything it flags is excluded from the operating expense totals
rather than quietly inflating them.

Rent, software, packaging, advertising, your accountant - those
belong here, and nowhere else.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from .common import (alert_cf, alert_chain, blank, date_num_col,
                     id_col, monthly_equivalent, month_col,
                     seq_num_col, yesno, zero)
from ..writer import date_check, list_from, number_check

NAME = "Expense Tracker"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        id_col("description", "ExpensePrefix",
               digits=cfg.digits("expense"), header="Expense ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("description", "What it was", 28, "text"),
        Col("category", "Category", 18, "text",
            validate=list_from("ExpenseCategoryList"),
            note="Pick carefully. Stock purchases and Raw materials "
                 "are treated as inventory, not as running costs."),
        Col("supplier", "Paid to", 18, "text"),
        Col("amount", "Amount", 11, "money",
            validate=number_check(0, 1000000)),
        Col("tax_amount", "Tax / VAT", 10, "money",
            validate=number_check(0, 1000000),
            note="The recoverable tax on the invoice, if your "
                 "jurisdiction and registration let you reclaim it."),
        Col("total", "Total paid", 11, "money",
            formula=zero("description", 'ROUND({amount}{r}+'
                                        '{tax_amount}{r},2)')),
        Col("method", "Paid by", 12, "text",
            validate=list_from("PaymentMethodList")),
        Col("recurring", "Recurring", 11, "text",
            validate=list_from("FrequencyList"),
            note="One-off, monthly, quarterly, yearly. Anything "
                 "repeating gets converted to a monthly figure in "
                 "the next column."),
        Col("monthly", "Per month", 11, "money",
            formula=blank("description",
                          monthly_equivalent("{amount}{r}",
                                             "{recurring}{r}")),
            note="What this really costs each month. A yearly "
                 "licence and a monthly subscription are not "
                 "comparable until both are stated the same way."),
        Col("tax_category", "Tax category", 16, "text",
            validate=list_from("TaxCategoryList"),
            note="How your accountant will want it treated. Not "
                 "advice - a label, so the conversation is short."),
        Col("receipt", "Receipt", 9, "flag",
            validate=yesno("Have you actually got it?"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "No", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("reference", "Reference", 13, "text",
            note="Invoice number, or the PO this relates to."),
        Col("notes", "Notes", 24, "wrap"),

        # ---- the separation ---------------------------------------
        Col("is_stock", "Stock purchase?", 13, "int",
            formula=zero("description",
                         'IF(COUNTIF(StockCategoryList,'
                         '{category}{r})>0,1,0)'),
            hidden=True,
            note="1 when the category is a stock category. These "
                 "are excluded from operating expenses because the "
                 "Purchase Orders tab already counts them as cost "
                 "of goods."),
        Col("operating", "Operating cost", 12, "money",
            formula=zero("description", 'IF({is_stock}{r}=1,0,'
                                        '{amount}{r})'),
            note="What this contributes to running the business, as "
                 "opposed to filling the shelves. Stock purchases "
                 "show zero here on purpose."),
        Col("inventory_spend", "Inventory spend", 13, "money",
            formula=zero("description", 'IF({is_stock}{r}=1,'
                                        '{amount}{r},0)'),
            hidden=True),
        Col("duplicate", "Possible duplicate", 14, "int",
            formula=zero("description",
                         'IF({amount}{r}<=0,0,'
                         'COUNTIFS({date_first}:{date_last},'
                         '{date}{r},{amount_first}:{amount_last},'
                         '{amount}{r},{supplier_first}:'
                         '{supplier_last},{supplier}{r}))'),
            hidden=True,
            note="Same day, same amount, same payee. Usually a "
                 "double entry rather than a double payment, but "
                 "worth a look either way."),
        Col("size_flag", "Large", 7, "int",
            formula=zero("description",
                         'IF({amount}{r}>LargeExpenseLevel,1,0)'),
            hidden=True),

        month_col("date"),
        date_num_col("date", "date_num"),
        seq_num_col("description"),

        Col("alert", "What needs doing", 30, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Expense Tracker",
        subtitle=("Running costs. Stock purchases are counted on the "
                  "Purchase Orders tab and are deliberately kept out "
                  "of these totals."),
        columns=cols,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("expenses", cfg),
        help_text=("Everything that is not stock: rent, postage, "
                   "software, advertising, fees, your accountant. "
                   "Recurring costs get converted to a monthly "
                   "figure so you can see what the business costs "
                   "to keep open."),
        extra_notes=[
            "If you do type a stock purchase in here - and sooner "
            "or later everybody does - the Stock purchase? flag "
            "catches it and keeps it out of the operating cost "
            "total. It is not an error you need to fix, just one "
            "the workbook refuses to let distort your margins.",
            "Tax category is a label for your accountant's benefit, "
            "not a ruling. What is deductible, and when, depends "
            "entirely on where you trade and how you are set up.",
        ],
    )


def _alert() -> str:
    return alert_chain("description", [
        ('{date}{r}=""', "Add - date"),
        ('{amount}{r}<=0', "Add - amount"),
        ('{category}{r}=""', "Add - category"),
        ('{duplicate}{r}>1',
         "Duplicate - same payee, same amount, same day"),
        ('AND({is_stock}{r}=1,{reference}{r}="")',
         "Check - stock purchase, is it already on a purchase "
         "order?"),
        ('{receipt}{r}="No"', "Missing - no receipt for this one"),
        ('AND({size_flag}{r}=1,{receipt}{r}<>"Yes")',
         "Urgent - large payment with no receipt"),
        ('{receipt}{r}=""', "Check - have you got the receipt?"),
        ('{tax_category}{r}=""',
         "Add - tax category, your accountant will ask"),
        ('{recurring}{r}=""', "Add - is this a one-off or regular?"),
        ('{date}{r}>TodayDate', "Check - dated in the future"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/profit.py`

*318 lines · 2 functions*

> 20. Profit Analysis - the business, month by month.
>
> Product Performance tells you which lines earn their keep. This sheet
> tells you whether the business does, which is a different question
> and the one that pays your wages.
>
> The structure is the one an accountant would recognise:
>
>     revenue (retail + wholesale, net of what platforms took)
>     less  cost of goods sold, at weighted average cost
>     =     gross profit
>     less  operating expenses, with stock purchases excluded
>     =     net profit
>
> Stock purchases are not in the expense line. Buying inventory moves
> cash into an asset; it becomes a cost when the thing sells, through
> COGS. Counting it in both places is the single commonest way a
> spreadsheet reports a loss in a month with a big delivery and a
> fortune in the month it sells through.
>
> Twelve months of rows, generated for you.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `_alert()` |  |

**Module constants:** `NAME` ('Profit Analysis').

````python
"""20. Profit Analysis - the business, month by month.

Product Performance tells you which lines earn their keep. This sheet
tells you whether the business does, which is a different question
and the one that pays your wages.

The structure is the one an accountant would recognise:

    revenue (retail + wholesale, net of what platforms took)
    less  cost of goods sold, at weighted average cost
    =     gross profit
    less  operating expenses, with stock purchases excluded
    =     net profit

Stock purchases are not in the expense line. Buying inventory moves
cash into an asset; it becomes a cost when the thing sells, through
COGS. Counting it in both places is the single commonest way a
spreadsheet reports a loss in a month with a big delivery and a
fortune in the month it sells through.

Twelve months of rows, generated for you.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from .common import (alert_cf, alert_chain, blank, margin, money_cf,
                     scale_cf, status_cf, target_cf, zero)

NAME = "Profit Analysis"


def spec(cfg: Config) -> SheetSpec:
    cols = [
        Col("month", "Month", 11, "text",
            formula='=IF(ROW()-4>12,"",TEXT(DATE(ReportYear,'
                    'ROW()-4,1),"yyyy-mm"))',
            note="Twelve rows, one per month of the reporting year "
                 "set on Settings. Nothing to type."),
        Col("month_name", "Name", 11, "text",
            formula='=IF({month}{r}="","",TEXT(DATE(ReportYear,'
                    'ROW()-4,1),"mmmm"))'),
        Col("start_num", "From", 9, "number",
            formula='=IF({month}{r}="","",DATE(ReportYear,'
                    'ROW()-4,1))', hidden=True),
        Col("end_num", "To", 9, "number",
            formula='=IF({month}{r}="","",DATE(ReportYear,'
                    'ROW()-3,1)-1)', hidden=True),

        # ---- revenue ----------------------------------------------
        Col("retail_revenue", "Retail", 12, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Retail Sales Tracker.gross]],'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],">="&{start_num}{r},'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],"<="&{end_num}{r},'
                                  '[[Retail Sales Tracker.is_sold]],'
                                  '1)')),
        Col("wholesale_revenue", "Wholesale", 12, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Wholesale Orders.line_total]],'
                                  '[[Wholesale Orders.ship_num]],'
                                  '">="&{start_num}{r},'
                                  '[[Wholesale Orders.ship_num]],'
                                  '"<="&{end_num}{r},'
                                  '[[Wholesale Orders.is_shipped]],'
                                  '1)')),
        Col("shipping_income", "Shipping in", 11, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Retail Sales Tracker.'
                                  'shipping_charged]],'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],">="&{start_num}{r},'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],"<="&{end_num}{r},'
                                  '[[Retail Sales Tracker.is_sold]],'
                                  '1)+SUMIFS('
                                  '[[Wholesale Orders.'
                                  'shipping_charge]],'
                                  '[[Wholesale Orders.ship_num]],'
                                  '">="&{start_num}{r},'
                                  '[[Wholesale Orders.ship_num]],'
                                  '"<="&{end_num}{r},'
                                  '[[Wholesale Orders.is_shipped]],'
                                  '1)'),
            hidden=True),
        Col("revenue", "Total revenue", 13, "money",
            formula=zero("month", 'ROUND({retail_revenue}{r}+'
                                  '{wholesale_revenue}{r}+'
                                  '{shipping_income}{r},2)'),
            cf=money_cf()),

        # ---- cost of sales ----------------------------------------
        Col("cogs", "COGS", 12, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Retail Sales Tracker.cogs]],'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],">="&{start_num}{r},'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],"<="&{end_num}{r},'
                                  '[[Retail Sales Tracker.is_sold]],'
                                  '1)+SUMIFS('
                                  '[[Wholesale Orders.line_cost]],'
                                  '[[Wholesale Orders.ship_num]],'
                                  '">="&{start_num}{r},'
                                  '[[Wholesale Orders.ship_num]],'
                                  '"<="&{end_num}{r},'
                                  '[[Wholesale Orders.is_shipped]],'
                                  '1)'),
            note="Units sold at weighted average cost. Not what you "
                 "spent on stock this month - what the things you "
                 "sold this month cost you."),
        Col("fees", "Platform fees", 12, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Retail Sales Tracker.'
                                  'platform_fee]],'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],">="&{start_num}{r},'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],"<="&{end_num}{r},'
                                  '[[Retail Sales Tracker.is_sold]],'
                                  '1)+SUMIFS('
                                  '[[Retail Sales Tracker.'
                                  'payment_fee]],'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],">="&{start_num}{r},'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],"<="&{end_num}{r},'
                                  '[[Retail Sales Tracker.is_sold]],'
                                  '1)')),
        Col("shipping_cost", "Shipping out", 12, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Retail Sales Tracker.'
                                  'shipping_cost]],'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],">="&{start_num}{r},'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],"<="&{end_num}{r},'
                                  '[[Retail Sales Tracker.is_sold]],'
                                  '1)'),
            hidden=True),
        Col("gross_profit", "Gross profit", 13, "money",
            formula=zero("month", 'ROUND({revenue}{r}-{cogs}{r}-'
                                  '{fees}{r}-'
                                  '{shipping_cost}{r},2)'),
            cf=money_cf()),
        Col("gross_margin", "GP%", 8, "percent0",
            formula=blank("month", margin("{gross_profit}{r}",
                                          "{revenue}{r}")),
            cf=target_cf("TargetMargin")),

        # ---- running the business ---------------------------------
        Col("expenses", "Operating costs", 14, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Expense Tracker.operating]],'
                                  '[[Expense Tracker.date_num]],'
                                  '">="&{start_num}{r},'
                                  '[[Expense Tracker.date_num]],'
                                  '"<="&{end_num}{r})'),
            note="Rent, software, advertising, your accountant. "
                 "Stock purchases are excluded - they are already "
                 "in COGS."),
        Col("returns_cost", "Returns", 11, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Returns & Damaged Goods.'
                                  'net_loss]],'
                                  '[[Returns & Damaged Goods.'
                                  'date_num]],">="&{start_num}{r},'
                                  '[[Returns & Damaged Goods.'
                                  'date_num]],"<="&{end_num}{r})'),
            hidden=True),
        Col("write_offs", "Write-offs", 11, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Stock Movement Log.value]],'
                                  '[[Stock Movement Log.date_num]],'
                                  '">="&{start_num}{r},'
                                  '[[Stock Movement Log.date_num]],'
                                  '"<="&{end_num}{r},'
                                  '[[Stock Movement Log.'
                                  'is_writeoff]],1)'),
            hidden=True,
            note="Damage and loss at cost. A real expense, and one "
                 "most spreadsheets never show."),
        Col("net_profit", "Net profit", 13, "money",
            formula=zero("month", 'ROUND({gross_profit}{r}-'
                                  '{expenses}{r}-'
                                  '{returns_cost}{r}-'
                                  '{write_offs}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bold_good"}]),
        Col("net_margin", "Net %", 8, "percent0",
            formula=blank("month", margin("{net_profit}{r}",
                                          "{revenue}{r}")),
            cf=scale_cf()),

        # ---- cash and stock ---------------------------------------
        Col("stock_spend", "Stock bought", 12, "money",
            formula=zero("month", 'SUMIFS('
                                  '[[Purchase Orders.'
                                  'received_value]],'
                                  '[[Purchase Orders.received_num]],'
                                  '">="&{start_num}{r},'
                                  '[[Purchase Orders.received_num]],'
                                  '"<="&{end_num}{r})'),
            note="Cash that went into stock this month. Shown "
                 "separately from profit on purpose: a big delivery "
                 "is not a bad month, it is a full shelf."),
        Col("cash_effect", "Cash effect", 12, "money",
            formula=zero("month", 'ROUND({net_profit}{r}+'
                                  '{cogs}{r}-{stock_spend}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}],
            note="Roughly what the month did to your bank balance: "
                 "profit, with the stock you sold added back and "
                 "the stock you bought taken off. Profitable months "
                 "can still be cash-negative, and this is where "
                 "that shows."),
        Col("orders", "Orders", 8, "int",
            formula=zero("month", 'COUNTIFS('
                                  '[[Retail Sales Tracker.'
                                  'date_num]],">="&{start_num}{r},'
                                  '[[Retail Sales Tracker.'
                                  'date_num]],"<="&{end_num}{r},'
                                  '[[Retail Sales Tracker.is_sold]],'
                                  '1)+SUMIFS('
                                  '[[Wholesale Orders.'
                                  'is_first_line]],'
                                  '[[Wholesale Orders.ship_num]],'
                                  '">="&{start_num}{r},'
                                  '[[Wholesale Orders.ship_num]],'
                                  '"<="&{end_num}{r},'
                                  '[[Wholesale Orders.is_shipped]],'
                                  '1)'),
            hidden=True),
        Col("avg_order", "Avg order", 11, "money",
            formula='=IF({month}{r}="","",IF({orders}{r}<=0,"",'
                    'ROUND({revenue}{r}/{orders}{r},2)))',
            hidden=True),
        Col("vs_target", "vs target", 10, "percent0",
            formula='=IF({month}{r}="","",'
                    'IF(MonthlyTarget<=0,"",'
                    'ROUND({revenue}{r}/MonthlyTarget-1,4)))',
            cf=[{"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": -0.15,
                 "fmt": "bad"}]),
        Col("vs_prev", "vs last month", 12, "percent0",
            formula='=IF({month}{r}="","",IF(ROW()-4=1,"",'
                    'IF({revenue}{r_prev}<=0,"",'
                    'ROUND({revenue}{r}/{revenue}{r_prev}-1,4))))',
            cf=[{"type": "cell", "criteria": "<", "value": -0.2,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.2,
                 "fmt": "good"}]),
        Col("band", "Month", 12, "status",
            formula='=IF({month}{r}="","",'
                    'IF({revenue}{r}<=0,"No trading",'
                    'IF({net_profit}{r}<0,"Loss",'
                    'IF({net_margin}{r}="","Review",'
                    'IF({net_margin}{r}>=0.2,"Strong",'
                    'IF({net_margin}{r}>=0.08,"Steady",'
                    '"Thin"))))))',
            cf=status_cf()),

        Col("alert", "What needs doing", 32, "text", cf=alert_cf(),
            formula=_alert()),
    ]

    return SheetSpec(
        name=NAME,
        title="Profit Analysis",
        subtitle=("Revenue, cost of goods, gross profit, operating "
                  "costs, net profit. Twelve months, calculated."),
        columns=cols,
        rows=14,
        freeze_cols=2,
        help_text=("Nothing to fill in. Change the reporting year "
                   "on Settings and all twelve rows follow. Revenue "
                   "is recognised when goods ship, which is why a "
                   "confirmed-but-unshipped order is not in here."),
        extra_notes=[
            "Cash effect is the column to read in a month that "
            "looks profitable but felt tight. Profit counts the "
            "cost of what you sold; cash counts what you paid for "
            "what you bought. A month with a big delivery and a "
            "slow start will show healthy profit and ugly cash, "
            "and both numbers are correct.",
            "Wholesale revenue is counted on the shipping date, "
            "not the order date, so an order confirmed in March "
            "and shipped in April belongs to April. Reserved stock "
            "has not earned anything yet.",
        ],
    )


def _alert() -> str:
    return alert_chain("month", [
        ('{revenue}{r}<=0', "OK"),
        ('{net_profit}{r}<0',
         "Urgent - the business lost money this month"),
        ('AND({gross_margin}{r}<>"",{gross_margin}{r}<MinGP)',
         "Action - gross margin below your minimum"),
        ('{cash_effect}{r}<0',
         "Check - profitable but cash went out, stock buying was "
         "heavy"),
        ('AND({net_margin}{r}<>"",{net_margin}{r}<0.05)',
         "Check - very thin net margin"),
        ('AND({vs_prev}{r}<>"",{vs_prev}{r}<-0.25)',
         "Check - revenue well down on last month"),
        ('AND({vs_target}{r}<>"",{vs_target}{r}<-0.15)',
         "Check - behind the monthly revenue target"),
        ('{expenses}{r}>{gross_profit}{r}',
         "Action - running costs ate the whole gross profit"),
    ])
````

### `inventory_wholesale_spreadsheet/sheets/monthly.py`

*276 lines · 1 function*

> `Monthly Reports` - one month, properly examined.
>
> Pick a month at the top and everything below it follows: what sold,
> what it cost, what the platforms took, what moved in and out of
> stock, which products carried the month and which quietly lost money.
>
> The month is a real drop-down, not a formula you have to edit, and
> every figure on the page is driven from that one cell.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Monthly Reports'), `LAST_COL` (11), `TOP_ROWS` (10), `CHANNEL_ROWS` (10), `KEY` ('ReportMonth').

````python
"""`Monthly Reports` - one month, properly examined.

Pick a month at the top and everything below it follows: what sold,
what it cost, what the platforms took, what moved in and out of
stock, which products carried the month and which quietly lost money.

The month is a real drop-down, not a formula you have to edit, and
every figure on the page is driven from that one cell.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from

NAME = "Monthly Reports"
LAST_COL = 11
TOP_ROWS = 10
CHANNEL_ROWS = 10

KEY = "ReportMonth"


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)

    s.banner("Monthly Reports",
             "Pick a month. Everything below it is worked out from "
             "the registers.", LAST_COL)
    s.set_col(0, 0, 26)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 14)

    row = 4
    s.merge(row, 0, 0, "Month", fmts.chrome("label_bold"))
    s.formula(row, 1, '=TEXT(TODAY(),"yyyy-mm")', fmts.input("text"))
    s.ws.data_validation(row, 1, row, 1, list_from("MonthKeyList"))
    book.define_name(KEY, f"='{NAME}'!$B${row + 1}")
    s.merge(row, 2, 3, "Pick any month in yyyy-mm form",
            fmts.chrome("note"))
    s.merge(row, 4, 4, "Business", fmts.chrome("label_right"))
    s.merge(row, 5, 7, None, fmts.calc("text"))
    s.formula(row, 5, "=BusinessName", fmts.calc("text"))
    s.merge(row, 8, 8, "Printed", fmts.chrome("label_right"))
    s.formula(row, 9, "=TodayDate", fmts.calc("date"))
    row += 2

    # -------------------------------------------------- the headlines --
    row = s.section(row, "The month at a glance", LAST_COL)
    retail = f"SUMIFS(SaleGross,SaleMonths,{KEY},SaleIsSold,1)"
    trade = f"SUMIFS(OrderTotal,OrderMonths,{KEY},OrderIsShipped,1)"
    s.kpi(row, 0, "Revenue", f"={retail}+{trade}", "money",
          note="Retail plus shipped wholesale")
    s.kpi(row, 2, "COGS", f"=SUMIFS(ProfitCOGS,ProfitMonths,{KEY})",
          "money", note="At weighted average cost")
    s.kpi(row, 4, "Gross profit",
          f"=SUMIFS(ProfitGross,ProfitMonths,{KEY})", "money")
    s.kpi(row, 6, "Operating costs",
          f"=SUMIFS(ProfitExpenses,ProfitMonths,{KEY})", "money",
          note="Stock purchases excluded")
    s.kpi(row, 8, "Net profit",
          f"=SUMIFS(ProfitNet,ProfitMonths,{KEY})", "money")
    row = s.kpi(row, 10, "Margin",
                f'=IFERROR(SUMIFS(ProfitNet,ProfitMonths,{KEY})/'
                f'SUMIFS(ProfitRevenue,ProfitMonths,{KEY}),"")',
                "percent0", span=2)
    row += 1

    # ------------------------------------------------------ the split --
    row = s.section(row, "Where the money came from", LAST_COL)
    s.table_head(row, 0, ["", "Retail", "Wholesale", "Total",
                          "Share retail"],
                 widths=[26, 14, 14, 14, 14])
    row += 1
    rows = [
        ("Revenue", retail, trade),
        ("Units sold",
         f"SUMIFS(SaleQty,SaleMonths,{KEY},SaleIsSold,1)",
         f"SUMIFS(OrderQty,OrderMonths,{KEY},OrderIsShipped,1)"),
        ("Cost of goods",
         f"SUMIFS(SaleCOGS,SaleMonths,{KEY},SaleIsSold,1)",
         f"SUMIFS(OrderCost,OrderMonths,{KEY},OrderIsShipped,1)"),
        ("Gross profit",
         f"SUMIFS(SaleProfit,SaleMonths,{KEY},SaleIsSold,1)",
         f"SUMIFS(OrderProfit,OrderMonths,{KEY},OrderIsShipped,1)"),
    ]
    for label, left, right in rows:
        style = "int" if label == "Units sold" else "money"
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, f"={left}", fmts.calc(style))
        s.formula(row, 2, f"={right}", fmts.calc(style))
        s.formula(row, 3, f"={left}+{right}", fmts.calc(style))
        s.formula(row, 4,
                  f'=IFERROR(({left})/(({left})+({right})),"")',
                  fmts.calc("percent0"))
        row += 1
    row += 1

    # ---------------------------------------------------- what it cost --
    row = s.section(row, "What came off the top", LAST_COL)
    s.table_head(row, 0, ["Cost", "Amount", "As % of revenue",
                          "Where it is recorded"],
                 widths=[26, 14, 16, 40])
    row += 1
    costs = [
        ("Cost of goods sold",
         f"SUMIFS(ProfitCOGS,ProfitMonths,{KEY})",
         "Weighted average cost of the units that shipped"),
        ("Marketplace and payment fees",
         f"SUMIFS(SaleFees,SaleMonths,{KEY},SaleIsSold,1)+"
         f"SUMIFS(SalePaymentFees,SaleMonths,{KEY},SaleIsSold,1)",
         "Retail Sales Tracker, from the rates on Settings"),
        ("Postage you paid",
         f"SUMIFS(SaleShipCost,SaleMonths,{KEY},SaleIsSold,1)",
         "Retail Sales Tracker"),
        ("Operating expenses",
         f"SUMIFS(ExpenseOperating,ExpenseMonths,{KEY})",
         "Expense Tracker, with stock purchases excluded"),
        ("Returns and refunds",
         f"SUMIFS(ReturnLoss,ReturnMonths,{KEY})",
         "Returns & Damaged Goods, net of supplier credits"),
        ("Stock written off",
         f"SUMIFS(MovementValue,MovementMonths,{KEY},"
         f"MovementIsWriteoff,1)",
         "Stock Movement Log - damage, loss and breakage"),
    ]
    for label, formula, where in costs:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, f"={formula}", fmts.calc("money"))
        s.formula(row, 2,
                  f'=IFERROR(({formula})/'
                  f'SUMIFS(ProfitRevenue,ProfitMonths,{KEY}),"")',
                  fmts.calc("percent0"))
        s.merge(row, 3, LAST_COL, where, fmts.chrome("note"))
        row += 1
    row += 1

    # ---------------------------------------------------- stock moves --
    row = s.section(row, "What happened to the stock", LAST_COL)
    s.kpi(row, 0, "Stock bought",
          f"=SUMIFS(ProfitStockSpend,ProfitMonths,{KEY})", "money",
          note="Received at landed cost")
    s.kpi(row, 2, "Units received",
          f"=SUMIFS(POReceived,POMonths,{KEY})", "int")
    s.kpi(row, 4, "Units out",
          f"=SUMIFS(SaleQty,SaleMonths,{KEY},SaleIsSold,1)+"
          f"SUMIFS(OrderQty,OrderMonths,{KEY},OrderIsShipped,1)",
          "int")
    s.kpi(row, 6, "Units written off",
          f"=SUMIFS(MovementDamaged,MovementMonths,{KEY})", "int",
          note="Damage and loss")
    s.kpi(row, 8, "Returns",
          f"=SUMIFS(ReturnQty,ReturnMonths,{KEY})", "int")
    row = s.kpi(row, 10, "Cash effect",
                f"=SUMIFS(ProfitCash,ProfitMonths,{KEY})", "money",
                span=2, note="Profit, adjusted for stock buying")
    row += 1

    # ---------------------------------------------------- by channel --
    row = s.section(row, "By sales channel", LAST_COL)
    s.table_head(row, 0, ["Channel", "Orders", "Units", "Revenue",
                          "Fees", "Net profit", "Margin"],
                 widths=[20, 10, 10, 14, 12, 13, 10])
    row += 1
    for n in range(1, CHANNEL_ROWS + 1):
        channel = f'IFERROR(INDEX(ChannelList,{n}),"")'
        s.formula(row, 0, f'={channel}', fmts.calc("text"))
        base = (f'SaleMonths,{KEY},SaleIsSold,1,SaleChannel,'
                f'{channel}')
        s.formula(row, 1,
                  f'=IF({channel}="","",COUNTIFS({base}))',
                  fmts.calc("int"))
        s.formula(row, 2,
                  f'=IF({channel}="","",SUMIFS(SaleQty,{base}))',
                  fmts.calc("int"))
        s.formula(row, 3,
                  f'=IF({channel}="","",SUMIFS(SaleGross,{base}))',
                  fmts.calc("money"))
        s.formula(row, 4,
                  f'=IF({channel}="","",SUMIFS(SaleFees,{base})+'
                  f'SUMIFS(SalePaymentFees,{base}))',
                  fmts.calc("money"))
        s.formula(row, 5,
                  f'=IF({channel}="","",SUMIFS(SaleProfit,{base}))',
                  fmts.calc("money"))
        s.formula(row, 6,
                  f'=IF({channel}="","",IFERROR(SUMIFS(SaleProfit,'
                  f'{base})/SUMIFS(SaleGross,{base}),""))',
                  fmts.calc("percent0"))
        row += 1
    row += 1

    # ------------------------------------------------- top performers --
    row = s.section(row, "What sold", LAST_COL)
    s.table_head(row, 0, ["SKU", "Product", "Units this month",
                          "Revenue this month", "Stock left",
                          "Days cover"],
                 widths=[16, 26, 16, 18, 13, 12])
    row += 1
    for n in range(1, TOP_ROWS + 1):
        sku = (f'IFERROR(INDEX(PerfSKU,MATCH({n},PerfRevenueRank,'
               f'0)),"")')
        s.formula(row, 0, f'={sku}', fmts.calc("id"))
        s.formula(row, 1,
                  f'=IFERROR(INDEX(PerfProduct,MATCH({n},'
                  f'PerfRevenueRank,0)),"")', fmts.calc("text"))
        s.formula(
            row, 2,
            f'=IF({sku}="","",SUMIFS(SaleQty,SaleMonths,{KEY},'
            f'SaleIsSold,1,SaleSKU,{sku})+SUMIFS(OrderQty,'
            f'OrderMonths,{KEY},OrderIsShipped,1,OrderSKU,{sku}))',
            fmts.calc("int"))
        s.formula(
            row, 3,
            f'=IF({sku}="","",SUMIFS(SaleGross,SaleMonths,{KEY},'
            f'SaleIsSold,1,SaleSKU,{sku})+SUMIFS(OrderTotal,'
            f'OrderMonths,{KEY},OrderIsShipped,1,OrderSKU,{sku}))',
            fmts.calc("money"))
        s.formula(
            row, 4,
            f'=IF({sku}="","",IFERROR(INDEX(StockAvailable,'
            f'MATCH({sku},StockSKU,0)),""))', fmts.calc("int"))
        s.formula(
            row, 5,
            f'=IF({sku}="","",IFERROR(INDEX(StockDaysCover,'
            f'MATCH({sku},StockSKU,0)),""))', fmts.calc("int"))
        row += 1
    row += 1

    # -------------------------------------------------- what went wrong --
    row = s.section(row, "What to look at", LAST_COL)
    s.table_head(row, 0, ["Question", "Answer", "Worth a look if"],
                 widths=[30, 14, 50])
    row += 1
    checks = [
        ("Sales that lost money",
         f"SUMIFS(SaleIsLoss,SaleMonths,{KEY},SaleIsSold,1)", "int",
         "More than a handful - check your fee rates and postage"),
        ("Stock counts with a variance",
         f"SUMIFS(AuditHasVariance,AuditDates,\">=\"&DATE(VALUE("
         f"LEFT({KEY},4)),VALUE(RIGHT({KEY},2)),1),AuditDates,"
         f"\"<=\"&DATE(VALUE(LEFT({KEY},4)),VALUE(RIGHT({KEY},2))"
         f"+1,0))", "int",
         "Any at all - and every one should be posted as an "
         "adjustment"),
        ("Deliveries that arrived late",
         f"SUMIFS(POIsLate,POMonths,{KEY})", "int",
         "A pattern from one supplier means raising their lead "
         "time"),
        ("Returns logged",
         f"COUNTIFS(ReturnMonths,{KEY})", "int",
         "Rising month on month, or clustered on one product"),
        ("Expenses with no receipt",
         f'COUNTIFS(ExpenseMonths,{KEY},ExpenseReceipt,"No")', "int",
         "Anything over zero, before your accountant asks"),
    ]
    for label, formula, style, when in checks:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, f"={formula}", fmts.calc(style))
        s.merge(row, 2, LAST_COL, when, fmts.chrome("note"))
        row += 1
    row += 1

    row = s.note(
        row,
        "Revenue is recognised when goods ship, not when an order "
        "is placed. An order confirmed on the 28th and shipped on "
        "the 2nd belongs to the following month, which is why a "
        "busy end of month can look quiet here and busy next time.",
        LAST_COL, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `inventory_wholesale_spreadsheet/sheets/annual.py`

*250 lines · 1 function*

> `Annual Reports` - the year, and whether it was a good one.
>
> Twelve months side by side, then the questions you only ask once a
> year: what is the stock actually worth at the year end, how many
> times did it turn, what proportion of it never sold, which
> twenty per cent of products made eighty per cent of the profit, and
> what is the business worth keeping doing.
>
> This is also the page to print for an accountant. The closing stock
> valuation, the cost of goods sold and the costing method are all
> stated on it explicitly, which is the first thing they will ask for.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Annual Reports'), `LAST_COL` (13), `TOP_ROWS` (12), `CATEGORY_ROWS` (8).

````python
"""`Annual Reports` - the year, and whether it was a good one.

Twelve months side by side, then the questions you only ask once a
year: what is the stock actually worth at the year end, how many
times did it turn, what proportion of it never sold, which
twenty per cent of products made eighty per cent of the profit, and
what is the business worth keeping doing.

This is also the page to print for an accountant. The closing stock
valuation, the cost of goods sold and the costing method are all
stated on it explicitly, which is the first thing they will ask for.
"""

from __future__ import annotations

from ..config import COSTING_NOTE, DISCLAIMER, Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Annual Reports"
LAST_COL = 13
TOP_ROWS = 12
CATEGORY_ROWS = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.deep)

    s.banner("Annual Reports",
             "The whole year, and the figures an accountant will ask "
             "for.", LAST_COL)
    s.set_col(0, 0, 24)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 12)

    row = 4
    s.merge(row, 0, 0, "Year", fmts.chrome("label_bold"))
    s.formula(row, 1, "=ReportYear", fmts.calc("int"))
    s.merge(row, 2, 4, "Set on Settings & Lists",
            fmts.chrome("note"))
    s.merge(row, 5, 5, "Business", fmts.chrome("label_right"))
    s.merge(row, 6, 9, None, fmts.calc("text"))
    s.formula(row, 6, "=BusinessName", fmts.calc("text"))
    s.merge(row, 10, 10, "Printed", fmts.chrome("label_right"))
    s.formula(row, 11, "=TodayDate", fmts.calc("date"))
    row += 2

    # ------------------------------------------------- the year's totals --
    row = s.section(row, "The year", LAST_COL)
    s.kpi(row, 0, "Revenue", "=SUM(ProfitRevenue)", "money")
    s.kpi(row, 2, "COGS", "=SUM(ProfitCOGS)", "money",
          note="Weighted average")
    s.kpi(row, 4, "Gross profit", "=SUM(ProfitGross)", "money")
    s.kpi(row, 6, "Operating costs", "=SUM(ProfitExpenses)", "money")
    s.kpi(row, 8, "Net profit", "=SUM(ProfitNet)", "money")
    s.kpi(row, 10, "Net margin",
          '=IFERROR(SUM(ProfitNet)/SUM(ProfitRevenue),"")',
          "percent0")
    row = s.kpi(row, 12, "vs target",
                '=IF(YearlyTarget<=0,"",SUM(ProfitRevenue)/'
                'YearlyTarget-1)', "percent0", span=1)
    row += 1

    # ------------------------------------------------------- the months --
    row = s.section(row, "Month by month", LAST_COL)
    s.table_head(row, 0, ["", "Jan", "Feb", "Mar", "Apr", "May",
                          "Jun", "Jul", "Aug", "Sep", "Oct", "Nov",
                          "Dec"],
                 widths=[24] + [12] * 12)
    row += 1
    lines = [
        ("Revenue", "revenue", "money"),
        ("Retail", "retail_revenue", "money"),
        ("Wholesale", "wholesale_revenue", "money"),
        ("COGS", "cogs", "money"),
        ("Gross profit", "gross_profit", "money"),
        ("Margin", "gross_margin", "percent0"),
        ("Operating costs", "expenses", "money"),
        ("Net profit", "net_profit", "money"),
        ("Stock bought", "stock_spend", "money"),
        ("Cash effect", "cash_effect", "money"),
    ]
    for label, key, style in lines:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        for month in range(1, 13):
            s.formula(
                row, month,
                f'=IFERROR(INDEX([[Profit Analysis.{key}]],'
                f'{month}),"")', fmts.calc(style))
        if style == "money":
            s.cf(f"B{row + 1}:M{row + 1}",
                 {"type": "cell", "criteria": "<", "value": 0,
                  "fmt": "bad"})
        row += 1
    row += 1

    # ------------------------------------------------ the stock position --
    row = s.section(row, "Stock at the year end", LAST_COL)
    row = s.note(
        row,
        "These are the numbers an accountant asks for. Closing "
        "stock is valued at weighted average cost, the same method "
        "used in every calculation in this workbook.",
        LAST_COL, height=32)
    row += 1
    s.kpi(row, 0, "Closing stock at cost", "=TotalStockValue",
          "money", note="Weighted average")
    s.kpi(row, 2, "At retail value", "=SUM(StockRetailValue)",
          "money")
    s.kpi(row, 4, "SKUs held",
          "=SUMPRODUCT((StockCurrent>0)*1)", "int")
    s.kpi(row, 6, "Units held", "=SUM(StockCurrent)", "int")
    s.kpi(row, 8, "Stock turns",
          '=IFERROR(SUM(ProfitCOGS)/TotalStockValue,"")', "number1",
          note="COGS divided by stock held")
    s.kpi(row, 10, "Dead stock value", "=SUM(ValuationDead)",
          "money", note="No sales in the window")
    row = s.kpi(row, 12, "Written off",
                "=SUM(ValuationWriteoff)", "money", span=1)
    row += 1

    # --------------------------------------------------- by category --
    row = s.section(row, "By category", LAST_COL)
    s.table_head(row, 0, ["Category", "SKUs", "Units sold",
                          "Revenue", "COGS", "Gross profit", "GP%",
                          "Stock held", "Share of stock"],
                 widths=[20, 9, 12, 13, 12, 13, 9, 13, 14])
    row += 1
    for n in range(1, CATEGORY_ROWS + 1):
        cat = f'IFERROR(INDEX(CategoryList,{n}),"")'
        s.formula(row, 0, f'={cat}', fmts.calc("text"))
        s.formula(row, 1,
                  f'=IF({cat}="","",COUNTIFS(PerfCategory,{cat}))',
                  fmts.calc("int"))
        s.formula(row, 2,
                  f'=IF({cat}="","",SUMIFS(PerfUnits,PerfCategory,'
                  f'{cat}))', fmts.calc("int"))
        s.formula(row, 3,
                  f'=IF({cat}="","",SUMIFS(PerfRevenue,'
                  f'PerfCategory,{cat}))', fmts.calc("money"))
        s.formula(row, 4,
                  f'=IF({cat}="","",SUMIFS(ValuationCOGS,'
                  f'StockCategory,{cat}))', fmts.calc("money"))
        s.formula(row, 5,
                  f'=IF({cat}="","",SUMIFS(PerfProfit,PerfCategory,'
                  f'{cat}))', fmts.calc("money"))
        s.formula(row, 6,
                  f'=IF({cat}="","",IFERROR(SUMIFS(PerfProfit,'
                  f'PerfCategory,{cat})/SUMIFS(PerfRevenue,'
                  f'PerfCategory,{cat}),""))',
                  fmts.calc("percent0"))
        s.formula(row, 7,
                  f'=IF({cat}="","",SUMIFS(StockValueCol,'
                  f'StockCategory,{cat}))', fmts.calc("money"))
        s.formula(row, 8,
                  f'=IF({cat}="","",IFERROR(SUMIFS(StockValueCol,'
                  f'StockCategory,{cat})/TotalStockValue,""))',
                  fmts.calc("percent0"))
        row += 1
    row += 1

    # ------------------------------------------------------ top products --
    row = s.section(row, "The products that carried the year",
                    LAST_COL)
    s.table_head(row, 0, ["Rank", "SKU", "Product", "Units",
                          "Revenue", "Profit", "GP%",
                          "Return on stock", "Running share"],
                 widths=[7, 16, 26, 10, 13, 13, 9, 14, 14])
    row += 1
    first_top = row
    for n in range(1, TOP_ROWS + 1):
        s.write(row, 0, n, fmts.calc("int"))
        for col, field, style in ((1, "PerfSKU", "id"),
                                  (2, "PerfProduct", "text"),
                                  (3, "PerfUnits", "int"),
                                  (4, "PerfRevenue", "money"),
                                  (5, "PerfProfit", "money"),
                                  (6, "PerfGP", "percent0"),
                                  (7, "PerfReturnOnStock",
                                   "percent0")):
            s.formula(row, col,
                      f'=IFERROR(INDEX({field},MATCH({n},'
                      f'PerfProfitRank,0)),"")', fmts.calc(style))
        s.formula(
            row, 8,
            f'=IFERROR(SUM($F${first_top + 1}:$F{row + 1})/'
            f'SUM(PerfProfit),"")', fmts.calc("percent0"))
        row += 1
    row = s.note(
        row,
        "Running share is the cumulative profit of the lines above, "
        "as a share of the whole year. If three products get you "
        "past sixty per cent, that is where your attention belongs "
        "- and it is also a concentration risk worth naming out "
        "loud.",
        LAST_COL, height=38)
    row += 1

    # --------------------------------------------------- the suppliers --
    row = s.section(row, "Suppliers", LAST_COL)
    s.kpi(row, 0, "Suppliers used",
          '=SUMPRODUCT((SupplierSpend>0)*1)', "int")
    s.kpi(row, 2, "Total spend", "=TotalPurchaseSpend", "money")
    s.kpi(row, 4, "Still committed", "=SUM(SupplierOpenValue)",
          "money", note="Ordered, not delivered")
    s.kpi(row, 6, "Late deliveries", "=SUM(POIsLate)", "int")
    s.kpi(row, 8, "Biggest supplier share",
          '=IFERROR(MAX(SupplierSpend)/TotalPurchaseSpend,"")',
          "percent0", note="Over half is a risk")
    row = s.kpi(row, 10, "Avg lead gap",
                '=IFERROR(SUMPRODUCT(SupplierLeadGap)/'
                'SUMPRODUCT((SupplierLeadGap<>"")*1),"")',
                "number1", span=3,
                note="Days slower than quoted, on average")
    row += 1

    # ---------------------------------------------------- the customers --
    row = s.section(row, "Wholesale customers", LAST_COL)
    s.kpi(row, 0, "Customers who ordered",
          "=SUMPRODUCT((CustomerOrders>0)*1)", "int")
    s.kpi(row, 2, "Wholesale revenue", "=WholesaleRevenue", "money")
    s.kpi(row, 4, "Wholesale profit", "=SUM(CustomerProfit)",
          "money")
    s.kpi(row, 6, "Still owed", "=SUM(CustomerOwing)", "money")
    s.kpi(row, 8, "Overdue", "=SUM(CustomerOverdue)", "money")
    row = s.kpi(row, 10, "Biggest account share",
                '=IFERROR(MAX(CustomerRevenue)/WholesaleRevenue,"")',
                "percent0", span=3,
                note="A third or more is a concentration risk")
    row += 1

    # ------------------------------------------------------- year notes --
    row = s.section(row, "For your accountant", LAST_COL)
    row = s.note(row, COSTING_NOTE, LAST_COL, height=80)
    row += 1
    row = s.note(
        row,
        "Closing stock at cost, cost of goods sold and gross "
        "profit are all on this page and all computed on the same "
        "weighted average basis. Operating expenses exclude stock "
        "purchases, which are capitalised into inventory and "
        "released through COGS as goods sell.",
        LAST_COL, height=44)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=76)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `inventory_wholesale_spreadsheet/sheets/printables.py`

*354 lines · 2 functions*

> `Printable Orders & Invoices` - three documents, ready to send.
>
> Type an ID at the top of each block and the document fills itself in:
> a purchase order to send a supplier, a wholesale invoice to send a
> customer, and a picking list for whoever is packing the boxes.
>
> Each one gathers every line that shares the ID you typed, which is
> why purchase order and order lines repeat their reference across
> rows. Nothing here is typed twice.
>
> Set your print area to the block you want and it comes out on one
> page.

| Defines | What it does |
|---------|--------------|
| `_line(sheet, key_name, id_cell, field, n)` | The nth line of whichever document the user asked for |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Printable Orders & Invoices'), `LAST_COL` (8), `PO_LINES` (14), `INV_LINES` (14), `PICK_LINES` (16).

````python
"""`Printable Orders & Invoices` - three documents, ready to send.

Type an ID at the top of each block and the document fills itself in:
a purchase order to send a supplier, a wholesale invoice to send a
customer, and a picking list for whoever is packing the boxes.

Each one gathers every line that shares the ID you typed, which is
why purchase order and order lines repeat their reference across
rows. Nothing here is typed twice.

Set your print area to the block you want and it comes out on one
page.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Printable Orders & Invoices"
LAST_COL = 8
PO_LINES = 14
INV_LINES = 14
PICK_LINES = 16


def _line(sheet: str, key_name: str, id_cell: str, field: str,
          n: int) -> str:
    """The nth line of whichever document the user asked for.

    Walks the register looking for the nth row whose ID matches, by
    counting matches as it goes - no array formula, no helper tab.
    """
    return (f'IFERROR(IF({id_cell}="","",'
            f'INDEX([[{sheet}.{field}]],'
            f'MATCH({n},[[{sheet}.{key_name}]],0))),"")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.accent)

    s.banner("Printable Orders & Invoices",
             "Type an ID, print the document. Everything else is "
             "filled in for you.", LAST_COL)
    s.set_col(0, 0, 22)
    s.set_col(1, 1, 30)
    for column in range(2, LAST_COL + 1):
        s.set_col(column, column, 15)

    row = 4
    row = s.note(
        row,
        "Three documents, one below the other. Each needs one thing "
        "from you: the PO number, the order ID or the order ID "
        "again. Every line that shares that reference is gathered "
        "automatically, so a six-line order prints as six lines.",
        LAST_COL, height=36)
    row += 1

    # =====================================================================
    # 1. PURCHASE ORDER
    # =====================================================================
    row = s.section(row, "Purchase order - to send to a supplier",
                    LAST_COL)
    s.merge(row, 0, 0, "PO number", fmts.chrome("label_bold"))
    s.write(row, 1, None, fmts.input("id"))
    po_cell = f"$B${row + 1}"
    book.define_name("PrintPO", f"='{NAME}'!{po_cell}")
    s.merge(row, 2, LAST_COL,
            "Type a PO number from the Purchase Orders tab",
            fmts.chrome("note"))
    row += 2

    s.merge(row, 0, 1, None, fmts.chrome("print_title"))
    s.formula(row, 0, '="PURCHASE ORDER"', fmts.chrome("print_title"))
    s.merge(row, 2, 4, None, fmts.calc("text"))
    s.formula(row, 2, f'=IF({po_cell}="","",{po_cell})',
              fmts.calc("id"))
    s.merge(row, 5, 6, "Date", fmts.chrome("label_right"))
    s.formula(
        row, 7,
        f'={_line("Purchase Orders", "po_print_rank", po_cell, "order_date", 1)}',
        fmts.calc("date"))
    row += 1

    s.merge(row, 0, 1, None, fmts.chrome("print_sub"))
    s.formula(row, 0, "=BusinessName", fmts.chrome("print_sub"))
    s.merge(row, 2, 4, "To", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, None, fmts.calc("text"))
    s.formula(
        row, 5,
        f'={_line("Purchase Orders", "po_print_rank", po_cell, "supplier_name", 1)}',
        fmts.calc("text"))
    row += 1

    for label, field in (("Expected delivery", "expected"),
                         ("Supplier reference", "supplier")):
        s.merge(row, 0, 1, label, fmts.chrome("label_bold"))
        style = "date" if field == "expected" else "id"
        s.formula(
            row, 2,
            f'={_line("Purchase Orders", "po_print_rank", po_cell, field, 1)}',
            fmts.calc(style))
        row += 1
    row += 1

    s.table_head(row, 0, ["SKU", "Product", "Their code", "Qty",
                          "Unit cost", "Discount", "Freight",
                          "Line total"],
                 widths=[16, 30, 15, 10, 13, 11, 11, 14])
    row += 1
    for n in range(1, PO_LINES + 1):
        for col, field, style in ((0, "sku", "id"),
                                  (1, "product", "text"),
                                  (2, "supplier_sku", "text"),
                                  (3, "qty", "int"),
                                  (4, "unit_cost", "money"),
                                  (5, "discount", "percent0"),
                                  (6, "freight", "money"),
                                  (7, "total", "money")):
            s.formula(
                row, col,
                f'={_line("Purchase Orders", "po_print_rank", po_cell, field, n)}',
                fmts.calc(style))
        row += 1
    s.merge(row, 0, 6, "Order total", fmts.chrome("total_label"))
    s.formula(row, 7,
              f'=IF({po_cell}="","",SUMIF(POId,{po_cell},POTotal))',
              fmts.chrome("total_value",
                          num_format=cfg.money_format))
    row += 1
    s.merge(row, 0, 6, "Already received", fmts.chrome("label_right"))
    s.formula(
        row, 7,
        f'=IF({po_cell}="","",SUMIF(POId,{po_cell},POReceivedValue))',
        fmts.calc("money"))
    row += 1
    s.merge(row, 0, 6, "Still outstanding", fmts.chrome("label_right"))
    s.formula(
        row, 7,
        f'=IF({po_cell}="","",SUMIF(POId,{po_cell},'
        f'POOutstandingValue))', fmts.calc("money"))
    row += 2

    s.merge(row, 0, LAST_COL,
            "Please confirm receipt of this order and advise if any "
            "line cannot be supplied in full.", fmts.chrome("note"))
    row += 2

    # =====================================================================
    # 2. WHOLESALE INVOICE
    # =====================================================================
    row = s.section(row, "Wholesale invoice - to send to a customer",
                    LAST_COL)
    s.merge(row, 0, 0, "Order ID", fmts.chrome("label_bold"))
    s.write(row, 1, None, fmts.input("id"))
    ord_cell = f"$B${row + 1}"
    book.define_name("PrintOrder", f"='{NAME}'!{ord_cell}")
    s.merge(row, 2, LAST_COL,
            "Type an Order ID from the Wholesale Orders tab",
            fmts.chrome("note"))
    row += 2

    s.merge(row, 0, 1, None, fmts.chrome("print_title"))
    s.formula(row, 0, '="INVOICE"', fmts.chrome("print_title"))
    s.merge(row, 2, 4, None, fmts.calc("text"))
    s.formula(
        row, 2,
        f'={_line("Wholesale Orders", "order_print_rank", ord_cell, "invoice_no", 1)}',
        fmts.calc("id"))
    s.merge(row, 5, 6, "Invoice date", fmts.chrome("label_right"))
    s.formula(
        row, 7,
        f'={_line("Wholesale Orders", "order_print_rank", ord_cell, "ship_date", 1)}',
        fmts.calc("date"))
    row += 1

    s.merge(row, 0, 1, None, fmts.chrome("print_sub"))
    s.formula(row, 0, "=BusinessName", fmts.chrome("print_sub"))
    s.merge(row, 2, 4, "Bill to", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, None, fmts.calc("text"))
    s.formula(
        row, 5,
        f'={_line("Wholesale Orders", "order_print_rank", ord_cell, "customer_name", 1)}',
        fmts.calc("text"))
    row += 1

    s.merge(row, 0, 1, "Order reference", fmts.chrome("label_bold"))
    s.formula(row, 2, f'=IF({ord_cell}="","",{ord_cell})',
              fmts.calc("id"))
    s.merge(row, 3, 4, "Their PO", fmts.chrome("label_right"))
    s.formula(
        row, 5,
        f'={_line("Wholesale Orders", "order_print_rank", ord_cell, "po_ref", 1)}',
        fmts.calc("text"))
    s.merge(row, 6, 6, "Due", fmts.chrome("label_right"))
    s.formula(
        row, 7,
        f'={_line("Wholesale Orders", "order_print_rank", ord_cell, "due_date", 1)}',
        fmts.calc("date"))
    row += 2

    s.table_head(row, 0, ["SKU", "Product", "Qty", "Unit price",
                          "Discount", "Line total"],
                 widths=[16, 34, 10, 14, 12, 15])
    row += 1
    for n in range(1, INV_LINES + 1):
        for col, field, style in ((0, "sku", "id"),
                                  (1, "product", "text"),
                                  (2, "qty", "int"),
                                  (3, "unit_price", "money"),
                                  (4, "line_discount", "percent0"),
                                  (5, "line_total", "money")):
            s.formula(
                row, col,
                f'={_line("Wholesale Orders", "order_print_rank", ord_cell, field, n)}',
                fmts.calc(style))
        row += 1

    goods = f'SUMIF(OrderId,{ord_cell},OrderTotal)'
    s.merge(row, 0, 4, "Goods total", fmts.chrome("label_right"))
    s.formula(row, 5, f'=IF({ord_cell}="","",{goods})',
              fmts.calc("money"))
    row += 1
    s.merge(row, 0, 4, "Shipping", fmts.chrome("label_right"))
    s.formula(
        row, 5,
        f'=IF({ord_cell}="","",SUMIF(OrderId,{ord_cell},'
        f'[[Wholesale Orders.shipping_charge]]))', fmts.calc("money"))
    ship_row = row
    row += 1
    s.merge(row, 0, 4, None, fmts.chrome("label_right"))
    s.formula(row, 0, '=TaxName&IF(TaxOnSales="Yes"," at "&'
                      'TEXT(TaxRate,"0%")," - not charged")',
              fmts.chrome("label_right"))
    s.formula(
        row, 5,
        f'=IF({ord_cell}="","",IF(TaxOnSales<>"Yes",0,'
        f'ROUND(({goods}+$F${ship_row + 1})*TaxRate,2)))',
        fmts.calc("money"))
    tax_row = row
    row += 1
    s.merge(row, 0, 4, "TOTAL DUE", fmts.chrome("total_label"))
    s.formula(
        row, 5,
        f'=IF({ord_cell}="","",{goods}+$F${ship_row + 1}+'
        f'$F${tax_row + 1})',
        fmts.chrome("total_value", num_format=cfg.money_format))
    row += 1
    s.merge(row, 0, 4, "Already paid", fmts.chrome("label_right"))
    s.formula(
        row, 5,
        f'=IF({ord_cell}="","",SUMIF(OrderId,{ord_cell},'
        f'[[Wholesale Orders.paid_amount]]))', fmts.calc("money"))
    row += 1
    s.merge(row, 0, 4, "Outstanding", fmts.chrome("label_right"))
    s.formula(
        row, 5,
        f'=IF({ord_cell}="","",SUMIF(OrderId,{ord_cell},'
        f'OrderOutstanding))', fmts.calc("money"))
    row += 2
    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    s.formula(
        row, 0,
        '="Payment terms: "&IFERROR(INDEX(CustomerTermsCol,'
        'MATCH(IFERROR(INDEX(OrderCustomer,MATCH(PrintOrder,'
        'OrderId,0)),""),CustomerList,0)),"as agreed")&'
        '". Please quote the order reference when paying."',
        fmts.chrome("note"))
    row += 2

    # =====================================================================
    # 3. PICKING LIST
    # =====================================================================
    row = s.section(row, "Picking list - for whoever packs the box",
                    LAST_COL)
    row = s.note(
        row,
        "Same order ID as the invoice above. This one shows where "
        "each item is stored and whether there is enough on the "
        "shelf to pick it, which is the only question the packer "
        "actually has.",
        LAST_COL, height=32)
    row += 1

    s.merge(row, 0, 1, None, fmts.chrome("print_title"))
    s.formula(row, 0, '="PICKING LIST"', fmts.chrome("print_title"))
    s.merge(row, 2, 4, None, fmts.calc("text"))
    s.formula(row, 2, f'=IF({ord_cell}="","",{ord_cell})',
              fmts.calc("id"))
    s.merge(row, 5, 6, "Customer", fmts.chrome("label_right"))
    s.formula(
        row, 7,
        f'={_line("Wholesale Orders", "order_print_rank", ord_cell, "customer_name", 1)}',
        fmts.calc("text"))
    row += 2

    s.table_head(row, 0, ["SKU", "Product", "Location", "Qty to "
                          "pick", "Available", "Enough?", "Picked"],
                 widths=[16, 30, 14, 12, 12, 12, 10])
    row += 1
    for n in range(1, PICK_LINES + 1):
        sku = _line("Wholesale Orders", "order_print_rank", ord_cell,
                    "sku", n)
        s.formula(row, 0, f'={sku}', fmts.calc("id"))
        s.formula(
            row, 1,
            f'={_line("Wholesale Orders", "order_print_rank", ord_cell, "product", n)}',
            fmts.calc("text"))
        s.formula(
            row, 2,
            f'=IFERROR(IF({sku}="","",INDEX(LocationCol,'
            f'MATCH({sku},LocationSKU,0))&""),"")',
            fmts.calc("text"))
        s.formula(
            row, 3,
            f'={_line("Wholesale Orders", "order_print_rank", ord_cell, "qty", n)}',
            fmts.calc("int"))
        s.formula(
            row, 4,
            f'=IFERROR(IF({sku}="","",INDEX(StockCurrent,'
            f'MATCH({sku},StockSKU,0))),"")', fmts.calc("int"))
        s.formula(
            row, 5,
            f'=IF({sku}="","",IF(IFERROR(INDEX(StockCurrent,'
            f'MATCH({sku},StockSKU,0)),0)>='
            f'IFERROR(INDEX([[Wholesale Orders.qty]],'
            f'MATCH({n},[[Wholesale Orders.order_print_rank]],0)),0),'
            f'"Yes","SHORT"))', fmts.calc("status"))
        s.write(row, 6, None, fmts.chrome("form_box"))
        s.cf(f"F{row + 1}:F{row + 1}",
             {"type": "text", "criteria": "containing",
              "value": "SHORT", "fmt": "bad"})
        row += 1
    row += 1
    s.merge(row, 0, 2, "Picked by", fmts.chrome("label_bold"))
    s.merge(row, 3, 5, None, fmts.chrome("form_line"))
    s.merge(row, 6, 6, "Date", fmts.chrome("label_right"))
    s.merge(row, 7, LAST_COL, None, fmts.chrome("form_line"))
    row += 2

    row = s.note(
        row,
        "Available on the picking list is the physical stock, not "
        "the available-to-sell figure - the packer wants to know "
        "whether the units are on the shelf, including the ones "
        "this very order has reserved.",
        LAST_COL, height=34)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

---

## 11. The quality tools

Not shipped with the workbook, but the reason it can be trusted: a static validator that checks the SKU join in both directions and holds the five technical decisions to account, a bracket and placeholder checker, a specification checker and a sample data checker that both run before the file is written, and a recalculation harness that really evaluates every formula.

### `tools/validate_inventory_workbook.py`

*1,238 lines · 18 functions*

> Static checks on a generated Ultimate Inventory & Wholesale Order Manager.
>
> Run it with any Python that has openpyxl installed::
>
>     python tools/validate_inventory_workbook.py output/inventory.xlsx
>
> It never opens Excel - it parses the XML and reports anything that would
> break for a buyer: unresolved template tokens, formulas pointing at the
> wrong sheet, defined names that do not exist, missing sheet protection,
> formula cells that are not locked, input cells that are, and - the one
> that matters most here - any break in the SKU chain that holds the
> product database, the stock tracker, the purchase orders, the wholesale
> orders, the sales, the valuation and the profit-by-product figures
> together.
>
> Five things this product cannot be allowed to get wrong, and all five
> are checked below:
>
> * stock is derived from transactions, never typed over;
> * reserved stock is separated from available stock;
> * the IDs connect - SKU, Supplier ID, Purchase Order ID, Order ID;
> * receiving the same purchase order twice cannot double-count;
> * one costing method, weighted average, used everywhere.

| Defines | What it does |
|---------|--------------|
| `_colnum(ref)` |  |
| `check_raw_xml(path)` | Overlapping merges and stray hyperlinks - Excel calls these corrupt |
| `paren_depth(formula)` | Open brackets minus closing ones, ignoring anything inside text |
| `_top_level(text, wanted)` | Positions of `wanted` characters outside brackets and strings |
| `_arguments(formula)` | Every comma separated argument, at every bracket depth |
| `precedence_slip(formula)` | `a+b/c` at the top level of a division - almost always a bug |
| `_span(text)` | (rows, columns) of an A1 range, or None if it is not one |
| `_calls(formula, names)` | Yield (function, [argument, ...]) for every call in *formula* |
| `ragged_criteria(formula)` | COUNTIFS/SUMIFS whose ranges are not all the same shape |
| `collapsed_criteria(formula)` | A COUNTIFS/SUMIFS criteria *range* that carries its own test |
| `_name_target(wb, name)` | The one cell a defined name points at, or None |
| `check_pages(wb, problems, notes)` | The summary pages, checked through their defined names |
| `_specs()` | Every register specification, keyed by sheet name |
| `check_chain(wb, problems, notes)` | The five workflows only hold while every one of these does |
| `check_sku_join(wb, problems, notes)` | The SKU is the join key |
| `check_decisions(wb, problems, notes)` | The five decisions this product is not allowed to get wrong |
| `check_promises(wb, problems, notes)` | Three promises this product makes to the buyer in writing |
| `check(path)` |  |

**Module constants:** `FUNCTION_RE` (re.compile(...)), `NAME_RE` (re.compile(...)), `BARE_COL_RE` (re.compile(...)), `KNOWN_FUNCTIONS` ({'IF', 'IFERROR', 'SUM', 'SUMIF', 'SUMIFS', 'COUNT', 'COU...), `MODERN_ONLY` ({'XLOOKUP', 'IFS', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCA...), `PRECEDENCE_OK` (0 entries), `TIME_DIVISOR_RE` (re.compile(...)), `COMPLEMENT_RE` (re.compile(...)), `DIVISOR_RE` (re.compile(...)), `RANGE_RE` (re.compile(...)), `SEEDED_INPUTS` ({'Monthly Reports!B5', 'Printable Orders & Invoices!B6', ...), `CHAIN_CHECKS` (104 items), `PAGE_FORMULAS` (9 items), `PAGE_INPUTS` (25 items) and 1 more.

````python
"""Static checks on a generated Ultimate Inventory & Wholesale Order Manager.

Run it with any Python that has openpyxl installed::

    python tools/validate_inventory_workbook.py output/inventory.xlsx

It never opens Excel - it parses the XML and reports anything that would
break for a buyer: unresolved template tokens, formulas pointing at the
wrong sheet, defined names that do not exist, missing sheet protection,
formula cells that are not locked, input cells that are, and - the one
that matters most here - any break in the SKU chain that holds the
product database, the stock tracker, the purchase orders, the wholesale
orders, the sales, the valuation and the profit-by-product figures
together.

Five things this product cannot be allowed to get wrong, and all five
are checked below:

* stock is derived from transactions, never typed over;
* reserved stock is separated from available stock;
* the IDs connect - SKU, Supplier ID, Purchase Order ID, Order ID;
* receiving the same purchase order twice cannot double-count;
* one costing method, weighted average, used everywhere.
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
from inventory_wholesale_spreadsheet.build import (                # noqa: E402
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


#: `1-x/y` is the complement of a ratio - an on-time rate, a stock
#: accuracy score, a discount depth. `(1-x)/y` is not a thing anybody
#: means by it.
COMPLEMENT_RE = re.compile(r"(?:^|[(,=<>&*+-])\s*\d+(?:\.\d+)?$")
#: the divisor after a `/`, so `a/n + b/n` - two terms averaged over
#: the same window - can be told apart from `(a+b)/n`
DIVISOR_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.!$']*|\d+(?:\.\d+)?")


def precedence_slip(formula):
    """`a+b/c` at the top level of a division - almost always a bug."""
    body = re.sub(r'"[^"]*"', '""', formula)
    for match in re.finditer(r"/", body):
        if TIME_DIVISOR_RE.match(body[match.start():]):
            continue
        divisor = DIVISOR_RE.match(body[match.start() + 1:])
        divisor = divisor.group(0) if divisor else ""
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
                left = body[:i]
                if ch == "-" and COMPLEMENT_RE.search(left):
                    break
                if ch == "+" and divisor and \
                        left.rstrip().endswith("/" + divisor):
                    break
                return body[max(0, i - 25):match.start() + 12]
            i -= 1
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
    "Monthly Reports!B5",                 # the month being reported on
    "Printable Orders & Invoices!B6",     # the PO being printed
    "Printable Orders & Invoices!B7",     # the order being printed
}

#: Every link below is a formula in row 5 of a register.  Together
#: they are the product: break one and a buyer gets a blank cell
#: instead of an answer, with nothing to tell them why.
#:
#: The SKU is the join key.  Every separately stocked variant has its
#: own SKU, sixteen registers reach back to the product database
#: through it, and stock itself is derived from four transaction
#: registers rather than typed anywhere.  Most of what follows is
#: those two decisions being checked in both directions.
#:
#: Every entry was read off the real formula, not guessed.  A guessed
#: chain check is worse than none: it fails on a correct product and
#: teaches you to ignore the tool.
CHAIN_CHECKS = [
    # ---- the catalogue: one row per SKU, the hub everything joins to ----
    ("Product & SKU Database", "id", ("text", "ProductPrefix")),
    ("Product & SKU Database", "on_hand", ("sheet", "Inventory Tracker")),
    ("Product & SKU Database", "available", ("sheet", "Inventory Tracker")),
    ("Product & SKU Database", "avg_cost", ("sheet",
                                            "Inventory Valuation & COGS")),
    ("Product & SKU Database", "units_sold", ("sheet",
                                              "Product Performance")),
    ("Product & SKU Database", "dup_count", ("column",
                                             "Product & SKU Database",
                                             "sku")),
    ("Product & SKU Database", "retail_margin", ("column",
                                                 "Product & SKU Database",
                                                 "cost")),

    # ---- workflow one: product -> SKU -> stock --------------------------
    # Stock is DERIVED. Every term below comes from a transaction
    # register; none of it is a number anybody types.
    ("Inventory Tracker", "sku", ("sheet", "Product & SKU Database")),
    ("Inventory Tracker", "opening", ("sheet", "Product & SKU Database")),
    ("Inventory Tracker", "received", ("sheet", "Purchase Orders")),
    ("Inventory Tracker", "sold", ("sheet", "Retail Sales Tracker")),
    ("Inventory Tracker", "returned", ("sheet", "Returns & Damaged Goods")),
    ("Inventory Tracker", "damaged", ("sheet", "Stock Movement Log")),
    ("Inventory Tracker", "adjusted", ("sheet", "Stock Movement Log")),
    ("Inventory Tracker", "current", ("column", "Inventory Tracker",
                                      "opening")),
    # ---- the reserved/available split, decision two ---------------------
    ("Inventory Tracker", "reserved", ("sheet", "Wholesale Orders")),
    ("Inventory Tracker", "available", ("column", "Inventory Tracker",
                                        "reserved")),
    ("Inventory Tracker", "on_order", ("sheet", "Purchase Orders")),
    ("Inventory Tracker", "daily_sales", ("text", "SalesWindow")),
    ("Inventory Tracker", "days_cover", ("column", "Inventory Tracker",
                                         "daily_sales")),
    ("Inventory Tracker", "stock_value", ("column", "Inventory Tracker",
                                          "avg_cost")),

    # ---- workflow two: supplier -> PO -> receiving ----------------------
    ("Purchase Orders", "supplier_name", ("sheet", "Supplier Database")),
    ("Purchase Orders", "product", ("sheet", "Product & SKU Database")),
    ("Purchase Orders", "due_date", ("sheet", "Supplier Database")),
    ("Purchase Orders", "landed_each", ("column", "Purchase Orders",
                                        "total")),
    ("Purchase Orders", "outstanding", ("column", "Purchase Orders",
                                        "received")),
    # the duplicate-receiving defence, decision four
    ("Purchase Orders", "dup_line", ("column", "Purchase Orders", "sku")),
    ("Purchase Orders", "received_value", ("column", "Purchase Orders",
                                           "landed_each")),
    ("Purchase Orders", "days_late", ("column", "Purchase Orders",
                                      "expected")),
    ("Purchase Orders", "alert", ("column", "Purchase Orders", "dup_line")),
    ("Supplier Database", "id", ("text", "SupplierPrefix")),
    ("Supplier Database", "orders", ("sheet", "Purchase Orders")),
    ("Supplier Database", "late_orders", ("sheet", "Purchase Orders")),
    ("Supplier Database", "actual_lead", ("sheet", "Purchase Orders")),
    ("Supplier Database", "lead_gap", ("column", "Supplier Database",
                                       "actual_lead")),
    ("Supplier Database", "on_time_rate", ("column", "Supplier Database",
                                           "late_orders")),

    # ---- workflow three: customer -> tiered price -> fulfilment ---------
    ("Wholesale Customer CRM", "id", ("text", "CustomerPrefix")),
    ("Wholesale Customer CRM", "revenue", ("sheet", "Wholesale Orders")),
    ("Wholesale Customer CRM", "owing", ("sheet", "Wholesale Orders")),
    ("Wholesale Customer CRM", "reserved_value", ("sheet",
                                                  "Wholesale Orders")),
    ("Wholesale Pricing", "sku", ("sheet", "Product & SKU Database")),
    ("Wholesale Pricing", "avg_cost", ("sheet", "Inventory Tracker")),
    ("Wholesale Pricing", "floor", ("sheet", "Product & SKU Database")),
    ("Wholesale Pricing", "tier3_gp", ("column", "Wholesale Pricing",
                                       "avg_cost")),
    ("Wholesale Orders", "customer_name", ("sheet",
                                           "Wholesale Customer CRM")),
    ("Wholesale Orders", "tier", ("sheet", "Wholesale Customer CRM")),
    ("Wholesale Orders", "tier_price", ("sheet", "Wholesale Pricing")),
    ("Wholesale Orders", "unit_cost", ("sheet", "Inventory Tracker")),
    ("Wholesale Orders", "available", ("sheet", "Inventory Tracker")),
    ("Wholesale Orders", "is_reserved", ("column", "Wholesale Orders",
                                         "status")),
    ("Wholesale Orders", "shortfall", ("column", "Wholesale Orders",
                                       "available")),
    ("Wholesale Orders", "line_profit", ("column", "Wholesale Orders",
                                         "line_cost")),
    ("Wholesale Orders", "due_date", ("column", "Wholesale Orders",
                                      "terms")),

    # ---- workflow four: movement -> reorder alert -> replenishment ------
    ("Stock Movement Log", "id", ("text", "MovementPrefix")),
    ("Stock Movement Log", "direction", ("text", "MovementSign")),
    ("Stock Movement Log", "cost_each", ("sheet", "Inventory Tracker")),
    ("Stock Movement Log", "stock_after", ("sheet", "Inventory Tracker")),
    ("Reorder Planner", "sku", ("sheet", "Product & SKU Database")),
    ("Reorder Planner", "available", ("sheet", "Inventory Tracker")),
    ("Reorder Planner", "daily_sales", ("sheet", "Inventory Tracker")),
    ("Reorder Planner", "quoted_lead", ("sheet", "Supplier Database")),
    ("Reorder Planner", "actual_lead", ("sheet", "Supplier Database")),
    ("Reorder Planner", "lead_days", ("column", "Reorder Planner",
                                      "quoted_lead")),
    ("Reorder Planner", "safety_stock", ("text", "SafetyDays")),
    ("Reorder Planner", "reorder_point", ("column", "Reorder Planner",
                                          "calc_point")),
    ("Reorder Planner", "order_by", ("text", "TodayDate")),
    ("Reorder Planner", "suggested", ("column", "Reorder Planner",
                                      "target")),

    # ---- workflow five: sales -> COGS -> profit by product --------------
    ("Retail Sales Tracker", "product", ("sheet",
                                         "Product & SKU Database")),
    ("Retail Sales Tracker", "unit_cost", ("sheet", "Inventory Tracker")),
    ("Retail Sales Tracker", "platform_fee", ("column",
                                              "Retail Sales Tracker",
                                              "fee_rate")),
    ("Retail Sales Tracker", "payment_fee", ("text", "PaymentFeePct")),
    ("Retail Sales Tracker", "cogs", ("column", "Retail Sales Tracker",
                                      "unit_cost")),
    ("Retail Sales Tracker", "profit", ("column", "Retail Sales Tracker",
                                        "cogs")),
    ("Inventory Valuation & COGS", "received_value", ("sheet",
                                                      "Purchase Orders")),
    ("Inventory Valuation & COGS", "avg_cost", ("sheet",
                                                "Purchase Orders")),
    ("Inventory Valuation & COGS", "cogs", ("column",
                                            "Inventory Valuation & COGS",
                                            "avg_cost")),
    ("Inventory Valuation & COGS", "gross_profit", ("column",
                                                    "Inventory Valuation "
                                                    "& COGS", "cogs")),
    ("Inventory Valuation & COGS", "value_share", ("text",
                                                   "TotalStockValue")),
    ("Inventory Valuation & COGS", "abc_class", ("column",
                                                 "Inventory Valuation "
                                                 "& COGS",
                                                 "value_share")),
    ("Product Performance", "units_retail", ("sheet",
                                             "Retail Sales Tracker")),
    ("Product Performance", "units_wholesale", ("sheet",
                                                "Wholesale Orders")),
    ("Product Performance", "cogs", ("sheet",
                                     "Inventory Valuation & COGS")),
    ("Product Performance", "returns_cost", ("sheet",
                                             "Returns & Damaged Goods")),
    ("Product Performance", "profit", ("column", "Product Performance",
                                       "cogs")),
    ("Product Performance", "turns", ("sheet",
                                      "Inventory Valuation & COGS")),

    # ---- the other registers --------------------------------------------
    ("Returns & Damaged Goods", "id", ("text", "ReturnPrefix")),
    ("Returns & Damaged Goods", "unit_cost", ("sheet",
                                              "Inventory Tracker")),
    ("Returns & Damaged Goods", "restocked", ("column",
                                              "Returns & Damaged Goods",
                                              "condition")),
    ("Warehouse Locations", "system_qty", ("sheet", "Inventory Tracker")),
    ("Warehouse Locations", "placed_gap", ("column", "Warehouse Locations",
                                           "system_qty")),
    ("Stock Audit", "id", ("text", "AuditPrefix")),
    ("Stock Audit", "system_qty", ("sheet", "Inventory Tracker")),
    ("Stock Audit", "variance", ("column", "Stock Audit", "system_qty")),
    ("Stock Audit", "variance_value", ("column", "Stock Audit",
                                       "variance")),
    ("Stock Audit", "unposted", ("column", "Stock Audit", "posted")),
    ("Expense Tracker", "is_stock", ("text", "StockCategoryList")),
    ("Expense Tracker", "operating", ("column", "Expense Tracker",
                                      "is_stock")),
    ("Profit Analysis", "retail_revenue", ("sheet",
                                           "Retail Sales Tracker")),
    ("Profit Analysis", "wholesale_revenue", ("sheet",
                                              "Wholesale Orders")),
    ("Profit Analysis", "expenses", ("sheet", "Expense Tracker")),
    ("Profit Analysis", "stock_spend", ("sheet", "Purchase Orders")),
    ("Profit Analysis", "gross_profit", ("column", "Profit Analysis",
                                         "cogs")),
    ("Profit Analysis", "net_profit", ("column", "Profit Analysis",
                                       "expenses")),
    ("Profit Analysis", "cash_effect", ("column", "Profit Analysis",
                                        "stock_spend")),
]

#: named cells that must still hold a live formula - a name pointing
#: at an empty cell means a page quietly stopped working
PAGE_FORMULAS = [
    ("TodayDate", "settings: the live date every alert reads"),
    ("YearStart", "settings: the first day of the reporting year"),
    ("YearEnd", "settings: the last day of the reporting year"),
    ("ThisMonthKey", "settings: the month key the dashboard groups on"),
    ("TotalStockValue", "settings: stock at cost, the ABC denominator"),
    ("TotalRevenue", "settings: revenue this year, both channels"),
    ("WholesaleRevenue", "settings: the wholesale half of revenue"),
    ("TotalPurchaseSpend", "settings: what has been spent on stock"),
    ("RetailRevenue", "settings: the retail half of revenue"),
    ]

#: cells a buyer has to be able to type in
PAGE_INPUTS = [
    ("BusinessName", "settings: the business name"),
    ("DefaultMarkup", "settings: the markup a new product starts at"),
    ("MinGP", "settings: the margin you refuse to go under"),
    ("TargetMargin", "settings: the margin you are aiming at"),
    ("DefaultLeadDays", "settings: lead time when a supplier has none"),
    ("SafetyDays", "settings: days of safety stock in the reorder point"),
    ("SalesWindow", "settings: the window the sales rate averages over"),
    ("MarketplaceFeePct", "settings: the marketplace commission"),
    ("PaymentFeePct", "settings: the payment processing percentage"),
    ("PaymentFeeFixed", "settings: the fixed part of a payment fee"),
    ("TaxRate", "settings: sales tax or VAT"),
    ("TaxOnSales", "settings: whether tax applies at all"),
    ("ReportYear", "settings: the year the reports cover"),
    ("YearlyTarget", "settings: the revenue target for the year"),
    ("TargetTurns", "settings: the stock turns you are aiming at"),
    ("PriceGapTolerance", "settings: when an invoice gap is worth a look"),
    ("SlowStockDays", "settings: when stock counts as slow moving"),
    ("ProductPrefix", "settings: the Product ID prefix"),
    ("DeadStockDays", "settings: when unsold stock counts as dead"),
    ("SupplierPrefix", "settings: the Supplier ID prefix"),
    ("CustomerPrefix", "settings: the Customer ID prefix"),
    ("MonthlyTarget", "settings: the revenue target a month"),
    ("ReportMonth", "monthly reports: the month being reported on"),
    ("PrintPO", "printables: the purchase order being printed"),
    ("PrintOrder", "printables: the wholesale order being printed"),
]

#: every drop down on the workbook reads one of these - a missing one
#: means a blank list and a column nobody can fill in
REQUIRED_LIST_NAMES = [
    "SKUList", "SupplierList", "CustomerList", "YesNoList",
    "ProductStatusList", "CategoryList", "SubcategoryList", "UnitList",
    "SupplierTypeList", "SupplierStatusList", "POStatusList",
    "DeliveryConditionList", "CustomerTypeList", "CustomerStatusList",
    "CustomerTierList", "LeadSourceList", "OrderStatusList",
    "ShippingMethodList", "ChannelList", "SaleStatusList",
    "MovementTypeList", "ReturnSourceList", "ReturnReasonList",
    "ReturnConditionList", "ReturnStatusList", "AuditTypeList",
    "AuditStatusList", "VarianceReasonList", "StorageTypeList",
    "LocationList", "StaffList", "ExpenseCategoryList",
    "StockCategoryList", "TaxCategoryList", "PaymentMethodList",
    "PaymentTermsList", "FrequencyList", "CurrencyList",
    "MonthNameList", "MonthKeyList",
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
    from inventory_wholesale_spreadsheet.build import REGISTER_MODULES
    from inventory_wholesale_spreadsheet.config import Config
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


def check_sku_join(wb, problems: list, notes: list) -> None:
    """The SKU is the join key. Verify it in both directions.

    This is the technical decision the whole product rests on: one SKU
    per separately stocked variant, listed once in the product
    database and chosen from a drop-down everywhere else.  Without it
    the purchase orders, the stock tracker, the wholesale orders, the
    sales and the valuation are five unrelated lists that happen to be
    in the same file.

    Four things have to hold:

    1. the product database really does mint a Product ID and carry a
       duplicate-SKU check, because a SKU typed twice silently splits
       one product's stock into two half-right rows;
    2. every register that joins back does so through a column
       validated against SKUList, so a buyer cannot mistype it;
    3. that drop-down resolves to a real defined name;
    4. the supplier, purchase order, customer and order IDs do the
       same job for their own chains.
    """
    specs = _specs()
    hub = specs.get("Product & SKU Database")
    if hub is None:
        problems.append("Product & SKU Database register has gone")
        return
    try:
        column = hub.index("id") + 1
    except Exception:                                  # noqa: B902
        problems.append("Product & SKU Database has no id column - "
                        "there is no generated Product ID left")
        return
    formula = wb["Product & SKU Database"].cell(row=5,
                                                column=column).value
    if not isinstance(formula, str) or "ProductPrefix" not in formula:
        problems.append("Product & SKU Database.id no longer builds "
                        "the Product ID from ProductPrefix")

    dup = next((c for c in hub.columns if c.key == "dup_count"), None)
    if dup is None or not dup.formula or "COUNTIF" not in dup.formula:
        problems.append("the duplicate-SKU check has gone from the "
                        "product database; one SKU entered twice "
                        "would now split a product's stock in half "
                        "with nothing to say so")

    defined = set(wb.defined_names)
    for ws in wb.worksheets:
        defined |= set(ws.defined_names)
    for name in ("SKUList", "SupplierList", "CustomerList"):
        if name not in defined:
            problems.append(f"{name} is missing - every drop-down "
                            f"that reads it is now empty")

    joined = []
    for name, spec in specs.items():
        if name == "Product & SKU Database":
            continue
        col = next((c for c in spec.columns if c.key == "sku"), None)
        if col is None or col.formula:
            continue                    # derived, not chosen
        joined.append(name)
        source = (col.validate or {}).get("source")
        if not (isinstance(source, str)
                and source.lstrip("=") == "SKUList"):
            problems.append(
                f"{name}.sku is not a SKUList drop-down, so a buyer "
                f"can mistype the SKU and silently disconnect the row "
                f"from its product, its cost and its stock")
    if len(joined) < 7:
        problems.append(f"only {len(joined)} registers join through a "
                        f"typed SKU; the design calls for at least 7")
    notes.append(f"SKU join: {len(joined)} registers joined to the "
                 f"product database through SKUList")


def check_decisions(wb, problems: list, notes: list) -> None:
    """The five decisions this product is not allowed to get wrong.

    Every one of them is a way an inventory spreadsheet quietly starts
    lying to its owner, and every one is checked against the real
    formulas rather than against a promise in the documentation.
    """
    specs = _specs()
    held = []

    # 1. stock is derived from transactions, never typed over
    tracker = specs.get("Inventory Tracker")
    if tracker is None:
        problems.append("the Inventory Tracker has gone")
    else:
        typed = [c.key for c in tracker.columns if not c.formula]
        if typed:
            problems.append(
                f"Inventory Tracker has typed columns {typed} - stock "
                f"must be derived from the transaction registers, "
                f"never overwritten by hand")
        else:
            held.append("stock derived, not typed")
        current = next((c for c in tracker.columns
                        if c.key == "current"), None)
        needed = ("opening", "received", "sold", "returned", "damaged",
                  "adjusted")
        if current is None or not all(f"{{{k}}}" in (current.formula or "")
                                      for k in needed):
            problems.append(
                "Inventory Tracker.current no longer adds up all six "
                "transaction terms, so the stock figure is no longer "
                "the sum of what actually happened")

        # 2. reserved stock is separated from available stock
        available = next((c for c in tracker.columns
                          if c.key == "available"), None)
        reserved = next((c for c in tracker.columns
                         if c.key == "reserved"), None)
        if available is None or reserved is None:
            problems.append("the reserved/available split has gone - a "
                            "confirmed order would be invisible until "
                            "it shipped")
        elif "{reserved}" not in (available.formula or ""):
            problems.append("Inventory Tracker.available no longer "
                            "subtracts reserved stock, so a confirmed "
                            "order can be sold twice")
        elif "is_reserved" not in (reserved.formula or ""):
            problems.append("Inventory Tracker.reserved no longer "
                            "reads the order status, so reserving has "
                            "stopped meaning anything")
        else:
            held.append("reserved separated from available")

    # 3. receiving the same PO twice cannot double-count
    po = specs.get("Purchase Orders")
    if po is None:
        problems.append("the Purchase Orders register has gone")
    else:
        dup = next((c for c in po.columns if c.key == "dup_line"), None)
        alert = next((c for c in po.columns if c.key == "alert"), None)
        if dup is None or "COUNTIFS" not in (dup.formula or ""):
            problems.append(
                "Purchase Orders has lost its duplicate-line check, so "
                "the same PO line can be received twice and the "
                "delivery counted twice into stock")
        elif alert is None or "{dup_line}" not in (alert.formula or ""):
            problems.append("the duplicate-line check is no longer "
                            "wired into the Purchase Orders alert, so "
                            "nobody would be told")
        else:
            held.append("duplicate receiving flagged")

    # 4. one costing method, used everywhere
    mixed = []
    for name, spec in specs.items():
        for col in spec.columns:
            text = (col.formula or "").upper()
            if "FIFO" in text or "LIFO" in text:
                mixed.append(f"{name}.{col.key}")
    if mixed:
        problems.append(f"a second costing method appears in {mixed} - "
                        f"this product is weighted average throughout "
                        f"and must not mix methods")
    else:
        held.append("one costing method (weighted average)")

    valuation = specs.get("Inventory Valuation & COGS")
    if valuation is not None:
        avg = next((c for c in valuation.columns
                    if c.key == "avg_cost"), None)
        if avg is None or "received_value" not in (avg.formula or ""):
            problems.append(
                "Inventory Valuation & COGS.avg_cost no longer "
                "weights cost by the value actually received, so it "
                "is not a weighted average any more")

    # 5. a count never overwrites stock - it raises a variance
    audit = specs.get("Stock Audit")
    if audit is not None:
        variance = next((c for c in audit.columns
                         if c.key == "variance"), None)
        posted = next((c for c in audit.columns
                       if c.key == "posted"), None)
        if variance is None or posted is None:
            problems.append("Stock Audit has lost the variance/posted "
                            "pair, which is what stops a count from "
                            "silently overwriting stock")
        else:
            held.append("counts raise a variance, never overwrite")

    notes.append(f"technical decisions held: {len(held)}/5 - "
                 + "; ".join(held))


def check_promises(wb, problems: list, notes: list) -> None:
    """Three promises this product makes to the buyer in writing.

    1. It is a workbook, not an integration: it must never claim to
       sync with Etsy, Shopify, Amazon or anything else.  Orders are
       typed in or imported by the owner, and saying otherwise is the
       single most common lie in this category.
    2. Buying stock is not an operating expense, and the workbook has
       to say so where a buyer will read it.
    3. The accounting figures carry a disclaimer pointing at a real
       accountant.
    """
    banned_sync = ("syncs with", "synchronise with", "synchronize with",
                   "live feed", "api connection", "automatically imports",
                   "real-time stock", "real-time inventory",
                   "connects to your shop", "connects to etsy",
                   "connects to shopify", "two-way sync",
                   "automatic sync")
    banned_magic = ("updates itself", "updates automatically from",
                    "pulls your orders", "imports your orders",
                    "downloads your sales")
    sync_hits = magic_hits = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or value.startswith("="):
                    continue
                low = value.lower()
                negated = ("not " in low or "never " in low
                           or "cannot" in low or "does not" in low
                           or "no " in low or "nothing " in low
                           or "without " in low or "doesn't" in low)
                for phrase in banned_sync:
                    if phrase in low and not negated:
                        sync_hits += 1
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: claims a "
                            f"live marketplace connection -> "
                            f"{value[:70]}")
                for phrase in banned_magic:
                    if phrase in low and not negated:
                        magic_hits += 1
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: claims "
                            f"orders arrive by themselves -> "
                            f"{value[:70]}")

    wanted = {
        "entered by you": ("you enter", "you type", "typed in",
                           "enter them", "paste", "import"),
        "no marketplace sync": ("does not connect", "not connect",
                                "no connection", "does not sync",
                                "not sync", "no sync"),
        "stock is not an expense": ("not an operating expense",
                                    "when it sells",
                                    "not a cost until",
                                    "separate from your operating",
                                    "kept out of"),
        "accountant": ("accountant",),
    }
    found = {key: False for key in wanted}
    for title in ("Getting Started", "Settings & Lists",
                  "Expense Tracker", "Profit Analysis", "Dashboard"):
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
                 f"{magic_hits} self-updating claims, honesty notes "
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
    check_sku_join(wb, problems, notes)
    check_decisions(wb, problems, notes)
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

### `tools/check_inventory_specs.py`

*206 lines · 5 functions*

> Fast structural sweep over the inventory product's register specs.
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
>     python3 tools/check_inventory_specs.py

| Defines | What it does |
|---------|--------------|
| `_strings_out(text)` |  |
| `_args(text, start)` | Split the argument list of a call whose '(' is at `start` |
| `collapsed_criteria(formula)` |  |
| `precedence_slip(formula)` | `a+b/c` at the top level of a division - almost always a bug |
| `main()` |  |

**Module constants:** `TOKEN` (re.compile(...)), `CALL` (re.compile(...)), `XREF` (re.compile(...)), `ENGINE_TOKENS` ({'r', 'r_prev', 'first', 'last'}), `TIME_DIVISORS` (8 items), `COMPLEMENT` (re.compile(...)), `DIVISOR` (re.compile(...)).

````python
#!/usr/bin/env python3
"""Fast structural sweep over the inventory product's register specs.

Runs in well under a second and catches, before a workbook is ever
built, the four mistakes that are expensive to find later:

1. a formula whose brackets do not balance (concatenated nested IFs
   lose a closing paren silently);
2. a `{token}` that names no column on that sheet;
3. `a+b/c` where `(a+b)/c` was meant;
4. a COUNTIFS/SUMIFS criteria range that carries its own test.

    python3 tools/check_inventory_specs.py
"""

from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from inventory_wholesale_spreadsheet.config import Config     # noqa: E402

TOKEN = re.compile(r"\{([a-z_][a-z0-9_]*)\}")
CALL = re.compile(r"\b(COUNTIFS|SUMIFS|AVERAGEIFS)\s*\(")
XREF = re.compile(r"\[\[([^\].#]+)\.([a-z_][a-z0-9_]*)(#\d*)?\]\]")

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


#: Dividing by one of these is a unit conversion, not a mis-bracket:
#: a date serial plus a fraction of a day, or a rate spread over a
#: fixed number of days. Bracketing those the other way would be the
#: actual bug.
TIME_DIVISORS = ("1440", "86400", "24", "60", "365", "30", "12", "7")


#: `1-x/y` is the complement of a ratio - an on-time rate, an accuracy
#: score, a discount depth. The left operand is a bare number, and
#: `(1-x)/y` is not a thing anybody means.
COMPLEMENT = re.compile(r"(?:^|[(,=<>&*+-])\s*\d+(?:\.\d+)?$")

#: the divisor that follows a `/`, so `a/n + b/n` - two terms each
#: averaged over the same window - can be told apart from `(a+b)/n`
DIVISOR = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*|\d+(?:\.\d+)?")


def precedence_slip(formula: str) -> str:
    """`a+b/c` at the top level of a division - almost always a bug."""
    body = _strings_out(formula)
    for match in re.finditer(r"/", body):
        tail = body[match.start() + 1:match.start() + 8]
        if re.match(r"(%s)\b" % "|".join(TIME_DIVISORS), tail):
            continue
        divisor = DIVISOR.match(body[match.start() + 1:])
        divisor = divisor.group(0) if divisor else ""
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
                left = body[:i]
                if ch == "-" and COMPLEMENT.search(left):
                    break           # 1-x/y, a proportion complement
                if ch == "+" and divisor and \
                        left.rstrip().endswith("/" + divisor):
                    break           # a/n+b/n, two terms, one window
                return body[max(0, i - 25):match.start() + 12]
            i -= 1
    return ""


def main() -> int:
    from inventory_wholesale_spreadsheet.build import REGISTER_MODULES

    cfg = Config()
    problems = 0
    total_cols = 0
    total_formulas = 0
    total_xrefs = 0

    # every register's column keys, so a [[Sheet.key]] pointing at a
    # column that does not exist is caught here and not in Excel
    catalogue = {m.spec(cfg).name: {c.key for c in m.spec(cfg).columns}
                 for m in REGISTER_MODULES}

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

            for sheet, key, _n in XREF.findall(text):
                total_xrefs += 1
                if sheet not in catalogue:
                    problems += 1
                    print(f"{where}: [[{sheet}.{key}]] - no such sheet")
                elif key not in catalogue[sheet]:
                    problems += 1
                    print(f"{where}: [[{sheet}.{key}]] - "
                          f"{sheet} has no column '{key}'")

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
          f"{total_formulas} formula columns, {total_xrefs} cross-sheet "
          f"references, {problems} problems")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
````

### `tools/check_inventory_sample.py`

*170 lines · 1 function*

> Validate the inventory product's example data before it is built.
>
> Five checks, every one of which has caught a real bug in this repo:
>
> 1. every key in a sample row is an *input* column on that register -
>    a value written into a formula column is silently dropped and takes
>    the formula for that row with it;
> 2. every value that feeds a drop-down exists in that drop-down's list -
>    otherwise Excel flags the cell the moment a buyer clicks it (199 of
>    these shipped undetected in an earlier product);
> 3. every ID that feeds a live list (SKU, Supplier ID, Customer ID,
>    location, staff name) actually exists on the register it points at -
>    one SKU typo makes every downstream lookup return blank, and on this
>    product a SKU is the join key for all seventeen registers;
> 4. the rows fit in the register;
> 5. nothing claims stock it cannot have: every SKU used by a purchase
>    order, order line, sale, return, movement, count or bin exists in
>    the product database.
>
>     python3 tools/check_inventory_sample.py

| Defines | What it does |
|---------|--------------|
| `main()` |  |

**Module constants:** `DYNAMIC_LISTS` (CROSS_SHEET_NAMES), `AUTO_IDS` (2 entries), `SKU_COLUMNS` (7 items).

````python
#!/usr/bin/env python3
"""Validate the inventory product's example data before it is built.

Five checks, every one of which has caught a real bug in this repo:

1. every key in a sample row is an *input* column on that register -
   a value written into a formula column is silently dropped and takes
   the formula for that row with it;
2. every value that feeds a drop-down exists in that drop-down's list -
   otherwise Excel flags the cell the moment a buyer clicks it (199 of
   these shipped undetected in an earlier product);
3. every ID that feeds a live list (SKU, Supplier ID, Customer ID,
   location, staff name) actually exists on the register it points at -
   one SKU typo makes every downstream lookup return blank, and on this
   product a SKU is the join key for all seventeen registers;
4. the rows fit in the register;
5. nothing claims stock it cannot have: every SKU used by a purchase
   order, order line, sale, return, movement, count or bin exists in
   the product database.

    python3 tools/check_inventory_sample.py
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from inventory_wholesale_spreadsheet.config import (Config,          # noqa: E402
                                                    MOVEMENT_TYPES,
                                                    PAYMENT_TERMS)
from inventory_wholesale_spreadsheet.build import REGISTER_MODULES   # noqa: E402
from inventory_wholesale_spreadsheet.names import (                  # noqa: E402
    CROSS_SHEET_NAMES)
from inventory_wholesale_spreadsheet.sheets.settings import LISTS    # noqa: E402

#: which register and column each live drop-down reads from
DYNAMIC_LISTS = CROSS_SHEET_NAMES

#: sheets whose ID column is generated, and the config key that names
#: the prefix - so the live list can be rebuilt from the row count
AUTO_IDS = {
    "Supplier Database": "supplier",
    "Wholesale Customer CRM": "customer",
}

#: every register that names a SKU, and the column it uses
SKU_COLUMNS = [
    ("Stock Movement Log", "sku"),
    ("Purchase Orders", "sku"),
    ("Wholesale Orders", "sku"),
    ("Retail Sales Tracker", "sku"),
    ("Returns & Damaged Goods", "sku"),
    ("Warehouse Locations", "sku"),
    ("Stock Audit", "sku"),
]


def main() -> int:
    cfg = Config()
    lists = {name: set(values) for name, values in LISTS}
    # the Settings blocks written row by row rather than as a list
    lists["MovementTypeList"] = {name for name, _sign, _blurb
                                 in MOVEMENT_TYPES}
    lists["PaymentTermsList"] = {label for label, _days in PAYMENT_TERMS}

    specs = {}
    for module in REGISTER_MODULES:
        spec = module.spec(cfg)
        specs[spec.name] = spec

    # Build the live contents of every dynamic list out of the sample
    # data itself, so a SKU typo in Purchase Orders is caught here
    # rather than by a buyer wondering why a row is blank.
    live: dict = {}
    for list_name, (sheet, key) in DYNAMIC_LISTS.items():
        spec = specs.get(sheet)
        if spec is None:
            continue
        column = next((c for c in spec.columns if c.key == key), None)
        values = set()
        if column is not None and column.formula:
            prefix_key = AUTO_IDS.get(sheet)
            if prefix_key:
                digits = cfg.digits(prefix_key)
                prefix = cfg.prefix(prefix_key)
                values = {f"{prefix}{n:0{digits}d}"
                          for n in range(1, len(spec.samples) + 1)}
        else:
            values = {str(row.get(key, "")) for row in spec.samples
                      if row.get(key, "") != ""}
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
                    continue            # an inline list, e.g. Yes/No
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

        print(f"  {spec.name:26s} {len(rows):4d} rows / {spec.rows:4d}, "
              f"{len(inputs)} input columns")

    # ---- the join key the whole product turns on -------------------
    catalogue = {str(row.get("sku", ""))
                 for row in specs["Product & SKU Database"].samples}
    orphans = 0
    for sheet, key in SKU_COLUMNS:
        spec = specs.get(sheet)
        if spec is None:
            continue
        for index, row in enumerate(spec.samples):
            sku = str(row.get(key, ""))
            if sku and sku not in catalogue:
                problems += 1
                orphans += 1
                print(f"{sheet} row {index}: SKU {sku!r} is not in the "
                      f"product database")
    if not orphans:
        print(f"\n  every SKU on every register resolves to one of the "
              f"{len(catalogue)} catalogue products")

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

#### Product & SKU Database

One row per stocked thing. A variant you count separately is a separate SKU - that is the whole rule.

41 columns, 300 rows of capacity, 27 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Product ID | `id` | id | formula |
| B | SKU | `sku` | id | you type it |
| C | Product name | `name` | text | you type it |
| D | Description | `description` | wrap | you type it |
| E | Category | `category` | text | drop down `CategoryList` |
| F | Subcategory | `subcategory` | text | drop down `SubcategoryList` |
| G | Brand | `brand` | text | you type it |
| H | Variant | `variant` | text | you type it |
| I | Size | `size` | text | you type it |
| J | Colour | `colour` | text | you type it |
| K | Material | `material` | text | you type it |
| L | Barcode | `barcode` | text | you type it |
| M | Unit | `unit` | text | drop down `UnitList` |
| N | Image link | `image` | link | you type it |
| O | Supplier ID | `supplier` | id | drop down `SupplierList` |
| P | Their code | `supplier_sku` | text | you type it |
| Q | Unit cost | `cost` | money | you type it (checked) |
| R | Retail price | `retail_price` | money | you type it (checked) |
| S | Wholesale | `wholesale_price` | money | you type it (checked) |
| T | Floor price | `min_price` | money | you type it (checked) |
| U | Retail GP% | `retail_margin` | percent | formula |
| V | Trade GP% | `wholesale_margin` | percent | formula |
| W | Markup | `markup` | number1 | formula |
| X | Opening stock | `opening_qty` | int | you type it (checked) |
| Y | Opening value | `opening_value` | money | you type it (checked) |
| Z | Min level | `min_stock` | int | you type it (checked) |
| AA | Max level | `max_stock` | int | you type it (checked) |
| AB | Reorder at | `reorder_point` | int | you type it (checked) |
| AC | Status | `status` | status | drop down `ProductStatusList` |
| AD | Launched | `launched` | date | you type it (checked) |
| AE | Notes | `notes` | wrap | you type it |
| AF | Full name | `variant_label` | text | formula, hidden |
| AG | Dups | `dup_count` | int | formula, hidden |
| AH | Active | `is_active` | int | formula, hidden |
| AI | In stock | `on_hand` | int | formula |
| AJ | Available | `available` | int | formula |
| AK | Stock value | `stock_value` | money | formula |
| AL | Avg cost | `avg_cost` | money4 | formula |
| AM | Sold | `units_sold` | int | formula, hidden |
| AN | Rank | `value_rank` | int | formula, hidden |
| AO | What needs doing | `alert` | text | formula |

#### Inventory Tracker

Every quantity here is calculated. There is nothing to type on this sheet - and no way to quietly overwrite a stock figure.

33 columns, 300 rows of capacity, 0 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | SKU | `sku` | id | formula |
| B | Product | `product` | text | formula |
| C | Category | `category` | text | formula |
| D | Unit | `unit` | text | formula |
| E | Product status | `status` | status | formula |
| F | Location | `location` | text | formula |
| G | Opening | `opening` | int | formula |
| H | Received | `received` | int | formula |
| I | Sold | `sold` | int | formula |
| J | Returned in | `returned` | int | formula |
| K | Damaged | `damaged` | int | formula |
| L | Adjusted | `adjusted` | int | formula |
| M | In stock | `current` | int | formula |
| N | Reserved | `reserved` | int | formula |
| O | Available | `available` | int | formula |
| P | On order | `on_order` | int | formula |
| Q | Next PO row | `next_po_row` | int | formula, hidden |
| R | Next delivery | `incoming` | date | formula |
| S | Avg cost | `avg_cost` | money4 | formula |
| T | Stock value | `stock_value` | money | formula |
| U | Retail value | `retail_value` | money | formula, hidden |
| V | Min | `min_stock` | int | formula, hidden |
| W | Max | `max_stock` | int | formula, hidden |
| X | Sold/day | `daily_sales` | number1 | formula |
| Y | Days cover | `days_cover` | int | formula |
| Z | Stock level | `stock_state` | status | formula |
| AA | Last moved (n) | `last_move_num` | number | formula, hidden |
| AB | Last moved | `last_movement` | date | formula |
| AC | Out | `is_out` | int | formula, hidden |
| AD | Low | `is_low` | int | formula, hidden |
| AE | Over | `is_over` | int | formula, hidden |
| AF | Dead | `is_dead` | int | formula, hidden |
| AG | What needs doing | `alert` | text | formula |

#### Stock Movement Log

Every stock change that is not a purchase, a sale or a return - with a reason attached to it.

25 columns, 900 rows of capacity, 9 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Movement ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | SKU | `sku` | id | drop down `SKUList` |
| D | Product name | `product` | text | formula |
| E | Movement type | `type` | text | drop down `MovementTypeList` |
| F | Quantity | `qty` | int | you type it (checked) |
| G | In / out | `direction` | int | formula, hidden |
| H | Change | `signed_qty` | int | formula |
| I | Reason | `reason` | wrap | you type it |
| J | Reference | `reference` | text | you type it |
| K | Location | `location` | text | drop down `LocationList` |
| L | Moved to | `to_location` | text | drop down `LocationList` |
| M | Recorded by | `recorded_by` | text | drop down `StaffList` |
| N | Cost each | `cost_each` | money | formula |
| O | Value | `value` | money | formula |
| P | Damaged qty | `damaged` | int | formula, hidden |
| Q | Adjustment qty | `adjustment` | int | formula, hidden |
| R | Write-off | `is_writeoff` | int | formula, hidden |
| S | Manual | `is_manual` | int | formula, hidden |
| T | Month | `month` | text | formula, hidden |
| U | date (number) | `date_num` | number | formula, hidden |
| V | Row number | `seq_num` | int | formula, hidden |
| W | Stock after | `stock_after` | int | formula |
| X | Check | `flag` | status | formula |
| Y | What needs doing | `alert` | text | formula |

#### Supplier Database

Contacts and terms on the left; how they actually perform, measured from your own orders, on the right.

39 columns, 60 rows of capacity, 23 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Supplier ID | `id` | id | formula |
| B | Supplier name | `name` | text | you type it |
| C | Contact | `contact` | text | you type it |
| D | Email | `email` | text | you type it |
| E | Phone | `phone` | text | you type it |
| F | Website | `website` | link | you type it |
| G | Address | `address` | wrap | you type it |
| H | City | `city` | text | you type it |
| I | Country | `country` | text | you type it |
| J | Supplier type | `type` | text | drop down `SupplierTypeList` |
| K | What they supply | `supplies` | wrap | you type it |
| L | Payment terms | `terms` | text | drop down `PaymentTermsList` |
| M | Currency | `currency` | text | drop down `CurrencyList` |
| N | Lead time (days) | `lead_time` | int | you type it (checked) |
| O | Min order | `moq` | money | you type it (checked) |
| P | Free ship over | `free_ship_over` | money | you type it (checked) |
| Q | Standard discount | `discount` | percent0 | you type it (checked) |
| R | Account no. | `account_no` | text | you type it |
| S | Tax / VAT no. | `tax_id` | text | you type it |
| T | Status | `status` | status | drop down `SupplierStatusList` |
| U | Preferred | `preferred` | flag | you type it (checked) |
| V | Rating | `rating` | rating | you type it (checked) |
| W | First order | `first_order` | date | you type it (checked) |
| X | Notes | `notes` | wrap | you type it |
| Y | Orders | `orders` | int | formula |
| Z | Total spend | `spend` | money | formula |
| AA | Spend this year | `spend_year` | money | formula |
| AB | Open | `open_orders` | int | formula |
| AC | Open value | `open_value` | money | formula |
| AD | Late | `late_orders` | int | formula |
| AE | On time | `on_time_rate` | percent0 | formula |
| AF | Actual lead | `actual_lead` | number1 | formula |
| AG | Lead gap | `lead_gap` | number1 | formula |
| AH | Last order | `last_order` | date | formula |
| AI | Last order (n) | `last_order_num` | number | formula, hidden |
| AJ | Days since | `days_since` | int | formula |
| AK | Products | `products` | int | formula |
| AL | Share of spend | `share` | percent0 | formula, hidden |
| AM | What needs doing | `alert` | text | formula |

#### Purchase Orders

One row per order line. One receiving cell per line - which is why the same delivery cannot be booked in twice.

48 columns, 500 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | PO number | `po_id` | id | you type it |
| B | Line | `line_no` | int | formula |
| C | Supplier ID | `supplier` | id | drop down `SupplierList` |
| D | Supplier | `supplier_name` | text | formula |
| E | Ordered | `order_date` | date | you type it (checked) |
| F | Expected | `expected` | date | you type it (checked) |
| G | SKU | `sku` | id | drop down `SKUList` |
| H | Product | `product` | text | formula |
| I | Their code | `supplier_sku` | text | formula |
| J | Qty ordered | `qty` | int | you type it (checked) |
| K | Unit cost | `unit_cost` | money | you type it (checked) |
| L | Discount | `discount` | percent0 | you type it (checked) |
| M | Freight | `freight` | money | you type it (checked) |
| N | Goods | `goods_value` | money | formula |
| O | Line total | `total` | money | formula |
| P | Landed each | `landed_each` | money4 | formula |
| Q | Qty received | `received` | int | you type it (checked) |
| R | Received | `received_date` | date | you type it (checked) |
| S | Checked by | `received_by` | text | drop down `StaffList` |
| T | Condition | `condition` | text | drop down `DeliveryConditionList` |
| U | Received value | `received_value` | money | formula |
| V | Still due | `outstanding` | int | formula |
| W | Due value | `outstanding_value` | money | formula, hidden |
| X | Status | `status` | status | drop down `POStatusList` |
| Y | Suggested | `suggested_status` | text | formula |
| Z | Invoice no. | `invoice_no` | text | you type it |
| AA | Invoiced | `invoice_amount` | money | you type it (checked) |
| AB | Paid | `paid` | flag | you type it (checked) |
| AC | Payment due | `due_date` | date | formula |
| AD | Notes | `notes` | wrap | you type it |
| AE | Duplicate line | `dup_line` | int | formula, hidden |
| AF | Open | `is_open` | int | formula, hidden |
| AG | Committed | `is_committed` | int | formula, hidden |
| AH | Received? | `is_received` | int | formula, hidden |
| AI | Over | `is_over` | int | formula, hidden |
| AJ | Days late | `days_late` | int | formula |
| AK | Late | `is_late` | int | formula, hidden |
| AL | Lead days | `lead_days` | int | formula, hidden |
| AM | Price gap | `price_gap` | money | formula |
| AN | Cost drift | `cost_drift` | percent0 | formula |
| AO | Month | `month` | text | formula, hidden |
| AP | order_date (number) | `order_num` | number | formula, hidden |
| AQ | received_date (number) | `received_num` | number | formula, hidden |
| AR | Row number | `seq_num` | int | formula, hidden |
| AS | Print line | `po_print_rank` | int | formula, hidden |
| AT | Chase key | `chase_key` | number | formula, hidden |
| AU | Chase rank | `chase_rank` | int | formula, hidden |
| AV | What needs doing | `alert` | text | formula |

#### Wholesale Customer CRM

Stockists, their terms and their tier - and, on the right, whether the account is actually worth having.

42 columns, 120 rows of capacity, 24 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Customer ID | `id` | id | formula |
| B | Business name | `business` | text | you type it |
| C | Buyer | `contact` | text | you type it |
| D | Email | `email` | text | you type it |
| E | Phone | `phone` | text | you type it |
| F | Website / shop | `website` | link | you type it |
| G | Address | `address` | wrap | you type it |
| H | City | `city` | text | you type it |
| I | Region | `region` | text | you type it |
| J | Postcode | `postcode` | text | you type it |
| K | Country | `country` | text | you type it |
| L | Business type | `type` | text | drop down `CustomerTypeList` |
| M | Price tier | `tier` | text | drop down `CustomerTierList` |
| N | Payment terms | `terms` | text | drop down `PaymentTermsList` |
| O | Credit limit | `credit_limit` | money | you type it (checked) |
| P | Min order | `min_order` | money | you type it (checked) |
| Q | Extra discount | `discount` | percent0 | you type it (checked) |
| R | Tax exempt | `tax_exempt` | flag | you type it (checked) |
| S | Resale cert. no. | `resale_no` | text | you type it |
| T | Shipping | `shipping_pref` | text | drop down `ShippingMethodList` |
| U | Status | `status` | status | drop down `CustomerStatusList` |
| V | Rating | `rating` | rating | you type it (checked) |
| W | First order | `first_order` | date | you type it (checked) |
| X | How they found you | `source` | text | drop down `LeadSourceList` |
| Y | Notes | `notes` | wrap | you type it |
| Z | Orders | `orders` | int | formula |
| AA | Units | `units` | int | formula, hidden |
| AB | Total revenue | `revenue` | money | formula |
| AC | This year | `revenue_year` | money | formula |
| AD | Gross profit | `profit` | money | formula |
| AE | GP% | `gp_pct` | percent0 | formula |
| AF | Avg order | `avg_order` | money | formula |
| AG | Owing now | `owing` | money | formula |
| AH | Overdue | `overdue` | money | formula |
| AI | Credit used | `credit_used` | percent0 | formula |
| AJ | Reserved for them | `reserved_value` | money | formula, hidden |
| AK | Last order (n) | `last_order_num` | number | formula, hidden |
| AL | Last order | `last_order` | date | formula |
| AM | Days since | `days_since` | int | formula |
| AN | Rank | `revenue_rank` | int | formula, hidden |
| AO | Share | `share` | percent0 | formula, hidden |
| AP | What needs doing | `alert` | text | formula |

#### Wholesale Pricing

Three quantity breaks per SKU, with the margin at each one shown next to it - because a price without its margin is just a hope.

31 columns, 300 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | SKU | `sku` | id | formula |
| B | Product | `product` | text | formula |
| C | Category | `category` | text | formula, hidden |
| D | Product status | `status` | status | formula, hidden |
| E | Avg cost | `avg_cost` | money4 | formula |
| F | Retail price | `retail` | money | formula |
| G | Floor price | `floor` | money | formula |
| H | Base wholesale | `wholesale` | money | you type it (checked) |
| I | Base GP% | `wholesale_gp` | percent0 | formula |
| J | Tier 1 from | `tier1_qty` | int | you type it (checked) |
| K | Tier 1 price | `tier1_price` | money | you type it (checked) |
| L | Tier 1 GP% | `tier1_gp` | percent0 | formula |
| M | Tier 2 from | `tier2_qty` | int | you type it (checked) |
| N | Tier 2 price | `tier2_price` | money | you type it (checked) |
| O | Tier 2 GP% | `tier2_gp` | percent0 | formula |
| P | Tier 3 from | `tier3_qty` | int | you type it (checked) |
| Q | Tier 3 price | `tier3_price` | money | you type it (checked) |
| R | Tier 3 GP% | `tier3_gp` | percent0 | formula |
| S | Best price | `best_price` | money | formula |
| T | Best GP% | `best_gp` | percent0 | formula |
| U | Discount | `discount_depth` | percent0 | formula, hidden |
| V | Of retail | `retail_ratio` | percent0 | formula, hidden |
| W | MOQ | `moq` | int | you type it (checked) |
| X | Case qty | `case_qty` | int | you type it (checked) |
| Y | Sell to trade | `available_trade` | flag | you type it (checked) |
| Z | Last reviewed | `reviewed` | date | you type it (checked) |
| AA | Notes | `notes` | wrap | you type it |
| AB | Below floor | `below_floor` | int | formula, hidden |
| AC | Thin | `thin_margin` | int | formula, hidden |
| AD | Priced | `is_priced` | int | formula, hidden |
| AE | What needs doing | `alert` | text | formula |

#### Wholesale Orders

Lines price themselves off the customer's tier. The status column decides whether stock is reserved or gone.

50 columns, 600 rows of capacity, 15 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Order ID | `order_id` | id | you type it |
| B | Line | `line_no` | int | formula |
| C | First line | `is_first_line` | int | formula, hidden |
| D | Order date | `date` | date | you type it (checked) |
| E | Customer ID | `customer` | id | drop down `CustomerList` |
| F | Customer | `customer_name` | text | formula |
| G | Tier | `tier` | text | formula |
| H | Their PO ref | `po_ref` | text | you type it |
| I | SKU | `sku` | id | drop down `SKUList` |
| J | Product | `product` | text | formula |
| K | Qty | `qty` | int | you type it (checked) |
| L | Tier price | `tier_price` | money | formula |
| M | Acct disc. | `cust_discount` | percent0 | formula, hidden |
| N | Line disc. | `line_discount` | percent0 | you type it (checked) |
| O | Unit price | `unit_price` | money | formula |
| P | Line total | `line_total` | money | formula |
| Q | Unit cost | `unit_cost` | money4 | formula, hidden |
| R | Line cost | `line_cost` | money | formula, hidden |
| S | Line profit | `line_profit` | money | formula |
| T | GP% | `line_gp` | percent0 | formula |
| U | Status | `status` | status | drop down `OrderStatusList` |
| V | Available | `available` | int | formula |
| W | Short by | `shortfall` | int | formula |
| X | Shipped | `ship_date` | date | you type it (checked) |
| Y | Tracking | `tracking` | text | you type it |
| Z | Method | `ship_method` | text | drop down `ShippingMethodList` |
| AA | Shipping | `shipping_charge` | money | you type it (checked) |
| AB | Invoice no. | `invoice_no` | text | you type it |
| AC | Terms | `terms` | text | formula, hidden |
| AD | Payment due | `due_date` | date | formula |
| AE | Paid | `paid_amount` | money | you type it (checked) |
| AF | Outstanding | `outstanding` | money | formula |
| AG | Overdue | `is_overdue` | int | formula, hidden |
| AH | Days over | `days_overdue` | int | formula |
| AI | Notes | `notes` | wrap | you type it |
| AJ | Reserved | `is_reserved` | int | formula, hidden |
| AK | Shipped? | `is_shipped` | int | formula, hidden |
| AL | Open | `is_open` | int | formula, hidden |
| AM | Units out | `units_shipped` | int | formula, hidden |
| AN | Days to ship | `days_to_ship` | int | formula |
| AO | Age | `age_days` | int | formula |
| AP | Below MOQ | `below_min` | int | formula, hidden |
| AQ | Month | `month` | text | formula, hidden |
| AR | date (number) | `date_num` | number | formula, hidden |
| AS | ship_date (number) | `ship_num` | number | formula, hidden |
| AT | Row number | `seq_num` | int | formula, hidden |
| AU | Print line | `order_print_rank` | int | formula, hidden |
| AV | Pick key | `pick_key` | number | formula, hidden |
| AW | Pick rank | `pick_rank` | int | formula, hidden |
| AX | What needs doing | `alert` | text | formula |

#### Retail Sales Tracker

Every single sale, with what the platform took out of it. Typed or pasted in - nothing here syncs with a marketplace.

34 columns, 900 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Sale ID | `sale_id` | id | you type it |
| B | Date | `date` | date | you type it (checked) |
| C | Channel | `channel` | text | drop down `ChannelList` |
| D | SKU | `sku` | id | drop down `SKUList` |
| E | Product | `product` | text | formula |
| F | Qty | `qty` | int | you type it (checked) |
| G | Unit price | `unit_price` | money | you type it (checked) |
| H | Price used | `price_used` | money | formula, hidden |
| I | Gross sale | `gross` | money | formula |
| J | Ship charged | `shipping_charged` | money | you type it (checked) |
| K | Ship cost | `shipping_cost` | money | you type it (checked) |
| L | Discount | `discount` | money | you type it (checked) |
| M | Fee rate | `fee_rate` | percent0 | formula, hidden |
| N | Platform fee | `platform_fee` | money | formula |
| O | Fee override | `fee_override` | money | you type it (checked) |
| P | Payment fee | `payment_fee` | money | formula |
| Q | Packaging | `packaging` | money | formula, hidden |
| R | Net revenue | `net_revenue` | money | formula |
| S | Unit cost | `unit_cost` | money4 | formula, hidden |
| T | COGS | `cogs` | money | formula |
| U | Net profit | `profit` | money | formula |
| V | Margin | `margin` | percent0 | formula |
| W | Status | `status` | status | drop down `SaleStatusList` |
| X | Customer | `customer` | text | you type it |
| Y | Sold from | `location` | text | drop down `LocationList` |
| Z | Notes | `notes` | wrap | you type it |
| AA | Sold? | `is_sold` | int | formula, hidden |
| AB | Loss | `is_loss` | int | formula, hidden |
| AC | Ship gap | `ship_gap` | money | formula, hidden |
| AD | Month | `month` | text | formula, hidden |
| AE | Day | `weekday` | text | formula |
| AF | date (number) | `date_num` | number | formula, hidden |
| AG | Row number | `seq_num` | int | formula, hidden |
| AH | What needs doing | `alert` | text | formula |

#### Reorder Planner

(average daily sales x real lead time) + safety stock. Everything on this sheet is calculated.

33 columns, 300 rows of capacity, 0 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | SKU | `sku` | id | formula |
| B | Product | `product` | text | formula |
| C | Category | `category` | text | formula, hidden |
| D | Product status | `status` | status | formula, hidden |
| E | Supplier ID | `supplier` | id | formula, hidden |
| F | Supplier | `supplier_name` | text | formula |
| G | Available | `available` | int | formula |
| H | On order | `on_order` | int | formula |
| I | Stock position | `position` | int | formula |
| J | Sold/day | `daily_sales` | number1 | formula |
| K | Quoted lead | `quoted_lead` | int | formula, hidden |
| L | Real lead | `actual_lead` | number1 | formula, hidden |
| M | Lead time | `lead_days` | number1 | formula |
| N | Safety stock | `safety_stock` | int | formula |
| O | Calculated ROP | `calc_point` | int | formula |
| P | Your ROP | `set_point` | int | formula, hidden |
| Q | Reorder at | `reorder_point` | int | formula |
| R | Max level | `max_stock` | int | formula, hidden |
| S | Order up to | `target` | int | formula, hidden |
| T | Order qty | `suggested` | int | formula |
| U | Case | `case_qty` | int | formula, hidden |
| V | Rounded | `order_rounded` | int | formula |
| W | Unit cost | `unit_cost` | money | formula, hidden |
| X | Order value | `order_value` | money | formula |
| Y | Days cover | `days_cover` | int | formula |
| Z | Runs out | `stockout_date` | date | formula |
| AA | Order by | `order_by` | date | formula |
| AB | Urgency | `urgency` | status | formula |
| AC | Rank | `priority` | int | formula, hidden |
| AD | To order | `needs_order` | int | formula, hidden |
| AE | Buy key | `buy_key` | number | formula, hidden |
| AF | Buy rank | `buy_rank` | int | formula, hidden |
| AG | What needs doing | `alert` | text | formula |

#### Inventory Valuation & COGS

Weighted average cost, shown term by term. One method, used everywhere, with no exceptions.

28 columns, 300 rows of capacity, 0 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | SKU | `sku` | id | formula |
| B | Product | `product` | text | formula |
| C | Category | `category` | text | formula |
| D | Status | `status` | status | formula, hidden |
| E | Opening units | `opening_qty` | int | formula |
| F | Opening value | `opening_value` | money | formula |
| G | Units received | `received_qty` | int | formula |
| H | Value received | `received_value` | money | formula |
| I | Units in | `total_qty` | int | formula, hidden |
| J | Value in | `total_value` | money | formula, hidden |
| K | Average cost | `avg_cost` | money4 | formula |
| L | Standard cost | `book_cost` | money | formula, hidden |
| M | Drift | `cost_drift` | percent0 | formula |
| N | Units sold | `units_sold` | int | formula |
| O | Units held | `on_hand` | int | formula |
| P | Stock at cost | `stock_value` | money | formula |
| Q | COGS | `cogs` | money | formula |
| R | Revenue | `revenue` | money | formula |
| S | Gross profit | `gross_profit` | money | formula |
| T | GP% | `gp_pct` | percent0 | formula |
| U | Written off | `write_offs` | int | formula |
| V | Write-off value | `writeoff_value` | money | formula |
| W | Shrinkage | `shrinkage` | percent0 | formula |
| X | Stock turns | `turns` | number1 | formula |
| Y | Share of value | `value_share` | percent0 | formula, hidden |
| Z | ABC | `abc_class` | text | formula |
| AA | Dead stock | `dead_value` | money | formula, hidden |
| AB | What needs doing | `alert` | text | formula |

#### Returns & Damaged Goods

Restocked or written off - the condition decides, and the Inventory Tracker follows.

32 columns, 200 rows of capacity, 17 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Return ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Came from | `source` | text | drop down `ReturnSourceList` |
| D | Order ref | `reference` | text | you type it |
| E | Customer | `customer` | text | you type it |
| F | SKU | `sku` | id | drop down `SKUList` |
| G | Product name | `product` | text | formula |
| H | Qty | `qty` | int | you type it (checked) |
| I | Reason | `reason` | text | drop down `ReturnReasonList` |
| J | Condition | `condition` | text | drop down `ReturnConditionList` |
| K | Restocked | `restocked` | int | formula |
| L | Back to | `location` | text | drop down `LocationList` |
| M | Unit cost | `unit_cost` | money4 | formula, hidden |
| N | Stock loss | `stock_loss` | money | formula |
| O | Refund given | `refund` | money | you type it (checked) |
| P | Replacement sent | `replacement` | flag | you type it (checked) |
| Q | Return postage | `return_shipping` | money | you type it (checked) |
| R | Total cost | `total_cost` | money | formula |
| S | Supplier ID | `supplier` | id | formula, hidden |
| T | Supplier fault | `supplier_fault` | flag | you type it (checked) |
| U | Claimed back | `claimed` | flag | you type it (checked) |
| V | Credit received | `credit_received` | money | you type it (checked) |
| W | Net loss | `net_loss` | money | formula |
| X | Status | `status` | status | drop down `ReturnStatusList` |
| Y | Notes | `notes` | wrap | you type it |
| Z | Write-off | `is_writeoff` | int | formula, hidden |
| AA | Open | `is_open` | int | formula, hidden |
| AB | Unclaimed | `unclaimed` | money | formula, hidden |
| AC | Month | `month` | text | formula, hidden |
| AD | date (number) | `date_num` | number | formula, hidden |
| AE | Row number | `seq_num` | int | formula, hidden |
| AF | What needs doing | `alert` | text | formula |

#### Warehouse Locations

Which shelf, which bin, how many. One row per SKU per location.

21 columns, 200 rows of capacity, 10 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | SKU | `sku` | id | drop down `SKUList` |
| B | Product name | `product` | text | formula |
| C | Category | `category` | text | formula, hidden |
| D | Location | `location` | text | drop down `LocationList` |
| E | Zone | `zone` | text | you type it |
| F | Bin / shelf | `bin` | text | you type it |
| G | Qty here | `qty_here` | int | you type it (checked) |
| H | Primary | `is_primary` | flag | you type it (checked) |
| I | Capacity | `capacity` | int | you type it (checked) |
| J | Full | `fill` | percent0 | formula |
| K | Storage type | `type` | text | drop down `StorageTypeList` |
| L | Last checked | `checked` | date | you type it (checked) |
| M | Notes | `notes` | wrap | you type it |
| N | All bins | `total_here` | int | formula, hidden |
| O | System says | `system_qty` | int | formula |
| P | Unplaced | `placed_gap` | int | formula |
| Q | Bins | `bins_used` | int | formula, hidden |
| R | Dup | `dup_bin` | int | formula, hidden |
| S | Primaries | `primary_count` | int | formula, hidden |
| T | checked (number) | `checked_num` | number | formula, hidden |
| U | What needs doing | `alert` | text | formula |

#### Stock Audit

Count it, compare it, then post the difference as an adjustment. A count never silently overwrites your stock.

27 columns, 300 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Audit ID | `id` | id | formula |
| B | Count date | `date` | date | you type it (checked) |
| C | Type | `audit_type` | text | drop down `AuditTypeList` |
| D | SKU | `sku` | id | drop down `SKUList` |
| E | Product name | `product` | text | formula |
| F | Location | `location` | text | drop down `LocationList` |
| G | Counted by | `counted_by` | text | drop down `StaffList` |
| H | System says | `system_qty` | int | formula |
| I | Counted | `counted_qty` | int | you type it (checked) |
| J | Variance | `variance` | int | formula |
| K | Variance % | `variance_pct` | percent0 | formula |
| L | Unit cost | `unit_cost` | money4 | formula, hidden |
| M | Value | `variance_value` | money | formula |
| N | Likely reason | `reason` | text | drop down `VarianceReasonList` |
| O | Recounted | `recount` | flag | you type it (checked) |
| P | Adjustment posted | `posted` | flag | you type it (checked) |
| Q | Movement ID | `movement_ref` | text | you type it |
| R | Status | `status` | status | drop down `AuditStatusList` |
| S | Notes | `notes` | wrap | you type it |
| T | Has variance | `has_variance` | int | formula, hidden |
| U | Unposted | `unposted` | int | formula, hidden |
| V | Abs value | `abs_value` | money | formula, hidden |
| W | Accuracy | `accuracy` | percent0 | formula, hidden |
| X | Month | `month` | text | formula, hidden |
| Y | date (number) | `date_num` | number | formula, hidden |
| Z | Row number | `seq_num` | int | formula, hidden |
| AA | What needs doing | `alert` | text | formula |

#### Product Performance

Units, revenue, profit and return on the stock held - ranked, so the decisions make themselves.

35 columns, 300 rows of capacity, 0 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | SKU | `sku` | id | formula |
| B | Product | `product` | text | formula |
| C | Category | `category` | text | formula |
| D | Status | `status` | status | formula |
| E | Retail units | `units_retail` | int | formula, hidden |
| F | Trade units | `units_wholesale` | int | formula, hidden |
| G | Units sold | `units_sold` | int | formula |
| H | Orders | `orders` | int | formula, hidden |
| I | Retail revenue | `revenue_retail` | money | formula, hidden |
| J | Trade revenue | `revenue_wholesale` | money | formula, hidden |
| K | Revenue | `revenue` | money | formula |
| L | COGS | `cogs` | money | formula |
| M | Platform fees | `fees` | money | formula, hidden |
| N | Returns cost | `returns_cost` | money | formula, hidden |
| O | Profit | `profit` | money | formula |
| P | GP% | `gp_pct` | percent0 | formula |
| Q | Profit/unit | `profit_per_unit` | money | formula |
| R | Sold/day | `daily_sales` | number1 | formula, hidden |
| S | Units/month | `monthly_units` | number1 | formula |
| T | Stock held | `stock_value` | money | formula |
| U | Return on stock | `return_on_stock` | percent0 | formula |
| V | Turns | `turns` | number1 | formula, hidden |
| W | Last sold (n) | `last_sold_num` | number | formula, hidden |
| X | Last sold | `last_sold` | date | formula |
| Y | Days since | `days_since_sale` | int | formula |
| Z | Rev rank | `revenue_rank` | int | formula, hidden |
| AA | Profit rank | `profit_rank` | int | formula, hidden |
| AB | Unit rank | `units_rank` | int | formula, hidden |
| AC | Share | `revenue_share` | percent0 | formula, hidden |
| AD | Performance | `band` | status | formula |
| AE | Dead | `is_dead` | int | formula, hidden |
| AF | Star | `is_star` | int | formula, hidden |
| AG | Dead key | `dead_key` | number | formula, hidden |
| AH | Dead rank | `dead_rank` | int | formula, hidden |
| AI | What needs doing | `alert` | text | formula |

#### Expense Tracker

Running costs. Stock purchases are counted on the Purchase Orders tab and are deliberately kept out of these totals.

24 columns, 400 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Expense ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | What it was | `description` | text | you type it |
| D | Category | `category` | text | drop down `ExpenseCategoryList` |
| E | Paid to | `supplier` | text | you type it |
| F | Amount | `amount` | money | you type it (checked) |
| G | Tax / VAT | `tax_amount` | money | you type it (checked) |
| H | Total paid | `total` | money | formula |
| I | Paid by | `method` | text | drop down `PaymentMethodList` |
| J | Recurring | `recurring` | text | drop down `FrequencyList` |
| K | Per month | `monthly` | money | formula |
| L | Tax category | `tax_category` | text | drop down `TaxCategoryList` |
| M | Receipt | `receipt` | flag | you type it (checked) |
| N | Reference | `reference` | text | you type it |
| O | Notes | `notes` | wrap | you type it |
| P | Stock purchase? | `is_stock` | int | formula, hidden |
| Q | Operating cost | `operating` | money | formula |
| R | Inventory spend | `inventory_spend` | money | formula, hidden |
| S | Possible duplicate | `duplicate` | int | formula, hidden |
| T | Large | `size_flag` | int | formula, hidden |
| U | Month | `month` | text | formula, hidden |
| V | date (number) | `date_num` | number | formula, hidden |
| W | Row number | `seq_num` | int | formula, hidden |
| X | What needs doing | `alert` | text | formula |

#### Profit Analysis

Revenue, cost of goods, gross profit, operating costs, net profit. Twelve months, calculated.

26 columns, 14 rows of capacity, 0 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Month | `month` | text | formula |
| B | Name | `month_name` | text | formula |
| C | From | `start_num` | number | formula, hidden |
| D | To | `end_num` | number | formula, hidden |
| E | Retail | `retail_revenue` | money | formula |
| F | Wholesale | `wholesale_revenue` | money | formula |
| G | Shipping in | `shipping_income` | money | formula, hidden |
| H | Total revenue | `revenue` | money | formula |
| I | COGS | `cogs` | money | formula |
| J | Platform fees | `fees` | money | formula |
| K | Shipping out | `shipping_cost` | money | formula, hidden |
| L | Gross profit | `gross_profit` | money | formula |
| M | GP% | `gross_margin` | percent0 | formula |
| N | Operating costs | `expenses` | money | formula |
| O | Returns | `returns_cost` | money | formula, hidden |
| P | Write-offs | `write_offs` | money | formula, hidden |
| Q | Net profit | `net_profit` | money | formula |
| R | Net % | `net_margin` | percent0 | formula |
| S | Stock bought | `stock_spend` | money | formula |
| T | Cash effect | `cash_effect` | money | formula |
| U | Orders | `orders` | int | formula, hidden |
| V | Avg order | `avg_order` | money | formula, hidden |
| W | vs target | `vs_target` | percent0 | formula |
| X | vs last month | `vs_prev` | percent0 | formula |
| Y | Month | `band` | status | formula |
| Z | What needs doing | `alert` | text | formula |

---

*12,759 lines of Python in 41 files.*
