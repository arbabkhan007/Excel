# Handmade Product Pricing & Profit Calculator — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 40 tab `.xlsx` workbook with over 97,000 formulas: a product catalog, a material library costed to four decimal places, a bill of materials, labour by the minute, packaging profiles, a postage rate card, overhead allocation, an editable marketplace fee table, a retail price calculator, a wholesale calculator, profit and break-even, six craft specific calculators (crochet, jewelry, candles, clothing, woodworking and digital downloads), bundles, variations, batch production, a sales ledger with estimated against actual fees, promotions, price history, competitor research, a waste log, development cost recovery, craft fairs, stockists, wholesale orders, pricing scenarios, reports, three printables and a help tab. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

The pricing is the point: cost is materials plus labour plus packaging plus postage plus overhead, and the price is `(cost + fixed fees) / (1 - target margin - fee rate)` - a division, not an addition, so the marketplace percentage comes out of the price rather than out of your profit.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_handmade_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --shop TEXT           shop name on the cover and printables
  --owner TEXT          the maker's name
  --currency SYMBOL     currency symbol  (default: $)
  --currency-code CODE  three letter code (default: USD)
  --year YYYY           the year the reports open on
  --rate AMOUNT         what you pay yourself an hour
  --margin PERCENT      target retail margin
  --wholesale PERCENT   the discount a stockist gets off retail
  --overhead AMOUNT     overhead a month, if you would rather
                        type it than list it
  --hours N             productive hours a month
  --rounding RULE       how a calculated price is rounded
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
  * [`build_handmade_workbook.py`](#build_handmade_workbookpy)
* [2. The package](#2-the-package)
  * [`handmade_pricing_spreadsheet/__init__.py`](#handmade_pricing_spreadsheet__init__py)
  * [`handmade_pricing_spreadsheet/__main__.py`](#handmade_pricing_spreadsheet__main__py)
  * [`handmade_pricing_spreadsheet/cli.py`](#handmade_pricing_spreadsheetclipy)
  * [`handmade_pricing_spreadsheet/config.py`](#handmade_pricing_spreadsheetconfigpy)
  * [`handmade_pricing_spreadsheet/layout.py`](#handmade_pricing_spreadsheetlayoutpy)
  * [`handmade_pricing_spreadsheet/formats.py`](#handmade_pricing_spreadsheetformatspy)
  * [`handmade_pricing_spreadsheet/writer.py`](#handmade_pricing_spreadsheetwriterpy)
  * [`handmade_pricing_spreadsheet/names.py`](#handmade_pricing_spreadsheetnamespy)
  * [`handmade_pricing_spreadsheet/build.py`](#handmade_pricing_spreadsheetbuildpy)
  * [`handmade_pricing_spreadsheet/sample_data.py`](#handmade_pricing_spreadsheetsample_datapy)
* [3. The costing registers](#3-the-costing-registers)
  * [`handmade_pricing_spreadsheet/sheets/__init__.py`](#handmade_pricing_spreadsheetsheets__init__py)
  * [`handmade_pricing_spreadsheet/sheets/common.py`](#handmade_pricing_spreadsheetsheetscommonpy)
  * [`handmade_pricing_spreadsheet/sheets/lists.py`](#handmade_pricing_spreadsheetsheetslistspy)
  * [`handmade_pricing_spreadsheet/sheets/products.py`](#handmade_pricing_spreadsheetsheetsproductspy)
  * [`handmade_pricing_spreadsheet/sheets/materials.py`](#handmade_pricing_spreadsheetsheetsmaterialspy)
  * [`handmade_pricing_spreadsheet/sheets/material_usage.py`](#handmade_pricing_spreadsheetsheetsmaterial_usagepy)
  * [`handmade_pricing_spreadsheet/sheets/labor.py`](#handmade_pricing_spreadsheetsheetslaborpy)
  * [`handmade_pricing_spreadsheet/sheets/packaging.py`](#handmade_pricing_spreadsheetsheetspackagingpy)
  * [`handmade_pricing_spreadsheet/sheets/shipping.py`](#handmade_pricing_spreadsheetsheetsshippingpy)
  * [`handmade_pricing_spreadsheet/sheets/overhead.py`](#handmade_pricing_spreadsheetsheetsoverheadpy)
  * [`handmade_pricing_spreadsheet/sheets/channels.py`](#handmade_pricing_spreadsheetsheetschannelspy)
* [4. The six craft calculators](#4-the-six-craft-calculators)
  * [`handmade_pricing_spreadsheet/sheets/crochet.py`](#handmade_pricing_spreadsheetsheetscrochetpy)
  * [`handmade_pricing_spreadsheet/sheets/jewelry.py`](#handmade_pricing_spreadsheetsheetsjewelrypy)
  * [`handmade_pricing_spreadsheet/sheets/candles.py`](#handmade_pricing_spreadsheetsheetscandlespy)
  * [`handmade_pricing_spreadsheet/sheets/clothing.py`](#handmade_pricing_spreadsheetsheetsclothingpy)
  * [`handmade_pricing_spreadsheet/sheets/woodworking.py`](#handmade_pricing_spreadsheetsheetswoodworkingpy)
  * [`handmade_pricing_spreadsheet/sheets/digital.py`](#handmade_pricing_spreadsheetsheetsdigitalpy)
* [5. The selling registers](#5-the-selling-registers)
  * [`handmade_pricing_spreadsheet/sheets/bundles.py`](#handmade_pricing_spreadsheetsheetsbundlespy)
  * [`handmade_pricing_spreadsheet/sheets/bundle_lines.py`](#handmade_pricing_spreadsheetsheetsbundle_linespy)
  * [`handmade_pricing_spreadsheet/sheets/variations.py`](#handmade_pricing_spreadsheetsheetsvariationspy)
  * [`handmade_pricing_spreadsheet/sheets/batches.py`](#handmade_pricing_spreadsheetsheetsbatchespy)
  * [`handmade_pricing_spreadsheet/sheets/sales.py`](#handmade_pricing_spreadsheetsheetssalespy)
  * [`handmade_pricing_spreadsheet/sheets/promotions.py`](#handmade_pricing_spreadsheetsheetspromotionspy)
  * [`handmade_pricing_spreadsheet/sheets/price_history.py`](#handmade_pricing_spreadsheetsheetsprice_historypy)
  * [`handmade_pricing_spreadsheet/sheets/competitors.py`](#handmade_pricing_spreadsheetsheetscompetitorspy)
  * [`handmade_pricing_spreadsheet/sheets/waste.py`](#handmade_pricing_spreadsheetsheetswastepy)
  * [`handmade_pricing_spreadsheet/sheets/development.py`](#handmade_pricing_spreadsheetsheetsdevelopmentpy)
  * [`handmade_pricing_spreadsheet/sheets/events.py`](#handmade_pricing_spreadsheetsheetseventspy)
  * [`handmade_pricing_spreadsheet/sheets/wholesale_customers.py`](#handmade_pricing_spreadsheetsheetswholesale_customerspy)
  * [`handmade_pricing_spreadsheet/sheets/wholesale_orders.py`](#handmade_pricing_spreadsheetsheetswholesale_orderspy)
* [6. The pages](#6-the-pages)
  * [`handmade_pricing_spreadsheet/sheets/start_here.py`](#handmade_pricing_spreadsheetsheetsstart_herepy)
  * [`handmade_pricing_spreadsheet/sheets/settings.py`](#handmade_pricing_spreadsheetsheetssettingspy)
  * [`handmade_pricing_spreadsheet/sheets/dashboard.py`](#handmade_pricing_spreadsheetsheetsdashboardpy)
  * [`handmade_pricing_spreadsheet/sheets/retail_calc.py`](#handmade_pricing_spreadsheetsheetsretail_calcpy)
  * [`handmade_pricing_spreadsheet/sheets/wholesale_calc.py`](#handmade_pricing_spreadsheetsheetswholesale_calcpy)
  * [`handmade_pricing_spreadsheet/sheets/break_even.py`](#handmade_pricing_spreadsheetsheetsbreak_evenpy)
  * [`handmade_pricing_spreadsheet/sheets/scenarios.py`](#handmade_pricing_spreadsheetsheetsscenariospy)
  * [`handmade_pricing_spreadsheet/sheets/reports.py`](#handmade_pricing_spreadsheetsheetsreportspy)
  * [`handmade_pricing_spreadsheet/sheets/printables.py`](#handmade_pricing_spreadsheetsheetsprintablespy)
  * [`handmade_pricing_spreadsheet/sheets/help_tab.py`](#handmade_pricing_spreadsheetsheetshelp_tabpy)
* [7. Quality checks (optional)](#7-quality-checks-optional)
  * [`tools/validate_handmade_workbook.py`](#toolsvalidate_handmade_workbookpy)
  * [`tools/check_handmade_sample.py`](#toolscheck_handmade_samplepy)
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

`build_handmade_workbook.py` parses the command line into a `Config`, then
`handmade_pricing_spreadsheet.build.build_workbook(path, cfg)` does six things in order:

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
6. **Activate the first tab** and close the file - the buyer should land on Start Here, not on the Lists tab.

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

40 tabs in the order they appear in the finished workbook. "Rows" is the capacity of a register at the default `--rows` setting.

| # | Tab | Kind | Written by | Columns | Rows |
|---|-----|------|------------|--------:|-----:|
| 1 | **Start Here** | page | `handmade_pricing_spreadsheet/sheets/start_here.py` | - | - |
| 2 | **Settings** | page | `handmade_pricing_spreadsheet/sheets/settings.py` | - | - |
| 3 | **Dashboard** | page | `handmade_pricing_spreadsheet/sheets/dashboard.py` | - | - |
| 4 | **Products** | register | `handmade_pricing_spreadsheet/sheets/products.py` | 71 | 300 |
| 5 | **Materials** | register | `handmade_pricing_spreadsheet/sheets/materials.py` | 33 | 250 |
| 6 | **Material Usage** | register | `handmade_pricing_spreadsheet/sheets/material_usage.py` | 18 | 900 |
| 7 | **Labor** | register | `handmade_pricing_spreadsheet/sheets/labor.py` | 22 | 300 |
| 8 | **Packaging** | register | `handmade_pricing_spreadsheet/sheets/packaging.py` | 26 | 80 |
| 9 | **Shipping** | register | `handmade_pricing_spreadsheet/sheets/shipping.py` | 18 | 60 |
| 10 | **Overhead** | register | `handmade_pricing_spreadsheet/sheets/overhead.py` | 18 | 40 |
| 11 | **Channels** | register | `handmade_pricing_spreadsheet/sheets/channels.py` | 29 | 14 |
| 12 | **Retail Calculator** | page | `handmade_pricing_spreadsheet/sheets/retail_calc.py` | - | - |
| 13 | **Wholesale Calculator** | page | `handmade_pricing_spreadsheet/sheets/wholesale_calc.py` | - | - |
| 14 | **Profit & Break-Even** | page | `handmade_pricing_spreadsheet/sheets/break_even.py` | - | - |
| 15 | **Crochet** | register | `handmade_pricing_spreadsheet/sheets/crochet.py` | 47 | 80 |
| 16 | **Jewelry** | register | `handmade_pricing_spreadsheet/sheets/jewelry.py` | 49 | 80 |
| 17 | **Candles** | register | `handmade_pricing_spreadsheet/sheets/candles.py` | 50 | 80 |
| 18 | **Clothing** | register | `handmade_pricing_spreadsheet/sheets/clothing.py` | 45 | 80 |
| 19 | **Woodworking** | register | `handmade_pricing_spreadsheet/sheets/woodworking.py` | 47 | 80 |
| 20 | **Digital** | register | `handmade_pricing_spreadsheet/sheets/digital.py` | 42 | 80 |
| 21 | **Bundles** | register | `handmade_pricing_spreadsheet/sheets/bundles.py` | 34 | 60 |
| 22 | **Bundle Lines** | register | `handmade_pricing_spreadsheet/sheets/bundle_lines.py` | 14 | 200 |
| 23 | **Variations** | register | `handmade_pricing_spreadsheet/sheets/variations.py` | 38 | 300 |
| 24 | **Batches** | register | `handmade_pricing_spreadsheet/sheets/batches.py` | 31 | 200 |
| 25 | **Sales** | register | `handmade_pricing_spreadsheet/sheets/sales.py` | 39 | 800 |
| 26 | **Promotions** | register | `handmade_pricing_spreadsheet/sheets/promotions.py` | 24 | 40 |
| 27 | **Price History** | register | `handmade_pricing_spreadsheet/sheets/price_history.py` | 23 | 300 |
| 28 | **Competitors** | register | `handmade_pricing_spreadsheet/sheets/competitors.py` | 22 | 150 |
| 29 | **Waste Log** | register | `handmade_pricing_spreadsheet/sheets/waste.py` | 14 | 200 |
| 30 | **Development** | register | `handmade_pricing_spreadsheet/sheets/development.py` | 22 | 80 |
| 31 | **Events** | register | `handmade_pricing_spreadsheet/sheets/events.py` | 28 | 40 |
| 32 | **Wholesale Customers** | register | `handmade_pricing_spreadsheet/sheets/wholesale_customers.py` | 27 | 60 |
| 33 | **Wholesale Orders** | register | `handmade_pricing_spreadsheet/sheets/wholesale_orders.py` | 30 | 200 |
| 34 | **Scenarios** | page | `handmade_pricing_spreadsheet/sheets/scenarios.py` | - | - |
| 35 | **Reports** | page | `handmade_pricing_spreadsheet/sheets/reports.py` | - | - |
| 36 | **Print Costing** | page | `handmade_pricing_spreadsheet/sheets/printables.py` | - | - |
| 37 | **Print Quote** | page | `handmade_pricing_spreadsheet/sheets/printables.py` | - | - |
| 38 | **Print Order** | page | `handmade_pricing_spreadsheet/sheets/printables.py` | - | - |
| 39 | **Lists** | page | `handmade_pricing_spreadsheet/sheets/lists.py` | - | - |
| 40 | **Help** | page | `handmade_pricing_spreadsheet/sheets/help_tab.py` | - | - |

### Every file at a glance

| File | Lines | What it is |
|------|------:|------------|
| [`build_handmade_workbook.py`](#build_handmade_workbookpy) | 17 | Build the Handmade Product Pricing & Profit Calculator workbook |
| [`handmade_pricing_spreadsheet/__init__.py`](#handmade_pricing_spreadsheet__init__py) | 19 | Handmade Product Pricing & Profit Calculator |
| [`handmade_pricing_spreadsheet/__main__.py`](#handmade_pricing_spreadsheet__main__py) | 6 | ``python -m handmade_pricing_spreadsheet`` |
| [`handmade_pricing_spreadsheet/cli.py`](#handmade_pricing_spreadsheetclipy) | 170 | Command line front end |
| [`handmade_pricing_spreadsheet/config.py`](#handmade_pricing_spreadsheetconfigpy) | 337 | Configuration for the Handmade Product Pricing & Profit Calculator |
| [`handmade_pricing_spreadsheet/layout.py`](#handmade_pricing_spreadsheetlayoutpy) | 253 | Layout engine |
| [`handmade_pricing_spreadsheet/formats.py`](#handmade_pricing_spreadsheetformatspy) | 202 | Cell format factory |
| [`handmade_pricing_spreadsheet/writer.py`](#handmade_pricing_spreadsheetwriterpy) | 458 | Sheet writing helpers shared by every tab of the workbook |
| [`handmade_pricing_spreadsheet/names.py`](#handmade_pricing_spreadsheetnamespy) | 54 | Workbook level defined names that point at the register sheets |
| [`handmade_pricing_spreadsheet/build.py`](#handmade_pricing_spreadsheetbuildpy) | 163 | Workbook builder: puts every tab together in the right order |
| [`handmade_pricing_spreadsheet/sample_data.py`](#handmade_pricing_spreadsheetsample_datapy) | 1,384 | Demo data: a season in the life of Willow & Thread Handmade |
| [`handmade_pricing_spreadsheet/sheets/__init__.py`](#handmade_pricing_spreadsheetsheets__init__py) | 0 |  |
| [`handmade_pricing_spreadsheet/sheets/common.py`](#handmade_pricing_spreadsheetsheetscommonpy) | 400 | Small building blocks shared by the register sheets |
| [`handmade_pricing_spreadsheet/sheets/lists.py`](#handmade_pricing_spreadsheetsheetslistspy) | 476 | `Lists` - every drop down in the workbook, plus the factor tables |
| [`handmade_pricing_spreadsheet/sheets/products.py`](#handmade_pricing_spreadsheetsheetsproductspy) | 512 | `Products` - the catalog, and the price every other tab argues about |
| [`handmade_pricing_spreadsheet/sheets/materials.py`](#handmade_pricing_spreadsheetsheetsmaterialspy) | 239 | `Materials` - every supply you buy, costed down to the gram |
| [`handmade_pricing_spreadsheet/sheets/material_usage.py`](#handmade_pricing_spreadsheetsheetsmaterial_usagepy) | 163 | `Material Usage` - the bill of materials: what goes into what |
| [`handmade_pricing_spreadsheet/sheets/labor.py`](#handmade_pricing_spreadsheetsheetslaborpy) | 196 | `Labor` - your hours, stage by stage, priced like the work they are |
| [`handmade_pricing_spreadsheet/sheets/packaging.py`](#handmade_pricing_spreadsheetsheetspackagingpy) | 220 | `Packaging` - the box, the tissue, the card, the sticker |
| [`handmade_pricing_spreadsheet/sheets/shipping.py`](#handmade_pricing_spreadsheetsheetsshippingpy) | 160 | `Shipping` - your postage rate card, and who is really paying it |
| [`handmade_pricing_spreadsheet/sheets/overhead.py`](#handmade_pricing_spreadsheetsheetsoverheadpy) | 149 | `Overhead` - the bills that exist whether you sell anything or not |
| [`handmade_pricing_spreadsheet/sheets/channels.py`](#handmade_pricing_spreadsheetsheetschannelspy) | 222 | `Channels` - what each place you sell costs you to sell there |
| [`handmade_pricing_spreadsheet/sheets/crochet.py`](#handmade_pricing_spreadsheetsheetscrochetpy) | 209 | `Crochet` - yarn by the gram, because that is how it is used |
| [`handmade_pricing_spreadsheet/sheets/jewelry.py`](#handmade_pricing_spreadsheetsheetsjewelrypy) | 196 | `Jewelry` - metal by the gram, stones by the piece, scrap included |
| [`handmade_pricing_spreadsheet/sheets/candles.py`](#handmade_pricing_spreadsheetsheetscandlespy) | 205 | `Candles` - wax by the gram, worked out from the vessel |
| [`handmade_pricing_spreadsheet/sheets/clothing.py`](#handmade_pricing_spreadsheetsheetsclothingpy) | 182 | `Clothing` - fabric with the layout waste in, notions, and the sewing |
| [`handmade_pricing_spreadsheet/sheets/woodworking.py`](#handmade_pricing_spreadsheetsheetswoodworkingpy) | 192 | `Woodworking` - timber with the offcuts in, hardware, finish, machine time |
| [`handmade_pricing_spreadsheet/sheets/digital.py`](#handmade_pricing_spreadsheetsheetsdigitalpy) | 193 | `Digital` - patterns and downloads: all the cost is up front |
| [`handmade_pricing_spreadsheet/sheets/bundles.py`](#handmade_pricing_spreadsheetsheetsbundlespy) | 239 | `Bundles` - gift sets and multi-buys, priced so the discount is a choice |
| [`handmade_pricing_spreadsheet/sheets/bundle_lines.py`](#handmade_pricing_spreadsheetsheetsbundle_linespy) | 96 | `Bundle Lines` - what is in each bundle |
| [`handmade_pricing_spreadsheet/sheets/variations.py`](#handmade_pricing_spreadsheetsheetsvariationspy) | 243 | `Variations` - sizes, colours, personalisation and rush orders |
| [`handmade_pricing_spreadsheet/sheets/batches.py`](#handmade_pricing_spreadsheetsheetsbatchespy) | 199 | `Batches` - what a production run actually cost, against what you estimated |
| [`handmade_pricing_spreadsheet/sheets/sales.py`](#handmade_pricing_spreadsheetsheetssalespy) | 261 | `Sales` - what really happened, order by order |
| [`handmade_pricing_spreadsheet/sheets/promotions.py`](#handmade_pricing_spreadsheetsheetspromotionspy) | 175 | `Promotions` - what a sale costs, and how much more you have to sell |
| [`handmade_pricing_spreadsheet/sheets/price_history.py`](#handmade_pricing_spreadsheetsheetsprice_historypy) | 156 | `Price History` - every price change, and whether it worked |
| [`handmade_pricing_spreadsheet/sheets/competitors.py`](#handmade_pricing_spreadsheetsheetscompetitorspy) | 144 | `Competitors` - what the market charges, and where you sit in it |
| [`handmade_pricing_spreadsheet/sheets/waste.py`](#handmade_pricing_spreadsheetsheetswastepy) | 121 | `Waste Log` - the material that never made it into a product |
| [`handmade_pricing_spreadsheet/sheets/development.py`](#handmade_pricing_spreadsheetsheetsdevelopmentpy) | 150 | `Development` - what a new design cost before it earned anything |
| [`handmade_pricing_spreadsheet/sheets/events.py`](#handmade_pricing_spreadsheetsheetseventspy) | 176 | `Events` - craft fairs, markets and whether the stall paid |
| [`handmade_pricing_spreadsheet/sheets/wholesale_customers.py`](#handmade_pricing_spreadsheetsheetswholesale_customerspy) | 160 | `Wholesale Customers` - the shops that buy from you, and what they are worth |
| [`handmade_pricing_spreadsheet/sheets/wholesale_orders.py`](#handmade_pricing_spreadsheetsheetswholesale_orderspy) | 174 | `Wholesale Orders` - line by line, priced at the shop's own discount |
| [`handmade_pricing_spreadsheet/sheets/start_here.py`](#handmade_pricing_spreadsheetsheetsstart_herepy) | 336 | `Start Here` - the first thing anyone sees when they open the file |
| [`handmade_pricing_spreadsheet/sheets/settings.py`](#handmade_pricing_spreadsheetsheetssettingspy) | 401 | The `Settings` tab - shop profile, labour rates, fees, margins, targets |
| [`handmade_pricing_spreadsheet/sheets/dashboard.py`](#handmade_pricing_spreadsheetsheetsdashboardpy) | 420 | `Dashboard` - one page that answers "am I actually making money?" Twelve numbers across the top, a list of... |
| [`handmade_pricing_spreadsheet/sheets/retail_calc.py`](#handmade_pricing_spreadsheetsheetsretail_calcpy) | 439 | `Retail Calculator` - the page the whole workbook exists for |
| [`handmade_pricing_spreadsheet/sheets/wholesale_calc.py`](#handmade_pricing_spreadsheetsheetswholesale_calcpy) | 316 | `Wholesale Calculator` - can you afford to sell at half price? Wholesale is simple arithmetic that catches... |
| [`handmade_pricing_spreadsheet/sheets/break_even.py`](#handmade_pricing_spreadsheetsheetsbreak_evenpy) | 333 | `Profit & Break-Even` - how many you have to sell before you earn anything |
| [`handmade_pricing_spreadsheet/sheets/scenarios.py`](#handmade_pricing_spreadsheetsheetsscenariospy) | 284 | `Scenarios` - the same product priced five ways, on every channel, with the materials going up |
| [`handmade_pricing_spreadsheet/sheets/reports.py`](#handmade_pricing_spreadsheetsheetsreportspy) | 338 | `Reports` - the numbers you would want if someone asked how the year went |
| [`handmade_pricing_spreadsheet/sheets/printables.py`](#handmade_pricing_spreadsheetsheetsprintablespy) | 526 | The three printable forms |
| [`handmade_pricing_spreadsheet/sheets/help_tab.py`](#handmade_pricing_spreadsheetsheetshelp_tabpy) | 282 | `Help` - troubleshooting, a glossary and a map of every feature |
| [`tools/validate_handmade_workbook.py`](#toolsvalidate_handmade_workbookpy) | 702 | Static checks on a generated Handmade Pricing workbook |
| [`tools/check_handmade_sample.py`](#toolscheck_handmade_samplepy) | 400 | Cross-check the example records in a generated handmade workbook |
| [`tools/recalc_check.py`](#toolsrecalc_checkpy) | 216 | Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error... |
| [`tools/export_source_md.py`](#toolsexport_source_mdpy) | 1,250 | Write every source file of a generator into one markdown document |
| [`tools/make_download_zip.py`](#toolsmake_download_zippy) | 286 | Pack a generator into a single zip people can download and run |

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_handmade_workbook.py
```

The workbook lands in `output/Handmade Product Pricing & Profit Calculator.xlsx`.

### `build_handmade_workbook.py`

*17 lines*

> Build the Handmade Product Pricing & Profit Calculator workbook.
>
>     python3 build_handmade_workbook.py
>     python3 build_handmade_workbook.py -o "output/My Pricing.xlsx" --blank
>     python3 build_handmade_workbook.py --currency "£" --rate 22 --margin 60
>
> Run `--help` for every option. The generator itself lives in the
> `handmade_pricing_spreadsheet` package next to this file.

````python
#!/usr/bin/env python3
"""Build the Handmade Product Pricing & Profit Calculator workbook.

    python3 build_handmade_workbook.py
    python3 build_handmade_workbook.py -o "output/My Pricing.xlsx" --blank
    python3 build_handmade_workbook.py --currency "£" --rate 22 --margin 60

Run `--help` for every option. The generator itself lives in the
`handmade_pricing_spreadsheet` package next to this file.
"""

import sys

from handmade_pricing_spreadsheet.cli import main

if __name__ == "__main__":
    sys.exit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer, the cross sheet names and the demo shop. These are the parts every tab leans on.

### `handmade_pricing_spreadsheet/__init__.py`

*19 lines*

> Handmade Product Pricing & Profit Calculator.
>
> Build the workbook with::
>
>     python -m handmade_pricing_spreadsheet -o "output/Pricing.xlsx"
>
> or from Python::
>
>     from handmade_pricing_spreadsheet import build_workbook
>     build_workbook("Handmade Pricing Calculator.xlsx")

````python
"""Handmade Product Pricing & Profit Calculator.

Build the workbook with::

    python -m handmade_pricing_spreadsheet -o "output/Pricing.xlsx"

or from Python::

    from handmade_pricing_spreadsheet import build_workbook
    build_workbook("Handmade Pricing Calculator.xlsx")
"""

from .build import build_workbook
from .config import (AUTHOR, PRODUCT_LISTING_TITLE, PRODUCT_NAME,
                     PRODUCT_SUBTITLE, PRODUCT_VERSION, Config)

__all__ = ["build_workbook", "Config", "AUTHOR", "PRODUCT_NAME",
           "PRODUCT_SUBTITLE", "PRODUCT_LISTING_TITLE", "PRODUCT_VERSION"]
__version__ = PRODUCT_VERSION
````

### `handmade_pricing_spreadsheet/__main__.py`

*6 lines*

> ``python -m handmade_pricing_spreadsheet``.

````python
"""``python -m handmade_pricing_spreadsheet``."""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
````

### `handmade_pricing_spreadsheet/cli.py`

*170 lines · 3 functions*

> Command line front end.
>
> python -m handmade_pricing_spreadsheet
> python -m handmade_pricing_spreadsheet -o "output/Pricing.xlsx" --blank
> python -m handmade_pricing_spreadsheet --currency "£" --rate 22 --rows 150

| Defines | What it does |
|---------|--------------|
| `build_parser()` |  |
| `config_from_args(args)` |  |
| `main(argv)` |  |

**Module constants:** `DEFAULT_OUTPUT` (os.path.join(...)).

````python
"""Command line front end.

    python -m handmade_pricing_spreadsheet
    python -m handmade_pricing_spreadsheet -o "output/Pricing.xlsx" --blank
    python -m handmade_pricing_spreadsheet --currency "£" --rate 22 --rows 150
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import (AUTHOR, PRODUCT_LISTING_TITLE, PRODUCT_NAME,
                     ROUNDING_RULES, Config)

DEFAULT_OUTPUT = os.path.join(
    "output", "Handmade Product Pricing & Profit Calculator.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="handmade_pricing_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty "
                             "(no example products)")
    parser.add_argument("--no-protect", action="store_true",
                        help="do not protect the sheets (for development)")
    parser.add_argument("--password", default=None,
                        help="sheet protection password")
    parser.add_argument("--author", default=AUTHOR,
                        help="document author / creator")
    parser.add_argument("--shop", default=None,
                        help="pre-fill the shop name on Settings")
    parser.add_argument("--owner", default=None,
                        help="pre-fill the maker / owner name")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="the year the reports open on")
    parser.add_argument("--rate", type=float, default=None,
                        help="what you pay yourself an hour")
    parser.add_argument("--margin", type=float, default=None,
                        help="target retail margin as a percentage, "
                             "e.g. 55 for 55%%")
    parser.add_argument("--wholesale-margin", type=float, default=None,
                        help="target wholesale margin as a percentage")
    parser.add_argument("--wholesale-discount", type=float, default=None,
                        help="wholesale price as a percentage off retail, "
                             "e.g. 50 for half price")
    parser.add_argument("--rounding", choices=ROUNDING_RULES, default=None,
                        help="how calculated prices are rounded")
    parser.add_argument("--transaction-fee", type=float, default=None,
                        help="marketplace transaction fee as a "
                             "percentage, e.g. 6.5")
    parser.add_argument("--listing-fee", type=float, default=None,
                        help="fee charged to list one item")
    parser.add_argument("--tax", type=float, default=None,
                        help="sales tax / VAT rate as a percentage")
    parser.add_argument("--units", choices=["Metric", "Imperial"],
                        default=None,
                        help="the measurement system used for materials")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the default row "
                             "counts are multiplied to fit this many "
                             "product rows)")
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
    if args.rate:
        kwargs["hourly_rate"] = args.rate
    if args.margin is not None:
        kwargs["target_retail_margin"] = round(args.margin / 100.0, 4)
    if args.wholesale_margin is not None:
        kwargs["target_wholesale_margin"] = round(
            args.wholesale_margin / 100.0, 4)
    if args.wholesale_discount is not None:
        kwargs["wholesale_discount"] = round(
            args.wholesale_discount / 100.0, 4)
    if args.rounding:
        kwargs["rounding_rule"] = args.rounding
    if args.transaction_fee is not None:
        kwargs["transaction_fee"] = round(args.transaction_fee / 100.0, 5)
    if args.listing_fee is not None:
        kwargs["listing_fee"] = args.listing_fee
    if args.tax is not None:
        kwargs["tax_rate"] = round(args.tax / 100.0, 4)
    if args.units:
        kwargs["measure_system"] = args.units
        if args.units == "Imperial":
            kwargs["weight_unit"] = "oz"
            kwargs["length_unit"] = "inch"
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.shop:
        cfg.shop.name = args.shop
    if args.owner:
        cfg.shop.owner = args.owner
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
    print(f"  sheets        : {len(SHEET_ORDER)}")
    print(f"  year          : {cfg.year}")
    print(f"  shop          : {cfg.shop.name}")
    print(f"  labour        : {cfg.currency_symbol}{cfg.hourly_rate:,.2f} an "
          f"hour ({cfg.currency_symbol}{cfg.labor_cost_per_hour:,.2f} with "
          f"on-costs)")
    print(f"  target margin : {cfg.target_retail_margin:.0%} retail, "
          f"{cfg.target_wholesale_margin:.0%} wholesale")
    print(f"  example data  : "
          f"{'no' if not cfg.include_sample_data else 'yes'}")
    print(f"  protection    : "
          f"{'off' if not cfg.protect_sheets else 'on, password ' + cfg.password}")
    print(f"  written to    : {os.path.abspath(path)}  ({size:,.0f} KB)")
    return 0


if __name__ == "__main__":        # pragma: no cover
    raise SystemExit(main())
````

### `handmade_pricing_spreadsheet/config.py`

*337 lines · 3 classes*

> Configuration for the Handmade Product Pricing & Profit Calculator.
>
> Everything a maker may want to re-brand lives here, so the workbook can
> be regenerated for a different shop, currency, country or pricing model
> without touching any of the sheet code.
>
> Author of the generated workbooks: Novality Store.

| Defines | What it does |
|---------|--------------|
| `class Palette` |  |
| `class ShopDefaults` | Sample values written into the business profile on Settings |
| `class Config` | Everything the generator needs to know to write one workbook |

**Module constants:** `AUTHOR` ('Novality Store'), `PRODUCT_NAME` ('Handmade Product Pricing & Profit Calculator'), `PRODUCT_SUBTITLE` ('Material & Labor Cost Tracker, Retail and Wholesale Pric...), `PRODUCT_VERSION` ('1.0.0'), `PRODUCT_LISTING_TITLE` ('Handmade Product Pricing Calculator, Material and Labor ...), `FORMULA_PASSWORD` ('premium'), `SALES_CHANNELS` (7 items), `PRODUCT_STATUSES` (8 items), `CRAFT_SHEETS` (6 items), `PRODUCT_CATEGORIES` (15 items), `ROUNDING_RULES` (8 items), `PRICING_SCENARIOS` (5 entries), `VOLUME_STEPS` (5 items), `UNIT_CONVERSIONS` (17 items) and 1 more.

````python
"""
Configuration for the Handmade Product Pricing & Profit Calculator.

Everything a maker may want to re-brand lives here, so the workbook can
be regenerated for a different shop, currency, country or pricing model
without touching any of the sheet code.

Author of the generated workbooks: Novality Store.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import date


# ---------------------------------------------------------------------------
# Product / branding constants
# ---------------------------------------------------------------------------

AUTHOR = "Novality Store"
PRODUCT_NAME = "Handmade Product Pricing & Profit Calculator"
PRODUCT_SUBTITLE = (
    "Material & Labor Cost Tracker, Retail and Wholesale Pricing, "
    "Profit Margin & Break-Even Calculator"
)
PRODUCT_VERSION = "1.0.0"

#: Title suggested for the shop listing.
PRODUCT_LISTING_TITLE = (
    "Handmade Product Pricing Calculator, Material and Labor Cost "
    "Tracker, Profit Margin and Wholesale Price Spreadsheet"
)

#: Password that protects every sheet so formula cells cannot be edited
#: by accident.  Input cells stay unlocked, formula cells are locked.
FORMULA_PASSWORD = "premium"

#: Where a product is sold.  Each one carries its own fee structure on
#: the Channels tab.
SALES_CHANNELS = ["Etsy", "Own website", "Craft fair", "Wholesale",
                  "Instagram / direct", "Amazon Handmade", "Other"]

#: Product status, in the order a product moves through them.
PRODUCT_STATUSES = ["Idea", "Prototype", "Draft", "Active", "Made to order",
                    "Seasonal", "Out of stock", "Discontinued"]

#: The six specialised calculators, in tab order.  The product catalog
#: pulls its costs from whichever of these a product belongs to.
CRAFT_SHEETS = ["Crochet", "Jewelry", "Candles", "Clothing",
                "Woodworking", "Digital"]

#: Product categories offered in the drop downs.
PRODUCT_CATEGORIES = [
    "Crochet & knitting", "Jewelry", "Candles & home fragrance",
    "Clothing & accessories", "Woodworking", "Digital download",
    "Ceramics & pottery", "Bath & body", "Paper & stationery",
    "Art & prints", "Bags & leather", "Toys & amigurumi",
    "Pet products", "Wedding & occasion", "Other handmade",
]

#: How a price is rounded once it has been calculated.
ROUNDING_RULES = ["None", "Nearest 0.99", "Nearest 0.95", "Nearest 0.50",
                  "Nearest 1", "Nearest 5", "Up to whole 1",
                  "Up to whole 5"]

#: Pricing scenarios compared on the Scenarios tab: a name and the
#: target margin that goes with it.
PRICING_SCENARIOS = {
    "Budget market": 0.15,
    "Conservative": 0.22,
    "Standard": 0.30,
    "Premium": 0.40,
    "Luxury": 0.50,
}

#: Sales volumes the profit analyser forecasts against.
VOLUME_STEPS = [5, 10, 25, 50, 100]

#: Unit conversions written into the Lists tab: unit -> (base unit,
#: how many base units one of it is).
UNIT_CONVERSIONS = [
    ("g", "g", 1.0), ("kg", "g", 1000.0), ("oz", "g", 28.3495),
    ("lb", "g", 453.592), ("ml", "ml", 1.0), ("l", "ml", 1000.0),
    ("fl oz", "ml", 29.5735), ("cm", "cm", 1.0), ("m", "cm", 100.0),
    ("inch", "cm", 2.54), ("ft", "cm", 30.48), ("yard", "cm", 91.44),
    ("each", "each", 1.0), ("pair", "each", 2.0), ("sheet", "each", 1.0),
    ("skein", "each", 1.0), ("metre", "cm", 100.0),
]


# ---------------------------------------------------------------------------
# Colours - warm craft studio palette: clay, terracotta, sage, cream
# ---------------------------------------------------------------------------

class Palette:
    ink = "#3B2A26"           # darkest clay - titles
    primary = "#8C4A32"       # terracotta headers
    primary_mid = "#B5704F"   # section bands
    primary_soft = "#F6E4D7"  # light band / alternating rows
    accent = "#5E7B5A"        # sage accent
    accent_soft = "#E4EDE1"
    input_bg = "#FFFFFF"      # editable cell background
    input_border = "#E3D2C6"
    calc_bg = "#FBF3EC"       # locked formula cell background (warm cream)
    calc_border = "#EADCCF"
    white = "#FFFFFF"
    grey_text = "#6E5C55"
    line = "#DCC9BB"
    good_bg = "#DCEBD8"
    good_text = "#2F5A2B"
    warn_bg = "#FDEDD2"
    warn_text = "#8A5A00"
    bad_bg = "#F8DAD9"
    bad_text = "#9C2A24"
    info_bg = "#E2E6F3"
    info_text = "#3A4578"
    note_bg = "#FBF7F3"


# ---------------------------------------------------------------------------
# Row capacity of every register
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Settings": 0,
    "Lists": 60,
    "Products": 300,
    "Materials": 250,
    "Material Usage": 900,
    "Labor": 300,
    "Packaging": 80,
    "Shipping": 60,
    "Overhead": 40,
    "Channels": 14,
    "Crochet": 80,
    "Jewelry": 80,
    "Candles": 80,
    "Clothing": 80,
    "Woodworking": 80,
    "Digital": 80,
    "Bundles": 60,
    "Bundle Lines": 200,
    "Variations": 300,
    "Batches": 200,
    "Sales": 800,
    "Promotions": 40,
    "Price History": 300,
    "Competitors": 150,
    "Waste Log": 200,
    "Development": 80,
    "Events": 40,
    "Wholesale Customers": 60,
    "Wholesale Orders": 200,
}


@dataclass
class ShopDefaults:
    """Sample values written into the business profile on Settings."""

    name: str = "Willow & Thread Handmade"
    owner: str = "Maya Albright"
    brand: str = "Willow & Thread"
    email: str = "hello@willowandthread.co"
    phone: str = "(828) 555 0176"
    address: str = "42 Larkspur Lane, Asheville, NC 28801"
    website: str = "willowandthread.co"
    etsy_shop: str = "WillowAndThreadCo"
    instagram: str = "@willowandthread"
    tagline: str = "Small batch crochet, candles and keepsakes since 2021"
    tax_id: str = "NC-HM-2021-4471"


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
    measure_system: str = "Metric"       # Metric | Imperial
    weight_unit: str = "g"               # g | oz
    length_unit: str = "cm"              # cm | inch

    # Shop profile pre-filled on the Settings tab ---------------------------
    shop: ShopDefaults = field(default_factory=ShopDefaults)
    plan_year: int = 0                   # 0 = the year of `today`

    # Labour ----------------------------------------------------------------
    hourly_rate: float = 25.00           # what the maker pays themselves
    minimum_wage: float = 15.00          # the floor the workbook warns at
    employee_rate: float = 18.00
    contractor_rate: float = 32.00
    payroll_burden: float = 0.12         # tax, insurance, holiday pay
    default_batch_size: int = 10

    # Pricing ---------------------------------------------------------------
    target_retail_margin: float = 0.30
    target_wholesale_margin: float = 0.20
    default_markup: float = 1.20         # 120% on cost when markup is used
    wholesale_discount: float = 0.40     # trade price = 60% of retail
    margin_warning: float = 0.15         # below this the row is flagged
    rounding_rule: str = "Nearest 0.99"
    default_discount: float = 0.10
    minimum_order_qty: int = 6
    minimum_order_value: float = 150.00

    # Costs -----------------------------------------------------------------
    default_packaging: float = 1.85
    default_shipping: float = 0.00       # what the seller absorbs
    default_waste: float = 0.05          # 5% of every material is lost
    default_defect_rate: float = 0.03    # 3% of a batch is unsellable
    supplies_per_hour: float = 0.60      # consumables while working

    # Marketplace fees (editable - these move, check the fee schedule) ------
    listing_fee: float = 0.20
    transaction_fee: float = 0.065       # 6.5% of the order
    processing_percent: float = 0.03     # payment processing
    processing_fixed: float = 0.25
    offsite_ads_fee: float = 0.15        # only on orders that come from ads
    offsite_ads_share: float = 0.20      # how many orders those are
    regulatory_fee: float = 0.0040
    currency_fee: float = 0.025
    fees_include_shipping: str = "Yes"
    fees_include_tax: str = "No"

    # Overhead --------------------------------------------------------------
    monthly_overhead: float = 420.00     # typed on the Overhead tab
    productive_hours: float = 70.0       # hours a month actually making
    units_per_month: int = 24
    tax_rate: float = 0.00               # sales tax, if you add it on

    # Targets the dashboard measures against --------------------------------
    monthly_revenue_target: float = 3500.0
    annual_revenue_target: float = 42000.0
    monthly_profit_target: float = 1400.0
    # the example shop is smaller than the defaults assume, so the demo
    # file carries targets it can actually be measured against
    sample_monthly_revenue_target: float = 2400.0
    sample_annual_revenue_target: float = 28800.0
    sample_monthly_profit_target: float = 900.0

    # Reminder windows ------------------------------------------------------
    price_review_days: int = 180         # how often a price gets a look
    fee_review_days: int = 180           # how often the fee schedule does
    low_stock_days: int = 21

    # Behaviour -------------------------------------------------------------
    top_rows: int = 5                    # size of the "top / best" tables
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
    def metric(self) -> bool:
        return self.measure_system.lower().startswith("met")

    @property
    def labor_cost_per_hour(self) -> float:
        """An hour of the maker's own time, once on-costs are added."""
        return round(self.hourly_rate * (1 + self.payroll_burden), 2)

    @property
    def overhead_per_hour(self) -> float:
        if self.productive_hours <= 0:
            return 0.0
        return round(self.monthly_overhead / self.productive_hours, 2)

    def scale_rows(self, product_rows: int) -> None:
        """Scale every register in place so Products holds N rows."""
        base = DEFAULT_ROWS["Products"]
        factor = max(0.02, float(product_rows) / float(base))
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
        """Unit costs can be fractions of a cent - a gram of yarn is."""
        s = self.currency_symbol
        return f'{s}#,##0.0000;[Red]-{s}#,##0.0000'
````

### `handmade_pricing_spreadsheet/layout.py`

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

### `handmade_pricing_spreadsheet/formats.py`

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

### `handmade_pricing_spreadsheet/writer.py`

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
    ("Products", "Products"),
    ("Materials", "Materials"),
    ("Labour", "Labor"),
    ("Retail Price", "Retail Calculator"),
    ("Wholesale", "Wholesale Calculator"),
    ("Sales", "Sales"),
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

### `handmade_pricing_spreadsheet/names.py`

*54 lines · 1 function*

> Workbook level defined names that point at the register sheets.
>
> The drop down lists on the `Lists` tab define their own names. These are
> the ones that have to follow live data - product IDs, material names, the
> shops you wholesale to - so they are built from the registry once every
> register exists. Each one is an ``OFFSET`` range that grows and shrinks
> with the number of filled rows, which is why a new material appears in
> the bill of materials drop down the moment you type it.

| Defines | What it does |
|---------|--------------|
| `define_cross_sheet_names(book, reg)` | Create the live lookup names |

**Module constants:** `CROSS_SHEET_NAMES` (21 entries).

````python
"""Workbook level defined names that point at the register sheets.

The drop down lists on the `Lists` tab define their own names. These are
the ones that have to follow live data - product IDs, material names, the
shops you wholesale to - so they are built from the registry once every
register exists. Each one is an ``OFFSET`` range that grows and shrinks
with the number of filled rows, which is why a new material appears in
the bill of materials drop down the moment you type it.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # the spine of the whole workbook
    "ProductIDs": ("Products", "id"),
    "ProductNames": ("Products", "name"),
    "ProductSKUs": ("Products", "sku"),
    "MaterialIDs": ("Materials", "id"),
    "MaterialNames": ("Materials", "name"),
    "ChannelNames": ("Channels", "name"),
    "ShipOptions": ("Shipping", "label"),
    "PackComponents": ("Packaging", "name"),
    "BundleIDs": ("Bundles", "id"),
    "BatchIDs": ("Batches", "id"),
    "PromoCodes": ("Promotions", "code"),
    "EventNames": ("Events", "name"),
    "CustomerIDs": ("Wholesale Customers", "id"),
    "CustomerNames": ("Wholesale Customers", "name"),
    "OrderIDs": ("Wholesale Orders", "id"),
    # columns the dashboard and the reports scan
    "ProductCategories": ("Products", "category"),
    "ProductMargins": ("Products", "margin"),
    "SaleDates": ("Sales", "date"),
    "SaleMonths": ("Sales", "month"),
    "SaleProducts": ("Sales", "product"),
    "SaleChannels": ("Sales", "channel"),
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

### `handmade_pricing_spreadsheet/build.py`

*163 lines · 3 functions*

> Workbook builder: puts every tab together in the right order.
>
> Typical use::
>
>     from handmade_pricing_spreadsheet import build_workbook
>     build_workbook("Handmade Product Pricing & Profit Calculator.xlsx")

| Defines | What it does |
|---------|--------------|
| `build_register(book, fmts, cfg, reg, module)` | Create one register sheet from its :class:`SheetSpec` |
| `build_workbook(path, cfg)` | Write the whole product to *path* and return the path |
| `_set_properties(book, cfg)` |  |

**Module constants:** `REGISTER_MODULES` (27 items), `SHEET_ORDER` (40 items), `PAGE_BUILDERS` (13 entries), `REGISTER_BY_NAME` ({module.NAME: module for module in REGISTER_MODULES}).

````python
"""Workbook builder: puts every tab together in the right order.

Typical use::

    from handmade_pricing_spreadsheet import build_workbook
    build_workbook("Handmade Product Pricing & Profit Calculator.xlsx")
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
from .sheets import (batches, break_even, bundle_lines, bundles, candles,
                     channels, clothing, competitors, crochet, dashboard,
                     development, digital, events, help_tab, jewelry,
                     labor, lists, material_usage, materials, overhead,
                     packaging, price_history, printables, products,
                     promotions, reports, retail_calc, sales, scenarios,
                     settings, shipping, start_here, variations,
                     waste, wholesale_calc, wholesale_customers,
                     wholesale_orders, woodworking)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    products, materials, material_usage, labor, packaging, shipping,
    overhead, channels, crochet, jewelry, candles, clothing, woodworking,
    digital, bundles, bundle_lines, variations, batches, sales,
    promotions, price_history, competitors, waste, development, events,
    wholesale_customers, wholesale_orders,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Start Here", "Settings", "Dashboard",
    "Products", "Materials", "Material Usage", "Labor",
    "Packaging", "Shipping", "Overhead", "Channels",
    "Retail Calculator", "Wholesale Calculator", "Profit & Break-Even",
    "Crochet", "Jewelry", "Candles", "Clothing", "Woodworking", "Digital",
    "Bundles", "Bundle Lines", "Variations", "Batches",
    "Sales", "Promotions", "Price History", "Competitors",
    "Waste Log", "Development", "Events",
    "Wholesale Customers", "Wholesale Orders",
    "Scenarios", "Reports",
    "Print Costing", "Print Quote", "Print Order",
    "Lists", "Help",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Start Here": start_here.build,
    "Settings": settings.build,
    "Dashboard": dashboard.build,
    "Retail Calculator": retail_calc.build,
    "Wholesale Calculator": wholesale_calc.build,
    "Profit & Break-Even": break_even.build,
    "Scenarios": scenarios.build,
    "Reports": reports.build,
    "Print Costing": printables.build_costing,
    "Print Quote": printables.build_quote,
    "Print Order": printables.build_order,
    "Lists": lists.build,
    "Help": help_tab.build,
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


def build_workbook(path: str = "Handmade Pricing Calculator.xlsx",
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

    # 3. Live lookup names (product IDs, material names, who buys wholesale).
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
        "subject": "Handmade product pricing, material and labour costing, "
                   "margins and wholesale prices",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("handmade pricing calculator, Etsy pricing "
                     "spreadsheet, craft pricing, material cost "
                     "tracker, labor cost calculator, profit margin "
                     "calculator, wholesale price calculator, "
                     "break-even calculator, craft business, "
                     "spreadsheet template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

### `handmade_pricing_spreadsheet/sample_data.py`

*1,384 lines · 1 class · 9 functions*

> Demo data: a season in the life of Willow & Thread Handmade.
>
> A one-maker studio in Asheville that crochets, pours candles, solders a
> little silver, sews a few bags and turns out the occasional walnut
> board - plus a pattern shop that sells while she sleeps. Every tab
> opens with believable numbers so the dashboard, the price calculators
> and the reports mean something the moment the file is opened, and
> `--blank` strips the lot out for your own data.
>
> Everything is dated relative to the build date, so the workbook never
> looks stale.

| Defines | What it does |
|---------|--------------|
| `d(base, offset)` |  |
| `pid(n)` |  |
| `mid(n)` |  |
| `bid(n)` |  |
| `btid(n)` |  |
| `wcid(n)` |  |
| `woid(n)` |  |
| `class Sample` | fields: products, materials, material_usage, labor, packaging, shipping, overhead, channels |
| `get(cfg)` | Build (and cache) the demo data for this configuration |
| `_build(cfg)` |  |

````python
"""Demo data: a season in the life of Willow & Thread Handmade.

A one-maker studio in Asheville that crochets, pours candles, solders a
little silver, sews a few bags and turns out the occasional walnut
board - plus a pattern shop that sells while she sleeps. Every tab
opens with believable numbers so the dashboard, the price calculators
and the reports mean something the moment the file is opened, and
`--blank` strips the lot out for your own data.

Everything is dated relative to the build date, so the workbook never
looks stale.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .config import Config


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def pid(n: int) -> str:
    return f"P-{n:04d}"


def mid(n: int) -> str:
    return f"M-{n:04d}"


def bid(n: int) -> str:
    return f"BN-{n:04d}"


def btid(n: int) -> str:
    return f"BT-{n:04d}"


def wcid(n: int) -> str:
    return f"WC-{n:04d}"


def woid(n: int) -> str:
    return f"WO-{n:04d}"


@dataclass
class Sample:
    products: List[Dict[str, Any]] = field(default_factory=list)
    materials: List[Dict[str, Any]] = field(default_factory=list)
    material_usage: List[Dict[str, Any]] = field(default_factory=list)
    labor: List[Dict[str, Any]] = field(default_factory=list)
    packaging: List[Dict[str, Any]] = field(default_factory=list)
    shipping: List[Dict[str, Any]] = field(default_factory=list)
    overhead: List[Dict[str, Any]] = field(default_factory=list)
    channels: List[Dict[str, Any]] = field(default_factory=list)
    crochet: List[Dict[str, Any]] = field(default_factory=list)
    jewelry: List[Dict[str, Any]] = field(default_factory=list)
    candles: List[Dict[str, Any]] = field(default_factory=list)
    clothing: List[Dict[str, Any]] = field(default_factory=list)
    woodworking: List[Dict[str, Any]] = field(default_factory=list)
    digital: List[Dict[str, Any]] = field(default_factory=list)
    bundles: List[Dict[str, Any]] = field(default_factory=list)
    bundle_lines: List[Dict[str, Any]] = field(default_factory=list)
    variations: List[Dict[str, Any]] = field(default_factory=list)
    batches: List[Dict[str, Any]] = field(default_factory=list)
    sales: List[Dict[str, Any]] = field(default_factory=list)
    promotions: List[Dict[str, Any]] = field(default_factory=list)
    price_history: List[Dict[str, Any]] = field(default_factory=list)
    competitors: List[Dict[str, Any]] = field(default_factory=list)
    waste: List[Dict[str, Any]] = field(default_factory=list)
    development: List[Dict[str, Any]] = field(default_factory=list)
    events: List[Dict[str, Any]] = field(default_factory=list)
    wholesale_customers: List[Dict[str, Any]] = field(default_factory=list)
    wholesale_orders: List[Dict[str, Any]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------

_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg) if cfg.include_sample_data else Sample()
    return _CACHE[key]


def _build(cfg: Config) -> Sample:
    t = cfg.today
    s = Sample()

    # ---------------------------------------------------------- materials --
    # pack_qty is the number of `unit`s in the pack, so the cost per gram,
    # per centimetre or per item falls out of the pack price by itself.
    s.materials = [
        dict(name="Merino worsted yarn, natural", type="Yarn & thread",
             detail="100 g skein, 200 m", supplier="Blue Ridge Fibre Co",
             code="BRF-MW-NAT", pack_qty=100, unit="g", pack_cost=9.40,
             pack_ship=0.00, discount=0.00, waste=0.04, stock=1850,
             reorder=600, last_cost=8.60, priced_on=d(t, -24),
             notes="Up 9% this spring. Hats and cowls."),
        dict(name="Merino worsted yarn, moss", type="Yarn & thread",
             detail="100 g skein, 200 m", supplier="Blue Ridge Fibre Co",
             code="BRF-MW-MSS", pack_qty=100, unit="g", pack_cost=9.40,
             pack_ship=0.00, discount=0.00, waste=0.04, stock=700,
             reorder=600, last_cost=8.60, priced_on=d(t, -24)),
        dict(name="Cotton 8/4 twine", type="Yarn & thread",
             detail="250 g cone", supplier="Carolina Cotton",
             code="CC-84-ECR", pack_qty=250, unit="g", pack_cost=12.80,
             pack_ship=0.00, discount=0.05, waste=0.03, stock=1100,
             reorder=500, last_cost=12.80, priced_on=d(t, -96),
             notes="Market bags. 5% trade discount over 10 cones."),
        dict(name="Cotton DK, mixed colours", type="Yarn & thread",
             detail="50 g ball", supplier="Blue Ridge Fibre Co",
             code="BRF-DK-MIX", pack_qty=50, unit="g", pack_cost=3.20,
             pack_ship=0.00, waste=0.08, stock=640, reorder=300,
             last_cost=3.05, priced_on=d(t, -62),
             notes="Amigurumi. Waste is high - colour changes."),
        dict(name="Polyester toy stuffing", type="Other supply",
             detail="500 g bag", supplier="Craft Warehouse",
             code="CW-STF-500", pack_qty=500, unit="g", pack_cost=8.90,
             pack_ship=0.00, waste=0.02, stock=1400, reorder=500,
             last_cost=8.90, priced_on=d(t, -120)),
        dict(name="Safety eyes 10 mm", type="Elastic & notions",
             detail="Pack of 20 pairs", supplier="Craft Warehouse",
             code="CW-EYE-10", pack_qty=20, unit="pair", pack_cost=4.60,
             pack_ship=0.00, waste=0.00, stock=34, reorder=10,
             last_cost=4.60, priced_on=d(t, -120)),
        dict(name="Sterling silver wire 1.5 mm", type="Metal & wire",
             detail="Half hard, 30 g coil", supplier="Rio Metals",
             code="RIO-SS-15", pack_qty=30, unit="g", pack_cost=42.00,
             pack_ship=4.50, discount=0.00, waste=0.06, stock=74,
             reorder=30, last_cost=36.00, priced_on=d(t, -18),
             notes="Silver moved 17% in six months. Re-price rings."),
        dict(name="Sterling silver chain 1.2 mm", type="Metal & wire",
             detail="Fine cable, sold by the metre",
             supplier="Rio Metals", code="RIO-CH-12", pack_qty=100,
             unit="cm", pack_cost=14.00, pack_ship=0.00, waste=0.05,
             stock=420, reorder=150, last_cost=12.40,
             priced_on=d(t, -18)),
        dict(name="Sea glass cabochons", type="Stones & crystals",
             detail="Tumbled, 15-20 mm", supplier="Tidewater Finds",
             code="TW-SG-MIX", pack_qty=10, unit="each", pack_cost=18.00,
             pack_ship=3.20, waste=0.05, stock=22, reorder=10,
             last_cost=18.00, priced_on=d(t, -55)),
        dict(name="Jump rings & clasps, sterling",
             type="Beads & findings", detail="Mixed pack of 50",
             supplier="Rio Metals", code="RIO-FND-50", pack_qty=50,
             unit="each", pack_cost=9.50, pack_ship=0.00, waste=0.02,
             stock=180, reorder=60, last_cost=8.90,
             priced_on=d(t, -18)),
        dict(name="Coconut soy wax", type="Wax",
             detail="5 kg box, 464 blend", supplier="Blue Ridge Candle",
             code="BRC-CS-5", pack_qty=5, unit="kg", pack_cost=38.00,
             pack_ship=6.00, discount=0.00, waste=0.03, stock=11,
             reorder=4, last_cost=34.50, priced_on=d(t, -31),
             notes="Box price includes freight - that is why pack "
                   "delivery is filled in."),
        dict(name="Amber & fir fragrance oil", type="Fragrance & oils",
             detail="500 ml bottle", supplier="Blue Ridge Candle",
             code="BRC-FO-AF", pack_qty=500, unit="ml", pack_cost=42.00,
             pack_ship=0.00, waste=0.02, stock=720, reorder=250,
             last_cost=39.00, priced_on=d(t, -31)),
        dict(name="Lavender fragrance oil", type="Fragrance & oils",
             detail="250 ml bottle", supplier="Blue Ridge Candle",
             code="BRC-FO-LAV", pack_qty=250, unit="ml", pack_cost=26.00,
             pack_ship=0.00, waste=0.02, stock=180, reorder=100,
             last_cost=26.00, priced_on=d(t, -31)),
        dict(name="Cotton wick CD-12, pretabbed", type="Wicks",
             detail="Pack of 100", supplier="Blue Ridge Candle",
             code="BRC-WK-CD12", pack_qty=100, unit="each",
             pack_cost=12.00, pack_ship=0.00, waste=0.02, stock=260,
             reorder=100, last_cost=12.00, priced_on=d(t, -31)),
        dict(name="Amber glass jar 220 ml", type="Vessels & jars",
             detail="With bamboo lid, case of 24",
             supplier="Glassware Direct", code="GD-AM-220", pack_qty=24,
             unit="each", pack_cost=33.60, pack_ship=8.40, waste=0.02,
             stock=46, reorder=24, last_cost=31.20,
             priced_on=d(t, -40),
             notes="Breakage in transit is why the waste allowance "
                   "is not zero."),
        dict(name="Candle tin 120 ml", type="Vessels & jars",
             detail="Seamless, case of 48", supplier="Glassware Direct",
             code="GD-TIN-120", pack_qty=48, unit="each",
             pack_cost=38.40, pack_ship=0.00, waste=0.01, stock=96,
             reorder=48, last_cost=38.40, priced_on=d(t, -40)),
        dict(name="Washed linen, flax", type="Fabric",
             detail="150 cm wide, sold by the metre",
             supplier="Mill End Fabrics", code="MEF-LIN-FLX",
             pack_qty=300, unit="cm", pack_cost=54.00, pack_ship=0.00,
             waste=0.12, stock=640, reorder=200, last_cost=49.50,
             priced_on=d(t, -27),
             notes="Quantities in centimetres of length at full "
                   "width. 12% layout waste on the bias."),
        dict(name="Cotton lining, natural", type="Fabric",
             detail="112 cm wide", supplier="Mill End Fabrics",
             code="MEF-COT-NAT", pack_qty=500, unit="cm",
             pack_cost=32.00, pack_ship=0.00, waste=0.10, stock=380,
             reorder=150, last_cost=32.00, priced_on=d(t, -27)),
        dict(name="Zip, webbing & hardware set",
             type="Hardware & fixings", detail="Antique brass, per bag",
             supplier="Bag Makers Supply", code="BMS-HW-AB",
             pack_qty=10, unit="each", pack_cost=16.00, pack_ship=0.00,
             waste=0.00, stock=18, reorder=10, last_cost=14.80,
             priced_on=d(t, -47)),
        dict(name="Woven label, folded", type="Labels & tags",
             detail="Pack of 100", supplier="Label Lab",
             code="LL-WV-100", pack_qty=100, unit="each",
             pack_cost=24.00, pack_ship=0.00, waste=0.01, stock=210,
             reorder=50, last_cost=24.00, priced_on=d(t, -150)),
        dict(name="Walnut board, 25 mm", type="Timber & boards",
             detail="Kiln dried, 180 cm x 20 cm",
             supplier="Appalachian Timber", code="AT-WAL-25",
             pack_qty=180, unit="cm", pack_cost=76.00, pack_ship=0.00,
             waste=0.25, stock=240, reorder=180, last_cost=68.00,
             priced_on=d(t, -21),
             notes="A quarter of every board is offcut and sawdust. "
                   "Pretending otherwise is why boards feel "
                   "unprofitable."),
        dict(name="White oak board, 20 mm", type="Timber & boards",
             detail="240 cm x 15 cm", supplier="Appalachian Timber",
             code="AT-OAK-20", pack_qty=240, unit="cm", pack_cost=68.00,
             pack_ship=0.00, waste=0.22, stock=180, reorder=120,
             last_cost=61.00, priced_on=d(t, -21)),
        dict(name="Hard wax oil, satin", type="Finish & glue",
             detail="500 ml tin", supplier="Appalachian Timber",
             code="AT-HWO-500", pack_qty=500, unit="ml",
             pack_cost=32.00, pack_ship=0.00, waste=0.05, stock=310,
             reorder=150, last_cost=32.00, priced_on=d(t, -21)),
        dict(name="Sanding discs & glue", type="Tools & consumables",
             detail="Mixed grits, per project pack",
             supplier="Appalachian Timber", code="AT-ABR-MIX",
             pack_qty=25, unit="each", pack_cost=18.50, pack_ship=0.00,
             waste=0.00, stock=14, reorder=10, last_cost=18.50,
             priced_on=d(t, -70)),
    ]

    # ----------------------------------------------------------- products --
    s.products = [
        dict(name="Willow cable beanie", sku_input="", category="Crochet & knitting",
             status="Active", season="Autumn", difficulty="Moderate",
             channel="Etsy", pack_profile="Standard",
             ship_option="First class - Domestic to 250g", weight=140,
             stock=9, price_input=78.00, made=d(t, -12),
             launched=d(t, -400),
             url="willowandthread.co/beanie",
             notes="Costed on the Crochet tab, not the bill of "
                   "materials."),
        dict(name="Chunky market bag", sku_input="",
             category="Bags & leather", status="Active",
             season="All year", difficulty="Simple", channel="Etsy",
             pack_profile="Budget",
             ship_option="First class - Domestic to 500g", weight=320,
             stock=6, price_input=75.00, made=d(t, -20),
             launched=d(t, -520)),
        dict(name="Woodland fox amigurumi", sku_input="",
             category="Toys & amigurumi", status="Active",
             season="Christmas", difficulty="Moderate", channel="Etsy",
             pack_profile="Gift",
             ship_option="First class - Domestic to 500g", weight=180,
             stock=4, price_input=78.00, made=d(t, -7),
             launched=d(t, -300)),
        dict(name="Hammered silver stacking ring", sku_input="",
             category="Jewelry", status="Active", season="All year",
             difficulty="Moderate", channel="Etsy",
             pack_profile="Gift",
             ship_option="First class - Domestic to 100g", weight=25,
             stock=14, price_input=58.00, made=d(t, -9),
             launched=d(t, -620)),
        dict(name="Sea glass pendant", sku_input="",
             category="Jewelry", status="Active", season="Summer",
             difficulty="Simple", channel="Etsy", pack_profile="Gift",
             ship_option="First class - Domestic to 100g", weight=30,
             stock=11, price_input=52.00, made=d(t, -15),
             launched=d(t, -420)),
        dict(name="Amber & fir candle 220 ml", sku_input="",
             category="Candles & home fragrance", status="Active",
             season="Winter", difficulty="Simple", channel="Etsy",
             pack_profile="Standard",
             ship_option="Priority - Domestic to 1,000g", weight=520,
             stock=22, price_input=32.00, made=d(t, -5),
             launched=d(t, -260)),
        dict(name="Lavender tin candle 120 ml", sku_input="",
             category="Candles & home fragrance", status="Active",
             season="All year", difficulty="Very simple",
             channel="Etsy", pack_profile="Budget",
             ship_option="First class - Domestic to 500g", weight=260,
             stock=31, price_input=19.00, made=d(t, -5),
             launched=d(t, -260)),
        dict(name="Linen crossbody bag", sku_input="",
             category="Bags & leather", status="Active",
             season="All year", difficulty="Complex", channel="Etsy",
             pack_profile="Standard",
             ship_option="Priority - Domestic to 1,000g", weight=380,
             stock=3, price_input=195.00, made=d(t, -30),
             launched=d(t, -180)),
        dict(name="Linen pinafore apron", sku_input="",
             category="Clothing & accessories", status="Active",
             season="All year", difficulty="Moderate",
             channel="Own website", pack_profile="Standard",
             ship_option="Priority - Domestic to 1,000g", weight=420,
             stock=5, price_input=148.00, made=d(t, -40),
             launched=d(t, -210)),
        dict(name="Walnut serving board", sku_input="",
             category="Woodworking", status="Active", season="All year",
             difficulty="Moderate", channel="Own website",
             pack_profile="Standard",
             ship_option="Courier - Domestic to 5,000g", weight=1400,
             stock=4, price_input=135.00, made=d(t, -18),
             launched=d(t, -340)),
        dict(name="Oak floating shelf 60 cm", sku_input="",
             category="Woodworking", status="Made to order",
             season="All year", difficulty="Simple",
             channel="Own website", pack_profile="Budget",
             ship_option="Courier - Domestic to 5,000g", weight=2100,
             stock=0, price_input=None, made=d(t, -60),
             launched=d(t, -150),
             notes="No price yet - the dashboard is right to "
                   "complain about it."),
        dict(name="Cable beanie crochet pattern", sku_input="",
             category="Digital download", status="Active",
             season="Autumn", difficulty="Very simple", channel="Etsy",
             weight=0, stock=0, price_input=8.50,
             launched=d(t, -380),
             notes="Instant download - no packaging profile, so "
                   "nothing is charged for a box that is never "
                   "posted."),
        dict(name="Woodland fox pattern", sku_input="",
             category="Digital download", status="Active",
             season="Christmas", difficulty="Very simple",
             channel="Etsy", weight=0, stock=0,
             price_input=9.50, launched=d(t, -120)),
        dict(name="Beeswax wrap set of three", sku_input="",
             category="Other handmade", status="Prototype",
             season="All year", difficulty="Simple", channel="Etsy",
             pack_profile="Budget",
             ship_option="First class - Domestic to 250g", weight=120,
             stock=0, price_input=None, launched=None,
             notes="Still costing it. No materials entered yet."),
    ]

    # ---------------------------------------------------- bill of materials
    # Only the products that are costed the generic way have BOM lines -
    # anything on a specialist tab is costed there instead, which is what
    # the check column on those tabs is watching for.
    s.material_usage = [
        dict(product=pid(2), material=mid(3), qty=240, waste=0.03,
             optional="No", notes="Weighed off the cone."),
        dict(product=pid(2), material=mid(20), qty=1, optional="No",
             notes="Woven label in the side seam."),
        dict(product=pid(3), material=mid(4), qty=85, waste=0.08,
             optional="No", notes="Three colours, so more waste."),
        dict(product=pid(3), material=mid(5), qty=60, optional="No"),
        dict(product=pid(3), material=mid(6), qty=1, optional="No",
             notes="One pair of safety eyes."),
        dict(product=pid(3), material=mid(20), qty=1, optional="No"),
        dict(product=pid(8), material=mid(17), qty=70, waste=0.12,
             optional="No", notes="Outer, cut on the straight grain."),
        dict(product=pid(8), material=mid(18), qty=60, waste=0.08,
             optional="No", notes="Lining and inside pocket."),
        dict(product=pid(8), material=mid(19), qty=1, optional="No",
             notes="Zip, webbing, slider and two rings."),
        dict(product=pid(8), material=mid(20), qty=1, optional="No"),
        dict(product=pid(10), material=mid(21), qty=55, waste=0.25,
             optional="No",
             notes="A quarter of the board is offcut - that is the "
                   "honest number."),
        dict(product=pid(10), material=mid(23), qty=18, waste=0.05,
             optional="No", notes="Three coats of hard wax oil."),
        dict(product=pid(10), material=mid(24), qty=1.5, optional="No",
             notes="Discs die fast on walnut."),
        dict(product=pid(10), material=mid(20), qty=1, optional="Yes",
             notes="Branded tag, only on gift orders."),
    ]

    # ----------------------------------------------------------- the hours
    s.labor = [
        dict(product=pid(2), stage="Making", who="Me", minutes=85,
             basis="Per item", skilled="No",
             notes="Base, sides and handles."),
        dict(product=pid(2), stage="Weaving in", who="Me", minutes=10,
             basis="Per item", skilled="No"),
        dict(product=pid(2), stage="Packing", who="Me", minutes=6,
             basis="Per item", skilled="No"),
        dict(product=pid(2), stage="Photography", who="Me", minutes=25,
             basis="Per batch", batch_qty=6, skilled="No",
             notes="One shoot covers the batch."),
        dict(product=pid(3), stage="Making", who="Me", minutes=90,
             basis="Per item", skilled="Yes",
             notes="Body, head, four legs, tail and ears."),
        dict(product=pid(3), stage="Assembly", who="Me", minutes=18,
             basis="Per item", skilled="Yes",
             notes="Sewing up and getting the face right."),
        dict(product=pid(3), stage="Finishing", who="Me", minutes=8,
             basis="Per item", skilled="No"),
        dict(product=pid(3), stage="Packing", who="Me", minutes=4,
             basis="Per item", skilled="No"),
        dict(product=pid(8), stage="Cutting & prep", who="Me",
             minutes=30, basis="Per item", skilled="Yes"),
        dict(product=pid(8), stage="Sewing", who="Me", minutes=80,
             basis="Per item", skilled="Yes",
             notes="Zip, lining, strap, topstitching."),
        dict(product=pid(8), stage="Finishing", who="Me", minutes=15,
             basis="Per item", skilled="No"),
        dict(product=pid(8), stage="Packing", who="Me", minutes=8,
             basis="Per item", skilled="No"),
        dict(product=pid(8), stage="Photography", who="Me", minutes=30,
             basis="Per batch", batch_qty=4, skilled="No"),
        dict(product=pid(10), stage="Cutting & prep", who="Me",
             minutes=25, basis="Per item", skilled="Yes",
             notes="Milling, thicknessing, cutting to length."),
        dict(product=pid(10), stage="Sanding & smoothing", who="Me",
             minutes=35, basis="Per item", skilled="No",
             notes="Always longer than anybody admits."),
        dict(product=pid(10), stage="Painting & staining", who="Me",
             minutes=20, basis="Per item", skilled="No",
             notes="Three coats of oil, with waiting in between."),
        dict(product=pid(10), stage="Packing", who="Me", minutes=12,
             basis="Per item", skilled="No",
             notes="Boards need real protection in the post."),
        dict(product=pid(10), stage="Photography", who="Me",
             minutes=25, basis="Per batch", batch_qty=3, skilled="No"),
    ]

    # --------------------------------------------------------- packaging --
    s.packaging = [
        dict(name="Kraft mailer box, small", type="Mailer & box",
             supplier="Packmart", pack_qty=50, pack_cost=34.50, qty=1,
             weight=45, budget="Yes", standard="Yes", gift="Yes",
             premium="Yes", reusable="No"),
        dict(name="Poly mailer", type="Polybag", supplier="Packmart",
             pack_qty=100, pack_cost=14.00, qty=1, weight=12,
             budget="Yes", standard="No", gift="No", premium="No",
             reusable="No", notes="Patterns and anything soft."),
        dict(name="Printed tissue paper", type="Tissue & wrap",
             supplier="Noissue", pack_qty=500, pack_cost=42.00, qty=2,
             weight=4, budget="No", standard="Yes", gift="Yes",
             premium="Yes", reusable="No"),
        dict(name="Kraft sticker seal", type="Label & sticker",
             supplier="Noissue", pack_qty=500, pack_cost=22.00, qty=2,
             weight=1, budget="Yes", standard="Yes", gift="Yes",
             premium="Yes", reusable="No"),
        dict(name="Thank you card", type="Thank you card",
             supplier="Moo", pack_qty=250, pack_cost=38.00, qty=1,
             weight=5, budget="No", standard="Yes", gift="Yes",
             premium="Yes", reusable="No",
             notes="Handwritten. Worth more than it costs."),
        dict(name="Care card", type="Care card", supplier="Moo",
             pack_qty=250, pack_cost=26.00, qty=1, weight=3,
             budget="No", standard="Yes", gift="Yes", premium="Yes",
             reusable="No", notes="Washing, burning and storing."),
        dict(name="Baker's twine", type="Ribbon & twine",
             supplier="Packmart", pack_qty=80, pack_cost=9.00, qty=1,
             weight=2, budget="No", standard="No", gift="Yes",
             premium="Yes", reusable="No"),
        dict(name="Rigid gift box", type="Gift box",
             supplier="Packmart", pack_qty=50, pack_cost=76.00, qty=1,
             weight=95, budget="No", standard="No", gift="Yes",
             premium="Yes", reusable="Yes",
             notes="The thing people keep. Worth the money on "
                   "jewelry."),
        dict(name="Bubble wrap sheet", type="Fragile protection",
             supplier="Packmart", pack_qty=100, pack_cost=19.00, qty=1,
             weight=15, budget="No", standard="Yes", gift="Yes",
             premium="Yes", reusable="No",
             notes="Candles and boards."),
        dict(name="Wax seal & ribbon", type="Ribbon & twine",
             supplier="Etsy supplier", pack_qty=100, pack_cost=28.00,
             qty=1, weight=4, budget="No", standard="No", gift="No",
             premium="Yes", reusable="No"),
    ]

    # ---------------------------------------------------------- shipping --
    s.shipping = [
        dict(service="First class", zone="Domestic", max_weight=100,
             cost=4.25, packaging_extra=0.00, charge=4.95,
             payer="Buyer pays", tracked="No", days=3, carrier="USPS",
             notes="Jewelry in a padded envelope."),
        dict(service="First class", zone="Domestic", max_weight=250,
             cost=5.40, packaging_extra=0.00, charge=5.95,
             payer="Buyer pays", tracked="Yes", days=3, carrier="USPS"),
        dict(service="First class", zone="Domestic", max_weight=500,
             cost=7.10, packaging_extra=0.00, charge=6.95,
             payer="Buyer pays", tracked="Yes", days=3, carrier="USPS",
             notes="You absorb 15 cents on every one of these."),
        dict(service="Priority", zone="Domestic", max_weight=1000,
             cost=9.85, packaging_extra=0.35, charge=8.95,
             payer="Buyer pays", tracked="Yes", days=2, carrier="USPS",
             notes="Candles and bags."),
        dict(service="Priority", zone="Domestic", max_weight=2000,
             cost=12.40, packaging_extra=0.35, charge=11.95,
             payer="Buyer pays", tracked="Yes", days=2, carrier="USPS"),
        dict(service="Courier", zone="Domestic", max_weight=5000,
             cost=18.60, packaging_extra=1.20, charge=16.00,
             payer="Flat rate", tracked="Yes", days=4, carrier="UPS",
             notes="Boards. The flat rate loses a little on every "
                   "order and sells a lot more of them."),
        dict(service="International tracked", zone="Rest of world",
             max_weight=500, cost=24.50, packaging_extra=0.00,
             charge=24.50, payer="Buyer pays", tracked="Yes", days=12,
             carrier="USPS", notes="Customs form every time."),
        dict(service="Collection / pickup", zone="Local", cost=0.00,
             charge=0.00, payer="Collection only", tracked="No",
             days=0, carrier="Studio",
             notes="Local buyers collect from the studio."),
    ]

    # ----------------------------------------------------------- overhead --
    s.overhead = [
        dict(name="Studio rent", type="Studio rent",
             supplier="River Arts District", amount=260.00,
             frequency="Monthly", start=d(t, -700), essential="Yes",
             notes="Shared bench space, 24 hour access."),
        dict(name="Electricity & heat", type="Utilities",
             supplier="Duke Energy", amount=48.00, frequency="Monthly",
             start=d(t, -700), essential="Yes",
             notes="The wax melter is not cheap to run."),
        dict(name="Maker's insurance", type="Insurance",
             supplier="ACT Insurance", amount=288.00,
             frequency="Yearly", start=d(t, -200), essential="Yes",
             notes="Product liability - candles need it."),
        dict(name="Etsy Plus", type="Shop subscription",
             supplier="Etsy", amount=10.00, frequency="Monthly",
             start=d(t, -400), essential="No"),
        dict(name="Website & domain", type="Website & domain",
             supplier="Squarespace", amount=192.00, frequency="Yearly",
             start=d(t, -300), essential="Yes"),
        dict(name="Design software", type="Software",
             supplier="Adobe", amount=14.99, frequency="Monthly",
             start=d(t, -500), essential="No",
             notes="Patterns and listing photos."),
        dict(name="Bookkeeping", type="Accounting",
             supplier="Hill & Co", amount=240.00, frequency="Yearly",
             start=d(t, -300), essential="Yes"),
        dict(name="Phone & internet (business share)",
             type="Phone & internet", supplier="Spectrum",
             amount=35.00, frequency="Monthly", start=d(t, -700),
             essential="Yes", notes="Half the household bill."),
        dict(name="Tool replacement fund", type="Tool replacement",
             supplier="-", amount=360.00, frequency="Yearly",
             start=d(t, -365), essential="Yes",
             notes="Hooks, blades, sanding belts, the next wax "
                   "melter. Put it aside before you need it."),
        dict(name="Photography props & backdrops",
             type="Photography props", supplier="Various",
             amount=120.00, frequency="Yearly", start=d(t, -200),
             essential="No"),
        dict(name="Guild membership", type="Memberships",
             supplier="Southern Highland Craft Guild", amount=60.00,
             frequency="Yearly", start=d(t, -250), essential="No",
             notes="Pays for itself in fair bookings."),
    ]

    # ----------------------------------------------------------- channels --
    reviewed = d(t, -38)
    s.channels = [
        dict(name="Etsy", kind="Percent + fixed", listing_fee=0.20,
             listing_sales=2, transaction=0.065, processing_pct=0.03,
             processing_fixed=0.25, ads_fee=0.15, ads_share=0.20,
             regulatory=0.0040, currency_fee=0.00, other_pct=0.00,
             other_fixed=0.00, subscription=10.00,
             fees_on_shipping="Yes", reviewed=reviewed,
             url="etsy.com/legal/fees",
             notes="Offsite ads are 15% but only on the orders that "
                   "come through them - about one in five here. "
                   "Verify the current official fee schedule for "
                   "your country before you rely on these."),
        dict(name="Own website", kind="Percent + fixed",
             listing_fee=0.00, listing_sales=1, transaction=0.00,
             processing_pct=0.029, processing_fixed=0.30, ads_fee=0.00,
             ads_share=0.00, regulatory=0.00, currency_fee=0.00,
             other_pct=0.00, other_fixed=0.00, subscription=16.00,
             fees_on_shipping="Yes", reviewed=reviewed,
             url="squarespace.com/pricing",
             notes="Card processing only. The subscription belongs "
                   "in overhead, not in the price."),
        dict(name="Craft fair", kind="Percent + fixed",
             listing_fee=0.00, listing_sales=1, transaction=0.00,
             processing_pct=0.026, processing_fixed=0.10, ads_fee=0.00,
             ads_share=0.00, regulatory=0.00, currency_fee=0.00,
             other_pct=0.00, other_fixed=0.00, subscription=0.00,
             fees_on_shipping="No", reviewed=reviewed,
             url="squareup.com/us/en/pricing",
             notes="Card reader rate. The stall fee is on the "
                   "Events tab, where it belongs."),
        dict(name="Wholesale", kind="None", listing_fee=0.00,
             listing_sales=1, transaction=0.00, processing_pct=0.00,
             processing_fixed=0.00, ads_fee=0.00, ads_share=0.00,
             regulatory=0.00, currency_fee=0.00, other_pct=0.00,
             other_fixed=0.00, subscription=0.00,
             fees_on_shipping="No", reviewed=reviewed,
             notes="No fees - the discount is the cost."),
        dict(name="Instagram / direct", kind="Percent + fixed",
             listing_fee=0.00, listing_sales=1, transaction=0.00,
             processing_pct=0.029, processing_fixed=0.30, ads_fee=0.00,
             ads_share=0.00, regulatory=0.00, currency_fee=0.00,
             other_pct=0.00, other_fixed=0.00, subscription=0.00,
             fees_on_shipping="Yes", reviewed=reviewed,
             notes="A payment link. The cheapest place you sell."),
        dict(name="Amazon Handmade", kind="Percent of order",
             listing_fee=0.00, listing_sales=1, transaction=0.15,
             processing_pct=0.00, processing_fixed=0.00, ads_fee=0.00,
             ads_share=0.00, regulatory=0.00, currency_fee=0.00,
             other_pct=0.00, other_fixed=0.00, subscription=0.00,
             fees_on_shipping="Yes", reviewed=reviewed,
             url="sell.amazon.com/handmade",
             notes="15% referral fee. Nothing listed here yet - "
                   "the margin would have to be built in first."),
        dict(name="Other", kind="None", listing_fee=0.00,
             listing_sales=1, transaction=0.00, processing_pct=0.00,
             processing_fixed=0.00, subscription=0.00,
             fees_on_shipping="No", reviewed=reviewed,
             notes="Gifts, swaps and anything that does not fit."),
    ]

    # ------------------------------------------------------ craft tabs --
    # Each specialist tab costs the pieces that belong to it.  A piece
    # with a product ID feeds the catalog; a piece without one is a
    # costing you are still thinking about.
    s.crochet = [
        dict(product=pid(1), name="Willow cable beanie",
             pattern="Willow cable (my own)", size="Standard",
             yarn_weight="Worsted / Aran", hook="5.0 mm", yarn=mid(1),
             grams=95, waste=0.05, notions=0.25, stuffing=0.00,
             pattern_cost=0.00, hours_input=1.25, finish_mins=10,
             price_input=78.00,
             notes="Weighed before and after. 95 g is the real "
                   "number, not the 80 g the pattern claims."),
        dict(name="Moss chunky cowl", pattern="Moss cowl (my own)",
             size="Standard", yarn_weight="Bulky / chunky",
             hook="8.0 mm", yarn=mid(2), grams=150, waste=0.05,
             notions=0.25, hours_input=1.8, finish_mins=12,
             price_input=78.00,
             notes="No product ID yet - costing it before I list it."),
        dict(name="Cotton baby blanket, 90 cm",
             pattern="Granny square blanket", size="Large",
             yarn_weight="DK", hook="4.0 mm", yarn=mid(4), grams=480,
             waste=0.06, notions=0.25, hours_input=9.5,
             finish_mins=35, price_input=215.00,
             notes="Nine and a half hours. Look at what the hourly "
                   "return column says before listing this."),
    ]

    s.jewelry = [
        dict(product=pid(4), name="Hammered silver stacking ring",
             metal="Sterling silver", metal_id=mid(7), metal_g=3.8,
             scrap=0.15, recovery=0.60, stone="None", stone_count=0,
             stone_cost=0.00, findings=0.00, solder=0.45, plating=0.00,
             hallmark=0.00, design_mins=0, bench_mins=32,
             polish_mins=10, price_input=58.00,
             notes="Scrap is real on a soldered ring - 15% of it, "
                   "and I re-melt about 60%."),
        dict(product=pid(5), name="Sea glass pendant",
             metal="Sterling silver", metal_id=mid(7), metal_g=1.6,
             scrap=0.10, recovery=0.60, stone="Sea glass",
             stone_count=1, stone_cost=2.12, chain_cm=45,
             chain_id=mid(8), findings=0.57, solder=0.25,
             plating=0.00, hallmark=0.00, design_mins=0,
             bench_mins=26, polish_mins=8, price_input=52.00,
             notes="Chain is 45 cm of the fine cable, costed by the "
                   "centimetre."),
        dict(name="Copper cuff, hammered", metal="Copper",
             metal_rate_input=0.0620, metal_g=28, scrap=0.08,
             recovery=0.00, stone="None", stone_count=0,
             findings=0.00, solder=0.20, design_mins=15,
             bench_mins=35, polish_mins=10, price_input=52.00,
             notes="No material row for copper yet, so the rate a "
                   "gram is typed in instead."),
    ]

    s.candles = [
        dict(product=pid(6), name="Amber & fir candle 220 ml",
             wax_type="Coconut soy", wax_id=mid(11),
             vessel="Amber jar", vessel_ml=220, fill=0.85,
             density=0.90, melt_loss=0.05, fragrance_id=mid(12),
             fragrance_load=0.09, wick="Cotton", wick_count=1,
             wick_cost=0.12, vessel_cost=1.75, dye_cost=0.00,
             label_cost=0.24, batch_size=12, batch_mins=55,
             item_mins=6, cure_days=10, price_input=32.00,
             notes="Fill to 85% of the jar. Nobody pours to the "
                   "brim."),
        dict(product=pid(7), name="Lavender tin candle 120 ml",
             wax_type="Coconut soy", wax_id=mid(11), vessel="Tin",
             vessel_ml=120, fill=0.88, density=0.90, melt_loss=0.05,
             fragrance_id=mid(13), fragrance_load=0.08, wick="Cotton",
             wick_count=1, wick_cost=0.12, vessel_cost=0.80,
             dye_cost=0.00, label_cost=0.24, batch_size=24,
             batch_mins=65, item_mins=4, cure_days=10,
             price_input=19.00,
             notes="The batch minutes are spread over 24 tins, "
                   "which is why these look cheap to make."),
        dict(name="Wax melt clamshell", wax_type="Coconut soy",
             wax_id=mid(11), vessel="Clamshell", vessel_ml=70,
             fill=0.95, density=0.90, melt_loss=0.04,
             fragrance_id=mid(13), fragrance_load=0.10, wick="Cotton",
             wick_count=0, wick_cost=0.00, vessel_cost=0.38,
             label_cost=0.18, batch_size=36, batch_mins=50,
             item_mins=2, cure_days=7, price_input=8.50,
             notes="No wick at all - wick count zero, and the maths "
                   "copes."),
    ]

    s.clothing = [
        dict(product=pid(9), name="Linen pinafore apron",
             garment_type="Apron", size="M", fabric_id=mid(17),
             fabric_cm=180, layout_waste=0.12, interfacing=0.40,
             thread=0.35, notions=1.20, labels=0.24, cut_mins=25,
             sew_mins=55, finish_mins=20, press_mins=10, fit_mins=0,
             price_input=148.00,
             notes="Size M is the base. The size factor scales the "
                   "fabric and the time for everything else."),
        dict(name="Toddler pinafore, age 3", garment_type="Pinafore",
             size="Toddler", size_factor=0.55, fabric_id=mid(17),
             fabric_cm=95, layout_waste=0.15, lining_id=mid(18),
             lining_cm=60, interfacing=0.20, thread=0.25,
             notions=0.90, labels=0.24, cut_mins=20, sew_mins=48,
             finish_mins=15, press_mins=8, fit_mins=0,
             price_input=85.00,
             notes="Half the fabric, nearly all of the time. This "
                   "is why children's clothes are hard to price."),
    ]

    s.woodworking = [
        dict(product=pid(11), name="Oak floating shelf 60 cm",
             species="Oak", wood_id=mid(22), wood_qty=65,
             offcut_waste=0.22, hardware=6.40, glue=0.60,
             finish_type="Hard wax oil", finish_ml=22,
             finish_rate=0.0640, abrasives=1.80, machine_mins=18,
             machine_rate=8.00, mill_mins=15, build_mins=25,
             sand_mins=40, finish_mins=25, dry_days=2,
             notes="No price typed in yet - the suggested price on "
                   "this row is the one to beat."),
        dict(name="Live edge walnut shelf 90 cm", species="Walnut",
             wood_id=mid(21), wood_qty=95, offcut_waste=0.30,
             hardware=11.80, glue=0.80, finish_type="Hard wax oil",
             finish_ml=34, finish_rate=0.0640, abrasives=2.60,
             machine_mins=26, machine_rate=8.00, mill_mins=25,
             build_mins=35, sand_mins=55, finish_mins=35, dry_days=2,
             price_input=210.00,
             notes="Thirty percent offcut on a live edge slab. "
                   "Anybody quoting less has not measured."),
    ]

    s.digital = [
        dict(product=pid(12), name="Cable beanie crochet pattern",
             kind="Crochet pattern", formats="PDF",
             licence="Personal use", design_hours=6, write_hours=11,
             test_hours=4, photo_hours=3, tester_cost=0.00,
             software_cost=18.00, expected_sales=220,
             support_mins=1.5, update_hours=2, sold_so_far=143,
             listed=d(t, -380), price_input=8.50,
             notes="Two testers worked for a free pattern, so the "
                   "tester fee is zero and the hours are mine."),
        dict(product=pid(13), name="Woodland fox pattern",
             kind="Crochet pattern", formats="PDF, Video",
             licence="Small commercial", design_hours=9,
             write_hours=14, test_hours=5, photo_hours=4,
             tester_cost=45.00, software_cost=18.00,
             expected_sales=250, support_mins=2.5, update_hours=2,
             sold_so_far=38, listed=d(t, -120), price_input=9.50,
             notes="Amigurumi patterns generate questions. Two and "
                   "a half minutes of support a sale is honest."),
        dict(name="Christmas ornament SVG bundle",
             kind="SVG cut file", formats="SVG, PNG, DXF",
             licence="Small commercial", design_hours=12,
             write_hours=3, test_hours=2, photo_hours=2,
             tester_cost=0.00, software_cost=24.00,
             expected_sales=300, support_mins=0.5, update_hours=1,
             sold_so_far=0, listed=None, price_input=12.00,
             notes="Not listed yet. The break-even sales column "
                   "says how many it has to sell to pay for the "
                   "design time."),
    ]

    # ------------------------------------------------------- bundles --
    s.bundles = [
        dict(name="Cosy night in gift set", kind="Gift set",
             channel="Etsy", pack_profile="Gift", assembly_mins=12,
             discount=0.12, status="Active", season="Winter",
             notes="Beanie and the amber candle. The discount is "
                   "what the set gives away against buying both."),
        dict(name="Candle duo", kind="Multi-buy", channel="Etsy",
             pack_profile="Standard", assembly_mins=6, discount=0.10,
             status="Active", season="All year",
             notes="Buyer picks the two scents."),
        dict(name="New baby keepsake set", kind="Gift set",
             channel="Own website", pack_profile="Premium",
             assembly_mins=15, discount=0.08, price_input=92.00,
             status="Active", season="All year",
             notes="Priced by hand at 92 - the suggested price "
                   "column says whether that was brave or silly."),
    ]

    s.bundle_lines = [
        dict(bundle=bid(1), product=pid(1), qty=1, swap="No"),
        dict(bundle=bid(1), product=pid(6), qty=1, swap="Yes",
             notes="Scent of their choice."),
        dict(bundle=bid(2), product=pid(6), qty=1, swap="Yes"),
        dict(bundle=bid(2), product=pid(7), qty=1, swap="Yes"),
        dict(bundle=bid(3), product=pid(3), qty=1, swap="No",
             notes="The fox is the piece they remember."),
        dict(bundle=bid(3), product=pid(7), qty=1, swap="No"),
    ]

    # ----------------------------------------------------- variations --
    s.variations = [
        dict(product=pid(1), name="Beanie, extra large", kind="Size",
             size="Large", extra_mins=25, rush="Standard",
             channel="Etsy", offered="Yes",
             notes="More yarn and more rounds - the size factor "
                   "handles both."),
        dict(product=pid(1), name="Beanie with faux fur pom",
             kind="Personalisation", extra_material=3.40,
             extra_mins=8, rush="Standard", channel="Etsy",
             price_input=86.00, offered="Yes"),
        dict(product=pid(3), name="Fox with embroidered name",
             kind="Embroidery", extra_mins=22, personalisation=8.00,
             rush="Standard", channel="Etsy", offered="Yes",
             notes="Eight dollars for twenty two minutes. Look at "
                   "the uplift column before saying yes to this."),
        dict(product=pid(4), name="Ring, rush order",
             kind="Rush order", rush="Same week", channel="Etsy",
             offered="Yes",
             notes="Half again on the labour, because everything "
                   "else waits."),
        dict(product=pid(9), name="Apron, made to measure",
             kind="Size", size="Custom size", extra_mins=35,
             rush="Standard", channel="Own website", offered="Yes",
             notes="Fitting, a toile and a second pressing."),
        dict(product=pid(6), name="Candle, custom scent",
             kind="Scent", extra_material=1.20, extra_mins=10,
             rush="Standard", channel="Instagram / direct",
             offered="No",
             notes="Costed and declined - a ten minute change on a "
                   "32 dollar candle does not pay."),
    ]

    # -------------------------------------------------------- batches --
    s.batches = [
        dict(product=pid(6), started=d(t, -40), finished=d(t, -38),
             planned=12, made=12, seconds=2, sold_seconds=24.00,
             material_actual=64.80, hours_actual=2.6, extra_cost=0.00,
             status="Complete",
             notes="Two with sink holes - sold as seconds at the "
                   "fair."),
        dict(product=pid(7), started=d(t, -40), finished=d(t, -39),
             planned=24, made=24, seconds=1, sold_seconds=9.00,
             material_actual=71.50, hours_actual=3.1, extra_cost=0.00,
             status="Complete"),
        dict(product=pid(1), started=d(t, -25), finished=d(t, -18),
             planned=6, made=6, seconds=0, material_actual=58.40,
             hours_actual=10.8, extra_cost=0.00, status="Complete",
             notes="Ten point eight hours for six. That is 1.8 "
                   "each, not the 1.57 the costing assumes - and "
                   "the costing is the one that needs changing."),
        dict(product=pid(3), started=d(t, -14), finished=d(t, -7),
             planned=4, made=4, seconds=1, sold_seconds=22.00,
             material_actual=31.20, hours_actual=9.6,
             extra_cost=0.00, status="Complete",
             notes="One fox has a wonky ear. Sold at the pop-up."),
        dict(product=pid(10), started=d(t, -22), finished=d(t, -18),
             planned=3, made=3, seconds=0, material_actual=92.00,
             hours_actual=5.6, extra_cost=4.50, status="Complete",
             notes="Extra cost is a replacement sanding belt."),
        dict(product=pid(2), started=d(t, -6), planned=8,
             status="In progress",
             notes="Half crocheted. Nothing recorded as made yet, "
                   "which is why the row is flagged."),
    ]

    # ---------------------------------------------------- promotions --
    s.promotions = [
        dict(code="WELCOME10", name="Welcome to the shop",
             kind="Percent off", value=0.10, scope="Everything",
             start=d(t, -200), min_spend=0.00, base_margin=0.30,
             notes="Permanent, on the shop banner."),
        dict(code="AUTUMN20", name="Autumn sale",
             kind="Percent off", value=0.20, scope="Everything",
             start=d(t, -95), ends=d(t, -80), min_spend=0.00,
             base_margin=0.30,
             notes="Twenty percent off ate most of the margin. "
                   "Look at what it left before running it again."),
        dict(code="FREESHIP50", name="Free postage over 50",
             kind="Free shipping", value=7.10, scope="Order",
             start=d(t, -120), min_spend=50.00, base_margin=0.30,
             notes="Value is the postage you absorb on an average "
                   "qualifying order."),
        dict(code="BUNDLE5", name="Five off any set",
             kind="Amount off", value=5.00, scope="Bundles",
             start=d(t, -60), min_spend=60.00, limit=100,
             base_margin=0.26),
    ]

    # ------------------------------------------------- price history --
    s.price_history = [
        dict(date=d(t, -400), product=pid(1), old_price=48.00,
             new_price=65.00, cost_then=52.00, reason="First price",
             channel="Etsy", sold_before=0, sold_after=18,
             notes="Opening price was set by looking at other shops. "
                   "Then it was costed properly."),
        dict(date=d(t, -26), product=pid(1), old_price=65.00,
             new_price=78.00, cost_then=59.40,
             reason="Material cost rose", channel="Etsy",
             sold_before=11, sold_after=9,
             notes="Merino went up 9%. Sales did not move."),
        dict(date=d(t, -200), product=pid(2), old_price=62.00,
             new_price=75.00, cost_then=68.00,
             reason="Competitor move", channel="Etsy",
             sold_before=14, sold_after=16,
             notes="Still does not cover the hours. The bag is on "
                   "the list to redesign or retire."),
        dict(date=d(t, -30), product=pid(3), old_price=62.00,
             new_price=78.00, cost_then=88.00,
             reason="Underpriced", channel="Etsy",
             sold_before=6, sold_after=3,
             notes="Sixteen dollars more and it still loses money. "
                   "A two hour toy is a two hour toy."),
        dict(date=d(t, -18), product=pid(4), old_price=52.00,
             new_price=58.00, cost_then=33.90,
             reason="Material cost rose", channel="Etsy",
             sold_before=7, sold_after=6,
             notes="Silver moved 17%. This one had to move."),
        dict(date=d(t, -45), product=pid(5), old_price=44.00,
             new_price=52.00, cost_then=31.40,
             reason="Underpriced", channel="Etsy",
             sold_before=9, sold_after=7,
             notes="The margin column said nine percent. Eight "
                   "dollars fixed it and nobody complained."),
        dict(date=d(t, -120), product=pid(6), old_price=28.00,
             new_price=32.00, cost_then=12.80,
             reason="Rebrand / repositioning", channel="Etsy",
             sold_before=22, sold_after=26,
             notes="New label, better photos, higher price - and "
                   "it sold more, not less."),
        dict(date=d(t, -120), product=pid(7), old_price=17.00,
             new_price=19.00, cost_then=6.80, reason="Rounding tidy-up",
             channel="Etsy", sold_before=30, sold_after=28),
        dict(date=d(t, -180), product=pid(8), old_price=165.00,
             new_price=195.00, cost_then=128.00,
             reason="Labour rate change", channel="Etsy",
             sold_before=5, sold_after=4,
             notes="Put my own rate up from 20 to 25."),
        dict(date=d(t, -60), product=pid(9), old_price=128.00,
             new_price=148.00, cost_then=114.00,
             reason="Underpriced", channel="Own website",
             sold_before=4, sold_after=3,
             notes="Linen went up twice in a year and the price "
                   "had not moved once."),
        dict(date=d(t, -70), product=pid(10), old_price=118.00,
             new_price=135.00, cost_then=98.00,
             reason="Material cost rose", channel="Own website",
             sold_before=6, sold_after=5,
             notes="Walnut up 12%. The old price never covered the "
                   "sanding either."),
        dict(date=d(t, -90), product=pid(12), old_price=7.50,
             new_price=8.50, cost_then=5.00,
             reason="Testing a price", channel="Etsy",
             sold_before=18, sold_after=17,
             notes="A dollar more, the same number of sales. The "
                   "test paid for itself."),
        dict(date=d(t, -50), product=pid(13), old_price=6.50,
             new_price=9.50, cost_then=6.20,
             reason="Underpriced", channel="Etsy",
             sold_before=12, sold_after=9,
             notes="Two and a half minutes of support a sale has "
                   "to be paid for by somebody."),
    ]

    # -------------------------------------------------- competitors --
    s.competitors = [
        dict(product=pid(1), competitor="Highland Knits",
             shop="Etsy star seller", channel="Etsy", price=84.00,
             shipping=5.95, quality="Similar", handmade="Yes",
             reviews=1240, checked=d(t, -20),
             notes="Same yarn weight, no cable. Charging more."),
        dict(product=pid(1), competitor="FastFashionCo",
             shop="Amazon", channel="Amazon Handmade", price=19.99,
             shipping=0.00, quality="Mass produced", handmade="No",
             reviews=8600, checked=d(t, -20),
             notes="Machine made in a factory. Not the same "
                   "product and not your customer."),
        dict(product=pid(3), competitor="Thistle & Thread Toys",
             shop="Etsy", channel="Etsy", price=92.00, shipping=6.50,
             quality="Above mine", handmade="Yes", reviews=430,
             checked=d(t, -32),
             notes="Better photographs. That is the gap, not the "
                   "crochet."),
        dict(product=pid(4), competitor="Asheville Silverworks",
             shop="Own website", channel="Own website", price=72.00,
             shipping=0.00, quality="Above mine", handmade="Yes",
             reviews=96, checked=d(t, -14),
             notes="Hallmarked and boxed. Room to move up."),
        dict(product=pid(5), competitor="Tidewater Jewelry",
             shop="Etsy", channel="Etsy", price=46.00, shipping=4.50,
             quality="Similar", handmade="Yes", reviews=610,
             checked=d(t, -14)),
        dict(product=pid(6), competitor="Smoky Mountain Candle",
             shop="Etsy", channel="Etsy", price=28.00, shipping=7.95,
             quality="Similar", handmade="Yes", reviews=2100,
             checked=d(t, -26),
             notes="Cheaper jar, same wax. Their postage is "
                   "dearer, so the delivered price is level."),
        dict(product=pid(7), competitor="Smoky Mountain Candle",
             shop="Etsy", channel="Etsy", price=16.00, shipping=7.95,
             quality="Similar", handmade="Yes", reviews=2100,
             checked=d(t, -26)),
        dict(product=pid(8), competitor="Foothills Canvas",
             shop="Own website", channel="Own website", price=215.00,
             shipping=0.00, quality="Above mine", handmade="Yes",
             reviews=58, checked=d(t, -45),
             notes="Leather trim and free postage."),
        dict(product=pid(10), competitor="Blue Ridge Boards",
             shop="Craft fair", channel="Craft fair", price=112.00,
             shipping=0.00, quality="Below mine", handmade="Yes",
             reviews=0, checked=d(t, -35),
             notes="Thinner stock, machine oiled. Undercuts you "
                   "at every fair."),
        dict(product=pid(12), competitor="Loop & Lane Patterns",
             shop="Etsy", channel="Etsy", price=7.00, shipping=0.00,
             quality="Similar", handmade="N/A", reviews=3400,
             checked=d(t, -60),
             notes="Patterns are a volume game - the price is set "
                   "by the market, not by your hours."),
    ]

    # ------------------------------------------------------- waste --
    s.waste = [
        dict(date=d(t, -8), material=mid(1), qty=22, reason="Offcut",
             product=pid(1), avoidable="No",
             notes="Tail ends from six hats."),
        dict(date=d(t, -12), material=mid(21), qty=38,
             reason="Miscut", product=pid(10), avoidable="Yes",
             notes="Measured once. Expensive lesson in walnut."),
        dict(date=d(t, -15), material=mid(11), qty=0.42,
             reason="Failed pour", product=pid(6), avoidable="Yes",
             notes="Poured too hot, two jars with sink holes."),
        dict(date=d(t, -19), material=mid(15), qty=2,
             reason="Breakage in post", product=pid(6),
             avoidable="No", notes="Courier. Claim submitted."),
        dict(date=d(t, -24), material=mid(17), qty=45,
             reason="Normal trim waste", product=pid(8),
             avoidable="No"),
        dict(date=d(t, -30), material=mid(4), qty=36,
             reason="Mistake / frogged", product=pid(3),
             avoidable="Yes",
             notes="Frogged a whole fox body - yarn is reusable "
                   "but tired."),
        dict(date=d(t, -44), material=mid(12), qty=18,
             reason="Spoiled in storage", avoidable="Yes",
             notes="Left the cap off. Fragrance flashes off fast."),
        dict(date=d(t, -52), material=mid(22), qty=26,
             reason="Warped", product=pid(11), avoidable="No",
             notes="Board cupped after milling."),
        dict(date=d(t, -66), material=mid(3), qty=30,
             reason="Dye lot mismatch", product=pid(2),
             avoidable="No"),
        dict(date=d(t, -80), material=mid(9), qty=1,
             reason="Cracked", product=pid(5), avoidable="No",
             notes="Sea glass split while setting."),
    ]

    # ------------------------------------------------- development --
    s.development = [
        dict(name="Willow cable beanie", product=pid(1),
             stage="Recovered cost", started=d(t, -430),
             design_hours=6, sample_hours=5, photo_hours=3,
             sample_cost=18.60, other_cost=0.00, samples_sold=2,
             decided=d(t, -400),
             notes="Three samples before the crown sat right. Paid "
                   "for itself in the first season."),
        dict(name="Woodland fox amigurumi", product=pid(3),
             stage="Listed", started=d(t, -330), design_hours=11,
             sample_hours=9, photo_hours=4, sample_cost=14.20,
             other_cost=0.00, samples_sold=1, decided=d(t, -300),
             notes="Twenty four hours of development. The break "
                   "even column says how many foxes that is."),
        dict(name="Linen crossbody bag", product=pid(8),
             stage="Listed", started=d(t, -220), design_hours=9,
             sample_hours=12, photo_hours=4, sample_cost=42.80,
             other_cost=16.00, samples_sold=1, decided=d(t, -180),
             notes="Two toiles and a pattern draft."),
        dict(name="Beeswax wrap set", product=pid(14),
             stage="Sampling", started=d(t, -40), design_hours=3,
             sample_hours=4, photo_hours=0, sample_cost=22.40,
             other_cost=0.00, samples_sold=0,
             notes="Still deciding. Nothing listed, nothing "
                   "recovered."),
        dict(name="Ceramic wick holder", stage="Shelved",
             started=d(t, -150), design_hours=5, sample_hours=6,
             photo_hours=0, sample_cost=38.00, other_cost=0.00,
             samples_sold=0, decided=d(t, -120),
             notes="Shelved - kiln time costs more than the "
                   "holders would ever sell for. Written off on "
                   "purpose, and visible."),
        dict(name="Christmas ornament SVG bundle", stage="Testing",
             started=d(t, -25), design_hours=12, sample_hours=2,
             photo_hours=2, sample_cost=0.00, other_cost=24.00,
             samples_sold=0,
             notes="Digital, so the only cost is time and the "
                   "stock licence."),
    ]

    # ------------------------------------------------------ events --
    s.events = [
        dict(name="Asheville Spring Market", kind="Craft fair",
             date=d(t, -70), days=1, stall_fee=45.00, travel=18.00,
             accommodation=0.00, other_cost=22.00, prep_hours=9,
             stall_hours=8, target=350.00, footfall=900,
             notes="First fair of the year. Good weather, slow "
                   "crowd."),
        dict(name="River Arts Night Market", kind="Farmers market",
             date=d(t, -35), days=1, stall_fee=30.00, travel=12.00,
             accommodation=0.00, other_cost=8.00, prep_hours=5,
             stall_hours=6, target=240.00, footfall=520,
             notes="Evening market. Candles sell under lights."),
        dict(name="Holiday Pop-up at The Mill", kind="Pop-up shop",
             date=d(t, -8), days=1, stall_fee=60.00, travel=15.00,
             accommodation=0.00, other_cost=10.00, prep_hours=6,
             stall_hours=7, target=320.00, footfall=700),
        dict(name="Winter Makers Fair", kind="Christmas market",
             date=d(t, 24), days=2, stall_fee=120.00, travel=40.00,
             accommodation=95.00, other_cost=35.00, prep_hours=16,
             stall_hours=14, target=1400.00, footfall=2600,
             notes="Booked. The stock needed row says what has to "
                   "be made before then."),
    ]

    # ------------------------------------------ wholesale customers --
    s.wholesale_customers = [
        dict(name="The Mill Mercantile", kind="Independent shop",
             contact="Dana Pryor", email="dana@millmercantile.com",
             phone="828 555 0142", city="Asheville NC",
             terms="Net 30", discount=0.50, min_order=250.00,
             first_order=d(t, -240), stockist_since=d(t, -240),
             exclusive="No", active="Yes",
             notes="Reorders candles every six weeks."),
        dict(name="Blue Ridge Gift Co", kind="Gift shop",
             contact="Hal Winters", email="orders@brgift.co",
             phone="828 555 0198", city="Black Mountain NC",
             terms="50% deposit, 50% on despatch", discount=0.45,
             min_order=200.00, first_order=d(t, -150),
             stockist_since=d(t, -150), exclusive="No", active="Yes",
             notes="Takes 45% rather than 50 - small orders."),
        dict(name="Pinecone Boutique", kind="Boutique",
             contact="Rae Olsen", email="rae@pineconeboutique.com",
             phone="828 555 0177", city="Asheville NC",
             terms="Net 14", discount=0.50, min_order=250.00,
             first_order=d(t, -90), stockist_since=d(t, -90),
             exclusive="No", active="Yes"),
        dict(name="Folk Art Gallery", kind="Gallery",
             contact="Marianne Shaw", email="hello@folkartgallery.org",
             phone="828 555 0121", city="Asheville NC",
             terms="Consignment", discount=0.40, min_order=0.00,
             first_order=d(t, -300), stockist_since=d(t, -300),
             exclusive="No", active="No",
             notes="Consignment, and gone quiet. Chase or drop."),
        dict(name="Hearthside Subscription Box",
             kind="Subscription box", contact="Toby Lin",
             email="buying@hearthsidebox.com", phone="404 555 0160",
             city="Atlanta GA", terms="Payment with order",
             discount=0.55, min_order=600.00, first_order=d(t, -40),
             stockist_since=d(t, -40), exclusive="No", active="Yes",
             notes="55% off, 200 units at a time. Volume only "
                   "works if the maths works - check the margin."),
    ]

    s.wholesale_orders = [
        dict(date=d(t, -240), customer=wcid(1), po_ref="MM-1041",
             product=pid(6), qty=24, due=d(t, -226),
             despatched=d(t, -228), carriage=18.00,
             carriage_cost=16.40, status="Paid", paid="Yes"),
        dict(date=d(t, -150), customer=wcid(2), po_ref="BRG-220",
             product=pid(7), qty=36, due=d(t, -136),
             despatched=d(t, -138), carriage=14.00,
             carriage_cost=12.80, status="Paid", paid="Yes"),
        dict(date=d(t, -90), customer=wcid(3), po_ref="PB-0012",
             product=pid(1), qty=12, due=d(t, -76),
             despatched=d(t, -78), carriage=12.00,
             carriage_cost=11.20, status="Paid", paid="Yes",
             notes="Sold through in three weeks - and still the worst "
                   "order in the book. Half of 78 does not cover a "
                   "beanie. The margin column said so first."),
        dict(date=d(t, -62), customer=wcid(1), po_ref="MM-1108",
             product=pid(6), qty=24, due=d(t, -48),
             despatched=d(t, -50), carriage=18.00,
             carriage_cost=16.40, status="Paid", paid="Yes"),
        dict(date=d(t, -40), customer=wcid(5), po_ref="HSB-A417",
             product=pid(7), qty=200, price_input=7.60,
             due=d(t, -12), despatched=d(t, -14), carriage=48.00,
             carriage_cost=52.00, status="Paid", paid="Yes",
             notes="Price typed in by hand at 7.60 because the "
                   "55% discount off 19 would have been 8.55 and "
                   "they would not go above 7.60. Check the "
                   "margin column - this is the order to watch."),
        dict(date=d(t, -34), customer=wcid(3), po_ref="PB-0019",
             product=pid(5), qty=8, due=d(t, -20),
             despatched=d(t, -22), carriage=9.00, carriage_cost=8.40,
             status="Invoiced", paid="No",
             notes="Net 14 and now overdue. Chase it."),
        dict(date=d(t, -21), customer=wcid(2), po_ref="BRG-268",
             product=pid(6), qty=18, due=d(t, -7),
             despatched=d(t, -9), carriage=14.00,
             carriage_cost=13.60, status="Invoiced", paid="No"),
        dict(date=d(t, -12), customer=wcid(1), po_ref="MM-1155",
             product=pid(1), qty=10, due=d(t, 2), carriage=12.00,
             carriage_cost=11.20, status="In production", paid="No",
             notes="Ten beanies, and the last trade order of them at "
                   "this price. Candles wholesale, crochet does not."),
        dict(date=d(t, -6), customer=wcid(5), po_ref="HSB-A512",
             product=pid(7), qty=200, price_input=7.60,
             due=d(t, 22), carriage=48.00, carriage_cost=52.00,
             status="Confirmed", paid="No",
             notes="Second box order. Same price, same question."),
        dict(date=d(t, -2), customer=wcid(4), po_ref="FAG-77",
             product=pid(4), qty=6, due=d(t, 26), carriage=0.00,
             carriage_cost=0.00, status="Enquiry", paid="No",
             notes="Consignment enquiry - they pay when it sells, "
                   "if it sells."),
    ]

    # -------------------------------------------------------- sales --
    # Five months of orders: mostly Etsy, three fairs, a handful of
    # direct sales, one breakage and one gift.  Postage charged and
    # postage paid are both recorded, because the difference between
    # them is real money and most sellers never add it up.
    def sale(days: int, product: int, qty: int, channel: str,
             price: float, **kw: Any) -> Dict[str, Any]:
        row: Dict[str, Any] = dict(
            date=d(t, -days), product=pid(product), qty=qty,
            channel=channel, unit_price=price,
            status=kw.pop("status", "Delivered"))
        row.update(kw)
        return row

    fair1 = "Asheville Spring Market"
    fair2 = "River Arts Night Market"
    fair3 = "Holiday Pop-up at The Mill"

    s.sales = [
        sale(1, 8, 1, "Own website", 195.00, ship_charged=8.95,
             ship_paid=9.85, status="Awaiting payment",
             customer="D. Fairweather",
             notes="Invoice sent, not paid yet."),
        sale(2, 6, 1, "Etsy", 32.00, ship_charged=8.95, ship_paid=9.85,
             status="Paid", customer="R. Mendez"),
        sale(3, 12, 1, "Etsy", 8.50, status="Delivered",
             customer="S. Udall",
             notes="Instant download. No postage, no packing, no "
                   "hours after the first sale."),
        sale(4, 4, 1, "Etsy", 58.00, ship_charged=4.95, ship_paid=4.25,
             fees_input=8.12, customer="J. Whitcomb",
             notes="Fee typed in from the statement - compare it "
                   "with the estimate in the next column."),
        sale(5, 7, 2, "Etsy", 19.00, ship_charged=6.95, ship_paid=7.10,
             customer="A. Fenwick"),
        sale(6, 1, 1, "Etsy", 78.00, discount=7.80, promo="WELCOME10",
             ship_charged=5.95, ship_paid=5.40, customer="N. Brodie"),
        sale(8, 6, 2, "Craft fair", 32.00, event=fair3,
             customer="Walk-in"),
        sale(8, 7, 3, "Craft fair", 19.00, event=fair3,
             customer="Walk-in"),
        sale(8, 1, 1, "Craft fair", 78.00, event=fair3,
             customer="Walk-in"),
        sale(8, 5, 1, "Craft fair", 52.00, event=fair3,
             customer="Walk-in"),
        sale(9, 13, 1, "Etsy", 9.50, customer="M. Kaur"),
        sale(10, 10, 1, "Own website", 135.00, ship_charged=16.00,
             ship_paid=18.60, customer="Hearthside Interiors",
             notes="Flat rate postage. The courier cost 2.60 more "
                   "than the buyer paid."),
        sale(11, 7, 1, "Other", 0.00, status="Gift",
             customer="Blue Ridge Living",
             notes="Sample to a magazine. Costed, not charged - "
                   "which is exactly how it should appear."),
        sale(12, 2, 1, "Etsy", 75.00, ship_charged=6.95,
             ship_paid=7.10, customer="P. Lasseter"),
        sale(13, 5, 1, "Etsy", 52.00, ship_charged=4.95,
             ship_paid=4.25, customer="C. Ngata"),
        sale(15, 9, 1, "Own website", 148.00, ship_charged=8.95,
             ship_paid=9.85, customer="T. Alcorn"),
        sale(16, 12, 2, "Etsy", 8.50, customer="E. Rossiter"),
        sale(18, 3, 1, "Etsy", 78.00, ship_charged=6.95,
             ship_paid=7.10, customer="H. Delacroix"),
        sale(19, 6, 1, "Instagram / direct", 32.00, ship_charged=8.95,
             ship_paid=9.85, customer="L. Byrne",
             notes="A payment link. Cheapest fees you pay."),
        sale(21, 6, 1, "Etsy", 32.00, ship_charged=8.95,
             ship_paid=9.85, refund=40.95, status="Refunded",
             customer="G. Hollis",
             notes="Jar cracked in the post. Refunded in full, "
                   "postage as well - and the wax is gone too."),
        sale(23, 8, 1, "Etsy", 195.00, ship_charged=8.95,
             ship_paid=9.85, customer="V. Amadi"),
        sale(24, 1, 1, "Etsy", 78.00, ship_charged=5.95,
             ship_paid=5.40, fees_input=21.20, customer="B. Soriano",
             notes="Came through offsite ads, so the fee was 15% "
                   "on top. One order in five looks like this."),
        sale(26, 7, 2, "Own website", 19.00, ship_charged=6.95,
             ship_paid=7.10, customer="F. Okonkwo"),
        sale(28, 13, 1, "Etsy", 9.50, customer="I. Petrova"),
        sale(30, 5, 1, "Etsy", 52.00, ship_charged=4.95,
             ship_paid=4.25, customer="D. Mwangi"),
        sale(31, 2, 1, "Etsy", 75.00, ship_charged=6.95,
             ship_paid=7.10, customer="K. Lindqvist"),
        sale(33, 12, 1, "Etsy", 8.50, customer="O. Reyes"),
        sale(35, 7, 4, "Craft fair", 19.00, event=fair2,
             customer="Walk-in"),
        sale(35, 6, 2, "Craft fair", 32.00, event=fair2,
             customer="Walk-in"),
        sale(35, 1, 2, "Craft fair", 65.00, event=fair2,
             customer="Walk-in"),
        sale(38, 3, 1, "Etsy", 62.00, ship_charged=6.95,
             ship_paid=7.10, customer="W. Castellano"),
        sale(40, 9, 1, "Instagram / direct", 148.00, ship_charged=8.95,
             ship_paid=9.85, customer="K. Osei"),
        sale(42, 12, 2, "Etsy", 8.50, customer="R. Bannerjee"),
        sale(44, 4, 1, "Etsy", 52.00, ship_charged=4.95,
             ship_paid=4.25, customer="Y. Tanaka",
             notes="Sold at the old price, before silver moved."),
        sale(46, 10, 1, "Own website", 135.00, ship_charged=16.00,
             ship_paid=18.60, customer="M. Quintero"),
        sale(48, 6, 1, "Etsy", 32.00, ship_charged=8.95,
             ship_paid=9.85, customer="S. Ellery"),
        sale(52, 1, 1, "Etsy", 65.00, ship_charged=5.95,
             ship_paid=5.40, customer="J. Pembroke",
             notes="Before the autumn price rise."),
        sale(55, 13, 1, "Etsy", 6.50, customer="A. Laurent"),
        sale(58, 5, 1, "Etsy", 44.00, ship_charged=4.95,
             ship_paid=4.25, refund=8.00, status="Partly refunded",
             customer="N. Oyelaran",
             notes="Chain was shorter than listed. Eight dollars "
                   "back rather than a return."),
        sale(60, 2, 1, "Etsy", 75.00, ship_charged=6.95,
             ship_paid=7.10, customer="T. Halloran"),
        sale(63, 7, 1, "Etsy", 19.00, ship_charged=6.95,
             ship_paid=7.10, customer="E. Vasquez"),
        sale(66, 12, 1, "Etsy", 8.50, customer="L. Marchetti"),
        sale(70, 5, 2, "Craft fair", 44.00, event=fair1,
             customer="Walk-in"),
        sale(70, 2, 2, "Craft fair", 75.00, event=fair1,
             customer="Walk-in"),
        sale(70, 7, 3, "Craft fair", 19.00, event=fair1,
             customer="Walk-in"),
        sale(70, 1, 1, "Craft fair", 65.00, event=fair1,
             customer="Walk-in"),
        sale(74, 8, 1, "Etsy", 195.00, ship_charged=8.95,
             ship_paid=9.85, customer="P. Ashworth"),
        sale(78, 12, 3, "Etsy", 8.50, customer="Three buyers",
             notes="A pattern round-up blog linked it."),
        sale(82, 4, 1, "Etsy", 52.00, ship_charged=4.95,
             ship_paid=4.25, customer="R. Stenberg"),
        sale(88, 6, 1, "Etsy", 32.00, discount=6.40, promo="AUTUMN20",
             ship_charged=8.95, ship_paid=9.85, customer="J. Farrow",
             notes="Twenty percent off. Look at what was left."),
        sale(92, 3, 1, "Own website", 62.00, ship_charged=6.95,
             ship_paid=7.10, customer="U. Bergstrom"),
        sale(96, 9, 1, "Etsy", 128.00, ship_charged=8.95,
             ship_paid=9.85, customer="C. Devereux"),
        sale(104, 5, 1, "Etsy", 44.00, ship_charged=24.50,
             ship_paid=24.50, customer="A. Kowalski",
             notes="International. Customs form, twelve days, "
                   "and the fees are charged on the postage too."),
        sale(112, 1, 1, "Etsy", 65.00, ship_charged=5.95,
             ship_paid=5.40, customer="M. Thorne"),
        sale(120, 7, 2, "Etsy", 19.00, ship_charged=6.95,
             ship_paid=7.10, customer="S. Nakamura"),
        sale(130, 12, 1, "Etsy", 7.50, customer="B. Ferreira",
             notes="Pattern was 7.50 before the price test."),
        sale(140, 10, 1, "Own website", 118.00, ship_charged=16.00,
             ship_paid=18.60, customer="Cottonwood Interiors"),
        sale(150, 2, 1, "Etsy", 75.00, ship_charged=6.95,
             ship_paid=7.10, customer="G. Vuković"),
    ]

    return s
````

---

## 3. The costing registers

One module per data tab. Each returns a `SheetSpec` - a list of columns with their formulas - and the writer turns that into a formatted, validated, protected sheet. These are the chain that turns a pile of materials and a number of minutes into a price: the catalog, the materials and the bill of materials, labour, packaging, postage, overhead and the marketplace fee table.

### `handmade_pricing_spreadsheet/sheets/__init__.py`

*0 lines*

````python

````

### `handmade_pricing_spreadsheet/sheets/common.py`

*400 lines · 26 functions*

> Small building blocks shared by the register sheets.

| Defines | What it does |
|---------|--------------|
| `id_col(trigger, prefix, digits, width, header, note)` | Auto numbered ID column (blank until the trigger column is filled) |
| `month_col(date_key, key, header, width)` | `yyyy-mm` text column - the key every monthly roll-up groups on |
| `weekday_col(date_key, key, header, width)` |  |
| `week_col(date_key, key, header, width)` | Week number of the year |
| `alert_cf()` | Traffic light colouring for the 'Alert' / 'Check' columns |
| `status_cf()` | Colouring for the job, quote, invoice and stock status columns |
| `rate_cf(target_name)` | Green at or above the target, amber below it |
| `bar_cf(color)` |  |
| `scale_cf()` |  |
| `money_cf()` |  |
| `loss_cf()` | Red when a profit column goes negative |
| `dup_cf()` |  |
| `flag_yes_no_cf()` |  |
| `yesno(note)` |  |
| `score_validation(maximum)` |  |
| `blank(trigger, body)` | `=IF(trigger="","",body)` with the standard row placeholders |
| `lookup_product(target, key, missing)` | Pull a column of the Products catalog for the ID in `key` |
| `lookup_material(target, key, missing)` | Pull a column of the Materials tab for the material in `key` |
| `lookup_channel(target, key, missing)` | Pull a column of the Channels tab for the sales channel in `key` |
| `lookup(sheet, target, key_col, key, missing)` | Generic `INDEX/MATCH` against any register, blank row safe |
| `round_price(cell, rule)` | Apply the rounding rule chosen on Settings to a price `cell` |
| `rank_desc(key, guard, header, width, note)` | Hidden dense rank, 1 = the biggest value in `key` |
| `rank_asc(key, guard, header, width, note)` | Hidden dense rank, 1 = the smallest value in `key` |
| `min_number(value_range, cond, cap)` | Smallest positive number in `value_range` where `cond` holds |
| `craft_tail(trigger)` | Cost, suggested price, margin and hourly pay for a craft tab |
| `craft_check(trigger, needs)` | The shared 'is this row sane' column for a specialist tab |

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
    ``ClientPrefix`` or a quoted literal such as ``'"JOB-"'``.
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
               note="Automatic. Every monthly report groups on this.")


def weekday_col(date_key: str, key: str = "weekday", header: str = "Day",
                width: float = 9) -> Col:
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"ddd"))'))


def week_col(date_key: str, key: str = "week", header: str = "Week",
             width: float = 7) -> Col:
    """Week number of the year."""
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'WEEKNUM({{{date_key}}}{{r}},2))'))


def alert_cf() -> List[Dict[str, Any]]:
    """Traffic light colouring for the 'Alert' / 'Check' columns."""
    return [
        cf_text("OK", "good", "begins with"),
        cf_text("Overdue", "bad"),
        cf_text("Late", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Conflict", "bad"),
        cf_text("Clash", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Duplicate", "bad"),
        cf_text("Unpaid", "bad"),
        cf_text("Loss", "bad"),
        cf_text("Below", "bad"),
        cf_text("Action", "warn"),
        cf_text("Due", "warn"),
        cf_text("Check", "warn"),
        cf_text("Chase", "warn"),
        cf_text("Assign", "warn"),
        cf_text("Confirm", "warn"),
        cf_text("Order", "warn"),
        cf_text("Today", "info"),
        cf_text("Upcoming", "info"),
        cf_text("Tomorrow", "info"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    """Colouring for the job, quote, invoice and stock status columns."""
    return [
        cf_text("Completed", "good"),
        cf_text("Complete", "good"),
        cf_text("Done", "good"),
        cf_text("Paid", "good"),
        cf_text("Accepted", "good"),
        cf_text("Won", "good"),
        cf_text("Active", "good"),
        cf_text("In stock", "good"),
        cf_text("Resolved", "good"),
        cf_text("Signed", "good"),
        cf_text("Yes", "good"),
        cf_text("Booked", "info"),
        cf_text("Confirmed", "info"),
        cf_text("Scheduled", "info"),
        cf_text("In progress", "info"),
        cf_text("Sent", "info"),
        cf_text("Invoiced", "info"),
        cf_text("Quoted", "info"),
        cf_text("Prospect", "info"),
        cf_text("Draft", "warn"),
        cf_text("Enquiry", "warn"),
        cf_text("Pending", "warn"),
        cf_text("Part paid", "warn"),
        cf_text("Followed up", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Due soon", "warn"),
        cf_text("On hold", "warn"),
        cf_text("Unassigned", "warn"),
        cf_text("Overdue", "bad"),
        cf_text("Cancelled", "bad"),
        cf_text("Declined", "bad"),
        cf_text("Lost", "bad"),
        cf_text("No access", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("Written off", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Discontinued", "muted"),
        cf_text("Idea", "muted"),
        cf_text("Prototype", "warn"),
        cf_text("Made to order", "info"),
        cf_text("Seasonal", "info"),
        cf_text("Review", "warn"),
        cf_text("Raise price", "warn"),
        cf_text("Inactive", "muted"),
        cf_text("Archived", "muted"),
        cf_text("Not needed", "muted"),
    ]


def rate_cf(target_name: str = "TargetRetailMargin") -> List[Dict[str, Any]]:
    """Green at or above the target, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def bar_cf(color: str = "#B5704F") -> List[Dict[str, Any]]:
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


def dup_cf() -> List[Dict[str, Any]]:
    return [cf_text("Duplicate", "bad"), cf_text("OK", "good")]


def flag_yes_no_cf() -> List[Dict[str, Any]]:
    return [cf_text("No", "warn"), cf_text("Yes", "good")]


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


def lookup_product(target: str, key: str = "product",
                   missing: str = "Not found") -> str:
    """Pull a column of the Products catalog for the ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Products.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Products.id]],0)),"{missing}"))')


def lookup_material(target: str, key: str = "material",
                    missing: str = "Add to Materials") -> str:
    """Pull a column of the Materials tab for the material in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Materials.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Materials.id]],0)),"{missing}"))')


def lookup_channel(target: str, key: str = "channel",
                   missing: str = "Add to Channels") -> str:
    """Pull a column of the Channels tab for the sales channel in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Channels.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Channels.name]],0)),"{missing}"))')


def lookup(sheet: str, target: str, key_col: str, key: str,
           missing: str = "") -> str:
    """Generic `INDEX/MATCH` against any register, blank row safe."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[{sheet}.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[{sheet}.{key_col}]],0)),"{missing}"))')


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
    times into every row of the catalog.
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


# ---------------------------------------------------------------------------
# The six specialist calculators share a tail
# ---------------------------------------------------------------------------

def craft_tail(trigger: str = "name") -> List[Col]:
    """Cost, suggested price, margin and hourly pay for a craft tab.

    Every specialist calculator works out its own `mat_cost`,
    `labor_cost`, `other_cost` and `minutes`; from there the sums are
    identical, so they live here once.  The suggested price uses the
    default channel's fees and the target margin from Settings, which
    makes each tab usable on its own even for a piece that is not in
    the catalog yet.
    """
    return [
        Col("total_cost", "Cost to make", 12, "money",
            formula=blank(
                trigger,
                'ROUND({mat_cost}{r}+{labor_cost}{r}+{other_cost}{r},2)'),
            note="Materials + labour + extras. Packaging, postage and "
                 "overhead are added in the catalog, which knows which "
                 "profile and channel this product uses."),
        Col("fee_pct", "Fee rate", 9, "percent",
            formula=blank(
                trigger,
                'IFERROR(INDEX([[Channels.pct_total]],MATCH(DefaultChannel,'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 9, "money",
            formula=blank(
                trigger,
                'IFERROR(INDEX([[Channels.fixed_total]],MATCH(DefaultChannel,'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("denom", "Denominator", 9, "number",
            formula=blank(trigger, '1-TargetRetailMargin-{fee_pct}{r}'),
            hidden=True),
        Col("suggested_raw", "Unrounded", 10, "money",
            formula=blank(
                trigger,
                'IF({denom}{r}<=0,"",({total_cost}{r}+{fee_fixed}{r})/'
                '{denom}{r})'),
            hidden=True),
        Col("suggested", "Suggested price", 13, "money",
            formula=blank(
                trigger,
                'IF({suggested_raw}{r}="","Check the settings",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="(Cost + fixed fees) / (1 - target margin - fee rate), "
                 "priced for the default channel on Settings."),
        Col("price", "Price used", 11, "money",
            formula=blank(
                trigger,
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF(ISNUMBER({suggested}{r}),{suggested}{r},""))')),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                trigger,
                'IF({price}{r}="","",ROUND({price}{r}-{total_cost}{r}-'
                '({price}{r}*{fee_pct}{r}+{fee_fixed}{r}),2))'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                trigger,
                'IF(OR({price}{r}="",{price}{r}=0),"",{profit}{r}/'
                '{price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("hours", "Hours", 9, "number",
            formula=blank(trigger, 'ROUND({minutes}{r}/60,3)'),
            hidden=True),
        Col("wage_hour", "You earn an hour", 12, "money",
            formula=blank(
                trigger,
                'IF(OR({hours}{r}=0,{hours}{r}="",{profit}{r}=""),"",'
                'ROUND(({profit}{r}+{labor_cost}{r})/{hours}{r},2))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MinimumWage",
                 "fmt": "bad"}],
            note="Wage plus profit over the hours it takes. The honest "
                 "test of whether this piece is worth making."),
    ]


def craft_check(trigger: str = "name", needs: str = "mat_cost") -> Col:
    """The shared 'is this row sane' column for a specialist tab."""
    return Col(
        "check", "Check", 30, "status",
        formula=blank(
            trigger,
            'IF(AND({product}{r}<>"",ISERROR(MATCH({product}{r},'
            '[[Products.id]],0))),"Check - no such product ID",'
            f'IF({{{needs}}}{{r}}=0,"Add - no material cost yet",'
            'IF({minutes}{r}=0,"Add - how long does it take?",'
            'IF(AND({product}{r}<>"",COUNTIFS([[Material Usage.product]],'
            '{product}{r})>0),'
            '"Check - this product also has a bill of materials",'
            'IF(AND({product}{r}<>"",COUNTIFS([[Labor.product]],'
            '{product}{r})>0),'
            '"Check - this product also has labour lines",'
            'IF({wage_hour}{r}<MinimumWage,'
            '"Below - pays under minimum wage","OK")))))) '),
        cf=alert_cf())
````

### `handmade_pricing_spreadsheet/sheets/lists.py`

*476 lines · 1 function*

> `Lists` - every drop down in the workbook, plus the factor tables.
>
> Each list becomes a workbook level defined name, so a column can point at
> it with ``validate=list_from("CategoryList")``. Edit a list here and
> every drop down that uses it follows.
>
> The tables at the bottom are the engine behind the calculators: one
> converts a buying unit into the unit you actually use (a 100 g skein
> into grams of yarn), one turns "how fiddly is this" into extra minutes,
> one turns a turnaround promise into a rush surcharge, and one holds the
> target margin behind each pricing scenario. Change a number there and
> every price in the workbook re-prices itself.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Lists'), `LAST_COL` (9), `LISTS` (53 items), `UNIT_TABLE` (list(...)), `DIFFICULTY_TABLE` (5 items), `RUSH_TABLE` (5 items), `SCENARIO_TABLE` (5 items), `SIZE_TABLE` (6 items).

````python
"""`Lists` - every drop down in the workbook, plus the factor tables.

Each list becomes a workbook level defined name, so a column can point at
it with ``validate=list_from("CategoryList")``. Edit a list here and
every drop down that uses it follows.

The tables at the bottom are the engine behind the calculators: one
converts a buying unit into the unit you actually use (a 100 g skein
into grams of yarn), one turns "how fiddly is this" into extra minutes,
one turns a turnaround promise into a rush surcharge, and one holds the
target margin behind each pricing scenario. Change a number there and
every price in the workbook re-prices itself.
"""

from __future__ import annotations

from typing import List, Tuple

from ..config import (PRICING_SCENARIOS, PRODUCT_CATEGORIES,
                      PRODUCT_STATUSES, ROUNDING_RULES, SALES_CHANNELS,
                      UNIT_CONVERSIONS, Config, Palette)
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Lists"
LAST_COL = 9

#: name -> (header, [values]) - written one list per column block
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("YesNoNAList", ["Yes", "No", "N/A"]),
    ("CategoryList", list(PRODUCT_CATEGORIES)),
    ("StatusList", list(PRODUCT_STATUSES)),
    ("ChannelList", list(SALES_CHANNELS)),
    ("PricingMethodList", ["Cost plus margin", "Cost plus markup",
                           "Target price", "Keystone (x2)",
                           "Market matched", "Hourly rate only",
                           "Fixed price", "Made to order quote"]),
    ("UnitList", [u for u, _b, _f in UNIT_CONVERSIONS]),
    ("MaterialTypeList", ["Yarn & thread", "Fabric", "Beads & findings",
                          "Metal & wire", "Stones & crystals", "Wax",
                          "Fragrance & oils", "Wicks", "Vessels & jars",
                          "Timber & boards", "Finish & glue",
                          "Hardware & fixings", "Clay & glaze",
                          "Paper & card", "Ink & paint", "Leather",
                          "Elastic & notions", "Labels & tags",
                          "Packaging material", "Tools & consumables",
                          "Other supply"]),
    ("StockStatusList", ["In stock", "Low stock", "Out of stock",
                         "On order", "Discontinued", "Not tracked"]),
    ("SeasonList", ["All year", "Spring", "Summer", "Autumn", "Winter",
                    "Christmas", "Valentine's", "Mother's Day",
                    "Halloween", "Wedding season", "Back to school"]),
    ("DifficultyList", ["Very simple", "Simple", "Moderate", "Complex",
                        "Very complex"]),
    ("LaborStageList", ["Design", "Pattern & setup", "Cutting & prep",
                        "Making", "Assembly", "Sewing", "Finishing",
                        "Sanding & smoothing", "Painting & staining",
                        "Curing & drying", "Quality check", "Weaving in",
                        "Blocking", "Photography", "Listing & admin",
                        "Packing", "Post office run", "Customer messages"]),
    ("LaborWhoList", ["Me", "Partner", "Employee", "Contractor",
                      "Apprentice", "Studio assistant", "Outsourced"]),
    ("PerBasisList", ["Per item", "Per batch"]),
    ("PackProfileList", ["Budget", "Standard", "Gift", "Premium"]),
    ("PackTypeList", ["Mailer & box", "Tissue & wrap", "Filler",
                      "Tape & seal", "Label & sticker", "Thank you card",
                      "Care card", "Ribbon & twine", "Gift box",
                      "Polybag", "Insert & flyer", "Fragile protection"]),
    ("ShipServiceList", ["First class", "Priority", "Priority express",
                         "Ground", "Tracked 24", "Tracked 48",
                         "Signed for", "Courier", "International economy",
                         "International tracked", "Local delivery",
                         "Collection / pickup"]),
    ("ShipZoneList", ["Local", "Domestic", "Domestic remote", "Europe",
                      "North America", "Rest of world"]),
    ("ShipPayerList", ["Buyer pays", "Free shipping (I absorb it)",
                       "Free over threshold", "Flat rate",
                       "Calculated at checkout", "Collection only"]),
    ("OverheadTypeList", ["Studio rent", "Utilities", "Insurance",
                          "Shop subscription", "Website & domain",
                          "Software", "Accounting", "Bank & card fees",
                          "Marketing & ads", "Photography props",
                          "Tool replacement", "Equipment lease",
                          "Training & courses", "Memberships",
                          "Phone & internet", "Travel & fuel",
                          "Market stall fees", "Storage", "Other"]),
    ("FrequencyList", ["Monthly", "Quarterly", "Twice a year", "Yearly",
                       "One-off", "Per order"]),
    ("AllocationList", ["Per hour worked", "Per unit made", "Per order",
                        "Share of revenue", "Do not allocate"]),
    ("FeeTypeList", ["Percent of order", "Fixed per order",
                     "Fixed per item", "Listing fee",
                     "Percent + fixed", "Subscription", "None"]),
    ("PromoTypeList", ["Percent off", "Amount off", "Free shipping",
                       "Buy one get one", "Bundle price",
                       "Spend and save", "Sale price", "Coupon code",
                       "Early bird", "Loyalty reward"]),
    ("VariationTypeList", ["Size", "Colour", "Yarn choice", "Scent",
                           "Length", "Chain length", "Metal finish",
                           "Wood species", "Personalisation",
                           "Embroidery", "Gift wrap", "Rush order",
                           "Bundle add-on", "Custom request"]),
    ("BatchStatusList", ["Planned", "Materials ordered", "In progress",
                         "Curing", "Finishing", "Complete",
                         "Listed", "Sold out", "Abandoned"]),
    ("SaleStatusList", ["Paid", "Awaiting payment", "Made", "Shipped",
                        "Delivered", "Refunded", "Partly refunded",
                        "Cancelled", "Exchange", "Gift"]),
    ("WasteReasonList", ["Offcut", "Miscut", "Mistake / frogged",
                         "Failed pour", "Cracked", "Warped",
                         "Dye lot mismatch", "Spoiled in storage",
                         "Expired", "Sample / test piece",
                         "Customer return", "Breakage in post",
                         "Lost", "Normal trim waste"]),
    ("EventTypeList", ["Craft fair", "Farmers market", "Christmas market",
                       "Pop-up shop", "Trade show", "Wholesale market",
                       "Workshop", "Online market", "Charity event",
                       "School fete"]),
    ("CustomerTypeList", ["Independent shop", "Boutique", "Gallery",
                          "Gift shop", "Museum shop", "Online retailer",
                          "Subscription box", "Corporate gifting",
                          "Interior designer", "Market trader",
                          "Distributor"]),
    ("TermsList", ["Payment with order", "50% deposit, 50% on despatch",
                   "Net 7", "Net 14", "Net 30", "Net 60",
                   "Consignment", "Card on file"]),
    ("OrderStatusList", ["Enquiry", "Quoted", "Confirmed", "In production",
                         "Ready", "Despatched", "Invoiced", "Paid",
                         "Cancelled", "On hold"]),
    ("RoundingList", list(ROUNDING_RULES)),
    ("PriceReasonList", ["First price", "Material cost rose",
                         "Material cost fell", "Labour rate change",
                         "Fee change", "Competitor move", "Sale price",
                         "Sale ended", "Rebrand / repositioning",
                         "Rounding tidy-up", "Wholesale request",
                         "Testing a price"]),
    ("QualityList", ["Below mine", "Similar", "Above mine",
                     "Mass produced", "Designer / luxury", "Unknown"]),
    ("DevStageList", ["Idea", "Sketching", "Sampling", "Testing",
                      "Photography", "Listed", "Shelved",
                      "Recovered cost"]),
    ("RatingList", ["5 - Loved it", "4 - Happy", "3 - Fine",
                    "2 - Unhappy", "1 - Complaint"]),
    ("YarnWeightList", ["Lace", "Light fingering", "Fingering / 4 ply",
                        "Sport", "DK", "Worsted / Aran", "Bulky / chunky",
                        "Super bulky", "Jumbo", "Thread"]),
    ("HookSizeList", ["2.0 mm", "2.5 mm", "3.0 mm", "3.5 mm", "4.0 mm",
                      "4.5 mm", "5.0 mm", "5.5 mm", "6.0 mm", "6.5 mm",
                      "8.0 mm", "10.0 mm", "12.0 mm"]),
    ("MetalList", ["Sterling silver", "Silver plated", "Gold filled",
                   "9ct gold", "14ct gold", "Gold plated", "Rose gold",
                   "Brass", "Copper", "Stainless steel", "Titanium",
                   "Aluminium", "Mixed metal"]),
    ("StoneList", ["None", "Glass bead", "Crystal", "Freshwater pearl",
                   "Amethyst", "Rose quartz", "Turquoise", "Onyx",
                   "Moonstone", "Labradorite", "Cubic zirconia",
                   "Resin cabochon", "Sea glass", "Other gemstone"]),
    ("WaxList", ["Soy", "Coconut soy", "Coconut apricot", "Beeswax",
                 "Rapeseed", "Paraffin", "Palm", "Gel", "Wax melt blend"]),
    ("WickList", ["Cotton", "Wood", "Crackling wood", "Hemp", "Eco",
                  "Braided", "Multiple wicks"]),
    ("VesselList", ["Glass jar", "Amber jar", "Tin", "Ceramic vessel",
                    "Concrete vessel", "Teacup", "Tumbler", "Clamshell",
                    "Refill bag", "No vessel (melt)"]),
    ("FabricList", ["Cotton", "Organic cotton", "Linen", "Cotton jersey",
                    "Double gauze", "Denim", "Corduroy", "Wool",
                    "Fleece", "Canvas", "Viscose", "Silk", "Satin",
                    "Ripstop", "Oilcloth", "Deadstock"]),
    ("GarmentSizeList", ["Baby", "Toddler", "Child", "XS", "S", "M", "L",
                         "XL", "2XL", "3XL", "One size", "Made to measure"]),
    ("WoodList", ["Pine", "Oak", "Ash", "Beech", "Birch ply", "Walnut",
                  "Maple", "Cherry", "Sapele", "Reclaimed", "Plywood",
                  "MDF", "Live edge slab", "Offcut stock"]),
    ("FinishList", ["Oil", "Hard wax oil", "Danish oil", "Beeswax polish",
                    "Varnish", "Lacquer", "Stain", "Paint", "Milk paint",
                    "Resin", "Raw / unfinished"]),
    ("DigitalTypeList", ["Crochet pattern", "Knitting pattern",
                         "Sewing pattern", "SVG cut file", "Printable",
                         "Planner", "Clipart set", "Font",
                         "Photo pack", "Course / workshop",
                         "E-book", "Template"]),
    ("FileFormatList", ["PDF", "SVG", "PNG", "JPG", "DXF", "EPS", "AI",
                        "ZIP bundle", "Video", "Canva link"]),
    ("LicenceList", ["Personal use", "Small commercial",
                     "Extended commercial", "Limited run",
                     "No resale of files", "Credit required"]),
    ("BundleTypeList", ["Gift set", "Starter kit", "Collection",
                        "Seasonal set", "Multi-buy", "Mix and match",
                        "Subscription box", "Wholesale case"]),
    ("MonthList", ["January", "February", "March", "April", "May",
                   "June", "July", "August", "September", "October",
                   "November", "December"]),
]

#: buying unit -> the unit you actually count in, and the multiplier
UNIT_TABLE: List[Tuple[str, str, float]] = list(UNIT_CONVERSIONS)

#: how fiddly the piece is -> time factor, and the price factor some
#: makers like to add on top of the honest cost
DIFFICULTY_TABLE: List[Tuple[str, float, float, str]] = [
    ("Very simple", 0.85, 1.00,
     "Repeat work you could do watching television."),
    ("Simple", 1.00, 1.00,
     "A piece you have made many times - no surprises."),
    ("Moderate", 1.15, 1.05,
     "Some shaping, colour changes or fiddly assembly."),
    ("Complex", 1.35, 1.12,
     "Needs concentration, a pattern in front of you, few distractions."),
    ("Very complex", 1.60, 1.25,
     "Heirloom work, a skill few makers have - charge for it."),
]

#: how fast the customer wants it -> what that is worth
RUSH_TABLE: List[Tuple[str, float, str]] = [
    ("Standard", 0.00, "Your normal turnaround, no surcharge."),
    ("Priority", 0.15, "Jumps the queue by a few days."),
    ("Rush", 0.30, "Everything else waits - evenings and a weekend."),
    ("Same week", 0.50, "Dropping other work entirely. Price it properly."),
    ("Next day", 0.75, "Only if you genuinely can, and only at a premium."),
]

#: pricing scenario -> the margin it aims at
SCENARIO_TABLE: List[Tuple[str, float, str]] = [
    ("Budget market", PRICING_SCENARIOS["Budget market"],
     "Craft fair impulse buys and market stalls."),
    ("Conservative", PRICING_SCENARIOS["Conservative"],
     "Safe pricing while you build reviews."),
    ("Standard", PRICING_SCENARIOS["Standard"],
     "What a healthy handmade shop should be aiming for."),
    ("Premium", PRICING_SCENARIOS["Premium"],
     "Distinctive work, good photography, a following."),
    ("Luxury", PRICING_SCENARIOS["Luxury"],
     "Commissions, galleries and one-of-a-kind pieces."),
]

#: size -> how much more material and time it takes
SIZE_TABLE: List[Tuple[str, float, float]] = [
    ("Mini", 0.45, 0.50),
    ("Small", 0.75, 0.80),
    ("Standard", 1.00, 1.00),
    ("Large", 1.35, 1.25),
    ("Extra large", 1.75, 1.55),
    ("Custom size", 1.00, 1.00),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.grey_text)

    s.banner("Drop Down Lists & Pricing Factors",
             "Every list the workbook offers you, and the four tables the "
             "calculators multiply by. Edit an entry and the drop down "
             "that uses it changes everywhere.",
             LAST_COL)

    row = 4
    row = s.note(
        row,
        "Add new entries directly under the last one in a column - the "
        "lists are sized with room to grow, and blanks are ignored. Do "
        "not delete a whole column: a drop down pointing at a missing "
        "list stops working.",
        LAST_COL, height=30)
    row += 1

    # ------------------------------------------------------------ the lists
    row = s.section(row, "The drop downs", LAST_COL)
    per_row = 5
    spare = 8                       # blank rows kept inside every name
    top = row
    col = 0
    tallest = 0
    for index, (name, values) in enumerate(LISTS):
        if index and index % per_row == 0:
            top += tallest + 2
            col = 0
            tallest = 0
        s.write(top, col, name.replace("List", ""), fmts.chrome("header"))
        s.set_col(col, col, 24)
        for offset, value in enumerate(values):
            s.write(top + 1 + offset, col, value, fmts.input("text"))
        for offset in range(len(values), len(values) + spare):
            s.write(top + 1 + offset, col, None, fmts.input("text"))
        first = top + 2             # 1 based first value row
        last = top + 1 + len(values) + spare
        book.define_name(
            name,
            f"={quote_sheet(NAME)}!${chr(65 + col)}${first}:"
            f"${chr(65 + col)}${last}")
        tallest = max(tallest, len(values) + spare)
        col += 2
    row = top + tallest + 3

    sheet_ref = quote_sheet(NAME)

    # ----------------------------------------------------- unit conversions
    row = s.section(row, "Units: what you buy, what you use", LAST_COL)
    row = s.note(
        row,
        "Yarn arrives in 100 g skeins and gets used by the gram. Wax "
        "arrives by the kilo and gets poured by the gram. This table is "
        "how the Materials tab turns a purchase unit into a usable one, "
        "so a bill of materials can say '38 g' and still cost it "
        "correctly. 'Multiply by' is how many base units one of this "
        "unit contains.",
        LAST_COL, height=44)
    row += 1

    s.table_head(row, 0, ["Unit", "Base unit", "Multiply by"],
                 [18, 16, 16])
    first_unit = row + 1
    for index, (unit, base, factor) in enumerate(UNIT_TABLE):
        r = first_unit + index
        s.write(r, 0, unit, fmts.input("text"))
        s.write(r, 1, base, fmts.input("text"))
        s.write(r, 2, factor, fmts.input("number"))
    spare_unit = 6
    for index in range(len(UNIT_TABLE), len(UNIT_TABLE) + spare_unit):
        r = first_unit + index
        for c, style in enumerate(["text", "text", "number"]):
            s.write(r, c, None, fmts.input(style))
    last_unit = first_unit + len(UNIT_TABLE) + spare_unit

    book.define_name("UnitNames",
                     f"={sheet_ref}!$A${first_unit + 1}:$A${last_unit}")
    book.define_name("UnitBases",
                     f"={sheet_ref}!$B${first_unit + 1}:$B${last_unit}")
    book.define_name("UnitFactors",
                     f"={sheet_ref}!$C${first_unit + 1}:$C${last_unit}")
    row = last_unit + 2

    # ------------------------------------------------------ difficulty table
    row = s.section(row, "How fiddly is it, and what that does to the time",
                    LAST_COL)
    row = s.note(
        row,
        "The same shawl takes one maker four hours and another seven. "
        "Pick a line per product and the labour calculator multiplies "
        "your estimated minutes by the time factor. The price factor is "
        "separate and optional - it is the premium skilled work can "
        "carry on top of its honest cost.",
        LAST_COL, height=44)
    row += 1

    s.table_head(row, 0, ["Difficulty", "Time factor", "Price factor",
                          "What it means"], [20, 14, 14, 54])
    first_diff = row + 1
    for index, (label, time_f, price_f, meaning) in \
            enumerate(DIFFICULTY_TABLE):
        r = first_diff + index
        s.write(r, 0, label, fmts.input("text"))
        s.write(r, 1, time_f, fmts.input("number"))
        s.write(r, 2, price_f, fmts.input("number"))
        s.write(r, 3, meaning, fmts.input("text"))
    spare_diff = 4
    for index in range(len(DIFFICULTY_TABLE),
                       len(DIFFICULTY_TABLE) + spare_diff):
        r = first_diff + index
        for c, style in enumerate(["text", "number", "number", "text"]):
            s.write(r, c, None, fmts.input(style))
    last_diff = first_diff + len(DIFFICULTY_TABLE) + spare_diff

    book.define_name("DiffNames",
                     f"={sheet_ref}!$A${first_diff + 1}:$A${last_diff}")
    book.define_name("DiffTimeFactors",
                     f"={sheet_ref}!$B${first_diff + 1}:$B${last_diff}")
    book.define_name("DiffPriceFactors",
                     f"={sheet_ref}!$C${first_diff + 1}:$C${last_diff}")
    row = last_diff + 2

    # ------------------------------------------------------------ rush table
    row = s.section(row, "Turnaround: what a hurry is worth", LAST_COL)
    row = s.note(
        row,
        "A rush order costs you evenings, not materials. The custom "
        "order and variation tabs add the surcharge below to the price, "
        "so saying yes to 'can I have it by Friday?' stops being a "
        "favour you pay for yourself.",
        LAST_COL, height=38)
    row += 1

    s.table_head(row, 0, ["Turnaround", "Surcharge", "When you use it"],
                 [20, 14, 62])
    first_rush = row + 1
    for index, (label, pct, meaning) in enumerate(RUSH_TABLE):
        r = first_rush + index
        s.write(r, 0, label, fmts.input("text"))
        s.write(r, 1, pct, fmts.input("percent"))
        s.write(r, 2, meaning, fmts.input("text"))
    spare_rush = 3
    for index in range(len(RUSH_TABLE), len(RUSH_TABLE) + spare_rush):
        r = first_rush + index
        for c, style in enumerate(["text", "percent", "text"]):
            s.write(r, c, None, fmts.input(style))
    last_rush = first_rush + len(RUSH_TABLE) + spare_rush

    book.define_name("RushNames",
                     f"={sheet_ref}!$A${first_rush + 1}:$A${last_rush}")
    book.define_name("RushFactors",
                     f"={sheet_ref}!$B${first_rush + 1}:$B${last_rush}")
    row = last_rush + 2

    # -------------------------------------------------------- scenario table
    row = s.section(row, "Pricing scenarios and the margin each one aims at",
                    LAST_COL)
    row = s.note(
        row,
        "The Scenarios tab prices every product five ways at once so you "
        "can see what repositioning would be worth. These are the "
        "margins behind those five columns - change them to match the "
        "market you actually sell into.",
        LAST_COL, height=38)
    row += 1

    s.table_head(row, 0, ["Scenario", "Target margin", "Who prices this way"],
                 [20, 16, 62])
    first_scen = row + 1
    for index, (label, margin, meaning) in enumerate(SCENARIO_TABLE):
        r = first_scen + index
        s.write(r, 0, label, fmts.input("text"))
        s.write(r, 1, margin, fmts.input("percent"))
        s.write(r, 2, meaning, fmts.input("text"))
    last_scen = first_scen + len(SCENARIO_TABLE)

    book.define_name("ScenarioNames",
                     f"={sheet_ref}!$A${first_scen + 1}:$A${last_scen}")
    book.define_name("ScenarioMargins",
                     f"={sheet_ref}!$B${first_scen + 1}:$B${last_scen}")
    row = last_scen + 2

    # ------------------------------------------------------------ size table
    row = s.section(row, "Size: more material, more minutes", LAST_COL)
    row = s.note(
        row,
        "A baby blanket and a king size one are the same pattern and "
        "nowhere near the same price. The Variations tab multiplies the "
        "parent product's material cost and labour time by the factors "
        "below, so a size run prices itself.",
        LAST_COL, height=38)
    row += 1

    s.table_head(row, 0, ["Size", "Material factor", "Time factor"],
                 [20, 18, 16])
    first_size = row + 1
    for index, (label, mat_f, time_f) in enumerate(SIZE_TABLE):
        r = first_size + index
        s.write(r, 0, label, fmts.input("text"))
        s.write(r, 1, mat_f, fmts.input("number"))
        s.write(r, 2, time_f, fmts.input("number"))
    spare_size = 4
    for index in range(len(SIZE_TABLE), len(SIZE_TABLE) + spare_size):
        r = first_size + index
        for c, style in enumerate(["text", "number", "number"]):
            s.write(r, c, None, fmts.input(style))
    last_size = first_size + len(SIZE_TABLE) + spare_size

    book.define_name("SizeNames",
                     f"={sheet_ref}!$A${first_size + 1}:$A${last_size}")
    book.define_name("SizeMaterialFactors",
                     f"={sheet_ref}!$B${first_size + 1}:$B${last_size}")
    book.define_name("SizeTimeFactors",
                     f"={sheet_ref}!$C${first_size + 1}:$C${last_size}")
    row = last_size + 2

    row = s.note(
        row,
        "A word on the factors: they are a starting point from real "
        "handmade businesses, not a law. Track what a batch actually "
        "takes you on the Batch Production tab for a month, compare it "
        "with what you estimated, then nudge these numbers until the "
        "estimates match your bench.",
        LAST_COL, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/products.py`

*512 lines · 3 functions*

> `Products` - the catalog, and the price every other tab argues about.
>
> This is where it all lands. Materials from the bill of materials (or
> from one of the six specialist calculators), labour from the time you
> logged, packaging from the profile you picked, the postage you absorb,
> your share of the overhead, and the fees the channel takes.
>
>     Cost  = materials + labour + packaging + shipping + overhead
>     Price = (Cost + fixed fees) / (1 - target margin - fee rate)
>
> Type a price of your own in "Your price" whenever you want to - the
> workbook keeps showing you what the maths said, what you actually
> charge, and what the difference is doing to your margin.

| Defines | What it does |
|---------|--------------|
| `_craft(field)` | Sum one column of all six specialist calculators for this product |
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the catalog |

**Module constants:** `NAME` ('Products').

````python
"""`Products` - the catalog, and the price every other tab argues about.

This is where it all lands. Materials from the bill of materials (or
from one of the six specialist calculators), labour from the time you
logged, packaging from the profile you picked, the postage you absorb,
your share of the overhead, and the fees the channel takes.

    Cost  = materials + labour + packaging + shipping + overhead
    Price = (Cost + fixed fees) / (1 - target margin - fee rate)

Type a price of your own in "Your price" whenever you want to - the
workbook keeps showing you what the maths said, what you actually
charge, and what the difference is doing to your margin.
"""

from __future__ import annotations

from ..config import CRAFT_SHEETS, Config, Palette
from ..layout import Col, SheetSpec, max_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, rank_desc,
                     round_price, status_cf)

NAME = "Products"


def _craft(field: str) -> str:
    """Sum one column of all six specialist calculators for this product."""
    return "+".join(
        f'SUMIFS([[{sheet}.{field}]],[[{sheet}.product]],{{id}}{{r}})'
        for sheet in CRAFT_SHEETS)


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ProductPrefix", digits=4, width=10,
               note="Generated automatically. Every other tab refers "
                    "to a product by this ID, so filter the catalog "
                    "rather than re-sorting it."),
        Col("name", "Product", 30, "text",
            note="The name a buyer would recognise. 'Cream cable "
                 "beanie' beats 'Hat 3'."),
        Col("sku_input", "SKU override", 14, "text",
            note="Leave blank and the workbook builds one from the "
                 "category and the row number."),
        Col("category", "Category", 22, "text",
            validate=list_from("CategoryList")),
        Col("status", "Status", 14, "status", validate=list_from("StatusList"),
            cf=status_cf()),
        Col("season", "Season", 14, "text", validate=list_from("SeasonList")),
        Col("difficulty", "Difficulty", 14, "text",
            validate=list_from("DifficultyList"),
            note="Feeds the time factor on the Lists tab, which "
                 "stretches your estimated minutes."),
        Col("channel", "Main channel", 16, "text",
            validate=list_from("ChannelList"),
            note="Which fee schedule this product is priced for. "
                 "Change it and watch the suggested price move."),
        Col("pack_profile", "Packaging", 13, "text",
            validate=list_from("PackProfileList"),
            note="Picks up the cost and the weight of that profile "
                 "from the Packaging tab."),
        Col("ship_option", "Shipping option", 30, "text",
            validate=list_from("ShipOptions"),
            note="From the Shipping rate card. Only the part you "
                 "absorb becomes a cost of this product."),
        Col("weight", "Item weight (g)", 12, "number1",
            validate=number_check(0, 1000000),
            note="The finished piece, before packaging. The workbook "
                 "warns you when item plus packaging goes over the "
                 "weight band you chose."),
        Col("stock", "In stock", 9, "int",
            validate=number_check(0, 100000, decimal=False)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 1000000),
            note="What you actually charge. Leave it blank to use "
                 "the suggested price."),
        Col("target_input", "Target margin", 11, "percent",
            validate=number_check(0, 0.95),
            note="Per product override. Blank uses the target on the "
                 "Settings tab."),
        Col("made", "Made so far", 10, "int",
            validate=number_check(0, 1000000, decimal=False),
            note="Lifetime units made, for your own reference - the "
                 "sales figures come from the Sales tab."),
        Col("launched", "Listed on", 11, "date", validate=date_check()),
        Col("url", "Listing link", 22, "text"),
        Col("notes", "Notes", 26, "wrap"),
        # ----------------------------------------------- what it costs --
        Col("sku", "SKU", 13, "id",
            formula=blank(
                "name",
                'IF({sku_input}{r}<>"",{sku_input}{r},'
                'UPPER(LEFT(IF({category}{r}="","GEN",{category}{r}),3))'
                '&"-"&TEXT(ROW()-4,"000"))'),
            note="Automatic unless you override it. Printed on the "
                 "costing sheet and the order forms."),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Material Usage.line_cost]],'
                '[[Material Usage.product]],{id}{r})+' + _craft("mat_cost") +
                ',2)'),
            note="The bill of materials plus anything costed on a "
                 "specialist calculator tab. Use one or the other "
                 "for a given product, never both."),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Labor.cost]],[[Labor.product]],{id}{r})+' +
                _craft("labor_cost") + ',2)')),
        Col("minutes_raw", "Minutes logged", 11, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Labor.mins_item]],[[Labor.product]],'
                '{id}{r})+' + _craft("minutes") + ',1)'),
            hidden=True),
        Col("time_factor", "Difficulty factor", 10, "number",
            formula=blank(
                "name",
                'IFERROR(INDEX(DiffTimeFactors,MATCH({difficulty}{r},'
                'DiffNames,0)),1)'),
            hidden=True),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'ROUND({minutes_raw}{r}*{time_factor}{r},1)'),
            note="Your logged minutes, stretched by the difficulty "
                 "factor on the Lists tab."),
        Col("hours", "Hours", 9, "number",
            formula=blank("name", 'ROUND({minutes}{r}/60,3)'),
            hidden=True),
        Col("labour_adj", "Labour at difficulty", 11, "money",
            formula=blank(
                "name",
                'ROUND({labor_cost}{r}*{time_factor}{r},2)'),
            note="The labour cost after the difficulty factor. This "
                 "is the figure the price is built on."),
        Col("consumables", "Consumables", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Labor.consumables]],[[Labor.product]],'
                '{id}{r})+' + _craft("other_cost") + ',2)'),
            note="Thread, glue, sandpaper - charged per hour worked "
                 "- plus anything a specialist tab calls an extra."),
        Col("pack_cost", "Packaging", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({pack_profile}{r}="Budget",PackBudget,'
                'IF({pack_profile}{r}="Standard",PackStandard,'
                'IF({pack_profile}{r}="Gift",PackGift,'
                'IF({pack_profile}{r}="Premium",PackPremium,'
                'DefaultPackaging)))),2)')),
        Col("pack_weight", "Packaging weight", 11, "number1",
            formula=blank(
                "name",
                'IF({pack_profile}{r}="Budget",PackWeightBudget,'
                'IF({pack_profile}{r}="Standard",PackWeightStandard,'
                'IF({pack_profile}{r}="Gift",PackWeightGift,'
                'IF({pack_profile}{r}="Premium",PackWeightPremium,0))))'),
            hidden=True),
        Col("ship_pay", "Postage costs", 11, "money",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Shipping.total_cost]],'
                'MATCH({ship_option}{r},[[Shipping.label]],0)),'
                'DefaultShipping)'),
            hidden=True),
        Col("ship_charge", "Postage charged", 11, "money",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Shipping.charge]],'
                'MATCH({ship_option}{r},[[Shipping.label]],0)),0)'),
            hidden=True),
        Col("ship_limit", "Weight band", 11, "number1",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Shipping.max_weight]],'
                'MATCH({ship_option}{r},[[Shipping.label]],0)),"")'),
            hidden=True),
        Col("ship_cost", "Postage you absorb", 12, "money",
            formula=blank(
                "name",
                'ROUND(MAX(0,{ship_pay}{r}-{ship_charge}{r}),2)'),
            note="What the label costs less what the buyer pays. "
                 "Free shipping lands here in full."),
        Col("cost_base", "Cost before overhead", 13, "money",
            formula=blank(
                "name",
                'ROUND({mat_cost}{r}+{labour_adj}{r}+{consumables}{r}+'
                '{pack_cost}{r}+{ship_cost}{r},2)'),
            note="Materials + labour + consumables + packaging + the "
                 "postage you absorb."),
        # ------------------------------------------------- what it sells
        Col("fee_pct", "Fee rate", 10, "percent",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Channels.pct_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            note="Every percentage fee that channel charges, added "
                 "up. The 'f' in the pricing formula."),
        Col("fee_ship", "Fees on postage?", 10, "flag",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Channels.fees_on_shipping]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),"No")'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 10, "money",
            formula=blank(
                "name",
                'ROUND(IFERROR(INDEX([[Channels.fixed_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)+IF({fee_ship}{r}="Yes",'
                '{ship_charge}{r}*{fee_pct}{r},0),4)'),
            note="Listing and payment fees per order, plus the fee "
                 "the channel charges on the postage the buyer paid."),
        Col("oh_rate", "Overhead rate", 10, "percent",
            formula=blank(
                "name",
                'IF(OverheadMethod="Share of revenue",'
                'IF(RevenueTargetMonth<=0,0,OverheadMonthly/'
                'RevenueTargetMonth),0)'),
            hidden=True),
        Col("oh_fixed", "Overhead in cost", 10, "money",
            formula=blank(
                "name",
                'IF(OverheadMethod="Per hour worked",'
                'ROUND(OverheadPerHour*{hours}{r},2),'
                'IF(OR(OverheadMethod="Per unit made",'
                'OverheadMethod="Per order"),OverheadPerUnit,0))'),
            hidden=True),
        Col("target", "Target margin used", 12, "percent",
            formula=blank(
                "name",
                'IF({target_input}{r}<>"",{target_input}{r},'
                'TargetRetailMargin)')),
        Col("denom", "Pricing denominator", 10, "number",
            formula=blank(
                "name",
                '1-{target}{r}-{fee_pct}{r}-{oh_rate}{r}'),
            hidden=True),
        Col("suggested_raw", "Suggested, unrounded", 12, "money",
            formula=blank(
                "name",
                'IF({denom}{r}<=0,"",({cost_base}{r}+{oh_fixed}{r}+'
                '{fee_fixed}{r})/{denom}{r})'),
            hidden=True),
        Col("suggested", "Suggested price", 12, "money",
            formula=blank(
                "name",
                'IF({suggested_raw}{r}="","Check the settings",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="(Cost + fixed fees) / (1 - target margin - fee "
                 "rate), rounded by the rule on Settings. The one "
                 "number this whole workbook exists to produce."),
        Col("price", "Price used", 11, "money",
            formula=blank(
                "name",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF(ISNUMBER({suggested}{r}),{suggested}{r},""))'),
            note="Your price if you typed one, otherwise the "
                 "suggested price. Everything below uses this."),
        Col("overhead", "Overhead", 10, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({oh_fixed}{r}+{price}{r}*'
                '{oh_rate}{r},2))')),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({cost_base}{r}+'
                '{overhead}{r},2))'),
            note="C = materials + labour + packaging + shipping + "
                 "overhead. Fees are not a cost of making it, so "
                 "they are shown separately."),
        Col("fees", "Channel fees", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({price}{r}*{fee_pct}{r}+'
                '{fee_fixed}{r},2))')),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({price}{r}-{total_cost}{r}-'
                '{fees}{r},2))'),
            cf=loss_cf(),
            note="What is left after everything - including paying "
                 "yourself for the hours. A small number here is "
                 "still a wage; a negative one is a hobby."),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({price}{r}="",{price}{r}=0),"",{profit}{r}/'
                '{price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}],
            note="Profit / price. The share of every dollar you keep."),
        Col("markup", "Markup", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({total_cost}{r}="",{total_cost}{r}=0),"",'
                '{profit}{r}/{total_cost}{r})'),
            note="Profit / cost. The same product, a bigger looking "
                 "number - 100% markup is a 50% margin. Never "
                 "confuse the two again."),
        Col("profit_hour", "Profit an hour", 11, "money",
            formula=blank(
                "name",
                'IF(OR({hours}{r}="",{hours}{r}=0,{profit}{r}=""),"",'
                'ROUND({profit}{r}/{hours}{r},2))'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}],
            note="On top of the wage already inside the labour cost. "
                 "This is what the business earns over and above "
                 "paying you."),
        Col("wage_hour", "You earn an hour", 11, "money",
            formula=blank(
                "name",
                'IF(OR({hours}{r}="",{hours}{r}=0,{profit}{r}=""),"",'
                'ROUND(({profit}{r}+{labour_adj}{r})/{hours}{r},2))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MinimumWage",
                 "fmt": "bad"}],
            note="Wage plus profit, divided by the hours. If this is "
                 "under your local minimum wage, the price is wrong "
                 "- not your speed."),
        Col("min_price", "Never sell below", 12, "money",
            formula=blank(
                "name",
                'IF(1-{fee_pct}{r}-{oh_rate}{r}<=0,"",ROUND(({cost_base}{r}'
                '+{oh_fixed}{r}+{fee_fixed}{r})/(1-{fee_pct}{r}-'
                '{oh_rate}{r}),2))'),
            note="The price at which this product makes exactly "
                 "nothing. Sale prices below this are costing you "
                 "money to give away."),
        Col("keystone", "Keystone (x2)", 11, "money",
            formula=blank(
                "name",
                'IF({total_cost}{r}="","",ROUND({total_cost}{r}*2,2))'),
            hidden=True,
            note="The old craft fair rule of thumb, for comparison "
                 "only. It ignores fees entirely, which is why it "
                 "leaves makers short."),
        Col("wholesale", "Wholesale price", 12, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({price}{r}*'
                '(1-WholesaleDiscount),2))'),
            note="Your retail price less the wholesale discount on "
                 "Settings. The Wholesale Calculator shows whether "
                 "it survives."),
        Col("ws_profit", "Wholesale profit", 12, "money",
            formula=blank(
                "name",
                'IF({wholesale}{r}="","",ROUND({wholesale}{r}-'
                '{cost_base}{r}-{oh_fixed}{r},2))'),
            cf=loss_cf(),
            note="No marketplace fees on a wholesale order, but no "
                 "retail price either."),
        Col("ws_margin", "Wholesale margin", 11, "percent",
            formula=blank(
                "name",
                'IF(OR({wholesale}{r}="",{wholesale}{r}=0),"",'
                '{ws_profit}{r}/{wholesale}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetWholesaleMargin", "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=TargetWholesaleMargin", "fmt": "good"}]),
        Col("comp_price", "Competitors average", 12, "money",
            formula=blank(
                "name",
                'IFERROR(ROUND(AVERAGEIFS([[Competitors.price]],'
                '[[Competitors.product]],{id}{r}),2),"")')),
        Col("comp_gap", "Versus the market", 11, "percent",
            formula=blank(
                "name",
                'IF(OR({comp_price}{r}="",{comp_price}{r}=0,'
                '{price}{r}=""),"",{price}{r}/{comp_price}{r}-1)'),
            note="Above the market is not a problem if the work "
                 "justifies it - but you should know you are there."),
        Col("sold", "Units sold", 10, "int",
            formula=blank(
                "name",
                'SUMIFS([[Sales.qty]],[[Sales.product]],{id}{r})')),
        Col("revenue", "Revenue", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.revenue]],[[Sales.product]],'
                '{id}{r}),2)')),
        Col("actual_profit", "Actual profit", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.profit]],[[Sales.product]],'
                '{id}{r}),2)'),
            cf=loss_cf()),
        Col("actual_margin", "Actual margin", 11, "percent",
            formula=blank(
                "name",
                'IF(OR({revenue}{r}="",{revenue}{r}=0),"",'
                '{actual_profit}{r}/{revenue}{r})'),
            note="What really happened, discounts and refunds "
                 "included. Compare it with the margin you planned."),
        Col("last_sold", "Last sold", 11, "date",
            formula=blank(
                "name",
                'IF(' + max_if("[[Sales.date]]",
                               "[[Sales.product]]={id}{r}") +
                '=0,"",' + max_if("[[Sales.date]]",
                                  "[[Sales.product]]={id}{r}") + ')')),
        Col("last_priced", "Price changed", 11, "date",
            formula=blank(
                "name",
                'IF(' + max_if("[[Price History.date]]",
                               "[[Price History.product]]={id}{r}") +
                '=0,IF({launched}{r}="","",{launched}{r}),' +
                max_if("[[Price History.date]]",
                       "[[Price History.product]]={id}{r}") + ')'),
            hidden=True),
        Col("price_age", "Days since review", 11, "int",
            formula=blank(
                "name",
                'IF({last_priced}{r}="","",TodayDate-{last_priced}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": "=PriceReviewDays",
                 "fmt": "warn"}],
            hidden=True),
        rank_desc("profit", guard='{name}{r}=""', header="Profit rank"),
        rank_desc("margin", guard='{name}{r}=""', header="Margin rank"),
        rank_desc("revenue", guard='{name}{r}=""', header="Revenue rank"),
        rank_desc("profit_hour", guard='{name}{r}=""',
                  header="Hourly rank"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({cost_base}{r}=0,"Add - no costs entered yet",'
                'IF({price}{r}="","Check - no price yet",'
                'IF({profit}{r}<0,"Loss - every sale costs you money",'
                'IF(AND({wage_hour}{r}<>"",{wage_hour}{r}<MinimumWage),'
                '"Below - pays under minimum wage",'
                'IF({margin}{r}<MarginWarning,"Low margin - re-price it",'
                'IF(AND({ship_limit}{r}<>"",{weight}{r}<>"",'
                '{weight}{r}+{pack_weight}{r}>{ship_limit}{r}),'
                '"Check - too heavy for that postage band",'
                'IF(AND({price_age}{r}<>"",'
                '{price_age}{r}>PriceReviewDays),'
                '"Due - price has not been reviewed",'
                'IF(AND({comp_gap}{r}<>"",{comp_gap}{r}>0.5),'
                '"Check - well above the market","OK")))))))) '),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Product Catalog",
        subtitle="Every product, every cost, the price the maths "
                 "asks for and the margin you actually make.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).products,
        help_text="Give a product a name, a category, a packaging "
                  "profile and a channel. Cost it on the Materials "
                  "and Labor tabs (or a specialist tab) and the "
                  "price appears here.",
        extra_notes=[
            "The suggested price is (cost + fixed fees) / (1 - "
            "target margin - fee rate). Dividing by what is left "
            "is the step most pricing advice skips, and it is why "
            "'cost x 2' leaves you short on a marketplace.",
            "Margin and markup are not the same thing. 50% margin "
            "is 100% markup. Both are shown so nobody has to "
            "remember which one the Facebook group meant.",
            "'You earn an hour' is the number to look at when a "
            "price feels too high. If it is under minimum wage, "
            "the price is not too high - it is too low.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the catalog."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Catalog summary", last)
    sheet.table_head(row, 0, ["Products", "Active", "Average price",
                              "Average cost", "Average margin",
                              "Losing money", "Below target margin",
                              "Under minimum wage", "Average profit an hour"],
                     [12, 10, 14, 14, 15, 14, 19, 19, 21])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Products.name]],"<>")',
        '=COUNTIFS([[Products.status]],"Active")+'
        'COUNTIFS([[Products.status]],"Made to order")',
        '=IFERROR(ROUND(AVERAGE([[Products.price]]),2),"")',
        '=IFERROR(ROUND(AVERAGE([[Products.total_cost]]),2),"")',
        '=IFERROR(AVERAGE([[Products.margin]]),"")',
        '=COUNTIFS([[Products.profit]],"<0")',
        '=COUNTIFS([[Products.margin]],"<"&MarginWarning)',
        '=COUNTIFS([[Products.wage_hour]],"<"&MinimumWage,'
        '[[Products.name]],"<>")',
        '=IFERROR(ROUND(AVERAGE([[Products.profit_hour]]),2),"")',
    ], ["int", "int", "money", "money", "percent", "int", "int", "int",
        "money"])
````

### `handmade_pricing_spreadsheet/sheets/materials.py`

*239 lines · 2 functions*

> `Materials` - every supply you buy, costed down to the gram.
>
> A 100 g skein that cost $8.50 is not an $8.50 cost on a product that
> used 38 g of it. This tab turns what you *bought* into what one gram,
> one bead, one inch of it costs, and the bill of materials on the next
> tab multiplies that by what you actually used.
>
> Waste is deliberately NOT added here - it belongs on the usage line, so
> nothing ever gets counted twice. The "with waste" column is a reference
> figure only.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the material list |

**Module constants:** `NAME` ('Materials').

````python
"""`Materials` - every supply you buy, costed down to the gram.

A 100 g skein that cost $8.50 is not an $8.50 cost on a product that
used 38 g of it. This tab turns what you *bought* into what one gram,
one bead, one inch of it costs, and the bill of materials on the next
tab multiplies that by what you actually used.

Waste is deliberately NOT added here - it belongs on the usage line, so
nothing ever gets counted twice. The "with waste" column is a reference
figure only.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, rank_desc, status_cf

NAME = "Materials"

_USE = "[[Material Usage.material]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "MaterialPrefix", digits=4, width=10,
               note="Generated automatically. The bill of materials "
                    "picks from these IDs, so filter the list rather "
                    "than re-sorting it."),
        Col("name", "Material", 26, "text",
            note="What you call it when you reorder. Be specific - "
                 "'Cascade 220 worsted, cream' beats 'wool'."),
        Col("type", "Type", 18, "text",
            validate=list_from("MaterialTypeList")),
        Col("detail", "Colour / size / detail", 20, "text"),
        Col("supplier", "Supplier", 18, "text"),
        Col("code", "Supplier code", 14, "text",
            note="Their SKU. Saves ten minutes every time you reorder."),
        Col("pack_qty", "Quantity in a pack", 12, "number1",
            validate=number_check(0, 1000000),
            note="How much one purchase contains: 100 (g in a skein), "
                 "144 (beads in a gross), 2.4 (metres on a board)."),
        Col("unit", "Unit", 11, "text", validate=list_from("UnitList"),
            note="The unit that quantity is measured in. The unit "
                 "table on the Lists tab converts it to a base unit."),
        Col("pack_cost", "Pack price", 11, "money",
            validate=number_check(0, 1000000),
            note="What the supplier charges for one pack, before any "
                 "trade discount and before delivery."),
        Col("pack_ship", "Delivery per pack", 12, "money",
            validate=number_check(0, 100000),
            note="Postage, divided by how many packs were on the "
                 "order. It is part of what the material costs you."),
        Col("discount", "Trade discount", 11, "percent",
            validate=number_check(0, 1)),
        Col("waste", "Waste allowance", 11, "percent",
            validate=number_check(0, 1),
            note="Leave blank to use the default on Settings. This is "
                 "the share that becomes offcuts and mistakes - it is "
                 "applied on the usage line, never here."),
        Col("stock", "Stock on hand", 11, "number1",
            validate=number_check(0, 10000000),
            note="In the unit above. Count the shelf now and again."),
        Col("reorder", "Reorder at", 10, "number1",
            validate=number_check(0, 10000000)),
        Col("last_cost", "Previous pack price", 12, "money",
            validate=number_check(0, 1000000),
            note="The price before the last rise. The workbook shows "
                 "you how far it has moved."),
        Col("priced_on", "Price checked", 11, "date", validate=date_check(),
            note="When you last confirmed this price. Material costs "
                 "drift upwards quietly."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("net_pack", "Pack cost to you", 12, "money",
            formula=blank(
                "name",
                'IF({pack_cost}{r}="","",ROUND({pack_cost}{r}*'
                '(1-IF({discount}{r}="",0,{discount}{r}))+'
                'IF({pack_ship}{r}="",0,{pack_ship}{r}),2))'),
            note="Pack price less the trade discount, plus delivery."),
        Col("unit_cost", "Cost a unit", 11, "money4",
            formula=blank(
                "name",
                'IF(OR({net_pack}{r}="",{pack_qty}{r}="",'
                '{pack_qty}{r}=0),"",ROUND({net_pack}{r}/'
                '{pack_qty}{r},4))'),
            note="The number everything else multiplies: what one "
                 "gram, bead or inch costs you."),
        Col("base_unit", "Counted in", 10, "text",
            formula=blank(
                "name",
                'IFERROR(INDEX(UnitBases,MATCH({unit}{r},UnitNames,0)),"")'),
            hidden=True),
        Col("base_factor", "Unit factor", 9, "number",
            formula=blank(
                "name",
                'IFERROR(INDEX(UnitFactors,MATCH({unit}{r},UnitNames,0)),"")'),
            hidden=True),
        Col("base_cost", "Cost a base unit", 13, "money4",
            formula=blank(
                "name",
                'IF(OR({unit_cost}{r}="",{base_factor}{r}="",'
                '{base_factor}{r}=0),"",ROUND({unit_cost}{r}/'
                '{base_factor}{r},4))'),
            note="The same cost expressed per gram, per millilitre or "
                 "per centimetre, so two suppliers selling different "
                 "pack sizes can be compared honestly."),
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "name",
                'IF({waste}{r}="",DefaultWaste,{waste}{r})'),
            hidden=True),
        Col("cost_waste", "Cost a unit with waste", 13, "money4",
            formula=blank(
                "name",
                'IF({unit_cost}{r}="","",ROUND({unit_cost}{r}*'
                '(1+{waste_rate}{r}),4))'),
            note="Reference only. The waste is added on the bill of "
                 "materials line, so it is not double counted here."),
        Col("stock_value", "Stock value", 11, "money",
            formula=blank(
                "name",
                'IF(OR({stock}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({stock}{r}*{unit_cost}{r},2))'),
            note="Add this column up. Most makers are holding far more "
                 "money on the shelf than they think."),
        Col("lines", "Used in", 9, "int",
            formula=blank("name", f'COUNTIFS({_USE},{{id}}{{r}})'),
            note="How many product recipes call for this material."),
        Col("qty_used", "Quantity committed", 12, "number1",
            formula=blank(
                "name",
                f'SUMIFS([[Material Usage.qty_eff]],{_USE},{{id}}{{r}})'),
            note="One of each product that uses it would consume this "
                 "much, waste included."),
        Col("cost_used", "Cost committed", 12, "money",
            formula=blank(
                "name",
                f'ROUND(SUMIFS([[Material Usage.line_cost]],{_USE},'
                f'{{id}}{{r}}),2)')),
        Col("cover", "Stock covers", 9, "int",
            formula=blank(
                "name",
                'IF(OR({stock}{r}="",{qty_used}{r}="",'
                '{qty_used}{r}<=0),"",ROUNDDOWN({stock}{r}/'
                '{qty_used}{r},0))'),
            note="How many complete sets of the products that use it "
                 "you could make from the stock on hand."),
        Col("change", "Price move", 10, "percent",
            formula=blank(
                "name",
                'IF(OR({last_cost}{r}="",{last_cost}{r}=0,'
                '{pack_cost}{r}=""),"",({pack_cost}{r}-{last_cost}{r})/'
                '{last_cost}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.075,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "good"}],
            note="A material up 10% is a product that quietly stopped "
                 "making money. Re-price it."),
        Col("status", "Stock", 12, "status",
            formula=blank(
                "name",
                'IF({stock}{r}="","Not tracked",'
                'IF({stock}{r}<=0,"Out of stock",'
                'IF(AND({reorder}{r}<>"",{stock}{r}<={reorder}{r}),'
                '"Low stock","In stock")))'),
            cf=status_cf()),
        rank_desc("cost_used", guard='{name}{r}=""',
                  header="Spend rank",
                  note="1 = the material you spend the most on. The "
                       "Reports tab reads this."),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({pack_cost}{r}="","Add - no price yet",'
                'IF(OR({pack_qty}{r}="",{pack_qty}{r}=0),'
                '"Add - how much is in a pack?",'
                'IF({status}{r}="Out of stock","Order - none left",'
                'IF({status}{r}="Low stock","Order - running low",'
                'IF(AND({change}{r}<>"",{change}{r}>0.075),'
                '"Check - price jumped, re-price the products",'
                'IF(AND({priced_on}{r}<>"",TodayDate-{priced_on}{r}>365),'
                '"Check - price is over a year old","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Materials & Supplies",
        subtitle="What you buy, and what one unit of it actually costs "
                 "once discount and delivery are in.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).materials,
        help_text="Enter what a pack costs and how much is in it. The "
                  "workbook works out the cost per unit and every "
                  "product that uses it re-prices itself.",
        extra_notes=[
            "Delivery is part of the cost. A $6 pack with $4 postage "
            "is a $10 pack, and pretending otherwise is how makers "
            "end up working for nothing.",
            "Keep the previous price in its column when you update "
            "one. The price move column is an early warning that a "
            "product needs re-pricing.",
            "Waste belongs on the bill of materials line, not here - "
            "that way a material used carefully in one product and "
            "wastefully in another is costed correctly in both.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the material list."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1        # 0 based row under the data
    last = spec_.last_col_index
    row = sheet.section(row, "Materials summary", last)
    sheet.table_head(row, 0, ["Materials listed", "Priced", "Missing a price",
                              "Value on the shelf", "Low or out of stock",
                              "Prices up since last check",
                              "Not checked in a year"],
                     [17, 12, 16, 18, 19, 22, 20])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Materials.name]],"<>")',
        '=COUNTIFS([[Materials.name]],"<>",[[Materials.pack_cost]],"<>")',
        '=COUNTIFS([[Materials.name]],"<>",[[Materials.pack_cost]],"")',
        '=SUM([[Materials.stock_value]])',
        '=COUNTIFS([[Materials.status]],"Low stock")+'
        'COUNTIFS([[Materials.status]],"Out of stock")',
        '=COUNTIFS([[Materials.change]],">0")',
        '=COUNTIFS([[Materials.name]],"<>",[[Materials.priced_on]],'
        '"<"&TodayDate-365)',
    ], ["int", "int", "int", "money", "int", "int", "int"])
````

### `handmade_pricing_spreadsheet/sheets/material_usage.py`

*163 lines · 2 functions*

> `Material Usage` - the bill of materials: what goes into what.
>
> One row per material in a product. 38 g of cream worsted, two wooden
> buttons, 40 cm of ribbon. Each line is costed from the Materials tab,
> the waste allowance is added here (and only here), and the product
> catalog adds the lines up.
>
> If a product is costed on one of the six specialist calculator tabs
> instead, leave it out of this tab - use the bill of materials *or* a
> specialist sheet, not both, or the material cost is counted twice.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the bill of materials |

**Module constants:** `NAME` ('Material Usage').

````python
"""`Material Usage` - the bill of materials: what goes into what.

One row per material in a product. 38 g of cream worsted, two wooden
buttons, 40 cm of ribbon. Each line is costed from the Materials tab,
the waste allowance is added here (and only here), and the product
catalog adds the lines up.

If a product is costed on one of the six specialist calculator tabs
instead, leave it out of this tab - use the bill of materials *or* a
specialist sheet, not both, or the material cost is counted twice.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, lookup_material, lookup_product

NAME = "Material Usage"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs"),
            note="The product ID from the Products tab. Type several "
                 "rows with the same ID - one per material."),
        Col("product_name", "Product name", 26, "text",
            formula=lookup_product("name"),
            note="Automatic, so you can see at a glance that the ID "
                 "you typed is the product you meant."),
        Col("material", "Material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("material_name", "Material name", 26, "text",
            formula=lookup_material("name")),
        Col("unit", "Unit", 9, "text", formula=lookup_material("unit")),
        Col("qty", "Quantity used", 12, "number1",
            validate=number_check(0, 1000000),
            note="In the unit shown. Weigh the yarn left over and "
                 "subtract - guessing is what makes a price wrong."),
        Col("waste", "Waste %", 9, "percent",
            validate=number_check(0, 1),
            note="Blank uses the material's own allowance, or the "
                 "default on Settings. Fabric cut on the bias wastes "
                 "far more than yarn does."),
        Col("optional", "Optional?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Yes for an upgrade or an add-on that not every "
                 "piece gets. Optional lines are still costed - the "
                 "flag is there so you can see what is core."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "material",
                'IF({waste}{r}<>"",{waste}{r},'
                'IFERROR(INDEX([[Materials.waste_rate]],'
                'MATCH({material}{r},[[Materials.id]],0)),DefaultWaste))'),
            hidden=True),
        Col("qty_eff", "Quantity with waste", 13, "number1",
            formula=blank(
                "material",
                'IF({qty}{r}="","",ROUND({qty}{r}*'
                '(1+{waste_rate}{r}),4))'),
            note="What you actually have to buy to make one."),
        Col("unit_cost", "Cost a unit", 11, "money4",
            formula=lookup_material("unit_cost", missing="")),
        Col("line_cost", "Line cost", 11, "money4",
            formula=blank(
                "material",
                'IF(OR({qty_eff}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({qty_eff}{r}*{unit_cost}{r},4))'),
            note="Quantity, plus waste, times the cost a unit. This "
                 "is what the catalog adds up."),
        Col("product_total", "Materials in this product", 14, "money",
            formula=blank(
                "product",
                'ROUND(SUMIFS([[Material Usage.line_cost]],'
                '[[Material Usage.product]],{product}{r}),2)'),
            hidden=True),
        Col("share", "Share of the product", 11, "percent",
            formula=blank(
                "product",
                'IF(OR({line_cost}{r}="",{product_total}{r}="",'
                '{product_total}{r}=0),"",{line_cost}{r}/'
                '{product_total}{r})'),
            note="Which material is driving the cost. Usually one "
                 "line is most of it - that is the one to shop "
                 "around for."),
        Col("seq", "Line", 7, "int",
            formula=blank(
                "product",
                'COUNTIFS({product_first}:{product}{r},{product}{r})'),
            hidden=True),
        Col("key", "Key", 14, "text",
            formula=blank("product", '{product}{r}&"|"&{seq}{r}'),
            hidden=True,
            note="Used by the printable costing sheet to pull the "
                 "lines of one product in order."),
        Col("check", "Check", 24, "status",
            formula=(
                '=IF(AND({product}{r}="",{material}{r}=""),"",'
                'IF({product}{r}="","Add - which product?",'
                'IF({material}{r}="","Add - which material?",'
                'IF({product_name}{r}="Not found",'
                '"Check - no such product ID",'
                'IF({material_name}{r}="Add to Materials",'
                '"Check - no such material ID",'
                'IF(OR({qty}{r}="",{qty}{r}=0),"Add - how much?",'
                'IF({unit_cost}{r}="",'
                '"Check - that material has no price","OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Bill of Materials",
        subtitle="What goes into each product, how much of it, and what "
                 "that costs.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).material_usage,
        help_text="One row per material per product. The catalog adds "
                  "the lines up by product ID, so the order of the "
                  "rows does not matter.",
        extra_notes=[
            "Weigh, do not guess. Make one, weigh what you used, and "
            "the price of everything you make from then on is built "
            "on a real number.",
            "Use this tab OR a specialist calculator tab for a given "
            "product, never both - the catalog adds both together, "
            "which is the one way to double count in this workbook.",
            "The share column tells you where the money goes. If one "
            "material is 70% of the cost, a 10% better price on it "
            "is worth more than every other saving combined.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the bill of materials."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Bill of materials summary", last)
    sheet.table_head(row, 0, ["Lines entered", "Products covered",
                              "Materials used", "Total material cost",
                              "Average per line", "Lines needing a look"],
                     [14, 17, 16, 19, 16, 20])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Material Usage.product]],"<>")',
        '=COUNTIFS([[Material Usage.seq]],1)',
        '=COUNTIF([[Materials.lines]],">0")',
        '=ROUND(SUM([[Material Usage.line_cost]]),2)',
        '=IFERROR(ROUND(SUM([[Material Usage.line_cost]])/'
        'COUNTIF([[Material Usage.product]],"<>"),2),"")',
        '=COUNTIFS([[Material Usage.check]],"Check*")+'
        'COUNTIFS([[Material Usage.check]],"Add*")',
    ], ["int", "int", "int", "money", "money", "int"])
````

### `handmade_pricing_spreadsheet/sheets/labor.py`

*196 lines · 2 functions*

> `Labor` - your hours, stage by stage, priced like the work they are.
>
> One row per stage per product: 35 minutes cutting, 240 minutes making,
> 20 minutes finishing, 10 minutes packing. The cost of a stage is
> minutes / 60 x the hourly rate of whoever does it - the formula every
> pricing guide quotes and almost no maker actually applies.
>
> Stages marked "per batch" are divided by the batch size, so the hour
> you spend setting up a run of twenty is twenty three-minute shares,
> not twenty hours.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the labour lines |

**Module constants:** `NAME` ('Labor').

````python
"""`Labor` - your hours, stage by stage, priced like the work they are.

One row per stage per product: 35 minutes cutting, 240 minutes making,
20 minutes finishing, 10 minutes packing. The cost of a stage is
minutes / 60 x the hourly rate of whoever does it - the formula every
pricing guide quotes and almost no maker actually applies.

Stages marked "per batch" are divided by the batch size, so the hour
you spend setting up a run of twenty is twenty three-minute shares,
not twenty hours.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, lookup_product

NAME = "Labor"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs"),
            note="The product ID. Add as many stages as the piece "
                 "really has - the total is what the catalog uses."),
        Col("product_name", "Product name", 26, "text",
            formula=lookup_product("name")),
        Col("stage", "Stage", 20, "text",
            validate=list_from("LaborStageList"),
            note="Making is never the whole story. Photography, "
                 "listing and packing are work too, and they are "
                 "where small items lose money."),
        Col("who", "Who does it", 14, "text",
            validate=list_from("LaborWhoList"),
            note="Sets the hourly rate: you, an employee or a "
                 "contractor. Override it in the next column if this "
                 "particular job pays differently."),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000),
            note="Leave blank to use the rate on the Settings tab for "
                 "whoever is doing it."),
        Col("minutes", "Minutes", 10, "number1",
            validate=number_check(0, 100000),
            note="Time it once with a timer, not from memory. Makers "
                 "underestimate their own hours by about a third."),
        Col("basis", "Per item or batch", 13, "text",
            validate=list_from("PerBasisList"),
            note="'Per batch' for anything you do once for a whole "
                 "run - setting up, pouring a pot of wax, a post "
                 "office trip."),
        Col("batch_qty", "Items in that batch", 12, "int",
            validate=number_check(1, 100000, decimal=False),
            note="Only used when the stage is per batch. Blank uses "
                 "the default batch size on Settings."),
        Col("skilled", "Skilled work?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Flag the stages that took years to learn. The "
                 "reports show how much of your making time is "
                 "skilled - that is the part worth protecting."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "product",
                'IF({rate_input}{r}<>"",{rate_input}{r},'
                'IF({who}{r}="Employee",ROUND(EmployeeRate*'
                '(1+PayrollBurden),2),'
                'IF(OR({who}{r}="Contractor",{who}{r}="Outsourced"),'
                'ROUND(ContractorRate*(1+PayrollBurden),2),'
                'LaborCostHour)))'),
            note="Your own hours carry the on-costs from Settings "
                 "too - holiday and tax do not stop existing because "
                 "you are self employed."),
        Col("hours", "Hours", 9, "number",
            formula=blank(
                "product",
                'IF({minutes}{r}="","",ROUND({minutes}{r}/60,3))'),
            note="Minutes / 60. The textbook formula."),
        Col("batch_size", "Batch used", 10, "int",
            formula=blank(
                "product",
                'IF({basis}{r}="Per batch",IF(OR({batch_qty}{r}="",'
                '{batch_qty}{r}=0),DefaultBatchSize,{batch_qty}{r}),1)'),
            hidden=True),
        Col("stage_cost", "Cost of the stage", 12, "money",
            formula=blank(
                "product",
                'IF(OR({hours}{r}="",{rate}{r}=""),"",'
                'ROUND({hours}{r}*{rate}{r},2))'),
            note="Hours x rate, before it is shared over a batch."),
        Col("cost", "Cost per item", 11, "money",
            formula=blank(
                "product",
                'IF({stage_cost}{r}="","",ROUND({stage_cost}{r}/'
                'MAX(1,{batch_size}{r}),2))'),
            note="What this stage adds to one finished piece. This is "
                 "the column the catalog sums."),
        Col("mins_item", "Minutes per item", 12, "number1",
            formula=blank(
                "product",
                'IF({minutes}{r}="","",ROUND({minutes}{r}/'
                'MAX(1,{batch_size}{r}),2))')),
        Col("consumables", "Consumables", 11, "money",
            formula=blank(
                "product",
                'IF({hours}{r}="","",ROUND({hours}{r}/'
                'MAX(1,{batch_size}{r})*SuppliesPerHour,2))'),
            note="Thread, glue, sandpaper, the sponge you ruin. "
                 "Charged per hour worked from the Settings rate."),
        Col("product_total", "Labour in this product", 13, "money",
            formula=blank(
                "product",
                'ROUND(SUMIFS([[Labor.cost]],[[Labor.product]],'
                '{product}{r}),2)'),
            hidden=True),
        Col("share", "Share of the labour", 11, "percent",
            formula=blank(
                "product",
                'IF(OR({cost}{r}="",{product_total}{r}="",'
                '{product_total}{r}=0),"",{cost}{r}/'
                '{product_total}{r})')),
        Col("seq", "Line", 7, "int",
            formula=blank(
                "product",
                'COUNTIFS({product_first}:{product}{r},{product}{r})'),
            hidden=True),
        Col("key", "Key", 14, "text",
            formula=blank("product", '{product}{r}&"|"&{seq}{r}'),
            hidden=True),
        Col("check", "Check", 24, "status",
            formula=(
                '=IF({product}{r}="","",'
                'IF({product_name}{r}="Not found",'
                '"Check - no such product ID",'
                'IF(OR({minutes}{r}="",{minutes}{r}=0),'
                '"Add - how many minutes?",'
                'IF({stage}{r}="","Add - which stage?",'
                'IF(AND({basis}{r}="Per batch",{batch_qty}{r}=""),'
                '"Check - using the default batch size",'
                'IF({rate}{r}<MinimumWage,'
                '"Check - below minimum wage","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Labor & Time",
        subtitle="Every stage of making it, timed and priced at a real "
                 "hourly rate.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).labor,
        help_text="One row per stage. Cost = minutes / 60 x hourly "
                  "rate, and batch stages are shared across the batch.",
        extra_notes=[
            "Time yourself once, properly, with a timer running. "
            "Nearly every maker who does this discovers the piece "
            "takes half again as long as they believed.",
            "Count the invisible stages. Photographing, listing, "
            "answering messages and standing in the post office "
            "queue are hours you are not paid for unless you put "
            "them here.",
            "If a product's price cannot carry an honest hourly "
            "rate, the answer is a different product or a different "
            "price - not a quieter rate.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the labour lines."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Labour summary", last)
    sheet.table_head(row, 0, ["Stages entered", "Products covered",
                              "Total minutes per item", "Total labour cost",
                              "Average rate", "Skilled stages",
                              "Lines needing a look"],
                     [15, 17, 20, 18, 14, 15, 20])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Labor.product]],"<>")',
        '=COUNTIFS([[Labor.seq]],1)',
        '=ROUND(SUM([[Labor.mins_item]]),0)',
        '=ROUND(SUM([[Labor.cost]]),2)',
        '=IFERROR(ROUND(AVERAGE([[Labor.rate]]),2),"")',
        '=COUNTIFS([[Labor.skilled]],"Yes")',
        '=COUNTIFS([[Labor.check]],"Check*")+'
        'COUNTIFS([[Labor.check]],"Add*")',
    ], ["int", "int", "int", "money", "money", "int", "int"])
````

### `handmade_pricing_spreadsheet/sheets/packaging.py`

*220 lines · 2 functions*

> `Packaging` - the box, the tissue, the card, the sticker.
>
> Packaging is the cost makers forget and buyers judge you on. List
> every component once, say which of the four packaging profiles it
> belongs to - Budget, Standard, Gift, Premium - and the workbook adds
> up what each profile costs and weighs. A product then picks a profile
> and inherits the lot.
>
> The totals at the bottom are published as PackBudget, PackStandard,
> PackGift and PackPremium (and the matching weights), which is how the
> catalog and both price calculators read them.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Profile totals, published as defined names |

**Module constants:** `NAME` ('Packaging'), `PROFILES` (4 items).

````python
"""`Packaging` - the box, the tissue, the card, the sticker.

Packaging is the cost makers forget and buyers judge you on. List
every component once, say which of the four packaging profiles it
belongs to - Budget, Standard, Gift, Premium - and the workbook adds
up what each profile costs and weighs. A product then picks a profile
and inherits the lot.

The totals at the bottom are published as PackBudget, PackStandard,
PackGift and PackPremium (and the matching weights), which is how the
catalog and both price calculators read them.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..layout import quote_sheet
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf

NAME = "Packaging"

PROFILES = [
    ("budget", "Budget", "PackBudget", "PackWeightBudget",
     "Plain mailer, a sticker, done. For low priced items where the "
     "packaging cannot be more than the profit."),
    ("standard", "Standard", "PackStandard", "PackWeightStandard",
     "What most of your orders go out in."),
    ("gift", "Gift", "PackGift", "PackWeightGift",
     "Ribbon, tissue, a hand written card - the one people "
     "photograph."),
    ("premium", "Premium", "PackPremium", "PackWeightPremium",
     "Rigid box, insert, branded sleeve. For the pieces that carry "
     "it."),
]


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("name", "Component", 26, "text",
            note="Mailer, tissue, thank you card, twine, fragile "
                 "sticker - one row each."),
        Col("type", "Type", 18, "text", validate=list_from("PackTypeList")),
        Col("supplier", "Supplier", 16, "text"),
        Col("pack_qty", "Units in a pack", 12, "number1",
            validate=number_check(0, 1000000),
            note="500 stickers on a roll, 100 mailers in a box."),
        Col("pack_cost", "Pack price", 11, "money",
            validate=number_check(0, 100000),
            note="Including delivery, as always."),
        Col("qty", "Used per order", 11, "number1",
            validate=number_check(0, 10000),
            note="Usually 1. Two sheets of tissue is 2."),
        Col("weight", "Weight each (g)", 12, "number1",
            validate=number_check(0, 100000),
            note="Packaging weight pushes orders into the next "
                 "postage band more often than the product does."),
        Col("budget", "Budget", 9, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("standard", "Standard", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("gift", "Gift", 9, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("premium", "Premium", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("reusable", "Branded?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Branded packaging is marketing as well as cost - "
                 "worth knowing how much of the spend it is."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("unit_cost", "Cost each", 10, "money4",
            formula=blank(
                "name",
                'IF(OR({pack_cost}{r}="",{pack_qty}{r}="",'
                '{pack_qty}{r}=0),"",ROUND({pack_cost}{r}/'
                '{pack_qty}{r},4))')),
        Col("line_cost", "Cost per order", 12, "money4",
            formula=blank(
                "name",
                'IF(OR({unit_cost}{r}="",{qty}{r}=""),"",'
                'ROUND({unit_cost}{r}*{qty}{r},4))')),
        Col("line_weight", "Weight per order", 13, "number1",
            formula=blank(
                "name",
                'IF(OR({weight}{r}="",{qty}{r}=""),"",'
                'ROUND({weight}{r}*{qty}{r},1))')),
        Col("c_budget", "Budget cost", 10, "money4",
            formula=blank(
                "name",
                'IF(AND({budget}{r}="Yes",{line_cost}{r}<>""),'
                '{line_cost}{r},0)'),
            hidden=True),
        Col("c_standard", "Standard cost", 10, "money4",
            formula=blank(
                "name",
                'IF(AND({standard}{r}="Yes",{line_cost}{r}<>""),'
                '{line_cost}{r},0)'),
            hidden=True),
        Col("c_gift", "Gift cost", 10, "money4",
            formula=blank(
                "name",
                'IF(AND({gift}{r}="Yes",{line_cost}{r}<>""),'
                '{line_cost}{r},0)'),
            hidden=True),
        Col("c_premium", "Premium cost", 10, "money4",
            formula=blank(
                "name",
                'IF(AND({premium}{r}="Yes",{line_cost}{r}<>""),'
                '{line_cost}{r},0)'),
            hidden=True),
        Col("w_budget", "Budget weight", 10, "number1",
            formula=blank(
                "name",
                'IF(AND({budget}{r}="Yes",{line_weight}{r}<>""),'
                '{line_weight}{r},0)'),
            hidden=True),
        Col("w_standard", "Standard weight", 10, "number1",
            formula=blank(
                "name",
                'IF(AND({standard}{r}="Yes",{line_weight}{r}<>""),'
                '{line_weight}{r},0)'),
            hidden=True),
        Col("w_gift", "Gift weight", 10, "number1",
            formula=blank(
                "name",
                'IF(AND({gift}{r}="Yes",{line_weight}{r}<>""),'
                '{line_weight}{r},0)'),
            hidden=True),
        Col("w_premium", "Premium weight", 10, "number1",
            formula=blank(
                "name",
                'IF(AND({premium}{r}="Yes",{line_weight}{r}<>""),'
                '{line_weight}{r},0)'),
            hidden=True),
        Col("profiles", "In profiles", 11, "int",
            formula=blank(
                "name",
                'COUNTIF({budget}{r}:{premium}{r},"Yes")'),
            note="How many of the four profiles include this "
                 "component."),
        Col("check", "Check", 24, "status",
            formula=blank(
                "name",
                'IF(OR({pack_cost}{r}="",{pack_qty}{r}=""),'
                '"Add - price and pack size please",'
                'IF({profiles}{r}=0,"Check - in no profile",'
                'IF({qty}{r}="","Add - how many per order?",'
                'IF({weight}{r}="","Check - no weight, postage may be '
                'wrong","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Packaging Costs",
        subtitle="Four packaging profiles, built from the components "
                 "you actually buy.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).packaging,
        help_text="Tick the profiles each component belongs to. The "
                  "four totals at the bottom are what products "
                  "inherit when they pick a profile.",
        extra_notes=[
            "Branded packaging is a marketing decision with a price "
            "tag. Know the price tag before you decide.",
            "Weight matters as much as cost: 40 g of box and tissue "
            "can push a 180 g order into the next postage band and "
            "cost you more than the tissue did.",
            "A gift profile that costs $3.40 on a $12 product is "
            "eating a quarter of the price. Either charge for gift "
            "wrap or keep it for the pieces that can carry it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Profile totals, published as defined names."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    book = sheet.book
    ref = quote_sheet(NAME)

    row = sheet.section(row, "What each packaging profile costs", last)
    row = sheet.note(
        row,
        "These four totals are the figures the product catalog, both "
        "price calculators and the order forms read. A product picks a "
        "profile by name and inherits the cost and the weight.",
        last, height=30)
    row += 1

    row = sheet.table_head(row, 0, ["Profile", "Components", "Cost",
                                    "Weight (g)", "What it is for"],
                           [16, 13, 13, 13, 58])
    for key, label, cost_name, weight_name, blurb in PROFILES:
        sheet.write(row, 0, label, fmts.chrome("label_bold"))
        sheet.calc_row(row, 1, [
            f'=COUNTIFS([[Packaging.{key}]],"Yes")',
            f'=ROUND(SUM([[Packaging.c_{key}]]),2)',
            f'=ROUND(SUM([[Packaging.w_{key}]]),1)',
        ], ["int", "money", "number1"])
        sheet.write(row, 4, blurb, fmts.chrome("note"))
        book.define_name(cost_name, f"={ref}!$C${row + 1}")
        book.define_name(weight_name, f"={ref}!$D${row + 1}")
        row += 1

    row += 1
    sheet.note(
        row,
        "No profile that suits a product? Add a component, tick the "
        "profile it belongs to, and these totals move - along with "
        "every price in the workbook that uses them.",
        last, height=26)
````

### `handmade_pricing_spreadsheet/sheets/shipping.py`

*160 lines · 2 functions*

> `Shipping` - your postage rate card, and who is really paying it.
>
> Two numbers per line that most sellers merge into one and regret:
> what the postage costs *you*, and what the buyer is charged. When the
> buyer pays it, postage is not a cost of the product - it passes
> through. When you offer free shipping, it is entirely your cost, and
> it has to come out of the price.
>
> This tab keeps them apart. Every product picks a shipping option by
> name and the catalog takes only the part you actually absorb.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the rate card |

**Module constants:** `NAME` ('Shipping').

````python
"""`Shipping` - your postage rate card, and who is really paying it.

Two numbers per line that most sellers merge into one and regret:
what the postage costs *you*, and what the buyer is charged. When the
buyer pays it, postage is not a cost of the product - it passes
through. When you offer free shipping, it is entirely your cost, and
it has to come out of the price.

This tab keeps them apart. Every product picks a shipping option by
name and the catalog takes only the part you actually absorb.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, loss_cf,
                     min_number)

NAME = "Shipping"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("service", "Service", 22, "text",
            validate=list_from("ShipServiceList"),
            note="The postage product you actually buy."),
        Col("zone", "Zone", 16, "text", validate=list_from("ShipZoneList")),
        Col("max_weight", "Up to (g)", 11, "number1",
            validate=number_check(0, 1000000),
            note="The weight limit of this band. Product weight plus "
                 "packaging weight has to fit under it."),
        Col("cost", "Postage costs you", 13, "money",
            validate=number_check(0, 10000),
            note="What the label costs, including any fuel or "
                 "handling surcharge."),
        Col("packaging_extra", "Extra packaging", 12, "money",
            validate=number_check(0, 10000),
            note="Anything this service needs that the profile does "
                 "not include - a box for fragile items, a tracked "
                 "label."),
        Col("charge", "Buyer is charged", 13, "money",
            validate=number_check(0, 10000),
            note="0 for free shipping. Whatever you put here, the "
                 "workbook never pretends it covers more than it "
                 "does."),
        Col("payer", "Who pays", 20, "text",
            validate=list_from("ShipPayerList"),
            note="'Free shipping (I absorb it)' moves the whole cost "
                 "into the product price - which is exactly why free "
                 "shipping needs a higher price, not a thinner "
                 "margin."),
        Col("tracked", "Tracked?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("days", "Days", 8, "int", validate=number_check(0, 365,
                                                            decimal=False)),
        Col("carrier", "Carrier", 16, "text"),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("label", "Option", 34, "text",
            formula=blank(
                "service",
                '{service}{r}&" - "&{zone}{r}&IF({max_weight}{r}="",""," to "'
                '&TEXT({max_weight}{r},"#,##0")&"g")'),
            note="The name products pick from. Built automatically so "
                 "two bands of the same service never collide."),
        Col("total_cost", "Costs you", 11, "money",
            formula=blank(
                "service",
                'ROUND(IF({cost}{r}="",0,{cost}{r})+'
                'IF({packaging_extra}{r}="",0,{packaging_extra}{r}),2)')),
        Col("absorbed", "You absorb", 11, "money",
            formula=blank(
                "service",
                'ROUND(MAX(0,{total_cost}{r}-IF({charge}{r}="",0,'
                '{charge}{r})),2)'),
            note="The part that has to come out of the product price. "
                 "This is what the catalog adds to a product's cost."),
        Col("surplus", "You keep", 11, "money",
            formula=blank(
                "service",
                'ROUND(IF({charge}{r}="",0,{charge}{r})-'
                '{total_cost}{r},2)'),
            cf=loss_cf(),
            note="Positive means the postage you charge more than "
                 "covers the label. Negative is the per-order hole "
                 "the price has to fill."),
        Col("per_gram", "Cost per 100 g", 12, "money",
            formula=blank(
                "service",
                'IF(OR({max_weight}{r}="",{max_weight}{r}=0),"",'
                'ROUND({total_cost}{r}/{max_weight}{r}*100,3))'),
            hidden=True),
        Col("used", "Products using it", 12, "int",
            formula=blank(
                "service",
                'COUNTIFS([[Products.ship_option]],{label}{r})')),
        Col("check", "Check", 26, "status",
            formula=blank(
                "service",
                'IF({cost}{r}="","Add - what does the label cost?",'
                'IF({max_weight}{r}="","Add - what is the weight limit?",'
                'IF(AND({payer}{r}="Buyer pays",{charge}{r}=""),'
                '"Add - what are you charging?",'
                'IF({surplus}{r}<-0.5,"Loss - postage is costing you",'
                'IF(AND({charge}{r}>0,{surplus}{r}>1),'
                '"Check - you are over charging postage","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Shipping Rates",
        subtitle="What postage costs you, what the buyer pays, and the "
                 "difference you have to price for.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).shipping,
        help_text="One row per service and weight band. Products pick "
                  "an option by name; only the part you absorb goes "
                  "into the product cost.",
        extra_notes=[
            "Free shipping is a price decision, not a kindness. If "
            "you absorb $4.20 of postage on a $18 item, your margin "
            "just dropped 23 points - the workbook shows you "
            "exactly that.",
            "Weigh a finished, packed order on kitchen scales. "
            "Guessing the band is the most expensive guess in "
            "handmade selling.",
            "Most marketplaces charge their transaction fee on the "
            "postage the buyer paid as well. The fee setting on the "
            "Settings tab controls whether the workbook does too.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the rate card."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Postage summary", last)
    sheet.table_head(row, 0, ["Options listed", "Free shipping options",
                              "Average label cost", "Average absorbed",
                              "Options losing money", "Cheapest domestic"],
                     [15, 20, 18, 17, 20, 18])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Shipping.service]],"<>")',
        '=COUNTIFS([[Shipping.payer]],"Free shipping (I absorb it)")',
        '=IFERROR(ROUND(AVERAGE([[Shipping.total_cost]]),2),"")',
        '=IFERROR(ROUND(AVERAGE([[Shipping.absorbed]]),2),"")',
        '=COUNTIFS([[Shipping.surplus]],"<-0.5")',
        '=IF(COUNTIFS([[Shipping.zone]],"Domestic")=0,"",ROUND(' +
        min_number("[[Shipping.total_cost]]",
                   '[[Shipping.zone]]="Domestic"') + ',2))',
    ], ["int", "int", "money", "money", "int", "money"])
````

### `handmade_pricing_spreadsheet/sheets/overhead.py`

*149 lines · 2 functions*

> `Overhead` - the bills that exist whether you sell anything or not.
>
> Studio rent, insurance, the shop subscription, the website, the
> accountant, the ring light you bought in January. None of it belongs
> to a single product, and all of it has to be paid out of the margin on
> the products you sell.
>
> Type the bills here with how often they come. The workbook turns them
> into a monthly figure, then into an hourly and a per-item rate - and
> the Settings tab decides which of those every product carries.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the bills |

**Module constants:** `NAME` ('Overhead').

````python
"""`Overhead` - the bills that exist whether you sell anything or not.

Studio rent, insurance, the shop subscription, the website, the
accountant, the ring light you bought in January. None of it belongs
to a single product, and all of it has to be paid out of the margin on
the products you sell.

Type the bills here with how often they come. The workbook turns them
into a monthly figure, then into an hourly and a per-item rate - and
the Settings tab decides which of those every product carries.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, bar_cf, blank, id_col

NAME = "Overhead"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"OH-"', digits=3, width=9),
        Col("name", "Cost", 26, "text",
            note="What the bill is called on your bank statement."),
        Col("type", "Type", 20, "text",
            validate=list_from("OverheadTypeList")),
        Col("supplier", "Paid to", 18, "text"),
        Col("amount", "Amount", 11, "money",
            validate=number_check(0, 1000000)),
        Col("frequency", "How often", 14, "text",
            validate=list_from("FrequencyList"),
            note="One-off costs are spread over twelve months, which "
                 "is the honest way to carry a kiln or a camera."),
        Col("start", "Started", 11, "date", validate=date_check()),
        Col("ends", "Ends", 11, "date", validate=date_check(),
            note="Leave blank for an ongoing cost. A finished cost "
                 "stops counting against your products."),
        Col("essential", "Essential?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Be honest. The reports total what you could stop "
                 "paying tomorrow."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("active", "Active", 8, "flag",
            formula=blank(
                "name",
                'IF(AND({ends}{r}<>"",{ends}{r}<TodayDate),"No","Yes")'),
            hidden=True),
        Col("monthly", "A month", 11, "money",
            formula=blank(
                "name",
                'IF(OR({amount}{r}="",{active}{r}="No"),0,'
                'ROUND({amount}{r}*IF({frequency}{r}="Monthly",1,'
                'IF({frequency}{r}="Quarterly",1/3,'
                'IF({frequency}{r}="Twice a year",1/6,'
                'IF({frequency}{r}="Yearly",1/12,'
                'IF({frequency}{r}="One-off",1/12,0))))),2))'),
            cf=bar_cf(),
            note="Everything converted to what it costs you each "
                 "month. The Settings tab adds this column up."),
        Col("yearly", "A year", 11, "money",
            formula=blank("name", 'ROUND({monthly}{r}*12,2)')),
        Col("share", "Share of overhead", 12, "percent",
            formula=blank(
                "name",
                'IF(OverheadMonthly=0,"",{monthly}{r}/OverheadMonthly)'),
            cf=bar_cf()),
        Col("per_hour", "Per hour worked", 12, "money",
            formula=blank(
                "name",
                'IF(ProductiveHours<=0,"",ROUND({monthly}{r}/'
                'ProductiveHours,2))')),
        Col("per_unit", "Per item made", 11, "money",
            formula=blank(
                "name",
                'IF(UnitsPerMonth<=0,"",ROUND({monthly}{r}/'
                'UnitsPerMonth,2))')),
        Col("units_to_cover", "Items to cover it", 13, "int",
            formula=blank(
                "name",
                'IF(OR(OverheadMonthly=0,{monthly}{r}=0),"",'
                'ROUNDUP({monthly}{r}/MAX(0.01,'
                'IFERROR(AVERAGE([[Products.profit]]),0.01)),0))'),
            note="How many average products you have to sell each "
                 "month just to pay this one bill."),
        Col("check", "Check", 24, "status",
            formula=blank(
                "name",
                'IF({amount}{r}="","Add - how much is it?",'
                'IF({frequency}{r}="","Add - how often?",'
                'IF({active}{r}="No","Ended - no longer counted",'
                'IF(AND({essential}{r}="No",{monthly}{r}>'
                'OverheadMonthly*0.15),'
                '"Check - a big bill you called optional","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Overhead Costs",
        subtitle="The bills that are not about any one product, spread "
                 "across everything you make.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).overhead,
        help_text="Enter the bill and how often it arrives. The "
                  "monthly column is what the Settings tab divides "
                  "across your hours or your items.",
        extra_notes=[
            "Overhead is the quiet difference between a hobby that "
            "breaks even and a business that pays you. A $420 month "
            "of overhead spread over 55 items is $7.64 a product - "
            "on a $20 item that is more than a third of the margin.",
            "Spread one-off purchases over a year rather than "
            "dropping the whole cost on one month. The kiln does "
            "not stop being useful in February.",
            "Review the 'not essential' lines twice a year. Most "
            "makers find a subscription they stopped using.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the bills."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Overhead summary", last)
    sheet.table_head(row, 0, ["Costs listed", "Total a month",
                              "Total a year", "Per hour worked",
                              "Per item made", "Could be cut",
                              "Items to break even"],
                     [13, 15, 15, 16, 15, 14, 19])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Overhead.name]],"<>")',
        '=ROUND(SUM([[Overhead.monthly]]),2)',
        '=ROUND(SUM([[Overhead.yearly]]),2)',
        '=OverheadPerHour',
        '=OverheadPerUnit',
        '=ROUND(SUMIFS([[Overhead.monthly]],[[Overhead.essential]],"No"),2)',
        '=IFERROR(ROUNDUP(OverheadMonthly/MAX(0.01,'
        'AVERAGE([[Products.profit]])),0),"")',
    ], ["int", "money", "money", "money", "money", "money", "int"])
````

### `handmade_pricing_spreadsheet/sheets/channels.py`

*222 lines · 2 functions*

> `Channels` - what each place you sell costs you to sell there.
>
> Etsy, your own site, a craft fair, wholesale, Instagram. Each one
> takes a different cut, and a product that makes 48% on your own
> website can make 22% on a marketplace once listing, transaction,
> processing, ads and regulatory fees have all had a bite.
>
> Every fee on this tab is a number you can edit. Nothing is hard coded
> anywhere in this workbook - and nothing here is advice about what the
> fees are. Marketplace fees change and they differ by country: check
> the current official fee schedule for your shop and type what you
> find.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the fee table |

**Module constants:** `NAME` ('Channels').

````python
"""`Channels` - what each place you sell costs you to sell there.

Etsy, your own site, a craft fair, wholesale, Instagram. Each one
takes a different cut, and a product that makes 48% on your own
website can make 22% on a marketplace once listing, transaction,
processing, ads and regulatory fees have all had a bite.

Every fee on this tab is a number you can edit. Nothing is hard coded
anywhere in this workbook - and nothing here is advice about what the
fees are. Marketplace fees change and they differ by country: check
the current official fee schedule for your shop and type what you
find.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf

NAME = "Channels"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("name", "Channel", 22, "text",
            note="The name products pick from. Add your own - a "
                 "second marketplace, a gallery, a subscription box."),
        Col("kind", "Fee shape", 18, "text",
            validate=list_from("FeeTypeList"),
            note="Documentation only - the individual fees below are "
                 "what the maths uses."),
        Col("listing_fee", "Listing fee", 11, "money",
            validate=number_check(0, 1000),
            note="Charged per listing, and again each time it "
                 "renews. Spread over the items a listing sells."),
        Col("listing_sales", "Sales per listing", 13, "number1",
            validate=number_check(0, 100000),
            note="How many items one listing sells before it renews. "
                 "2 means each sale carries half a listing fee."),
        Col("transaction", "Transaction fee", 12, "percent",
            validate=number_check(0, 1)),
        Col("processing_pct", "Processing %", 12, "percent",
            validate=number_check(0, 1)),
        Col("processing_fixed", "Processing fixed", 13, "money",
            validate=number_check(0, 100),
            note="The flat part of the payment fee. On a $6 order "
                 "this hurts far more than the percentage."),
        Col("ads_fee", "Offsite ads fee", 12, "percent",
            validate=number_check(0, 1)),
        Col("ads_share", "Orders from ads", 12, "percent",
            validate=number_check(0, 1),
            note="What share of your orders arrive through the "
                 "marketplace's own advertising. Your statements "
                 "will tell you."),
        Col("regulatory", "Regulatory fee", 11, "percent",
            validate=number_check(0, 1)),
        Col("currency_fee", "Currency fee", 11, "percent",
            validate=number_check(0, 1)),
        Col("other_pct", "Other %", 10, "percent",
            validate=number_check(0, 1),
            note="Anything else that scales with the order - a "
                 "gallery commission, an affiliate cut."),
        Col("other_fixed", "Other fixed", 11, "money",
            validate=number_check(0, 10000),
            note="Stall fee per item, packing surcharge, a per-order "
                 "charge."),
        Col("subscription", "Monthly subscription", 14, "money",
            validate=number_check(0, 100000),
            note="A shop plan or a pitch rent. Not charged per "
                 "order - it belongs in your overhead, and the "
                 "reminder column says so."),
        Col("fees_on_shipping", "Fees on postage?", 13, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Most marketplaces charge their percentage on the "
                 "postage the buyer paid as well as the item."),
        Col("reviewed", "Fees checked", 11, "date", validate=date_check(),
            note="The date you last read the official fee page. Put "
                 "today's date in when you do."),
        Col("url", "Fee schedule", 26, "text",
            note="Where you read it. Paste the link to the official "
                 "page so the next check takes ten seconds."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("pct_total", "Percentage fees", 12, "percent",
            formula=blank(
                "name",
                'IF({transaction}{r}="",0,{transaction}{r})+'
                'IF({processing_pct}{r}="",0,{processing_pct}{r})+'
                'IF({regulatory}{r}="",0,{regulatory}{r})+'
                'IF({currency_fee}{r}="",0,{currency_fee}{r})+'
                'IF({other_pct}{r}="",0,{other_pct}{r})+'
                'IF({ads_fee}{r}="",0,{ads_fee}{r})*'
                'IF({ads_share}{r}="",0,{ads_share}{r})'),
            note="Everything that scales with the price, added up, "
                 "with offsite ads weighted by how many orders "
                 "actually come from them. This is the 'r' in the "
                 "pricing formula."),
        Col("fixed_total", "Fixed fees", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({processing_fixed}{r}="",0,'
                '{processing_fixed}{r})+'
                'IF({other_fixed}{r}="",0,{other_fixed}{r})+'
                'IF(OR({listing_fee}{r}="",{listing_sales}{r}="",'
                '{listing_sales}{r}=0),0,{listing_fee}{r}/'
                '{listing_sales}{r}),4)'),
            note="Everything charged per order regardless of price, "
                 "with the listing fee spread over the sales one "
                 "listing makes. This is the 'F'."),
        Col("fee_20", "Fee on a 20 sale", 12, "money",
            formula=blank(
                "name",
                'ROUND(20*{pct_total}{r}+{fixed_total}{r},2)'),
            note="What this channel would take from a sale of 20 in "
                 "your currency. The quickest way to compare two "
                 "places to sell."),
        Col("keep_20", "You keep", 11, "money",
            formula=blank("name", 'ROUND(20-{fee_20}{r},2)')),
        Col("take_rate", "Effective take", 11, "percent",
            formula=blank("name", '{fee_20}{r}/20'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.15,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0.06,
                 "fmt": "good"}],
            note="The honest headline number: what this channel "
                 "costs you as a share of a typical sale."),
        Col("products", "Products on it", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Products.channel]],{name}{r})')),
        Col("sales_count", "Sales recorded", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Sales.channel]],{name}{r})')),
        Col("sales_value", "Sold through it", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.revenue]],[[Sales.channel]],'
                '{name}{r}),2)')),
        Col("fees_paid", "Fees actually paid", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.fees]],[[Sales.channel]],'
                '{name}{r}),2)'),
            note="From the Sales tab - what this channel really took "
                 "once the orders were in. Compare it with the "
                 "estimate above."),
        Col("actual_rate", "Actual take", 11, "percent",
            formula=blank(
                "name",
                'IF(OR({sales_value}{r}="",{sales_value}{r}=0),"",'
                '{fees_paid}{r}/{sales_value}{r})'),
            note="If this is well above the estimate, your fee "
                 "settings are out of date."),
        Col("check", "Check", 26, "status",
            formula=blank(
                "name",
                'IF({reviewed}{r}="","Check - fees never verified",'
                'IF(TodayDate-{reviewed}{r}>FeeReviewDays,'
                '"Check - verify the fee schedule",'
                'IF({pct_total}{r}=0,"Check - no fees entered",'
                'IF(AND({actual_rate}{r}<>"",{actual_rate}{r}>'
                '{take_rate}{r}+0.02),'
                '"Check - real fees are higher than these settings",'
                'IF({subscription}{r}>0,'
                '"Note - put the subscription in Overhead too",'
                '"OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Sales Channels & Fees",
        subtitle="Every fee, editable, per place you sell - and what "
                 "each one really takes from a sale.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).channels,
        help_text="Check the official fee schedule for your country, "
                  "type what you find, and date the row. Every price "
                  "in the workbook uses these numbers.",
        extra_notes=[
            "VERIFY THESE FIGURES. The defaults this workbook ships "
            "with are a plausible starting point from public fee "
            "pages, not a quotation - fees change, and they are "
            "different in different countries.",
            "The fixed part of a payment fee is what kills small "
            "items. A 30 cent fee on a $4 sticker is 7.5% before "
            "anybody has taken a percentage.",
            "Comparing channels is the point of this tab. The same "
            "product can be worth making for one channel and not "
            "another, and that is a decision you can only make with "
            "the take rate in front of you.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the fee table."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Channel comparison", last)
    sheet.table_head(row, 0, ["Channels set up", "Cheapest to sell on",
                              "Dearest to sell on", "Average take",
                              "Fees paid this year", "Needing a fee check"],
                     [16, 22, 22, 15, 19, 21])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Channels.name]],"<>")',
        '=IFERROR(INDEX([[Channels.name]],MATCH(MIN([[Channels.take_rate]]),'
        '[[Channels.take_rate]],0)),"")',
        '=IFERROR(INDEX([[Channels.name]],MATCH(MAX([[Channels.take_rate]]),'
        '[[Channels.take_rate]],0)),"")',
        '=IFERROR(AVERAGE([[Channels.take_rate]]),"")',
        '=ROUND(SUM([[Channels.fees_paid]]),2)',
        '=COUNTIFS([[Channels.check]],"Check*")',
    ], ["int", "text", "text", "percent", "money", "int"])
````

---

## 4. The six craft calculators

The reason this is not a generic costing sheet. Crochet costs yarn by the gram, jewelry works in metal weight and scrap recovery, candles fill a vessel and dose the fragrance, clothing cuts fabric with a layout allowance, woodworking carries offcut waste and machine time, and a digital download spreads its design hours over the sales it expects.

### `handmade_pricing_spreadsheet/sheets/crochet.py`

*209 lines · 1 function*

> `Crochet` - yarn by the gram, because that is how it is used.
>
> Nobody uses a whole skein. This tab costs a crochet or knitting piece
> the way it is actually made: so many grams of this yarn, a few of
> that, a couple of buttons, the pattern you bought spread over the
> pieces you will make from it, and the hours it really takes.
>
> Quantities are in the material's base unit - grams for yarn - and the
> unit table on the Lists tab converts whatever the skein was sold in.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Crochet').

````python
"""`Crochet` - yarn by the gram, because that is how it is used.

Nobody uses a whole skein. This tab costs a crochet or knitting piece
the way it is actually made: so many grams of this yarn, a few of
that, a couple of buttons, the pattern you bought spread over the
pieces you will make from it, and the hours it really takes.

Quantities are in the material's base unit - grams for yarn - and the
unit table on the Lists tab converts whatever the skein was sold in.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Crochet"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')
_PACK = ('IFERROR(INDEX([[Materials.pack_qty]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')
_FACTOR = ('IFERROR(INDEX([[Materials.base_factor]],MATCH({key}{r},'
           '[[Materials.id]],0)),1)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs"),
            note="Link it to the catalog and the costs below flow "
                 "straight into the product's price. Leave it blank "
                 "to cost a piece you have not listed yet."),
        Col("name", "Piece", 28, "text",
            note="What you are making: 'Cable beanie, adult', "
                 "'Granny square blanket, throw'."),
        Col("pattern", "Pattern", 22, "text"),
        Col("size", "Size", 14, "text", validate=list_from("SizeNames"),
            note="Stretches both the yarn and the time by the size "
                 "factors on the Lists tab."),
        Col("yarn_weight", "Yarn weight", 16, "text",
            validate=list_from("YarnWeightList")),
        Col("hook", "Hook", 10, "text", validate=list_from("HookSizeList")),
        Col("yarn", "Main yarn", 12, "id", validate=list_from("MaterialIDs"),
            note="The material ID from the Materials tab."),
        Col("grams", "Grams used", 11, "number1",
            validate=number_check(0, 100000),
            note="Weigh the finished piece, or weigh the skein "
                 "before and after. This single number is the "
                 "difference between a guess and a price."),
        Col("yarn2", "Second yarn", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("grams2", "Grams used", 11, "number1",
            validate=number_check(0, 100000)),
        Col("waste", "Yarn waste", 10, "percent",
            validate=number_check(0, 1),
            note="Tails, swatches, the bit you frogged. Blank uses "
                 "the default on Settings."),
        Col("notions", "Buttons, zips, eyes", 13, "money",
            validate=number_check(0, 10000)),
        Col("stuffing", "Stuffing / lining", 12, "money",
            validate=number_check(0, 10000)),
        Col("pattern_cost", "Pattern cost", 11, "money",
            validate=number_check(0, 10000),
            note="What you paid for the pattern or the licence."),
        Col("pattern_uses", "Pieces from it", 11, "int",
            validate=number_check(1, 100000, decimal=False),
            note="How many pieces you expect to make from that "
                 "pattern. The cost is spread over them."),
        Col("hours_input", "Hours hooking", 11, "number",
            validate=number_check(0, 1000),
            note="Actual hook time. Most makers are surprised the "
                 "first time they measure it honestly."),
        Col("finish_mins", "Finishing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Weaving in ends, blocking, sewing up, buttons. "
                 "The part everybody forgets to charge for."),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("size_mat", "Size factor", 9, "number",
            formula=blank(
                "name",
                'IFERROR(INDEX(SizeMaterialFactors,MATCH({size}{r},'
                'SizeNames,0)),1)'),
            hidden=True),
        Col("size_time", "Time factor", 9, "number",
            formula=blank(
                "name",
                'IFERROR(INDEX(SizeTimeFactors,MATCH({size}{r},'
                'SizeNames,0)),1)'),
            hidden=True),
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "name", 'IF({waste}{r}="",DefaultWaste,{waste}{r})'),
            hidden=True),
        Col("grams_total", "Grams, with waste", 12, "number1",
            formula=blank(
                "name",
                'ROUND((IF({grams}{r}="",0,{grams}{r})+'
                'IF({grams2}{r}="",0,{grams2}{r}))*{size_mat}{r}*'
                '(1+{waste_rate}{r}),1)'),
            note="Everything the piece really consumes, size and "
                 "waste included."),
        Col("yarn_cost", "Main yarn cost", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({grams}{r}="",0,{grams}{r})*{size_mat}{r}*'
                '(1+{waste_rate}{r})*' + _BASE.replace("{key}", "{yarn}") + ',2)')),
        Col("yarn2_cost", "Second yarn cost", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({grams2}{r}="",0,{grams2}{r})*{size_mat}{r}*'
                '(1+{waste_rate}{r})*' + _BASE.replace("{key}", "{yarn2}") +
                ',2)')),
        Col("skeins", "Skeins needed", 11, "number1",
            formula=blank(
                "name",
                'IF(' + _PACK.replace("{key}", "{yarn}") + '=0,"",'
                'ROUNDUP(IF({grams}{r}="",0,{grams}{r})*{size_mat}{r}*'
                '(1+{waste_rate}{r})/(' + _PACK.replace("{key}", "{yarn}") + '*' +
                _FACTOR.replace("{key}", "{yarn}") + '),0))'),
            note="How many balls you have to buy to make one. Buy "
                 "whole skeins, cost the grams - but know the "
                 "difference, because the leftovers are real money."),
        Col("leftover", "Grams left over", 12, "number1",
            formula=blank(
                "name",
                'IF({skeins}{r}="","",ROUND({skeins}{r}*' +
                _PACK.replace("{key}", "{yarn}") + '*' +
                _FACTOR.replace("{key}", "{yarn}") + '-IF({grams}{r}="",0,'
                '{grams}{r})*{size_mat}{r}*(1+{waste_rate}{r}),1))'),
            hidden=True),
        Col("pattern_share", "Pattern per piece", 12, "money",
            formula=blank(
                "name",
                'IF({pattern_cost}{r}="",0,ROUND({pattern_cost}{r}/'
                'MAX(1,IF({pattern_uses}{r}="",1,{pattern_uses}{r})),2))'),
            hidden=True),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({yarn_cost}{r}+{yarn2_cost}{r}+'
                'IF({notions}{r}="",0,{notions}{r})+'
                'IF({stuffing}{r}="",0,{stuffing}{r})+'
                '{pattern_share}{r},2)'),
            note="Yarn, notions, stuffing and the pattern's share. "
                 "This is what flows into the catalog."),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'ROUND((IF({hours_input}{r}="",0,{hours_input}{r})*60+'
                'IF({finish_mins}{r}="",0,{finish_mins}{r}))*'
                '{size_time}{r},1)')),
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank(
                "name",
                'ROUND({minutes}{r}/60*{rate}{r},2)'),
            note="Minutes / 60 x your hourly rate. A blanket that "
                 "takes 22 hours is not a $45 blanket."),
        Col("other_cost", "Extras", 10, "money",
            formula=blank(
                "name",
                'ROUND({minutes}{r}/60*SuppliesPerHour,2)'),
            hidden=True),
        Col("cost_per_gram", "Cost a gram", 11, "money4",
            formula=blank(
                "name",
                'IF(OR({grams_total}{r}="",{grams_total}{r}=0),"",'
                'ROUND({mat_cost}{r}/{grams_total}{r},4))'),
            hidden=True),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Crochet & Knitting Calculator",
        subtitle="Yarn by the gram, notions, the pattern's share and "
                 "the hours it really takes.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).crochet,
        help_text="Weigh the piece, enter the grams, and the yarn "
                  "costs itself. Link the row to a product ID and "
                  "the catalog picks the whole lot up.",
        extra_notes=[
            "Weigh the finished piece on kitchen scales. A 94 g "
            "beanie from a $9.40 100 g skein used $8.84 of yarn - "
            "no estimate gets that close.",
            "Hours are the real cost of crochet. At $25 an hour a "
            "six hour shawl carries $150 of labour before a single "
            "gram of yarn, which is why handmade crochet cannot "
            "compete with machine made and should never try.",
            "Use this tab OR the bill of materials for a given "
            "product, not both - the catalog adds them together.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/jewelry.py`

*196 lines · 1 function*

> `Jewelry` - metal by the gram, stones by the piece, scrap included.
>
> Silver is priced by weight and the price moves, findings are pennies
> that add up, and the bench time on a soldered piece is most of what
> you are selling. This tab costs all of it, including the metal that
> ends up as scrap on the bench mat.
>
> Quantities are in the material's base unit: grams of metal,
> centimetres of chain and wire.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Jewelry').

````python
"""`Jewelry` - metal by the gram, stones by the piece, scrap included.

Silver is priced by weight and the price moves, findings are pennies
that add up, and the bench time on a soldered piece is most of what
you are selling. This tab costs all of it, including the metal that
ends up as scrap on the bench mat.

Quantities are in the material's base unit: grams of metal,
centimetres of chain and wire.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Jewelry"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Piece", 28, "text",
            note="'Hammered silver bangle', 'Moonstone drop "
                 "earrings'."),
        Col("metal", "Metal", 16, "text", validate=list_from("MetalList")),
        Col("metal_id", "Metal material", 12, "id",
            validate=list_from("MaterialIDs"),
            note="The material ID you buy the metal as - sheet, "
                 "wire, casting grain."),
        Col("metal_g", "Metal used (g)", 12, "number1",
            validate=number_check(0, 100000),
            note="Weigh the finished piece. Add the sprue, the "
                 "filings and the bits that became scrap in the "
                 "scrap column."),
        Col("metal_rate_input", "Price a gram", 11, "money4",
            validate=number_check(0, 10000),
            note="Override the material price - useful when the "
                 "metal price has moved since you bought it."),
        Col("scrap", "Scrap / loss", 10, "percent",
            validate=number_check(0, 1),
            note="Filings, cut-offs, failed solders. Real on every "
                 "soldered piece, and worth recovering."),
        Col("recovery", "Scrap recovered", 11, "percent",
            validate=number_check(0, 1),
            note="How much of that scrap you actually send back or "
                 "re-melt. Credited against the cost."),
        Col("stone", "Stone", 16, "text", validate=list_from("StoneList")),
        Col("stone_count", "Stones", 9, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("stone_cost", "Cost a stone", 11, "money",
            validate=number_check(0, 100000)),
        Col("chain_cm", "Chain (cm)", 10, "number1",
            validate=number_check(0, 10000)),
        Col("chain_id", "Chain material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("wire_cm", "Wire (cm)", 10, "number1",
            validate=number_check(0, 10000)),
        Col("wire_id", "Wire material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("findings", "Findings", 11, "money",
            validate=number_check(0, 10000),
            note="Clasps, jump rings, ear wires, posts and backs. "
                 "Count them - they are 60 cents each and they are "
                 "on every piece."),
        Col("solder", "Solder & consumables", 13, "money",
            validate=number_check(0, 10000),
            note="Solder, flux, pickle, polishing compound, saw "
                 "blades."),
        Col("plating", "Plating / finishing", 12, "money",
            validate=number_check(0, 10000)),
        Col("hallmark", "Hallmarking", 11, "money",
            validate=number_check(0, 10000),
            note="Per piece, including the courier. Required over "
                 "the weight threshold in several countries."),
        Col("design_mins", "Design minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("bench_mins", "Bench minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Sawing, forming, soldering, setting."),
        Col("polish_mins", "Polishing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("metal_rate", "Metal a gram", 11, "money4",
            formula=blank(
                "name",
                'IF({metal_rate_input}{r}<>"",{metal_rate_input}{r},' +
                _BASE.replace("{key}", "{metal_id}") + ')'),
            hidden=True),
        Col("metal_gross", "Metal with scrap", 11, "number1",
            formula=blank(
                "name",
                'ROUND(IF({metal_g}{r}="",0,{metal_g}{r})*'
                '(1+IF({scrap}{r}="",0,{scrap}{r})),2)'),
            hidden=True),
        Col("metal_cost", "Metal cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({metal_gross}{r}*{metal_rate}{r}-'
                'IF({metal_g}{r}="",0,{metal_g}{r})*'
                'IF({scrap}{r}="",0,{scrap}{r})*'
                'IF({recovery}{r}="",0,{recovery}{r})*{metal_rate}{r},2)'),
            note="Metal used plus the scrap it makes, less whatever "
                 "of that scrap you recover."),
        Col("stones_cost", "Stones cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({stone_count}{r}="",0,{stone_count}{r})*'
                'IF({stone_cost}{r}="",0,{stone_cost}{r}),2)')),
        Col("chain_cost", "Chain cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({chain_cm}{r}="",0,{chain_cm}{r})*' +
                _BASE.replace("{key}", "{chain_id}") + ',2)')),
        Col("wire_cost", "Wire cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({wire_cm}{r}="",0,{wire_cm}{r})*' +
                _BASE.replace("{key}", "{wire_id}") + ',2)')),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({metal_cost}{r}+{stones_cost}{r}+{chain_cost}{r}+'
                '{wire_cost}{r}+IF({findings}{r}="",0,{findings}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'IF({design_mins}{r}="",0,{design_mins}{r})+'
                'IF({bench_mins}{r}="",0,{bench_mins}{r})+'
                'IF({polish_mins}{r}="",0,{polish_mins}{r})')),
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank(
                "name", 'ROUND({minutes}{r}/60*{rate}{r},2)')),
        Col("other_cost", "Extras", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({solder}{r}="",0,{solder}{r})+'
                'IF({plating}{r}="",0,{plating}{r})+'
                'IF({hallmark}{r}="",0,{hallmark}{r})+'
                '{minutes}{r}/60*SuppliesPerHour,2)'),
            note="Solder, plating, hallmarking and the bench "
                 "consumables charged per hour."),
        Col("metal_share", "Metal share", 10, "percent",
            formula=blank(
                "name",
                'IF(OR({mat_cost}{r}="",{mat_cost}{r}=0),"",'
                '{metal_cost}{r}/{mat_cost}{r})'),
            hidden=True,
            note="When this is most of your material cost, a silver "
                 "price rise is a re-pricing job, not a rounding "
                 "error."),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Jewelry Calculator",
        subtitle="Metal by the gram with scrap, stones, findings, "
                 "chain and bench time.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).jewelry,
        help_text="Weigh the metal, count the findings, time the "
                  "bench work. Link the row to a product ID and the "
                  "catalog prices it.",
        extra_notes=[
            "Metal prices move. Keep the price a gram current - a "
            "25% rise in silver on a piece that is 60% metal is a "
            "15% cost rise you will not feel until the quarter is "
            "over.",
            "Findings are the quiet killer. Ear wires, jump rings, "
            "a clasp and a tag can be $1.80 on a pair of earrings "
            "that cost $2.10 in silver.",
            "Recovering scrap is real money. 8% scrap on 12 g of "
            "sterling, recovered, is most of a cup of coffee on "
            "every single piece.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/candles.py`

*205 lines · 1 function*

> `Candles` - wax by the gram, worked out from the vessel.
>
> Tell the workbook how big the jar is and how full you pour it and it
> works out the wax, the fragrance load on top of that, the wick, the
> dye and the vessel. Melt loss is included, because the wax left in
> the pouring pitcher was still bought and paid for.
>
> Quantities are in base units: grams of wax and fragrance, millilitres
> of vessel capacity.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Candles').

````python
"""`Candles` - wax by the gram, worked out from the vessel.

Tell the workbook how big the jar is and how full you pour it and it
works out the wax, the fragrance load on top of that, the wick, the
dye and the vessel. Melt loss is included, because the wax left in
the pouring pitcher was still bought and paid for.

Quantities are in base units: grams of wax and fragrance, millilitres
of vessel capacity.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Candles"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Candle", 28, "text",
            note="'Fig & cedar, 220 ml amber jar'."),
        Col("wax_type", "Wax", 16, "text", validate=list_from("WaxList")),
        Col("wax_id", "Wax material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("vessel", "Vessel", 16, "text",
            validate=list_from("VesselList")),
        Col("vessel_ml", "Vessel size (ml)", 13, "number1",
            validate=number_check(0, 100000),
            note="The capacity to the brim. The fill level below "
                 "turns it into the wax you actually pour."),
        Col("fill", "Fill level", 10, "percent",
            validate=number_check(0, 1),
            note="How full you pour it. 85% is typical - leave the "
                 "headroom or the lid sticks."),
        Col("density", "Wax density", 10, "number",
            validate=number_check(0.1, 2),
            note="Grams per millilitre. Soy is about 0.9. Blank "
                 "uses 0.9."),
        Col("wax_input", "Wax override (g)", 13, "number1",
            validate=number_check(0, 100000),
            note="If you already know the pour weight, type it and "
                 "the vessel maths is ignored."),
        Col("melt_loss", "Melt loss", 10, "percent",
            validate=number_check(0, 1),
            note="Wax left in the pitcher and on the thermometer. "
                 "Blank uses the default waste on Settings."),
        Col("fragrance_id", "Fragrance material", 13, "id",
            validate=list_from("MaterialIDs")),
        Col("fragrance_load", "Fragrance load", 12, "percent",
            validate=number_check(0, 0.3),
            note="As a share of the WAX weight, which is how "
                 "fragrance oil is specified. 8-10% is typical."),
        Col("wick", "Wick", 14, "text", validate=list_from("WickList")),
        Col("wick_count", "Wicks", 8, "int",
            validate=number_check(0, 100, decimal=False)),
        Col("wick_cost", "Cost a wick", 10, "money",
            validate=number_check(0, 1000)),
        Col("vessel_cost", "Vessel cost", 11, "money",
            validate=number_check(0, 10000),
            note="Including the lid. Vessels are usually the "
                 "biggest single cost in a candle."),
        Col("dye_cost", "Dye & additives", 11, "money",
            validate=number_check(0, 10000)),
        Col("label_cost", "Labels & warning", 12, "money",
            validate=number_check(0, 10000),
            note="Front label, base warning label. The warning "
                 "label is a legal requirement in most places."),
        Col("batch_size", "Pour batch", 10, "int",
            validate=number_check(1, 10000, decimal=False),
            note="How many you pour at once. The setup and "
                 "clean-down time is shared across them."),
        Col("batch_mins", "Batch minutes", 11, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Melting, measuring, cleaning the pitcher - once "
                 "per batch, not once per candle."),
        Col("item_mins", "Minutes a candle", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Wicking, pouring, trimming, labelling."),
        Col("cure_days", "Cure days", 10, "int",
            validate=number_check(0, 365, decimal=False),
            note="Not a cost, but it is working capital sitting on "
                 "a shelf. Two weeks of cure is two weeks you "
                 "cannot sell it."),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 22, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("wax_g", "Wax poured (g)", 12, "number1",
            formula=blank(
                "name",
                'IF({wax_input}{r}<>"",{wax_input}{r},'
                'ROUND(IF({vessel_ml}{r}="",0,{vessel_ml}{r})*'
                'IF({fill}{r}="",0.85,{fill}{r})*'
                'IF({density}{r}="",0.9,{density}{r}),1))'),
            note="Vessel size x fill level x density, or your own "
                 "pour weight."),
        Col("wax_total", "Wax bought (g)", 12, "number1",
            formula=blank(
                "name",
                'ROUND({wax_g}{r}*(1+IF({melt_loss}{r}="",DefaultWaste,'
                '{melt_loss}{r})),1)'),
            hidden=True),
        Col("wax_cost", "Wax cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({wax_total}{r}*' + _BASE.replace("{key}", "{wax_id}") +
                ',2)')),
        Col("fragrance_g", "Fragrance (g)", 11, "number1",
            formula=blank(
                "name",
                'ROUND({wax_g}{r}*IF({fragrance_load}{r}="",0,'
                '{fragrance_load}{r}),2)'),
            hidden=True),
        Col("fragrance_cost", "Fragrance cost", 12, "money",
            formula=blank(
                "name",
                'ROUND({fragrance_g}{r}*' +
                _BASE.replace("{key}", "{fragrance_id}") + ',2)'),
            note="Fragrance oil is the expensive bottle on the "
                 "shelf, and the load is a percentage of the wax, "
                 "so a bigger jar costs more twice over."),
        Col("wicks_cost", "Wicks cost", 10, "money",
            formula=blank(
                "name",
                'ROUND(IF({wick_count}{r}="",0,{wick_count}{r})*'
                'IF({wick_cost}{r}="",0,{wick_cost}{r}),2)')),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({wax_cost}{r}+{fragrance_cost}{r}+{wicks_cost}{r}+'
                'IF({vessel_cost}{r}="",0,{vessel_cost}{r})+'
                'IF({dye_cost}{r}="",0,{dye_cost}{r})+'
                'IF({label_cost}{r}="",0,{label_cost}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'ROUND(IF({item_mins}{r}="",0,{item_mins}{r})+'
                'IF({batch_mins}{r}="",0,{batch_mins}{r})/'
                'MAX(1,IF({batch_size}{r}="",DefaultBatchSize,'
                '{batch_size}{r})),1)'),
            note="Your per candle time plus this candle's share of "
                 "the batch setup."),
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank("name", 'ROUND({minutes}{r}/60*{rate}{r},2)')),
        Col("other_cost", "Extras", 10, "money",
            formula=blank(
                "name", 'ROUND({minutes}{r}/60*SuppliesPerHour,2)'),
            hidden=True),
        Col("cost_per_100g", "Cost per 100 g", 12, "money",
            formula=blank(
                "name",
                'IF(OR({wax_g}{r}="",{wax_g}{r}=0),"",'
                'ROUND({mat_cost}{r}/{wax_g}{r}*100,2))'),
            hidden=True),
        Col("burn_hours", "Burn hours", 10, "int",
            formula=blank(
                "name",
                'IF({wax_g}{r}="","",ROUND({wax_g}{r}/7,0))'),
            hidden=True,
            note="Roughly, at 7 g an hour. Buyers compare candles "
                 "on burn time, so it is worth knowing yours."),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Candle Calculator",
        subtitle="Wax from the vessel size, fragrance load on top, "
                 "wicks, dye, vessel and the pour.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).candles,
        help_text="Enter the vessel size and the fill level - the "
                  "wax, the fragrance and the cost work themselves "
                  "out from there.",
        extra_notes=[
            "Fragrance load is a percentage of the wax, not of the "
            "finished candle. 10% of 180 g of wax is 18 g of oil, "
            "and at $0.09 a gram that is $1.62 before the jar.",
            "The vessel is usually the single biggest cost. A "
            "beautiful ceramic pot can be more than the wax, the "
            "oil and the wick together - price accordingly.",
            "Cure time is working capital. Two weeks on a shelf "
            "before a candle can be sold is two weeks of money you "
            "have already spent.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/clothing.py`

*182 lines · 1 function*

> `Clothing` - fabric with the layout waste in, notions, and the sewing.
>
> Cutting a garment out of a rectangle always wastes fabric, and the
> bigger the size the more of everything it takes. This tab costs a
> sewn piece honestly: fabric by the centimetre with a layout
> allowance, lining, interfacing, thread, zips, labels, and the four
> separate jobs that make up "sewing".
>
> Fabric quantities are in centimetres of length, which is the base
> unit the Lists table converts metres and yards into.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Clothing').

````python
"""`Clothing` - fabric with the layout waste in, notions, and the sewing.

Cutting a garment out of a rectangle always wastes fabric, and the
bigger the size the more of everything it takes. This tab costs a
sewn piece honestly: fabric by the centimetre with a layout
allowance, lining, interfacing, thread, zips, labels, and the four
separate jobs that make up "sewing".

Fabric quantities are in centimetres of length, which is the base
unit the Lists table converts metres and yards into.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Clothing"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Garment", 28, "text",
            note="'Linen pinafore', 'Toddler bucket hat'."),
        Col("garment_type", "Type", 18, "text",
            validate=list_from("FabricList"),
            note="The main fabric. The list is on the Lists tab and "
                 "you can add your own."),
        Col("size", "Size", 14, "text",
            validate=list_from("GarmentSizeList")),
        Col("size_factor", "Size factor", 10, "number",
            validate=number_check(0.1, 5),
            note="How much more fabric and time this size takes "
                 "than your standard one. Blank is 1."),
        Col("fabric_id", "Fabric material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("fabric_cm", "Fabric (cm)", 11, "number1",
            validate=number_check(0, 100000),
            note="Length of fabric at its usual width. Measure the "
                 "pattern layout, do not guess."),
        Col("layout_waste", "Layout waste", 11, "percent",
            validate=number_check(0, 1),
            note="What the cutting layout wastes - 10 to 25% is "
                 "normal, more on a directional print or a bias "
                 "cut. Blank uses the Settings default."),
        Col("lining_id", "Lining material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("lining_cm", "Lining (cm)", 11, "number1",
            validate=number_check(0, 100000)),
        Col("interfacing", "Interfacing", 11, "money",
            validate=number_check(0, 10000)),
        Col("thread", "Thread", 10, "money",
            validate=number_check(0, 10000),
            note="A reel is $3 and does ten garments. 30 cents is "
                 "not nothing when the margin is thin."),
        Col("notions", "Zips, buttons, elastic", 14, "money",
            validate=number_check(0, 10000)),
        Col("labels", "Labels & care tag", 12, "money",
            validate=number_check(0, 10000)),
        Col("cut_mins", "Cutting minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Laying out, pinning, cutting. On a one-off this "
                 "is often longer than the sewing."),
        Col("sew_mins", "Sewing minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("finish_mins", "Finishing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Hems, buttonholes, hand finishing, labels."),
        Col("press_mins", "Pressing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Pressing is sewing. Count it."),
        Col("fit_mins", "Fitting & alterations", 14, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Made to measure work lives or dies on this "
                 "number."),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 22, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("factor", "Size factor used", 10, "number",
            formula=blank(
                "name",
                'IF({size_factor}{r}="",1,{size_factor}{r})'),
            hidden=True),
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "name",
                'IF({layout_waste}{r}="",DefaultWaste,'
                '{layout_waste}{r})'),
            hidden=True),
        Col("fabric_total", "Fabric bought (cm)", 13, "number1",
            formula=blank(
                "name",
                'ROUND(IF({fabric_cm}{r}="",0,{fabric_cm}{r})*'
                '{factor}{r}*(1+{waste_rate}{r}),1)'),
            note="What you have to buy, layout waste and size "
                 "included."),
        Col("fabric_cost", "Fabric cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({fabric_total}{r}*' +
                _BASE.replace("{key}", "{fabric_id}") + ',2)')),
        Col("lining_cost", "Lining cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({lining_cm}{r}="",0,{lining_cm}{r})*'
                '{factor}{r}*(1+{waste_rate}{r})*' +
                _BASE.replace("{key}", "{lining_id}") + ',2)')),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({fabric_cost}{r}+{lining_cost}{r}+'
                'IF({interfacing}{r}="",0,{interfacing}{r})+'
                'IF({thread}{r}="",0,{thread}{r})+'
                'IF({notions}{r}="",0,{notions}{r})+'
                'IF({labels}{r}="",0,{labels}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'ROUND((IF({cut_mins}{r}="",0,{cut_mins}{r})+'
                'IF({sew_mins}{r}="",0,{sew_mins}{r})+'
                'IF({finish_mins}{r}="",0,{finish_mins}{r})+'
                'IF({press_mins}{r}="",0,{press_mins}{r})+'
                'IF({fit_mins}{r}="",0,{fit_mins}{r}))*{factor}{r},1)')),
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank("name", 'ROUND({minutes}{r}/60*{rate}{r},2)')),
        Col("other_cost", "Extras", 10, "money",
            formula=blank(
                "name", 'ROUND({minutes}{r}/60*SuppliesPerHour,2)'),
            hidden=True),
        Col("waste_cost", "Cost of the waste", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({fabric_cm}{r}="",0,{fabric_cm}{r})*'
                '{factor}{r}*{waste_rate}{r}*' +
                _BASE.replace("{key}", "{fabric_id}") + ',2)'),
            hidden=True,
            note="What the offcuts cost. Worth knowing before you "
                 "decide whether a scrap-busting product is free "
                 "money or not."),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Clothing & Sewing Calculator",
        subtitle="Fabric with the layout waste in, notions, and all "
                 "four jobs that make up sewing.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).clothing,
        help_text="Enter the fabric at its usual width in "
                  "centimetres and a layout allowance. Size factors "
                  "stretch both the fabric and the time.",
        extra_notes=[
            "Cutting and pressing are sewing. A maker who only "
            "counts time at the machine undercharges by a third on "
            "every garment.",
            "Layout waste is real. 20% on $18 a metre linen is "
            "$3.60 a metre you paid for and did not use - price it "
            "in, then look for a layout that beats it.",
            "Made to measure needs a fitting allowance. If you do "
            "not charge for alterations you are working for free "
            "exactly when the customer is most demanding.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/woodworking.py`

*192 lines · 1 function*

> `Woodworking` - timber with the offcuts in, hardware, finish, machine time.
>
> Wood is sold in lengths and used in pieces, and the difference is
> sawdust and offcuts you paid for. This tab costs the board, the
> hardware, the oil and the abrasives, and it charges for machine time
> separately - because a planer and a sander wear out, use power and
> have to earn their keep.
>
> Timber quantities are in base units: centimetres of length (the Lists
> table converts metres, feet and yards).

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Woodworking').

````python
"""`Woodworking` - timber with the offcuts in, hardware, finish, machine time.

Wood is sold in lengths and used in pieces, and the difference is
sawdust and offcuts you paid for. This tab costs the board, the
hardware, the oil and the abrasives, and it charges for machine time
separately - because a planer and a sander wear out, use power and
have to earn their keep.

Timber quantities are in base units: centimetres of length (the Lists
table converts metres, feet and yards).
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Woodworking"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Piece", 28, "text",
            note="'Walnut serving board, large', 'Oak floating "
                 "shelf 60 cm'."),
        Col("species", "Species", 16, "text",
            validate=list_from("WoodList")),
        Col("wood_id", "Timber material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("wood_qty", "Timber used (cm)", 13, "number1",
            validate=number_check(0, 1000000),
            note="Length of board at the width and thickness you "
                 "buy it in. Measure the cut list, not the "
                 "finished piece."),
        Col("offcut_waste", "Offcut waste", 11, "percent",
            validate=number_check(0, 1),
            note="Saw kerf, snipe, the end you cut off because it "
                 "split. 15-30% is normal on small work."),
        Col("second_id", "Second timber", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("second_qty", "Second timber (cm)", 13, "number1",
            validate=number_check(0, 1000000)),
        Col("hardware", "Hardware", 11, "money",
            validate=number_check(0, 100000),
            note="Hinges, screws, feet, magnets, brackets."),
        Col("glue", "Glue & fixings", 11, "money",
            validate=number_check(0, 10000)),
        Col("finish_type", "Finish", 16, "text",
            validate=list_from("FinishList")),
        Col("finish_ml", "Finish used (ml)", 13, "number1",
            validate=number_check(0, 100000)),
        Col("finish_rate", "Finish cost a ml", 12, "money4",
            validate=number_check(0, 1000),
            note="A 500 ml tin of hard wax oil at $32 is $0.064 a "
                 "millilitre."),
        Col("abrasives", "Abrasives & blades", 12, "money",
            validate=number_check(0, 10000),
            note="Sanding discs die on hardwood. They are a cost "
                 "per piece, not an annual surprise."),
        Col("machine_mins", "Machine minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Time under power: planer, bandsaw, router, "
                 "sander. Charged at the machine rate as well as "
                 "your time."),
        Col("machine_rate", "Machine rate an hour", 14, "money",
            validate=number_check(0, 1000),
            note="Power, dust extraction, blade and belt wear, "
                 "tool depreciation. $6-$12 an hour is a sane "
                 "starting point for a small shop."),
        Col("mill_mins", "Milling minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("build_mins", "Assembly minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("sand_mins", "Sanding minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Always longer than anybody admits."),
        Col("finish_mins", "Finishing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Coats, waiting, de-nibbing, another coat."),
        Col("dry_days", "Drying days", 10, "int",
            validate=number_check(0, 365, decimal=False)),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 22, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "name",
                'IF({offcut_waste}{r}="",DefaultWaste,'
                '{offcut_waste}{r})'),
            hidden=True),
        Col("wood_total", "Timber bought (cm)", 13, "number1",
            formula=blank(
                "name",
                'ROUND((IF({wood_qty}{r}="",0,{wood_qty}{r})+'
                'IF({second_qty}{r}="",0,{second_qty}{r}))*'
                '(1+{waste_rate}{r}),1)'),
            hidden=True),
        Col("wood_cost", "Timber cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({wood_qty}{r}="",0,{wood_qty}{r})*'
                '(1+{waste_rate}{r})*' + _BASE.replace("{key}", "{wood_id}") +
                '+IF({second_qty}{r}="",0,{second_qty}{r})*'
                '(1+{waste_rate}{r})*' + _BASE.replace("{key}", "{second_id}") +
                ',2)')),
        Col("finish_cost", "Finish cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({finish_ml}{r}="",0,{finish_ml}{r})*'
                'IF({finish_rate}{r}="",0,{finish_rate}{r}),2)')),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({wood_cost}{r}+{finish_cost}{r}+'
                'IF({hardware}{r}="",0,{hardware}{r})+'
                'IF({glue}{r}="",0,{glue}{r})+'
                'IF({abrasives}{r}="",0,{abrasives}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'IF({mill_mins}{r}="",0,{mill_mins}{r})+'
                'IF({build_mins}{r}="",0,{build_mins}{r})+'
                'IF({sand_mins}{r}="",0,{sand_mins}{r})+'
                'IF({finish_mins}{r}="",0,{finish_mins}{r})')),
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank("name", 'ROUND({minutes}{r}/60*{rate}{r},2)')),
        Col("machine_cost", "Machine time", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({machine_mins}{r}="",0,{machine_mins}{r})/60*'
                'IF({machine_rate}{r}="",0,{machine_rate}{r}),2)'),
            note="The tools earning their replacement. Leave it out "
                 "and the lathe is a hobby that happens to make "
                 "things."),
        Col("other_cost", "Extras", 10, "money",
            formula=blank(
                "name",
                'ROUND({machine_cost}{r}+{minutes}{r}/60*'
                'SuppliesPerHour,2)'),
            hidden=True),
        Col("waste_cost", "Cost of the offcuts", 13, "money",
            formula=blank(
                "name",
                'ROUND((IF({wood_qty}{r}="",0,{wood_qty}{r})*' +
                _BASE.replace("{key}", "{wood_id}") +
                '+IF({second_qty}{r}="",0,{second_qty}{r})*' +
                _BASE.replace("{key}", "{second_id}") + ')*{waste_rate}{r},2)'),
            hidden=True),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Woodworking Calculator",
        subtitle="Timber with the offcuts in, hardware, finish, "
                 "abrasives and machine time.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).woodworking,
        help_text="Cost the cut list, not the finished piece, and "
                  "add an honest offcut allowance. Machine time is "
                  "charged on top of your hours.",
        extra_notes=[
            "Offcuts are paid-for wood. A 25% allowance on $0.42 a "
            "centimetre walnut is real money on every board, and "
            "pretending otherwise is why small woodworkers run out "
            "of timber budget.",
            "Charge machine time. Blades, belts, bits and dust "
            "bags are consumables, and a thicknesser that is not "
            "paying for its own replacement is costing you later.",
            "Drying days are not free either - a finish that needs "
            "three days between coats is three days the bench is "
            "occupied.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/digital.py`

*193 lines · 1 function*

> `Digital` - patterns and downloads: all the cost is up front.
>
> A PDF pattern costs nothing to make a second copy of, so the whole
> question is how many you will sell and whether the design time is
> ever recovered. This tab spreads the design, testing, photography and
> writing hours over the sales you honestly expect, adds the per sale
> costs that do exist, and tells you how many you have to sell before
> the file has paid for itself.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Digital').

````python
"""`Digital` - patterns and downloads: all the cost is up front.

A PDF pattern costs nothing to make a second copy of, so the whole
question is how many you will sell and whether the design time is
ever recovered. This tab spreads the design, testing, photography and
writing hours over the sales you honestly expect, adds the per sale
costs that do exist, and tells you how many you have to sell before
the file has paid for itself.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Digital"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Download", 28, "text",
            note="'Cable beanie crochet pattern', 'Farmhouse SVG "
                 "bundle'."),
        Col("kind", "Type", 20, "text",
            validate=list_from("DigitalTypeList")),
        Col("formats", "Formats", 16, "text",
            validate=list_from("FileFormatList")),
        Col("licence", "Licence", 20, "text",
            validate=list_from("LicenceList"),
            note="A commercial licence is a different product at a "
                 "different price, not a favour."),
        Col("design_hours", "Design hours", 11, "number",
            validate=number_check(0, 10000),
            note="Making the thing in the first place - the "
                 "sample, the grading, the drawing."),
        Col("write_hours", "Writing hours", 11, "number",
            validate=number_check(0, 10000),
            note="Writing the pattern up, charts, diagrams, "
                 "layout. Usually longer than making it."),
        Col("test_hours", "Testing hours", 11, "number",
            validate=number_check(0, 10000),
            note="Your time running a test round and correcting "
                 "it. Tester payments go in the fees column."),
        Col("photo_hours", "Photography hours", 13, "number",
            validate=number_check(0, 10000)),
        Col("tester_cost", "Tester & editor fees", 14, "money",
            validate=number_check(0, 100000),
            note="Free patterns for testers, a tech editor, a "
                 "proofreader. Real money, paid once."),
        Col("software_cost", "Software & stock", 12, "money",
            validate=number_check(0, 100000),
            note="The share of a subscription or stock licence "
                 "that belongs to this file."),
        Col("expected_sales", "Sales expected", 12, "int",
            validate=number_check(1, 1000000, decimal=False),
            note="Over the life of the listing. Be honest - this "
                 "single number decides whether the design time "
                 "is ever recovered."),
        Col("support_mins", "Support minutes a sale", 15, "number1",
            validate=number_check(0, 1000),
            note="Questions, errata, 'it will not open'. Two "
                 "minutes a sale is a real cost at volume."),
        Col("update_hours", "Updates a year", 11, "number",
            validate=number_check(0, 10000),
            note="Revisions, new sizes, a format nobody could "
                 "open. Spread over a year of sales."),
        Col("sold_so_far", "Sold so far", 11, "int",
            validate=number_check(0, 10000000, decimal=False),
            note="Lifetime sales of this file, for the recovery "
                 "figure."),
        Col("listed", "Listed on", 11, "date", validate=date_check()),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("dev_hours", "Development hours", 13, "number",
            formula=blank(
                "name",
                'IF({design_hours}{r}="",0,{design_hours}{r})+'
                'IF({write_hours}{r}="",0,{write_hours}{r})+'
                'IF({test_hours}{r}="",0,{test_hours}{r})+'
                'IF({photo_hours}{r}="",0,{photo_hours}{r})+'
                'IF({update_hours}{r}="",0,{update_hours}{r})')),
        Col("dev_cost", "Development cost", 13, "money",
            formula=blank(
                "name",
                'ROUND({dev_hours}{r}*{rate}{r}+'
                'IF({tester_cost}{r}="",0,{tester_cost}{r}),2)'),
            note="Everything the file cost to exist, paid once, "
                 "before a single sale."),
        Col("expected", "Sales used", 10, "int",
            formula=blank(
                "name",
                'MAX(1,IF({expected_sales}{r}="",1,'
                '{expected_sales}{r}))'),
            hidden=True),
        Col("mat_cost", "Materials", 10, "money",
            formula=blank("name", '0'),
            hidden=True,
            note="A download has no materials - which is exactly "
                 "why the development time has to be taken "
                 "seriously."),
        Col("minutes", "Minutes a sale", 12, "number1",
            formula=blank(
                "name",
                'ROUND({dev_hours}{r}*60/{expected}{r}+'
                'IF({support_mins}{r}="",0,{support_mins}{r}),1)')),
        Col("labor_cost", "Labour a sale", 12, "money",
            formula=blank(
                "name",
                'ROUND({dev_cost}{r}/{expected}{r}+'
                'IF({support_mins}{r}="",0,{support_mins}{r})/60*'
                '{rate}{r},2)'),
            note="The development cost spread over the sales you "
                 "expect, plus the support each sale really takes."),
        Col("other_cost", "Extras a sale", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({software_cost}{r}="",0,{software_cost}{r})/'
                '{expected}{r},2)'),
            hidden=True),
        Col("recovered", "Recovered so far", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({sold_so_far}{r}="",0,{sold_so_far}{r})*'
                'MAX(0,{profit}{r}+{labor_cost}{r}-'
                'IF({support_mins}{r}="",0,{support_mins}{r})/60*'
                '{rate}{r}),2)'),
            note="What the sales so far have put back against the "
                 "development cost."),
        Col("recovery", "Development recovered", 13, "percent",
            formula=blank(
                "name",
                'IF({dev_cost}{r}=0,"",{recovered}{r}/{dev_cost}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "warn"}]),
        Col("break_even", "Sales to break even", 14, "int",
            formula=blank(
                "name",
                'IF(OR({price}{r}="",{price}{r}-{other_cost}{r}-'
                '({price}{r}*{fee_pct}{r}+{fee_fixed}{r})-'
                'IF({support_mins}{r}="",0,{support_mins}{r})/60*'
                '{rate}{r}<=0),"Never at this price",'
                'ROUNDUP({dev_cost}{r}/({price}{r}-{other_cost}{r}-'
                '({price}{r}*{fee_pct}{r}+{fee_fixed}{r})-'
                'IF({support_mins}{r}="",0,{support_mins}{r})/60*'
                '{rate}{r}),0))'),
            note="How many copies pay for the design time. If that "
                 "number is bigger than the sales you expect, the "
                 "price is too low or the file is too ambitious."),
    ] + craft_tail() + [craft_check(needs="labor_cost")]

    return SheetSpec(
        name=NAME,
        title="Digital Product Calculator",
        subtitle="Design time spread over the sales you expect, and "
                 "the day it pays for itself.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).digital,
        help_text="All the cost is up front. Enter the hours "
                  "honestly and the sales you expect "
                  "conservatively, and the per sale cost falls out.",
        extra_notes=[
            "Writing a pattern up takes longer than making the "
            "thing. Twelve hours of design and sixteen of writing "
            "at $25 is $700 before a single sale - over 200 copies "
            "that is $3.50 each.",
            "A download that sells steadily is the best margin in "
            "handmade. One that sells twice is an expensive way to "
            "spend a fortnight. Expect the second, be delighted by "
            "the first.",
            "Support time is real. Two minutes a sale at $25 an "
            "hour is 83 cents - on a $6 pattern that is 14% of the "
            "price.",
        ],
    )
````

---

## 5. The selling registers

What happens to a product once it has a price: bundles and their lines, variations and custom orders, batch production, the sales ledger, promotions, price history, competitor prices, the waste log, development cost recovery, craft fairs, stockists and wholesale orders.

### `handmade_pricing_spreadsheet/sheets/bundles.py`

*239 lines · 1 function*

> `Bundles` - gift sets and multi-buys, priced so the discount is a choice.
>
> A bundle is three products in a box with a ribbon and a discount, and
> it is the easiest place in handmade to lose money without noticing:
> the discount comes off the retail price while every cost stays
> exactly where it was, plus a bigger box and the time to assemble it.
>
> This tab adds the components up from the Bundle Lines tab, adds the
> assembly time and the packaging, and shows you what the discount is
> really costing.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Bundles').

````python
"""`Bundles` - gift sets and multi-buys, priced so the discount is a choice.

A bundle is three products in a box with a ribbon and a discount, and
it is the easiest place in handmade to lose money without noticing:
the discount comes off the retail price while every cost stays
exactly where it was, plus a bigger box and the time to assemble it.

This tab adds the components up from the Bundle Lines tab, adds the
assembly time and the packaging, and shows you what the discount is
really costing.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, round_price,
                     status_cf)

NAME = "Bundles"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "BundlePrefix", digits=4, width=10),
        Col("name", "Bundle", 28, "text",
            note="'Cosy night in gift set', 'Starter crochet kit'."),
        Col("kind", "Type", 18, "text",
            validate=list_from("BundleTypeList")),
        Col("channel", "Channel", 16, "text",
            validate=list_from("ChannelList")),
        Col("pack_profile", "Packaging", 12, "text",
            validate=list_from("PackProfileList"),
            note="Bundles usually need the bigger box. Pick the "
                 "profile that actually goes out."),
        Col("assembly_mins", "Assembly minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Boxing, wrapping, tying, the card. Ten minutes a "
                 "set at $25 an hour is $4.17 - more than most "
                 "bundle discounts."),
        Col("discount", "Discount off retail", 13, "percent",
            validate=number_check(0, 0.9),
            note="What the buyer saves against buying the pieces "
                 "separately. Leave blank for no discount."),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("status", "Status", 13, "status",
            validate=list_from("StatusList"), cf=status_cf()),
        Col("season", "Season", 14, "text", validate=list_from("SeasonList")),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("items", "Items", 8, "int",
            formula=blank(
                "name",
                'SUMIFS([[Bundle Lines.qty]],[[Bundle Lines.bundle]],'
                '{id}{r})')),
        Col("lines", "Products", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Bundle Lines.bundle]],{id}{r})'),
            hidden=True),
        Col("parts_cost", "Components cost", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Bundle Lines.line_cost]],'
                '[[Bundle Lines.bundle]],{id}{r}),2)'),
            note="What the pieces cost to make, before the box and "
                 "the ribbon."),
        Col("parts_price", "Bought separately", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Bundle Lines.line_price]],'
                '[[Bundle Lines.bundle]],{id}{r}),2)'),
            note="What the same items would cost the buyer one at "
                 "a time. The discount comes off this."),
        Col("pack_cost", "Packaging", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({pack_profile}{r}="Budget",PackBudget,'
                'IF({pack_profile}{r}="Standard",PackStandard,'
                'IF({pack_profile}{r}="Gift",PackGift,'
                'IF({pack_profile}{r}="Premium",PackPremium,'
                'DefaultPackaging)))),2)')),
        Col("assembly_cost", "Assembly", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({assembly_mins}{r}="",0,{assembly_mins}{r})/60*'
                'LaborCostHour,2)')),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({parts_cost}{r}+{pack_cost}{r}+'
                '{assembly_cost}{r},2)'),
            note="Components already carry their own packaging "
                 "only if you listed them that way - check you are "
                 "not boxing twice."),
        Col("fee_pct", "Fee rate", 9, "percent",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Channels.pct_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 9, "money",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Channels.fixed_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("denom", "Denominator", 9, "number",
            formula=blank("name", '1-TargetRetailMargin-{fee_pct}{r}'),
            hidden=True),
        Col("suggested_raw", "Unrounded", 10, "money",
            formula=blank(
                "name",
                'IF({denom}{r}<=0,"",({total_cost}{r}+{fee_fixed}{r})/'
                '{denom}{r})'),
            hidden=True),
        Col("suggested", "Suggested price", 13, "money",
            formula=blank(
                "name",
                'IF({suggested_raw}{r}="","Check the settings",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="Priced from the bundle's own costs - ignore the "
                 "discount for a moment and see what it needs."),
        Col("discounted", "Discounted price", 13, "money",
            formula=blank(
                "name",
                'IF({parts_price}{r}=0,"",ROUND({parts_price}{r}*'
                '(1-IF({discount}{r}="",0,{discount}{r})),2))'),
            hidden=True),
        Col("price", "Price used", 11, "money",
            formula=blank(
                "name",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF({discounted}{r}<>"",{discounted}{r},'
                'IF(ISNUMBER({suggested}{r}),{suggested}{r},"")))'),
            note="Your price, otherwise the discounted total of the "
                 "pieces, otherwise what the maths suggests."),
        Col("fees", "Channel fees", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({price}{r}*{fee_pct}{r}+'
                '{fee_fixed}{r},2))')),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({price}{r}-{total_cost}{r}-'
                '{fees}{r},2))'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({price}{r}="",{price}{r}=0),"",{profit}{r}/'
                '{price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("saving", "Buyer saves", 11, "money",
            formula=blank(
                "name",
                'IF(OR({price}{r}="",{parts_price}{r}=0),"",'
                'ROUND({parts_price}{r}-{price}{r},2))'),
            note="The headline for the listing. Make sure you can "
                 "afford the headline."),
        Col("saving_pct", "Saving", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({saving}{r}="",{parts_price}{r}=0),"",'
                '{saving}{r}/{parts_price}{r})'),
            hidden=True),
        Col("parts_margin", "Separate margin", 12, "percent",
            formula=blank(
                "name",
                'IF({parts_price}{r}=0,"",({parts_price}{r}-'
                '{parts_cost}{r})/{parts_price}{r})'),
            hidden=True,
            note="The margin those items make when sold one at a "
                 "time. Compare it with the bundle margin - that "
                 "gap is what the discount cost you."),
        Col("margin_lost", "Margin given away", 12, "percent",
            formula=blank(
                "name",
                'IF(OR({margin}{r}="",{parts_margin}{r}=""),"",'
                '{parts_margin}{r}-{margin}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.15,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0.05,
                 "fmt": "good"}]),
        Col("sold", "Bundles sold", 11, "int",
            formula=blank(
                "name",
                'SUMIFS([[Sales.qty]],[[Sales.product]],{id}{r})')),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({items}{r}=0,"Add - no products in it yet",'
                'IF({price}{r}="","Check - no price yet",'
                'IF({profit}{r}<0,"Loss - the discount is too deep",'
                'IF({margin}{r}<MarginWarning,'
                '"Low margin - trim the discount",'
                'IF(AND({margin_lost}{r}<>"",{margin_lost}{r}>0.15),'
                '"Check - giving away a lot of margin",'
                'IF({assembly_mins}{r}="",'
                '"Add - how long does it take to box?","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Bundles & Gift Sets",
        subtitle="Multi-item sets, with the discount, the box and "
                 "the assembly time all counted.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).bundles,
        help_text="Build the contents on the Bundle Lines tab, then "
                  "set a discount here and watch what it does to "
                  "the margin.",
        extra_notes=[
            "A 15% bundle discount on products that make 45% "
            "margin leaves 35% - before the bigger box and ten "
            "minutes of ribbon. Bundles are a volume play, not a "
            "margin one.",
            "Sell the saving, not the discount. 'Save $9' reads "
            "better than '15% off' and costs you exactly the same.",
            "Check you are not paying for packaging twice. If each "
            "piece is individually wrapped inside the gift box, "
            "both costs are real - put the extra in the assembly "
            "line or pick a richer profile.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/bundle_lines.py`

*96 lines · 1 function*

> `Bundle Lines` - what is in each bundle.
>
> One row per product per bundle. The costs and the retail prices come
> straight from the catalog, so a bundle re-prices itself the moment a
> component does.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Bundle Lines').

````python
"""`Bundle Lines` - what is in each bundle.

One row per product per bundle. The costs and the retail prices come
straight from the catalog, so a bundle re-prices itself the moment a
component does.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, lookup_product

NAME = "Bundle Lines"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("bundle", "Bundle", 12, "id", validate=list_from("BundleIDs"),
            note="The bundle ID from the Bundles tab."),
        Col("bundle_name", "Bundle name", 24, "text",
            formula=blank(
                "bundle",
                'IFERROR(INDEX([[Bundles.name]],MATCH({bundle}{r},'
                '[[Bundles.id]],0)),"Not found")')),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs")),
        Col("product_name", "Product name", 26, "text",
            formula=lookup_product("name")),
        Col("qty", "Quantity", 10, "int",
            validate=number_check(1, 10000, decimal=False)),
        Col("swap", "Substitutable?", 12, "flag",
            validate=list_from("YesNoList"),
            note="Yes if a buyer can pick a colour or scent. "
                 "Substitutes should cost about the same or the "
                 "bundle margin moves with every order."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("unit_cost", "Cost each", 11, "money",
            formula=lookup_product("total_cost", missing="")),
        Col("unit_price", "Retail each", 11, "money",
            formula=lookup_product("price", missing="")),
        Col("line_cost", "Line cost", 11, "money",
            formula=blank(
                "product",
                'IF(OR({qty}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({qty}{r}*{unit_cost}{r},2))')),
        Col("line_price", "Line retail", 11, "money",
            formula=blank(
                "product",
                'IF(OR({qty}{r}="",{unit_price}{r}=""),"",'
                'ROUND({qty}{r}*{unit_price}{r},2))')),
        Col("seq", "Line", 7, "int",
            formula=blank(
                "bundle",
                'COUNTIFS({bundle_first}:{bundle}{r},{bundle}{r})'),
            hidden=True),
        Col("key", "Key", 14, "text",
            formula=blank("bundle", '{bundle}{r}&"|"&{seq}{r}'),
            hidden=True),
        Col("check", "Check", 24, "status",
            formula=(
                '=IF(AND({bundle}{r}="",{product}{r}=""),"",'
                'IF({bundle}{r}="","Add - which bundle?",'
                'IF({product}{r}="","Add - which product?",'
                'IF({product_name}{r}="Not found",'
                '"Check - no such product ID",'
                'IF({qty}{r}="","Add - how many?",'
                'IF({unit_price}{r}="",'
                '"Check - that product has no price","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Bundle Contents",
        subtitle="What goes in each set, costed from the catalog.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).bundle_lines,
        help_text="One row per product per bundle. Change a "
                  "component's price in the catalog and every "
                  "bundle that contains it follows.",
        extra_notes=[
            "Put the quantity in rather than listing the same "
            "product twice - the bundle totals count quantities, "
            "not rows.",
            "If a buyer can choose a substitute, make sure the "
            "alternatives cost about the same. A free upgrade "
            "chosen by every customer is a price cut.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/variations.py`

*243 lines · 1 function*

> `Variations` - sizes, colours, personalisation and rush orders.
>
> The same product in three sizes is three different costs, and
> "can you add a name on it?" is twenty minutes nobody charged for.
> This tab starts from a product in the catalog, applies the size
> factors, adds whatever extra material and time the variation takes,
> and prices it - including a surcharge for a turnaround you did not
> offer.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Variations').

````python
"""`Variations` - sizes, colours, personalisation and rush orders.

The same product in three sizes is three different costs, and
"can you add a name on it?" is twenty minutes nobody charged for.
This tab starts from a product in the catalog, applies the size
factors, adds whatever extra material and time the variation takes,
and prices it - including a surcharge for a turnaround you did not
offer.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, lookup_product,
                     round_price)

NAME = "Variations"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "VariationPrefix", digits=4, width=10),
        Col("product", "Base product", 12, "id",
            validate=list_from("ProductIDs"),
            note="The variation inherits this product's costs and "
                 "then adjusts them."),
        Col("name", "Variation", 26, "text",
            note="'Large, personalised', 'Rush order, navy'."),
        Col("kind", "What changes", 18, "text",
            validate=list_from("VariationTypeList")),
        Col("size", "Size", 14, "text", validate=list_from("SizeNames"),
            note="Applies the material and time factors from the "
                 "Lists tab to the base product."),
        Col("extra_material", "Extra materials", 13, "money",
            validate=number_check(0, 100000),
            note="Anything this version uses that the base does "
                 "not - a bigger clasp, a second colour, a gift "
                 "box insert."),
        Col("extra_mins", "Extra minutes", 12, "int",
            validate=number_check(0, 100000, decimal=False),
            note="Personalisation, colour changes, re-threading "
                 "the machine, an extra fitting."),
        Col("rush", "Turnaround", 14, "text",
            validate=list_from("RushNames"),
            note="Adds the surcharge from the Lists tab. Saying "
                 "yes to Friday should be worth something."),
        Col("personalisation", "Personalisation fee", 14, "money",
            validate=number_check(0, 100000),
            note="A flat charge you add on top, if you prefer that "
                 "to pricing the minutes."),
        Col("channel", "Channel", 15, "text",
            validate=list_from("ChannelList")),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("offered", "Offered?", 10, "flag",
            validate=list_from("YesNoList"),
            note="No for a variation you costed and decided "
                 "against - the workbook keeps the reasoning."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("base_name", "Base product name", 24, "text",
            formula=lookup_product("name"),
            hidden=True),
        Col("base_mat", "Base materials", 11, "money",
            formula=lookup_product("mat_cost", missing=""),
            hidden=True),
        Col("base_labour", "Base labour", 11, "money",
            formula=lookup_product("labour_adj", missing=""),
            hidden=True),
        Col("base_other", "Base other", 11, "money",
            formula=blank(
                "product",
                'IFERROR(INDEX([[Products.cost_base]],MATCH({product}{r},'
                '[[Products.id]],0))-INDEX([[Products.mat_cost]],'
                'MATCH({product}{r},[[Products.id]],0))-'
                'INDEX([[Products.labour_adj]],MATCH({product}{r},'
                '[[Products.id]],0)),0)'),
            hidden=True,
            note="Consumables, packaging and postage from the base "
                 "product - unchanged by the variation."),
        Col("base_minutes", "Base minutes", 11, "number1",
            formula=lookup_product("minutes", missing=""),
            hidden=True),
        Col("base_price", "Base price", 11, "money",
            formula=lookup_product("price", missing="")),
        Col("size_mat", "Size factor", 9, "number",
            formula=blank(
                "product",
                'IFERROR(INDEX(SizeMaterialFactors,MATCH({size}{r},'
                'SizeNames,0)),1)'),
            hidden=True),
        Col("size_time", "Time factor", 9, "number",
            formula=blank(
                "product",
                'IFERROR(INDEX(SizeTimeFactors,MATCH({size}{r},'
                'SizeNames,0)),1)'),
            hidden=True),
        Col("rush_rate", "Rush surcharge", 10, "percent",
            formula=blank(
                "product",
                'IFERROR(INDEX(RushFactors,MATCH({rush}{r},'
                'RushNames,0)),0)'),
            hidden=True),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "product",
                'ROUND(IF({base_mat}{r}="",0,{base_mat}{r})*'
                '{size_mat}{r}+IF({extra_material}{r}="",0,'
                '{extra_material}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "product",
                'ROUND(IF({base_minutes}{r}="",0,{base_minutes}{r})*'
                '{size_time}{r}+IF({extra_mins}{r}="",0,'
                '{extra_mins}{r}),1)')),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank(
                "product",
                'ROUND(IF({base_labour}{r}="",0,{base_labour}{r})*'
                '{size_time}{r}+IF({extra_mins}{r}="",0,'
                '{extra_mins}{r})/60*LaborCostHour,2)')),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "product",
                'ROUND({mat_cost}{r}+{labor_cost}{r}+'
                'IF({base_other}{r}="",0,{base_other}{r}),2)')),
        Col("fee_pct", "Fee rate", 9, "percent",
            formula=blank(
                "product",
                'IFERROR(INDEX([[Channels.pct_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 9, "money",
            formula=blank(
                "product",
                'IFERROR(INDEX([[Channels.fixed_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("denom", "Denominator", 9, "number",
            formula=blank("product", '1-TargetRetailMargin-{fee_pct}{r}'),
            hidden=True),
        Col("suggested_raw", "Unrounded", 10, "money",
            formula=blank(
                "product",
                'IF({denom}{r}<=0,"",(({total_cost}{r}+{fee_fixed}{r})/'
                '{denom}{r})*(1+{rush_rate}{r})+'
                'IF({personalisation}{r}="",0,{personalisation}{r}))'),
            hidden=True),
        Col("suggested", "Suggested price", 13, "money",
            formula=blank(
                "product",
                'IF({suggested_raw}{r}="","Check the settings",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="The base maths, stretched by the size, plus the "
                 "rush surcharge and any flat personalisation fee."),
        Col("price", "Price used", 11, "money",
            formula=blank(
                "product",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF(ISNUMBER({suggested}{r}),{suggested}{r},""))')),
        Col("uplift", "Versus the base", 11, "money",
            formula=blank(
                "product",
                'IF(OR({price}{r}="",{base_price}{r}=""),"",'
                'ROUND({price}{r}-{base_price}{r},2))'),
            note="What you are charging extra. Put this in the "
                 "listing as the upgrade price."),
        Col("fees", "Channel fees", 10, "money",
            formula=blank(
                "product",
                'IF({price}{r}="","",ROUND({price}{r}*{fee_pct}{r}+'
                '{fee_fixed}{r},2))'),
            hidden=True),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                "product",
                'IF({price}{r}="","",ROUND({price}{r}-{total_cost}{r}-'
                '{fees}{r},2))'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "product",
                'IF(OR({price}{r}="",{price}{r}=0),"",{profit}{r}/'
                '{price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("wage_hour", "You earn an hour", 12, "money",
            formula=blank(
                "product",
                'IF(OR({minutes}{r}=0,{minutes}{r}="",{profit}{r}=""),"",'
                'ROUND(({profit}{r}+{labor_cost}{r})/'
                '({minutes}{r}/60),2))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MinimumWage",
                 "fmt": "bad"}]),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "product",
                'IF({base_name}{r}="Not found","Check - no such product ID",'
                'IF({price}{r}="","Check - no price yet",'
                'IF({profit}{r}<0,"Loss - this version costs you money",'
                'IF(AND({wage_hour}{r}<>"",{wage_hour}{r}<MinimumWage),'
                '"Below - pays under minimum wage",'
                'IF(AND({extra_mins}{r}>0,{uplift}{r}<=0),'
                '"Check - extra work, no extra charge",'
                'IF({margin}{r}<MarginWarning,'
                '"Low margin - re-price it","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Variations & Custom Orders",
        subtitle="Sizes, personalisation and rush work, priced from "
                 "the base product.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).variations,
        help_text="Pick a base product, say what changes, and the "
                  "variation prices itself from the size factors "
                  "and the extra minutes.",
        extra_notes=[
            "Personalisation is not free. Twenty minutes of "
            "embroidery at $25 an hour is $8.33, and a $3 "
            "'personalisation fee' is you paying the customer to "
            "give you more work.",
            "Rush orders should cost more. The surcharges on the "
            "Lists tab are a starting point - the real question is "
            "what dropping everything else is worth to you.",
            "Costing a variation and deciding against it is a win. "
            "Set 'Offered' to No and the reasoning stays here for "
            "the next time somebody asks.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/batches.py`

*199 lines · 1 function*

> `Batches` - what a production run actually cost, against what you estimated.
>
> Estimates are where pricing starts; batches are where it gets tested.
> Log a run - how many you planned, what the materials really came to,
> how long it really took, how many came out wrong - and this tab tells
> you the true cost of a good unit and how far your estimate was out.
>
> The seconds are not free. A batch of twenty with three failures is a
> batch of seventeen, and all twenty pieces of wax were paid for.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Batches').

````python
"""`Batches` - what a production run actually cost, against what you estimated.

Estimates are where pricing starts; batches are where it gets tested.
Log a run - how many you planned, what the materials really came to,
how long it really took, how many came out wrong - and this tab tells
you the true cost of a good unit and how far your estimate was out.

The seconds are not free. A batch of twenty with three failures is a
batch of seventeen, and all twenty pieces of wax were paid for.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, lookup_product,
                     status_cf)

NAME = "Batches"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("product", "BatchPrefix", digits=4, width=10),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs")),
        Col("product_name", "Product name", 24, "text",
            formula=lookup_product("name")),
        Col("started", "Started", 11, "date", validate=date_check()),
        Col("finished", "Finished", 11, "date", validate=date_check()),
        Col("planned", "Planned", 9, "int",
            validate=number_check(0, 100000, decimal=False)),
        Col("made", "Made", 8, "int",
            validate=number_check(0, 100000, decimal=False)),
        Col("seconds", "Seconds / failures", 13, "int",
            validate=number_check(0, 100000, decimal=False),
            note="Pieces that cannot be sold at full price. Their "
                 "cost belongs to the ones that can."),
        Col("sold_seconds", "Seconds sold for", 13, "money",
            validate=number_check(0, 1000000),
            note="Total recovered from selling the seconds as "
                 "seconds. Credited back to the batch."),
        Col("material_actual", "Materials actually", 14, "money",
            validate=number_check(0, 1000000),
            note="What the run really consumed. Leave blank to use "
                 "the estimate from the catalog."),
        Col("hours_actual", "Hours actually", 12, "number",
            validate=number_check(0, 10000),
            note="Bench time for the whole run, setup included."),
        Col("extra_cost", "Other costs", 11, "money",
            validate=number_check(0, 1000000),
            note="A courier, a rush delivery of supplies, studio "
                 "hire for the day."),
        Col("status", "Status", 15, "status",
            validate=list_from("BatchStatusList"), cf=status_cf()),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("good", "Good units", 10, "int",
            formula=blank(
                "product",
                'MAX(0,IF({made}{r}="",0,{made}{r})-'
                'IF({seconds}{r}="",0,{seconds}{r}))')),
        Col("defect_rate", "Failure rate", 10, "percent",
            formula=blank(
                "product",
                'IF(OR({made}{r}="",{made}{r}=0),"",'
                'IF({seconds}{r}="",0,{seconds}{r})/{made}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": "=DefaultDefect",
                 "fmt": "warn"}]),
        Col("est_mat_unit", "Materials estimated", 13, "money",
            formula=lookup_product("mat_cost", missing=""),
            hidden=True),
        Col("est_minutes", "Minutes estimated", 12, "number1",
            formula=lookup_product("minutes", missing=""),
            hidden=True),
        Col("mat_total", "Materials cost", 12, "money",
            formula=blank(
                "product",
                'ROUND(IF({material_actual}{r}<>"",{material_actual}{r},'
                'IF({est_mat_unit}{r}="",0,{est_mat_unit}{r})*'
                'IF({made}{r}="",0,{made}{r})),2)')),
        Col("labour_total", "Labour cost", 11, "money",
            formula=blank(
                "product",
                'ROUND(IF({hours_actual}{r}<>"",{hours_actual}{r},'
                'IF({est_minutes}{r}="",0,{est_minutes}{r})*'
                'IF({made}{r}="",0,{made}{r})/60)*LaborCostHour,2)')),
        Col("batch_cost", "Batch cost", 11, "money",
            formula=blank(
                "product",
                'ROUND({mat_total}{r}+{labour_total}{r}+'
                'IF({extra_cost}{r}="",0,{extra_cost}{r})-'
                'IF({sold_seconds}{r}="",0,{sold_seconds}{r}),2)'),
            note="Everything the run cost, less whatever the "
                 "seconds brought back in."),
        Col("cost_good", "Cost a good unit", 12, "money",
            formula=blank(
                "product",
                'IF({good}{r}=0,"",ROUND({batch_cost}{r}/{good}{r},2))'),
            note="The honest cost of a sellable piece - the "
                 "failures are carried by the survivors."),
        Col("est_unit", "Estimated a unit", 12, "money",
            formula=blank(
                "product",
                'IFERROR(INDEX([[Products.cost_base]],MATCH({product}{r},'
                '[[Products.id]],0)),"")'),
            hidden=True),
        Col("variance", "Versus estimate", 11, "money",
            formula=blank(
                "product",
                'IF(OR({cost_good}{r}="",{est_unit}{r}=""),"",'
                'ROUND({cost_good}{r}-{est_unit}{r},2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}],
            note="Positive means the real thing costs more than "
                 "the catalog thinks. Two batches saying the same "
                 "means the catalog is wrong."),
        Col("variance_pct", "Out by", 9, "percent",
            formula=blank(
                "product",
                'IF(OR({variance}{r}="",{est_unit}{r}="",'
                '{est_unit}{r}=0),"",{variance}{r}/{est_unit}{r})'),
            hidden=True),
        Col("mins_actual", "Minutes a unit", 11, "number1",
            formula=blank(
                "product",
                'IF(OR({hours_actual}{r}="",{made}{r}="",{made}{r}=0),"",'
                'ROUND({hours_actual}{r}*60/{made}{r},1))')),
        Col("time_variance", "Minutes out by", 12, "number1",
            formula=blank(
                "product",
                'IF(OR({mins_actual}{r}="",{est_minutes}{r}=""),"",'
                'ROUND({mins_actual}{r}-{est_minutes}{r},1))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="The number that re-prices a whole product line. "
                 "If every batch takes longer than the estimate, "
                 "the estimate is a wish."),
        Col("value", "Batch retail value", 13, "money",
            formula=blank(
                "product",
                'IFERROR(ROUND({good}{r}*INDEX([[Products.price]],'
                'MATCH({product}{r},[[Products.id]],0)),2),"")')),
        Col("batch_profit", "Profit if it all sells", 14, "money",
            formula=blank(
                "product",
                'IF({value}{r}="","",ROUND({value}{r}-{batch_cost}{r},2))'),
            cf=loss_cf()),
        Col("days", "Days taken", 10, "int",
            formula=blank(
                "product",
                'IF(OR({started}{r}="",{finished}{r}=""),"",'
                '{finished}{r}-{started}{r})'),
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "product",
                'IF({made}{r}="","Planned - nothing made yet",'
                'IF({good}{r}=0,"Check - no good units",'
                'IF(AND({variance_pct}{r}<>"",{variance_pct}{r}>0.15),'
                '"Check - costing 15% more than the estimate",'
                'IF(AND({defect_rate}{r}<>"",'
                '{defect_rate}{r}>DefaultDefect*2),'
                '"Check - high failure rate",'
                'IF(AND({time_variance}{r}<>"",{time_variance}{r}>5),'
                '"Check - taking longer than estimated",'
                'IF({batch_profit}{r}<0,"Loss - this run lost money",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Batch Production",
        subtitle="What a run really cost, what it really took, and "
                 "how far the estimate was out.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).batches,
        help_text="Log the run as it happens. Leave the actuals "
                  "blank and the batch is costed from the catalog "
                  "estimate instead.",
        extra_notes=[
            "This is the tab that makes the rest of the workbook "
            "honest. Estimates drift; batches do not lie.",
            "Failures are paid for. Twenty poured candles with "
            "three sink holes is seventeen sellable candles and "
            "twenty jars of wax - the cost a good unit column does "
            "that arithmetic for you.",
            "Two batches in a row costing more than the estimate "
            "is not bad luck. Go and update the bill of materials "
            "or the labour minutes.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/sales.py`

*261 lines · 2 functions*

> `Sales` - what really happened, order by order.
>
> Estimated profit is a plan. This is the tab where the plan meets
> discounts, refunds, a postage label that cost more than you charged
> and a fee that was not quite what the fee page said.
>
> Costs come from the catalog, with the actual postage swapped in for
> the estimate, so nothing is counted twice. Type the fee the
> marketplace really took whenever you have the statement in front of
> you - the Channels tab compares it with your settings.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the order book |

**Module constants:** `NAME` ('Sales').

````python
"""`Sales` - what really happened, order by order.

Estimated profit is a plan. This is the tab where the plan meets
discounts, refunds, a postage label that cost more than you charged
and a fee that was not quite what the fee page said.

Costs come from the catalog, with the actual postage swapped in for
the estimate, so nothing is counted twice. Type the fee the
marketplace really took whenever you have the statement in front of
you - the Channels tab compares it with your settings.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, lookup_product,
                     month_col, status_cf)

NAME = "Sales"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "SalePrefix", digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs"),
            note="A product ID, or a bundle ID - both work, the "
                 "costs come from whichever tab the ID lives on."),
        Col("qty", "Qty", 7, "int",
            validate=number_check(1, 10000, decimal=False)),
        Col("channel", "Channel", 16, "text",
            validate=list_from("ChannelList")),
        Col("unit_price", "Price each", 11, "money",
            validate=number_check(0, 1000000),
            note="What you actually charged, before any discount. "
                 "Blank uses the catalog price."),
        Col("discount", "Discount given", 12, "money",
            validate=number_check(0, 1000000),
            note="In money, for the whole order. A 15% code on a "
                 "$28 order is 4.20."),
        Col("promo", "Promo code", 12, "text",
            validate=list_from("PromoCodes")),
        Col("ship_charged", "Postage charged", 12, "money",
            validate=number_check(0, 10000)),
        Col("ship_paid", "Postage paid", 11, "money",
            validate=number_check(0, 10000),
            note="What the label actually cost. The difference is "
                 "yours either way."),
        Col("fees_input", "Fees actually taken", 14, "money",
            validate=number_check(0, 1000000),
            note="From the marketplace statement. Leave blank and "
                 "the workbook estimates it from the Channels tab."),
        Col("refund", "Refunded", 10, "money",
            validate=number_check(0, 1000000)),
        Col("status", "Status", 14, "status",
            validate=list_from("SaleStatusList"), cf=status_cf()),
        Col("customer", "Customer", 18, "text"),
        Col("event", "Event", 16, "text", validate=list_from("EventNames"),
            note="For craft fair sales - ties the order to the "
                 "event so you can see if the stall paid."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("product_name", "Product name", 24, "text",
            formula=lookup_product("name", missing="Bundle or deleted")),
        month_col("date"),
        Col("catalog_price", "Catalog price", 11, "money",
            formula=lookup_product("price", missing=""),
            hidden=True),
        Col("price_used", "Price used", 11, "money",
            formula=blank(
                "date",
                'IF({unit_price}{r}<>"",{unit_price}{r},'
                'IF({catalog_price}{r}="",0,{catalog_price}{r}))'),
            hidden=True),
        Col("gross", "Gross", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({qty}{r}="",0,{qty}{r})*{price_used}{r},2)'),
            hidden=True),
        Col("revenue", "Revenue", 11, "money",
            formula=blank(
                "date",
                'ROUND({gross}{r}-IF({discount}{r}="",0,{discount}{r})-'
                'IF({refund}{r}="",0,{refund}{r}),2)'),
            note="What you actually took for the items, after "
                 "discounts and refunds. Postage is kept separate."),
        Col("unit_cost_total", "Catalog cost", 11, "money",
            formula=lookup_product("total_cost", missing=""),
            hidden=True),
        Col("unit_ship", "Catalog postage", 11, "money",
            formula=lookup_product("ship_cost", missing=""),
            hidden=True),
        Col("cogs", "Cost of the goods", 13, "money",
            formula=blank(
                "date",
                'ROUND(IF({qty}{r}="",0,{qty}{r})*'
                'MAX(0,IF({unit_cost_total}{r}="",0,{unit_cost_total}{r})-'
                'IF({unit_ship}{r}="",0,{unit_ship}{r})),2)'),
            note="Materials, labour, packaging and overhead from "
                 "the catalog. The estimated postage is taken out "
                 "because the real postage is on this row."),
        Col("ship_absorbed", "Postage you absorbed", 14, "money",
            formula=blank(
                "date",
                'ROUND(MAX(0,IF({ship_paid}{r}="",0,{ship_paid}{r})-'
                'IF({ship_charged}{r}="",0,{ship_charged}{r})),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("fee_pct", "Fee rate", 9, "percent",
            formula=blank(
                "date",
                'IFERROR(INDEX([[Channels.pct_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 9, "money",
            formula=blank(
                "date",
                'IFERROR(INDEX([[Channels.fixed_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_ship", "Fees on postage?", 10, "flag",
            formula=blank(
                "date",
                'IFERROR(INDEX([[Channels.fees_on_shipping]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),"No")'),
            hidden=True),
        Col("fees_est", "Fees estimated", 12, "money",
            formula=blank(
                "date",
                'ROUND(({revenue}{r}+IF({fee_ship}{r}="Yes",'
                'IF({ship_charged}{r}="",0,{ship_charged}{r}),0))*'
                '{fee_pct}{r}+{fee_fixed}{r},2)'),
            hidden=True),
        Col("fees", "Fees", 11, "money",
            formula=blank(
                "date",
                'IF({fees_input}{r}<>"",{fees_input}{r},'
                '{fees_est}{r})')),
        Col("fee_gap", "Fees versus estimate", 13, "money",
            formula=blank(
                "date",
                'IF({fees_input}{r}="","",ROUND({fees_input}{r}-'
                '{fees_est}{r},2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.25,
                 "fmt": "bad"}],
            hidden=True,
            note="Real fees consistently above the estimate means "
                 "the fee settings need updating."),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                "date",
                'ROUND({revenue}{r}-{cogs}{r}-{fees}{r}-'
                '{ship_absorbed}{r},2)'),
            cf=loss_cf(),
            note="Revenue less the goods, the fees and whatever "
                 "postage you swallowed."),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "date",
                'IF(OR({revenue}{r}="",{revenue}{r}=0),"",{profit}{r}/'
                '{revenue}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("planned_margin", "Planned margin", 12, "percent",
            formula=lookup_product("margin", missing=""),
            hidden=True),
        Col("margin_gap", "Versus plan", 10, "percent",
            formula=blank(
                "date",
                'IF(OR({margin}{r}="",{planned_margin}{r}=""),"",'
                '{margin}{r}-{planned_margin}{r})'),
            hidden=True),
        Col("hours", "Hours of making", 12, "number",
            formula=blank(
                "date",
                'IFERROR(ROUND(IF({qty}{r}="",0,{qty}{r})*'
                'INDEX([[Products.minutes]],MATCH({product}{r},'
                '[[Products.id]],0))/60,2),"")'),
            hidden=True),
        Col("profit_hour", "Profit an hour", 11, "money",
            formula=blank(
                "date",
                'IF(OR({hours}{r}="",{hours}{r}=0),"",'
                'ROUND({profit}{r}/{hours}{r},2))'),
            hidden=True),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "date",
                'IF({product}{r}="","Add - which product?",'
                'IF({profit}{r}<0,"Loss - this order cost you money",'
                'IF({status}{r}="Refunded","Refunded - check why",'
                'IF(AND({fee_gap}{r}<>"",{fee_gap}{r}>0.25),'
                '"Check - fees higher than your settings",'
                'IF({ship_absorbed}{r}>1,'
                '"Check - postage is eating the margin",'
                'IF(AND({margin}{r}<>"",{margin}{r}<MarginWarning),'
                '"Low margin on this order","OK")))))) '),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Sales & Actual Profit",
        subtitle="Every order, with the real postage, the real fees "
                 "and the real margin.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).sales,
        help_text="Enter the order. Leave the price blank to use "
                  "the catalog price, and type the real fees "
                  "whenever you have the statement.",
        extra_notes=[
            "Estimated and actual will not match, and that is the "
            "point. The gap is where a price quietly stopped "
            "working.",
            "Postage is a separate line because it behaves "
            "differently: the buyer usually pays it, the fee is "
            "usually charged on it, and when you offer free "
            "shipping it comes straight out of your margin.",
            "Record craft fair sales against the event. A stall "
            "that took $380 and cost $95 plus two days is a "
            "decision you can only make with the numbers in front "
            "of you.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the order book."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Sales summary", last)
    sheet.table_head(row, 0, ["Orders", "Items", "Revenue", "Cost of goods",
                              "Fees", "Postage absorbed", "Profit",
                              "Margin", "Loss making orders"],
                     [10, 10, 14, 15, 12, 17, 13, 10, 18])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Sales.date]],"<>")',
        '=SUM([[Sales.qty]])',
        '=ROUND(SUM([[Sales.revenue]]),2)',
        '=ROUND(SUM([[Sales.cogs]]),2)',
        '=ROUND(SUM([[Sales.fees]]),2)',
        '=ROUND(SUM([[Sales.ship_absorbed]]),2)',
        '=ROUND(SUM([[Sales.profit]]),2)',
        '=IFERROR(ROUND(SUM([[Sales.profit]])/SUM([[Sales.revenue]]),4),"")',
        '=COUNTIFS([[Sales.profit]],"<0")',
    ], ["int", "int", "money", "money", "money", "money", "money",
        "percent", "int"])
````

### `handmade_pricing_spreadsheet/sheets/promotions.py`

*175 lines · 1 function*

> `Promotions` - what a sale costs, and how much more you have to sell.
>
> A 20% discount on a 45% margin product does not cost you 20%. It
> costs you 44% of your profit, and you need to sell 80% more units to
> stand still. That arithmetic is on this tab, next to what each code
> actually did.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Promotions').

````python
"""`Promotions` - what a sale costs, and how much more you have to sell.

A 20% discount on a 45% margin product does not cost you 20%. It
costs you 44% of your profit, and you need to sell 80% more units to
stand still. That arithmetic is on this tab, next to what each code
actually did.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, loss_cf, status_cf

NAME = "Promotions"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("code", "Code", 14, "id",
            note="The code buyers type, or a short name for a sale "
                 "with no code. Sales rows pick from these."),
        Col("name", "Promotion", 26, "text"),
        Col("kind", "Type", 18, "text", validate=list_from("PromoTypeList")),
        Col("value", "Value", 10, "number",
            validate=number_check(0, 100000),
            note="A percentage as a decimal (0.15 for 15% off), or "
                 "an amount for 'Amount off'."),
        Col("scope", "Applies to", 18, "text",
            note="'Everything', a category, or a product ID. Free "
                 "text - it is for your own records."),
        Col("start", "Starts", 11, "date", validate=date_check()),
        Col("ends", "Ends", 11, "date", validate=date_check()),
        Col("min_spend", "Minimum spend", 12, "money",
            validate=number_check(0, 100000)),
        Col("limit", "Use limit", 10, "int",
            validate=number_check(0, 1000000, decimal=False)),
        Col("base_margin", "Normal margin", 12, "percent",
            validate=number_check(0, 0.95),
            note="The margin these products normally make. Blank "
                 "uses the target margin on Settings."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("margin_used", "Margin used", 11, "percent",
            formula=blank(
                "code",
                'IF({base_margin}{r}<>"",{base_margin}{r},'
                'TargetRetailMargin)'),
            hidden=True),
        Col("discount_pct", "Discount", 10, "percent",
            formula=blank(
                "code",
                'IF({kind}{r}="Percent off",{value}{r},'
                'IF({kind}{r}="Free shipping","",'
                'IF(OR({min_spend}{r}="",{min_spend}{r}=0),"",'
                '{value}{r}/{min_spend}{r})))'),
            hidden=True,
            note="An amount off is turned into a percentage using "
                 "the minimum spend, which is the only honest way "
                 "to compare the two."),
        Col("margin_after", "Margin after it", 12, "percent",
            formula=blank(
                "code",
                'IF({discount_pct}{r}="","",'
                '({margin_used}{r}-{discount_pct}{r})/'
                '(1-{discount_pct}{r}))'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}],
            note="What is left of the margin once the discount is "
                 "off both the price and the profit."),
        Col("profit_lost", "Profit given away", 12, "percent",
            formula=blank(
                "code",
                'IF(OR({discount_pct}{r}="",{margin_used}{r}=0),"",'
                '{discount_pct}{r}/{margin_used}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.3,
                 "fmt": "warn"}],
            note="The share of your profit the discount hands "
                 "over. 20% off a 45% margin gives away 44% of the "
                 "profit on every order."),
        Col("uplift_needed", "Extra sales needed", 14, "percent",
            formula=blank(
                "code",
                'IF({discount_pct}{r}="","Never - it loses money",'
                'IF({margin_used}{r}-{discount_pct}{r}<=0,'
                '"Never - it loses money",'
                '{discount_pct}{r}/({margin_used}{r}-'
                '{discount_pct}{r})))'),
            note="How much more you have to sell to make the same "
                 "profit as before the sale. This is the number "
                 "that stops a 30% off weekend."),
        Col("orders", "Orders", 9, "int",
            formula=blank(
                "code",
                'COUNTIFS([[Sales.promo]],{code}{r})')),
        Col("revenue", "Revenue", 11, "money",
            formula=blank(
                "code",
                'ROUND(SUMIFS([[Sales.revenue]],[[Sales.promo]],'
                '{code}{r}),2)')),
        Col("given", "Discount given", 12, "money",
            formula=blank(
                "code",
                'ROUND(SUMIFS([[Sales.discount]],[[Sales.promo]],'
                '{code}{r}),2)')),
        Col("profit", "Profit made", 11, "money",
            formula=blank(
                "code",
                'ROUND(SUMIFS([[Sales.profit]],[[Sales.promo]],'
                '{code}{r}),2)'),
            cf=loss_cf()),
        Col("actual_margin", "Actual margin", 11, "percent",
            formula=blank(
                "code",
                'IF(OR({revenue}{r}="",{revenue}{r}=0),"",'
                '{profit}{r}/{revenue}{r})')),
        Col("avg_order", "Average order", 11, "money",
            formula=blank(
                "code",
                'IF({orders}{r}=0,"",ROUND({revenue}{r}/'
                '{orders}{r},2))'),
            hidden=True),
        Col("status", "Status", 13, "status",
            formula=blank(
                "code",
                'IF(AND({start}{r}<>"",{start}{r}>TodayDate),"Scheduled",'
                'IF(AND({ends}{r}<>"",{ends}{r}<TodayDate),"Expired",'
                'IF(AND({limit}{r}<>"",{orders}{r}>={limit}{r}),'
                '"Used up","Active")))'),
            cf=status_cf()),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "code",
                'IF({value}{r}="","Add - how much off?",'
                'IF(AND({margin_after}{r}<>"",{margin_after}{r}<0),'
                '"Loss - this discount sells below cost",'
                'IF(AND({profit_lost}{r}<>"",{profit_lost}{r}>0.5),'
                '"Check - over half the profit given away",'
                'IF({status}{r}="Expired","Ended - review what it did",'
                'IF(AND({orders}{r}=0,{status}{r}="Active"),'
                '"Waiting - no orders yet","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Discounts & Promotions",
        subtitle="What each sale costs in margin, and how many "
                 "extra orders it has to bring.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).promotions,
        help_text="Set the discount and your normal margin, and "
                  "the workbook shows what is left and how much "
                  "more you would have to sell.",
        extra_notes=[
            "Discounting is the fastest way to turn a good margin "
            "into a busy fortnight with nothing to show for it. "
            "The 'extra sales needed' column is the reality check.",
            "Free shipping is a discount too. Price it as one: "
            "work out what you absorb and treat it as money off.",
            "A code that brings new buyers who come back is worth "
            "more than the margin it cost. Keep notes here - the "
            "numbers alone will not tell you that.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/price_history.py`

*156 lines · 1 function*

> `Price History` - every price change, and whether it worked.
>
> Raising a price feels like a risk until you can see what happened
> last time you did it. Log the change, the reason and what the
> product cost that week, and this tab keeps the story: what the
> margin was before, what it became, and how long the price has stood
> since.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Price History').

````python
"""`Price History` - every price change, and whether it worked.

Raising a price feels like a risk until you can see what happened
last time you did it. Log the change, the reason and what the
product cost that week, and this tab keeps the story: what the
margin was before, what it became, and how long the price has stood
since.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, loss_cf, lookup_product, month_col

NAME = "Price History"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs")),
        Col("product_name", "Product name", 24, "text",
            formula=lookup_product("name")),
        Col("old_price", "Old price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("new_price", "New price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("cost_then", "Cost at the time", 13, "money",
            validate=number_check(0, 1000000),
            note="What the product cost to make that week. Blank "
                 "uses today's cost from the catalog, which is "
                 "fine for a recent change and misleading for an "
                 "old one."),
        Col("reason", "Reason", 22, "text",
            validate=list_from("PriceReasonList")),
        Col("channel", "Channel", 15, "text",
            validate=list_from("ChannelList")),
        Col("sold_before", "Sold in the 30 days before", 16, "int",
            validate=number_check(0, 1000000, decimal=False),
            note="Optional, but this is what turns a price change "
                 "into evidence."),
        Col("sold_after", "Sold in the 30 days after", 16, "int",
            validate=number_check(0, 1000000, decimal=False)),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        month_col("date"),
        Col("change", "Change", 10, "money",
            formula=blank(
                "date",
                'IF(OR({old_price}{r}="",{new_price}{r}=""),"",'
                'ROUND({new_price}{r}-{old_price}{r},2))'),
            cf=loss_cf()),
        Col("change_pct", "Change", 9, "percent",
            formula=blank(
                "date",
                'IF(OR({old_price}{r}="",{old_price}{r}=0,'
                '{new_price}{r}=""),"",{new_price}{r}/'
                '{old_price}{r}-1)')),
        Col("cost_used", "Cost used", 11, "money",
            formula=blank(
                "date",
                'IF({cost_then}{r}<>"",{cost_then}{r},'
                'IFERROR(INDEX([[Products.total_cost]],MATCH({product}{r},'
                '[[Products.id]],0)),""))'),
            hidden=True),
        Col("margin_before", "Margin before", 12, "percent",
            formula=blank(
                "date",
                'IF(OR({old_price}{r}="",{old_price}{r}=0,'
                '{cost_used}{r}=""),"",({old_price}{r}-'
                '{cost_used}{r})/{old_price}{r})')),
        Col("margin_after", "Margin after", 12, "percent",
            formula=blank(
                "date",
                'IF(OR({new_price}{r}="",{new_price}{r}=0,'
                '{cost_used}{r}=""),"",({new_price}{r}-'
                '{cost_used}{r})/{new_price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("margin_move", "Margin moved", 11, "percent",
            formula=blank(
                "date",
                'IF(OR({margin_before}{r}="",{margin_after}{r}=""),"",'
                '{margin_after}{r}-{margin_before}{r})'),
            hidden=True),
        Col("volume_move", "Sales moved", 10, "percent",
            formula=blank(
                "date",
                'IF(OR({sold_before}{r}="",{sold_before}{r}=0,'
                '{sold_after}{r}=""),"",{sold_after}{r}/'
                '{sold_before}{r}-1)'),
            note="Did the units fall? A 15% rise that cost you 5% "
                 "of the volume made you money."),
        Col("revenue_move", "Revenue effect", 12, "percent",
            formula=blank(
                "date",
                'IF(OR({volume_move}{r}="",{change_pct}{r}=""),"",'
                '(1+{change_pct}{r})*(1+{volume_move}{r})-1)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"}],
            note="Price change and volume change together. "
                 "Positive means the change was worth making."),
        Col("seq", "Change number", 9, "int",
            formula=blank(
                "product",
                'COUNTIFS({product_first}:{product}{r},{product}{r})'),
            hidden=True),
        Col("days_since", "Days since", 10, "int",
            formula=blank(
                "date", 'TodayDate-{date}{r}'),
            hidden=True),
        Col("verdict", "Verdict", 26, "status",
            formula=blank(
                "date",
                'IF({change}{r}="","Add - old and new price",'
                'IF({margin_after}{r}<0,"Loss - priced below cost",'
                'IF({revenue_move}{r}="","Watch - no volume data yet",'
                'IF({revenue_move}{r}>0.02,"Worked - revenue up",'
                'IF({revenue_move}{r}<-0.02,'
                '"Hurt - revenue down","Neutral - no real change")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Price History",
        subtitle="Every change, the reason for it, and what it did "
                 "to the margin and the volume.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).price_history,
        help_text="Log the change on the day you make it. The "
                  "catalog uses the latest date here to tell you "
                  "when a price is going stale.",
        extra_notes=[
            "Record the units sold before and after and you stop "
            "guessing about price sensitivity. Most makers "
            "discover their buyers never noticed.",
            "A price rise that loses 10% of the volume on a 45% "
            "margin product still leaves you better off. The "
            "revenue effect column does that sum.",
            "Put the reason in. 'Yarn went up 18%' is the "
            "difference between a price rise you can explain and "
            "one that feels like greed six months later.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/competitors.py`

*144 lines · 1 function*

> `Competitors` - what the market charges, and where you sit in it.
>
> Competitor research is not about copying a price. It is about
> knowing whether you are the cheap option, the middle one or the
> expensive one - and making that a decision instead of an accident.
>
> Log what comparable pieces sell for, including their postage,
> because a $24 item with free shipping beats a $21 one with $5
> postage in every buyer's head.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Competitors').

````python
"""`Competitors` - what the market charges, and where you sit in it.

Competitor research is not about copying a price. It is about
knowing whether you are the cheap option, the middle one or the
expensive one - and making that a decision instead of an accident.

Log what comparable pieces sell for, including their postage,
because a $24 item with free shipping beats a $21 one with $5
postage in every buyer's head.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, lookup_product

NAME = "Competitors"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Your product", 12, "id",
            validate=list_from("ProductIDs"),
            note="Which of your products this one competes with."),
        Col("product_name", "Your product name", 24, "text",
            formula=lookup_product("name")),
        Col("competitor", "Their item", 28, "text"),
        Col("shop", "Shop", 20, "text"),
        Col("channel", "Where", 15, "text",
            validate=list_from("ChannelList")),
        Col("price", "Their price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("shipping", "Their postage", 11, "money",
            validate=number_check(0, 10000),
            note="0 for free shipping - which usually means it is "
                 "in the price."),
        Col("quality", "Quality", 16, "text",
            validate=list_from("QualityList"),
            note="Honestly. 'Above mine' is useful information, "
                 "not an insult."),
        Col("handmade", "Really handmade?", 13, "flag",
            validate=list_from("YesNoNAList"),
            note="A resold import is not your competition, "
                 "whatever the listing says."),
        Col("reviews", "Reviews", 9, "int",
            validate=number_check(0, 10000000, decimal=False),
            note="A rough proxy for how much they sell."),
        Col("checked", "Checked on", 11, "date", validate=date_check()),
        Col("url", "Link", 24, "text"),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("their_total", "Their total", 11, "money",
            formula=blank(
                "competitor",
                'ROUND(IF({price}{r}="",0,{price}{r})+'
                'IF({shipping}{r}="",0,{shipping}{r}),2)'),
            note="Price plus postage: what the buyer actually "
                 "pays, which is the only number they compare."),
        Col("your_price", "Your price", 11, "money",
            formula=lookup_product("price", missing=""),
            hidden=True),
        Col("your_ship", "Your postage", 11, "money",
            formula=lookup_product("ship_charge", missing=""),
            hidden=True),
        Col("your_total", "Your total", 11, "money",
            formula=blank(
                "competitor",
                'IF({your_price}{r}="","",ROUND({your_price}{r}+'
                'IF({your_ship}{r}="",0,{your_ship}{r}),2))')),
        Col("gap", "Difference", 11, "money",
            formula=blank(
                "competitor",
                'IF(OR({your_total}{r}="",{their_total}{r}=0),"",'
                'ROUND({your_total}{r}-{their_total}{r},2))')),
        Col("gap_pct", "You are", 10, "percent",
            formula=blank(
                "competitor",
                'IF(OR({your_total}{r}="",{their_total}{r}=0),"",'
                '{your_total}{r}/{their_total}{r}-1)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.35,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": -0.25,
                 "fmt": "bad"}],
            note="Well below the market is not a bargain, it is "
                 "an unpaid wage. Well above it needs better "
                 "photographs and a reason."),
        Col("position", "Position", 16, "status",
            formula=blank(
                "competitor",
                'IF({gap_pct}{r}="","",'
                'IF({gap_pct}{r}>0.15,"Dearer",'
                'IF({gap_pct}{r}<-0.15,"Cheaper","Similar")))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Cheaper", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Similar", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Dearer", "fmt": "info"}]),
        Col("stale", "Days old", 9, "int",
            formula=blank(
                "competitor",
                'IF({checked}{r}="","",TodayDate-{checked}{r})'),
            hidden=True),
        Col("check", "Note", 30, "status",
            formula=blank(
                "competitor",
                'IF({price}{r}="","Add - what do they charge?",'
                'IF({your_price}{r}="","Check - your product has no price",'
                'IF(AND({stale}{r}<>"",{stale}{r}>180),'
                '"Check - research is over six months old",'
                'IF(AND({position}{r}="Cheaper",{quality}{r}="Below mine"),'
                '"Raise - better work, lower price",'
                'IF({position}{r}="Dearer",'
                '"Fine - make sure the listing shows why","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Competitor Research",
        subtitle="What comparable work sells for, postage included, "
                 "and where that puts you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).competitors,
        help_text="Three or four honest comparisons per product "
                  "are worth more than twenty screenshots. The "
                  "catalog shows the average against your price.",
        extra_notes=[
            "Compare the total the buyer pays. Free shipping is a "
            "price, not a kindness, and buyers compare totals.",
            "Being cheaper than everybody is the most common "
            "mistake in handmade. If your work is better and your "
            "price is lower, the market is not telling you to "
            "stay there.",
            "Check again every few months. Everybody else's "
            "materials went up too.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/waste.py`

*121 lines · 2 functions*

> `Waste Log` - the material that never made it into a product.
>
> The frogged shawl, the failed pour, the board that split, the dye lot
> that did not match. None of it is free, all of it was bought, and
> most makers have no idea what it adds up to in a year.
>
> Logging waste does two things: it tells you what your real material
> cost is, and it shows you which mistakes are worth engineering out.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the waste log |

**Module constants:** `NAME` ('Waste Log').

````python
"""`Waste Log` - the material that never made it into a product.

The frogged shawl, the failed pour, the board that split, the dye lot
that did not match. None of it is free, all of it was bought, and
most makers have no idea what it adds up to in a year.

Logging waste does two things: it tells you what your real material
cost is, and it shows you which mistakes are worth engineering out.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, lookup_material,
                     month_col)

NAME = "Waste Log"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("material", "Material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("material_name", "Material name", 24, "text",
            formula=lookup_material("name")),
        Col("qty", "Quantity lost", 12, "number1",
            validate=number_check(0, 1000000),
            note="In the material's own unit."),
        Col("reason", "What happened", 22, "text",
            validate=list_from("WasteReasonList")),
        Col("product", "On which product", 12, "id",
            validate=list_from("ProductIDs"),
            note="Optional. Useful when one design keeps going "
                 "wrong."),
        Col("avoidable", "Avoidable?", 11, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Trim waste is not avoidable. A miscut is. Be "
                 "honest and the totals become useful."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        month_col("date"),
        Col("unit", "Unit", 9, "text", formula=lookup_material("unit"),
            hidden=True),
        Col("unit_cost", "Cost a unit", 11, "money4",
            formula=lookup_material("unit_cost", missing=""),
            hidden=True),
        Col("cost", "Cost", 11, "money",
            formula=blank(
                "date",
                'IF(OR({qty}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({qty}{r}*{unit_cost}{r},2))'),
            note="What walked out of the door. Add the column up "
                 "before you decide a mistake does not matter."),
        Col("share", "Share of that material", 13, "percent",
            formula=blank(
                "date",
                'IFERROR({cost}{r}/SUMIFS([[Waste Log.cost]],'
                '[[Waste Log.material]],{material}{r}),"")'),
            hidden=True),
        Col("check", "Note", 28, "status",
            formula=blank(
                "date",
                'IF({material}{r}="","Add - which material?",'
                'IF({qty}{r}="","Add - how much?",'
                'IF({unit_cost}{r}="",'
                '"Check - that material has no price",'
                'IF({avoidable}{r}="Yes",'
                '"Avoidable - worth a second look","Logged"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Material Waste Log",
        subtitle="What got wasted, what it cost, and how much of "
                 "it you could stop.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).waste,
        help_text="Log it when it happens - nobody remembers at "
                  "the end of the month. The totals tell you "
                  "whether your waste allowance is right.",
        extra_notes=[
            "Compare the total here with the waste allowance in "
            "your prices. If you allow 5% and waste 12%, every "
            "product is under-priced.",
            "Separate avoidable from unavoidable. Offcuts are "
            "physics; a miscut because you did not check the "
            "pattern twice is a process.",
            "Wasted material is bought material. A $9.40 skein "
            "frogged and unusable is $9.40 off the month, not a "
            "learning experience.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the waste log."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Waste summary", last)
    sheet.table_head(row, 0, ["Entries", "Cost of the waste",
                              "Avoidable", "This month",
                              "Waste as % of material spend"],
                     [12, 18, 14, 14, 26])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Waste Log.date]],"<>")',
        '=ROUND(SUM([[Waste Log.cost]]),2)',
        '=ROUND(SUMIFS([[Waste Log.cost]],[[Waste Log.avoidable]],'
        '"Yes"),2)',
        '=ROUND(SUMIFS([[Waste Log.cost]],[[Waste Log.date]],'
        '">="&MonthStart),2)',
        '=IFERROR(SUM([[Waste Log.cost]])/SUM([[Materials.cost_used]]),"")',
    ], ["int", "money", "money", "money", "percent"])
````

### `handmade_pricing_spreadsheet/sheets/development.py`

*150 lines · 1 function*

> `Development` - what a new design cost before it earned anything.
>
> Every product started as a sample that did not sell, an afternoon of
> sketching and a photo shoot. That time is real and it has to come
> back out of the pieces you go on to sell - or you have to admit the
> design was a hobby.
>
> This tab tracks development cost per idea and tells you how many
> sales recover it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Development').

````python
"""`Development` - what a new design cost before it earned anything.

Every product started as a sample that did not sell, an afternoon of
sketching and a photo shoot. That time is real and it has to come
back out of the pieces you go on to sell - or you have to admit the
design was a hobby.

This tab tracks development cost per idea and tells you how many
sales recover it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, lookup_product, status_cf

NAME = "Development"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "DevPrefix", digits=4, width=10),
        Col("name", "Design", 28, "text",
            note="The idea, whether or not it ever became a "
                 "product."),
        Col("product", "Became product", 12, "id",
            validate=list_from("ProductIDs"),
            note="Link it once the design is listed."),
        Col("stage", "Stage", 16, "text", validate=list_from("DevStageList"),
            cf=status_cf()),
        Col("started", "Started", 11, "date", validate=date_check()),
        Col("design_hours", "Design hours", 11, "number",
            validate=number_check(0, 10000)),
        Col("sample_hours", "Sample hours", 11, "number",
            validate=number_check(0, 10000),
            note="Making the prototypes, including the ones that "
                 "went in the bin."),
        Col("photo_hours", "Photo & listing hours", 15, "number",
            validate=number_check(0, 10000),
            note="Shooting, editing, writing the listing, "
                 "keywords. Hours that never feel like work and "
                 "always are."),
        Col("sample_cost", "Sample materials", 13, "money",
            validate=number_check(0, 1000000)),
        Col("other_cost", "Other costs", 11, "money",
            validate=number_check(0, 1000000),
            note="A pattern you bought to learn from, a tool you "
                 "needed, a model you paid."),
        Col("samples_sold", "Samples sold for", 13, "money",
            validate=number_check(0, 1000000),
            note="Prototypes often sell as seconds. Credited "
                 "back."),
        Col("decided", "Decision date", 11, "date", validate=date_check()),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("hours", "Hours", 9, "number",
            formula=blank(
                "name",
                'IF({design_hours}{r}="",0,{design_hours}{r})+'
                'IF({sample_hours}{r}="",0,{sample_hours}{r})+'
                'IF({photo_hours}{r}="",0,{photo_hours}{r})'),
            hidden=True),
        Col("dev_cost", "Development cost", 13, "money",
            formula=blank(
                "name",
                'ROUND({hours}{r}*LaborCostHour+'
                'IF({sample_cost}{r}="",0,{sample_cost}{r})+'
                'IF({other_cost}{r}="",0,{other_cost}{r})-'
                'IF({samples_sold}{r}="",0,{samples_sold}{r}),2)'),
            note="Hours at your rate, plus materials and costs, "
                 "less whatever the samples sold for."),
        Col("unit_profit", "Profit a sale", 11, "money",
            formula=lookup_product("profit", missing=""),
            hidden=True),
        Col("sold", "Sold since", 10, "int",
            formula=blank(
                "name",
                'IF({product}{r}="",0,SUMIFS([[Sales.qty]],'
                '[[Sales.product]],{product}{r}))')),
        Col("recovered", "Recovered", 11, "money",
            formula=blank(
                "name",
                'IF({unit_profit}{r}="",0,ROUND({sold}{r}*'
                '{unit_profit}{r},2))')),
        Col("recovery", "Recovered", 10, "percent",
            formula=blank(
                "name",
                'IF({dev_cost}{r}<=0,"",{recovered}{r}/{dev_cost}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "warn"}]),
        Col("to_recover", "Sales to recover it", 14, "int",
            formula=blank(
                "name",
                'IF(OR({unit_profit}{r}="",{unit_profit}{r}<=0),'
                '"Needs a profitable price",'
                'ROUNDUP(MAX(0,{dev_cost}{r}-{recovered}{r})/'
                '{unit_profit}{r},0))'),
            note="How many more you have to sell before the "
                 "design has paid for its own development."),
        Col("per_unit", "Development a sale", 13, "money",
            formula=blank(
                "name",
                'IF(OR({sold}{r}=0,{dev_cost}{r}<=0),"",'
                'ROUND({dev_cost}{r}/{sold}{r},2))'),
            hidden=True,
            note="Spread over what has sold so far. It falls "
                 "every time one goes out of the door."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({dev_cost}{r}<=0,"Add - hours and sample costs",'
                'IF({product}{r}="","Idea - not listed yet",'
                'IF({recovery}{r}>=1,"Recovered - this one paid for itself",'
                'IF(AND({sold}{r}=0,{started}{r}<>"",'
                'TodayDate-{started}{r}>90),'
                '"Check - three months, no sales",'
                'IF({recovery}{r}<0.5,'
                '"Watch - under half recovered","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Product Development",
        subtitle="What each design cost to create, and how many "
                 "sales pay that back.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).development,
        help_text="Log the hours while you are doing them. Link "
                  "the design to its product and the recovery "
                  "figure follows the sales.",
        extra_notes=[
            "Development is a real cost and it belongs somewhere. "
            "Either recover it through the price of that design "
            "or accept it as the cost of having a range.",
            "A design that has not recovered its development in "
            "six months is telling you something - usually about "
            "the photographs or the price rather than the design.",
            "Count the listing and photography hours. For a "
            "digital product they are most of the cost.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/events.py`

*176 lines · 1 function*

> `Events` - craft fairs, markets and whether the stall paid.
>
> A stall is a fixed cost you pay before you sell anything: the pitch,
> the petrol, the parking, the two days you were not making. This tab
> works out what the day has to take before it breaks even, and
> compares that with what it actually took.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Events').

````python
"""`Events` - craft fairs, markets and whether the stall paid.

A stall is a fixed cost you pay before you sell anything: the pitch,
the petrol, the parking, the two days you were not making. This tab
works out what the day has to take before it breaks even, and
compares that with what it actually took.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, loss_cf, status_cf

NAME = "Events"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("name", "Event", 26, "text",
            note="The name sales rows pick from, so keep it "
                 "short and distinct."),
        Col("kind", "Type", 18, "text", validate=list_from("EventTypeList")),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("days", "Days", 7, "number1",
            validate=number_check(0, 60)),
        Col("stall_fee", "Stall fee", 11, "money",
            validate=number_check(0, 100000)),
        Col("travel", "Travel & parking", 13, "money",
            validate=number_check(0, 100000)),
        Col("accommodation", "Accommodation", 13, "money",
            validate=number_check(0, 100000)),
        Col("other_cost", "Stock, display, other", 15, "money",
            validate=number_check(0, 100000),
            note="Banners, a new table cover, card reader fees, "
                 "the coffee."),
        Col("prep_hours", "Prep hours", 11, "number",
            validate=number_check(0, 1000),
            note="Making extra stock, pricing, packing the car."),
        Col("stall_hours", "Hours on the stall", 14, "number",
            validate=number_check(0, 1000),
            note="Including setting up and packing down. They "
                 "count."),
        Col("target", "Sales target", 11, "money",
            validate=number_check(0, 1000000)),
        Col("footfall", "Footfall", 10, "int",
            validate=number_check(0, 10000000, decimal=False)),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("cash_cost", "Money out", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({stall_fee}{r}="",0,{stall_fee}{r})+'
                'IF({travel}{r}="",0,{travel}{r})+'
                'IF({accommodation}{r}="",0,{accommodation}{r})+'
                'IF({other_cost}{r}="",0,{other_cost}{r}),2)')),
        Col("hours", "Hours", 9, "number",
            formula=blank(
                "name",
                'IF({prep_hours}{r}="",0,{prep_hours}{r})+'
                'IF({stall_hours}{r}="",0,{stall_hours}{r})'),
            hidden=True),
        Col("time_cost", "Your time", 11, "money",
            formula=blank(
                "name", 'ROUND({hours}{r}*LaborCostHour,2)'),
            note="At your own hourly rate. A stall is two days you "
                 "did not spend making, and that is the real cost "
                 "of a craft fair."),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "name", 'ROUND({cash_cost}{r}+{time_cost}{r},2)')),
        Col("break_even", "Sales to break even", 15, "money",
            formula=blank(
                "name",
                'IF(TargetRetailMargin<=0,"",ROUND({total_cost}{r}/'
                'TargetRetailMargin,2))'),
            note="At your target margin. Taking this much is not a "
                 "good day - it is the day costing you nothing."),
        Col("be_cash", "Cash break-even", 13, "money",
            formula=blank(
                "name",
                'IF(TargetRetailMargin<=0,"",ROUND({cash_cost}{r}/'
                'TargetRetailMargin,2))'),
            hidden=True,
            note="Ignoring your time - the number most stallholders "
                 "use, and the reason craft fairs feel better than "
                 "they are."),
        Col("sales", "Actually took", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.revenue]],[[Sales.event]],'
                '{name}{r}),2)')),
        Col("orders", "Orders", 8, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Sales.event]],{name}{r})')),
        Col("gross_profit", "Profit on the goods", 14, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.profit]],[[Sales.event]],'
                '{name}{r}),2)'),
            hidden=True),
        Col("net", "Profit after the stall", 15, "money",
            formula=blank(
                "name",
                'IF({sales}{r}=0,"",ROUND({gross_profit}{r}-'
                '{total_cost}{r},2))'),
            cf=loss_cf(),
            note="What the day really made once the pitch, the "
                 "petrol and your hours are paid."),
        Col("per_hour", "Earned an hour", 11, "money",
            formula=blank(
                "name",
                'IF(OR({hours}{r}=0,{sales}{r}=0),"",'
                'ROUND(({gross_profit}{r}-{cash_cost}{r})/'
                '{hours}{r},2))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MinimumWage",
                 "fmt": "bad"}],
            note="Profit on what you sold, less the cash costs, "
                 "over every hour the event took."),
        Col("avg_order", "Average sale", 11, "money",
            formula=blank(
                "name",
                'IF({orders}{r}=0,"",ROUND({sales}{r}/{orders}{r},2))'),
            hidden=True),
        Col("conversion", "Buyers per 100 passing", 15, "percent",
            formula=blank(
                "name",
                'IF(OR({footfall}{r}="",{footfall}{r}=0),"",'
                '{orders}{r}/{footfall}{r})'),
            hidden=True),
        Col("status", "Status", 13, "status",
            formula=blank(
                "name",
                'IF(AND({date}{r}<>"",{date}{r}>TodayDate),"Booked",'
                'IF({sales}{r}=0,"No sales logged","Done"))'),
            cf=status_cf()),
        Col("verdict", "Verdict", 30, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Booked","Coming up - take enough stock",'
                'IF({sales}{r}=0,"Add - log the sales against this event",'
                'IF({net}{r}<0,"Loss - the stall cost more than it made",'
                'IF(AND({per_hour}{r}<>"",{per_hour}{r}<MinimumWage),'
                '"Below - paid under minimum wage for the day",'
                'IF({net}{r}>{total_cost}{r},'
                '"Worth repeating - book it again","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Craft Fairs & Events",
        subtitle="What the stall costs, what it has to take, and "
                 "what it actually did.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).events,
        help_text="Enter the costs before the event and the sales "
                  "after it - tag the Sales rows with the event "
                  "name and they total themselves.",
        extra_notes=[
            "Count your hours. A $45 pitch is not the cost of a "
            "craft fair; a $45 pitch plus fourteen hours is.",
            "Work out the break-even figure before you book. If "
            "it is more than you have ever taken at a similar "
            "event, you already know the answer.",
            "A fair that breaks even but brings three wholesale "
            "enquiries was worth it. Write that in the notes - "
            "the money column cannot see it.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/wholesale_customers.py`

*160 lines · 1 function*

> `Wholesale Customers` - the shops that buy from you, and what they are worth.
>
> Wholesale is a different business from retail: bigger orders, half
> the price, no marketplace fees and a shop that either reorders or
> does not. This tab keeps the accounts, the terms and the honest
> answer to "is this stockist worth the discount?".

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Wholesale Customers').

````python
"""`Wholesale Customers` - the shops that buy from you, and what they are worth.

Wholesale is a different business from retail: bigger orders, half
the price, no marketplace fees and a shop that either reorders or
does not. This tab keeps the accounts, the terms and the honest
answer to "is this stockist worth the discount?".
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..layout import max_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, loss_cf, status_cf

NAME = "Wholesale Customers"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "CustomerPrefix", digits=4, width=10),
        Col("name", "Shop", 26, "text"),
        Col("kind", "Type", 18, "text",
            validate=list_from("CustomerTypeList")),
        Col("contact", "Contact", 18, "text"),
        Col("email", "Email", 24, "text"),
        Col("phone", "Phone", 15, "text"),
        Col("city", "Town / city", 16, "text"),
        Col("terms", "Payment terms", 22, "text",
            validate=list_from("TermsList"),
            note="Net 30 means you are lending a shop money for a "
                 "month. Price that in or ask for a deposit."),
        Col("discount", "Their discount", 11, "percent",
            validate=number_check(0, 0.9),
            note="Off your retail price. Blank uses the standard "
                 "wholesale discount on Settings."),
        Col("min_order", "Their minimum", 11, "money",
            validate=number_check(0, 1000000),
            hidden=True),
        Col("first_order", "First order", 11, "date",
            validate=date_check()),
        Col("stockist_since", "Stockist since", 12, "date",
            validate=date_check(), hidden=True),
        Col("exclusive", "Exclusive area?", 12, "flag",
            validate=list_from("YesNoList"),
            note="Granting exclusivity is worth something. Make "
                 "sure you got something for it."),
        Col("active", "Still buying?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("discount_used", "Discount used", 11, "percent",
            formula=blank(
                "name",
                'IF({discount}{r}<>"",{discount}{r},WholesaleDiscount)'),
            hidden=True),
        Col("orders", "Orders", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Wholesale Orders.customer]],{id}{r})')),
        Col("units", "Items bought", 11, "int",
            formula=blank(
                "name",
                'SUMIFS([[Wholesale Orders.qty]],'
                '[[Wholesale Orders.customer]],{id}{r})')),
        Col("value", "Total spent", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Wholesale Orders.order_value]],'
                '[[Wholesale Orders.customer]],{id}{r}),2)')),
        Col("profit", "Profit from them", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Wholesale Orders.profit]],'
                '[[Wholesale Orders.customer]],{id}{r}),2)'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({value}{r}="",{value}{r}=0),"",{profit}{r}/'
                '{value}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetWholesaleMargin", "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=TargetWholesaleMargin", "fmt": "good"}]),
        Col("avg_order", "Average order", 11, "money",
            formula=blank(
                "name",
                'IF({orders}{r}=0,"",ROUND({value}{r}/{orders}{r},2))')),
        Col("last_order", "Last order", 11, "date",
            formula=blank(
                "name",
                'IF(' + max_if("[[Wholesale Orders.date]]",
                               "[[Wholesale Orders.customer]]={id}{r}") +
                '=0,"",' + max_if("[[Wholesale Orders.date]]",
                                  "[[Wholesale Orders.customer]]={id}{r}") +
                ')')),
        Col("days_quiet", "Days since", 10, "int",
            formula=blank(
                "name",
                'IF({last_order}{r}="","",TodayDate-{last_order}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 180,
                 "fmt": "warn"}],
            hidden=True),
        Col("outstanding", "Unpaid", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Wholesale Orders.order_value]],'
                '[[Wholesale Orders.customer]],{id}{r},'
                '[[Wholesale Orders.paid]],"No"),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("status", "Status", 13, "status",
            formula=blank(
                "name",
                'IF({orders}{r}=0,"Prospect",'
                'IF({active}{r}="No","Inactive",'
                'IF(AND({days_quiet}{r}<>"",{days_quiet}{r}>180),'
                '"Going quiet","Active")))'),
            cf=status_cf()),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({orders}{r}=0,"New - no orders yet",'
                'IF({outstanding}{r}>0,"Chase - money outstanding",'
                'IF({margin}{r}<0,"Loss - this stockist costs you money",'
                'IF(AND({margin}{r}<>"",{margin}{r}<'
                'TargetWholesaleMargin),'
                '"Low margin - their discount is too deep",'
                'IF({status}{r}="Going quiet",'
                '"Chase - six months without an order","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Wholesale Customers",
        subtitle="The shops that stock you, their terms, and what "
                 "each one is really worth.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).wholesale_customers,
        help_text="Set each shop's discount and terms here; the "
                  "order tab prices against them automatically.",
        extra_notes=[
            "A stockist who reorders every quarter at 50% off is "
            "worth more than three who buy once at 40%. Look at "
            "the order count before the discount.",
            "Net 30 is a loan. If a shop wants terms, a deposit "
            "or a slightly smaller discount is a fair trade.",
            "Exclusivity has a price. If a shop wants to be the "
            "only one in town, that is worth a bigger minimum "
            "order, not a bigger discount.",
        ],
    )
````

### `handmade_pricing_spreadsheet/sheets/wholesale_orders.py`

*174 lines · 1 function*

> `Wholesale Orders` - line by line, priced at the shop's own discount.
>
> One row per product per order. The price comes from the catalog at
> whatever discount that customer gets, the cost comes from the
> catalog too, and the margin is what is left - with no marketplace
> fees, because a wholesale order does not pay any.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Wholesale Orders').

````python
"""`Wholesale Orders` - line by line, priced at the shop's own discount.

One row per product per order. The price comes from the catalog at
whatever discount that customer gets, the cost comes from the
catalog too, and the margin is what is left - with no marketplace
fees, because a wholesale order does not pay any.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, lookup_product,
                     month_col, status_cf)

NAME = "Wholesale Orders"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "OrderPrefix", digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("customer", "Customer", 12, "id",
            validate=list_from("CustomerIDs")),
        Col("po_ref", "Their order ref", 14, "text"),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs")),
        Col("qty", "Quantity", 9, "int",
            validate=number_check(1, 1000000, decimal=False)),
        Col("price_input", "Agreed price each", 13, "money",
            validate=number_check(0, 1000000),
            note="Leave blank to use the catalog price at this "
                 "customer's discount."),
        Col("due", "Due", 11, "date", validate=date_check()),
        Col("despatched", "Despatched", 11, "date", validate=date_check()),
        Col("carriage", "Carriage charged", 13, "money",
            validate=number_check(0, 100000),
            hidden=True),
        Col("carriage_cost", "Carriage cost", 12, "money",
            validate=number_check(0, 100000),
            hidden=True),
        Col("status", "Status", 14, "status",
            validate=list_from("OrderStatusList"), cf=status_cf()),
        Col("paid", "Paid?", 8, "flag", validate=list_from("YesNoList")),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("customer_name", "Shop", 22, "text",
            formula=blank(
                "customer",
                'IFERROR(INDEX([[Wholesale Customers.name]],'
                'MATCH({customer}{r},[[Wholesale Customers.id]],0)),'
                '"Not found")')),
        Col("product_name", "Product name", 24, "text",
            formula=lookup_product("name")),
        month_col("date"),
        Col("discount", "Discount", 9, "percent",
            formula=blank(
                "customer",
                'IFERROR(INDEX([[Wholesale Customers.discount_used]],'
                'MATCH({customer}{r},[[Wholesale Customers.id]],0)),'
                'WholesaleDiscount)'),
            hidden=True),
        Col("retail", "Retail each", 11, "money",
            formula=lookup_product("price", missing=""),
            hidden=True),
        Col("unit_price", "Price each", 11, "money",
            formula=blank(
                "product",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF({retail}{r}="","",ROUND({retail}{r}*'
                '(1-{discount}{r}),2)))')),
        Col("unit_cost", "Cost each", 11, "money",
            formula=lookup_product("cost_base", missing=""),
            hidden=True,
            note="The making cost without the retail postage - a "
                 "wholesale order ships as one box."),
        Col("order_value", "Order value", 12, "money",
            formula=blank(
                "product",
                'IF(OR({qty}{r}="",{unit_price}{r}=""),"",'
                'ROUND({qty}{r}*{unit_price}{r},2))')),
        Col("order_cost", "Order cost", 11, "money",
            formula=blank(
                "product",
                'IF(OR({qty}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({qty}{r}*{unit_cost}{r}+'
                'MAX(0,IF({carriage_cost}{r}="",0,{carriage_cost}{r})-'
                'IF({carriage}{r}="",0,{carriage}{r})),2))')),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                "product",
                'IF({order_value}{r}="","",ROUND({order_value}{r}-'
                '{order_cost}{r},2))'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "product",
                'IF(OR({order_value}{r}="",{order_value}{r}=0),"",'
                '{profit}{r}/{order_value}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetWholesaleMargin", "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=TargetWholesaleMargin", "fmt": "good"}]),
        Col("hours", "Hours of making", 12, "number",
            formula=blank(
                "product",
                'IFERROR(ROUND({qty}{r}*INDEX([[Products.minutes]],'
                'MATCH({product}{r},[[Products.id]],0))/60,1),"")'),
            note="A 60 piece order at 25 minutes each is 25 hours "
                 "of your life. Check the diary before you say "
                 "yes."),
        Col("per_hour", "Profit an hour", 11, "money",
            formula=blank(
                "product",
                'IF(OR({hours}{r}="",{hours}{r}=0),"",'
                'ROUND({profit}{r}/{hours}{r},2))'),
            hidden=True),
        Col("meets_min", "Meets minimum?", 12, "flag",
            formula=blank(
                "product",
                'IF({order_value}{r}="","",IF(AND({order_value}{r}>='
                'MinOrderValue,{qty}{r}>=MinOrderQty),"Yes","No"))'),
            hidden=True),
        Col("lead_days", "Days to despatch", 12, "int",
            formula=blank(
                "date",
                'IF({despatched}{r}="","",{despatched}{r}-{date}{r})'),
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({product}{r}="","Add - which product?",'
                'IF({profit}{r}<0,"Loss - this order costs you money",'
                'IF(AND({margin}{r}<>"",{margin}{r}<'
                'TargetWholesaleMargin),'
                '"Low margin - the discount is too deep",'
                'IF(AND({due}{r}<>"",{despatched}{r}="",'
                '{due}{r}<TodayDate),"Overdue - not despatched",'
                'IF(AND({paid}{r}="No",{despatched}{r}<>"",'
                'TodayDate-{despatched}{r}>30),'
                '"Chase - unpaid over 30 days",'
                'IF({meets_min}{r}="No",'
                '"Check - below your minimum order","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Wholesale Orders",
        subtitle="Line by line, at the shop's own discount, with "
                 "the hours it will take you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=3,
        samples=sample(cfg).wholesale_orders,
        help_text="One row per product per order. The price comes "
                  "from the catalog at that customer's discount "
                  "unless you agreed something else.",
        extra_notes=[
            "Wholesale margin is thinner on purpose - there are "
            "no marketplace fees and the shop does the selling. "
            "It still has to be positive after your own wage.",
            "Check the hours before you accept. A big order at a "
            "thin margin that fills six weeks is how makers end "
            "up unable to take the profitable work.",
            "If a line will not make the target wholesale margin, "
            "the retail price is too low. Fix the retail price "
            "rather than refusing the shop.",
        ],
    )
````

---

## 6. The pages

The tabs that read the registers: the welcome page, settings, the dashboard, the retail price calculator, the wholesale calculator, profit and break-even, the pricing scenarios, the reports, the three printables and the help tab.

### `handmade_pricing_spreadsheet/sheets/start_here.py`

*336 lines · 1 function*

> `Start Here` - the first thing anyone sees when they open the file.
>
> An hour of setup, then this workbook answers the question that keeps
> handmade businesses small: what should I actually charge? It explains
> the order to do things in, what fills itself in, the pricing formula
> in plain words, and - just as honestly - what a spreadsheet cannot do.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Start Here'), `LAST_COL` (9).

````python
"""`Start Here` - the first thing anyone sees when they open the file.

An hour of setup, then this workbook answers the question that keeps
handmade businesses small: what should I actually charge? It explains
the order to do things in, what fills itself in, the pricing formula
in plain words, and - just as honestly - what a spreadsheet cannot do.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Start Here"
LAST_COL = 9


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    ws = s.ws

    s.banner(cfg.product_name,
             "Materials, labour, packaging, fees and overhead - "
             "turned into a price that pays you.", LAST_COL)

    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 18)
    s.set_col(3, LAST_COL, 15)

    row = 4
    row = s.note(row,
                 "Most handmade pricing advice stops at 'materials "
                 "times two'. That is why so many makers work for "
                 "nothing. This workbook counts the yarn, the hours, "
                 "the box, the postage, the Etsy fees and the studio "
                 "rent, then solves for the price that leaves the "
                 "margin you actually want.",
                 LAST_COL, height=42)
    row += 1

    # ------------------------------------------------- the five minutes
    row = s.section(row, "Five minutes to your first honest price",
                    LAST_COL)
    row = s.table_head(row, 1, ["Step", "Where", "What to do"],
                       [10, 24, 62])
    quick = [
        ("1", "Settings", "Put your hourly rate in. That one number "
                          "changes more than anything else here."),
        ("2", "Materials", "Add the three or four materials the "
                           "product uses, with what you paid and how "
                           "much came in the pack."),
        ("3", "Material Usage", "List how much of each goes into one "
                                "piece."),
        ("4", "Labor", "Put the minutes in - all of them, including "
                       "finishing and packing."),
        ("5", "Products", "Add the product. Cost, suggested price, "
                          "margin and your hourly wage appear."),
        ("6", "Retail Calculator",
         "Check the price, try the ladder, and decide."),
    ]
    for step, where, what in quick:
        s.write(row, 1, step, fmts.chrome("label_bold"))
        s.write(row, 2, where, fmts.chrome("label_bold"))
        s.merge(row, 3, LAST_COL, what, fmts.chrome("label"))
        s.row_height(row, 22)
        row += 1
    row += 1

    # ----------------------------------------------------- the formula
    row = s.section(row, "The formula everything is built on", LAST_COL)
    row = s.note(row,
                 "Marketplace fees are charged on the price - the "
                 "price you have not worked out yet. That is why you "
                 "cannot just add 20% to your costs: the fee is "
                 "charged on the fee. Solving it properly looks like "
                 "this.",
                 LAST_COL, first_col=1, height=34)
    formulas = [
        ("Cost", "C = materials + labour + packaging + postage + "
                 "overhead", "What one piece costs you."),
        ("Labour", "L = minutes / 60 x your hourly rate",
         "Every minute, not just the fun ones."),
        ("Price", "Price = (C + fixed fees) / (1 - target margin - "
                  "fee rate)",
         "The price that leaves the margin you asked for, after the "
         "percentage fees."),
        ("Margin", "Margin = profit / price x 100",
         "The share of the price you keep."),
        ("Markup", "Markup = profit / cost x 100",
         "The share of the cost you added. Not the same thing."),
        ("Break-even", "Units = fixed costs / contribution a unit",
         "Rounded up, because you cannot sell two thirds of a "
         "candle."),
    ]
    for label, formula, meaning in formulas:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, 5, formula, fmts.chrome("label_bold"))
        s.merge(row, 6, LAST_COL, meaning, fmts.chrome("note"))
        s.row_height(row, 22)
        row += 1
    row += 1
    row = s.note(row,
                 "Worked example. A beanie costs $8.40 in yarn and "
                 "labour, $1.85 to package and post, and carries "
                 "$1.20 of overhead - $11.45. Etsy takes about 9.5% "
                 "plus $0.45. You want a 55% margin. The price is "
                 "(11.45 + 0.45) / (1 - 0.55 - 0.095) = $33.52, "
                 "rounded to $33.99. Adding 55% to the cost would "
                 "have given $17.75 and a real margin of 25%.",
                 LAST_COL, first_col=1, height=56)
    row += 1

    # ------------------------------------------------------ the tabs --
    row = s.section(row, "What is in the file", LAST_COL)
    row = s.table_head(row, 1, ["Tab", "What it does", "Your job"],
                       [26, 54, 28])
    tabs = [
        ("Settings", "Your rates, targets, fee rates, rounding rule "
                     "and every default", "Fill in once, review "
                                          "twice a year"),
        ("Dashboard", "Your margins, your month and everything that "
                      "needs fixing", "Read it weekly"),
        ("Products", "The catalog: cost, price, margin, hourly wage "
                     "and what sold", "The heart of the workbook"),
        ("Materials", "Every yarn, bead, wax and board, with the "
                      "cost per gram worked out", "Add as you buy"),
        ("Material Usage", "How much of each material goes into "
                           "each product", "The bill of materials"),
        ("Labor", "Minutes per stage, per item or per batch",
         "Be honest about the time"),
        ("Packaging / Shipping", "Boxes, tissue, labels, postage "
                                 "services and what you absorb",
         "Set up once"),
        ("Overhead", "Studio, insurance, subscriptions - shared "
                     "across what you make", "Add your bills"),
        ("Channels", "Etsy, Shopify, markets, wholesale - every fee, "
                     "fully editable", "Check twice a year"),
        ("Retail Calculator", "The price, the ladder and the "
                              "sanity check", "Use before you list"),
        ("Wholesale Calculator", "What a stockist leaves you and "
                                 "what retail price supports it",
         "Use before you say yes"),
        ("Profit & Break-Even", "Contribution, break-even units, "
                                "margin against markup",
         "Use when you are planning"),
        ("Crochet / Jewelry / Candles", "Specialist calculators that "
                                        "work the way those crafts "
                                        "actually work", "Use yours"),
        ("Clothing / Woodworking / Digital", "Fabric and sizes, "
                                             "timber and offcuts, "
                                             "design time spread "
                                             "over sales",
         "Use yours"),
        ("Bundles / Variations", "Gift sets, sizes, personalisation "
                                 "and rush orders", "As needed"),
        ("Batches", "What a run really cost against the estimate",
         "Log your runs"),
        ("Sales", "Every order with the real fees and real postage",
         "Log them weekly"),
        ("Promotions / Price History / Competitors",
         "What discounts cost, what price changes did, what the "
         "market charges", "Keep it current"),
        ("Waste Log / Development / Events",
         "Material lost, design time, craft fair stalls",
         "Log as it happens"),
        ("Wholesale Customers / Orders", "Stockists, their terms and "
                                         "what each is worth",
         "If you wholesale"),
        ("Scenarios / Reports", "Reposition the range, compare "
                                "channels, survive a price rise",
         "Monthly"),
        ("Print Costing / Quote / Order", "Three printable forms, "
                                          "filled in from your data",
         "Print when needed"),
        ("Lists / Help", "Every drop-down, every factor table, and "
                         "the troubleshooting", "Edit the lists"),
    ]
    for tab, what, job in tabs:
        s.write(row, 1, tab, fmts.chrome("label_bold"))
        s.merge(row, 2, 6, what, fmts.chrome("label"))
        s.merge(row, 7, LAST_COL, job, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ---------------------------------------------------- the rhythm --
    row = s.section(row, "A rhythm that keeps it accurate", LAST_COL)
    rhythm = [
        ("As you buy", "Add new materials with the pack price and "
                       "pack size. Nothing else is needed.",
         "1 minute"),
        ("As you make", "Log batches with the real hours. Two runs "
                        "and your estimates become facts.",
         "2 minutes"),
        ("Every week", "Enter the week's sales with the real fees "
                       "and postage.", "15 minutes"),
        ("Every month", "Read the Dashboard. Fix anything below "
                        "minimum wage before you make more of it.",
         "20 minutes"),
        ("Every quarter", "Check your material prices and the fee "
                          "rates on the Channels tab. Both move.",
         "30 minutes"),
        ("Twice a year", "Re-price. Log every change on Price "
                         "History so you can see what worked.",
         "1 hour"),
    ]
    for when, what, how_long in rhythm:
        s.write(row, 1, when, fmts.chrome("label_bold"))
        s.merge(row, 2, 7, what, fmts.chrome("label"))
        s.merge(row, 8, LAST_COL, how_long, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ------------------------------------------------- what it does ---
    row = s.section(row, "What the workbook does by itself", LAST_COL)
    row = s.bullets(row, [
        "Turns a pack price into a cost per gram, millilitre or "
        "centimetre, and applies your waste allowance.",
        "Adds every material and every labour stage into a cost per "
        "piece, with a difficulty factor for the fiddly ones.",
        "Solves for the price that leaves your target margin after "
        "percentage fees and fixed fees - separately, because they "
        "behave differently.",
        "Applies your rounding rule so prices come out at $33.99 "
        "rather than $33.52.",
        "Warns you when a product pays under minimum wage, loses "
        "money, or has not been re-priced while its materials moved.",
        "Prices six crafts the way those crafts actually work - "
        "skeins and hooks, grams of silver and scrap recovery, wax "
        "and fragrance load, fabric and size factors, timber and "
        "offcuts, design time spread over expected sales.",
        "Works out wholesale both ways: what a discount leaves you, "
        "and what retail price would make it work.",
        "Shows break-even in units, in money and in hours.",
        "Compares what you estimated with what the batch really cost "
        "and what the marketplace really charged.",
        "Tells you what a discount costs in profit and how much more "
        "you would have to sell to stand still.",
    ], LAST_COL, first_col=1, height=16)
    row += 1

    # ------------------------------------------------ honest limits ---
    row = s.section(row, "What it honestly cannot do", LAST_COL)
    row = s.note(row,
                 "Every one of these needs something outside a "
                 "spreadsheet. Anyone who tells you otherwise is "
                 "selling you something.", LAST_COL, first_col=1,
                 height=22)
    row = s.bullets(row, [
        "Know the current Etsy, Shopify or Amazon Handmade fees. "
        "They change, they vary by country, and they are different "
        "for offsite ads. Put your real rates on the Channels tab "
        "and check them against the official fee schedule twice a "
        "year.",
        "Pull your orders in automatically. Export a CSV from your "
        "shop and paste it into the Sales tab.",
        "Tell you what your market will pay. It tells you what you "
        "need - the Competitors tab is where you record the other "
        "half.",
        "Make a product sell. A correct price on a bad photograph "
        "sells nothing.",
        "Handle sales tax, VAT or import duty properly. It has a "
        "simple tax rate and nothing more.",
        "Replace an accountant, and it does not know your local "
        "rules on labelling, safety testing or candle regulations.",
    ], LAST_COL, first_col=1, height=16)
    row += 1

    # ------------------------------------------------ the fee warning -
    row = s.section(row, "About marketplace fees", LAST_COL)
    s.merge(row, 1, LAST_COL, None, fmts.calc("text"))
    ws.write_formula(
        row, 1,
        '="Every fee in this workbook is a number you can edit on '
        'the Channels tab, because published fees change and vary by '
        'country. The rates in this file were last reviewed on "&'
        'TEXT(FeesReviewed,"d mmmm yyyy")&". Before you rely on a '
        'price, verify the current official fee schedule for your '
        'marketplace and your country, and update the Channels tab."',
        fmts.calc("text"))
    s.row_height(row, 44)
    row += 2

    # ------------------------------------------------ housekeeping ----
    row = s.section(row, "Housekeeping", LAST_COL)
    row = s.table_head(row, 1, ["Question", "Answer"], [30, 68])
    facts = [
        ("Is there example data in it?",
         "Both versions are in your download: one filled with a "
         "worked example shop, one empty and ready for yours."),
        ("How do I clear the example?",
         "Select the data rows on a tab and press Delete. The "
         "formula columns refill themselves - they are not deleted."),
        ("Why are some cells locked?",
         f"Every calculated cell is protected so a stray paste "
         f"cannot break it. The password is '{cfg.password}'. "
         f"Review, Unprotect Sheet if you want to change a formula."),
        ("Which cells do I type in?",
         "The white ones. Shaded cells work themselves out."),
        ("Can I change the currency?",
         "Yes - the Settings tab holds the symbol, and every money "
         "cell follows it."),
        ("Does it work in Google Sheets?",
         "Yes. File, Import, Upload. Drop-downs and colours come "
         "across; sheet protection behaves a little differently."),
        ("Which Excel versions?",
         "Excel 2010 and newer, Excel for Mac, Excel online, "
         "LibreOffice Calc and Google Sheets. No macros, so there "
         "is nothing to enable."),
        ("Can I add more rows?",
         "Yes - insert rows inside the existing block and copy the "
         "formulas down. Never insert above the header row."),
        ("My craft is not one of the six tabs",
         "Use the Products, Materials and Labor tabs - they work "
         "for anything. The six specialist tabs are shortcuts, not "
         "requirements."),
    ]
    for question, answer in facts:
        s.write(row, 1, question, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, answer, fmts.chrome("label"))
        s.row_height(row, 26)
        row += 1
    row += 1

    s.note(row,
           f"{cfg.product_name} - version {cfg.version}, by "
           f"{cfg.author}. Built with formulas only: no macros, no "
           f"add-ins, nothing to install. The Help tab has the "
           f"troubleshooting and a map of every feature.",
           LAST_COL, first_col=1, height=34)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/settings.py`

*401 lines · 1 function*

> The `Settings` tab - shop profile, labour rates, fees, margins, targets.
>
> Everything the rest of the workbook needs to know about your craft
> business is entered once here and published as a defined name
> (ShopName, HourlyRate, TargetRetailMargin, TransactionFee ...) so the
> formulas stay readable and there is exactly one place to change a
> number. Nothing in this workbook hard codes a fee or a rate.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `SHEET` ('Settings'), `LAST_COL` (8).

````python
"""
The `Settings` tab - shop profile, labour rates, fees, margins, targets.

Everything the rest of the workbook needs to know about your craft
business is entered once here and published as a defined name
(ShopName, HourlyRate, TargetRetailMargin, TransactionFee ...) so the
formulas stay readable and there is exactly one place to change a
number. Nothing in this workbook hard codes a fee or a rate.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

SHEET = "Settings"
LAST_COL = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, SHEET, tab_color=Palette.accent)
    b = cfg.shop

    s.banner("Business Settings",
             "Fill in the white cells once. Shaded cells are locked "
             "formulas. Every price, every margin and every warning in "
             "the workbook reads these values.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 38, fmts.chrome("label"))
    s.set_col(2, 2, 14, fmts.chrome("label"))
    s.set_col(3, 3, 24)
    s.set_col(4, 4, 18)
    s.set_col(5, 8, 17, fmts.chrome("note"))

    r = 4  # 0 based row index -> Excel row 5

    # ------------------------------------------------------------ 1. the shop
    r = s.section(r, "1. Your shop", LAST_COL)
    r = s.field(r, "Business name", b.name, name="ShopName", value_end=4,
                help_text="Shown on every quote, order form and "
                          "printable page.")
    r = s.field(r, "Maker / owner", b.owner, name="ShopOwner", value_end=4)
    r = s.field(r, "Brand name", b.brand, name="ShopBrand", value_end=4,
                help_text="If you trade under a shorter name than the "
                          "business one.")
    r = s.field(r, "Email", b.email, name="ShopEmail", value_end=4)
    r = s.field(r, "Phone", b.phone, name="ShopPhone", value_end=4)
    r = s.field(r, "Address", b.address, name="ShopAddress", value_end=4)
    r = s.field(r, "Website", b.website, name="ShopWebsite", value_end=4)
    r = s.field(r, "Marketplace shop name", b.etsy_shop, name="ShopEtsy",
                value_end=4,
                help_text="Your Etsy / Amazon Handmade / Folksy shop "
                          "name, for the order forms.")
    r = s.field(r, "Social handle", b.instagram, name="ShopSocial",
                value_end=4)
    r = s.field(r, "Tagline", b.tagline, name="ShopTagline", value_end=4)
    r = s.field(r, "Business / tax number", b.tax_id, name="ShopTaxId",
                value_end=4)

    # ------------------------------------------------------- 2. the calendar
    r += 1
    r = s.section(r, "2. Dates", LAST_COL)
    r = s.field(r, "Year", cfg.year, style="int", name="PlanYear",
                validate=number_check(2000, 2100, decimal=False),
                help_text="The year the reports and the sales summaries "
                          "open on.")
    r = s.field(r, "Today", None, style="date", mode="calc",
                name="TodayDate", formula="=TODAY()",
                help_text="Every 'due', 'stale price' and 'review the "
                          "fees' flag counts from this cell, so they all "
                          "move on by themselves overnight.")
    r = s.field(r, "This month starts", None, style="date", mode="calc",
                name="MonthStart",
                formula="=DATE(YEAR(TodayDate),MONTH(TodayDate),1)")
    r = s.field(r, "This month ends", None, style="date", mode="calc",
                name="MonthEnd",
                formula="=DATE(YEAR(TodayDate),MONTH(TodayDate)+1,1)-1")
    r = s.field(r, "This year starts", None, style="date", mode="calc",
                name="YearStart", formula="=DATE(PlanYear,1,1)")

    # --------------------------------------------------- 3. your time is cost
    r += 1
    r = s.section(r, "3. What your time is worth", LAST_COL)
    r = s.note(r, "The single number most handmade sellers get wrong. If "
                  "your products do not pay you for the hours you spend "
                  "making them, you do not have a business, you have an "
                  "expensive hobby. Pay yourself a real wage here and let "
                  "the prices follow.",
               LAST_COL, first_col=1, height=38)
    r = s.field(r, "What you pay yourself an hour", cfg.hourly_rate,
                style="money", name="HourlyRate",
                validate=number_check(0, 1000),
                help_text="What one hour of your making time costs. Not "
                          "what is left over - what you are worth.")
    r = s.field(r, "Local minimum wage (for the warning)", cfg.minimum_wage,
                style="money", name="MinimumWage",
                validate=number_check(0, 1000),
                help_text="The workbook flags any product whose price "
                          "pays you less than this an hour.")
    r = s.field(r, "Employee hourly rate", cfg.employee_rate, style="money",
                name="EmployeeRate", validate=number_check(0, 1000))
    r = s.field(r, "Contractor / outsourced hourly rate",
                cfg.contractor_rate, style="money", name="ContractorRate",
                validate=number_check(0, 1000),
                help_text="Used when a labour line says somebody else "
                          "made it.")
    r = s.field(r, "On-costs on paid labour", cfg.payroll_burden,
                style="percent", name="PayrollBurden",
                validate=number_check(0, 1),
                help_text="Employer tax, insurance and holiday pay. "
                          "Added to employee and contractor hours.")
    r = s.field(r, "Your hour, with on-costs", None, style="money",
                mode="calc", name="LaborCostHour",
                formula="=ROUND(HourlyRate*(1+PayrollBurden),2)",
                help_text="This is the rate the labour calculator "
                          "multiplies your minutes by.")
    r = s.field(r, "Hours a month you actually make things",
                cfg.productive_hours, style="number1",
                name="ProductiveHours", validate=number_check(1, 744),
                help_text="Bench hours only - not admin, not photos, not "
                          "scrolling. Overhead per hour is divided by "
                          "this.")
    r = s.field(r, "Items a month you usually finish", cfg.units_per_month,
                style="int", name="UnitsPerMonth",
                validate=number_check(1, 100000, decimal=False),
                help_text="Used when overhead is spread per item instead "
                          "of per hour.")
    r = s.field(r, "Default batch size", cfg.default_batch_size,
                style="int", name="DefaultBatchSize",
                validate=number_check(1, 10000, decimal=False))
    r = s.field(r, "Consumables an hour", cfg.supplies_per_hour,
                style="money", name="SuppliesPerHour",
                validate=number_check(0, 1000),
                help_text="Glue, needles, sandpaper, thread - the small "
                          "things nobody itemises. Added per labour hour.")

    # ---------------------------------------------------- 4. materials, waste
    r += 1
    r = s.section(r, "4. Materials and waste", LAST_COL)
    r = s.field(r, "Measurement system", cfg.measure_system,
                name="MeasureSystem",
                validate={"validate": "list",
                          "source": ["Metric", "Imperial"]},
                help_text="Metric weighs in grams, imperial in ounces. "
                          "The unit table on Lists converts either way.")
    r = s.field(r, "Default waste allowance", cfg.default_waste,
                style="percent", name="DefaultWaste",
                validate=number_check(0, 1),
                help_text="Offcuts, frogged rows, the last inch of a "
                          "board. Added to every material line that does "
                          "not set its own.")
    r = s.field(r, "Default defect / seconds rate", cfg.default_defect_rate,
                style="percent", name="DefaultDefect",
                validate=number_check(0, 1),
                help_text="The share of a batch that never makes it to "
                          "the shop. The batch tab spreads its cost over "
                          "the pieces that do.")
    r = s.field(r, "Days of stock that counts as low", cfg.low_stock_days,
                style="int", name="LowStockDays",
                validate=number_check(0, 365, decimal=False))

    # ------------------------------------------------ 5. packaging, shipping
    r += 1
    r = s.section(r, "5. Packaging and shipping", LAST_COL)
    r = s.field(r, "Fallback packaging cost", cfg.default_packaging,
                style="money", name="DefaultPackaging",
                validate=number_check(0, 1000),
                help_text="Used for a product that has not picked a "
                          "packaging profile on the Packaging tab.")
    r = s.field(r, "Shipping you absorb", cfg.default_shipping,
                style="money", name="DefaultShipping",
                validate=number_check(0, 1000),
                help_text="Only what you pay and do not charge on. "
                          "Postage the buyer pays is not your cost - the "
                          "workbook keeps the two apart.")
    r = s.field(r, "Free shipping over", cfg.minimum_order_value / 3.0,
                style="money", name="FreeShipThreshold",
                validate=number_check(0, 100000),
                help_text="Orders above this are shipped free, so the "
                          "postage becomes your cost.")

    # ------------------------------------------------------------ 6. overhead
    r += 1
    r = s.section(r, "6. Overhead", LAST_COL)
    r = s.note(r, "Rent, insurance, the shop subscription, the website, "
                  "the accountant. Type the bills on the Overhead tab - "
                  "this is only how they get spread across your "
                  "products.",
               LAST_COL, first_col=1, height=30)
    r = s.field(r, "Spread overhead", "Per hour worked",
                name="OverheadMethod", validate=list_from("AllocationList"),
                help_text="Per hour is the fairest for slow, careful "
                          "work. Per item suits a maker who turns out "
                          "many small, similar pieces.")
    r = s.field(r, "Overhead a month (from the Overhead tab)", None,
                style="money", mode="calc", name="OverheadMonthly",
                formula="=ROUND(SUMPRODUCT(IFERROR("
                        "[[Overhead.monthly]]*1,0)),2)")
    r = s.field(r, "Overhead per hour worked", None, style="money",
                mode="calc", name="OverheadPerHour",
                formula="=IF(ProductiveHours<=0,0,"
                        "ROUND(OverheadMonthly/ProductiveHours,2))")
    r = s.field(r, "Overhead per item made", None, style="money",
                mode="calc", name="OverheadPerUnit",
                formula="=IF(UnitsPerMonth<=0,0,"
                        "ROUND(OverheadMonthly/UnitsPerMonth,2))")

    # ------------------------------------------------- 7. marketplace fees
    r += 1
    r = s.section(r, "7. Marketplace fees", LAST_COL)
    r = s.note(r, "IMPORTANT: marketplace fees change, and they differ by "
                  "country. The figures below are the defaults this "
                  "workbook shipped with - verify them against the "
                  "current official fee schedule for your shop and your "
                  "region before you trust a price. Every one of them is "
                  "editable here, and every channel can override them on "
                  "the Channels tab.",
               LAST_COL, first_col=1, height=52)
    r = s.field(r, "Listing fee (per listing)", cfg.listing_fee,
                style="money", name="ListingFee",
                validate=number_check(0, 100),
                help_text="Charged when you list and again when the "
                          "listing renews.")
    r = s.field(r, "Transaction fee", cfg.transaction_fee, style="percent",
                name="TransactionFee", validate=number_check(0, 1),
                help_text="A percentage of what the buyer pays.")
    r = s.field(r, "Payment processing %", cfg.processing_percent,
                style="percent", name="ProcessingPercent",
                validate=number_check(0, 1))
    r = s.field(r, "Payment processing fixed", cfg.processing_fixed,
                style="money", name="ProcessingFixed",
                validate=number_check(0, 100),
                help_text="The flat part - this is what quietly eats "
                          "small orders alive.")
    r = s.field(r, "Offsite ads fee", cfg.offsite_ads_fee, style="percent",
                name="OffsiteAdsFee", validate=number_check(0, 1),
                help_text="Only charged on orders that came from the "
                          "marketplace's own advertising.")
    r = s.field(r, "Share of orders from offsite ads",
                cfg.offsite_ads_share, style="percent",
                name="OffsiteAdsShare", validate=number_check(0, 1),
                help_text="Used to work out the average ads cost per "
                          "order. Check your own statements.")
    r = s.field(r, "Regulatory operating fee", cfg.regulatory_fee,
                style="percent", name="RegulatoryFee",
                validate=number_check(0, 1),
                help_text="Charged in some countries only. Set it to 0 "
                          "if your shop is not in one.")
    r = s.field(r, "Currency conversion fee", cfg.currency_fee,
                style="percent", name="CurrencyFee",
                validate=number_check(0, 1),
                help_text="Applies when you are paid in a currency other "
                          "than your bank's.")
    r = s.field(r, "Fees charged on shipping too?", cfg.fees_include_shipping,
                name="FeesOnShipping", validate=list_from("YesNoList"),
                help_text="On most marketplaces the transaction fee "
                          "applies to the postage the buyer paid as well.")
    r = s.field(r, "Fees charged on sales tax?", cfg.fees_include_tax,
                name="FeesOnTax", validate=list_from("YesNoList"))
    r = s.field(r, "Fee schedule last reviewed", cfg.today, style="date",
                name="FeesReviewed",
                help_text="Put today's date here each time you check the "
                          "official fee page. The dashboard nags you "
                          "when it goes stale.")
    r = s.field(r, "Review the fee schedule every (days)",
                cfg.fee_review_days, style="int", name="FeeReviewDays",
                validate=number_check(0, 3650, decimal=False))
    r = s.field(r, "Fees are stale", None, mode="calc", name="FeesStale",
                formula='=IF(FeesReviewed="","Never reviewed",'
                        'IF(TodayDate-FeesReviewed>FeeReviewDays,'
                        '"Check the fee schedule","OK"))')
    r = s.field(r, "Default sales channel", "Etsy", name="DefaultChannel",
                validate=list_from("ChannelList"),
                help_text="The channel a new product is priced for "
                          "until you pick another.")

    # ------------------------------------------------------ 8. pricing rules
    r += 1
    r = s.section(r, "8. Pricing rules", LAST_COL)
    r = s.field(r, "Target retail margin", cfg.target_retail_margin,
                style="percent", name="TargetRetailMargin",
                validate=number_check(0, 0.95),
                help_text="Margin, not markup. 55% margin means 55 cents "
                          "of every dollar you take is left after costs "
                          "and fees.")
    r = s.field(r, "Target wholesale margin", cfg.target_wholesale_margin,
                style="percent", name="TargetWholesaleMargin",
                validate=number_check(0, 0.95),
                help_text="Lower than retail, because the shop does the "
                          "selling. It must still be positive.")
    r = s.field(r, "Default markup on cost", cfg.default_markup,
                style="percent", name="DefaultMarkup",
                validate=number_check(0, 20),
                help_text="Markup is a percentage OF YOUR COST. 100% "
                          "markup is a 50% margin - the workbook shows "
                          "both so you never confuse them again.")
    r = s.field(r, "Wholesale is this much off retail",
                cfg.wholesale_discount, style="percent",
                name="WholesaleDiscount", validate=number_check(0, 0.9),
                help_text="The traditional handmade rule is half the "
                          "retail price. Only sustainable if the retail "
                          "price was built properly in the first place.")
    r = s.field(r, "Warn below this margin", cfg.margin_warning,
                style="percent", name="MarginWarning",
                validate=number_check(0, 0.95),
                help_text="Any product under this turns amber on the "
                          "catalog and appears on the dashboard.")
    r = s.field(r, "Round prices to", cfg.rounding_rule,
                name="RoundingRule", validate=list_from("RoundingList"),
                help_text="Charm pricing always rounds UP, so a rounded "
                          "price can never fall below the price the "
                          "maths asked for.")
    r = s.field(r, "Usual discount", cfg.default_discount, style="percent",
                name="DefaultDiscount", validate=number_check(0, 1),
                help_text="What a sale usually knocks off. The promotion "
                          "tab shows what each one costs you.")
    r = s.field(r, "Sales tax / VAT on your prices", cfg.tax_rate,
                style="percent", name="TaxRate",
                validate=number_check(0, 1),
                help_text="Leave at 0 if you are not registered or the "
                          "marketplace handles it for you.")
    r = s.field(r, "Minimum wholesale order (items)", cfg.minimum_order_qty,
                style="int", name="MinOrderQty",
                validate=number_check(0, 10000, decimal=False))
    r = s.field(r, "Minimum wholesale order (value)",
                cfg.minimum_order_value, style="money",
                name="MinOrderValue", validate=number_check(0, 1000000))
    r = s.field(r, "Review a price after (days)", cfg.price_review_days,
                style="int", name="PriceReviewDays",
                validate=number_check(0, 3650, decimal=False),
                help_text="Material costs drift upwards quietly. This is "
                          "how long a price can go unchecked.")

    # ----------------------------------------------------------- 9. targets
    r += 1
    r = s.section(r, "9. Targets", LAST_COL)
    r = s.field(r, "Revenue target a month", cfg.revenue_target_month,
                style="money", name="RevenueTargetMonth",
                validate=number_check(0, 10000000))
    r = s.field(r, "Revenue target a year", cfg.revenue_target_year,
                style="money", name="RevenueTargetYear",
                validate=number_check(0, 100000000))
    r = s.field(r, "Profit target a month", cfg.profit_target_month,
                style="money", name="ProfitTargetMonth",
                validate=number_check(0, 10000000))
    r = s.field(r, "Hourly pay you are aiming for", cfg.hourly_rate,
                style="money", name="TargetHourly",
                validate=number_check(0, 1000),
                help_text="The dashboard works out what you are actually "
                          "earning an hour and compares it with this.")

    # ------------------------------------------------------ 10. ID prefixes
    r += 1
    r = s.section(r, "10. ID prefixes", LAST_COL)
    r = s.note(r, "The letters in front of the automatic numbers. Change "
                  "them before you start entering data, not after.",
               LAST_COL, first_col=1, height=24)
    for label, value, name in [
            ("Product IDs", "P-", "ProductPrefix"),
            ("Material IDs", "M-", "MaterialPrefix"),
            ("Bundle IDs", "BN-", "BundlePrefix"),
            ("Variation IDs", "VR-", "VariationPrefix"),
            ("Batch IDs", "BT-", "BatchPrefix"),
            ("Sale IDs", "SL-", "SalePrefix"),
            ("Wholesale customer IDs", "WC-", "CustomerPrefix"),
            ("Wholesale order IDs", "WO-", "OrderPrefix"),
            ("Development IDs", "DV-", "DevPrefix"),
            ("Event IDs", "EV-", "EventPrefix"),
            ("Waste log IDs", "WL-", "WastePrefix"),
            ("Crochet sheet IDs", "CRO-", "CrochetPrefix"),
            ("Jewelry sheet IDs", "JWL-", "JewelryPrefix"),
            ("Candle sheet IDs", "CND-", "CandlePrefix"),
            ("Clothing sheet IDs", "CLO-", "ClothingPrefix"),
            ("Woodworking sheet IDs", "WDW-", "WoodPrefix"),
            ("Digital sheet IDs", "DIG-", "DigitalPrefix")]:
        r = s.field(r, label, value, name=name)

    # --------------------------------------------------------- 11. about it
    r += 1
    r = s.section(r, "11. About this workbook", LAST_COL)
    r = s.field(r, "Product", cfg.product_name, mode="calc", value_end=5)
    r = s.field(r, "Version", cfg.version, mode="calc")
    r = s.field(r, "Created by", cfg.author, mode="calc")
    r = s.field(r, "Sheet protection password", cfg.password, mode="calc",
                help_text="Review > Unprotect Sheet if you need to change "
                          "a formula.")
    r = s.field(r, "Currency", f"{cfg.currency_symbol} "
                               f"({cfg.currency_code})", mode="calc")
    r += 1
    s.note(r,
           "Changed a rate here? Everything updates straight away - every "
           "product in the catalog, both price calculators, the "
           "break-even tab and the dashboard all read these cells rather "
           "than keeping their own copy. Sales you have already recorded "
           "keep the price you actually charged.",
           LAST_COL, first_col=1, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/dashboard.py`

*420 lines · 2 functions*

> `Dashboard` - one page that answers "am I actually making money?"
>
> Twelve numbers across the top, a list of the things that need fixing
> today, your best and worst products by the measures that matter, and
> twelve months of revenue and profit. Everything reads from the
> registers, so it is never out of date and there is nothing to refresh.

| Defines | What it does |
|---------|--------------|
| `_by_rank(rank_key, field, n)` | The nth product by one of the catalog's hidden rank columns |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dashboard'), `LAST_COL` (13), `MONTHS` (12), `TOP_ROWS` (8).

````python
"""`Dashboard` - one page that answers "am I actually making money?"

Twelve numbers across the top, a list of the things that need fixing
today, your best and worst products by the measures that matter, and
twelve months of revenue and profit. Everything reads from the
registers, so it is never out of date and there is nothing to refresh.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 13
MONTHS = 12
TOP_ROWS = 8


def _by_rank(rank_key: str, field: str, n: int) -> str:
    """The nth product by one of the catalog's hidden rank columns."""
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH({n},'
            f'[[Products.{rank_key}]],0)),"")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Pricing Dashboard",
             "Your range, your margins and your month - worked out "
             "from the products, costs and sales you have already "
             "entered.", LAST_COL)
    s.set_col(0, 0, 24)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ---------------------------------------------------------- the header
    row = 4
    s.merge(row, 0, 0, "Shop", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    ws.write_formula(row, 1, "=ShopName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    ws.write_formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 6, 6, "Month key", fmts.chrome("label_right"))
    ws.write_formula(row, 7, '=TEXT(TodayDate,"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$H${row + 1}")
    s.merge(row, 8, 9, "Today", fmts.chrome("label_right"))
    s.merge(row, 10, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 10, "=TodayDate", fmts.calc("date"))
    row += 2

    # ------------------------------------------------------------ the KPIs
    money = cfg.currency_symbol
    tiles_1 = [
        ("Products priced",
         '=COUNTIF([[Products.name]],"<>")', "int",
         "in the catalog"),
        ("Average margin",
         '=IFERROR(AVERAGE([[Products.margin]]),"")', "percent",
         f"target {cfg.target_retail_margin:.0%}"),
        ("Average you earn an hour",
         '=IFERROR(AVERAGE([[Products.wage_hour]]),"")', "money",
         f"your rate is {money}{cfg.hourly_rate:,.2f}"),
        ("Products below target",
         '=COUNTIFS([[Products.margin]],"<"&MarginWarning,'
         '[[Products.name]],"<>")', "int",
         "margin under your warning level"),
    ]
    tiles_2 = [
        ("Revenue this month",
         '=ROUND(SUMIFS([[Sales.revenue]],[[Sales.month]],DashKey),2)',
         "money0", f"target {money}{cfg.revenue_target_month:,.0f}"),
        ("Profit this month",
         '=ROUND(SUMIFS([[Sales.profit]],[[Sales.month]],DashKey),2)',
         "money0", "after fees and postage"),
        ("Margin this month",
         '=IFERROR(SUMIFS([[Sales.profit]],[[Sales.month]],DashKey)/'
         'SUMIFS([[Sales.revenue]],[[Sales.month]],DashKey),"")',
         "percent", "what you actually kept"),
        ("Items sold this month",
         '=SUMIFS([[Sales.qty]],[[Sales.month]],DashKey)', "int",
         "across every channel"),
    ]
    tiles_3 = [
        ("Overhead a month",
         "=ROUND(OverheadMonthly,2)", "money0",
         "studio, insurance, software"),
        ("Covered by this month",
         '=IFERROR(SUMIFS([[Sales.profit]],[[Sales.month]],DashKey)/'
         'MAX(0.01,OverheadMonthly),"")', "percent0",
         "100% means the studio has paid for itself"),
        ("Value of your materials",
         '=ROUND(SUM([[Materials.stock_value]]),2)', "money0",
         "money sitting on the shelf"),
        ("Hours of work sold",
         '=ROUND(SUMIFS([[Sales.hours]],[[Sales.month]],DashKey),1)',
         "number1", "this month"),
    ]
    for tiles in (tiles_1, tiles_2, tiles_3):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, formula, style=style, span=3,
                  note=note)
        row += 4

    # --------------------------------------------------- needs attention --
    row = s.section(row, "Needs your attention", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [26, 11, 46, 20])
    alerts = [
        ("Products with no price",
         '=COUNTIFS([[Products.name]],"<>",[[Products.price]],"")',
         '="Listed in the catalog with nothing to sell them at"',
         "Products"),
        ("Products that lose money",
         '=COUNTIF([[Products.alert]],"Loss*")',
         '="The price does not cover the cost and the fees"',
         "Products"),
        ("Products below minimum wage",
         '=COUNTIF([[Products.alert]],"Below*")',
         '="You would earn less than "&TEXT(MinimumWage,"0.00")&'
         '" an hour making these"', "Products"),
        ("Products under your margin warning",
         '=COUNTIFS([[Products.margin]],"<"&MarginWarning,'
         '[[Products.name]],"<>")',
         '="Thinner than "&TEXT(MarginWarning,"0%")&" - one bad month '
         'away from nothing"', "Retail Calculator"),
        ("Products with no costs entered",
         '=COUNTIF([[Products.alert]],"Add - no costs*")',
         '="Nothing on the Material Usage or Labor tabs for these"',
         "Material Usage"),
        ("Materials with no price",
         '=COUNTIF([[Materials.alert]],"Add*")',
         '="Every product using them is under-costed"', "Materials"),
        ("Materials running low",
         '=COUNTIF([[Materials.status]],"Low stock")+'
         'COUNTIF([[Materials.status]],"Out of stock")',
         '="Below the reorder level or less than "&LowStockDays&'
         '" days of cover"', "Materials"),
        ("Prices going stale",
         '=COUNTIF([[Products.alert]],"Review*")',
         '="Not changed in "&PriceReviewDays&" days, while your costs '
         'moved"', "Price History"),
        ("Fee settings to check",
         '=IF(FeesStale="Yes",1,0)',
         '="Marketplace fees change. Yours were last reviewed "&'
         'TEXT(FeesReviewed,"d mmm yyyy")', "Channels"),
        ("Orders that lost money",
         '=COUNTIFS([[Sales.profit]],"<0")',
         '="Real orders where the money went the wrong way"',
         "Sales"),
        ("Batches costing more than the estimate",
         '=COUNTIF([[Batches.alert]],"Check - costing*")',
         '="The catalog is optimistic about these"', "Batches"),
        ("Wholesale orders below target margin",
         '=COUNTIF([[Wholesale Orders.alert]],"Low margin*")',
         '="The discount is deeper than the product can carry"',
         "Wholesale Orders"),
        ("Unpaid wholesale invoices",
         '=ROUND(SUMIFS([[Wholesale Orders.order_value]],'
         '[[Wholesale Orders.paid]],"No"),0)',
         '="Money shops owe you right now"', "Wholesale Customers"),
        ("Avoidable waste this year",
         '=ROUND(SUMIFS([[Waste Log.cost]],[[Waste Log.avoidable]],'
         '"Yes"),0)',
         '="Material you paid for and did not sell"', "Waste Log"),
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

    # ------------------------------------------------------ best and worst
    row = s.section(row, "What is working", LAST_COL)
    row = s.table_head(row, 0, ["Most profit a sale", "Profit",
                                "Best margin", "Margin",
                                "Best paid an hour", "An hour",
                                "Most sold", "Sold"],
                       [24, 10, 24, 10, 24, 10, 24, 9])
    top_first = row
    for n in range(1, TOP_ROWS + 1):
        pairs = [
            (0, "profit_rank", "name", "text"),
            (1, "profit_rank", "profit", "money"),
            (2, "margin_rank", "name", "text"),
            (3, "margin_rank", "margin", "percent"),
            (4, "profit_hour_rank", "name", "text"),
            (5, "profit_hour_rank", "profit_hour", "money"),
            (6, "revenue_rank", "name", "text"),
            (7, "revenue_rank", "sold", "int"),
        ]
        for col, rank_key, field, style in pairs:
            ws.write_formula(row, col,
                             s.reg.render("=" + _by_rank(rank_key, field, n)),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(top_first, 1, row - 1, 1,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    ws.conditional_format(top_first, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    ws.conditional_format(top_first, 5, row - 1, 5,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    row += 1
    s.merge(row, 0, LAST_COL,
            "The four columns rarely agree. The piece that makes the "
            "most profit is often not the one that pays best per hour "
            "- and the hourly column is the one that decides what you "
            "should be making more of.", fmts.chrome("note"))
    row += 2

    # ------------------------------------------------- by category --------
    row = s.section(row, "By category", LAST_COL)
    row = s.table_head(row, 0, ["Category", "Products", "Average cost",
                                "Average price", "Average margin",
                                "Sold", "Revenue", "Profit"],
                       [22, 10, 13, 13, 14, 10, 13, 13])
    cat_first = row
    for n in range(1, 11):
        cat = f'IFERROR(INDEX(CategoryList,{n}),"")'
        ws.write_formula(row, 0, f"={cat}", fmts.calc("text"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IF({cat}="","",COUNTIFS([[Products.category]],'
                         f'{cat}))'), fmts.calc("int"))
        for col, field, style in ((2, "total_cost", "money"),
                                  (3, "price", "money"),
                                  (4, "margin", "percent")):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IF({cat}="","",IFERROR(AVERAGEIFS('
                             f'[[Products.{field}]],[[Products.category]],'
                             f'{cat}),""))'), fmts.calc(style))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=IF({cat}="","",SUMIFS([[Products.sold]],'
                         f'[[Products.category]],{cat}))'),
            fmts.calc("int"))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=IF({cat}="","",ROUND(SUMIFS('
                         f'[[Products.revenue]],[[Products.category]],'
                         f'{cat}),2))'), fmts.calc("money0"))
        ws.write_formula(
            row, 7,
            s.reg.render(f'=IF({cat}="","",ROUND(SUMIFS('
                         f'[[Products.actual_profit]],'
                         f'[[Products.category]],{cat}),2))'),
            fmts.calc("money0"))
        row += 1
    ws.conditional_format(cat_first, 4, row - 1, 4,
                          {"type": "3_color_scale"})
    ws.conditional_format(cat_first, 7, row - 1, 7,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    row += 2

    # ----------------------------------------------------- twelve months --
    row = s.section(row, "The last twelve months", LAST_COL)
    s.write(row, 0, "Month", fmts.chrome("header_left"))

    for i in range(MONTHS):
        back = MONTHS - 1 - i
        ws.write_formula(
            row, i + 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{back},1),'
            f'"mmm yy")', fmts.chrome("header"))
        ws.write_formula(
            row + 1, i + 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{back},1),'
            f'"yyyy-mm")', fmts.chrome("header", font_size=8))
    s.write(row + 1, 0, "Key", fmts.chrome("label"))
    key_row = row + 1
    row += 2
    lines = [
        ("Revenue", "SUMIFS([[Sales.revenue]],[[Sales.month]],{key})",
         "money0"),
        ("Cost of goods", "SUMIFS([[Sales.cogs]],[[Sales.month]],{key})",
         "money0"),
        ("Fees", "SUMIFS([[Sales.fees]],[[Sales.month]],{key})", "money0"),
        ("Profit", "SUMIFS([[Sales.profit]],[[Sales.month]],{key})",
         "money0"),
        ("Items", "SUMIFS([[Sales.qty]],[[Sales.month]],{key})", "int"),
        ("Hours sold", "SUMIFS([[Sales.hours]],[[Sales.month]],{key})",
         "number1"),
    ]
    band = {}
    for label, body, style in lines:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        for i in range(MONTHS):
            cell = f"{chr(66 + i)}${key_row + 1}"
            ws.write_formula(row, i + 1,
                             s.reg.render("=ROUND(" +
                                          body.format(key=cell) + ",2)"),
                             fmts.calc(style))
        band[label] = row
        row += 1
    s.write(row, 0, "Margin", fmts.chrome("label_bold"))
    for i in range(MONTHS):
        col = chr(66 + i)
        ws.write_formula(
            row, i + 1,
            f'=IFERROR({col}{band["Profit"] + 1}/{col}'
            f'{band["Revenue"] + 1},"")', fmts.calc("percent0"))
    margin_row = row
    row += 1
    for label in ("Revenue", "Profit"):
        ws.add_sparkline(band[label], MONTHS + 1, {
            "range": (f"{quote_sheet(NAME)}!B{band[label] + 1}:"
                      f"M{band[label] + 1}"),
            "type": "column",
            "series_color": "#8C4A32" if label == "Revenue" else "#5E7B5A",
        })
    ws.conditional_format(band["Revenue"], 1, band["Revenue"], MONTHS,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    ws.conditional_format(margin_row, 1, margin_row, MONTHS,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("warn")})
    s.set_col(MONTHS + 1, MONTHS + 1, 12)
    row += 1
    s.merge(row, 0, LAST_COL,
            "Blank months are months with nothing logged on the Sales "
            "tab, not months with no sales.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------------------- channels --
    row = s.section(row, "Where the money comes from", LAST_COL)
    row = s.table_head(row, 0, ["Channel", "Orders", "Revenue", "Fees",
                                "Fee rate", "Profit", "Margin",
                                "Share of revenue"],
                       [22, 10, 13, 12, 11, 13, 10, 16])
    ch_first = row
    for n in range(1, 11):
        name = f'IFERROR(INDEX([[Channels.name]],{n})&"","")'
        ws.write_formula(row, 0, s.reg.render(f"={name}"),
                         fmts.calc("text"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IF({name}="","",COUNTIFS([[Sales.channel]],'
                         f'{name}))'), fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=IF({name}="","",ROUND(SUMIFS('
                         f'[[Sales.revenue]],[[Sales.channel]],{name}),2))'),
            fmts.calc("money0"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF({name}="","",ROUND(SUMIFS([[Sales.fees]],'
                         f'[[Sales.channel]],{name}),2))'),
            fmts.calc("money0"))
        ws.write_formula(
            row, 4,
            f'=IFERROR(D{row + 1}/C{row + 1},"")', fmts.calc("percent"))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=IF({name}="","",ROUND(SUMIFS('
                         f'[[Sales.profit]],[[Sales.channel]],{name}),2))'),
            fmts.calc("money0"))
        ws.write_formula(
            row, 6, f'=IFERROR(F{row + 1}/C{row + 1},"")',
            fmts.calc("percent"))
        ws.write_formula(
            row, 7,
            s.reg.render(f'=IFERROR(C{row + 1}/SUM([[Sales.revenue]]),"")'),
            fmts.calc("percent"))
        row += 1
    ws.conditional_format(ch_first, 7, row - 1, 7,
                          {"type": "data_bar", "bar_color": "#8C4A32"})
    ws.conditional_format(ch_first, 4, row - 1, 4,
                          {"type": "cell", "criteria": ">", "value": 0.25,
                           "format": fmts.cf("warn")})
    row += 2

    # ------------------------------------------------------- what to do --
    row = s.section(row, "If you only do one thing this week", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.calc("text"))
    ws.write_formula(
        row, 0,
        s.reg.render(
            '=IF(COUNTIF([[Products.alert]],"Loss*")>0,'
            '"Fix the products that lose money. There are "&'
            'COUNTIF([[Products.alert]],"Loss*")&" of them on the '
            'Product Catalog - every sale of those makes things worse.",'
            'IF(COUNTIF([[Products.alert]],"Below*")>0,'
            '"Re-price the "&COUNTIF([[Products.alert]],"Below*")&'
            '" products that pay you under minimum wage. Start with '
            'the Retail Price Calculator.",'
            'IF(COUNTIFS([[Products.name]],"<>",[[Products.price]],"")>0,'
            '"Price the products that have no price yet.",'
            'IF(FeesStale="Yes","Check your marketplace fee rates on '
            'the Channels tab - they were last reviewed "&'
            'TEXT(FeesReviewed,"d mmm yyyy")&".",'
            'IF(COUNTIF([[Products.alert]],"Review*")>0,'
            '"Review the prices that have not moved in "&'
            'PriceReviewDays&" days. Your costs have.",'
            '"Everything is priced, nothing is losing money and your '
            'fees are current. Log this month\'s sales and look at '
            'what pays you best an hour.")))))'),
        fmts.calc("text"))
    s.row_height(row, 34)
    row += 2

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/retail_calc.py`

*439 lines · 2 functions*

> `Retail Calculator` - the page the whole workbook exists for.
>
> Pick a product, or type the costs in by hand. The page adds materials,
> labour, packaging, postage and overhead, works out what the
> marketplace takes, and solves for the price that leaves the margin you
> asked for:
>
>     Price = (Cost + fixed fees) / (1 - target margin - fee rate)
>
> Then it shows you the same product at a dozen other prices, so the
> number you choose is a decision rather than a hope.

| Defines | What it does |
|---------|--------------|
| `_pick(field, default)` | Look a field up on the catalog by the product name chosen above |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Retail Calculator'), `LAST_COL` (11), `LADDER` (10 items), `STEPS` (10 items).

````python
"""`Retail Calculator` - the page the whole workbook exists for.

Pick a product, or type the costs in by hand. The page adds materials,
labour, packaging, postage and overhead, works out what the
marketplace takes, and solves for the price that leaves the margin you
asked for:

    Price = (Cost + fixed fees) / (1 - target margin - fee rate)

Then it shows you the same product at a dozen other prices, so the
number you choose is a decision rather than a hope.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check
from .common import round_price

NAME = "Retail Calculator"
LAST_COL = 11

#: margins offered in the ladder
LADDER = [0.25, 0.30, 0.35, 0.40, 0.45, 0.50, 0.55, 0.60, 0.65, 0.70]
#: price points shown around the suggested price
STEPS = [-0.30, -0.20, -0.10, -0.05, 0.0, 0.05, 0.10, 0.20, 0.30, 0.50]


def _pick(field: str, default: str = "0") -> str:
    """Look a field up on the catalog by the product name chosen above."""
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH(RcProduct,'
            f'[[Products.name]],0)),{default})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    hours = s.reg.render(_pick("hours", "0"))

    s.banner("Retail Price Calculator",
             "Everything that goes into the price, and the number it "
             "comes out at. Fill in the white cells.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 30, fmts.chrome("label"))
    s.set_col(2, 2, 11, fmts.chrome("label"))
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 15)
    s.set_col(5, 5, 15)
    s.set_col(6, LAST_COL, 14, fmts.chrome("note"))

    r = 4

    # ----------------------------------------------- 1. which product --
    r = s.section(r, "1. What are you pricing?", LAST_COL)
    r = s.note(r, "Choose a product from your catalog and every cost "
                  "below fills itself in. Or leave it blank and type "
                  "the costs by hand - nothing on this page changes "
                  "your saved data.",
               LAST_COL, first_col=1, height=26)
    r = s.field(r, "Product", "Willow cable beanie", name="RcProduct",
                validate=list_from("ProductNames"), value_end=5,
                help_text="From the Product Catalog. Clear the cell to "
                          "price something that is not listed yet.")
    r = s.field(r, "Where are you selling it?", "Etsy",
                name="RcChannel", validate=list_from("ChannelList"),
                value_end=5,
                help_text="The fees come from the Channels tab, which "
                          "you can edit.")
    r = s.field(r, "Target margin", None, style="percent",
                name="RcTarget", validate=number_check(0, 0.95),
                formula="=TargetRetailMargin",
                help_text="Profit as a share of the price. Change it "
                          "here to test, or on Settings to change it "
                          "everywhere.")
    r = s.field(r, "Quantity you make at a time", 1, style="int",
                name="RcBatch", validate=number_check(1, 10000,
                                                      decimal=False),
                help_text="Used for the per unit setup time below.")
    r += 1

    # ------------------------------------------------- 2. the costs ----
    r = s.section(r, "2. What it costs you to make", LAST_COL)
    r = s.note(r, "The middle column is what the catalog says. Type "
                  "anything in the override column to try a different "
                  "number - an override always wins.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Cost", "From the catalog", "Your override",
                            "Used", "What it is"],
                     [30, 16, 15, 15, 44])
    cost_rows = [
        ("Materials", _pick("mat_cost"), "RcMat", "money",
         "Everything that ends up in the piece, waste included."),
        ("Labour", _pick("labour_adj"), "RcLab", "money",
         "Your hours at your rate - the cost most makers leave out."),
        ("Consumables", _pick("consumables"), "RcCons", "money",
         "Glue, thread, sandpaper, gas. Small, real, forgotten."),
        ("Packaging", _pick("pack_cost"), "RcPack", "money",
         "Box, tissue, card, label, tape."),
        ("Postage you absorb", _pick("ship_cost"), "RcShip", "money",
         "What the label costs you less what the buyer pays."),
        ("Other costs", "0", "RcOther", "money",
         "Anything else this piece carries - a pattern licence, a "
         "stamp, a fee."),
    ]
    first_cost = r
    for label, formula, name, style, what in cost_rows:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render("=" + formula),
                         fmts.calc(style))
        s.write(r, 3, None, fmts.input(style))
        ws.data_validation(r, 3, r, 3, number_check(0, 1000000))
        book.define_name(name, f"='{NAME}'!$D${r + 1}")
        ws.write_formula(r, 4, f'=IF(D{r + 1}<>"",D{r + 1},C{r + 1})',
                         fmts.calc(style))
        book.define_name(name + "Used", f"='{NAME}'!$E${r + 1}")
        s.merge(r, 5, LAST_COL, what, fmts.chrome("note"))
        r += 1
    direct_first, direct_last = first_cost, r - 1
    s.write(r, 1, "Direct cost", fmts.chrome("total_label"))
    ws.write_formula(r, 4,
                     f'=ROUND(SUM(E{direct_first + 1}:E{direct_last + 1}),2)',
                     fmts.chrome("total_value",
                                 num_format=s.num_format("money")))
    book.define_name("RcDirect", f"='{NAME}'!$E${r + 1}")
    s.merge(r, 5, LAST_COL, "Materials + labour + packaging + postage - "
                            "everything except overhead.",
            fmts.chrome("note"))
    r += 2

    # ----------------------------------------------- 3. the overhead ---
    r = s.section(r, "3. Overhead - the costs that are not in the piece",
                  LAST_COL)
    r = s.note(r, "Studio rent, insurance, your website, the shop "
                  "subscription. The Overhead tab works out how it is "
                  "shared; this is that share applied to this product.",
               LAST_COL, first_col=1, height=26)
    r = s.field(r, "How overhead is shared", None, style="text",
                mode="calc", formula="=OverheadMethod", value_end=5,
                help_text="Change the method on the Settings tab.")
    r = s.field(r, "Overhead on this product", None, style="money",
                mode="calc", name="RcOverhead", value_end=5,
                formula='=IF(OverheadMethod="Per hour worked",'
                        'ROUND(OverheadPerHour*' + _pick("hours") + ',2),'
                        'IF(OR(OverheadMethod="Per unit made",'
                        'OverheadMethod="Per order"),OverheadPerUnit,0))',
                help_text="With 'Share of revenue' the overhead goes "
                          "into the price as a percentage instead - "
                          "see the next section.")
    r = s.field(r, "Total cost", None, style="money", mode="calc",
                name="RcCost", value_end=5,
                formula="=ROUND(RcDirect+RcOverhead,2)",
                help_text="This is C in the pricing formula.")
    r += 1

    # ------------------------------------------- 4. what the shop takes
    r = s.section(r, "4. What the marketplace takes", LAST_COL)
    r = s.note(r, "Percentage fees and fixed fees behave completely "
                  "differently, which is why adding a flat 'Etsy 20%' "
                  "to your cost gets the price wrong. Percentages come "
                  "out of the price you are solving for; fixed fees "
                  "are added to the cost.",
               LAST_COL, first_col=1, height=30)
    r = s.field(r, "Percentage fees", None, style="percent", mode="calc",
                name="RcFeePct", value_end=5,
                formula='=IFERROR(INDEX([[Channels.pct_total]],'
                        'MATCH(RcChannel,[[Channels.name]],0)),0)',
                help_text="Listing share, transaction, payment "
                          "processing, ads, regulatory - all editable "
                          "on the Channels tab.")
    r = s.field(r, "Fixed fees", None, style="money", mode="calc",
                name="RcFeeFixed", value_end=5,
                formula='=IFERROR(INDEX([[Channels.fixed_total]],'
                        'MATCH(RcChannel,[[Channels.name]],0)),0)',
                help_text="The listing fee and the fixed part of the "
                          "payment processing charge.")
    r = s.field(r, "Overhead as a share of revenue", None, style="percent",
                mode="calc", name="RcOhRate", value_end=5,
                formula='=IF(OverheadMethod="Share of revenue",'
                        'IF(RevenueTargetMonth<=0,0,OverheadMonthly/'
                        'RevenueTargetMonth),0)',
                help_text="Only used with the 'Share of revenue' "
                          "method - it comes out of the price, like a "
                          "fee.")
    r += 1

    # ------------------------------------------------ 5. the answer ----
    r = s.section(r, "5. The price", LAST_COL)
    denom_row = r
    s.write(r, 1, "Denominator", fmts.chrome("label_bold"))
    ws.write_formula(r, 3, "=1-RcTarget-RcFeePct-RcOhRate",
                     fmts.calc("number"))
    book.define_name("RcDenom", f"='{NAME}'!$D${denom_row + 1}")
    s.merge(r, 4, LAST_COL,
            "1 minus the margin you want, minus the percentage fees, "
            "minus the overhead share. If this is zero or less, no "
            "price works and the maths says so.",
            fmts.chrome("note"))
    r += 1
    s.write(r, 1, "Suggested price", fmts.chrome("label_bold"))
    ws.write_formula(
        r, 3,
        '=IF(RcDenom<=0,"No price works",ROUND((RcCost+RcFeeFixed)/'
        'RcDenom,2))',
        fmts.calc("money"))
    book.define_name("RcRaw", f"='{NAME}'!$D${r + 1}")
    s.merge(r, 4, LAST_COL,
            "(Cost + fixed fees) divided by the denominator. This is "
            "the price that leaves exactly the margin you asked for.",
            fmts.chrome("note"))
    r += 1
    s.write(r, 1, "Rounded for the listing", fmts.chrome("label_bold"))
    ws.write_formula(
        r, 3,
        '=IF(ISNUMBER(RcRaw),ROUND(' + round_price("RcRaw") + ',2),"")',
        fmts.calc("money"))
    book.define_name("RcPrice", f"='{NAME}'!$D${r + 1}")
    s.merge(r, 4, LAST_COL,
            "Using the rounding rule on Settings. Rounding up is free "
            "margin; rounding down is a gift.",
            fmts.chrome("note"))
    r += 2

    # ---------------------------------------------------- the KPI tiles
    tiles = [
        ("Price", "=IF(RcPrice=\"\",\"\",RcPrice)", "money",
         "what to charge"),
        ("Fees on that price",
         '=IF(RcPrice="","",ROUND(RcPrice*RcFeePct+RcFeeFixed,2))',
         "money", "what the channel takes"),
        ("Profit",
         '=IF(RcPrice="","",ROUND(RcPrice-RcCost-(RcPrice*RcFeePct+'
         'RcFeeFixed),2))', "money", "what you keep"),
        ("Margin",
         '=IF(OR(RcPrice="",RcPrice=0),"",(RcPrice-RcCost-'
         '(RcPrice*RcFeePct+RcFeeFixed))/RcPrice)', "percent",
         "profit over price"),
    ]
    for i, (label, formula, style, note) in enumerate(tiles):
        s.kpi(r, 1 + i * 2, label, formula, style=style, span=2, note=note)
    r += 3
    tiles_2 = [
        ("Markup",
         '=IF(OR(RcPrice="",RcCost=0),"",(RcPrice-RcCost-'
         '(RcPrice*RcFeePct+RcFeeFixed))/RcCost)', "percent",
         "profit over cost"),
        ("Hours in it", "=" + _pick("hours", '0'), "number",
         "from the catalog"),
        ("You earn an hour",
         '=IF(OR(' + _pick("hours", "0") + '=0,RcPrice=""),"",'
         'ROUND((RcPrice-RcCost-(RcPrice*RcFeePct+RcFeeFixed)+'
         'RcLabUsed)/' + _pick("hours", "1") + ',2))', "money",
         "profit plus your wage"),
        ("Wholesale at this price",
         '=IF(RcPrice="","",ROUND(RcPrice*(1-WholesaleDiscount),2))',
         "money", "half retail, by default"),
    ]
    for i, (label, formula, style, note) in enumerate(tiles_2):
        s.kpi(r, 1 + i * 2, label, formula, style=style, span=2, note=note)
    r += 4

    # --------------------------------------------------- the workings --
    r = s.section(r, "6. How that number was worked out", LAST_COL)
    r = s.table_head(r, 1, ["Step", "Value", "Where it comes from"],
                     [30, 15, 60])
    steps = [
        ("Materials, labour and the rest", "=RcDirect", "money",
         "Section 2 above"),
        ("Overhead", "=RcOverhead", "money",
         "The Overhead tab, shared by your chosen method"),
        ("C - total cost", "=RcCost", "money", "The two lines above"),
        ("F - fixed fees", "=RcFeeFixed", "money",
         "Channels tab: listing fee plus the fixed processing charge"),
        ("r - target margin", "=RcTarget", "percent",
         "Your target, from Settings or the box above"),
        ("f - percentage fees", "=RcFeePct", "percent",
         "Channels tab: everything charged as a percentage"),
        ("1 - r - f", "=RcDenom", "number",
         "The share of the price that is left to cover the cost"),
        ("(C + F) / (1 - r - f)", "=RcRaw", "money",
         "The price that leaves exactly your target margin"),
        ("Rounded", "=RcPrice", "money",
         "Your rounding rule from the Settings tab"),
    ]
    for label, formula, style, where in steps:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, formula, fmts.calc(style))
        s.merge(r, 3, LAST_COL, where, fmts.chrome("note"))
        r += 1
    r += 1

    # ------------------------------------------------- the margin ladder
    r = s.section(r, "7. The same product at other margins", LAST_COL)
    r = s.note(r, "What you would have to charge for each level of "
                  "margin, and what you would earn an hour at it. The "
                  "row matching your target is the suggestion above.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Target margin", "Price", "Fees", "Profit",
                            "You earn an hour", "Markup"],
                     [16, 14, 12, 13, 17, 12])
    ladder_first = r
    for margin in LADDER:
        denom = f"(1-{margin}-RcFeePct-RcOhRate)"
        price = f"(RcCost+RcFeeFixed)/{denom}"
        s.write(r, 1, margin, fmts.calc("percent"))
        ws.write_formula(r, 2, f'=IF({denom}<=0,"",ROUND({price},2))',
                         fmts.calc("money"))
        ws.write_formula(
            r, 3, f'=IF(C{r + 1}="","",ROUND(C{r + 1}*RcFeePct+'
                  f'RcFeeFixed,2))', fmts.calc("money"))
        ws.write_formula(
            r, 4, f'=IF(C{r + 1}="","",ROUND(C{r + 1}-RcCost-'
                  f'D{r + 1},2))', fmts.calc("money"))
        ws.write_formula(
            r, 5, '=IF(OR(C{0}="",{1}=0),"",ROUND((E{0}+RcLabUsed)/'
                  '{1},2))'.format(r + 1, hours),
            fmts.calc("money"))
        ws.write_formula(
            r, 6, f'=IF(OR(C{r + 1}="",RcCost=0),"",E{r + 1}/RcCost)',
            fmts.calc("percent"))
        r += 1
    ws.conditional_format(ladder_first, 2, r - 1, 2,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    ws.conditional_format(ladder_first, 5, r - 1, 5,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    ws.conditional_format(ladder_first, 5, r - 1, 5,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetHourly",
                           "format": fmts.cf("good")})
    r += 1

    # ----------------------------------------------- what if you charged
    r = s.section(r, "8. What if you charged something else?", LAST_COL)
    r = s.note(r, "The suggested price, moved up and down. This is the "
                  "table to look at before you knock a few dollars off "
                  "because a price feels high.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Price", "Versus suggested", "Fees", "Profit",
                            "Margin", "Markup", "You earn an hour",
                            "Verdict"],
                     [14, 14, 12, 12, 10, 10, 16, 26])
    step_first = r
    for step in STEPS:
        price = f'IF(ISNUMBER(RcRaw),ROUND(RcRaw*{1 + step},2),"")'
        ws.write_formula(r, 1, f"={price}", fmts.calc("money"))
        s.write(r, 2, step, fmts.calc("percent0"))
        ws.write_formula(
            r, 3, f'=IF(B{r + 1}="","",ROUND(B{r + 1}*RcFeePct+'
                  f'RcFeeFixed,2))', fmts.calc("money"))
        ws.write_formula(
            r, 4, f'=IF(B{r + 1}="","",ROUND(B{r + 1}-RcCost-'
                  f'D{r + 1},2))', fmts.calc("money"))
        ws.write_formula(
            r, 5, f'=IF(OR(B{r + 1}="",B{r + 1}=0),"",E{r + 1}/'
                  f'B{r + 1})', fmts.calc("percent"))
        ws.write_formula(
            r, 6, f'=IF(OR(B{r + 1}="",RcCost=0),"",E{r + 1}/RcCost)',
            fmts.calc("percent"))
        ws.write_formula(
            r, 7, '=IF(OR(B{0}="",{1}=0),"",ROUND((E{0}+RcLabUsed)/'
                  '{1},2))'.format(r + 1, hours),
            fmts.calc("money"))
        ws.write_formula(
            r, 8,
            f'=IF(B{r + 1}="","",IF(E{r + 1}<0,"Loss - below cost",'
            f'IF(F{r + 1}<MarginWarning,"Too thin",'
            f'IF(F{r + 1}<RcTarget,"Under your target",'
            f'IF(F{r + 1}>RcTarget+0.1,"Above target - test it",'
            f'"On target")))))', fmts.calc("text"))
        r += 1
    ws.conditional_format(step_first, 4, r - 1, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(step_first, 8, r - 1, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "Loss", "format": fmts.cf("bad")})
    ws.conditional_format(step_first, 8, r - 1, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "On target", "format": fmts.cf("good")})
    r += 1

    # -------------------------------------------------- the sanity check
    r = s.section(r, "9. Does this price make sense?", LAST_COL)
    checks = [
        ("Covers your costs",
         '=IF(RcPrice="","Pick a product or type the costs",'
         'IF(RcPrice>RcCost+RcPrice*RcFeePct+RcFeeFixed,'
         '"Yes - there is profit in it","No - this price loses money"))'),
        ("Pays you properly",
         '=IF(OR(RcPrice="",' + _pick("hours", "0") + '=0),'
         '"Add the labour minutes on the Labor tab",'
         'IF((RcPrice-RcCost-(RcPrice*RcFeePct+RcFeeFixed)+RcLabUsed)/'
         + _pick("hours", "1") + '>=MinimumWage,'
         '"Yes - above minimum wage",'
         '"No - you are paying to work"))'),
        ("Works for wholesale",
         '=IF(RcPrice="","-",IF(RcPrice*(1-WholesaleDiscount)>RcCost,'
         '"Yes - there is room for a stockist",'
         '"No - wholesale would sell below cost"))'),
        ("Against the market",
         '=IF(OR(RcPrice="",' + _pick("comp_price", '""') + '=""),'
         '"No competitor research for this product yet",'
         '"The shops you logged average "&TEXT(' +
         _pick("comp_price", "0") + ',"0.00")&" - you would be "&'
         'TEXT(ABS(RcPrice-' + _pick("comp_price", "0") +
         '),"0.00")&IF(RcPrice>' + _pick("comp_price", "0") +
         '," dearer"," cheaper"))'),
        ("Fee settings",
         '=IF(FeesStale="Yes","Check - your fee rates were last '
         'reviewed "&TEXT(FeesReviewed,"d mmm yyyy")&". Marketplace '
         'fees change; verify the current official fee schedule.",'
         '"Fee rates reviewed "&TEXT(FeesReviewed,"d mmm yyyy"))'),
    ]
    check_first = r
    for label, formula in checks:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc("text"))
        r += 1
    for word, kind in (("No -", "bad"), ("Yes -", "good"),
                       ("Check", "warn")):
        ws.conditional_format(check_first, 2, r - 1, 2,
                              {"type": "text", "criteria": "containing",
                               "value": word, "format": fmts.cf(kind)})
    r += 1

    s.note(r,
           "Happy with the number? Put it in the 'Your price' column "
           "on the Product Catalog and log the change on the Price "
           "History tab. Nothing on this page is saved - it is a "
           "calculator, not a record.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/wholesale_calc.py`

*316 lines · 2 functions*

> `Wholesale Calculator` - can you afford to sell at half price?
>
> Wholesale is simple arithmetic that catches a lot of makers out. A
> shop expects around 50% off your retail price, sells it at your
> retail price, and keeps the difference. That only works if your
> retail price was built on your real costs in the first place - which
> is exactly what the rest of this workbook is for.
>
> This page works it both ways: what a given retail price leaves you at
> wholesale, and what retail price you would need for wholesale to work
> at the margin you want.

| Defines | What it does |
|---------|--------------|
| `_pick(field, default)` |  |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Wholesale Calculator'), `LAST_COL` (11), `DISCOUNTS` (7 items).

````python
"""`Wholesale Calculator` - can you afford to sell at half price?

Wholesale is simple arithmetic that catches a lot of makers out. A
shop expects around 50% off your retail price, sells it at your
retail price, and keeps the difference. That only works if your
retail price was built on your real costs in the first place - which
is exactly what the rest of this workbook is for.

This page works it both ways: what a given retail price leaves you at
wholesale, and what retail price you would need for wholesale to work
at the margin you want.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check
from .common import round_price

NAME = "Wholesale Calculator"
LAST_COL = 11

#: discounts a shop might ask for
DISCOUNTS = [0.30, 0.35, 0.40, 0.45, 0.50, 0.55, 0.60]


def _pick(field: str, default: str = "0") -> str:
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH(WcProduct,'
            f'[[Products.name]],0)),{default})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Wholesale Price Calculator",
             "What a stockist costs you, what you need to charge, and "
             "whether the order is worth the fortnight.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 32, fmts.chrome("label"))
    s.set_col(2, 2, 11, fmts.chrome("label"))
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 15)
    s.set_col(5, 5, 15)
    s.set_col(6, LAST_COL, 14, fmts.chrome("note"))

    r = 4

    # ------------------------------------------------------- the inputs
    r = s.section(r, "1. The product and the deal", LAST_COL)
    r = s.field(r, "Product", "Willow cable beanie", name="WcProduct",
                validate=list_from("ProductNames"), value_end=5,
                help_text="From the Product Catalog.")
    r = s.field(r, "Discount the shop wants", None, style="percent",
                name="WcDiscount", validate=number_check(0, 0.9),
                formula="=WholesaleDiscount",
                help_text="50% is the usual ask. 40% is common for "
                          "small independents, 60% for distributors.")
    r = s.field(r, "Units in the order", 24, style="int", name="WcQty",
                validate=number_check(1, 100000, decimal=False),
                help_text="Used for the order totals and the hours "
                          "further down.")
    r = s.field(r, "Carriage you pay", 0, style="money",
                name="WcCarriage", validate=number_check(0, 100000),
                help_text="One box to the shop, not one box per item "
                          "- which is why wholesale postage is so "
                          "much cheaper per piece.")
    r = s.field(r, "Target wholesale margin", None, style="percent",
                name="WcTarget", validate=number_check(0, 0.95),
                formula="=TargetWholesaleMargin",
                help_text="Thinner than retail on purpose: no "
                          "marketplace fees and the shop does the "
                          "selling.")
    r += 1

    # -------------------------------------------------------- the costs
    r = s.section(r, "2. What the piece costs you", LAST_COL)
    r = s.note(r, "Wholesale skips the gift packaging and the retail "
                  "postage, so the cost base is lower than it is for "
                  "a retail sale. Override anything that is "
                  "different for a trade order.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Cost", "From the catalog", "Your override",
                            "Used", "Note"],
                     [32, 16, 15, 15, 42])
    rows = [
        ("Materials", _pick("mat_cost"), "WcMat", "money",
         "The same whoever buys it."),
        ("Labour", _pick("labour_adj"), "WcLab", "money",
         "Often lower per piece on a batch - override it if so."),
        ("Consumables", _pick("consumables"), "WcCons", "money",
         "Glue, thread, gas, sandpaper."),
        ("Trade packaging", "0.35", "WcPack", "money",
         "A poly bag and a swing tag, not a gift box."),
        ("Overhead", _pick("overhead"), "WcOh", "money",
         "The same share of the studio as any other piece."),
    ]
    first = r
    for label, formula, name, style, note in rows:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render("=" + formula),
                         fmts.calc(style))
        s.write(r, 3, None, fmts.input(style))
        ws.data_validation(r, 3, r, 3, number_check(0, 1000000))
        book.define_name(name, f"='{NAME}'!$D${r + 1}")
        ws.write_formula(r, 4, f'=IF(D{r + 1}<>"",D{r + 1},C{r + 1})',
                         fmts.calc(style))
        book.define_name(name + "Used", f"='{NAME}'!$E${r + 1}")
        s.merge(r, 5, LAST_COL, note, fmts.chrome("note"))
        r += 1
    s.write(r, 1, "Cost a unit", fmts.chrome("total_label"))
    ws.write_formula(r, 4, f'=ROUND(SUM(E{first + 1}:E{r}),2)',
                     fmts.chrome("total_value",
                                 num_format=s.num_format("money")))
    book.define_name("WcCost", f"='{NAME}'!$E${r + 1}")
    r += 2

    # ------------------------------------------------- the two directions
    r = s.section(r, "3. Working it both ways", LAST_COL)
    r = s.subsection(r, "A. From your retail price down", LAST_COL,
                     first_col=1)
    rows_a = [
        ("Your retail price", "=" + _pick("price", '""'), "money",
         "WcRetail", "From the catalog."),
        ("Wholesale price", '=IF(WcRetail="","",ROUND(WcRetail*'
         '(1-WcDiscount),2))', "money", "WcPrice",
         "Retail less the discount the shop is asking for."),
        ("Profit a unit", '=IF(WcPrice="","",ROUND(WcPrice-WcCost-'
         'WcCarriage/MAX(1,WcQty),2))', "money", "WcProfit",
         "After the cost above and a share of the carriage."),
        ("Wholesale margin", '=IF(OR(WcPrice="",WcPrice=0),"",'
         'WcProfit/WcPrice)', "percent", "WcMargin",
         "Compare it with your target above."),
        ("You earn an hour", '=IF(OR(WcPrice="",' + _pick("hours", "0") +
         '=0),"",ROUND((WcProfit+WcLabUsed)/' + _pick("hours", "1") +
         ',2))', "money", "WcWage",
         "Profit plus your own wage, over the hours it takes."),
    ]
    for label, formula, style, name, note in rows_a:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 3, s.reg.render(formula), fmts.calc(style))
        book.define_name(name, f"='{NAME}'!$D${r + 1}")
        s.merge(r, 4, LAST_COL, note, fmts.chrome("note"))
        r += 1
    ws.conditional_format(r - 2, 3, r - 2, 3,
                          {"type": "cell", "criteria": "<",
                           "value": "=WcTarget",
                           "format": fmts.cf("warn")})
    ws.conditional_format(r - 2, 3, r - 2, 3,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    r += 1

    r = s.subsection(r, "B. From the margin you need, back up", LAST_COL,
                     first_col=1)
    rows_b = [
        ("Wholesale price you need",
         '=IF(WcTarget>=1,"",ROUND((WcCost+WcCarriage/MAX(1,WcQty))/'
         '(1-WcTarget),2))', "money", "WcNeedWholesale",
         "Cost divided by one minus your target margin."),
        ("Retail price that supports it",
         '=IF(OR(WcNeedWholesale="",WcDiscount>=1),"",ROUND(' +
         round_price("(WcNeedWholesale/(1-WcDiscount))") + ',2))',
         "money", "WcNeedRetail",
         "The retail price the shop will put on the shelf. If it is "
         "above what the market pays, the costs have to come down."),
        ("Versus your current retail",
         '=IF(OR(WcRetail="",WcNeedRetail=""),"",ROUND(WcNeedRetail-'
         'WcRetail,2))', "money", "WcGap",
         "Positive means your retail price is too low to wholesale "
         "from."),
        ("Keystone check (cost x 4)",
         '=ROUND(WcCost*4,2)', "money", "WcKeystone",
         "The old trade rule: double for wholesale, double again for "
         "retail. Crude, but a useful second opinion."),
    ]
    for label, formula, style, name, note in rows_b:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 3, s.reg.render(formula), fmts.calc(style))
        book.define_name(name, f"='{NAME}'!$D${r + 1}")
        s.merge(r, 4, LAST_COL, note, fmts.chrome("note"))
        r += 1
    r += 1

    # ------------------------------------------------------- the order --
    r = s.section(r, "4. The order itself", LAST_COL)
    r = s.table_head(r, 1, ["Units", "Order value", "Cost to make",
                            "Carriage", "Profit", "Margin",
                            "Hours of work", "Profit an hour"],
                     [10, 14, 14, 11, 12, 10, 14, 14])
    ws.write_formula(r, 1, "=WcQty", fmts.calc("int"))
    ws.write_formula(r, 2, '=IF(WcPrice="","",ROUND(WcQty*WcPrice,2))',
                     fmts.calc("money"))
    ws.write_formula(r, 3, "=ROUND(WcQty*WcCost,2)", fmts.calc("money"))
    ws.write_formula(r, 4, "=WcCarriage", fmts.calc("money"))
    ws.write_formula(r, 5, f'=IF(C{r + 1}="","",ROUND(C{r + 1}-'
                           f'D{r + 1}-E{r + 1},2))', fmts.calc("money"))
    ws.write_formula(r, 6, f'=IF(OR(C{r + 1}="",C{r + 1}=0),"",'
                           f'F{r + 1}/C{r + 1})', fmts.calc("percent"))
    ws.write_formula(r, 7, "=ROUND(WcQty*" + s.reg.render(_pick("hours")) +
                     ",1)", fmts.calc("number1"))
    ws.write_formula(r, 8, f'=IF(OR(H{r + 1}=0,F{r + 1}=""),"",'
                           f'ROUND(F{r + 1}/H{r + 1},2))',
                     fmts.calc("money"))
    order_row = r
    r += 1
    s.merge(r, 1, LAST_COL,
            "The hours column is the one to look at before you say "
            "yes. A profitable order that fills six weeks can still "
            "be the wrong order.", fmts.chrome("note"))
    ws.conditional_format(order_row, 5, order_row, 5,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(order_row, 8, order_row, 8,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    r += 2

    # --------------------------------------------------- discount ladder
    r = s.section(r, "5. What each discount does", LAST_COL)
    r = s.note(r, "The same product at every discount a shop might "
                  "ask for. The last column is the honest answer to "
                  "'can you do a bit better on the price?'.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Discount", "Wholesale price", "Profit a unit",
                            "Margin", "You earn an hour",
                            "On a full order", "Verdict"],
                     [11, 16, 14, 10, 16, 15, 26])
    ladder = r
    for disc in DISCOUNTS:
        price = f'IF(WcRetail="","",ROUND(WcRetail*{1 - disc},2))'
        s.write(r, 1, disc, fmts.calc("percent0"))
        ws.write_formula(r, 2, f"={price}", fmts.calc("money"))
        ws.write_formula(r, 3, f'=IF(C{r + 1}="","",ROUND(C{r + 1}-'
                               f'WcCost-WcCarriage/MAX(1,WcQty),2))',
                         fmts.calc("money"))
        ws.write_formula(r, 4, f'=IF(OR(C{r + 1}="",C{r + 1}=0),"",'
                               f'D{r + 1}/C{r + 1})',
                         fmts.calc("percent"))
        ws.write_formula(
            r, 5, '=IF(OR(C{0}="",{1}=0),"",ROUND((D{0}+WcLabUsed)/'
                  '{1},2))'.format(r + 1, s.reg.render(_pick("hours"))),
            fmts.calc("money"))
        ws.write_formula(r, 6, f'=IF(D{r + 1}="","",ROUND(D{r + 1}*'
                               f'WcQty,2))', fmts.calc("money"))
        ws.write_formula(
            r, 7,
            f'=IF(C{r + 1}="","",IF(D{r + 1}<0,"No - below cost",'
            f'IF(E{r + 1}<0.1,"No - nothing left in it",'
            f'IF(E{r + 1}<WcTarget,"Thin - only for a big order",'
            f'"Workable"))))', fmts.calc("text"))
        r += 1
    ws.conditional_format(ladder, 3, r - 1, 3,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(ladder, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "No -", "format": fmts.cf("bad")})
    ws.conditional_format(ladder, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "Workable", "format": fmts.cf("good")})
    r += 1

    # ------------------------------------------------------- the verdict
    r = s.section(r, "6. Should you take it?", LAST_COL)
    checks = [
        ("The order makes money",
         '=IF(WcPrice="","Pick a product first",'
         'IF(WcProfit>0,"Yes - there is profit in every piece",'
         '"No - every piece loses money at this discount"))'),
        ("It pays you properly",
         '=IF(OR(WcWage="",WcWage=0),"Add the labour minutes first",'
         'IF(WcWage>=MinimumWage,"Yes - above minimum wage",'
         '"No - you would be paying to make them"))'),
        ("Your retail price supports it",
         '=IF(WcGap="","-",IF(WcGap<=0,'
         '"Yes - your retail price has room in it",'
         '"No - retail needs to be "&TEXT(WcNeedRetail,"0.00")&'
         '" for this to work"))'),
        ("It meets your minimum",
         '=IF(WcPrice="","-",IF(AND(WcQty*WcPrice>=MinOrderValue,'
         'WcQty>=MinOrderQty),"Yes - above your minimum order",'
         '"Check - below the minimum on your Settings tab"))'),
        ("Can you actually make them?",
         '=IF(' + _pick("hours", "0") + '=0,'
         '"Add the labour minutes to find out",'
         '"It is "&TEXT(WcQty*' + _pick("hours", "0") +
         ',"0.0")&" hours of making - about "&'
         'TEXT(ROUNDUP(WcQty*' + _pick("hours", "0") +
         '/8,0),"0")&" full days")'),
    ]
    first_check = r
    for label, formula in checks:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc("text"))
        r += 1
    for word, kind in (("No -", "bad"), ("Yes -", "good"),
                       ("Check", "warn")):
        ws.conditional_format(first_check, 2, r - 1, 2,
                              {"type": "text", "criteria": "containing",
                               "value": word, "format": fmts.cf(kind)})
    r += 1

    s.note(r,
           "Taking the order? Add the shop to the Wholesale Customers "
           "tab with its discount and terms, then enter the lines on "
           "Wholesale Orders - the margins there follow whatever you "
           "agree here.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/break_even.py`

*333 lines · 2 functions*

> `Profit & Break-Even` - how many you have to sell before you earn anything.
>
> Break-even is the least glamorous number in the workbook and the most
> useful. It answers one question: given what the studio costs every
> month and what each piece contributes, how many do you have to sell
> before you are working for yourself rather than for the landlord?
>
>     Contribution a unit = price - fees - the costs that vary
>     Break-even units    = fixed costs / contribution a unit
>
> If the contribution is zero or negative, no number of sales works and
> the page says so rather than printing something enormous.

| Defines | What it does |
|---------|--------------|
| `_pick(field, default)` |  |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Profit & Break-Even'), `LAST_COL` (11), `PRODUCT_ROWS` (15), `PRICE_STEPS` (6 items).

````python
"""`Profit & Break-Even` - how many you have to sell before you earn anything.

Break-even is the least glamorous number in the workbook and the most
useful. It answers one question: given what the studio costs every
month and what each piece contributes, how many do you have to sell
before you are working for yourself rather than for the landlord?

    Contribution a unit = price - fees - the costs that vary
    Break-even units    = fixed costs / contribution a unit

If the contribution is zero or negative, no number of sales works and
the page says so rather than printing something enormous.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

NAME = "Profit & Break-Even"
LAST_COL = 11
PRODUCT_ROWS = 15
PRICE_STEPS = [-0.20, -0.10, 0.0, 0.10, 0.20, 0.30]


def _pick(field: str, default: str = "0") -> str:
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH(BeProduct,'
            f'[[Products.name]],0)),{default})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Profit, Margin & Break-Even",
             "What each piece contributes, how many pay for the month, "
             "and the difference between margin and markup.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 32, fmts.chrome("label"))
    s.set_col(2, 2, 12, fmts.chrome("label"))
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 15)
    s.set_col(5, 5, 15)
    s.set_col(6, LAST_COL, 14, fmts.chrome("note"))

    r = 4

    # -------------------------------------------------------- the inputs
    r = s.section(r, "1. The numbers it works from", LAST_COL)
    r = s.field(r, "Product", "Willow cable beanie", name="BeProduct",
                validate=list_from("ProductNames"), value_end=5,
                help_text="Or leave it and use the override boxes "
                          "below to test something hypothetical.")
    r = s.field(r, "Fixed costs a month", None, style="money",
                name="BeFixed", validate=number_check(0, 10000000),
                formula="=ROUND(OverheadMonthly,2)",
                help_text="From the Overhead tab. Type over it to "
                          "include a wage you want to draw.")
    r = s.field(r, "Wage you want to draw a month", 0, style="money",
                name="BeWage", validate=number_check(0, 10000000),
                help_text="Treat your own pay as a fixed cost and "
                          "break-even becomes a living, not a "
                          "survival.")
    r = s.field(r, "Price", None, style="money", name="BePrice",
                validate=number_check(0, 1000000),
                formula="=" + _pick("price", '""'),
                help_text="The catalog price. Override it to test a "
                          "different one.")
    r = s.field(r, "Variable cost a unit", None, style="money",
                name="BeVariable", validate=number_check(0, 1000000),
                formula="=" + _pick("cost_base", "0"),
                help_text="Materials, labour, packaging and postage - "
                          "everything that only happens when you make "
                          "one. Overhead is the fixed part and is "
                          "counted above.")
    r = s.field(r, "Percentage fees", None, style="percent",
                mode="calc", name="BeFeePct",
                formula="=" + _pick("fee_pct", "0"),
                help_text="From the channel this product sells on.")
    r = s.field(r, "Fixed fees a sale", None, style="money", mode="calc",
                name="BeFeeFixed",
                formula="=" + _pick("fee_fixed", "0"))
    r += 1

    # --------------------------------------------------- the calculation
    r = s.section(r, "2. The calculation, step by step", LAST_COL)
    r = s.table_head(r, 1, ["Step", "Value", "What it means"],
                     [32, 15, 58])
    steps = [
        ("Fixed costs a month", "=ROUND(BeFixed+BeWage,2)", "money",
         "BeFixedTotal",
         "Overhead plus whatever you want to pay yourself."),
        ("Price", '=IF(BePrice="",0,BePrice)', "money", "BeP",
         "What the buyer pays for the item."),
        ("Fees on that price",
         "=ROUND(BeP*BeFeePct+BeFeeFixed,2)", "money", "BeFees",
         "Percentage fees plus the fixed ones."),
        ("Variable cost", '=IF(BeVariable="",0,BeVariable)', "money",
         "BeVar", "Everything that only happens when you make one."),
        ("Contribution a unit",
         "=ROUND(BeP-BeFees-BeVar,2)", "money", "BeContribution",
         "What each sale puts towards the fixed costs. This is the "
         "number that matters."),
        ("Contribution margin",
         '=IF(BeP=0,"",BeContribution/BeP)', "percent", "BeCm",
         "Contribution as a share of the price."),
    ]
    for label, formula, style, name, note in steps:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc(style))
        book.define_name(name, f"='{NAME}'!$C${r + 1}")
        s.merge(r, 3, LAST_COL, note, fmts.chrome("note"))
        r += 1
    ws.conditional_format(r - 2, 2, r - 2, 2,
                          {"type": "cell", "criteria": "<=", "value": 0,
                           "format": fmts.cf("bad")})
    r += 1

    # ----------------------------------------------------- the KPI tiles
    tiles = [
        ("Break-even units a month",
         '=IF(BeContribution<=0,"",ROUNDUP(BeFixedTotal/'
         'BeContribution,0))', "int", "fixed costs / contribution"),
        ("Break-even revenue",
         '=IF(BeContribution<=0,"",ROUNDUP(BeFixedTotal/'
         'BeContribution,0)*BeP)', "money0", "at this price"),
        ("Units a week",
         '=IF(BeContribution<=0,"",ROUNDUP(BeFixedTotal/'
         'BeContribution/4.33,0))', "int", "to keep pace"),
        ("Hours a month to make them",
         '=IF(BeContribution<=0,"",ROUND(ROUNDUP(BeFixedTotal/'
         'BeContribution,0)*' + _pick("hours", "0") + ',1))',
         "number1", "is that even possible?"),
    ]
    for i, (label, formula, style, note) in enumerate(tiles):
        s.kpi(r, 1 + i * 2, label, formula, style=style, span=2,
              note=note)
    r += 4

    warn_row = r
    s.merge(r, 1, LAST_COL, None, fmts.calc("text"))
    ws.write_formula(
        r, 1,
        '=IF(BeContribution>0,"Every sale puts "&TEXT(BeContribution,'
        '"0.00")&" towards the "&TEXT(BeFixedTotal,"0")&" of fixed '
        'costs - "&TEXT(ROUNDUP(BeFixedTotal/BeContribution,0),"0")&'
        '" sales a month and the studio pays for itself.",'
        '"Warning: this product contributes nothing. At this price '
        'the fees and the variable costs are more than the price, so '
        'no number of sales will ever break even. Raise the price or '
        'cut the cost.")', fmts.calc("text"))
    s.row_height(r, 30)
    ws.conditional_format(warn_row, 1, warn_row, 1,
                          {"type": "text", "criteria": "containing",
                           "value": "Warning", "format": fmts.cf("bad")})
    r += 2

    # ----------------------------------------------- margin versus markup
    r = s.section(r, "3. Margin and markup are not the same thing",
                  LAST_COL)
    r = s.note(r, "Margin is profit over the price; markup is profit "
                  "over the cost. A 50% markup is a 33% margin, and "
                  "mixing them up is how a shop ends up pricing at "
                  "cost plus a bit.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["This product", "Value", "Worked out as"],
                     [32, 15, 58])
    pair = [
        ("Profit a unit", "=ROUND(BeContribution,2)", "money",
         "Price - fees - variable cost"),
        ("Margin", '=IF(BeP=0,"",BeContribution/BeP)', "percent",
         "Profit / price x 100"),
        ("Markup", '=IF(BeVar=0,"",BeContribution/BeVar)', "percent",
         "Profit / cost x 100"),
        ("Multiplier on cost", '=IF(BeVar=0,"",BeP/BeVar)', "number",
         "How many times your cost the price is"),
    ]
    for label, formula, style, how in pair:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, formula, fmts.calc(style))
        s.merge(r, 3, LAST_COL, how, fmts.chrome("note"))
        r += 1
    r += 1

    r = s.table_head(r, 1, ["If you want this margin", "The markup is",
                            "Multiply your cost by", "Example on $10 of cost"],
                     [24, 16, 22, 24])
    conv_first = r
    for margin in (0.20, 0.30, 0.35, 0.40, 0.45, 0.50, 0.55, 0.60, 0.65,
                   0.70):
        s.write(r, 1, margin, fmts.calc("percent0"))
        ws.write_formula(r, 2, f'={margin}/(1-{margin})',
                         fmts.calc("percent"))
        ws.write_formula(r, 3, f'=1/(1-{margin})', fmts.calc("number"))
        ws.write_formula(r, 4, f'=ROUND(10/(1-{margin}),2)',
                         fmts.calc("money"))
        r += 1
    ws.conditional_format(conv_first, 3, r - 1, 3,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    r += 1

    # -------------------------------------------------- price sensitivity
    r = s.section(r, "4. Break-even at other prices", LAST_COL)
    r = s.table_head(r, 1, ["Price", "Change", "Contribution",
                            "Margin", "Break-even units",
                            "Break-even revenue", "Verdict"],
                     [13, 10, 14, 10, 17, 18, 24])
    sens = r
    for step in PRICE_STEPS:
        price = f"ROUND(BeP*{1 + step},2)"
        contrib = f"ROUND({price}-({price}*BeFeePct+BeFeeFixed)-BeVar,2)"
        ws.write_formula(r, 1, f"={price}", fmts.calc("money"))
        s.write(r, 2, step, fmts.calc("percent0"))
        ws.write_formula(r, 3, f"={contrib}", fmts.calc("money"))
        ws.write_formula(r, 4, f'=IF(B{r + 1}=0,"",D{r + 1}/B{r + 1})',
                         fmts.calc("percent"))
        ws.write_formula(r, 5, f'=IF(D{r + 1}<=0,"Never",'
                               f'ROUNDUP(BeFixedTotal/D{r + 1},0))',
                         fmts.calc("int"))
        ws.write_formula(r, 6, f'=IF(D{r + 1}<=0,"",ROUNDUP('
                               f'BeFixedTotal/D{r + 1},0)*B{r + 1})',
                         fmts.calc("money0"))
        ws.write_formula(
            r, 7,
            f'=IF(D{r + 1}<=0,"Loses money on every sale",'
            f'IF(E{r + 1}<MarginWarning,"Too thin to live on",'
            f'IF(F{r + 1}<=UnitsPerMonth,"Reachable at your volume",'
            f'"More than you currently make")))', fmts.calc("text"))
        r += 1
    ws.conditional_format(sens, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "Reachable", "format": fmts.cf("good")})
    ws.conditional_format(sens, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "Loses", "format": fmts.cf("bad")})
    ws.conditional_format(sens, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "More than", "format": fmts.cf("warn")})
    r += 1

    # ----------------------------------------------- every product's view
    r = s.section(r, "5. Every product against the monthly overhead",
                  LAST_COL)
    r = s.note(r, "How many of each you would have to sell in a month "
                  "to pay the fixed costs on that product alone. It is "
                  "a blunt test and a very clear one.",
               LAST_COL, first_col=1, height=26)
    r = s.table_head(r, 1, ["Product", "Price", "Cost", "Fees",
                            "Contribution", "Margin",
                            "Units to cover the month", "Hours that takes"],
                     [30, 11, 11, 11, 13, 10, 20, 15])
    table_first = r
    for i in range(1, PRODUCT_ROWS + 1):
        name = f'IFERROR(INDEX([[Products.name]],{i})&"","")'
        price = f'IFERROR(INDEX([[Products.price]],{i}),"")'
        cost = f'IFERROR(INDEX([[Products.cost_base]],{i}),"")'
        fees = f'IFERROR(INDEX([[Products.fees]],{i}),"")'
        hours = f'IFERROR(INDEX([[Products.hours]],{i}),0)'
        ws.write_formula(r, 1, s.reg.render(f"={name}"),
                         fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(f'=IF({name}="","",{price})'),
                         fmts.calc("money"))
        ws.write_formula(r, 3, s.reg.render(f'=IF({name}="","",{cost})'),
                         fmts.calc("money"))
        ws.write_formula(r, 4, s.reg.render(f'=IF({name}="","",{fees})'),
                         fmts.calc("money"))
        ws.write_formula(
            r, 5,
            f'=IF(C{r + 1}="","",ROUND(C{r + 1}-D{r + 1}-E{r + 1},2))',
            fmts.calc("money"))
        ws.write_formula(
            r, 6,
            f'=IF(OR(C{r + 1}="",C{r + 1}=0),"",F{r + 1}/C{r + 1})',
            fmts.calc("percent"))
        ws.write_formula(
            r, 7,
            f'=IF(C{r + 1}="","",IF(F{r + 1}<=0,"Never",'
            f'ROUNDUP(BeFixedTotal/F{r + 1},0)))', fmts.calc("int"))
        ws.write_formula(
            r, 8,
            s.reg.render(f'=IF(OR(C{r + 1}="",F{r + 1}<=0),"",'
                         f'ROUND(H{r + 1}*{hours},0))'),
            fmts.calc("int"))
        r += 1
    ws.conditional_format(table_first, 6, r - 1, 6,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("warn")})
    ws.conditional_format(table_first, 7, r - 1, 7,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    r += 1

    # ------------------------------------------------- where you are now
    r = s.section(r, "6. Where you actually are", LAST_COL)
    r = s.table_head(r, 1, ["This month", "Value", "What it tells you"],
                     [30, 15, 58])
    now = [
        ("Revenue so far",
         '=ROUND(SUMIFS([[Sales.revenue]],[[Sales.date]],">="&MonthStart,'
         '[[Sales.date]],"<="&MonthEnd),2)', "money",
         "Everything logged on the Sales tab this month."),
        ("Contribution so far",
         '=ROUND(SUMIFS([[Sales.profit]],[[Sales.date]],">="&MonthStart,'
         '[[Sales.date]],"<="&MonthEnd),2)', "money",
         "Profit on the orders logged, before the studio overhead."),
        ("Fixed costs to cover", "=ROUND(BeFixed+BeWage,2)", "money",
         "What the month has to pay for."),
        ("Still to find",
         '=ROUND(MAX(0,BeFixed+BeWage-SUMIFS([[Sales.profit]],'
         '[[Sales.date]],">="&MonthStart,[[Sales.date]],"<="&MonthEnd)),2)',
         "money", "The gap between here and break-even."),
        ("Days left in the month",
         "=MAX(0,MonthEnd-TodayDate)", "int",
         "How long you have to find it."),
    ]
    for label, formula, style, what in now:
        s.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc(style))
        s.merge(r, 3, LAST_COL, what, fmts.chrome("note"))
        r += 1
    r += 1

    s.note(r,
           "Break-even is not a target, it is a floor. Pay yourself a "
           "wage in the box at the top and the number becomes the "
           "honest answer to 'how much do I have to sell to make this "
           "a job?'.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/scenarios.py`

*284 lines · 2 functions*

> `Scenarios` - the same product priced five ways, on every channel, with
> the materials going up.
>
> Three questions this page answers that no single price can:
>
> * What would my whole range look like if I repositioned it?
> * What do I have to charge on each channel to keep the same money?
> * What happens to my margins when materials go up 15%?
>
> Nothing here changes your catalog. It is a thinking tool.

| Defines | What it does |
|---------|--------------|
| `_pick(field, default)` |  |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Scenarios'), `LAST_COL` (11), `PRODUCT_ROWS` (15), `SCENARIOS` (5), `CHANNEL_ROWS` (10), `SHOCKS` (6 items).

````python
"""`Scenarios` - the same product priced five ways, on every channel, with
the materials going up.

Three questions this page answers that no single price can:

* What would my whole range look like if I repositioned it?
* What do I have to charge on each channel to keep the same money?
* What happens to my margins when materials go up 15%?

Nothing here changes your catalog. It is a thinking tool.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check
from .common import round_price

NAME = "Scenarios"
LAST_COL = 11
PRODUCT_ROWS = 15
SCENARIOS = 5
CHANNEL_ROWS = 10
SHOCKS = [0.0, 0.05, 0.10, 0.15, 0.25, 0.40]


def _pick(field: str, default: str = "0") -> str:
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH(ScProduct,'
            f'[[Products.name]],0)),{default})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Advanced Pricing Scenarios",
             "Reposition the whole range, compare every channel, and "
             "see what a price rise in your materials really does.",
             LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 30, fmts.chrome("label"))
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 14)

    r = 4

    # ------------------------------------------- 1. the whole range, 5 ways
    r = s.section(r, "1. Your range at five different positions", LAST_COL)
    r = s.note(r, "Each column prices every product at that scenario's "
                  "target margin, using its own costs and its own "
                  "channel fees. The margins behind the five columns "
                  "are on the Lists tab and you can change them.",
               LAST_COL, first_col=1, height=30)

    r = s.table_head(r, 1, ["Product", "Cost", "Your price", "", "", "",
                            "", "", "Where your price sits"],
                     [30, 11, 12, 12, 12, 12, 12, 12, 22])
    header_row = r - 1
    for i in range(1, SCENARIOS + 1):
        ws.write_formula(header_row, 3 + i,
                         f'=IFERROR(INDEX(ScenarioNames,{i}),"")',
                         fmts.chrome("header"))

    first = r
    for n in range(1, PRODUCT_ROWS + 1):
        # &"" so an empty catalog row reads as blank: INDEX on an
        # empty cell hands back a zero, and a zero is not ""
        name = f'IFERROR(INDEX([[Products.name]],{n})&"","")'
        cost = f'IFERROR(INDEX([[Products.total_cost]],{n}),"")'
        price = f'IFERROR(INDEX([[Products.price]],{n}),"")'
        # *1 turns the blank an empty catalog row returns into an
        # error, which IFERROR then reads as zero - subtracting "" from
        # 1 is a #VALUE! and it would spread across the whole block
        fee_pct = f'IFERROR(INDEX([[Products.fee_pct]],{n})*1,0)'
        fee_fix = f'IFERROR(INDEX([[Products.fee_fixed]],{n})*1,0)'
        ws.write_formula(r, 1, s.reg.render(f"={name}"), fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(f'=IF({name}="","",{cost})'),
                         fmts.calc("money"))
        ws.write_formula(r, 3, s.reg.render(f'=IF({name}="","",{price})'),
                         fmts.calc("money"))
        for i in range(1, SCENARIOS + 1):
            margin = f'IFERROR(INDEX(ScenarioMargins,{i}),0)'
            denom = f"(1-{margin}-{fee_pct})"
            body = f"(C{r + 1}+{fee_fix})/{denom}"
            ws.write_formula(
                r, 3 + i,
                s.reg.render(f'=IF(C{r + 1}="","",IF({denom}<=0,"",'
                             f'ROUND(' + round_price(f"({body})") + ',2)))'),
                fmts.calc("money"))
        ws.write_formula(
            r, 9,
            f'=IF(D{r + 1}="","",IF(D{r + 1}<E{r + 1},"Below the lowest",'
            f'IF(D{r + 1}<F{r + 1},"Budget",IF(D{r + 1}<G{r + 1},'
            f'"Conservative",IF(D{r + 1}<H{r + 1},"Standard",'
            f'IF(D{r + 1}<I{r + 1},"Premium","Luxury"))))))',
            fmts.calc("text"))
        r += 1
    ws.conditional_format(first, 9, r - 1, 9,
                          {"type": "text", "criteria": "containing",
                           "value": "Below the lowest",
                           "format": fmts.cf("bad")})
    ws.conditional_format(first, 9, r - 1, 9,
                          {"type": "text", "criteria": "containing",
                           "value": "Budget", "format": fmts.cf("warn")})
    r += 1
    s.merge(r, 1, LAST_COL,
            "If most of your range sits in the 'Budget' column, that "
            "is a positioning decision you probably never made on "
            "purpose.", fmts.chrome("note"))
    r += 2

    # ------------------------------------------- 2. one product, each channel
    r = s.section(r, "2. One product on every channel", LAST_COL)
    r = s.field(r, "Product", "Willow cable beanie", name="ScProduct",
                validate=list_from("ProductNames"), value_end=5,
                help_text="The channel comparison below follows this "
                          "choice.")
    r = s.field(r, "Target margin", None, style="percent", name="ScTarget",
                validate=number_check(0, 0.95),
                formula="=TargetRetailMargin")
    r += 1
    r = s.note(r, "The same piece, the same costs, different fees. The "
                  "'price to keep the same margin' column is what you "
                  "should actually list it at on each site - charging "
                  "one price everywhere quietly gives the dearest "
                  "channel your profit.",
               LAST_COL, first_col=1, height=30)
    r = s.table_head(r, 1, ["Channel", "Percentage fees", "Fixed fees",
                            "Price for the same margin", "Fees you pay",
                            "You keep", "At your current price",
                            "Margin there"],
                     [22, 15, 12, 20, 13, 12, 18, 13])
    ch_first = r
    for n in range(1, CHANNEL_ROWS + 1):
        name = f'IFERROR(INDEX([[Channels.name]],{n})&"","")'
        pct = f'IFERROR(INDEX([[Channels.pct_total]],{n})*1,0)'
        fix = f'IFERROR(INDEX([[Channels.fixed_total]],{n})*1,0)'
        ws.write_formula(r, 1, s.reg.render(f"={name}"), fmts.calc("text"))
        ws.write_formula(r, 2, s.reg.render(f'=IF({name}="","",{pct})'),
                         fmts.calc("percent"))
        ws.write_formula(r, 3, s.reg.render(f'=IF({name}="","",{fix})'),
                         fmts.calc("money"))
        cost = s.reg.render(_pick("total_cost"))
        ws.write_formula(
            r, 4,
            f'=IF(B{r + 1}="","",IF(1-ScTarget-C{r + 1}<=0,"",'
            f'ROUND(({cost}+D{r + 1})/(1-ScTarget-C{r + 1}),2)))',
            fmts.calc("money"))
        ws.write_formula(
            r, 5,
            f'=IF(E{r + 1}="","",ROUND(E{r + 1}*C{r + 1}+D{r + 1},2))',
            fmts.calc("money"))
        ws.write_formula(
            r, 6,
            f'=IF(E{r + 1}="","",ROUND(E{r + 1}-F{r + 1}-{cost},2))',
            fmts.calc("money"))
        price = s.reg.render(_pick("price", '""'))
        ws.write_formula(
            r, 7,
            f'=IF(OR(B{r + 1}="",{price}=""),"",ROUND({price}-'
            f'({price}*C{r + 1}+D{r + 1})-{cost},2))',
            fmts.calc("money"))
        ws.write_formula(
            r, 8,
            f'=IF(OR(H{r + 1}="",{price}=""),"",H{r + 1}/{price})',
            fmts.calc("percent"))
        r += 1
    ws.conditional_format(ch_first, 7, r - 1, 7,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(ch_first, 8, r - 1, 8,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("warn")})
    ws.conditional_format(ch_first, 4, r - 1, 4,
                          {"type": "data_bar", "bar_color": "#8C4A32"})
    r += 2

    # --------------------------------------------- 3. the cost shock table
    r = s.section(r, "3. What if your materials go up?", LAST_COL)
    r = s.note(r, "Yarn, silver, wax and timber all moved a long way "
                  "in the last few years. This is what each rise does "
                  "to the product above if you hold your price - and "
                  "what you would have to charge instead.",
               LAST_COL, first_col=1, height=30)
    r = s.table_head(r, 1, ["Materials go up by", "New material cost",
                            "New total cost", "Margin if you hold your price",
                            "Price to keep your margin", "Rise needed",
                            "Verdict"],
                     [18, 17, 16, 24, 20, 13, 24])
    shock_first = r
    mat = s.reg.render(_pick("mat_cost"))
    cost = s.reg.render(_pick("total_cost"))
    price = s.reg.render(_pick("price", '""'))
    fee_pct = s.reg.render(_pick("fee_pct"))
    fee_fix = s.reg.render(_pick("fee_fixed"))
    for shock in SHOCKS:
        s.write(r, 1, shock, fmts.calc("percent0"))
        ws.write_formula(r, 2, f"=ROUND({mat}*{1 + shock},2)",
                         fmts.calc("money"))
        ws.write_formula(r, 3, f"=ROUND({cost}+{mat}*{shock},2)",
                         fmts.calc("money"))
        ws.write_formula(
            r, 4,
            f'=IF(OR({price}="",{price}=0),"",({price}-D{r + 1}-'
            f'({price}*{fee_pct}+{fee_fix}))/{price})',
            fmts.calc("percent"))
        ws.write_formula(
            r, 5,
            f'=IF(1-ScTarget-{fee_pct}<=0,"",ROUND((D{r + 1}+{fee_fix})/'
            f'(1-ScTarget-{fee_pct}),2))', fmts.calc("money"))
        ws.write_formula(
            r, 6,
            f'=IF(OR({price}="",{price}=0,F{r + 1}=""),"",F{r + 1}/'
            f'{price}-1)', fmts.calc("percent"))
        ws.write_formula(
            r, 7,
            f'=IF(E{r + 1}="","",IF(E{r + 1}<0,"You would be losing money",'
            f'IF(E{r + 1}<MarginWarning,"Margin gone - must re-price",'
            f'IF(G{r + 1}>0.1,"Needs a real price rise",'
            f'"Absorbable"))))', fmts.calc("text"))
        r += 1
    ws.conditional_format(shock_first, 4, r - 1, 4,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("bad")})
    ws.conditional_format(shock_first, 7, r - 1, 7,
                          {"type": "text", "criteria": "containing",
                           "value": "Absorbable", "format": fmts.cf("good")})
    r += 2

    # ------------------------------------------------ 4. volume and time
    r = s.section(r, "4. Quantity, time and what your hour is worth",
                  LAST_COL)
    r = s.note(r, "Making more of the same thing is usually faster per "
                  "piece. This is what that saving is worth - and what "
                  "it lets you offer a shop without losing anything.",
               LAST_COL, first_col=1, height=28)
    r = s.table_head(r, 1, ["If a batch of", "Minutes each",
                            "Labour each", "Cost each", "At your price",
                            "Margin", "You earn an hour"],
                     [14, 13, 13, 13, 14, 10, 16])
    minutes = s.reg.render(_pick("minutes"))
    labour = s.reg.render(_pick("labour_adj"))
    vol_first = r
    for qty, factor in ((1, 1.0), (3, 0.88), (6, 0.80), (12, 0.72),
                        (24, 0.66), (48, 0.60)):
        s.write(r, 1, qty, fmts.calc("int"))
        ws.write_formula(r, 2, f"=ROUND({minutes}*{factor},1)",
                         fmts.calc("number1"))
        ws.write_formula(r, 3, f"=ROUND({labour}*{factor},2)",
                         fmts.calc("money"))
        ws.write_formula(r, 4, f"=ROUND({cost}-{labour}+D{r + 1},2)",
                         fmts.calc("money"))
        ws.write_formula(
            r, 5, f'=IF({price}="","",ROUND({price}-E{r + 1}-'
                  f'({price}*{fee_pct}+{fee_fix}),2))',
            fmts.calc("money"))
        ws.write_formula(
            r, 6, f'=IF(OR({price}="",{price}=0),"",F{r + 1}/{price})',
            fmts.calc("percent"))
        ws.write_formula(
            r, 7, f'=IF(OR(C{r + 1}=0,F{r + 1}=""),"",ROUND((F{r + 1}+'
                  f'D{r + 1})/(C{r + 1}/60),2))', fmts.calc("money"))
        r += 1
    ws.conditional_format(vol_first, 7, r - 1, 7,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    r += 1
    s.merge(r, 1, LAST_COL,
            "The batch factors here are a reasonable rule of thumb. "
            "Log a few runs on the Batch Production tab and you will "
            "have your own.", fmts.chrome("note"))
    r += 2

    s.note(r,
           "Nothing on this page writes to your catalog. When a "
           "scenario convinces you, change the price on the Product "
           "Catalog and log it on Price History.",
           LAST_COL, first_col=1, height=28)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/reports.py`

*338 lines · 2 functions*

> `Reports` - the numbers you would want if someone asked how the year went.
>
> Six reports, all built from the same registers: profit by product,
> where the material money goes, what your hours are worth, the month by
> month profit and loss, how your prices sit against the market, and the
> year against your targets.
>
> Nothing here needs refreshing and nothing here is typed twice.

| Defines | What it does |
|---------|--------------|
| `_rank(rank_key, field, n)` |  |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Reports'), `LAST_COL` (11), `TOP` (12), `MATERIAL_ROWS` (12), `MONTHS` (12).

````python
"""`Reports` - the numbers you would want if someone asked how the year went.

Six reports, all built from the same registers: profit by product,
where the material money goes, what your hours are worth, the month by
month profit and loss, how your prices sit against the market, and the
year against your targets.

Nothing here needs refreshing and nothing here is typed twice.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Reports"
LAST_COL = 11
TOP = 12
MATERIAL_ROWS = 12
MONTHS = 12


def _rank(rank_key: str, field: str, n: int) -> str:
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH({n},'
            f'[[Products.{rank_key}]],0)),"")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Reports & Analytics",
             "Profit by product, where the material money goes, what "
             "your hours earn, and the year against your targets.",
             LAST_COL)
    s.set_col(0, 0, 30)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4

    # -------------------------------------------- 1. profit by product --
    row = s.section(row, "1. Profitability by product", LAST_COL)
    row = s.note(row, "Ranked by the profit each sale makes. The hourly "
                      "column is the one that should decide what you "
                      "make next week.",
                 LAST_COL, height=26)
    row = s.table_head(row, 0, ["Product", "Category", "Cost", "Price",
                                "Profit", "Margin", "Markup", "Hours",
                                "An hour", "Sold", "Revenue", "Profit made"],
                       [30, 16, 11, 11, 11, 10, 10, 9, 11, 9, 13, 13])
    first = row
    for n in range(1, TOP + 1):
        fields = [("name", "text"), ("category", "text"),
                  ("total_cost", "money"), ("price", "money"),
                  ("profit", "money"), ("margin", "percent"),
                  ("markup", "percent"), ("hours", "number1"),
                  ("profit_hour", "money"), ("sold", "int"),
                  ("revenue", "money0"), ("actual_profit", "money0")]
        for col, (field, style) in enumerate(fields):
            ws.write_formula(row, col,
                             s.reg.render("=" + _rank("profit_rank", field,
                                                      n)),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(first, 4, row - 1, 4,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    ws.conditional_format(first, 5, row - 1, 5,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("warn")})
    ws.conditional_format(first, 8, row - 1, 8,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    row += 2

    # ------------------------------------------------ 2. material spend --
    row = s.section(row, "2. Where the material money goes", LAST_COL)
    row = s.table_head(row, 0, ["Material", "Type", "Used in", "Unit cost",
                                "Cost used", "Stock value", "Waste",
                                "Share of material cost"],
                       [30, 18, 11, 13, 13, 13, 11, 20])
    mat_first = row
    for n in range(1, MATERIAL_ROWS + 1):
        name = f'IFERROR(INDEX([[Materials.name]],MATCH({n},' \
               f'[[Materials.cost_used_rank]],0)),"")'
        mat_id = f'IFERROR(INDEX([[Materials.id]],MATCH({n},' \
                 f'[[Materials.cost_used_rank]],0)),"")'
        ws.write_formula(row, 0, s.reg.render(f"={name}"),
                         fmts.calc("text"))
        for col, field, style in ((1, "type", "text"),
                                  (2, "lines", "int"),
                                  (3, "unit_cost", "money4"),
                                  (4, "cost_used", "money"),
                                  (5, "stock_value", "money")):
            body = (f'IFERROR(INDEX([[Materials.{field}]],MATCH({n},'
                    f'[[Materials.cost_used_rank]],0)),"")')
            ws.write_formula(row, col, s.reg.render(f"={body}"),
                             fmts.calc(style))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=IF({mat_id}="","",ROUND(SUMIFS('
                         f'[[Waste Log.cost]],[[Waste Log.material]],'
                         f'{mat_id}),2))'), fmts.calc("money"))
        ws.write_formula(
            row, 7,
            s.reg.render(f'=IFERROR(E{row + 1}/SUM([[Materials.cost_used]]),'
                         f'"")'), fmts.calc("percent"))
        row += 1
    ws.conditional_format(mat_first, 7, row - 1, 7,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    row += 1
    s.merge(row, 0, LAST_COL,
            "If two or three materials dominate this list, that is "
            "where a supplier conversation or a bulk order is worth "
            "the most.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------------- 3. your hours --
    row = s.section(row, "3. What your hours are worth", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "Value", "What it means"],
                       [30, 15, 60])
    hours = [
        ("Your hourly rate", "=HourlyRate", "money",
         "What you decided your time is worth, on the Settings tab."),
        ("Average earned an hour, by product",
         '=IFERROR(AVERAGE([[Products.wage_hour]]),"")', "money",
         "Across the catalog, at your current prices."),
        ("Best paid product an hour",
         '=IFERROR(MAX([[Products.wage_hour]]),"")', "money",
         "Worth knowing what it is and making more of it."),
        ("Worst paid product an hour",
         '=IFERROR(MIN([[Products.wage_hour]]),"")', "money",
         "Worth knowing what it is and re-pricing it."),
        ("Products that pay under minimum wage",
         '=COUNTIFS([[Products.wage_hour]],"<"&MinimumWage,'
         '[[Products.name]],"<>")', "int",
         "Each one of these is a decision to work for less."),
        ("Hours sold in the last twelve months",
         '=ROUND(SUM([[Sales.hours]]),1)', "number1",
         "Making time that actually turned into money."),
        ("Profit an hour, actual",
         '=IFERROR(ROUND(SUM([[Sales.profit]])/SUM([[Sales.hours]]),2),"")',
         "money",
         "Total profit over total making hours - the real figure."),
        ("Productive hours you assumed",
         "=ProductiveHours", "number",
         "What the overhead per hour is based on. Be honest about "
         "it."),
    ]
    hour_first = row
    for label, formula, style, what in hours:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula), fmts.calc(style))
        s.merge(row, 2, LAST_COL, what, fmts.chrome("note"))
        row += 1
    ws.conditional_format(hour_first + 1, 1, hour_first + 3, 1,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    row += 2

    # --------------------------------------------- 4. profit and loss ----
    row = s.section(row, "4. Month by month", LAST_COL)
    s.write(row, 0, "Month", fmts.chrome("header_left"))
    for i in range(MONTHS):
        back = MONTHS - 1 - i
        ws.write_formula(
            row, i + 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{back},1),'
            f'"mmm yy")', fmts.chrome("header"))
        ws.write_formula(
            row + 1, i + 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{back},1),'
            f'"yyyy-mm")', fmts.chrome("header", font_size=8))
    s.write(row + 1, 0, "Key", fmts.chrome("label"))
    key_row = row + 1
    row += 2
    band = {}
    pl_lines = [
        ("Revenue", "SUMIFS([[Sales.revenue]],[[Sales.month]],{key})",
         "money0"),
        ("Materials & labour",
         "SUMIFS([[Sales.cogs]],[[Sales.month]],{key})", "money0"),
        ("Selling fees", "SUMIFS([[Sales.fees]],[[Sales.month]],{key})",
         "money0"),
        ("Postage absorbed",
         "SUMIFS([[Sales.ship_absorbed]],[[Sales.month]],{key})",
         "money0"),
    ]
    for label, body, style in pl_lines:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        for i in range(MONTHS):
            cell = f"{chr(66 + i)}${key_row + 1}"
            ws.write_formula(row, i + 1,
                             s.reg.render("=ROUND(" +
                                          body.format(key=cell) + ",2)"),
                             fmts.calc(style))
        band[label] = row
        row += 1
    s.write(row, 0, "Overhead", fmts.chrome("label_bold"))
    for i in range(MONTHS):
        ws.write_formula(row, i + 1, "=ROUND(OverheadMonthly,2)",
                         fmts.calc("money0"))
    band["Overhead"] = row
    row += 1
    s.write(row, 0, "Profit after overhead", fmts.chrome("total_label"))
    for i in range(MONTHS):
        col = chr(66 + i)
        parts = "-".join(f"{col}{band[k] + 1}" for k in
                         ("Materials & labour", "Selling fees",
                          "Postage absorbed", "Overhead"))
        ws.write_formula(
            row, i + 1,
            f'=ROUND({col}{band["Revenue"] + 1}-{parts},2)',
            fmts.chrome("total_value", num_format=s.num_format("money0")))
    profit_row = row
    row += 1
    s.write(row, 0, "Margin after overhead", fmts.chrome("label_bold"))
    for i in range(MONTHS):
        col = chr(66 + i)
        ws.write_formula(
            row, i + 1,
            f'=IFERROR({col}{profit_row + 1}/{col}'
            f'{band["Revenue"] + 1},"")', fmts.calc("percent0"))
    margin_row = row
    row += 1
    ws.add_sparkline(profit_row, MONTHS + 1, {
        "range": (f"{quote_sheet(NAME)}!B{profit_row + 1}:"
                  f"M{profit_row + 1}"),
        "type": "column", "negative_points": True,
        "series_color": "#5E7B5A", "negative_color": "#B3261E",
    })
    ws.conditional_format(profit_row, 1, profit_row, MONTHS,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(margin_row, 1, margin_row, MONTHS,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetRetailMargin",
                           "format": fmts.cf("good")})
    s.set_col(MONTHS + 1, MONTHS + 1, 12)
    row += 1
    s.merge(row, 0, LAST_COL,
            "Overhead is charged at the same amount every month "
            "because that is how it arrives. A month that does not "
            "cover it is a month the studio paid for itself out of "
            "another month.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------- 5. against the market
    row = s.section(row, "5. Your prices against the market", LAST_COL)
    row = s.table_head(row, 0, ["Product", "Your price", "Market average",
                                "Difference", "You are", "Research notes"],
                       [30, 12, 15, 12, 14, 34])
    comp_first = row
    for n in range(1, TOP + 1):
        name = _rank("margin_rank", "name", n)
        ws.write_formula(row, 0, s.reg.render("=" + name),
                         fmts.calc("text"))
        ws.write_formula(row, 1,
                         s.reg.render("=" + _rank("margin_rank", "price", n)),
                         fmts.calc("money"))
        ws.write_formula(
            row, 2,
            s.reg.render("=" + _rank("margin_rank", "comp_price", n)),
            fmts.calc("money"))
        ws.write_formula(
            row, 3,
            f'=IF(OR(B{row + 1}="",C{row + 1}="",C{row + 1}=0),"",'
            f'ROUND(B{row + 1}-C{row + 1},2))', fmts.calc("money"))
        ws.write_formula(
            row, 4,
            f'=IF(D{row + 1}="","No research yet",'
            f'IF(D{row + 1}>C{row + 1}*0.15,"Dearer",'
            f'IF(D{row + 1}<-C{row + 1}*0.15,"Cheaper","In line")))',
            fmts.calc("text"))
        ws.write_formula(
            row, 5,
            f'=IF(A{row + 1}="","",IF(E{row + 1}="Cheaper",'
            f'"Worth testing a rise",IF(E{row + 1}="Dearer",'
            f'"Make sure the listing shows why","")))',
            fmts.calc("text"))
        row += 1
    ws.conditional_format(comp_first, 4, row - 1, 4,
                          {"type": "text", "criteria": "containing",
                           "value": "Cheaper", "format": fmts.cf("warn")})
    ws.conditional_format(comp_first, 4, row - 1, 4,
                          {"type": "text", "criteria": "containing",
                           "value": "In line", "format": fmts.cf("good")})
    row += 2

    # ------------------------------------------------ 6. year to date ----
    row = s.section(row, "6. The year against your targets", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "This year so far", "Target",
                                "Difference", "Where you are"],
                       [30, 16, 14, 14, 30])
    targets = [
        ("Revenue",
         '=ROUND(SUMIFS([[Sales.revenue]],[[Sales.date]],">="&YearStart),2)',
         "=RevenueTargetYear", "money0"),
        ("Profit",
         '=ROUND(SUMIFS([[Sales.profit]],[[Sales.date]],">="&YearStart),2)',
         "=ProfitTargetMonth*12", "money0"),
        ("Items sold",
         '=SUMIFS([[Sales.qty]],[[Sales.date]],">="&YearStart)',
         "=UnitsPerMonth*12", "int"),
        ("Earned an hour",
         '=IFERROR(ROUND(SUMIFS([[Sales.profit]],[[Sales.date]],'
         '">="&YearStart)/MAX(0.1,SUMIFS([[Sales.hours]],'
         '[[Sales.date]],">="&YearStart)),2),"")',
         "=TargetHourly", "money"),
    ]
    for label, actual, target, style in targets:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(actual), fmts.calc(style))
        ws.write_formula(row, 2, target, fmts.calc(style))
        ws.write_formula(row, 3,
                         f'=IF(B{row + 1}="","",ROUND(B{row + 1}-'
                         f'C{row + 1},2))', fmts.calc(style))
        s.merge(row, 4, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(
            row, 4,
            f'=IF(OR(B{row + 1}="",C{row + 1}=0),"",'
            f'TEXT(B{row + 1}/C{row + 1},"0%")&" of the year\'s '
            f'target")', fmts.calc("text"))
        row += 1
    row += 1

    s.note(row,
           "Everything on this page comes from the Sales, Products, "
           "Materials and Waste tabs. If a report looks wrong, the "
           "fix is on one of those tabs rather than here.",
           LAST_COL, height=28)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/printables.py`

*526 lines · 5 functions*

> The three printable forms.
>
> * `Print Costing` - a one page costing card for a single product:
>   every material, every labour stage, the packaging, the fees and the
>   price, in a form you can put in a ring binder or hand to an
>   accountant.
> * `Print Quote` - a custom order quote for a buyer, with the lines, the
>   deposit and an acceptance box.
> * `Print Order` - a wholesale line sheet and order form: your range at
>   retail and trade prices, with a column the shop fills in.
>
> Pick an ID at the top of each page and the form fills itself in.

| Defines | What it does |
|---------|--------------|
| `_letterhead(s, fmts, row, last_col, heading)` | Shop name, contact details and the form's title |
| `_picker(s, fmts, book, row, label, default, list_name, define, help_text, last_col)` | The 'which one are we printing?' box |
| `build_costing(book, fmts, cfg, reg)` |  |
| `build_quote(book, fmts, cfg, reg)` |  |
| `build_order(book, fmts, cfg, reg)` |  |

**Module constants:** `COSTING_NAME` ('Print Costing'), `QUOTE_NAME` ('Print Quote'), `ORDER_NAME` ('Print Order'), `BOM_LINES` (12), `LABOUR_LINES` (7), `QUOTE_LINES` (8), `ORDER_LINES` (16).

````python
"""The three printable forms.

* `Print Costing` - a one page costing card for a single product:
  every material, every labour stage, the packaging, the fees and the
  price, in a form you can put in a ring binder or hand to an
  accountant.
* `Print Quote` - a custom order quote for a buyer, with the lines, the
  deposit and an acceptance box.
* `Print Order` - a wholesale line sheet and order form: your range at
  retail and trade prices, with a column the shop fills in.

Pick an ID at the top of each page and the form fills itself in.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from, number_check

COSTING_NAME = "Print Costing"
QUOTE_NAME = "Print Quote"
ORDER_NAME = "Print Order"

BOM_LINES = 12
LABOUR_LINES = 7
QUOTE_LINES = 8
ORDER_LINES = 16


def _letterhead(s: SheetWriter, fmts, row: int, last_col: int,
                heading: str) -> int:
    """Shop name, contact details and the form's title."""
    ws = s.ws
    s.merge(row, 0, 3, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0, "=ShopName", fmts.chrome("print_title"))
    s.merge(row, 4, last_col, heading, fmts.chrome("print_title"))
    row += 1
    s.merge(row, 0, 3, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0, "=ShopAddress", fmts.chrome("print_sub"))
    s.merge(row, 4, last_col, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 4,
                     '=ShopEmail&IF(ShopPhone="",""," - "&ShopPhone)&'
                     'IF(ShopWebsite="",""," - "&ShopWebsite)',
                     fmts.chrome("print_sub"))
    return row + 2


def _picker(s: SheetWriter, fmts, book, row: int, label: str,
            default: str, list_name: str, define: str,
            help_text: str, last_col: int = 7) -> int:
    """The 'which one are we printing?' box."""
    ws = s.ws
    s.write(row, 0, label, fmts.chrome("label_bold"))
    s.write(row, 1, default, fmts.input("id"))
    ws.data_validation(row, 1, row, 1, list_from(list_name))
    book.define_name(define, f"={quote_sheet(s.name)}!$B${row + 1}")
    s.merge(row, 2, last_col, help_text, fmts.chrome("note"))
    return row + 2


# ---------------------------------------------------------------------------
# Print Costing
# ---------------------------------------------------------------------------

def build_costing(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = COSTING_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Costing Sheet", "", LAST_COL)
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 28)
    s.set_col(2, 2, 14)
    s.set_col(3, 3, 11)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 13)
    s.set_col(7, 7, 15)

    row = 4
    row = _picker(s, fmts, book, row, "Product", "P-0001", "ProductIDs",
                  "PrintProduct",
                  "Pick a product. Every line below fills itself in "
                  "from the Materials, Material Usage and Labor tabs.",
                  LAST_COL)
    row = _letterhead(s, fmts, row, LAST_COL, "COSTING SHEET")

    def p(key: str, style: str = "text") -> str:
        return s.reg.render(f'=IFERROR(INDEX([[Products.{key}]],'
                            f'MATCH(PrintProduct,[[Products.id]],0)),"")')

    # ------------------------------------------------------- the header --
    pairs = [
        ("Product", "name", "text", "SKU", "sku", "text"),
        ("Category", "category", "text", "Status", "status", "text"),
        ("Channel", "channel", "text", "Difficulty", "difficulty", "text"),
        ("Packaging", "pack_profile", "text", "Postage", "ship_option",
         "text"),
    ]
    for label_a, key_a, style_a, label_b, key_b, style_b in pairs:
        s.write(row, 0, label_a, fmts.chrome("label_bold"))
        s.merge(row, 1, 2, None, fmts.cell(style_a, "plain"))
        ws.write_formula(row, 1, p(key_a), fmts.cell(style_a, "plain"))
        s.write(row, 4, label_b, fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.cell(style_b, "plain"))
        ws.write_formula(row, 5, p(key_b), fmts.cell(style_b, "plain"))
        row += 1
    row += 1

    # ---------------------------------------------------- the materials --
    row = s.section(row, "Materials", LAST_COL)
    row = s.table_head(row, 0, ["Material", "Detail", "Unit", "Quantity",
                                "Waste", "Unit cost", "Line cost"],
                       [18, 28, 14, 11, 10, 12, 13])
    bom_first = row
    for n in range(1, BOM_LINES + 1):
        key = f'PrintProduct&"|{n}"'
        fields = [("material_name", "text"), ("notes", "text"),
                  ("unit", "text"), ("qty", "number1"),
                  ("waste_rate", "percent"), ("unit_cost", "money4"),
                  ("line_cost", "money")]
        for col, (field, style) in enumerate(fields):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Material Usage.{field}]],'
                             f'MATCH({key},[[Material Usage.key]],0)),"")'),
                fmts.cell(style, "plain"))
        row += 1
    s.merge(row, 0, 5, "Materials total", fmts.chrome("total_label"))
    ws.write_formula(row, 6,
                     f"=ROUND(SUM(G{bom_first + 1}:G{row}),2)",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    # ------------------------------------------------------- the labour --
    row = s.section(row, "Labour", LAST_COL)
    row = s.table_head(row, 0, ["Stage", "Who", "Basis", "Minutes",
                                "Per item", "Rate", "Cost"],
                       [18, 28, 14, 11, 10, 12, 13])
    lab_first = row
    for n in range(1, LABOUR_LINES + 1):
        key = f'PrintProduct&"|{n}"'
        fields = [("stage", "text"), ("who", "text"), ("basis", "text"),
                  ("minutes", "number1"), ("mins_item", "number1"),
                  ("rate", "money"), ("cost", "money")]
        for col, (field, style) in enumerate(fields):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Labor.{field}]],'
                             f'MATCH({key},[[Labor.key]],0)),"")'),
                fmts.cell(style, "plain"))
        row += 1
    s.merge(row, 0, 5, "Labour total", fmts.chrome("total_label"))
    ws.write_formula(row, 6, f"=ROUND(SUM(G{lab_first + 1}:G{row}),2)",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    # ------------------------------------------------- everything else --
    row = s.section(row, "The rest of the cost", LAST_COL)
    row = s.table_head(row, 0, ["Cost", "What it covers", "", "", "", "",
                                "Amount"],
                       [18, 28, 14, 11, 10, 12, 13])
    others = [
        ("Materials", "Everything in the piece, waste included",
         "mat_cost"),
        ("Labour", "Your hours at your rate, difficulty applied",
         "labour_adj"),
        ("Consumables", "Thread, glue, gas, abrasives", "consumables"),
        ("Packaging", "Box, tissue, card, label", "pack_cost"),
        ("Postage absorbed", "What the label costs less what the buyer "
                             "pays", "ship_cost"),
        ("Overhead", "This product's share of the studio", "overhead"),
    ]
    other_first = row
    for label, what, key in others:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 5, what, fmts.chrome("note"))
        ws.write_formula(row, 6, p(key), fmts.cell("money", "plain"))
        row += 1
    s.merge(row, 0, 5, "Total cost", fmts.chrome("total_label"))
    ws.write_formula(row, 6,
                     f"=ROUND(SUM(G{other_first + 1}:G{row}),2)",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    # -------------------------------------------------------- the price --
    row = s.section(row, "The price", LAST_COL)
    price_rows = [
        ("Price", "price", "money", "What it sells for."),
        ("Channel fees", "fees", "money",
         "What the marketplace takes on that price."),
        ("Profit", "profit", "money", "What is left for you."),
        ("Margin", "margin", "percent", "Profit over the price."),
        ("Markup", "markup", "percent", "Profit over the cost."),
        ("Minutes of work", "minutes", "number1",
         "Including the difficulty factor."),
        ("You earn an hour", "wage_hour", "money",
         "Your wage plus the profit, over the hours."),
        ("Wholesale price", "wholesale", "money",
         "At the standard trade discount."),
        ("Suggested price", "suggested", "money",
         "What the target margin says it should be."),
    ]
    for label, key, style, what in price_rows:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 5, what, fmts.chrome("note"))
        ws.write_formula(row, 6, p(key), fmts.cell(style, "plain"))
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(
        row, 0,
        '="Costed by "&ShopName&" on "&TEXT(TodayDate,"d mmmm yyyy")&'
        '". Labour at "&TEXT(LaborCostHour,"0.00")&" an hour. '
        'Marketplace fee rates last reviewed "&'
        'TEXT(FeesReviewed,"d mmm yyyy")&"."',
        fmts.chrome("note"))
    row += 1

    s.finish(freeze_row=None, print_area=f"A1:H{row + 1}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Quote
# ---------------------------------------------------------------------------

def build_quote(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = QUOTE_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Quote", "", LAST_COL)
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 16)
    s.set_col(3, 3, 9)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 13)
    s.set_col(7, 7, 14)

    row = 4
    s.merge(row, 0, LAST_COL,
            "For custom and commission work. Type the customer's "
            "details and pick the products - the prices come from "
            "your catalog and the totals work themselves out.",
            fmts.chrome("note_box"))
    row += 2
    row = _letterhead(s, fmts, row, LAST_COL, "QUOTE")

    # -------------------------------------------------- the two blocks --
    s.write(row, 0, "Quote for", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.input("text"))
    s.write(row, 4, "Quote number", fmts.chrome("label_bold"))
    s.write(row, 5, None, fmts.input("text"))
    book.define_name("QuoteCustomer", f"={quote_sheet(NAME)}!$B${row + 1}")
    row += 1
    s.write(row, 0, "Email", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.input("text"))
    s.write(row, 4, "Date", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, "=TodayDate", fmts.cell("date", "plain"))
    row += 1
    s.write(row, 0, "What they want", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.input("wrap"))
    s.write(row, 4, "Valid for", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, "30 days", fmts.input("text"))
    s.row_height(row, 32)
    row += 2

    # ------------------------------------------------------- the lines --
    row = s.table_head(row, 0, ["Product", "Description", "Variation",
                                "Qty", "Price each", "Line total"],
                       [18, 30, 16, 9, 12, 13])
    first_line = row
    for _ in range(QUOTE_LINES):
        s.write(row, 0, None, fmts.input("text"))
        ws.data_validation(row, 0, row, 0, list_from("ProductNames"))
        s.merge(row, 1, 2, None, fmts.input("text"))
        s.write(row, 3, None, fmts.input("text"))
        s.write(row, 4, None, fmts.input("int"))
        ws.data_validation(row, 4, row, 4,
                           number_check(0, 10000, decimal=False))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=IF(A{row + 1}="","",IFERROR(INDEX('
                         f'[[Products.price]],MATCH(A{row + 1},'
                         f'[[Products.name]],0)),""))'),
            fmts.cell("money", "plain"))
        ws.write_formula(
            row, 6,
            f'=IF(OR(E{row + 1}="",F{row + 1}=""),"",ROUND(E{row + 1}*'
            f'F{row + 1},2))', fmts.cell("money", "plain"))
        row += 1
    last_line = row - 1
    row += 1

    # ------------------------------------------------------ the totals --
    totals = [
        ("Subtotal",
         f"=ROUND(SUM(G{first_line + 1}:G{last_line + 1}),2)", "money",
         "The lines above."),
        ("Rush or custom surcharge", None, "money",
         "Type it in if the turnaround is short."),
        ("Postage", None, "money", "What you will charge for postage."),
        ("Discount", None, "money", "Anything you are knocking off."),
    ]
    total_cells = []
    for label, formula, style, note in totals:
        s.merge(row, 0, 2, note, fmts.chrome("note"))
        s.merge(row, 3, 5, label, fmts.chrome("label_right"))
        if formula:
            ws.write_formula(row, 6, formula, fmts.cell(style, "plain"))
        else:
            s.write(row, 6, None, fmts.input(style))
        total_cells.append(row)
        row += 1
    sub, rush, post, disc = (f"G{n + 1}" for n in total_cells)
    s.merge(row, 3, 5, "Total", fmts.chrome("total_label"))
    ws.write_formula(
        row, 6,
        f'=ROUND({sub}+IF({rush}="",0,{rush})+IF({post}="",0,{post})-'
        f'IF({disc}="",0,{disc}),2)',
        fmts.chrome("total_value", num_format=cfg.money_format))
    total_row = row
    row += 1
    s.merge(row, 3, 5, "Deposit to start", fmts.chrome("label_right"))
    ws.write_formula(row, 6, f"=ROUND(G{total_row + 1}*0.5,2)",
                     fmts.cell("money", "plain"))
    s.merge(row, 0, 2, "Half up front is normal for custom work.",
            fmts.chrome("note"))
    row += 2

    # --------------------------------------------------------- the terms
    row = s.section(row, "What you are agreeing to", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        '="Custom work begins once the deposit is received. The '
        'balance is due before despatch. Handmade pieces vary '
        'slightly - colour, grain and texture will not be identical '
        'to the photographs. Custom and personalised orders cannot '
        'be returned unless faulty. Prices hold for 30 days from '
        'the date above. "&ShopName&" - "&ShopEmail',
        fmts.cell("wrap", "plain"))
    s.row_height(row, 56)
    row += 2

    row = s.section(row, "Acceptance", LAST_COL)
    for label in ("Signed", "Print name"):
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 3, None, fmts.chrome("form_line"))
        s.write(row, 4, "Date", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 26)
        row += 1
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(row, 0,
                     '="Quote prepared by "&ShopName&" on "&'
                     'TEXT(TodayDate,"d mmmm yyyy")',
                     fmts.chrome("note"))
    row += 1

    s.finish(freeze_row=None, print_area=f"A1:H{row + 1}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Order
# ---------------------------------------------------------------------------

def build_order(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = ORDER_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Wholesale Line Sheet & Order Form", "", LAST_COL)
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 16)
    s.set_col(3, 3, 11)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 10)
    s.set_col(6, 6, 13)
    s.set_col(7, 7, 14)

    row = 4
    row = _picker(s, fmts, book, row, "Customer", "WC-0001",
                  "CustomerIDs", "PrintCustomer",
                  "Pick a stockist and the sheet prices itself at "
                  "their discount. Leave the quantity column blank - "
                  "that is the bit they fill in.", LAST_COL)
    row = _letterhead(s, fmts, row, LAST_COL, "LINE SHEET")

    def c(key: str) -> str:
        return s.reg.render(
            f'=IFERROR(INDEX([[Wholesale Customers.{key}]],'
            f'MATCH(PrintCustomer,[[Wholesale Customers.id]],0)),"")')

    s.write(row, 0, "Shop", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 1, c("name"), fmts.cell("text", "plain"))
    s.write(row, 4, "Discount", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, c("discount_used"),
                     fmts.cell("percent", "plain"))
    row += 1
    s.write(row, 0, "Contact", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 1, c("contact"), fmts.cell("text", "plain"))
    s.write(row, 4, "Terms", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 5, c("terms"), fmts.cell("text", "plain"))
    row += 1
    s.write(row, 0, "Date", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=TodayDate", fmts.cell("date", "plain"))
    s.write(row, 4, "Order ref", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, None, fmts.input("text"))
    row += 2

    # -------------------------------------------------------- the range --
    row = s.table_head(row, 0, ["SKU", "Product", "Category", "Retail",
                                "Trade price", "Qty", "Line total"],
                       [18, 30, 16, 11, 12, 10, 13])
    first = row
    discount = 'IFERROR(INDEX([[Wholesale Customers.discount_used]],' \
               'MATCH(PrintCustomer,[[Wholesale Customers.id]],0)),' \
               'WholesaleDiscount)'
    for n in range(1, ORDER_LINES + 1):
        for col, field, style in ((0, "sku", "text"),
                                  (1, "name", "text"),
                                  (2, "category", "text")):
            # &"" keeps an empty catalog row blank on the printed sheet
            # instead of printing a zero
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Products.{field}]],{n})'
                             f'&"","")'), fmts.cell(style, "plain"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF(B{row + 1}="","",IFERROR('
                         f'INDEX([[Products.price]],{n}),""))'),
            fmts.cell("money", "plain"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IF(D{row + 1}="","",ROUND(D{row + 1}*(1-'
                         f'{discount}),2))'),
            fmts.cell("money", "plain"))
        s.write(row, 5, None, fmts.input("int"))
        ws.data_validation(row, 5, row, 5,
                           number_check(0, 100000, decimal=False))
        ws.write_formula(
            row, 6,
            f'=IF(OR(E{row + 1}="",F{row + 1}=""),"",ROUND(E{row + 1}*'
            f'F{row + 1},2))', fmts.cell("money", "plain"))
        row += 1
    last = row - 1
    row += 1

    s.merge(row, 0, 3, "Items ordered", fmts.chrome("label_right"))
    ws.write_formula(row, 4, f"=SUM(F{first + 1}:F{last + 1})",
                     fmts.cell("int", "plain"))
    s.merge(row, 5, 5, "Total", fmts.chrome("total_label"))
    ws.write_formula(row, 6,
                     f"=ROUND(SUM(G{first + 1}:G{last + 1}),2)",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    total_row = row
    row += 1
    s.merge(row, 0, 5, None, fmts.cell("text", "plain"))
    ws.write_formula(
        row, 0,
        f'=IF(G{total_row + 1}=0,"Add quantities in the Qty column.",'
        f'IF(G{total_row + 1}<MinOrderValue,"This order is under the '
        f'minimum of "&TEXT(MinOrderValue,"0.00")&".",'
        f'"Order accepted - thank you."))',
        fmts.cell("text", "plain"))
    ws.conditional_format(row, 0, row, 5,
                          {"type": "text", "criteria": "containing",
                           "value": "under the minimum",
                           "format": fmts.cf("warn")})
    row += 2

    row = s.section(row, "Trade terms", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        '="Minimum order "&TEXT(MinOrderValue,"0.00")&" or "&'
        'TEXT(MinOrderQty,"0")&" pieces. Lead time 2-4 weeks '
        'depending on the season. Carriage charged at cost unless '
        'agreed otherwise. Everything is made by hand, so small '
        'variations in colour and texture are part of the product '
        'rather than a fault. Payment terms as agreed above. "&'
        'ShopName&" - "&ShopEmail',
        fmts.cell("wrap", "plain"))
    s.row_height(row, 56)
    row += 2

    for label in ("Signed for the shop", "Print name"):
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 3, None, fmts.chrome("form_line"))
        s.write(row, 4, "Date", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 26)
        row += 1
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(row, 0,
                     '="Line sheet produced by "&ShopName&" on "&'
                     'TEXT(TodayDate,"d mmmm yyyy")&". Prices hold '
                     'for 60 days."', fmts.chrome("note"))
    row += 1

    s.finish(freeze_row=None, print_area=f"A1:H{row + 1}",
             gridlines=False, fit_width=True)
    return s
````

### `handmade_pricing_spreadsheet/sheets/help_tab.py`

*282 lines · 1 function*

> `Help` - troubleshooting, a glossary and a map of every feature.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Help'), `LAST_COL` (9).

````python
"""`Help` - troubleshooting, a glossary and a map of every feature."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Help"
LAST_COL = 9


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)

    s.banner("Help & Troubleshooting",
             "What the colours mean, what to do when a number looks "
             "wrong, and where every feature lives.", LAST_COL)

    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, LAST_COL, 14)

    row = 4

    # ------------------------------------------------ the colour code --
    row = s.section(row, "The colour code", LAST_COL)
    row = s.table_head(row, 1, ["Looks like", "Means"], [30, 68])
    colours = [
        ("White cell with a coloured edge",
         "Type here. These are the only cells you need to fill in."),
        ("Shaded cell", "A formula. It is locked so you cannot "
                        "overwrite it by accident."),
        ("Green", "Good - on target, profitable, above your "
                  "minimum wage."),
        ("Amber", "Worth a look - thin margin, stock low, research "
                  "going stale."),
        ("Red", "Act today - losing money, below minimum wage, out "
                "of stock."),
        ("Blue", "For information - a position, a status, a note."),
        ("A coloured bar inside a cell",
         "A data bar: the longer the bar, the bigger the number "
         "compared with the rest of the column."),
        ("A tiny chart", "A sparkline - twelve months of that "
                         "number in one cell."),
    ]
    for look, means in colours:
        s.write(row, 1, look, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, means, fmts.chrome("label"))
        s.row_height(row, 22)
        row += 1
    row += 1

    # ----------------------------------------------- troubleshooting ---
    row = s.section(row, "When something looks wrong", LAST_COL)
    row = s.table_head(row, 1, ["What you see", "Why", "The fix"],
                       [26, 36, 36])
    problems = [
        ("'Not found' in a name column",
         "The ID you typed is not on the tab it points at",
         "Pick from the drop-down instead of typing"),
        ("#DIV/0!",
         "Something is being divided by zero - usually a price or "
         "a pack size that is still empty",
         "Fill in the missing number"),
        ("#VALUE!",
         "Text where a number is expected, often from a pasted "
         "value with a currency symbol in it",
         "Paste as values and strip the symbols"),
        ("A price of 'Check the settings'",
         "Your target margin plus the fee rate is 100% or more, so "
         "no price can work",
         "Lower the target margin or check the fee rates"),
        ("'No price works'",
         "The same thing on the Retail Calculator",
         "Settings, then target margin; or Channels, then fees"),
        ("Suggested price looks far too high",
         "Usually the labour minutes - a figure in hours typed into "
         "a minutes box",
         "Check the Labor tab for that product"),
        ("Suggested price looks far too low",
         "Usually a missing cost: no packaging, no overhead or no "
         "labour at all",
         "Open the Retail Calculator and read the cost block"),
        ("Cost per gram looks wrong",
         "The pack size and the unit do not match - 100 g entered "
         "as 1 pack, or ml against a gram unit",
         "Materials tab: check pack quantity and unit"),
        ("A product shows no cost",
         "Nothing on the Material Usage or Labor tabs refers to it",
         "Add the bill of materials and the minutes"),
        ("Two sets of costs on one product",
         "It has rows on a craft tab and on Material Usage",
         "Use one or the other - the craft tab's check column "
         "warns you"),
        ("Margin is lower than you set",
         "The target is applied before fees; the fees then come "
         "off the price",
         "That is correct - the price already covers them"),
        ("Numbers do not update",
         "Calculation has been switched to manual",
         "Formulas, Calculation Options, Automatic - or press F9"),
        ("A cell will not let you type",
         "It is a formula and the sheet is protected",
         f"Review, Unprotect Sheet. The password is "
         f"'{cfg.password}'"),
        ("Drop-down list is missing an option",
         "The lists live on one tab",
         "Lists tab - add your option and it appears everywhere"),
    ]
    for what, why, fix in problems:
        s.write(row, 1, what, fmts.chrome("label_bold"))
        s.merge(row, 2, 5, why, fmts.chrome("label"))
        s.merge(row, 6, LAST_COL, fix, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ---------------------------------------------------- the glossary -
    row = s.section(row, "The words, in plain English", LAST_COL)
    row = s.table_head(row, 1, ["Term", "What it means here"],
                       [26, 72])
    glossary = [
        ("Cost", "Everything one piece costs you: materials, "
                 "labour, packaging, the postage you absorb and a "
                 "share of your overhead."),
        ("Margin", "Profit as a share of the price. $10 profit on a "
                   "$40 sale is a 25% margin."),
        ("Markup", "Profit as a share of the cost. The same $10 on "
                   "$30 of cost is a 33% markup. Always a bigger "
                   "number than the margin."),
        ("Contribution", "What one sale puts towards your fixed "
                         "costs: price less fees less the costs that "
                         "only happen when you make one."),
        ("Break-even", "How many you have to sell to cover the "
                       "fixed costs. Fixed costs divided by "
                       "contribution, rounded up."),
        ("Overhead", "Costs that are not about any one product - "
                     "studio rent, insurance, software, your website."),
        ("Base unit", "The unit a material is really used in: grams, "
                      "millilitres or centimetres. Costs are worked "
                      "out per base unit so quantities are simple."),
        ("Waste allowance", "The share of a material that never "
                            "makes it into the product - offcuts, "
                            "frogging, trim loss."),
        ("Percentage fees", "Fees charged as a share of the sale: "
                            "transaction, payment processing, "
                            "offsite ads."),
        ("Fixed fees", "Fees charged per order or per listing, "
                       "whatever the price: the listing fee, the "
                       "fixed part of card processing."),
        ("Keystone", "The old trade rule: double your cost for "
                     "wholesale, double again for retail. A sanity "
                     "check, not a method."),
        ("Line sheet", "The one page catalogue you send a shop, "
                       "with retail and trade prices."),
        ("Seconds", "Pieces that cannot be sold at full price. "
                    "Their cost is carried by the good ones."),
        ("Difficulty factor", "A multiplier on the time a fiddly "
                              "product takes, so you do not have to "
                              "re-time every variation."),
    ]
    for term, meaning in glossary:
        s.write(row, 1, term, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, meaning, fmts.chrome("label"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ------------------------------------------------- common questions
    row = s.section(row, "Questions people actually ask", LAST_COL)
    row = s.table_head(row, 1, ["Question", "Answer"], [30, 68])
    questions = [
        ("Should I really pay myself for every hour?",
         "Yes. If the price only works because your time is free, "
         "it is not a price, it is a subsidy. Start at your local "
         "minimum wage if your proper rate feels impossible, then "
         "raise it."),
        ("My price is now higher than everyone else's.",
         "Then either your costs are higher, your margin target is "
         "ambitious, or they are not paying themselves. Check the "
         "Competitors tab, then decide which of those it is."),
        ("What margin should I aim for?",
         "55% retail and 35% wholesale are reasonable starting "
         "points for handmade goods. Below 30% retail there is no "
         "room for a bad month."),
        ("Do I include my own labour in wholesale?",
         "Always. A wholesale price that only covers materials is "
         "a donation to a shop."),
        ("How do I handle free shipping?",
         "Put what the label costs in the Shipping tab and charge "
         "the buyer nothing. The workbook moves it into your cost "
         "where you can see it."),
        ("Should the pattern or design time be in every item?",
         "Spread it over the pieces you expect to sell, on the "
         "Development tab or the Digital tab. Loading it all onto "
         "the first sale makes the first sale impossible."),
        ("Why is my hourly wage lower than my hourly rate?",
         "Because the price is not covering the full cost. The "
         "wage column is what you really earn once the materials "
         "and fees are paid."),
        ("Can I use this for a craft fair?",
         "Yes - the Events tab costs the stall and tells you what "
         "the day has to take before it breaks even."),
        ("How often should I re-price?",
         "Twice a year, and immediately when a main material moves "
         "more than 10%. The catalog flags both."),
    ]
    for question, answer in questions:
        s.write(row, 1, question, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, answer, fmts.chrome("label"))
        s.row_height(row, 32)
        row += 1
    row += 1

    # ----------------------------------------------------- feature map -
    row = s.section(row, "Where every feature lives", LAST_COL)
    row = s.table_head(row, 1, ["If you want to...", "Go to"],
                       [46, 52])
    where = [
        ("Work out what to charge for something new",
         "Retail Calculator"),
        ("Check a product is paying you properly", "Products"),
        ("Cost a new material", "Materials"),
        ("Change your hourly rate", "Settings"),
        ("Change the Etsy or Shopify fee rates", "Channels"),
        ("Decide whether to accept a wholesale order",
         "Wholesale Calculator"),
        ("Find out how many you must sell this month",
         "Profit & Break-Even"),
        ("Price a crochet piece by yarn weight and hook",
         "Crochet"),
        ("Price a ring by grams of silver and scrap recovery",
         "Jewelry"),
        ("Price a candle by vessel size and fragrance load",
         "Candles"),
        ("Price a garment by fabric, size and layout waste",
         "Clothing"),
        ("Price a board by timber length and offcut allowance",
         "Woodworking"),
        ("Work out whether a pattern will ever pay for itself",
         "Digital"),
        ("See what a 20% discount really costs", "Promotions"),
        ("Compare what a run cost with what you estimated",
         "Batches"),
        ("See whether a craft fair was worth the day", "Events"),
        ("Test a price rise across the whole range", "Scenarios"),
        ("Print a costing sheet for your records", "Print Costing"),
        ("Send a shop a line sheet", "Print Order"),
        ("Quote for a commission", "Print Quote"),
        ("Add an option to a drop-down", "Lists"),
    ]
    for want, go in where:
        s.merge(row, 1, 5, want, fmts.chrome("label"))
        s.merge(row, 6, LAST_COL, go, fmts.chrome("label_bold"))
        s.row_height(row, 20)
        row += 1
    row += 1

    # ------------------------------------------------------- the small print
    row = s.section(row, "The small print", LAST_COL)
    row = s.note(row,
                 "This workbook is a calculator, not advice. Fee "
                 "rates, tax rules, safety testing and labelling "
                 "requirements differ by country and change without "
                 "notice - check the current official sources for "
                 "your marketplace and your jurisdiction before you "
                 "rely on a number. The sheet password protects "
                 "formulas from accidental edits; it is not "
                 "encryption and it is not security.",
                 LAST_COL, first_col=1, height=56)
    row += 1

    s.note(row,
           f"{cfg.product_name} - version {cfg.version}, by "
           f"{cfg.author}. No macros, no add-ins, nothing to "
           f"install. Every formula is readable and every list is "
           f"editable.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

---

## 7. Quality checks (optional)

Not needed to build the workbook - these are the checks that keep it honest. The first two need `openpyxl`, the third needs `formulas`.

### `tools/validate_handmade_workbook.py`

*702 lines · 11 functions*

> Static checks on a generated Handmade Pricing workbook.
>
> Run it with any Python that has openpyxl installed::
>
>     python tools/validate_handmade_workbook.py output/pricing.xlsx
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

**Module constants:** `FUNCTION_RE` (re.compile(...)), `NAME_RE` (re.compile(...)), `BARE_COL_RE` (re.compile(...)), `KNOWN_FUNCTIONS` ({'IF', 'IFERROR', 'SUM', 'SUMIF', 'SUMIFS', 'COUNT', 'COU...), `MODERN_ONLY` ({'XLOOKUP', 'IFS', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCA...), `PRECEDENCE_OK` ({'IF(P#="",#,P#)*U#+IF(G#="",#,G#)/#*LaborCostHour', 'IF(...), `SEEDED_INPUTS` (7 entries), `CHAIN_CHECKS` (82 items), `PAGE_FORMULAS` (11 items), `PAGE_INPUTS` (10 items).

````python
"""Static checks on a generated Handmade Pricing workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_handmade_workbook.py output/pricing.xlsx

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
from handmade_pricing_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402

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
    # variation labour = the base labour scaled, plus the extra minutes
    # over sixty at the labour rate - only the minutes are divided
    'IF(P#="",#,P#)*U#+IF(G#="",#,G#)/#*LaborCostHour',
    # jewelry consumables = solder, plating and hallmarking, plus bench
    # consumables charged by the hour
    'IF(Q#="",#,Q#)+IF(R#="",#,R#)+IF(S#="",#,S#)+AG#/#*SuppliesPerHour',
    # candle minutes = the minutes for one, plus the batch set-up spread
    # over the batch
    'IF(U#="",#,U#)+IF(T#="",#,T#)/MAX(#,IF(S#="",DefaultBatchSize,S#))',
    # woodworking consumables = glue and abrasives, plus shop
    # consumables by the hour
    'AG#+AD#/#*SuppliesPerHour',
    # a download: development spread over the sales it is expected to
    # make, plus the support minutes every sale really costs
    'V#/W#+IF(M#="",#,M#)/#*T#',
    'AL#+Z#-IF(M#="",#,M#)/#*T#',
    'AK#-AA#-(AK#*AF#+AG#)-IF(M#="",#,M#)/#*T#',
    'AK#-AA#-(AK#*AF#+AG#)-IF(M#="",#,M#)/#*T#<=#',
    # a wholesale line: carriage is per order, so only carriage is
    # divided by the quantity
    'C#-WcCost-WcCarriage/MAX(#,WcQty)',
    'WcPrice-WcCost-WcCarriage/MAX(#,WcQty)',
    'WcCost+WcCarriage/MAX(#,WcQty)',
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
    "Retail Calculator!D9": "target margin starts at the Settings figure",
    "Wholesale Calculator!D7": "trade discount starts at the Settings "
                               "figure",
    "Wholesale Calculator!D10": "target wholesale margin starts at the "
                                "Settings figure",
    "Profit & Break-Even!D7": "fixed costs start at the overhead total",
    "Profit & Break-Even!D9": "the price starts at the catalog price",
    "Profit & Break-Even!D10": "the cost starts at the catalog cost",
    "Scenarios!D28": "the scenario target margin starts at the Settings figure",
}

#: the automation chain, link by link: (sheet, column key, what the
#: formula in that column has to still be pointing at)
CHAIN_CHECKS = [
    # ---- a material's cost per unit, and what it is used in -------------
    ("Materials", "net_pack", ("column", "Materials", "pack_cost")),
    ("Materials", "unit_cost", ("column", "Materials", "net_pack")),
    ("Materials", "base_cost", ("column", "Materials", "unit_cost")),
    ("Materials", "cost_used", ("sheet", "Material Usage")),
    ("Materials", "cover", ("column", "Materials", "stock")),
    # ---- bill of materials ----------------------------------------------
    ("Material Usage", "unit_cost", ("sheet", "Materials")),
    ("Material Usage", "qty_eff", ("column", "Material Usage",
                                   "waste_rate")),
    ("Material Usage", "line_cost", ("column", "Material Usage",
                                     "unit_cost")),
    # ---- labour: minutes over sixty, times a rate -----------------------
    ("Labor", "rate", ("text", "LaborCostHour")),
    ("Labor", "hours", ("column", "Labor", "minutes")),
    ("Labor", "stage_cost", ("column", "Labor", "rate")),
    ("Labor", "cost", ("column", "Labor", "stage_cost")),
    # ---- packaging, postage and overhead --------------------------------
    ("Packaging", "unit_cost", ("column", "Packaging", "pack_cost")),
    ("Packaging", "line_cost", ("column", "Packaging", "unit_cost")),
    ("Shipping", "total_cost", ("column", "Shipping", "cost")),
    ("Shipping", "absorbed", ("column", "Shipping", "charge")),
    ("Overhead", "monthly", ("column", "Overhead", "amount")),
    ("Overhead", "per_hour", ("text", "ProductiveHours")),
    ("Overhead", "per_unit", ("text", "UnitsPerMonth")),
    # ---- marketplace fees, which must stay editable ---------------------
    ("Channels", "pct_total", ("column", "Channels", "transaction")),
    ("Channels", "fixed_total", ("column", "Channels", "processing_fixed")),
    ("Channels", "take_rate", ("column", "Channels", "fee_20")),
    # ---- the catalog: C = M + L + P + S + O -----------------------------
    ("Products", "mat_cost", ("sheet", "Material Usage")),
    ("Products", "labor_cost", ("sheet", "Labor")),
    ("Products", "pack_cost", ("text", "PackStandard")),
    ("Products", "ship_pay", ("sheet", "Shipping")),
    ("Products", "oh_rate", ("text", "OverheadMethod")),
    ("Products", "cost_base", ("column", "Products", "mat_cost")),
    ("Products", "fee_pct", ("sheet", "Channels")),
    ("Products", "fee_fixed", ("sheet", "Channels")),
    ("Products", "denom", ("column", "Products", "fee_pct")),
    ("Products", "suggested_raw", ("column", "Products", "denom")),
    ("Products", "suggested", ("text", "RoundingRule")),
    ("Products", "price", ("column", "Products", "price_input")),
    ("Products", "total_cost", ("column", "Products", "overhead")),
    ("Products", "fees", ("column", "Products", "fee_fixed")),
    ("Products", "profit", ("column", "Products", "total_cost")),
    ("Products", "margin", ("column", "Products", "profit")),
    ("Products", "markup", ("column", "Products", "total_cost")),
    ("Products", "profit_hour", ("column", "Products", "hours")),
    ("Products", "wholesale", ("text", "WholesaleDiscount")),
    ("Products", "comp_price", ("sheet", "Competitors")),
    ("Products", "sold", ("sheet", "Sales")),
    ("Products", "revenue", ("sheet", "Sales")),
    # ---- the six specialist calculators each reach the catalog ----------
    ("Crochet", "yarn_cost", ("sheet", "Materials")),
    ("Crochet", "labor_cost", ("column", "Crochet", "rate")),
    ("Crochet", "suggested_raw", ("column", "Crochet", "denom")),
    ("Jewelry", "metal_cost", ("column", "Jewelry", "metal_rate")),
    ("Jewelry", "suggested_raw", ("column", "Jewelry", "denom")),
    ("Candles", "wax_cost", ("column", "Candles", "wax_total")),
    ("Candles", "fragrance_cost", ("column", "Candles", "fragrance_g")),
    ("Clothing", "fabric_cost", ("column", "Clothing", "fabric_total")),
    ("Woodworking", "wood_cost", ("sheet", "Materials")),
    ("Digital", "dev_cost", ("column", "Digital", "dev_hours")),
    ("Digital", "break_even", ("column", "Digital", "dev_cost")),
    # ---- what actually happened -----------------------------------------
    ("Sales", "catalog_price", ("sheet", "Products")),
    ("Sales", "price_used", ("column", "Sales", "catalog_price")),
    ("Sales", "cogs", ("column", "Sales", "unit_cost_total")),
    ("Sales", "fees_est", ("column", "Sales", "fee_pct")),
    ("Sales", "fee_gap", ("column", "Sales", "fees_est")),
    ("Sales", "profit", ("column", "Sales", "cogs")),
    ("Sales", "margin", ("column", "Sales", "profit")),
    ("Sales", "profit_hour", ("column", "Sales", "hours")),
    # ---- bundles, variations, batches, wholesale ------------------------
    ("Bundles", "parts_cost", ("sheet", "Bundle Lines")),
    ("Bundles", "profit", ("column", "Bundles", "fees")),
    ("Bundle Lines", "unit_cost", ("sheet", "Products")),
    ("Variations", "base_mat", ("sheet", "Products")),
    ("Variations", "total_cost", ("column", "Variations", "mat_cost")),
    ("Batches", "mat_total", ("column", "Batches", "est_mat_unit")),
    ("Batches", "cost_good", ("column", "Batches", "batch_cost")),
    ("Batches", "variance", ("column", "Batches", "est_unit")),
    ("Wholesale Orders", "unit_price", ("column", "Wholesale Orders",
                                        "discount")),
    ("Wholesale Orders", "unit_cost", ("sheet", "Products")),
    ("Wholesale Orders", "margin", ("column", "Wholesale Orders",
                                    "profit")),
    ("Wholesale Customers", "value", ("sheet", "Wholesale Orders")),
    ("Promotions", "margin_after", ("column", "Promotions",
                                    "discount_pct")),
    ("Events", "net", ("column", "Events", "total_cost")),
    ("Development", "recovery", ("column", "Development", "dev_cost")),
    ("Price History", "change_pct", ("column", "Price History",
                                     "old_price")),
    ("Competitors", "gap", ("column", "Competitors", "your_total")),
    ("Waste Log", "unit_cost", ("sheet", "Materials")),
    ("Waste Log", "cost", ("column", "Waste Log", "unit_cost")),
]

#: cells reached by a defined name that have to hold a live formula -
#: the calculator pages are where the value of the file lives, and a
#: name pointing at an empty cell means a page quietly stopped working.
PAGE_FORMULAS = [
    ("RcCost", "retail calculator: cost a unit"),
    ("RcDenom", "retail calculator: 1 - margin - fee rate"),
    ("RcPrice", "retail calculator: the price it recommends"),
    ("RcOverhead", "retail calculator: overhead per unit"),
    ("WcPrice", "wholesale calculator: the trade price"),
    ("WcMargin", "wholesale calculator: margin at that price"),
    ("WcRetail", "wholesale calculator: the retail it implies"),
    ("BeContribution", "break-even: contribution a unit"),
    ("BeFixedTotal", "break-even: fixed costs a month"),
    ("BeCm", "break-even: units needed"),
    ("DashKey", "dashboard: the month the tiles report on"),
]

#: names that must point at an unlocked cell the buyer types into
PAGE_INPUTS = [
    ("RcProduct", "retail calculator: which product"),
    ("RcTarget", "retail calculator: target margin"),
    ("RcChannel", "retail calculator: which channel"),
    ("WcProduct", "wholesale calculator: which product"),
    ("WcDiscount", "wholesale calculator: trade discount"),
    ("BeProduct", "break-even: which product"),
    ("BePrice", "break-even: price being tested"),
    ("ScProduct", "scenarios: which product"),
    ("PrintProduct", "print costing: which product"),
    ("QuoteCustomer", "print quote: who it is for"),
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
    from handmade_pricing_spreadsheet.build import REGISTER_MODULES
    from handmade_pricing_spreadsheet.config import Config
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
              else "output/Handmade Product Pricing & Profit "
                   "Calculator.xlsx")
    raise SystemExit(check(target))
````

### `tools/check_handmade_sample.py`

*400 lines · 6 functions*

> Cross-check the example records in a generated handmade workbook.
>
> The demo data has to join up, otherwise the dashboard opens full of
> zeros and every lookup fills with "Not found".  This reads the literal
> values out of the finished file - never the formula columns - and makes
> sure every bill of materials line points at a real material, every sale
> at a real product, every wholesale order at a real stockist, that the
> shipping option typed against a product matches a real postage band,
> and that no product is costed twice over (once on its specialist tab
> and again through the bill of materials), which would silently double
> its cost.
>
>     python tools/check_handmade_sample.py output/pricing.xlsx

| Defines | What it does |
|---------|--------------|
| `data_rows()` | Last data row of every register, straight from its specification |
| `headers(ws, header_row)` |  |
| `column_values(ws, header, header_row)` | Every non-empty literal value under the given header |
| `row_dicts(ws, wanted)` | The literal values of the named columns, row by row |
| `check(path)` |  |
| `main(argv)` |  |

**Module constants:** `HEADER_ROW` (4), `CRAFT_SHEETS` (6 items), `LAST_ROW` (0 entries).

````python
"""Cross-check the example records in a generated handmade workbook.

The demo data has to join up, otherwise the dashboard opens full of
zeros and every lookup fills with "Not found".  This reads the literal
values out of the finished file - never the formula columns - and makes
sure every bill of materials line points at a real material, every sale
at a real product, every wholesale order at a real stockist, that the
shipping option typed against a product matches a real postage band,
and that no product is costed twice over (once on its specialist tab
and again through the bill of materials), which would silently double
its cost.

    python tools/check_handmade_sample.py output/pricing.xlsx
"""

from __future__ import annotations

import os
import sys
from collections import Counter

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

HEADER_ROW = 4

#: the six specialist calculators, which cost a product instead of the
#: bill of materials doing it
CRAFT_SHEETS = ["Crochet", "Jewelry", "Candles", "Clothing",
                "Woodworking", "Digital"]


def data_rows() -> dict:
    """Last data row of every register, straight from its specification.

    Without this the checks read whatever sits underneath the grid and
    report a summary label as a missing material.
    """
    from handmade_pricing_spreadsheet.build import REGISTER_MODULES
    from handmade_pricing_spreadsheet.config import Config
    cfg = Config()
    return {module.NAME: module.spec(cfg).last_data_row + 1
            for module in REGISTER_MODULES}


LAST_ROW: dict = {}


def headers(ws, header_row: int = HEADER_ROW) -> dict:
    out = {}
    for cell in ws[header_row]:
        if isinstance(cell.value, str) and cell.value.strip():
            out.setdefault(cell.value.strip(), cell.column)
    return out


def column_values(ws, header: str, header_row: int = HEADER_ROW) -> list:
    """Every non-empty literal value under the given header."""
    found = headers(ws, header_row)
    if header not in found:
        near = [name for name in found if name.startswith(header)]
        if len(near) != 1:
            raise KeyError(f"{ws.title}: no column called {header!r} "
                           f"(have {sorted(found)[:10]} ...)")
        header = near[0]
    col = found[header]
    out = []
    last = min(LAST_ROW.get(ws.title, ws.max_row), ws.max_row)
    for row in range(header_row + 1, last + 1):
        value = ws.cell(row=row, column=col).value
        if value is None:
            continue
        if isinstance(value, str):
            value = value.strip()
            if not value or value.startswith("="):
                continue
        out.append(value)
    return out


def row_dicts(ws, wanted: list) -> list:
    """The literal values of the named columns, row by row."""
    found = headers(ws)
    missing = [name for name in wanted if name not in found]
    if missing:
        raise KeyError(f"{ws.title}: no column called {missing[0]!r}")
    last = min(LAST_ROW.get(ws.title, ws.max_row), ws.max_row)
    rows = []
    for row in range(HEADER_ROW + 1, last + 1):
        record = {}
        for name in wanted:
            value = ws.cell(row=row, column=found[name]).value
            if isinstance(value, str):
                value = value.strip()
                if value.startswith("="):
                    value = None
            record[name] = value or None
        if any(v is not None for v in record.values()):
            rows.append(record)
    return rows


def check(path: str) -> int:
    LAST_ROW.update(data_rows())
    wb = openpyxl.load_workbook(path, data_only=False)
    problems: list = []
    notes: list = []

    def sheet(name):
        if name not in wb.sheetnames:
            problems.append(f"missing sheet {name}")
            return None
        return wb[name]

    products = sheet("Products")
    materials = sheet("Materials")
    if products is None or materials is None:
        for problem in problems:
            print(f"    - {problem}")
        return 1

    product_names = column_values(products, "Product")
    material_names = column_values(materials, "Material")
    channel_names = column_values(sheet("Channels"), "Channel")
    notes.append(f"{len(product_names)} products, {len(material_names)} "
                 f"materials, {len(channel_names)} channels")
    for label, values in (("product", product_names),
                          ("material", material_names),
                          ("channel", channel_names)):
        duplicates = [v for v, n in Counter(values).items() if n > 1]
        if duplicates:
            problems.append(f"duplicate {label} names: {duplicates[:5]}")

    def id_link(from_sheet: str, from_col: str, to_sheet: str,
                count_col: str) -> None:
        """IDs are generated by formulas, so check the number is a row."""
        src, dst = sheet(from_sheet), sheet(to_sheet)
        if src is None or dst is None:
            return
        try:
            values = [str(v) for v in column_values(src, from_col)]
            limit = len(column_values(dst, count_col))
        except KeyError as error:
            problems.append(str(error))
            return
        bad = []
        for value in values:
            digits = "".join(c for c in value if c.isdigit())
            if not digits or not 1 <= int(digits) <= limit:
                bad.append(value)
        if bad:
            problems.append(
                f"{from_sheet}.{from_col}: {len(bad)} reference(s) point "
                f"past the {limit} rows on {to_sheet} -> "
                f"{sorted(set(bad))[:5]}")
        else:
            notes.append(f"{from_sheet}.{from_col} -> {to_sheet} "
                         f"({len(values)} rows, max {limit}) ok")

    def link(from_sheet: str, from_col: str, to_sheet: str,
             to_col: str) -> None:
        src, dst = sheet(from_sheet), sheet(to_sheet)
        if src is None or dst is None:
            return
        try:
            values = column_values(src, from_col)
            allowed = set(column_values(dst, to_col))
        except KeyError as error:
            problems.append(str(error))
            return
        missing = sorted({str(v) for v in values}
                         - {str(a) for a in allowed})
        if missing:
            problems.append(
                f"{from_sheet}.{from_col}: {len(missing)} value(s) not in "
                f"{to_sheet}.{to_col} -> {missing[:5]}")
        else:
            notes.append(f"{from_sheet}.{from_col} -> {to_sheet}.{to_col} "
                         f"({len(values)} rows) ok")

    # ---- the costing spine ------------------------------------------
    id_link("Material Usage", "Product", "Products", "Product")
    id_link("Material Usage", "Material", "Materials", "Material")
    id_link("Labor", "Product", "Products", "Product")
    id_link("Bundle Lines", "Bundle", "Bundles", "Bundle")
    id_link("Bundle Lines", "Product", "Products", "Product")
    id_link("Variations", "Base product", "Products", "Product")
    id_link("Batches", "Product", "Products", "Product")
    id_link("Sales", "Product", "Products", "Product")
    id_link("Price History", "Product", "Products", "Product")
    id_link("Competitors", "Your product", "Products", "Product")
    id_link("Waste Log", "Material", "Materials", "Material")
    id_link("Waste Log", "On which product", "Products", "Product")
    id_link("Development", "Became product", "Products", "Product")
    id_link("Wholesale Orders", "Customer", "Wholesale Customers", "Shop")
    id_link("Wholesale Orders", "Product", "Products", "Product")
    for craft in CRAFT_SHEETS:
        id_link(craft, "Product ID", "Products", "Product")

    # ---- drop down values that have to match a row somewhere ---------
    link("Products", "Main channel", "Channels", "Channel")
    link("Sales", "Channel", "Channels", "Channel")
    link("Sales", "Promo code", "Promotions", "Code")
    link("Sales", "Event", "Events", "Event")
    link("Bundles", "Channel", "Channels", "Channel")
    link("Price History", "Channel", "Channels", "Channel")
    id_link("Crochet", "Main yarn", "Materials", "Material")
    id_link("Candles", "Wax material", "Materials", "Material")
    id_link("Jewelry", "Metal material", "Materials", "Material")
    id_link("Clothing", "Fabric material", "Materials",
            "Material")
    id_link("Woodworking", "Timber material", "Materials",
            "Material")

    # ---- the shipping option is a built label, so rebuild it ---------
    ship = sheet("Shipping")
    if ship is not None:
        bands = row_dicts(ship, ["Service", "Zone", "Up to (g)"])
        labels = set()
        for band in bands:
            if not band["Service"]:
                continue
            label = f"{band['Service']} - {band['Zone']}"
            if band["Up to (g)"]:
                label += f" to {int(band['Up to (g)']):,}g"
            labels.add(label)
        used = set(column_values(products, "Shipping option"))
        unknown = sorted(used - labels)
        if unknown:
            problems.append(
                f"Products.Shipping option: {len(unknown)} value(s) match "
                f"no postage band -> {unknown[:3]}")
        else:
            notes.append(f"{len(used)} shipping options used, all of them "
                         f"real bands")

    # ---- nothing may be costed twice --------------------------------
    bom_products = {str(v) for v in column_values(sheet("Material Usage"),
                                                  "Product")}
    bom_products |= {str(v) for v in column_values(sheet("Labor"),
                                                   "Product")}
    craft_products: dict = {}
    for craft in CRAFT_SHEETS:
        ws = sheet(craft)
        if ws is None:
            continue
        for value in column_values(ws, "Product ID"):
            craft_products.setdefault(str(value), []).append(craft)
    both = sorted(bom_products & set(craft_products))
    if both:
        problems.append(
            f"{len(both)} product(s) are costed twice - bill of materials "
            f"and a specialist tab -> {both[:4]}")
    else:
        notes.append(f"{len(bom_products)} products costed from the bill "
                     f"of materials, {len(craft_products)} on a "
                     f"specialist tab, none on both")
    twice = sorted(k for k, v in craft_products.items() if len(v) > 1)
    if twice:
        problems.append(f"product(s) on two specialist tabs: {twice[:4]}")

    # every priced product needs costs behind it, or the margin is 100%
    priced = row_dicts(products, ["Product", "Your price"])
    costed = bom_products | set(craft_products)
    uncosted = []
    for index, record in enumerate(priced, start=1):
        if record["Your price"] is None:
            continue
        if f"P-{index:04d}" not in costed:
            uncosted.append(record["Product"])
    if uncosted:
        problems.append(
            f"{len(uncosted)} priced product(s) have no materials and no "
            f"labour, so they look pure profit -> {uncosted[:4]}")
    else:
        notes.append(f"{len(priced)} catalog rows, every priced one has "
                     f"costs behind it")

    # ---- enough of everything to make the tabs look alive ------------
    for name, header, least, label in (
            ("Products", "Product", 10, "products"),
            ("Materials", "Material", 18, "materials"),
            ("Material Usage", "Quantity used", 10, "bill of material lines"),
            ("Labor", "Minutes", 12, "labour rows"),
            ("Packaging", "Component", 8, "packaging components"),
            ("Shipping", "Service", 6, "postage bands"),
            ("Overhead", "Cost", 8, "overhead costs"),
            ("Channels", "Channel", 5, "sales channels"),
            ("Bundles", "Bundle", 3, "bundles"),
            ("Bundle Lines", "Quantity", 5, "bundle lines"),
            ("Variations", "Variation", 4, "variations"),
            ("Batches", "Planned", 4, "production batches"),
            ("Sales", "Date", 40, "sales"),
            ("Promotions", "Code", 3, "promotions"),
            ("Price History", "New price", 6, "price changes"),
            ("Competitors", "Their item", 8, "competitor prices"),
            ("Waste Log", "Quantity lost", 8, "waste entries"),
            ("Development", "Design", 5, "development projects"),
            ("Events", "Event", 3, "craft fairs"),
            ("Wholesale Customers", "Shop", 4, "stockists"),
            ("Wholesale Orders", "Quantity", 8, "wholesale orders")):
        ws = sheet(name)
        if ws is None:
            continue
        try:
            values = column_values(ws, header)
        except KeyError as error:
            problems.append(str(error))
            continue
        if len(values) < least:
            problems.append(f"only {len(values)} {label} - expected at "
                            f"least {least}")
        else:
            notes.append(f"{len(values)} {label}")

    # the six specialist tabs each need a worked example
    for craft in CRAFT_SHEETS:
        ws = sheet(craft)
        if ws is None:
            continue
        header = "Piece" if craft in ("Crochet", "Woodworking") else {
            "Jewelry": "Piece", "Candles": "Candle", "Clothing": "Garment",
            "Digital": "Download"}[craft]
        rows = column_values(ws, header)
        if len(rows) < 2:
            problems.append(f"only {len(rows)} worked example(s) on the "
                            f"{craft} tab - a buyer wants to see it full")

    # ---- the sales ledger has to look like a trading year ------------
    sales = sheet("Sales")
    dates = [v for v in column_values(sales, "Date") if hasattr(v, "year")]
    if len(dates) < 40:
        problems.append(f"only {len(dates)} example sales")
    else:
        span = (max(dates) - min(dates)).days + 1
        if span < 90:
            problems.append(f"the sales only span {span} days - the "
                            f"monthly charts need a year of them")
        else:
            notes.append(f"{len(dates)} sales over {span} days "
                         f"({min(dates):%d %b} to {max(dates):%d %b})")
    months = Counter(f"{d:%Y-%m}" for d in dates)
    if len(months) < 4:
        problems.append(f"the sales only touch {len(months)} month(s)")
    else:
        notes.append(f"sales across {len(months)} months, busiest "
                     f"{months.most_common(1)[0][0]} with "
                     f"{months.most_common(1)[0][1]}")
    channels = Counter(column_values(sales, "Channel"))
    if len(channels) < 3:
        problems.append("the sales all come through "
                        f"{len(channels)} channel(s) - the channel mix "
                        f"on the dashboard needs more")
    else:
        notes.append("sales by channel: " + ", ".join(
            f"{k} {v}" for k, v in channels.most_common(4)))

    # prices and costs: a material with no pack price costs nothing
    packs = row_dicts(materials, ["Material", "Quantity in a pack",
                                  "Pack price"])
    blank = [p["Material"] for p in packs
             if p["Material"] and (not p["Quantity in a pack"]
                                   or not p["Pack price"])]
    if blank:
        problems.append(f"{len(blank)} material(s) have no pack price or "
                        f"quantity, so they cost nothing -> {blank[:4]}")
    else:
        notes.append(f"all {len(packs)} materials have a pack price and a "
                     f"pack quantity")

    # fees must carry a review date, because they go out of date
    reviewed = column_values(sheet("Channels"), "Fees checked")
    if len(reviewed) < len(channel_names):
        problems.append(f"only {len(reviewed)} of {len(channel_names)} "
                        f"channels have a fees-checked date")
    else:
        notes.append(f"all {len(reviewed)} channels carry a "
                     f"fees-checked date")

    print(f"--- {path}")
    for note in notes:
        print(f"  . {note}")
    if problems:
        print(f"  ! {len(problems)} problem(s):")
        for problem in problems[:30]:
            print(f"    - {problem}")
        return 1
    print("  OK - the example data joins up")
    return 0


def main(argv: list) -> int:
    targets = argv[1:] or ["output/Handmade Product Pricing & Profit "
                           "Calculator.xlsx"]
    return max(check(path) for path in targets)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
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

*1,250 lines · 10 functions*

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

**Module constants:** `ROOT` (os.path.dirname(...)), `SHARED_TOOLS` (2 items), `PET_SECTIONS` (5 items), `PET` (9 entries), `SOCIAL_SECTIONS` (7 items), `SOCIAL` (9 entries), `MEAL_SECTIONS` (6 items), `MEAL` (11 entries), `CLEANING_SECTIONS` (6 items), `CLEANING` (11 entries), `HANDMADE_SECTIONS` (7 items), `HANDMADE` (11 entries), `PRODUCTS` (5 entries), `ARCHITECTURE` ('Everything here is plain Python writing a `.xlsx` file t...) and 3 more.

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

# ---------------------------------------------------------------------------
# Product D - cleaning business CRM
# ---------------------------------------------------------------------------

_CLEAN = "cleaning_business_spreadsheet"

CLEANING_SECTIONS = [
    ("1. Run it",
     "Save the files below keeping these exact paths, then run the one "
     "liner. Nothing else is needed - no project file, no build step.\n\n"
     "```bash\n"
     "pip install XlsxWriter\n"
     "python build_cleaning_workbook.py\n"
     "```\n\n"
     "The workbook lands in "
     "`output/Cleaning Business CRM & Pricing Calculator.xlsx`.",
     ["build_cleaning_workbook.py"]),
    ("2. The package",
     "Configuration, the formula template engine, cell formats, the sheet "
     "writer, the cross sheet names and the demo cleaning company. These "
     "are the parts every tab leans on.",
     [f"{_CLEAN}/__init__.py",
      f"{_CLEAN}/__main__.py",
      f"{_CLEAN}/cli.py",
      f"{_CLEAN}/config.py",
      f"{_CLEAN}/layout.py",
      f"{_CLEAN}/formats.py",
      f"{_CLEAN}/writer.py",
      f"{_CLEAN}/names.py",
      f"{_CLEAN}/build.py",
      f"{_CLEAN}/sample_data.py"]),
    ("3. The work registers",
     "One module per data tab. Each returns a `SheetSpec` - a list of "
     "columns with their formulas - and the writer turns that into a "
     "formatted, validated, protected sheet. These are the chain that "
     "turns an enquiry into a priced job and a paid invoice.",
     [f"{_CLEAN}/sheets/__init__.py",
      f"{_CLEAN}/sheets/common.py",
      f"{_CLEAN}/sheets/lists.py",
      f"{_CLEAN}/sheets/settings.py",
      f"{_CLEAN}/sheets/clients.py",
      f"{_CLEAN}/sheets/properties.py",
      f"{_CLEAN}/sheets/leads.py",
      f"{_CLEAN}/sheets/services.py",
      f"{_CLEAN}/sheets/quotes.py",
      f"{_CLEAN}/sheets/quote_lines.py",
      f"{_CLEAN}/sheets/jobs.py",
      f"{_CLEAN}/sheets/job_tasks.py",
      f"{_CLEAN}/sheets/checklists.py"]),
    ("4. The people and money registers",
     "Who did the work, what it cost and what came back in: staff, "
     "timesheets, job costing, invoices and their lines, payments, "
     "supplies, equipment, expenses, mileage, complaints, incidents and "
     "compliance documents.",
     [f"{_CLEAN}/sheets/employees.py",
      f"{_CLEAN}/sheets/timesheets.py",
      f"{_CLEAN}/sheets/job_costing.py",
      f"{_CLEAN}/sheets/invoices.py",
      f"{_CLEAN}/sheets/invoice_lines.py",
      f"{_CLEAN}/sheets/payments.py",
      f"{_CLEAN}/sheets/supplies.py",
      f"{_CLEAN}/sheets/supply_usage.py",
      f"{_CLEAN}/sheets/equipment.py",
      f"{_CLEAN}/sheets/expenses.py",
      f"{_CLEAN}/sheets/mileage.py",
      f"{_CLEAN}/sheets/complaints.py",
      f"{_CLEAN}/sheets/incidents.py",
      f"{_CLEAN}/sheets/safety.py"]),
    ("5. The pages",
     "The tabs that read the registers: the welcome page, the dashboard, "
     "the quote calculator, the month calendar, the staff rota, "
     "retention, the reports, the financial plan, the three printables "
     "and the help tab.",
     [f"{_CLEAN}/sheets/start_here.py",
      f"{_CLEAN}/sheets/dashboard.py",
      f"{_CLEAN}/sheets/quote_calc.py",
      f"{_CLEAN}/sheets/calendar_page.py",
      f"{_CLEAN}/sheets/staff_schedule.py",
      f"{_CLEAN}/sheets/retention.py",
      f"{_CLEAN}/sheets/reports.py",
      f"{_CLEAN}/sheets/financial.py",
      f"{_CLEAN}/sheets/printables.py",
      f"{_CLEAN}/sheets/help_tab.py"]),
    ("6. Quality checks (optional)",
     "Not needed to build the workbook - these are the checks that keep "
     "it honest. The first two need `openpyxl`, the third needs "
     "`formulas`.",
     ["tools/validate_cleaning_workbook.py",
      "tools/check_cleaning_sample.py",
      "tools/recalc_check.py"] + SHARED_TOOLS),
]

CLEANING = {
    "package": "cleaning_business_spreadsheet",
    "builder": "build_cleaning_workbook.py",
    "doc": "README_CLEANING.md",
    "tabs": 37,
    "title": "Cleaning Business CRM & Pricing Calculator",
    "subtitle": ("Complete source code of the generator that writes the "
                 "`.xlsx` product."),
    "output": "SOURCE_CODE_CLEANING.md",
    "closing": ("**Activate the first tab** and close the file - there "
                "is no hidden sheet to hide in this one."),
    "engine_heading": "Ranking columns instead of a hidden engine",
    "intro": [
        "Writes a 37 tab `.xlsx` workbook with over 110,000 formulas: a "
        "client CRM and property database, a lead pipeline, a priced "
        "service catalogue, a quote calculator that sizes a job from the "
        "property and prices it, quotes and estimates, a job scheduler "
        "with recurring visits, a month calendar, property specific "
        "cleaning checklists, staff records, a rota with attendance and "
        "timesheets, job costing and profit per job, invoices, payments, "
        "supplies, equipment, expenses, mileage, retention, complaints, "
        "safety documents, a dashboard, reports, a financial plan and "
        "three printable forms. Document author is **Novality Store**; "
        "every formula cell is locked and every sheet protected with the "
        "password **`premium`**, while the cells you are meant to type "
        "in stay open.",
        "",
        "Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on "
        "the much older 0.3.0, so it does not rely on anything recent. "
        "The workbook itself is Excel 2010 safe \u2014 no dynamic arrays, "
        "no `XLOOKUP`, no `IFS`, nothing that needs array entry \u2014 and "
        "it opens in Google Sheets, LibreOffice and Numbers.",
        "",
        "```bash",
        "python build_cleaning_workbook.py --help",
        "",
        "  -o, --output PATH     where to write the file",
        "  --blank               no example data, just the empty workbook",
        "  --no-protect          leave the sheets unlocked",
        "  --password TEXT       protection password (default: premium)",
        "  --author TEXT         document author  (default: Novality Store)",
        "  --business TEXT       business name on the cover and printables",
        "  --owner TEXT          owner / manager name",
        "  --currency SYMBOL     currency symbol  (default: $)",
        "  --currency-code CODE  three letter code (default: USD)",
        "  --year YYYY           the year the planner and reports open on",
        "  --week-start DAY      Monday or Sunday",
        "  --rate AMOUNT         hourly rate charged to the client",
        "  --pay AMOUNT          hourly pay rate of a cleaner",
        "  --minimum AMOUNT      minimum charge for any job",
        "  --tax PERCENT         sales tax / VAT rate, e.g. 20 for 20%",
        "  --terms DAYS          payment terms in days",
        "  --units {sq ft,sq m}  the area unit used for pricing",
        "  --distance {miles,km} the distance unit used for mileage",
        "  --rows N              rows of capacity in the big registers",
        "  --today YYYY-MM-DD    pretend it is this date (for the demo data)",
        "```",
    ],
    "sections": CLEANING_SECTIONS,
}

# ---------------------------------------------------------------------------
# Product E - handmade product pricing calculator
# ---------------------------------------------------------------------------

_MAKE = "handmade_pricing_spreadsheet"

HANDMADE_SECTIONS = [
    ("1. Run it",
     "Save the files below keeping these exact paths, then run the one "
     "liner. Nothing else is needed - no project file, no build step.\n\n"
     "```bash\n"
     "pip install XlsxWriter\n"
     "python build_handmade_workbook.py\n"
     "```\n\n"
     "The workbook lands in "
     "`output/Handmade Product Pricing & Profit Calculator.xlsx`.",
     ["build_handmade_workbook.py"]),
    ("2. The package",
     "Configuration, the formula template engine, cell formats, the sheet "
     "writer, the cross sheet names and the demo shop. These are the "
     "parts every tab leans on.",
     [f"{_MAKE}/__init__.py",
      f"{_MAKE}/__main__.py",
      f"{_MAKE}/cli.py",
      f"{_MAKE}/config.py",
      f"{_MAKE}/layout.py",
      f"{_MAKE}/formats.py",
      f"{_MAKE}/writer.py",
      f"{_MAKE}/names.py",
      f"{_MAKE}/build.py",
      f"{_MAKE}/sample_data.py"]),
    ("3. The costing registers",
     "One module per data tab. Each returns a `SheetSpec` - a list of "
     "columns with their formulas - and the writer turns that into a "
     "formatted, validated, protected sheet. These are the chain that "
     "turns a pile of materials and a number of minutes into a price: "
     "the catalog, the materials and the bill of materials, labour, "
     "packaging, postage, overhead and the marketplace fee table.",
     [f"{_MAKE}/sheets/__init__.py",
      f"{_MAKE}/sheets/common.py",
      f"{_MAKE}/sheets/lists.py",
      f"{_MAKE}/sheets/products.py",
      f"{_MAKE}/sheets/materials.py",
      f"{_MAKE}/sheets/material_usage.py",
      f"{_MAKE}/sheets/labor.py",
      f"{_MAKE}/sheets/packaging.py",
      f"{_MAKE}/sheets/shipping.py",
      f"{_MAKE}/sheets/overhead.py",
      f"{_MAKE}/sheets/channels.py"]),
    ("4. The six craft calculators",
     "The reason this is not a generic costing sheet. Crochet costs yarn "
     "by the gram, jewelry works in metal weight and scrap recovery, "
     "candles fill a vessel and dose the fragrance, clothing cuts fabric "
     "with a layout allowance, woodworking carries offcut waste and "
     "machine time, and a digital download spreads its design hours over "
     "the sales it expects.",
     [f"{_MAKE}/sheets/crochet.py",
      f"{_MAKE}/sheets/jewelry.py",
      f"{_MAKE}/sheets/candles.py",
      f"{_MAKE}/sheets/clothing.py",
      f"{_MAKE}/sheets/woodworking.py",
      f"{_MAKE}/sheets/digital.py"]),
    ("5. The selling registers",
     "What happens to a product once it has a price: bundles and their "
     "lines, variations and custom orders, batch production, the sales "
     "ledger, promotions, price history, competitor prices, the waste "
     "log, development cost recovery, craft fairs, stockists and "
     "wholesale orders.",
     [f"{_MAKE}/sheets/bundles.py",
      f"{_MAKE}/sheets/bundle_lines.py",
      f"{_MAKE}/sheets/variations.py",
      f"{_MAKE}/sheets/batches.py",
      f"{_MAKE}/sheets/sales.py",
      f"{_MAKE}/sheets/promotions.py",
      f"{_MAKE}/sheets/price_history.py",
      f"{_MAKE}/sheets/competitors.py",
      f"{_MAKE}/sheets/waste.py",
      f"{_MAKE}/sheets/development.py",
      f"{_MAKE}/sheets/events.py",
      f"{_MAKE}/sheets/wholesale_customers.py",
      f"{_MAKE}/sheets/wholesale_orders.py"]),
    ("6. The pages",
     "The tabs that read the registers: the welcome page, settings, the "
     "dashboard, the retail price calculator, the wholesale calculator, "
     "profit and break-even, the pricing scenarios, the reports, the "
     "three printables and the help tab.",
     [f"{_MAKE}/sheets/start_here.py",
      f"{_MAKE}/sheets/settings.py",
      f"{_MAKE}/sheets/dashboard.py",
      f"{_MAKE}/sheets/retail_calc.py",
      f"{_MAKE}/sheets/wholesale_calc.py",
      f"{_MAKE}/sheets/break_even.py",
      f"{_MAKE}/sheets/scenarios.py",
      f"{_MAKE}/sheets/reports.py",
      f"{_MAKE}/sheets/printables.py",
      f"{_MAKE}/sheets/help_tab.py"]),
    ("7. Quality checks (optional)",
     "Not needed to build the workbook - these are the checks that keep "
     "it honest. The first two need `openpyxl`, the third needs "
     "`formulas`.",
     ["tools/validate_handmade_workbook.py",
      "tools/check_handmade_sample.py",
      "tools/recalc_check.py"] + SHARED_TOOLS),
]

HANDMADE = {
    "package": "handmade_pricing_spreadsheet",
    "builder": "build_handmade_workbook.py",
    "doc": "README_HANDMADE.md",
    "tabs": 40,
    "title": "Handmade Product Pricing & Profit Calculator",
    "subtitle": ("Complete source code of the generator that writes the "
                 "`.xlsx` product."),
    "output": "SOURCE_CODE_HANDMADE.md",
    "closing": ("**Activate the first tab** and close the file - the "
                "buyer should land on Start Here, not on the Lists tab."),
    "engine_heading": "Ranking columns instead of a hidden engine",
    "intro": [
        "Writes a 40 tab `.xlsx` workbook with over 97,000 formulas: a "
        "product catalog, a material library costed to four decimal "
        "places, a bill of materials, labour by the minute, packaging "
        "profiles, a postage rate card, overhead allocation, an editable "
        "marketplace fee table, a retail price calculator, a wholesale "
        "calculator, profit and break-even, six craft specific "
        "calculators (crochet, jewelry, candles, clothing, woodworking "
        "and digital downloads), bundles, variations, batch production, "
        "a sales ledger with estimated against actual fees, promotions, "
        "price history, competitor research, a waste log, development "
        "cost recovery, craft fairs, stockists, wholesale orders, "
        "pricing scenarios, reports, three printables and a help tab. "
        "Document author is **Novality Store**; every formula cell is "
        "locked and every sheet protected with the password "
        "**`premium`**, while the cells you are meant to type in stay "
        "open.",
        "",
        "The pricing is the point: cost is materials plus labour plus "
        "packaging plus postage plus overhead, and the price is "
        "`(cost + fixed fees) / (1 - target margin - fee rate)` - a "
        "division, not an addition, so the marketplace percentage comes "
        "out of the price rather than out of your profit.",
        "",
        "Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on "
        "the much older 0.3.0, so it does not rely on anything recent. "
        "The workbook itself is Excel 2010 safe \u2014 no dynamic arrays, "
        "no `XLOOKUP`, no `IFS`, nothing that needs array entry \u2014 and "
        "it opens in Google Sheets, LibreOffice and Numbers.",
        "",
        "```bash",
        "python build_handmade_workbook.py --help",
        "",
        "  -o, --output PATH     where to write the file",
        "  --blank               no example data, just the empty workbook",
        "  --no-protect          leave the sheets unlocked",
        "  --password TEXT       protection password (default: premium)",
        "  --author TEXT         document author  (default: Novality Store)",
        "  --shop TEXT           shop name on the cover and printables",
        "  --owner TEXT          the maker's name",
        "  --currency SYMBOL     currency symbol  (default: $)",
        "  --currency-code CODE  three letter code (default: USD)",
        "  --year YYYY           the year the reports open on",
        "  --rate AMOUNT         what you pay yourself an hour",
        "  --margin PERCENT      target retail margin",
        "  --wholesale PERCENT   the discount a stockist gets off retail",
        "  --overhead AMOUNT     overhead a month, if you would rather",
        "                        type it than list it",
        "  --hours N             productive hours a month",
        "  --rounding RULE       how a calculated price is rounded",
        "  --rows N              rows of capacity in the big registers",
        "  --today YYYY-MM-DD    pretend it is this date (for the demo data)",
        "```",
    ],
    "sections": HANDMADE_SECTIONS,
}

PRODUCTS = {"pet": PET, "social": SOCIAL, "meal": MEAL,
            "cleaning": CLEANING, "handmade": HANDMADE}


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

*286 lines · 2 functions*

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

**Module constants:** `ROOT` (os.path.dirname(...)), `STAMP` (6 items), `SHARED_TOOLS` (3 items), `REQUIREMENTS` ('# The generator needs one package.\nXlsxWriter>=3.0\n\n#...), `PET_HOW_TO_RUN` ('Pet Sitting & Dog Walking Business Spreadsheet\n========...), `SOCIAL_HOW_TO_RUN` ('Social Media Content Calendar & Analytics Planner\n=====...), `MEAL_HOW_TO_RUN` ('Automated Meal Planner & Grocery List Spreadsheet\n=====...), `CLEANING_HOW_TO_RUN` ('Cleaning Business CRM & Pricing Calculator\n============...), `PRODUCTS` (4 entries).

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


CLEANING_HOW_TO_RUN = """\
Cleaning Business CRM & Pricing Calculator
============================================

    pip install XlsxWriter
    python build_cleaning_workbook.py

The workbook appears in:

    output/Cleaning Business CRM & Pricing Calculator.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --business "Name Ltd"   your business name on the cover and prints
    --owner "Your Name"     the owner / manager
    --rate 45               hourly rate charged to the client
    --pay 18                hourly pay rate of a cleaner
    --minimum 90            minimum charge for any job
    --tax 20                sales tax / VAT rate, 0 if you are not registered
    --terms 14              payment terms in days
    --units "sq m"          price area work in square metres
    --distance km           log mileage in kilometres
    --year 2027             the year the planner and reports open on
    --week-start Sunday     weeks run Sunday to Saturday
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 1200             more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

Fill in Settings, then Clients and their Properties, then the Service
Catalog - after that the Pricing & Quote Calculator will price a job
from the property, the Job Scheduler will cost it and the Invoices tab
will bill it.  README_CLEANING.md explains how it is built and what is
on each of the 37 tabs; SOURCE_CODE_CLEANING.md is the whole program in
one readable document.
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
    "cleaning": {
        "folder": "cleaning-business-crm",
        "output": os.path.join(
            "download", "cleaning-business-crm-source.zip"),
        "files": ["build_cleaning_workbook.py", "README_CLEANING.md",
                  "SOURCE_CODE_CLEANING.md"],
        "trees": ["cleaning_business_spreadsheet"],
        "tools": ["tools/validate_cleaning_workbook.py",
                  "tools/check_cleaning_sample.py"] + SHARED_TOOLS,
        "how_to_run": CLEANING_HOW_TO_RUN,
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

#### Products

Every product, every cost, the price the maths asks for and the margin you actually make.

71 columns, 300 rows of capacity, 17 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Product | `name` | text | you type it |
| C | SKU override | `sku_input` | text | you type it |
| D | Category | `category` | text | drop down `CategoryList` |
| E | Status | `status` | status | drop down `StatusList` |
| F | Season | `season` | text | drop down `SeasonList` |
| G | Difficulty | `difficulty` | text | drop down `DifficultyList` |
| H | Main channel | `channel` | text | drop down `ChannelList` |
| I | Packaging | `pack_profile` | text | drop down `PackProfileList` |
| J | Shipping option | `ship_option` | text | drop down `ShipOptions` |
| K | Item weight (g) | `weight` | number1 | you type it (checked) |
| L | In stock | `stock` | int | you type it (checked) |
| M | Your price | `price_input` | money | you type it (checked) |
| N | Target margin | `target_input` | percent | you type it (checked) |
| O | Made so far | `made` | int | you type it (checked) |
| P | Listed on | `launched` | date | you type it (checked) |
| Q | Listing link | `url` | text | you type it |
| R | Notes | `notes` | wrap | you type it |
| S | SKU | `sku` | id | formula |
| T | Materials | `mat_cost` | money | formula |
| U | Labour | `labor_cost` | money | formula |
| V | Minutes logged | `minutes_raw` | number1 | formula, hidden |
| W | Difficulty factor | `time_factor` | number | formula, hidden |
| X | Minutes | `minutes` | number1 | formula |
| Y | Hours | `hours` | number | formula, hidden |
| Z | Labour at difficulty | `labour_adj` | money | formula |
| AA | Consumables | `consumables` | money | formula |
| AB | Packaging | `pack_cost` | money | formula |
| AC | Packaging weight | `pack_weight` | number1 | formula, hidden |
| AD | Postage costs | `ship_pay` | money | formula, hidden |
| AE | Postage charged | `ship_charge` | money | formula, hidden |
| AF | Weight band | `ship_limit` | number1 | formula, hidden |
| AG | Postage you absorb | `ship_cost` | money | formula |
| AH | Cost before overhead | `cost_base` | money | formula |
| AI | Fee rate | `fee_pct` | percent | formula |
| AJ | Fees on postage? | `fee_ship` | flag | formula, hidden |
| AK | Fixed fees | `fee_fixed` | money | formula |
| AL | Overhead rate | `oh_rate` | percent | formula, hidden |
| AM | Overhead in cost | `oh_fixed` | money | formula, hidden |
| AN | Target margin used | `target` | percent | formula |
| AO | Pricing denominator | `denom` | number | formula, hidden |
| AP | Suggested, unrounded | `suggested_raw` | money | formula, hidden |
| AQ | Suggested price | `suggested` | money | formula |
| AR | Price used | `price` | money | formula |
| AS | Overhead | `overhead` | money | formula |
| AT | Total cost | `total_cost` | money | formula |
| AU | Channel fees | `fees` | money | formula |
| AV | Profit | `profit` | money | formula |
| AW | Margin | `margin` | percent | formula |
| AX | Markup | `markup` | percent | formula |
| AY | Profit an hour | `profit_hour` | money | formula |
| AZ | You earn an hour | `wage_hour` | money | formula |
| BA | Never sell below | `min_price` | money | formula |
| BB | Keystone (x2) | `keystone` | money | formula, hidden |
| BC | Wholesale price | `wholesale` | money | formula |
| BD | Wholesale profit | `ws_profit` | money | formula |
| BE | Wholesale margin | `ws_margin` | percent | formula |
| BF | Competitors average | `comp_price` | money | formula |
| BG | Versus the market | `comp_gap` | percent | formula |
| BH | Units sold | `sold` | int | formula |
| BI | Revenue | `revenue` | money | formula |
| BJ | Actual profit | `actual_profit` | money | formula |
| BK | Actual margin | `actual_margin` | percent | formula |
| BL | Last sold | `last_sold` | date | formula |
| BM | Price changed | `last_priced` | date | formula, hidden |
| BN | Days since review | `price_age` | int | formula, hidden |
| BO | Profit rank | `profit_rank` | int | formula, hidden |
| BP | Margin rank | `margin_rank` | int | formula, hidden |
| BQ | Revenue rank | `revenue_rank` | int | formula, hidden |
| BR | Hourly rank | `profit_hour_rank` | int | formula, hidden |
| BS | Action | `alert` | status | formula |

#### Materials

What you buy, and what one unit of it actually costs once discount and delivery are in.

33 columns, 250 rows of capacity, 16 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Material | `name` | text | you type it |
| C | Type | `type` | text | drop down `MaterialTypeList` |
| D | Colour / size / detail | `detail` | text | you type it |
| E | Supplier | `supplier` | text | you type it |
| F | Supplier code | `code` | text | you type it |
| G | Quantity in a pack | `pack_qty` | number1 | you type it (checked) |
| H | Unit | `unit` | text | drop down `UnitList` |
| I | Pack price | `pack_cost` | money | you type it (checked) |
| J | Delivery per pack | `pack_ship` | money | you type it (checked) |
| K | Trade discount | `discount` | percent | you type it (checked) |
| L | Waste allowance | `waste` | percent | you type it (checked) |
| M | Stock on hand | `stock` | number1 | you type it (checked) |
| N | Reorder at | `reorder` | number1 | you type it (checked) |
| O | Previous pack price | `last_cost` | money | you type it (checked) |
| P | Price checked | `priced_on` | date | you type it (checked) |
| Q | Notes | `notes` | wrap | you type it |
| R | Pack cost to you | `net_pack` | money | formula |
| S | Cost a unit | `unit_cost` | money4 | formula |
| T | Counted in | `base_unit` | text | formula, hidden |
| U | Unit factor | `base_factor` | number | formula, hidden |
| V | Cost a base unit | `base_cost` | money4 | formula |
| W | Waste used | `waste_rate` | percent | formula, hidden |
| X | Cost a unit with waste | `cost_waste` | money4 | formula |
| Y | Stock value | `stock_value` | money | formula |
| Z | Used in | `lines` | int | formula |
| AA | Quantity committed | `qty_used` | number1 | formula |
| AB | Cost committed | `cost_used` | money | formula |
| AC | Stock covers | `cover` | int | formula |
| AD | Price move | `change` | percent | formula |
| AE | Stock | `status` | status | formula |
| AF | Spend rank | `cost_used_rank` | int | formula, hidden |
| AG | Action | `alert` | status | formula |

#### Material Usage

What goes into each product, how much of it, and what that costs.

18 columns, 900 rows of capacity, 6 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Product | `product` | id | drop down `ProductIDs` |
| B | Product name | `product_name` | text | formula |
| C | Material | `material` | id | drop down `MaterialIDs` |
| D | Material name | `material_name` | text | formula |
| E | Unit | `unit` | text | formula |
| F | Quantity used | `qty` | number1 | you type it (checked) |
| G | Waste % | `waste` | percent | you type it (checked) |
| H | Optional? | `optional` | flag | drop down `YesNoList` |
| I | Notes | `notes` | wrap | you type it |
| J | Waste used | `waste_rate` | percent | formula, hidden |
| K | Quantity with waste | `qty_eff` | number1 | formula |
| L | Cost a unit | `unit_cost` | money4 | formula |
| M | Line cost | `line_cost` | money4 | formula |
| N | Materials in this product | `product_total` | money | formula, hidden |
| O | Share of the product | `share` | percent | formula |
| P | Line | `seq` | int | formula, hidden |
| Q | Key | `key` | text | formula, hidden |
| R | Check | `check` | status | formula |

#### Labor

Every stage of making it, timed and priced at a real hourly rate.

22 columns, 300 rows of capacity, 9 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Product | `product` | id | drop down `ProductIDs` |
| B | Product name | `product_name` | text | formula |
| C | Stage | `stage` | text | drop down `LaborStageList` |
| D | Who does it | `who` | text | drop down `LaborWhoList` |
| E | Rate override | `rate_input` | money | you type it (checked) |
| F | Minutes | `minutes` | number1 | you type it (checked) |
| G | Per item or batch | `basis` | text | drop down `PerBasisList` |
| H | Items in that batch | `batch_qty` | int | you type it (checked) |
| I | Skilled work? | `skilled` | flag | drop down `YesNoList` |
| J | Notes | `notes` | wrap | you type it |
| K | Hourly rate | `rate` | money | formula |
| L | Hours | `hours` | number | formula |
| M | Batch used | `batch_size` | int | formula, hidden |
| N | Cost of the stage | `stage_cost` | money | formula |
| O | Cost per item | `cost` | money | formula |
| P | Minutes per item | `mins_item` | number1 | formula |
| Q | Consumables | `consumables` | money | formula |
| R | Labour in this product | `product_total` | money | formula, hidden |
| S | Share of the labour | `share` | percent | formula |
| T | Line | `seq` | int | formula, hidden |
| U | Key | `key` | text | formula, hidden |
| V | Check | `check` | status | formula |

#### Packaging

Four packaging profiles, built from the components you actually buy.

26 columns, 80 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Component | `name` | text | you type it |
| B | Type | `type` | text | drop down `PackTypeList` |
| C | Supplier | `supplier` | text | you type it |
| D | Units in a pack | `pack_qty` | number1 | you type it (checked) |
| E | Pack price | `pack_cost` | money | you type it (checked) |
| F | Used per order | `qty` | number1 | you type it (checked) |
| G | Weight each (g) | `weight` | number1 | you type it (checked) |
| H | Budget | `budget` | flag | drop down `YesNoList` |
| I | Standard | `standard` | flag | drop down `YesNoList` |
| J | Gift | `gift` | flag | drop down `YesNoList` |
| K | Premium | `premium` | flag | drop down `YesNoList` |
| L | Branded? | `reusable` | flag | drop down `YesNoList` |
| M | Notes | `notes` | wrap | you type it |
| N | Cost each | `unit_cost` | money4 | formula |
| O | Cost per order | `line_cost` | money4 | formula |
| P | Weight per order | `line_weight` | number1 | formula |
| Q | Budget cost | `c_budget` | money4 | formula, hidden |
| R | Standard cost | `c_standard` | money4 | formula, hidden |
| S | Gift cost | `c_gift` | money4 | formula, hidden |
| T | Premium cost | `c_premium` | money4 | formula, hidden |
| U | Budget weight | `w_budget` | number1 | formula, hidden |
| V | Standard weight | `w_standard` | number1 | formula, hidden |
| W | Gift weight | `w_gift` | number1 | formula, hidden |
| X | Premium weight | `w_premium` | number1 | formula, hidden |
| Y | In profiles | `profiles` | int | formula |
| Z | Check | `check` | status | formula |

#### Shipping

What postage costs you, what the buyer pays, and the difference you have to price for.

18 columns, 60 rows of capacity, 11 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Service | `service` | text | drop down `ShipServiceList` |
| B | Zone | `zone` | text | drop down `ShipZoneList` |
| C | Up to (g) | `max_weight` | number1 | you type it (checked) |
| D | Postage costs you | `cost` | money | you type it (checked) |
| E | Extra packaging | `packaging_extra` | money | you type it (checked) |
| F | Buyer is charged | `charge` | money | you type it (checked) |
| G | Who pays | `payer` | text | drop down `ShipPayerList` |
| H | Tracked? | `tracked` | flag | drop down `YesNoList` |
| I | Days | `days` | int | you type it (checked) |
| J | Carrier | `carrier` | text | you type it |
| K | Notes | `notes` | wrap | you type it |
| L | Option | `label` | text | formula |
| M | Costs you | `total_cost` | money | formula |
| N | You absorb | `absorbed` | money | formula |
| O | You keep | `surplus` | money | formula |
| P | Cost per 100 g | `per_gram` | money | formula, hidden |
| Q | Products using it | `used` | int | formula |
| R | Check | `check` | status | formula |

#### Overhead

The bills that are not about any one product, spread across everything you make.

18 columns, 40 rows of capacity, 9 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Cost | `name` | text | you type it |
| C | Type | `type` | text | drop down `OverheadTypeList` |
| D | Paid to | `supplier` | text | you type it |
| E | Amount | `amount` | money | you type it (checked) |
| F | How often | `frequency` | text | drop down `FrequencyList` |
| G | Started | `start` | date | you type it (checked) |
| H | Ends | `ends` | date | you type it (checked) |
| I | Essential? | `essential` | flag | drop down `YesNoList` |
| J | Notes | `notes` | wrap | you type it |
| K | Active | `active` | flag | formula, hidden |
| L | A month | `monthly` | money | formula |
| M | A year | `yearly` | money | formula |
| N | Share of overhead | `share` | percent | formula |
| O | Per hour worked | `per_hour` | money | formula |
| P | Per item made | `per_unit` | money | formula |
| Q | Items to cover it | `units_to_cover` | int | formula |
| R | Check | `check` | status | formula |

#### Channels

Every fee, editable, per place you sell - and what each one really takes from a sale.

29 columns, 14 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Channel | `name` | text | you type it |
| B | Fee shape | `kind` | text | drop down `FeeTypeList` |
| C | Listing fee | `listing_fee` | money | you type it (checked) |
| D | Sales per listing | `listing_sales` | number1 | you type it (checked) |
| E | Transaction fee | `transaction` | percent | you type it (checked) |
| F | Processing % | `processing_pct` | percent | you type it (checked) |
| G | Processing fixed | `processing_fixed` | money | you type it (checked) |
| H | Offsite ads fee | `ads_fee` | percent | you type it (checked) |
| I | Orders from ads | `ads_share` | percent | you type it (checked) |
| J | Regulatory fee | `regulatory` | percent | you type it (checked) |
| K | Currency fee | `currency_fee` | percent | you type it (checked) |
| L | Other % | `other_pct` | percent | you type it (checked) |
| M | Other fixed | `other_fixed` | money | you type it (checked) |
| N | Monthly subscription | `subscription` | money | you type it (checked) |
| O | Fees on postage? | `fees_on_shipping` | flag | drop down `YesNoList` |
| P | Fees checked | `reviewed` | date | you type it (checked) |
| Q | Fee schedule | `url` | text | you type it |
| R | Notes | `notes` | wrap | you type it |
| S | Percentage fees | `pct_total` | percent | formula |
| T | Fixed fees | `fixed_total` | money | formula |
| U | Fee on a 20 sale | `fee_20` | money | formula |
| V | You keep | `keep_20` | money | formula |
| W | Effective take | `take_rate` | percent | formula |
| X | Products on it | `products` | int | formula |
| Y | Sales recorded | `sales_count` | int | formula |
| Z | Sold through it | `sales_value` | money | formula |
| AA | Fees actually paid | `fees_paid` | money | formula |
| AB | Actual take | `actual_rate` | percent | formula |
| AC | Check | `check` | status | formula |

#### Crochet

Yarn by the gram, notions, the pattern's share and the hours it really takes.

47 columns, 80 rows of capacity, 20 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Product ID | `product` | id | drop down `ProductIDs` |
| B | Piece | `name` | text | you type it |
| C | Pattern | `pattern` | text | you type it |
| D | Size | `size` | text | drop down `SizeNames` |
| E | Yarn weight | `yarn_weight` | text | drop down `YarnWeightList` |
| F | Hook | `hook` | text | drop down `HookSizeList` |
| G | Main yarn | `yarn` | id | drop down `MaterialIDs` |
| H | Grams used | `grams` | number1 | you type it (checked) |
| I | Second yarn | `yarn2` | id | drop down `MaterialIDs` |
| J | Grams used | `grams2` | number1 | you type it (checked) |
| K | Yarn waste | `waste` | percent | you type it (checked) |
| L | Buttons, zips, eyes | `notions` | money | you type it (checked) |
| M | Stuffing / lining | `stuffing` | money | you type it (checked) |
| N | Pattern cost | `pattern_cost` | money | you type it (checked) |
| O | Pieces from it | `pattern_uses` | int | you type it (checked) |
| P | Hours hooking | `hours_input` | number | you type it (checked) |
| Q | Finishing minutes | `finish_mins` | int | you type it (checked) |
| R | Rate override | `rate_input` | money | you type it (checked) |
| S | Your price | `price_input` | money | you type it (checked) |
| T | Notes | `notes` | wrap | you type it |
| U | Size factor | `size_mat` | number | formula, hidden |
| V | Time factor | `size_time` | number | formula, hidden |
| W | Waste used | `waste_rate` | percent | formula, hidden |
| X | Grams, with waste | `grams_total` | number1 | formula |
| Y | Main yarn cost | `yarn_cost` | money | formula |
| Z | Second yarn cost | `yarn2_cost` | money | formula |
| AA | Skeins needed | `skeins` | number1 | formula |
| AB | Grams left over | `leftover` | number1 | formula, hidden |
| AC | Pattern per piece | `pattern_share` | money | formula, hidden |
| AD | Materials | `mat_cost` | money | formula |
| AE | Minutes | `minutes` | number1 | formula |
| AF | Hourly rate | `rate` | money | formula, hidden |
| AG | Labour | `labor_cost` | money | formula |
| AH | Extras | `other_cost` | money | formula, hidden |
| AI | Cost a gram | `cost_per_gram` | money4 | formula, hidden |
| AJ | Cost to make | `total_cost` | money | formula |
| AK | Fee rate | `fee_pct` | percent | formula, hidden |
| AL | Fixed fees | `fee_fixed` | money | formula, hidden |
| AM | Denominator | `denom` | number | formula, hidden |
| AN | Unrounded | `suggested_raw` | money | formula, hidden |
| AO | Suggested price | `suggested` | money | formula |
| AP | Price used | `price` | money | formula |
| AQ | Profit | `profit` | money | formula |
| AR | Margin | `margin` | percent | formula |
| AS | Hours | `hours` | number | formula, hidden |
| AT | You earn an hour | `wage_hour` | money | formula |
| AU | Check | `check` | status | formula |

#### Jewelry

Metal by the gram with scrap, stones, findings, chain and bench time.

49 columns, 80 rows of capacity, 25 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Product ID | `product` | id | drop down `ProductIDs` |
| B | Piece | `name` | text | you type it |
| C | Metal | `metal` | text | drop down `MetalList` |
| D | Metal material | `metal_id` | id | drop down `MaterialIDs` |
| E | Metal used (g) | `metal_g` | number1 | you type it (checked) |
| F | Price a gram | `metal_rate_input` | money4 | you type it (checked) |
| G | Scrap / loss | `scrap` | percent | you type it (checked) |
| H | Scrap recovered | `recovery` | percent | you type it (checked) |
| I | Stone | `stone` | text | drop down `StoneList` |
| J | Stones | `stone_count` | int | you type it (checked) |
| K | Cost a stone | `stone_cost` | money | you type it (checked) |
| L | Chain (cm) | `chain_cm` | number1 | you type it (checked) |
| M | Chain material | `chain_id` | id | drop down `MaterialIDs` |
| N | Wire (cm) | `wire_cm` | number1 | you type it (checked) |
| O | Wire material | `wire_id` | id | drop down `MaterialIDs` |
| P | Findings | `findings` | money | you type it (checked) |
| Q | Solder & consumables | `solder` | money | you type it (checked) |
| R | Plating / finishing | `plating` | money | you type it (checked) |
| S | Hallmarking | `hallmark` | money | you type it (checked) |
| T | Design minutes | `design_mins` | int | you type it (checked) |
| U | Bench minutes | `bench_mins` | int | you type it (checked) |
| V | Polishing minutes | `polish_mins` | int | you type it (checked) |
| W | Rate override | `rate_input` | money | you type it (checked) |
| X | Your price | `price_input` | money | you type it (checked) |
| Y | Notes | `notes` | wrap | you type it |
| Z | Metal a gram | `metal_rate` | money4 | formula, hidden |
| AA | Metal with scrap | `metal_gross` | number1 | formula, hidden |
| AB | Metal cost | `metal_cost` | money | formula |
| AC | Stones cost | `stones_cost` | money | formula |
| AD | Chain cost | `chain_cost` | money | formula |
| AE | Wire cost | `wire_cost` | money | formula |
| AF | Materials | `mat_cost` | money | formula |
| AG | Minutes | `minutes` | number1 | formula |
| AH | Hourly rate | `rate` | money | formula, hidden |
| AI | Labour | `labor_cost` | money | formula |
| AJ | Extras | `other_cost` | money | formula |
| AK | Metal share | `metal_share` | percent | formula, hidden |
| AL | Cost to make | `total_cost` | money | formula |
| AM | Fee rate | `fee_pct` | percent | formula, hidden |
| AN | Fixed fees | `fee_fixed` | money | formula, hidden |
| AO | Denominator | `denom` | number | formula, hidden |
| AP | Unrounded | `suggested_raw` | money | formula, hidden |
| AQ | Suggested price | `suggested` | money | formula |
| AR | Price used | `price` | money | formula |
| AS | Profit | `profit` | money | formula |
| AT | Margin | `margin` | percent | formula |
| AU | Hours | `hours` | number | formula, hidden |
| AV | You earn an hour | `wage_hour` | money | formula |
| AW | Check | `check` | status | formula |

#### Candles

Wax from the vessel size, fragrance load on top, wicks, dye, vessel and the pour.

50 columns, 80 rows of capacity, 25 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Product ID | `product` | id | drop down `ProductIDs` |
| B | Candle | `name` | text | you type it |
| C | Wax | `wax_type` | text | drop down `WaxList` |
| D | Wax material | `wax_id` | id | drop down `MaterialIDs` |
| E | Vessel | `vessel` | text | drop down `VesselList` |
| F | Vessel size (ml) | `vessel_ml` | number1 | you type it (checked) |
| G | Fill level | `fill` | percent | you type it (checked) |
| H | Wax density | `density` | number | you type it (checked) |
| I | Wax override (g) | `wax_input` | number1 | you type it (checked) |
| J | Melt loss | `melt_loss` | percent | you type it (checked) |
| K | Fragrance material | `fragrance_id` | id | drop down `MaterialIDs` |
| L | Fragrance load | `fragrance_load` | percent | you type it (checked) |
| M | Wick | `wick` | text | drop down `WickList` |
| N | Wicks | `wick_count` | int | you type it (checked) |
| O | Cost a wick | `wick_cost` | money | you type it (checked) |
| P | Vessel cost | `vessel_cost` | money | you type it (checked) |
| Q | Dye & additives | `dye_cost` | money | you type it (checked) |
| R | Labels & warning | `label_cost` | money | you type it (checked) |
| S | Pour batch | `batch_size` | int | you type it (checked) |
| T | Batch minutes | `batch_mins` | int | you type it (checked) |
| U | Minutes a candle | `item_mins` | int | you type it (checked) |
| V | Cure days | `cure_days` | int | you type it (checked) |
| W | Rate override | `rate_input` | money | you type it (checked) |
| X | Your price | `price_input` | money | you type it (checked) |
| Y | Notes | `notes` | wrap | you type it |
| Z | Wax poured (g) | `wax_g` | number1 | formula |
| AA | Wax bought (g) | `wax_total` | number1 | formula, hidden |
| AB | Wax cost | `wax_cost` | money | formula |
| AC | Fragrance (g) | `fragrance_g` | number1 | formula, hidden |
| AD | Fragrance cost | `fragrance_cost` | money | formula |
| AE | Wicks cost | `wicks_cost` | money | formula |
| AF | Materials | `mat_cost` | money | formula |
| AG | Minutes | `minutes` | number1 | formula |
| AH | Hourly rate | `rate` | money | formula, hidden |
| AI | Labour | `labor_cost` | money | formula |
| AJ | Extras | `other_cost` | money | formula, hidden |
| AK | Cost per 100 g | `cost_per_100g` | money | formula, hidden |
| AL | Burn hours | `burn_hours` | int | formula, hidden |
| AM | Cost to make | `total_cost` | money | formula |
| AN | Fee rate | `fee_pct` | percent | formula, hidden |
| AO | Fixed fees | `fee_fixed` | money | formula, hidden |
| AP | Denominator | `denom` | number | formula, hidden |
| AQ | Unrounded | `suggested_raw` | money | formula, hidden |
| AR | Suggested price | `suggested` | money | formula |
| AS | Price used | `price` | money | formula |
| AT | Profit | `profit` | money | formula |
| AU | Margin | `margin` | percent | formula |
| AV | Hours | `hours` | number | formula, hidden |
| AW | You earn an hour | `wage_hour` | money | formula |
| AX | Check | `check` | status | formula |

#### Clothing

Fabric with the layout waste in, notions, and all four jobs that make up sewing.

45 columns, 80 rows of capacity, 22 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Product ID | `product` | id | drop down `ProductIDs` |
| B | Garment | `name` | text | you type it |
| C | Type | `garment_type` | text | drop down `FabricList` |
| D | Size | `size` | text | drop down `GarmentSizeList` |
| E | Size factor | `size_factor` | number | you type it (checked) |
| F | Fabric material | `fabric_id` | id | drop down `MaterialIDs` |
| G | Fabric (cm) | `fabric_cm` | number1 | you type it (checked) |
| H | Layout waste | `layout_waste` | percent | you type it (checked) |
| I | Lining material | `lining_id` | id | drop down `MaterialIDs` |
| J | Lining (cm) | `lining_cm` | number1 | you type it (checked) |
| K | Interfacing | `interfacing` | money | you type it (checked) |
| L | Thread | `thread` | money | you type it (checked) |
| M | Zips, buttons, elastic | `notions` | money | you type it (checked) |
| N | Labels & care tag | `labels` | money | you type it (checked) |
| O | Cutting minutes | `cut_mins` | int | you type it (checked) |
| P | Sewing minutes | `sew_mins` | int | you type it (checked) |
| Q | Finishing minutes | `finish_mins` | int | you type it (checked) |
| R | Pressing minutes | `press_mins` | int | you type it (checked) |
| S | Fitting & alterations | `fit_mins` | int | you type it (checked) |
| T | Rate override | `rate_input` | money | you type it (checked) |
| U | Your price | `price_input` | money | you type it (checked) |
| V | Notes | `notes` | wrap | you type it |
| W | Size factor used | `factor` | number | formula, hidden |
| X | Waste used | `waste_rate` | percent | formula, hidden |
| Y | Fabric bought (cm) | `fabric_total` | number1 | formula |
| Z | Fabric cost | `fabric_cost` | money | formula |
| AA | Lining cost | `lining_cost` | money | formula |
| AB | Materials | `mat_cost` | money | formula |
| AC | Minutes | `minutes` | number1 | formula |
| AD | Hourly rate | `rate` | money | formula, hidden |
| AE | Labour | `labor_cost` | money | formula |
| AF | Extras | `other_cost` | money | formula, hidden |
| AG | Cost of the waste | `waste_cost` | money | formula, hidden |
| AH | Cost to make | `total_cost` | money | formula |
| AI | Fee rate | `fee_pct` | percent | formula, hidden |
| AJ | Fixed fees | `fee_fixed` | money | formula, hidden |
| AK | Denominator | `denom` | number | formula, hidden |
| AL | Unrounded | `suggested_raw` | money | formula, hidden |
| AM | Suggested price | `suggested` | money | formula |
| AN | Price used | `price` | money | formula |
| AO | Profit | `profit` | money | formula |
| AP | Margin | `margin` | percent | formula |
| AQ | Hours | `hours` | number | formula, hidden |
| AR | You earn an hour | `wage_hour` | money | formula |
| AS | Check | `check` | status | formula |

#### Woodworking

Timber with the offcuts in, hardware, finish, abrasives and machine time.

47 columns, 80 rows of capacity, 24 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Product ID | `product` | id | drop down `ProductIDs` |
| B | Piece | `name` | text | you type it |
| C | Species | `species` | text | drop down `WoodList` |
| D | Timber material | `wood_id` | id | drop down `MaterialIDs` |
| E | Timber used (cm) | `wood_qty` | number1 | you type it (checked) |
| F | Offcut waste | `offcut_waste` | percent | you type it (checked) |
| G | Second timber | `second_id` | id | drop down `MaterialIDs` |
| H | Second timber (cm) | `second_qty` | number1 | you type it (checked) |
| I | Hardware | `hardware` | money | you type it (checked) |
| J | Glue & fixings | `glue` | money | you type it (checked) |
| K | Finish | `finish_type` | text | drop down `FinishList` |
| L | Finish used (ml) | `finish_ml` | number1 | you type it (checked) |
| M | Finish cost a ml | `finish_rate` | money4 | you type it (checked) |
| N | Abrasives & blades | `abrasives` | money | you type it (checked) |
| O | Machine minutes | `machine_mins` | int | you type it (checked) |
| P | Machine rate an hour | `machine_rate` | money | you type it (checked) |
| Q | Milling minutes | `mill_mins` | int | you type it (checked) |
| R | Assembly minutes | `build_mins` | int | you type it (checked) |
| S | Sanding minutes | `sand_mins` | int | you type it (checked) |
| T | Finishing minutes | `finish_mins` | int | you type it (checked) |
| U | Drying days | `dry_days` | int | you type it (checked) |
| V | Rate override | `rate_input` | money | you type it (checked) |
| W | Your price | `price_input` | money | you type it (checked) |
| X | Notes | `notes` | wrap | you type it |
| Y | Waste used | `waste_rate` | percent | formula, hidden |
| Z | Timber bought (cm) | `wood_total` | number1 | formula, hidden |
| AA | Timber cost | `wood_cost` | money | formula |
| AB | Finish cost | `finish_cost` | money | formula |
| AC | Materials | `mat_cost` | money | formula |
| AD | Minutes | `minutes` | number1 | formula |
| AE | Hourly rate | `rate` | money | formula, hidden |
| AF | Labour | `labor_cost` | money | formula |
| AG | Machine time | `machine_cost` | money | formula |
| AH | Extras | `other_cost` | money | formula, hidden |
| AI | Cost of the offcuts | `waste_cost` | money | formula, hidden |
| AJ | Cost to make | `total_cost` | money | formula |
| AK | Fee rate | `fee_pct` | percent | formula, hidden |
| AL | Fixed fees | `fee_fixed` | money | formula, hidden |
| AM | Denominator | `denom` | number | formula, hidden |
| AN | Unrounded | `suggested_raw` | money | formula, hidden |
| AO | Suggested price | `suggested` | money | formula |
| AP | Price used | `price` | money | formula |
| AQ | Profit | `profit` | money | formula |
| AR | Margin | `margin` | percent | formula |
| AS | Hours | `hours` | number | formula, hidden |
| AT | You earn an hour | `wage_hour` | money | formula |
| AU | Check | `check` | status | formula |

#### Digital

Design time spread over the sales you expect, and the day it pays for itself.

42 columns, 80 rows of capacity, 19 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Product ID | `product` | id | drop down `ProductIDs` |
| B | Download | `name` | text | you type it |
| C | Type | `kind` | text | drop down `DigitalTypeList` |
| D | Formats | `formats` | text | drop down `FileFormatList` |
| E | Licence | `licence` | text | drop down `LicenceList` |
| F | Design hours | `design_hours` | number | you type it (checked) |
| G | Writing hours | `write_hours` | number | you type it (checked) |
| H | Testing hours | `test_hours` | number | you type it (checked) |
| I | Photography hours | `photo_hours` | number | you type it (checked) |
| J | Tester & editor fees | `tester_cost` | money | you type it (checked) |
| K | Software & stock | `software_cost` | money | you type it (checked) |
| L | Sales expected | `expected_sales` | int | you type it (checked) |
| M | Support minutes a sale | `support_mins` | number1 | you type it (checked) |
| N | Updates a year | `update_hours` | number | you type it (checked) |
| O | Sold so far | `sold_so_far` | int | you type it (checked) |
| P | Listed on | `listed` | date | you type it (checked) |
| Q | Rate override | `rate_input` | money | you type it (checked) |
| R | Your price | `price_input` | money | you type it (checked) |
| S | Notes | `notes` | wrap | you type it |
| T | Hourly rate | `rate` | money | formula, hidden |
| U | Development hours | `dev_hours` | number | formula |
| V | Development cost | `dev_cost` | money | formula |
| W | Sales used | `expected` | int | formula, hidden |
| X | Materials | `mat_cost` | money | formula, hidden |
| Y | Minutes a sale | `minutes` | number1 | formula |
| Z | Labour a sale | `labor_cost` | money | formula |
| AA | Extras a sale | `other_cost` | money | formula, hidden |
| AB | Recovered so far | `recovered` | money | formula |
| AC | Development recovered | `recovery` | percent | formula |
| AD | Sales to break even | `break_even` | int | formula |
| AE | Cost to make | `total_cost` | money | formula |
| AF | Fee rate | `fee_pct` | percent | formula, hidden |
| AG | Fixed fees | `fee_fixed` | money | formula, hidden |
| AH | Denominator | `denom` | number | formula, hidden |
| AI | Unrounded | `suggested_raw` | money | formula, hidden |
| AJ | Suggested price | `suggested` | money | formula |
| AK | Price used | `price` | money | formula |
| AL | Profit | `profit` | money | formula |
| AM | Margin | `margin` | percent | formula |
| AN | Hours | `hours` | number | formula, hidden |
| AO | You earn an hour | `wage_hour` | money | formula |
| AP | Check | `check` | status | formula |

#### Bundles

Multi-item sets, with the discount, the box and the assembly time all counted.

34 columns, 60 rows of capacity, 10 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Bundle | `name` | text | you type it |
| C | Type | `kind` | text | drop down `BundleTypeList` |
| D | Channel | `channel` | text | drop down `ChannelList` |
| E | Packaging | `pack_profile` | text | drop down `PackProfileList` |
| F | Assembly minutes | `assembly_mins` | int | you type it (checked) |
| G | Discount off retail | `discount` | percent | you type it (checked) |
| H | Your price | `price_input` | money | you type it (checked) |
| I | Status | `status` | status | drop down `StatusList` |
| J | Season | `season` | text | drop down `SeasonList` |
| K | Notes | `notes` | wrap | you type it |
| L | Items | `items` | int | formula |
| M | Products | `lines` | int | formula, hidden |
| N | Components cost | `parts_cost` | money | formula |
| O | Bought separately | `parts_price` | money | formula |
| P | Packaging | `pack_cost` | money | formula |
| Q | Assembly | `assembly_cost` | money | formula |
| R | Total cost | `total_cost` | money | formula |
| S | Fee rate | `fee_pct` | percent | formula, hidden |
| T | Fixed fees | `fee_fixed` | money | formula, hidden |
| U | Denominator | `denom` | number | formula, hidden |
| V | Unrounded | `suggested_raw` | money | formula, hidden |
| W | Suggested price | `suggested` | money | formula |
| X | Discounted price | `discounted` | money | formula, hidden |
| Y | Price used | `price` | money | formula |
| Z | Channel fees | `fees` | money | formula |
| AA | Profit | `profit` | money | formula |
| AB | Margin | `margin` | percent | formula |
| AC | Buyer saves | `saving` | money | formula |
| AD | Saving | `saving_pct` | percent | formula, hidden |
| AE | Separate margin | `parts_margin` | percent | formula, hidden |
| AF | Margin given away | `margin_lost` | percent | formula |
| AG | Bundles sold | `sold` | int | formula |
| AH | Action | `alert` | status | formula |

#### Bundle Lines

What goes in each set, costed from the catalog.

14 columns, 200 rows of capacity, 5 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Bundle | `bundle` | id | drop down `BundleIDs` |
| B | Bundle name | `bundle_name` | text | formula |
| C | Product | `product` | id | drop down `ProductIDs` |
| D | Product name | `product_name` | text | formula |
| E | Quantity | `qty` | int | you type it (checked) |
| F | Substitutable? | `swap` | flag | drop down `YesNoList` |
| G | Notes | `notes` | wrap | you type it |
| H | Cost each | `unit_cost` | money | formula |
| I | Retail each | `unit_price` | money | formula |
| J | Line cost | `line_cost` | money | formula |
| K | Line retail | `line_price` | money | formula |
| L | Line | `seq` | int | formula, hidden |
| M | Key | `key` | text | formula, hidden |
| N | Check | `check` | status | formula |

#### Variations

Sizes, personalisation and rush work, priced from the base product.

38 columns, 300 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Base product | `product` | id | drop down `ProductIDs` |
| C | Variation | `name` | text | you type it |
| D | What changes | `kind` | text | drop down `VariationTypeList` |
| E | Size | `size` | text | drop down `SizeNames` |
| F | Extra materials | `extra_material` | money | you type it (checked) |
| G | Extra minutes | `extra_mins` | int | you type it (checked) |
| H | Turnaround | `rush` | text | drop down `RushNames` |
| I | Personalisation fee | `personalisation` | money | you type it (checked) |
| J | Channel | `channel` | text | drop down `ChannelList` |
| K | Your price | `price_input` | money | you type it (checked) |
| L | Offered? | `offered` | flag | drop down `YesNoList` |
| M | Notes | `notes` | wrap | you type it |
| N | Base product name | `base_name` | text | formula, hidden |
| O | Base materials | `base_mat` | money | formula, hidden |
| P | Base labour | `base_labour` | money | formula, hidden |
| Q | Base other | `base_other` | money | formula, hidden |
| R | Base minutes | `base_minutes` | number1 | formula, hidden |
| S | Base price | `base_price` | money | formula |
| T | Size factor | `size_mat` | number | formula, hidden |
| U | Time factor | `size_time` | number | formula, hidden |
| V | Rush surcharge | `rush_rate` | percent | formula, hidden |
| W | Materials | `mat_cost` | money | formula |
| X | Minutes | `minutes` | number1 | formula |
| Y | Labour | `labor_cost` | money | formula |
| Z | Total cost | `total_cost` | money | formula |
| AA | Fee rate | `fee_pct` | percent | formula, hidden |
| AB | Fixed fees | `fee_fixed` | money | formula, hidden |
| AC | Denominator | `denom` | number | formula, hidden |
| AD | Unrounded | `suggested_raw` | money | formula, hidden |
| AE | Suggested price | `suggested` | money | formula |
| AF | Price used | `price` | money | formula |
| AG | Versus the base | `uplift` | money | formula |
| AH | Channel fees | `fees` | money | formula, hidden |
| AI | Profit | `profit` | money | formula |
| AJ | Margin | `margin` | percent | formula |
| AK | You earn an hour | `wage_hour` | money | formula |
| AL | Action | `alert` | status | formula |

#### Batches

What a run really cost, what it really took, and how far the estimate was out.

31 columns, 200 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Product | `product` | id | drop down `ProductIDs` |
| C | Product name | `product_name` | text | formula |
| D | Started | `started` | date | you type it (checked) |
| E | Finished | `finished` | date | you type it (checked) |
| F | Planned | `planned` | int | you type it (checked) |
| G | Made | `made` | int | you type it (checked) |
| H | Seconds / failures | `seconds` | int | you type it (checked) |
| I | Seconds sold for | `sold_seconds` | money | you type it (checked) |
| J | Materials actually | `material_actual` | money | you type it (checked) |
| K | Hours actually | `hours_actual` | number | you type it (checked) |
| L | Other costs | `extra_cost` | money | you type it (checked) |
| M | Status | `status` | status | drop down `BatchStatusList` |
| N | Notes | `notes` | wrap | you type it |
| O | Good units | `good` | int | formula |
| P | Failure rate | `defect_rate` | percent | formula |
| Q | Materials estimated | `est_mat_unit` | money | formula, hidden |
| R | Minutes estimated | `est_minutes` | number1 | formula, hidden |
| S | Materials cost | `mat_total` | money | formula |
| T | Labour cost | `labour_total` | money | formula |
| U | Batch cost | `batch_cost` | money | formula |
| V | Cost a good unit | `cost_good` | money | formula |
| W | Estimated a unit | `est_unit` | money | formula, hidden |
| X | Versus estimate | `variance` | money | formula |
| Y | Out by | `variance_pct` | percent | formula, hidden |
| Z | Minutes a unit | `mins_actual` | number1 | formula |
| AA | Minutes out by | `time_variance` | number1 | formula |
| AB | Batch retail value | `value` | money | formula |
| AC | Profit if it all sells | `batch_profit` | money | formula |
| AD | Days taken | `days` | int | formula, hidden |
| AE | Action | `alert` | status | formula |

#### Sales

Every order, with the real postage, the real fees and the real margin.

39 columns, 800 rows of capacity, 15 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Product | `product` | id | drop down `ProductIDs` |
| D | Qty | `qty` | int | you type it (checked) |
| E | Channel | `channel` | text | drop down `ChannelList` |
| F | Price each | `unit_price` | money | you type it (checked) |
| G | Discount given | `discount` | money | you type it (checked) |
| H | Promo code | `promo` | text | drop down `PromoCodes` |
| I | Postage charged | `ship_charged` | money | you type it (checked) |
| J | Postage paid | `ship_paid` | money | you type it (checked) |
| K | Fees actually taken | `fees_input` | money | you type it (checked) |
| L | Refunded | `refund` | money | you type it (checked) |
| M | Status | `status` | status | drop down `SaleStatusList` |
| N | Customer | `customer` | text | you type it |
| O | Event | `event` | text | drop down `EventNames` |
| P | Notes | `notes` | wrap | you type it |
| Q | Product name | `product_name` | text | formula |
| R | Month | `month` | text | formula |
| S | Catalog price | `catalog_price` | money | formula, hidden |
| T | Price used | `price_used` | money | formula, hidden |
| U | Gross | `gross` | money | formula, hidden |
| V | Revenue | `revenue` | money | formula |
| W | Catalog cost | `unit_cost_total` | money | formula, hidden |
| X | Catalog postage | `unit_ship` | money | formula, hidden |
| Y | Cost of the goods | `cogs` | money | formula |
| Z | Postage you absorbed | `ship_absorbed` | money | formula |
| AA | Fee rate | `fee_pct` | percent | formula, hidden |
| AB | Fixed fees | `fee_fixed` | money | formula, hidden |
| AC | Fees on postage? | `fee_ship` | flag | formula, hidden |
| AD | Fees estimated | `fees_est` | money | formula, hidden |
| AE | Fees | `fees` | money | formula |
| AF | Fees versus estimate | `fee_gap` | money | formula, hidden |
| AG | Profit | `profit` | money | formula |
| AH | Margin | `margin` | percent | formula |
| AI | Planned margin | `planned_margin` | percent | formula, hidden |
| AJ | Versus plan | `margin_gap` | percent | formula, hidden |
| AK | Hours of making | `hours` | number | formula, hidden |
| AL | Profit an hour | `profit_hour` | money | formula, hidden |
| AM | Action | `alert` | status | formula |

#### Promotions

What each sale costs in margin, and how many extra orders it has to bring.

24 columns, 40 rows of capacity, 11 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Code | `code` | id | you type it |
| B | Promotion | `name` | text | you type it |
| C | Type | `kind` | text | drop down `PromoTypeList` |
| D | Value | `value` | number | you type it (checked) |
| E | Applies to | `scope` | text | you type it |
| F | Starts | `start` | date | you type it (checked) |
| G | Ends | `ends` | date | you type it (checked) |
| H | Minimum spend | `min_spend` | money | you type it (checked) |
| I | Use limit | `limit` | int | you type it (checked) |
| J | Normal margin | `base_margin` | percent | you type it (checked) |
| K | Notes | `notes` | wrap | you type it |
| L | Margin used | `margin_used` | percent | formula, hidden |
| M | Discount | `discount_pct` | percent | formula, hidden |
| N | Margin after it | `margin_after` | percent | formula |
| O | Profit given away | `profit_lost` | percent | formula |
| P | Extra sales needed | `uplift_needed` | percent | formula |
| Q | Orders | `orders` | int | formula |
| R | Revenue | `revenue` | money | formula |
| S | Discount given | `given` | money | formula |
| T | Profit made | `profit` | money | formula |
| U | Actual margin | `actual_margin` | percent | formula |
| V | Average order | `avg_order` | money | formula, hidden |
| W | Status | `status` | status | formula |
| X | Action | `alert` | status | formula |

#### Price History

Every change, the reason for it, and what it did to the margin and the volume.

23 columns, 300 rows of capacity, 10 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Date | `date` | date | you type it (checked) |
| B | Product | `product` | id | drop down `ProductIDs` |
| C | Product name | `product_name` | text | formula |
| D | Old price | `old_price` | money | you type it (checked) |
| E | New price | `new_price` | money | you type it (checked) |
| F | Cost at the time | `cost_then` | money | you type it (checked) |
| G | Reason | `reason` | text | drop down `PriceReasonList` |
| H | Channel | `channel` | text | drop down `ChannelList` |
| I | Sold in the 30 days before | `sold_before` | int | you type it (checked) |
| J | Sold in the 30 days after | `sold_after` | int | you type it (checked) |
| K | Notes | `notes` | wrap | you type it |
| L | Month | `month` | text | formula |
| M | Change | `change` | money | formula |
| N | Change | `change_pct` | percent | formula |
| O | Cost used | `cost_used` | money | formula, hidden |
| P | Margin before | `margin_before` | percent | formula |
| Q | Margin after | `margin_after` | percent | formula |
| R | Margin moved | `margin_move` | percent | formula, hidden |
| S | Sales moved | `volume_move` | percent | formula |
| T | Revenue effect | `revenue_move` | percent | formula |
| U | Change number | `seq` | int | formula, hidden |
| V | Days since | `days_since` | int | formula, hidden |
| W | Verdict | `verdict` | status | formula |

#### Competitors

What comparable work sells for, postage included, and where that puts you.

22 columns, 150 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Your product | `product` | id | drop down `ProductIDs` |
| B | Your product name | `product_name` | text | formula |
| C | Their item | `competitor` | text | you type it |
| D | Shop | `shop` | text | you type it |
| E | Where | `channel` | text | drop down `ChannelList` |
| F | Their price | `price` | money | you type it (checked) |
| G | Their postage | `shipping` | money | you type it (checked) |
| H | Quality | `quality` | text | drop down `QualityList` |
| I | Really handmade? | `handmade` | flag | drop down `YesNoNAList` |
| J | Reviews | `reviews` | int | you type it (checked) |
| K | Checked on | `checked` | date | you type it (checked) |
| L | Link | `url` | text | you type it |
| M | Notes | `notes` | wrap | you type it |
| N | Their total | `their_total` | money | formula |
| O | Your price | `your_price` | money | formula, hidden |
| P | Your postage | `your_ship` | money | formula, hidden |
| Q | Your total | `your_total` | money | formula |
| R | Difference | `gap` | money | formula |
| S | You are | `gap_pct` | percent | formula |
| T | Position | `position` | status | formula |
| U | Days old | `stale` | int | formula, hidden |
| V | Note | `check` | status | formula |

#### Waste Log

What got wasted, what it cost, and how much of it you could stop.

14 columns, 200 rows of capacity, 7 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Date | `date` | date | you type it (checked) |
| B | Material | `material` | id | drop down `MaterialIDs` |
| C | Material name | `material_name` | text | formula |
| D | Quantity lost | `qty` | number1 | you type it (checked) |
| E | What happened | `reason` | text | drop down `WasteReasonList` |
| F | On which product | `product` | id | drop down `ProductIDs` |
| G | Avoidable? | `avoidable` | flag | drop down `YesNoList` |
| H | Notes | `notes` | wrap | you type it |
| I | Month | `month` | text | formula |
| J | Unit | `unit` | text | formula, hidden |
| K | Cost a unit | `unit_cost` | money4 | formula, hidden |
| L | Cost | `cost` | money | formula |
| M | Share of that material | `share` | percent | formula, hidden |
| N | Note | `check` | status | formula |

#### Development

What each design cost to create, and how many sales pay that back.

22 columns, 80 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Design | `name` | text | you type it |
| C | Became product | `product` | id | drop down `ProductIDs` |
| D | Stage | `stage` | text | drop down `DevStageList` |
| E | Started | `started` | date | you type it (checked) |
| F | Design hours | `design_hours` | number | you type it (checked) |
| G | Sample hours | `sample_hours` | number | you type it (checked) |
| H | Photo & listing hours | `photo_hours` | number | you type it (checked) |
| I | Sample materials | `sample_cost` | money | you type it (checked) |
| J | Other costs | `other_cost` | money | you type it (checked) |
| K | Samples sold for | `samples_sold` | money | you type it (checked) |
| L | Decision date | `decided` | date | you type it (checked) |
| M | Notes | `notes` | wrap | you type it |
| N | Hours | `hours` | number | formula, hidden |
| O | Development cost | `dev_cost` | money | formula |
| P | Profit a sale | `unit_profit` | money | formula, hidden |
| Q | Sold since | `sold` | int | formula |
| R | Recovered | `recovered` | money | formula |
| S | Recovered | `recovery` | percent | formula |
| T | Sales to recover it | `to_recover` | int | formula |
| U | Development a sale | `per_unit` | money | formula, hidden |
| V | Action | `alert` | status | formula |

#### Events

What the stall costs, what it has to take, and what it actually did.

28 columns, 40 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Event | `name` | text | you type it |
| B | Type | `kind` | text | drop down `EventTypeList` |
| C | Date | `date` | date | you type it (checked) |
| D | Days | `days` | number1 | you type it (checked) |
| E | Stall fee | `stall_fee` | money | you type it (checked) |
| F | Travel & parking | `travel` | money | you type it (checked) |
| G | Accommodation | `accommodation` | money | you type it (checked) |
| H | Stock, display, other | `other_cost` | money | you type it (checked) |
| I | Prep hours | `prep_hours` | number | you type it (checked) |
| J | Hours on the stall | `stall_hours` | number | you type it (checked) |
| K | Sales target | `target` | money | you type it (checked) |
| L | Footfall | `footfall` | int | you type it (checked) |
| M | Notes | `notes` | wrap | you type it |
| N | Money out | `cash_cost` | money | formula |
| O | Hours | `hours` | number | formula, hidden |
| P | Your time | `time_cost` | money | formula |
| Q | Total cost | `total_cost` | money | formula |
| R | Sales to break even | `break_even` | money | formula |
| S | Cash break-even | `be_cash` | money | formula, hidden |
| T | Actually took | `sales` | money | formula |
| U | Orders | `orders` | int | formula |
| V | Profit on the goods | `gross_profit` | money | formula, hidden |
| W | Profit after the stall | `net` | money | formula |
| X | Earned an hour | `per_hour` | money | formula |
| Y | Average sale | `avg_order` | money | formula, hidden |
| Z | Buyers per 100 passing | `conversion` | percent | formula, hidden |
| AA | Status | `status` | status | formula |
| AB | Verdict | `verdict` | status | formula |

#### Wholesale Customers

The shops that stock you, their terms, and what each one is really worth.

27 columns, 60 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Shop | `name` | text | you type it |
| C | Type | `kind` | text | drop down `CustomerTypeList` |
| D | Contact | `contact` | text | you type it |
| E | Email | `email` | text | you type it |
| F | Phone | `phone` | text | you type it |
| G | Town / city | `city` | text | you type it |
| H | Payment terms | `terms` | text | drop down `TermsList` |
| I | Their discount | `discount` | percent | you type it (checked) |
| J | Their minimum | `min_order` | money | you type it (checked), hidden |
| K | First order | `first_order` | date | you type it (checked) |
| L | Stockist since | `stockist_since` | date | you type it (checked), hidden |
| M | Exclusive area? | `exclusive` | flag | drop down `YesNoList` |
| N | Still buying? | `active` | flag | drop down `YesNoList` |
| O | Notes | `notes` | wrap | you type it |
| P | Discount used | `discount_used` | percent | formula, hidden |
| Q | Orders | `orders` | int | formula |
| R | Items bought | `units` | int | formula |
| S | Total spent | `value` | money | formula |
| T | Profit from them | `profit` | money | formula |
| U | Margin | `margin` | percent | formula |
| V | Average order | `avg_order` | money | formula |
| W | Last order | `last_order` | date | formula |
| X | Days since | `days_quiet` | int | formula, hidden |
| Y | Unpaid | `outstanding` | money | formula |
| Z | Status | `status` | status | formula |
| AA | Action | `alert` | status | formula |

#### Wholesale Orders

Line by line, at the shop's own discount, with the hours it will take you.

30 columns, 200 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Customer | `customer` | id | drop down `CustomerIDs` |
| D | Their order ref | `po_ref` | text | you type it |
| E | Product | `product` | id | drop down `ProductIDs` |
| F | Quantity | `qty` | int | you type it (checked) |
| G | Agreed price each | `price_input` | money | you type it (checked) |
| H | Due | `due` | date | you type it (checked) |
| I | Despatched | `despatched` | date | you type it (checked) |
| J | Carriage charged | `carriage` | money | you type it (checked), hidden |
| K | Carriage cost | `carriage_cost` | money | you type it (checked), hidden |
| L | Status | `status` | status | drop down `OrderStatusList` |
| M | Paid? | `paid` | flag | drop down `YesNoList` |
| N | Notes | `notes` | wrap | you type it |
| O | Shop | `customer_name` | text | formula |
| P | Product name | `product_name` | text | formula |
| Q | Month | `month` | text | formula |
| R | Discount | `discount` | percent | formula, hidden |
| S | Retail each | `retail` | money | formula, hidden |
| T | Price each | `unit_price` | money | formula |
| U | Cost each | `unit_cost` | money | formula, hidden |
| V | Order value | `order_value` | money | formula |
| W | Order cost | `order_cost` | money | formula |
| X | Profit | `profit` | money | formula |
| Y | Margin | `margin` | percent | formula |
| Z | Hours of making | `hours` | number | formula |
| AA | Profit an hour | `per_hour` | money | formula, hidden |
| AB | Meets minimum? | `meets_min` | flag | formula, hidden |
| AC | Days to despatch | `lead_days` | int | formula, hidden |
| AD | Action | `alert` | status | formula |

---

*15,800 lines of Python in 56 files.*
