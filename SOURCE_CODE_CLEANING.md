# Cleaning Business CRM & Pricing Calculator — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 37 tab `.xlsx` workbook with over 110,000 formulas: a client CRM and property database, a lead pipeline, a priced service catalogue, a quote calculator that sizes a job from the property and prices it, quotes and estimates, a job scheduler with recurring visits, a month calendar, property specific cleaning checklists, staff records, a rota with attendance and timesheets, job costing and profit per job, invoices, payments, supplies, equipment, expenses, mileage, retention, complaints, safety documents, a dashboard, reports, a financial plan and three printable forms. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_cleaning_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --business TEXT       business name on the cover and printables
  --owner TEXT          owner / manager name
  --currency SYMBOL     currency symbol  (default: $)
  --currency-code CODE  three letter code (default: USD)
  --year YYYY           the year the planner and reports open on
  --week-start DAY      Monday or Sunday
  --rate AMOUNT         hourly rate charged to the client
  --pay AMOUNT          hourly pay rate of a cleaner
  --minimum AMOUNT      minimum charge for any job
  --tax PERCENT         sales tax / VAT rate, e.g. 20 for 20%
  --terms DAYS          payment terms in days
  --units {sq ft,sq m}  the area unit used for pricing
  --distance {miles,km} the distance unit used for mileage
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
  * [`build_cleaning_workbook.py`](#build_cleaning_workbookpy)
* [2. The package](#2-the-package)
  * [`cleaning_business_spreadsheet/__init__.py`](#cleaning_business_spreadsheet__init__py)
  * [`cleaning_business_spreadsheet/__main__.py`](#cleaning_business_spreadsheet__main__py)
  * [`cleaning_business_spreadsheet/cli.py`](#cleaning_business_spreadsheetclipy)
  * [`cleaning_business_spreadsheet/config.py`](#cleaning_business_spreadsheetconfigpy)
  * [`cleaning_business_spreadsheet/layout.py`](#cleaning_business_spreadsheetlayoutpy)
  * [`cleaning_business_spreadsheet/formats.py`](#cleaning_business_spreadsheetformatspy)
  * [`cleaning_business_spreadsheet/writer.py`](#cleaning_business_spreadsheetwriterpy)
  * [`cleaning_business_spreadsheet/names.py`](#cleaning_business_spreadsheetnamespy)
  * [`cleaning_business_spreadsheet/build.py`](#cleaning_business_spreadsheetbuildpy)
  * [`cleaning_business_spreadsheet/sample_data.py`](#cleaning_business_spreadsheetsample_datapy)
* [3. The work registers](#3-the-work-registers)
  * [`cleaning_business_spreadsheet/sheets/__init__.py`](#cleaning_business_spreadsheetsheets__init__py)
  * [`cleaning_business_spreadsheet/sheets/common.py`](#cleaning_business_spreadsheetsheetscommonpy)
  * [`cleaning_business_spreadsheet/sheets/lists.py`](#cleaning_business_spreadsheetsheetslistspy)
  * [`cleaning_business_spreadsheet/sheets/settings.py`](#cleaning_business_spreadsheetsheetssettingspy)
  * [`cleaning_business_spreadsheet/sheets/clients.py`](#cleaning_business_spreadsheetsheetsclientspy)
  * [`cleaning_business_spreadsheet/sheets/properties.py`](#cleaning_business_spreadsheetsheetspropertiespy)
  * [`cleaning_business_spreadsheet/sheets/leads.py`](#cleaning_business_spreadsheetsheetsleadspy)
  * [`cleaning_business_spreadsheet/sheets/services.py`](#cleaning_business_spreadsheetsheetsservicespy)
  * [`cleaning_business_spreadsheet/sheets/quotes.py`](#cleaning_business_spreadsheetsheetsquotespy)
  * [`cleaning_business_spreadsheet/sheets/quote_lines.py`](#cleaning_business_spreadsheetsheetsquote_linespy)
  * [`cleaning_business_spreadsheet/sheets/jobs.py`](#cleaning_business_spreadsheetsheetsjobspy)
  * [`cleaning_business_spreadsheet/sheets/job_tasks.py`](#cleaning_business_spreadsheetsheetsjob_taskspy)
  * [`cleaning_business_spreadsheet/sheets/checklists.py`](#cleaning_business_spreadsheetsheetschecklistspy)
* [4. The people and money registers](#4-the-people-and-money-registers)
  * [`cleaning_business_spreadsheet/sheets/employees.py`](#cleaning_business_spreadsheetsheetsemployeespy)
  * [`cleaning_business_spreadsheet/sheets/timesheets.py`](#cleaning_business_spreadsheetsheetstimesheetspy)
  * [`cleaning_business_spreadsheet/sheets/job_costing.py`](#cleaning_business_spreadsheetsheetsjob_costingpy)
  * [`cleaning_business_spreadsheet/sheets/invoices.py`](#cleaning_business_spreadsheetsheetsinvoicespy)
  * [`cleaning_business_spreadsheet/sheets/invoice_lines.py`](#cleaning_business_spreadsheetsheetsinvoice_linespy)
  * [`cleaning_business_spreadsheet/sheets/payments.py`](#cleaning_business_spreadsheetsheetspaymentspy)
  * [`cleaning_business_spreadsheet/sheets/supplies.py`](#cleaning_business_spreadsheetsheetssuppliespy)
  * [`cleaning_business_spreadsheet/sheets/supply_usage.py`](#cleaning_business_spreadsheetsheetssupply_usagepy)
  * [`cleaning_business_spreadsheet/sheets/equipment.py`](#cleaning_business_spreadsheetsheetsequipmentpy)
  * [`cleaning_business_spreadsheet/sheets/expenses.py`](#cleaning_business_spreadsheetsheetsexpensespy)
  * [`cleaning_business_spreadsheet/sheets/mileage.py`](#cleaning_business_spreadsheetsheetsmileagepy)
  * [`cleaning_business_spreadsheet/sheets/complaints.py`](#cleaning_business_spreadsheetsheetscomplaintspy)
  * [`cleaning_business_spreadsheet/sheets/incidents.py`](#cleaning_business_spreadsheetsheetsincidentspy)
  * [`cleaning_business_spreadsheet/sheets/safety.py`](#cleaning_business_spreadsheetsheetssafetypy)
* [5. The pages](#5-the-pages)
  * [`cleaning_business_spreadsheet/sheets/start_here.py`](#cleaning_business_spreadsheetsheetsstart_herepy)
  * [`cleaning_business_spreadsheet/sheets/dashboard.py`](#cleaning_business_spreadsheetsheetsdashboardpy)
  * [`cleaning_business_spreadsheet/sheets/quote_calc.py`](#cleaning_business_spreadsheetsheetsquote_calcpy)
  * [`cleaning_business_spreadsheet/sheets/calendar_page.py`](#cleaning_business_spreadsheetsheetscalendar_pagepy)
  * [`cleaning_business_spreadsheet/sheets/staff_schedule.py`](#cleaning_business_spreadsheetsheetsstaff_schedulepy)
  * [`cleaning_business_spreadsheet/sheets/retention.py`](#cleaning_business_spreadsheetsheetsretentionpy)
  * [`cleaning_business_spreadsheet/sheets/reports.py`](#cleaning_business_spreadsheetsheetsreportspy)
  * [`cleaning_business_spreadsheet/sheets/financial.py`](#cleaning_business_spreadsheetsheetsfinancialpy)
  * [`cleaning_business_spreadsheet/sheets/printables.py`](#cleaning_business_spreadsheetsheetsprintablespy)
  * [`cleaning_business_spreadsheet/sheets/help_tab.py`](#cleaning_business_spreadsheetsheetshelp_tabpy)
* [6. Quality checks (optional)](#6-quality-checks-optional)
  * [`tools/validate_cleaning_workbook.py`](#toolsvalidate_cleaning_workbookpy)
  * [`tools/check_cleaning_sample.py`](#toolscheck_cleaning_samplepy)
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

`build_cleaning_workbook.py` parses the command line into a `Config`, then
`cleaning_business_spreadsheet.build.build_workbook(path, cfg)` does six things in order:

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
6. **Activate the first tab** and close the file - there is no hidden sheet to hide in this one.

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

37 tabs in the order they appear in the finished workbook. "Rows" is the capacity of a register at the default `--rows` setting.

| # | Tab | Kind | Written by | Columns | Rows |
|---|-----|------|------------|--------:|-----:|
| 1 | **Start Here** | page | `cleaning_business_spreadsheet/sheets/start_here.py` | - | - |
| 2 | **Settings** | page | `cleaning_business_spreadsheet/sheets/settings.py` | - | - |
| 3 | **Dashboard** | page | `cleaning_business_spreadsheet/sheets/dashboard.py` | - | - |
| 4 | **Quote Calculator** | page | `cleaning_business_spreadsheet/sheets/quote_calc.py` | - | - |
| 5 | **Clients** | register | `cleaning_business_spreadsheet/sheets/clients.py` | 45 | 200 |
| 6 | **Properties** | register | `cleaning_business_spreadsheet/sheets/properties.py` | 34 | 250 |
| 7 | **Leads** | register | `cleaning_business_spreadsheet/sheets/leads.py` | 26 | 150 |
| 8 | **Services** | register | `cleaning_business_spreadsheet/sheets/services.py` | 29 | 60 |
| 9 | **Quotes** | register | `cleaning_business_spreadsheet/sheets/quotes.py` | 37 | 200 |
| 10 | **Quote Lines** | register | `cleaning_business_spreadsheet/sheets/quote_lines.py` | 21 | 600 |
| 11 | **Jobs** | register | `cleaning_business_spreadsheet/sheets/jobs.py` | 50 | 600 |
| 12 | **Job Tasks** | register | `cleaning_business_spreadsheet/sheets/job_tasks.py` | 20 | 900 |
| 13 | **Calendar** | page | `cleaning_business_spreadsheet/sheets/calendar_page.py` | - | - |
| 14 | **Checklists** | register | `cleaning_business_spreadsheet/sheets/checklists.py` | 24 | 400 |
| 15 | **Employees** | register | `cleaning_business_spreadsheet/sheets/employees.py` | 39 | 40 |
| 16 | **Staff Schedule** | page | `cleaning_business_spreadsheet/sheets/staff_schedule.py` | - | - |
| 17 | **Timesheets** | register | `cleaning_business_spreadsheet/sheets/timesheets.py` | 26 | 800 |
| 18 | **Job Costing** | register | `cleaning_business_spreadsheet/sheets/job_costing.py` | 28 | 600 |
| 19 | **Invoices** | register | `cleaning_business_spreadsheet/sheets/invoices.py` | 31 | 400 |
| 20 | **Invoice Lines** | register | `cleaning_business_spreadsheet/sheets/invoice_lines.py` | 18 | 800 |
| 21 | **Payments** | register | `cleaning_business_spreadsheet/sheets/payments.py` | 18 | 400 |
| 22 | **Supplies** | register | `cleaning_business_spreadsheet/sheets/supplies.py` | 26 | 150 |
| 23 | **Supply Usage** | register | `cleaning_business_spreadsheet/sheets/supply_usage.py` | 13 | 300 |
| 24 | **Equipment** | register | `cleaning_business_spreadsheet/sheets/equipment.py` | 26 | 80 |
| 25 | **Expenses** | register | `cleaning_business_spreadsheet/sheets/expenses.py` | 21 | 400 |
| 26 | **Mileage** | register | `cleaning_business_spreadsheet/sheets/mileage.py` | 20 | 400 |
| 27 | **Client Retention** | page | `cleaning_business_spreadsheet/sheets/retention.py` | - | - |
| 28 | **Complaints** | register | `cleaning_business_spreadsheet/sheets/complaints.py` | 25 | 80 |
| 29 | **Incidents** | register | `cleaning_business_spreadsheet/sheets/incidents.py` | 24 | 60 |
| 30 | **Safety** | register | `cleaning_business_spreadsheet/sheets/safety.py` | 18 | 80 |
| 31 | **Financial Plan** | page | `cleaning_business_spreadsheet/sheets/financial.py` | - | - |
| 32 | **Reports** | page | `cleaning_business_spreadsheet/sheets/reports.py` | - | - |
| 33 | **Print Quote** | page | `cleaning_business_spreadsheet/sheets/printables.py` | - | - |
| 34 | **Print Invoice** | page | `cleaning_business_spreadsheet/sheets/printables.py` | - | - |
| 35 | **Print Job Sheet** | page | `cleaning_business_spreadsheet/sheets/printables.py` | - | - |
| 36 | **Lists** | page | `cleaning_business_spreadsheet/sheets/lists.py` | - | - |
| 37 | **Help** | page | `cleaning_business_spreadsheet/sheets/help_tab.py` | - | - |

### Every file at a glance

| File | Lines | What it is |
|------|------:|------------|
| [`build_cleaning_workbook.py`](#build_cleaning_workbookpy) | 17 | Build the Cleaning Business CRM & Pricing Calculator workbook |
| [`cleaning_business_spreadsheet/__init__.py`](#cleaning_business_spreadsheet__init__py) | 19 | Cleaning Business CRM & Pricing Calculator |
| [`cleaning_business_spreadsheet/__main__.py`](#cleaning_business_spreadsheet__main__py) | 6 | ``python -m cleaning_business_spreadsheet`` |
| [`cleaning_business_spreadsheet/cli.py`](#cleaning_business_spreadsheetclipy) | 154 | Command line front end |
| [`cleaning_business_spreadsheet/config.py`](#cleaning_business_spreadsheetconfigpy) | 315 | Configuration for the Cleaning Business CRM & Pricing Calculator |
| [`cleaning_business_spreadsheet/layout.py`](#cleaning_business_spreadsheetlayoutpy) | 253 | Layout engine |
| [`cleaning_business_spreadsheet/formats.py`](#cleaning_business_spreadsheetformatspy) | 200 | Cell format factory |
| [`cleaning_business_spreadsheet/writer.py`](#cleaning_business_spreadsheetwriterpy) | 457 | Sheet writing helpers shared by every tab of the workbook |
| [`cleaning_business_spreadsheet/names.py`](#cleaning_business_spreadsheetnamespy) | 53 | Workbook level defined names that point at the register sheets |
| [`cleaning_business_spreadsheet/build.py`](#cleaning_business_spreadsheetbuildpy) | 163 | Workbook builder: puts every tab together in the right order |
| [`cleaning_business_spreadsheet/sample_data.py`](#cleaning_business_spreadsheetsample_datapy) | 1,855 | Demo data: a month in the life of Bright & Tidy Cleaning Co |
| [`cleaning_business_spreadsheet/sheets/__init__.py`](#cleaning_business_spreadsheetsheets__init__py) | 6 | One module per tab |
| [`cleaning_business_spreadsheet/sheets/common.py`](#cleaning_business_spreadsheetsheetscommonpy) | 267 | Small building blocks shared by the register sheets |
| [`cleaning_business_spreadsheet/sheets/lists.py`](#cleaning_business_spreadsheetsheetslistspy) | 386 | `Lists` - every drop down in the workbook, plus the pricing factor tables |
| [`cleaning_business_spreadsheet/sheets/settings.py`](#cleaning_business_spreadsheetsheetssettingspy) | 342 | The `Settings` tab - business profile, pricing, pay rates, targets, alerts |
| [`cleaning_business_spreadsheet/sheets/clients.py`](#cleaning_business_spreadsheetsheetsclientspy) | 255 | `Clients` - the customer database every other tab points at |
| [`cleaning_business_spreadsheet/sheets/properties.py`](#cleaning_business_spreadsheetsheetspropertiespy) | 217 | `Properties` - the places you actually clean |
| [`cleaning_business_spreadsheet/sheets/leads.py`](#cleaning_business_spreadsheetsheetsleadspy) | 148 | `Leads` - every enquiry, from first phone call to won or lost |
| [`cleaning_business_spreadsheet/sheets/services.py`](#cleaning_business_spreadsheetsheetsservicespy) | 182 | `Services` - the price list |
| [`cleaning_business_spreadsheet/sheets/quotes.py`](#cleaning_business_spreadsheetsheetsquotespy) | 185 | `Quotes` - one row per quote, totalled from its lines |
| [`cleaning_business_spreadsheet/sheets/quote_lines.py`](#cleaning_business_spreadsheetsheetsquote_linespy) | 140 | `Quote Lines` - what is actually on each quote |
| [`cleaning_business_spreadsheet/sheets/jobs.py`](#cleaning_business_spreadsheetsheetsjobspy) | 315 | `Jobs` - the job scheduler, and the spine of the whole workbook |
| [`cleaning_business_spreadsheet/sheets/job_tasks.py`](#cleaning_business_spreadsheetsheetsjob_taskspy) | 120 | `Job Tasks` - the checklist as it was actually done, job by job |
| [`cleaning_business_spreadsheet/sheets/checklists.py`](#cleaning_business_spreadsheetsheetschecklistspy) | 153 | `Checklists` - the master task lists behind every clean |
| [`cleaning_business_spreadsheet/sheets/employees.py`](#cleaning_business_spreadsheetsheetsemployeespy) | 208 | `Employees` - the team |
| [`cleaning_business_spreadsheet/sheets/timesheets.py`](#cleaning_business_spreadsheetsheetstimesheetspy) | 144 | `Timesheets` - hours worked, what they cost and whether they were billed |
| [`cleaning_business_spreadsheet/sheets/job_costing.py`](#cleaning_business_spreadsheetsheetsjob_costingpy) | 166 | `Job Costing` - what every job really made, line by line |
| [`cleaning_business_spreadsheet/sheets/invoices.py`](#cleaning_business_spreadsheetsheetsinvoicespy) | 175 | `Invoices` - what you have billed and what is still owed |
| [`cleaning_business_spreadsheet/sheets/invoice_lines.py`](#cleaning_business_spreadsheetsheetsinvoice_linespy) | 115 | `Invoice Lines` - the detail on each invoice |
| [`cleaning_business_spreadsheet/sheets/payments.py`](#cleaning_business_spreadsheetsheetspaymentspy) | 116 | `Payments` - the money that actually arrived |
| [`cleaning_business_spreadsheet/sheets/supplies.py`](#cleaning_business_spreadsheetsheetssuppliespy) | 158 | `Supplies` - the cupboard, and what it is costing you |
| [`cleaning_business_spreadsheet/sheets/supply_usage.py`](#cleaning_business_spreadsheetsheetssupply_usagepy) | 85 | `Supply Usage` - what got used, where |
| [`cleaning_business_spreadsheet/sheets/equipment.py`](#cleaning_business_spreadsheetsheetsequipmentpy) | 138 | `Equipment` - vacuums, machines, vans and ladders |
| [`cleaning_business_spreadsheet/sheets/expenses.py`](#cleaning_business_spreadsheetsheetsexpensespy) | 116 | `Expenses` - everything the business spends |
| [`cleaning_business_spreadsheet/sheets/mileage.py`](#cleaning_business_spreadsheetsheetsmileagepy) | 108 | `Mileage & Travel` - the miles between the jobs |
| [`cleaning_business_spreadsheet/sheets/complaints.py`](#cleaning_business_spreadsheetsheetscomplaintspy) | 127 | `Complaints` - what went wrong and what you did about it |
| [`cleaning_business_spreadsheet/sheets/incidents.py`](#cleaning_business_spreadsheetsheetsincidentspy) | 96 | `Incidents` - accidents, near misses and damage |
| [`cleaning_business_spreadsheet/sheets/safety.py`](#cleaning_business_spreadsheetsheetssafetypy) | 113 | `Safety & Compliance` - the paperwork that has to be in date |
| [`cleaning_business_spreadsheet/sheets/start_here.py`](#cleaning_business_spreadsheetsheetsstart_herepy) | 294 | `Start Here` - the first thing anyone sees when they open the file |
| [`cleaning_business_spreadsheet/sheets/dashboard.py`](#cleaning_business_spreadsheetsheetsdashboardpy) | 375 | `Dashboard` - the one page that answers 'how is the business doing?' Twelve numbers across the top, a list... |
| [`cleaning_business_spreadsheet/sheets/quote_calc.py`](#cleaning_business_spreadsheetsheetsquote_calcpy) | 424 | `Quote Calculator` - price a clean in thirty seconds, honestly |
| [`cleaning_business_spreadsheet/sheets/calendar_page.py`](#cleaning_business_spreadsheetsheetscalendar_pagepy) | 176 | `Calendar` - the month at a glance, and the next jobs in order |
| [`cleaning_business_spreadsheet/sheets/staff_schedule.py`](#cleaning_business_spreadsheetsheetsstaff_schedulepy) | 305 | `Staff Schedule` - the week's rota, and who has room for one more |
| [`cleaning_business_spreadsheet/sheets/retention.py`](#cleaning_business_spreadsheetsheetsretentionpy) | 233 | `Client Retention` - keeping the clients you already have |
| [`cleaning_business_spreadsheet/sheets/reports.py`](#cleaning_business_spreadsheetsheetsreportspy) | 379 | `Reports & Analytics` - the questions a cleaning business should ask |
| [`cleaning_business_spreadsheet/sheets/financial.py`](#cleaning_business_spreadsheetsheetsfinancialpy) | 341 | `Financial Plan` - targets, break-even and what if you grew |
| [`cleaning_business_spreadsheet/sheets/printables.py`](#cleaning_business_spreadsheetsheetsprintablespy) | 538 | The three printable forms |
| [`cleaning_business_spreadsheet/sheets/help_tab.py`](#cleaning_business_spreadsheetsheetshelp_tabpy) | 255 | `Help` - troubleshooting, a glossary and a map of every feature |
| [`tools/validate_cleaning_workbook.py`](#toolsvalidate_cleaning_workbookpy) | 537 | Static checks on a generated Cleaning Business workbook |
| [`tools/check_cleaning_sample.py`](#toolscheck_cleaning_samplepy) | 309 | Cross-check the example records in a generated cleaning workbook |
| [`tools/recalc_check.py`](#toolsrecalc_checkpy) | 216 | Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error... |
| [`tools/export_source_md.py`](#toolsexport_source_mdpy) | 1,076 | Write every source file of a generator into one markdown document |
| [`tools/make_download_zip.py`](#toolsmake_download_zippy) | 286 | Pack a generator into a single zip people can download and run |

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_cleaning_workbook.py
```

The workbook lands in `output/Cleaning Business CRM & Pricing Calculator.xlsx`.

### `build_cleaning_workbook.py`

*17 lines*

> Build the Cleaning Business CRM & Pricing Calculator workbook.
>
>     python3 build_cleaning_workbook.py
>     python3 build_cleaning_workbook.py -o "output/My Cleaning CRM.xlsx" --blank
>     python3 build_cleaning_workbook.py --currency "£" --rate 38 --tax 20
>
> Run `--help` for every option. The generator itself lives in the
> `cleaning_business_spreadsheet` package next to this file.

````python
#!/usr/bin/env python3
"""Build the Cleaning Business CRM & Pricing Calculator workbook.

    python3 build_cleaning_workbook.py
    python3 build_cleaning_workbook.py -o "output/My Cleaning CRM.xlsx" --blank
    python3 build_cleaning_workbook.py --currency "£" --rate 38 --tax 20

Run `--help` for every option. The generator itself lives in the
`cleaning_business_spreadsheet` package next to this file.
"""

import sys

from cleaning_business_spreadsheet.cli import main

if __name__ == "__main__":
    sys.exit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer, the cross sheet names and the demo cleaning company. These are the parts every tab leans on.

### `cleaning_business_spreadsheet/__init__.py`

*19 lines*

> Cleaning Business CRM & Pricing Calculator.
>
> Build the workbook with::
>
>     python -m cleaning_business_spreadsheet -o "output/Cleaning CRM.xlsx"
>
> or from Python::
>
>     from cleaning_business_spreadsheet import build_workbook
>     build_workbook("Cleaning Business CRM.xlsx")

````python
"""Cleaning Business CRM & Pricing Calculator.

Build the workbook with::

    python -m cleaning_business_spreadsheet -o "output/Cleaning CRM.xlsx"

or from Python::

    from cleaning_business_spreadsheet import build_workbook
    build_workbook("Cleaning Business CRM.xlsx")
"""

from .build import build_workbook
from .config import (AUTHOR, PRODUCT_LISTING_TITLE, PRODUCT_NAME,
                     PRODUCT_SUBTITLE, PRODUCT_VERSION, Config)

__all__ = ["build_workbook", "Config", "AUTHOR", "PRODUCT_NAME",
           "PRODUCT_SUBTITLE", "PRODUCT_LISTING_TITLE", "PRODUCT_VERSION"]
__version__ = PRODUCT_VERSION
````

### `cleaning_business_spreadsheet/__main__.py`

*6 lines*

> ``python -m cleaning_business_spreadsheet``.

````python
"""``python -m cleaning_business_spreadsheet``."""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
````

### `cleaning_business_spreadsheet/cli.py`

*154 lines · 3 functions*

> Command line front end.
>
> python -m cleaning_business_spreadsheet
> python -m cleaning_business_spreadsheet -o "output/Cleaning.xlsx" --blank
> python -m cleaning_business_spreadsheet --currency "£" --rate 38 --rows 300

| Defines | What it does |
|---------|--------------|
| `build_parser()` |  |
| `config_from_args(args)` |  |
| `main(argv)` |  |

**Module constants:** `DEFAULT_OUTPUT` (os.path.join(...)).

````python
"""Command line front end.

    python -m cleaning_business_spreadsheet
    python -m cleaning_business_spreadsheet -o "output/Cleaning.xlsx" --blank
    python -m cleaning_business_spreadsheet --currency "£" --rate 38 --rows 300
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import AUTHOR, PRODUCT_LISTING_TITLE, PRODUCT_NAME, Config

DEFAULT_OUTPUT = os.path.join(
    "output", "Cleaning Business CRM & Pricing Calculator.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="cleaning_business_spreadsheet",
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
    parser.add_argument("--owner", default=None,
                        help="pre-fill the owner / manager name")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="the year the planner and reports open on")
    parser.add_argument("--week-start", choices=["Monday", "Sunday"],
                        default=None, help="first day of the week")
    parser.add_argument("--rate", type=float, default=None,
                        help="default hourly rate charged to the client")
    parser.add_argument("--pay", type=float, default=None,
                        help="default hourly pay rate of a cleaner")
    parser.add_argument("--minimum", type=float, default=None,
                        help="minimum charge for any job")
    parser.add_argument("--tax", type=float, default=None,
                        help="sales tax / VAT rate as a percentage, "
                             "e.g. 20 for 20%%")
    parser.add_argument("--terms", type=int, default=None,
                        help="payment terms in days")
    parser.add_argument("--units", choices=["sq ft", "sq m"], default=None,
                        help="the area unit used for pricing")
    parser.add_argument("--distance", choices=["miles", "km"], default=None,
                        help="the distance unit used for mileage")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the default row "
                             "counts are multiplied to fit this many "
                             "job rows)")
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
    if args.week_start:
        kwargs["week_start"] = args.week_start
    if args.year:
        kwargs["plan_year"] = args.year
    if args.rate:
        kwargs["hourly_rate"] = args.rate
    if args.pay:
        kwargs["pay_rate"] = args.pay
    if args.minimum:
        kwargs["minimum_charge"] = args.minimum
    if args.tax is not None:
        kwargs["tax_rate"] = round(args.tax / 100.0, 4)
    if args.terms:
        kwargs["payment_terms_days"] = args.terms
    if args.units:
        kwargs["area_unit"] = args.units
    if args.distance:
        kwargs["distance_unit"] = args.distance
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.business:
        cfg.business.name = args.business
    if args.owner:
        cfg.business.owner = args.owner
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
    print(f"  business      : {cfg.business.name}")
    print(f"  pricing       : {cfg.currency_symbol}{cfg.hourly_rate:,.2f} an "
          f"hour, minimum {cfg.currency_symbol}{cfg.minimum_charge:,.2f}")
    print(f"  example data  : "
          f"{'no' if not cfg.include_sample_data else 'yes'}")
    print(f"  protection    : "
          f"{'off' if not cfg.protect_sheets else 'on, password ' + cfg.password}")
    print(f"  written to    : {os.path.abspath(path)}  ({size:,.0f} KB)")
    return 0


if __name__ == "__main__":        # pragma: no cover
    raise SystemExit(main())
````

### `cleaning_business_spreadsheet/config.py`

*315 lines · 3 classes*

> Configuration for the Cleaning Business CRM & Pricing Calculator.
>
> Everything a shop owner may want to re-brand lives here, so the workbook
> can be regenerated for a different company, currency, country or pricing
> model without touching any of the sheet code.
>
> Author of the generated workbooks: Novality Store.

| Defines | What it does |
|---------|--------------|
| `class Palette` |  |
| `class BusinessDefaults` | Sample values written into the Business Profile (Settings) sheet |
| `class Config` | Everything the generator needs to know to write one workbook |

**Module constants:** `AUTHOR` ('Novality Store'), `PRODUCT_NAME` ('Cleaning Business CRM & Pricing Calculator'), `PRODUCT_SUBTITLE` ('Client CRM, Job Scheduler, Quote Calculator, Employee Pl...), `PRODUCT_VERSION` ('1.0.0'), `PRODUCT_LISTING_TITLE` ('Cleaning Business Spreadsheet, Cleaning Service CRM, Job...), `FORMULA_PASSWORD` ('premium'), `JOB_STATUSES` (8 items), `DONE_STATUSES` (2 items), `QUOTE_STATUSES` (6 items), `INVOICE_STATUSES` (6 items), `FREQUENCIES` (7 items), `FREQUENCY_FACTORS` (7 entries), `CONDITION_FACTORS` (5 entries), `DEFAULT_ROWS` (25 entries).

````python
"""
Configuration for the Cleaning Business CRM & Pricing Calculator.

Everything a shop owner may want to re-brand lives here, so the workbook
can be regenerated for a different company, currency, country or pricing
model without touching any of the sheet code.

Author of the generated workbooks: Novality Store.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import date


# ---------------------------------------------------------------------------
# Product / branding constants
# ---------------------------------------------------------------------------

AUTHOR = "Novality Store"
PRODUCT_NAME = "Cleaning Business CRM & Pricing Calculator"
PRODUCT_SUBTITLE = (
    "Client CRM, Job Scheduler, Quote Calculator, Employee Planner, "
    "Invoices, Expenses & Profit Dashboard"
)
PRODUCT_VERSION = "1.0.0"

#: Title suggested for the shop listing.
PRODUCT_LISTING_TITLE = (
    "Cleaning Business Spreadsheet, Cleaning Service CRM, Job Scheduler, "
    "Quote Calculator, Employee Planner, Invoices, Expenses and Profit "
    "Dashboard"
)

#: Password that protects every sheet so formula cells cannot be edited by
#: accident.  Input cells stay unlocked, formula cells are locked.
FORMULA_PASSWORD = "premium"

#: The job statuses that drive the scheduler, the dashboard and the alerts.
JOB_STATUSES = ["Enquiry", "Booked", "Confirmed", "In progress",
                "Completed", "Invoiced", "Cancelled", "No access"]

#: Statuses that mean the work actually happened.
DONE_STATUSES = ["Completed", "Invoiced"]

#: Quote statuses, in the order a quote moves through them.
QUOTE_STATUSES = ["Draft", "Sent", "Followed up", "Accepted", "Declined",
                  "Expired"]

#: Invoice statuses.
INVOICE_STATUSES = ["Draft", "Sent", "Part paid", "Paid", "Overdue",
                    "Written off"]

#: How often a recurring clean can repeat.
FREQUENCIES = ["One-off", "Weekly", "Fortnightly", "Every 3 weeks",
               "Monthly", "Quarterly", "Twice a week"]

#: Multiplier applied to the standard price for each frequency - regular
#: work is cheaper per visit because the place never gets badly dirty.
FREQUENCY_FACTORS = {
    "One-off": 1.00,
    "Weekly": 0.85,
    "Fortnightly": 0.90,
    "Every 3 weeks": 0.95,
    "Monthly": 1.00,
    "Quarterly": 1.05,
    "Twice a week": 0.80,
}

#: How dirty the place is, and what that does to the hours.
CONDITION_FACTORS = {
    "Very clean - upkeep only": 0.85,
    "Normal": 1.00,
    "Needs attention": 1.20,
    "Heavily soiled": 1.45,
    "Extreme / hoarding": 1.80,
}


# ---------------------------------------------------------------------------
# Colours - a clean, professional palette: deep teal, fresh aqua, citrus
# ---------------------------------------------------------------------------

class Palette:
    ink = "#102A30"           # darkest teal - titles
    primary = "#14616B"       # headers
    primary_mid = "#2F8F9D"   # section bands
    primary_soft = "#DCEEF0"  # light band / alternating rows
    accent = "#E8A33D"        # citrus accent
    accent_soft = "#FDF0D9"
    input_bg = "#FFFFFF"      # editable cell background (clean white)
    input_border = "#D5E3E5"
    calc_bg = "#EDF4F5"       # locked formula cell background (pale teal)
    calc_border = "#CFE0E3"
    white = "#FFFFFF"
    grey_text = "#5A6A6D"
    line = "#BFD6DA"
    good_bg = "#D7F2E3"
    good_text = "#11684A"
    warn_bg = "#FFF0CE"
    warn_text = "#8A5A00"
    bad_bg = "#FBD9DD"
    bad_text = "#9C1C2E"
    info_bg = "#DDE9FA"
    info_text = "#27467F"
    note_bg = "#F5FAFB"


# ---------------------------------------------------------------------------
# Row capacity of every register (how many pre-formatted / pre-formulated rows)
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Settings": 0,
    "Lists": 40,
    "Clients": 200,
    "Properties": 250,
    "Leads": 150,
    "Services": 60,
    "Quotes": 200,
    "Quote Lines": 600,
    "Jobs": 600,
    "Job Tasks": 900,
    "Checklists": 400,
    "Employees": 40,
    "Timesheets": 800,
    "Job Costing": 600,
    "Invoices": 400,
    "Invoice Lines": 800,
    "Payments": 400,
    "Supplies": 150,
    "Supply Usage": 300,
    "Equipment": 80,
    "Expenses": 400,
    "Mileage": 400,
    "Complaints": 80,
    "Safety": 80,
    "Incidents": 60,
}


@dataclass
class BusinessDefaults:
    """Sample values written into the Business Profile (Settings) sheet."""

    name: str = "Bright & Tidy Cleaning Co."
    owner: str = "Dana Whitfield"
    email: str = "hello@brightandtidy.co"
    phone: str = "(503) 555 0142"
    address: str = "18 Juniper Way, Portland, OR 97214"
    area: str = "Portland metro - 20 mile radius"
    website: str = "brightandtidy.co"
    tagline: str = "Reliable domestic and commercial cleaning since 2019"
    reg_number: str = "BT-CLEAN-2019-0146"
    insurance: str = "Hartfield Commercial - policy HCX-884213"


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
    area_unit: str = "sq ft"             # sq ft | sq m
    distance_unit: str = "miles"         # miles | km

    # Business profile pre-filled on the Settings tab -----------------------
    business: BusinessDefaults = field(default_factory=BusinessDefaults)
    plan_year: int = 0                   # 0 = the year of `today`
    week_start: str = "Monday"

    # Pricing ---------------------------------------------------------------
    hourly_rate: float = 45.00           # what the client pays an hour
    minimum_charge: float = 90.00        # no job leaves for less than this
    rate_per_area: float = 0.09          # per sq ft of a standard clean
    rate_per_bedroom: float = 25.00
    rate_per_bathroom: float = 30.00
    rate_per_room: float = 18.00
    callout_fee: float = 15.00
    deep_clean_uplift: float = 0.45      # +45% on a deep clean
    move_clean_uplift: float = 0.35
    pet_surcharge: float = 20.00
    supplies_per_hour: float = 2.50      # what the chemicals cost you
    target_margin: float = 0.45          # the margin a quote aims for
    margin_floor: float = 0.25           # below this the quote is flagged
    deposit_percent: float = 0.25

    # Staff -----------------------------------------------------------------
    pay_rate: float = 18.00              # what a cleaner costs an hour
    overtime_multiplier: float = 1.5
    employer_oncost: float = 0.18        # tax, insurance, holiday pay
    cleaners_per_job: int = 2
    productive_hours: float = 6.0        # billable hours in a working day
    travel_rate: float = 0.67            # reimbursement a mile
    travel_minutes: int = 20             # allowance between jobs

    # Billing ---------------------------------------------------------------
    tax_rate: float = 0.00               # sales tax / VAT on invoices
    tax_label: str = "Sales tax"
    payment_terms_days: int = 14
    late_fee_percent: float = 0.05
    quote_valid_days: int = 30

    # Targets the dashboard measures against --------------------------------
    monthly_revenue_target: float = 26000.0
    annual_revenue_target: float = 312000.0
    monthly_expense_budget: float = 15000.0
    # the example company is a smaller operation than the defaults
    # assume, so the demo file carries targets it can be measured
    # against - a dashboard showing 16% of target teaches nobody
    # anything
    sample_monthly_revenue_target: float = 6000.0
    sample_annual_revenue_target: float = 72000.0
    sample_monthly_expense_budget: float = 4000.0
    utilisation_target: float = 0.75     # share of paid hours that are billed
    retention_target: float = 0.85
    satisfaction_target: float = 4.5     # out of 5
    quote_win_target: float = 0.50

    # Reminder windows ------------------------------------------------------
    quote_followup_days: int = 3
    overdue_grace_days: int = 0
    reminder_days: int = 2               # how early a job is flagged
    inactive_client_days: int = 60
    low_stock_days: int = 14
    maintenance_days: int = 90           # service interval for equipment
    certificate_warning_days: int = 45
    renewal_warning_days: int = 60

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
    def expense_budget_month(self) -> float:
        return (self.sample_monthly_expense_budget
                if self.include_sample_data else self.monthly_expense_budget)

    @property
    def imperial(self) -> bool:
        return self.area_unit.replace(" ", "").lower().startswith("sqf")

    def scale_rows(self, job_rows: int) -> None:
        """Scale every register in place so Jobs holds N rows."""
        base = DEFAULT_ROWS["Jobs"]
        factor = max(0.02, float(job_rows) / float(base))
        self.rows = {k: (max(5, int(round(v * factor))) if v else v)
                     for k, v in self.rows.items()}

    def scaled(self, factor: float) -> "Config":
        """Return a copy with every register scaled."""
        rows = {k: max(5, int(round(v * factor))) if v else v
                for k, v in self.rows.items()}
        return replace(self, rows=rows)

    # Number format strings derived from the chosen currency ------------------
    @property
    def money_format(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0.00;[Red]-{s}#,##0.00'

    @property
    def money_format_0(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0;[Red]-{s}#,##0'

    @property
    def cost_per_hour(self) -> float:
        """What one cleaner-hour really costs once on-costs are added."""
        return round(self.pay_rate * (1 + self.employer_oncost), 2)

    @property
    def week_start_date(self) -> date:
        """The Monday (or Sunday) of the week `today` falls in."""
        day = self.today
        offset = day.weekday() if self.week_start == "Monday" \
            else (day.weekday() + 1) % 7
        return date.fromordinal(day.toordinal() - offset)
````

### `cleaning_business_spreadsheet/layout.py`

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

### `cleaning_business_spreadsheet/formats.py`

*200 lines · 1 class*

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

### `cleaning_business_spreadsheet/writer.py`

*457 lines · 1 class · 8 functions*

> Sheet writing helpers shared by every tab of the workbook.
>
> :class:`SheetWriter` wraps an XlsxWriter worksheet and adds the house style:
> title band, navigation strip, section headers, KPI tiles, label/value settings
> rows and - most importantly - :meth:`SheetWriter.grid`, which turns a
> :class:`~cleaning_business_spreadsheet.layout.SheetSpec` into a finished register
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
:class:`~cleaning_business_spreadsheet.layout.SheetSpec` into a finished register
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
    ("Jobs", "Jobs"),
    ("Calendar", "Calendar"),
    ("Quotes", "Quotes"),
    ("Invoices", "Invoices"),
    ("Staff", "Employees"),
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

### `cleaning_business_spreadsheet/names.py`

*53 lines · 1 function*

> Workbook level defined names that point at the register sheets.
>
> The drop down lists on the `Lists` tab define their own names. These are
> the ones that have to follow live data - client IDs, property references,
> the people on the payroll - so they are built from the registry once
> every register exists. Each one is an ``OFFSET`` range that grows and
> shrinks with the number of filled rows, which is why a new client appears
> in the job scheduler's drop down the moment you type their name.

| Defines | What it does |
|---------|--------------|
| `define_cross_sheet_names(book, reg)` | Create the live lookup names |

**Module constants:** `CROSS_SHEET_NAMES` (20 entries).

````python
"""Workbook level defined names that point at the register sheets.

The drop down lists on the `Lists` tab define their own names. These are
the ones that have to follow live data - client IDs, property references,
the people on the payroll - so they are built from the registry once
every register exists. Each one is an ``OFFSET`` range that grows and
shrinks with the number of filled rows, which is why a new client appears
in the job scheduler's drop down the moment you type their name.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # the spine of the whole workbook
    "ClientIDs": ("Clients", "id"),
    "ClientNames": ("Clients", "name"),
    "PropertyIDs": ("Properties", "id"),
    "PropertyNames": ("Properties", "name"),
    "ServiceNames": ("Services", "name"),
    "ServiceIDs": ("Services", "id"),
    "JobIDs": ("Jobs", "id"),
    "QuoteIDs": ("Quotes", "id"),
    "InvoiceIDs": ("Invoices", "id"),
    "EmployeeNames": ("Employees", "name"),
    "EmployeeIDs": ("Employees", "id"),
    "SupplyNames": ("Supplies", "name"),
    "EquipmentNames": ("Equipment", "name"),
    "LeadIDs": ("Leads", "id"),
    # columns the dashboard, the calendar and the reports scan
    "JobDates": ("Jobs", "date"),
    "JobClients": ("Jobs", "client"),
    "JobStatuses": ("Jobs", "status"),
    "JobMonths": ("Jobs", "month"),
    "InvoiceMonths": ("Invoices", "month"),
    "ExpenseMonths": ("Expenses", "month"),
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

### `cleaning_business_spreadsheet/build.py`

*163 lines · 3 functions*

> Workbook builder: puts every tab together in the right order.
>
> Typical use::
>
>     from cleaning_business_spreadsheet import build_workbook
>     build_workbook("Cleaning Business CRM & Pricing Calculator.xlsx")

| Defines | What it does |
|---------|--------------|
| `build_register(book, fmts, cfg, reg, module)` | Create one register sheet from its :class:`SheetSpec` |
| `build_workbook(path, cfg)` | Write the whole product to *path* and return the path |
| `_set_properties(book, cfg)` |  |

**Module constants:** `REGISTER_MODULES` (23 items), `SHEET_ORDER` (37 items), `PAGE_BUILDERS` (14 entries), `REGISTER_BY_NAME` ({module.NAME: module for module in REGISTER_MODULES}).

````python
"""Workbook builder: puts every tab together in the right order.

Typical use::

    from cleaning_business_spreadsheet import build_workbook
    build_workbook("Cleaning Business CRM & Pricing Calculator.xlsx")
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
from .sheets import (calendar_page, checklists, clients, complaints,
                     dashboard, employees, equipment, expenses,
                     financial, help_tab, incidents, invoice_lines,
                     invoices, job_costing, job_tasks, jobs, leads,
                     lists, mileage, payments, printables, properties,
                     quote_calc, quote_lines, quotes, reports,
                     retention, safety, services, settings,
                     staff_schedule, start_here, supplies, supply_usage,
                     timesheets)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    clients, properties, leads, services, quotes, quote_lines, jobs,
    job_tasks, checklists, employees, timesheets, job_costing, invoices,
    invoice_lines, payments, supplies, supply_usage, equipment, expenses,
    mileage, complaints, incidents, safety,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Start Here", "Settings", "Dashboard",
    "Quote Calculator",
    "Clients", "Properties", "Leads", "Services",
    "Quotes", "Quote Lines", "Jobs", "Job Tasks", "Calendar",
    "Checklists",
    "Employees", "Staff Schedule", "Timesheets",
    "Job Costing",
    "Invoices", "Invoice Lines", "Payments",
    "Supplies", "Supply Usage", "Equipment", "Expenses", "Mileage",
    "Client Retention", "Complaints", "Incidents", "Safety",
    "Financial Plan", "Reports",
    "Print Quote", "Print Invoice", "Print Job Sheet",
    "Lists", "Help",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Start Here": start_here.build,
    "Settings": settings.build,
    "Dashboard": dashboard.build,
    "Quote Calculator": quote_calc.build,
    "Calendar": calendar_page.build,
    "Staff Schedule": staff_schedule.build,
    "Client Retention": retention.build,
    "Financial Plan": financial.build,
    "Reports": reports.build,
    "Print Quote": printables.build_quote,
    "Print Invoice": printables.build_invoice,
    "Print Job Sheet": printables.build_job,
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


def build_workbook(path: str = "Cleaning Business CRM.xlsx",
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

    # 3. Live lookup names (client IDs, property names, who works here).
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
        "subject": "Cleaning business CRM, scheduling, quoting and "
                   "accounts",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("cleaning business spreadsheet, cleaning service "
                     "CRM, job scheduler, quote calculator, cleaning "
                     "price calculator, employee planner, invoice "
                     "tracker, expenses, profit dashboard, housekeeping "
                     "business, spreadsheet template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

### `cleaning_business_spreadsheet/sample_data.py`

*1,855 lines · 1 class · 11 functions*

> Demo data: a month in the life of Bright & Tidy Cleaning Co.
>
> A two-van, eight-cleaner operation in Portland with fifteen clients, a
> couple of commercial contracts, an Airbnb changeover round, one awkward
> end-of-tenancy job and a complaint that had to be put right. Every tab
> opens with believable numbers so the dashboard, the quote calculator and
> the reports mean something the moment the file is opened - and `--blank`
> strips the lot out for your own data.
>
> Everything is dated relative to the build date, so the workbook never
> looks stale.

| Defines | What it does |
|---------|--------------|
| `d(base, offset)` |  |
| `week_monday(day)` |  |
| `cid(n)` |  |
| `pid(n)` |  |
| `jid(n)` |  |
| `qid(n)` |  |
| `iid(n)` |  |
| `class Sample` | fields: clients, properties, leads, services, quotes, quote_lines, jobs, job_tasks |
| `get(cfg)` | Build (and cache) the demo data for this configuration |
| `_month_key(day)` |  |
| `_build(cfg)` |  |
| `_job_price(cfg, job)` | Mirror the Jobs price formula closely enough for the demo totals |

**Module constants:** `CLIENTS` (15 items), `PROPERTIES` (16 items), `SERVICES` (21 items), `EMPLOYEES` (8 items), `PROPERTY_EXTRA` (16 entries), `ENTRY_NOTES` (11 entries), `CHECKLIST_TASKS` (7 entries), `SUPPLIES` (16 items), `EQUIPMENT` (12 items), `SAFETY_DOCS` (16 items), `LEADS` (10 items), `QUOTES` (8 items), `ROUND` (16 items), `ROUND_STOPS` (1 entries) and 6 more.

````python
"""Demo data: a month in the life of Bright & Tidy Cleaning Co.

A two-van, eight-cleaner operation in Portland with fifteen clients, a
couple of commercial contracts, an Airbnb changeover round, one awkward
end-of-tenancy job and a complaint that had to be put right. Every tab
opens with believable numbers so the dashboard, the quote calculator and
the reports mean something the moment the file is opened - and `--blank`
strips the lot out for your own data.

Everything is dated relative to the build date, so the workbook never
looks stale.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .config import Config


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def week_monday(day: _dt.date) -> _dt.date:
    return day - _dt.timedelta(days=day.weekday())


def cid(n: int) -> str:
    return f"CL-{n:04d}"


def pid(n: int) -> str:
    return f"PR-{n:04d}"


def jid(n: int) -> str:
    return f"JOB-{n:04d}"


def qid(n: int) -> str:
    return f"Q-{n:04d}"


def iid(n: int) -> str:
    return f"INV-{n:04d}"


@dataclass
class Sample:
    clients: List[Dict[str, Any]] = field(default_factory=list)
    properties: List[Dict[str, Any]] = field(default_factory=list)
    leads: List[Dict[str, Any]] = field(default_factory=list)
    services: List[Dict[str, Any]] = field(default_factory=list)
    quotes: List[Dict[str, Any]] = field(default_factory=list)
    quote_lines: List[Dict[str, Any]] = field(default_factory=list)
    jobs: List[Dict[str, Any]] = field(default_factory=list)
    job_tasks: List[Dict[str, Any]] = field(default_factory=list)
    checklists: List[Dict[str, Any]] = field(default_factory=list)
    employees: List[Dict[str, Any]] = field(default_factory=list)
    timesheets: List[Dict[str, Any]] = field(default_factory=list)
    invoices: List[Dict[str, Any]] = field(default_factory=list)
    invoice_lines: List[Dict[str, Any]] = field(default_factory=list)
    payments: List[Dict[str, Any]] = field(default_factory=list)
    supplies: List[Dict[str, Any]] = field(default_factory=list)
    supply_usage: List[Dict[str, Any]] = field(default_factory=list)
    equipment: List[Dict[str, Any]] = field(default_factory=list)
    expenses: List[Dict[str, Any]] = field(default_factory=list)
    mileage: List[Dict[str, Any]] = field(default_factory=list)
    complaints: List[Dict[str, Any]] = field(default_factory=list)
    incidents: List[Dict[str, Any]] = field(default_factory=list)
    safety: List[Dict[str, Any]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# The client book
# ---------------------------------------------------------------------------
# name, type, contact, phone, email, address, status, source, since offset,
# frequency, day, slot, cleaner, rate, terms, contract, vip, notes
CLIENTS = [
    ("Harper Residence", "Residential", "Megan Harper", "(503) 555 0188",
     "megan.harper@example.com", "42 Larch Street, Portland OR 97215",
     "Regular", "Google search", -760, "Weekly", "Tuesday",
     "Morning (9-12)", "Marta Kovacs", 45, 7, "Rolling monthly", "No",
     "Two cats. Hates the smell of bleach - use the eco spray."),
    ("Okafor Family", "Residential", "Tobi Okafor", "(503) 555 0231",
     "tobi.okafor@example.com", "9 Elmwood Court, Portland OR 97212",
     "Regular", "Referral - client", -540, "Fortnightly", "Thursday",
     "Afternoon (2-5)", "Jade Ellis", 45, 7, "Rolling monthly", "No",
     "Twins under three - toys everywhere, allow extra tidy time."),
    ("Riverbend Dental", "Medical & dental", "Dr. Alan Reyes",
     "(503) 555 0410", "practice@riverbenddental.example",
     "118 Riverbend Plaza, Suite 3, Portland OR 97209", "Active",
     "Directory listing", -430, "Weekly", "Friday", "Late (5-8)",
     "Tomas Ruiz", 52, 14, "12 months", "Yes",
     "Clinical waste stays with the practice. Sign in at reception."),
    ("Cedar & Co Accountants", "Commercial office", "Priya Shah",
     "(503) 555 0366", "priya@cedarandco.example",
     "77 Cedar Avenue, Floor 2, Portland OR 97204", "Regular",
     "Referral - client", -620, "Twice a week", "Monday", "Late (5-8)",
     "Tomas Ruiz", 50, 14, "12 months", "Yes",
     "Out of hours only. Alarm code is on the Properties tab."),
    ("Lakeview Lettings", "Letting agent", "Sam Whitlock",
     "(503) 555 0177", "sam@lakeviewlettings.example",
     "3 Lakeview Road, Portland OR 97202", "Active", "Letting agent",
     "", "One-off", "Flexible", "Flexible", "Priya Raman", 48, 21,
     "No contract", "No",
     "End of tenancy work. Always wants photos and a receipt."),
    ("The Maple Cafe", "Restaurant & cafe", "Gio Bellini",
     "(503) 555 0145", "gio@maplecafe.example",
     "210 Maple Street, Portland OR 97214", "Active", "Word of mouth",
     -300, "Weekly", "Sunday", "Early (7-9)", "Amara Diallo", 55, 7,
     "6 months", "No",
     "Before opening. Kitchen deep clean quarterly, billed separately."),
    ("Bright Start Nursery", "School & nursery", "Carla Jensen",
     "(503) 555 0122", "carla@brightstart.example",
     "15 Juniper Lane, Portland OR 97213", "Regular", "Google search",
     -410, "Weekly", "Friday", "Late (5-8)", "Amara Diallo", 50, 14,
     "12 months", "Yes",
     "Background checks required for everyone on site. No harsh "
     "chemicals in the baby room."),
    ("Nguyen Household", "Residential", "Lien Nguyen", "(503) 555 0259",
     "lien.nguyen@example.com", "88 Alder Way, Portland OR 97211",
     "Active", "Facebook", -210, "Monthly", "Wednesday",
     "Morning (9-12)", "Jade Ellis", 45, 7, "Rolling monthly", "No",
     "Shoes off at the door. Key is in the lockbox."),
    ("Summit Fitness Studio", "Gym & studio", "Dev Patel",
     "(503) 555 0333", "dev@summitfitness.example",
     "501 Summit Drive, Portland OR 97201", "Active", "Instagram",
     -180, "Twice a week", "Monday", "Early (7-9)", "Luke Bennett", 52,
     14, "6 months", "No",
     "Mirrors and mats matter more than floors here."),
    ("Hartley Holiday Lets", "Holiday let / Airbnb", "Rosie Hartley",
     "(503) 555 0288", "rosie@hartleylets.example",
     "6 Quay Street, Portland OR 97209", "Regular", "Nextdoor", -330,
     "Weekly", "Saturday", "Midday (12-2)", "Jade Ellis", 50, 7,
     "Rolling monthly", "Yes",
     "Changeover between 11 and 3. Linen is in the hall cupboard."),
    ("Donovan Residence", "Residential", "Paul Donovan", "(503) 555 0194",
     "paul.donovan@example.com", "27 Birch Close, Portland OR 97215",
     "Paused", "Van signage", -480, "Fortnightly", "Thursday",
     "Morning (9-12)", "Marta Kovacs", 45, 7, "Rolling monthly", "No",
     "Paused for the summer, back in September - call in August."),
    ("Willow Park Estates", "Property manager", "Grace Lin",
     "(503) 555 0401", "grace@willowpark.example",
     "Willow Park, Block A-D, Portland OR 97217", "Active",
     "Referral - staff", -150, "Monthly", "Wednesday",
     "Afternoon (2-5)", "Luke Bennett", 48, 30, "12 months", "No",
     "Four communal stairwells and the bin stores. Invoice monthly."),
    ("Byrne Residence", "Residential", "Aoife Byrne", "(503) 555 0276",
     "aoife.byrne@example.com", "14 Rosewood Terrace, Portland OR 97206",
     "New", "Google Business Profile", -20, "Fortnightly", "Tuesday",
     "Afternoon (2-5)", "Marta Kovacs", 45, 7, "Trial period", "No",
     "Came from a deep clean quote. Trialling a fortnightly visit."),
    ("Stonebridge Law", "Commercial office", "Marcus Webb",
     "(503) 555 0388", "m.webb@stonebridgelaw.example",
     "4 Stonebridge Court, Portland OR 97204", "Active", "Yelp", -95,
     "Weekly", "Friday", "Late (5-8)", "Tomas Ruiz", 52, 14,
     "Rolling monthly", "No",
     "Confidential paperwork - never move anything off a desk."),
    ("Alvarez Family", "Residential", "Ines Alvarez", "(503) 555 0215",
     "ines.alvarez@example.com", "63 Hawthorn Road, Portland OR 97203",
     "Inactive", "Flyer / leaflet", -700, "One-off", "Flexible",
     "Flexible", "Jade Ellis", 45, 7, "No contract", "No",
     "Moved out of the area in the spring. Left a five star review."),
]

# ---------------------------------------------------------------------------
# The properties we actually clean
# ---------------------------------------------------------------------------
# client index, name, type, address, beds, baths, rooms, area, condition,
# access, key, parking, pets, floors, hours, notes
PROPERTIES = [
    (1, "Harper - family home", "House",
     "42 Larch Street, Portland OR 97215", 3, 2, 4, 1450, "Normal",
     "We hold a key", "Held by us", "Driveway", "Dog & cat",
     "Mixed", 3.0, "Cat flap in the utility door. Dog is friendly."),
    (2, "Okafor - townhouse", "Townhouse",
     "9 Elmwood Court, Portland OR 97212", 4, 2, 5, 1700, "Needs attention",
     "Client at home", "Not applicable", "Free street parking", "None",
     "Carpet & tile", 3.5, "Playroom gets the worst of it."),
    (3, "Riverbend Dental - practice", "Clinic",
     "118 Riverbend Plaza, Suite 3, Portland OR 97209", 0, 3, 8, 2200,
     "Normal", "Reception sign-in", "Not applicable", "Visitor bay",
     "None", "Vinyl", 2.5,
     "Four surgeries, waiting room, two staff rooms. Clinical bins "
     "are not ours."),
    (4, "Cedar & Co - second floor", "Office suite",
     "77 Cedar Avenue, Floor 2, Portland OR 97204", 0, 2, 9, 3100,
     "Very clean - upkeep only", "Keypad code", "Not applicable",
     "Car park", "None", "Carpet", 2.0,
     "Alarm: enter code within 30 seconds, panel by the lift."),
    (5, "Lakeview - 12B Lakeview Road", "Apartment",
     "12B Lakeview Road, Portland OR 97202", 2, 1, 3, 780,
     "Heavily soiled", "Agent meets us", "Not applicable",
     "Metered street", "None", "Laminate", 6.0,
     "End of tenancy. Oven and extractor are the big jobs."),
    (6, "The Maple Cafe - front of house", "Restaurant",
     "210 Maple Street, Portland OR 97214", 0, 2, 4, 1200, "Normal",
     "We hold a key", "Held by us", "Loading bay", "None", "Tile", 2.0,
     "Before 8am. Chairs down, floors mopped, glass polished."),
    (7, "Bright Start Nursery - main building", "School",
     "15 Juniper Lane, Portland OR 97213", 0, 4, 10, 2600, "Normal",
     "We hold a key", "In the key safe", "Driveway", "None", "Vinyl",
     3.5, "Toy sanitising on Fridays. Log it on the checklist."),
    (8, "Nguyen - family home", "House",
     "88 Alder Way, Portland OR 97211", 4, 2, 5, 1900, "Normal",
     "Lockbox", "In the key safe", "Driveway", "Cat", "Hardwood", 3.5,
     "Lockbox code changes every quarter - check before you go."),
    (9, "Summit Fitness - studio", "Gym",
     "501 Summit Drive, Portland OR 97201", 0, 2, 5, 2400, "Normal",
     "Keypad code", "Not applicable", "Car park", "None", "Rubber",
     2.5, "Mats, mirrors, changing rooms. Before 7am."),
    (10, "Hartley - Quay Street loft", "Holiday let",
     "6 Quay Street, Portland OR 97209", 2, 1, 3, 950, "Normal",
     "Lockbox", "In the key safe", "Permit needed", "None", "Hardwood",
     2.0, "Full changeover: linen, towels, welcome pack, photos."),
    (10, "Hartley - Riverside studio", "Holiday let",
     "19 Riverside Walk, Portland OR 97209", 1, 1, 2, 540, "Normal",
     "Lockbox", "In the key safe", "Pay and display", "None",
     "Laminate", 1.5, "Small but fiddly. Same checklist as the loft."),
    (11, "Donovan - bungalow", "Bungalow",
     "27 Birch Close, Portland OR 97215", 2, 1, 4, 1100, "Normal",
     "We hold a key", "Held by us", "Driveway", "Dog", "Carpet", 3.0,
     "Paused until September."),
    (12, "Willow Park - Block A", "Communal area",
     "Willow Park Block A, Portland OR 97217", 0, 0, 3, 1600, "Normal",
     "Concierge / front desk", "Not applicable", "Visitor bay", "None",
     "Stone", 1.5, "Stairwell, lift lobby, bin store."),
    (12, "Willow Park - Block B", "Communal area",
     "Willow Park Block B, Portland OR 97217", 0, 0, 3, 1600, "Normal",
     "Concierge / front desk", "Not applicable", "Visitor bay", "None",
     "Stone", 1.5, "Same scope as Block A."),
    (13, "Byrne - terrace", "House",
     "14 Rosewood Terrace, Portland OR 97206", 3, 1, 4, 1250,
     "Needs attention", "Client at home", "Not applicable",
     "Free street parking", "Dog", "Mixed", 3.0,
     "First deep clean done, now on a fortnightly trial."),
    (14, "Stonebridge Law - ground floor", "Office suite",
     "4 Stonebridge Court, Portland OR 97204", 0, 2, 7, 2100,
     "Very clean - upkeep only", "Smart lock", "Not applicable",
     "Underground garage", "None", "Carpet", 2.0,
     "Desks are never to be cleared. Bins and floors only."),
]

# ---------------------------------------------------------------------------
# What we sell
# ---------------------------------------------------------------------------
# name, category, model, hours, price, unit, uplift, supplies, popular,
# description
SERVICES = [
    ("Regular home clean", "Regular domestic clean", "Per hour", 3.0, 45,
     "hour", 0, 7.50, "Yes",
     "Kitchen, bathrooms, floors, dusting and beds made. The weekly "
     "or fortnightly visit most homes book."),
    ("First deep clean", "Deep clean", "Per hour", 6.0, 45, "hour", 0.45,
     18.00, "Yes",
     "The one before the regular visits start: skirting boards, "
     "inside the oven, limescale, behind furniture."),
    ("End of tenancy clean", "End of tenancy", "Per area", 6.0, 0.09,
     "sq ft", 0.35, 24.00, "Yes",
     "Priced on floor area. Everything cleaned inside and out so the "
     "deposit comes back."),
    ("Move in clean", "Move in / move out", "Per area", 5.0, 0.08,
     "sq ft", 0.35, 20.00, "No",
     "Before the furniture arrives. Cupboards, appliances, floors."),
    ("Office clean - small", "Commercial & office", "Flat price per visit",
     2.0, 110, "visit", 0, 6.00, "Yes",
     "Up to about 2,000 sq ft: desks, kitchen, washrooms, bins, "
     "floors."),
    ("Office clean - large", "Commercial & office", "Flat price per visit",
     3.5, 185, "visit", 0, 10.00, "No",
     "2,000 to 4,000 sq ft, same scope as the small office clean."),
    ("Airbnb changeover", "Airbnb changeover", "Flat price per visit",
     2.0, 95, "visit", 0, 8.00, "Yes",
     "Full reset between guests: linen, towels, bathroom, kitchen, "
     "welcome pack and photos."),
    ("Communal area clean", "Commercial & office", "Flat price per visit",
     1.5, 75, "visit", 0, 4.50, "No",
     "Stairwells, lobbies and bin stores for blocks and estates."),
    ("Nursery & school clean", "Commercial & office",
     "Flat price per visit", 3.5, 175, "visit", 0, 11.00, "No",
     "Classroom floors, toilets, toy sanitising and touch points."),
    ("Restaurant & kitchen clean", "Commercial & office",
     "Flat price per visit", 2.0, 125, "visit", 0, 8.00, "No",
     "Front of house, kitchen floors, surfaces and the bin area. "
     "Before opening or after close."),
    ("Oven clean", "Add-on / extra", "Per item", 1.5, 85, "oven", 0,
     9.00, "Yes",
     "Dip tank treatment for the racks, full degrease of the cavity, "
     "door glass split and cleaned."),
    ("Interior window clean", "Add-on / extra", "Per item", 0.75, 45,
     "property", 0, 3.00, "No",
     "Glass, frames and sills inside. Exterior is priced separately."),
    ("Exterior window clean", "Window cleaning", "Per item", 1.0, 60,
     "property", 0, 2.00, "No",
     "Pole fed, up to three storeys. Subcontracted to Chris."),
    ("Carpet clean", "Carpet & upholstery", "Per room", 1.0, 55, "room",
     0, 7.00, "No",
     "Hot water extraction. Allow four hours drying before furniture "
     "goes back."),
    ("Upholstery clean", "Carpet & upholstery", "Per item", 1.0, 65,
     "item", 0, 6.00, "No",
     "Sofas and armchairs, fabric only. Leather is a different "
     "treatment."),
    ("Inside fridge", "Add-on / extra", "Per item", 0.5, 30, "item", 0,
     2.50, "No", "Shelves out, washed, dried and put back."),
    ("Inside kitchen cupboards", "Add-on / extra", "Per item", 1.0, 45,
     "kitchen", 0, 3.50, "No",
     "Emptied, wiped inside and out, lined if you supply the liner."),
    ("Spring clean", "Spring clean", "Per hour", 5.0, 45, "hour", 0.25,
     14.00, "No",
     "The annual going-over: windows, blinds, cupboards, skirtings "
     "and the backs of things."),
    ("Post-construction clean", "Post-construction", "Per area", 8.0,
     0.14, "sq ft", 0.35, 32.00, "No",
     "Dust everywhere, paint specks, sticker removal. Price it after "
     "a site visit."),
    ("Pressure wash patio", "Pressure washing", "Per area", 2.0, 0.35,
     "sq ft", 0, 4.00, "No",
     "Patios, paths and driveways. Weather dependent."),
    ("Laundry & ironing", "Add-on / extra", "Per hour", 1.0, 38, "hour",
     0, 1.50, "No", "Charged by the hour at the property."),
]

# ---------------------------------------------------------------------------
# The team
# ---------------------------------------------------------------------------
# name, role, type, phone, email, start offset, pay, hours/week, status,
# skills, checked, notes
EMPLOYEES = [
    ("Dana Whitfield", "Owner", "Owner", "(503) 555 0142",
     "dana@brightandtidy.co", -2190, 25.00, 20, "Active",
     "Quoting, deep cleans, everything else", "Yes",
     "Quotes, invoices and the jobs nobody else can fit in."),
    ("Marta Kovacs", "Team leader", "Full time", "(503) 555 0511",
     "marta@brightandtidy.co", -1460, 21.00, 38, "Active",
     "Residential, deep clean, training", "Yes",
     "Runs van 1 and trains every new starter."),
    ("Jade Ellis", "Cleaner", "Full time", "(503) 555 0522",
     "jade@brightandtidy.co", -980, 18.50, 37, "Active",
     "Residential, changeovers", "Yes", "Fast and reliable on homes."),
    ("Tomas Ruiz", "Cleaner", "Full time", "(503) 555 0533",
     "tomas@brightandtidy.co", -740, 18.50, 36, "Active",
     "Commercial, out of hours", "Yes",
     "Does the evening office round. Holds the alarm codes."),
    ("Priya Raman", "Deep clean specialist", "Full time",
     "(503) 555 0544", "priya@brightandtidy.co", -610, 20.00, 35,
     "Active", "Deep clean, end of tenancy, carpets", "Yes",
     "The one to send when the place is in a state."),
    ("Luke Bennett", "Cleaner", "Part time", "(503) 555 0555",
     "luke@brightandtidy.co", -300, 17.50, 20, "Active",
     "Commercial, communal areas", "Yes",
     "Mornings only - studies in the afternoon."),
    ("Amara Diallo", "Cleaner", "Full time", "(503) 555 0566",
     "amara@brightandtidy.co", -200, 18.50, 35, "Probation",
     "Residential, cafe, nursery", "Yes",
     "Three months in and already asked for by name."),
    ("Chris Boyd", "Window cleaner", "Subcontractor", "(503) 555 0577",
     "chris@clearviewwindows.example", -520, 28.00, 8, "Active",
     "Exterior windows, pressure washing", "Yes",
     "Invoices us per property. Has his own insurance - on file."),
]


# ---------------------------------------------------------------------------
# Per property extras: how far away, which checklist, how to get in
# ---------------------------------------------------------------------------
# property number -> distance, checklist (a service name), on-site supplies
PROPERTY_EXTRA = {
    1: (4.2, "Regular home clean", "No"),
    2: (6.0, "Regular home clean", "No"),
    3: (3.1, "Office clean - small", "Yes"),
    4: (2.4, "Office clean - small", "Yes"),
    5: (7.5, "End of tenancy clean", "No"),
    6: (5.0, "Restaurant & kitchen clean", "Yes"),
    7: (4.6, "Nursery & school clean", "Yes"),
    8: (8.2, "Regular home clean", "No"),
    9: (3.8, "Office clean - small", "Yes"),
    10: (2.9, "Airbnb changeover", "Yes"),
    11: (3.4, "Airbnb changeover", "Yes"),
    12: (6.8, "Regular home clean", "No"),
    13: (9.1, "Communal area clean", "No"),
    14: (9.1, "Communal area clean", "No"),
    15: (5.5, "Regular home clean", "No"),
    16: (2.2, "Office clean - small", "Yes"),
}

ENTRY_NOTES = {
    "We hold a key": "Key 12 on the board. Alarm panel inside the front "
                     "door, code on the client card.",
    "Client at home": "Someone is always in. Ring, do not knock - the "
                      "bell wakes the baby.",
    "Reception sign-in": "Sign in at reception, collect the visitor "
                         "badge, sign out on the way through.",
    "Keypad code": "Keypad code then the green button. Re-arm on the "
                   "way out or the alarm company calls the owner.",
    "Agent meets us": "The agent unlocks and leaves. Call them when "
                      "you are finished so they can lock up.",
    "Lockbox": "Lockbox on the gatepost. Put the key back before you "
               "leave, never in your pocket.",
    "Smart lock": "One-time code sent the morning of the visit. It "
                  "expires at midnight.",
    "Open access": "Communal doors are open during the day. Bin store "
                   "key is on the ring.",
    "Key collected each visit": "Collect the key from the office before "
                                "you set off.",
    "Concierge / front desk": "Concierge holds the key and will want "
                              "the signature book.",
    "Neighbour holds key": "Number 44 next door holds the key. Text "
                           "first, she is often out before nine.",
}

# ---------------------------------------------------------------------------
# Master checklists, one per service
# ---------------------------------------------------------------------------
# area, task, how it should look, minutes, how often, required, photo
CHECKLIST_TASKS = {
    "Regular home clean": [
        ("Kitchen", "Wipe down worktops and splashback",
         "Clear, dry and streak free - nothing left on the surface",
         10, "Every visit", "Yes", "No"),
        ("Kitchen", "Clean sink and taps, polish dry",
         "No limescale round the base of the tap", 8, "Every visit",
         "Yes", "No"),
        ("Kitchen", "Front of cupboards and appliances",
         "No finger marks, handles wiped", 10, "Every visit", "Yes",
         "No"),
        ("Kitchen", "Hob and extractor front",
         "Degreased, no burnt-on spots", 10, "Every visit", "Yes",
         "No"),
        ("Kitchen", "Inside the microwave",
         "Wiped out, turntable washed", 5, "Weekly", "No", "No"),
        ("Bathroom", "Descale and polish the shower screen",
         "Clear glass, no water spots", 12, "Every visit", "Yes",
         "Yes"),
        ("Bathroom", "Clean and disinfect the toilet",
         "Inside, outside, behind and the seat hinges", 8,
         "Every visit", "Yes", "No"),
        ("Bathroom", "Basin, taps and mirror",
         "Polished, no splash marks on the mirror", 8, "Every visit",
         "Yes", "No"),
        ("Bathroom", "Floor washed, edges included",
         "Hair cleared from corners and behind the door", 8,
         "Every visit", "Yes", "No"),
        ("Bedroom", "Change the bed linen if left out",
         "Bed made properly, old linen in the basket", 8, "Weekly",
         "No", "No"),
        ("Bedroom", "Dust surfaces and headboard",
         "Nothing moved out of place, everything put back", 8,
         "Every visit", "Yes", "No"),
        ("Living room", "Dust and polish all surfaces",
         "Including skirtings, sills and the top of the frames", 12,
         "Every visit", "Yes", "No"),
        ("Living room", "Vacuum carpets and under the cushions",
         "Lines left in the carpet, cushions plumped", 12,
         "Every visit", "Yes", "No"),
        ("Hallway & stairs", "Vacuum stairs and hall, wipe bannister",
         "Edges done, no dust on the spindles", 10, "Every visit",
         "Yes", "No"),
        ("Whole property", "Empty bins and replace liners",
         "Bins wiped if sticky, rubbish taken to the outside bin", 6,
         "Every visit", "Yes", "No"),
        ("Whole property", "Final walk through and lock up",
         "Lights off, windows shut, door locked, key where it should "
         "be", 5, "Every visit", "Yes", "No"),
    ],
    "Office clean - small": [
        ("Reception", "Reception desk and glass door",
         "No finger marks on the glass at eye height", 8,
         "Every visit", "Yes", "No"),
        ("Open plan office", "Empty every bin, replace liners",
         "Including the recycling under each desk", 10, "Every visit",
         "Yes", "No"),
        ("Open plan office", "Wipe desks around equipment",
         "Nothing moved, no papers disturbed", 12, "Every visit",
         "Yes", "No"),
        ("Open plan office", "Vacuum throughout",
         "Under the desks as well as the walkways", 15, "Every visit",
         "Yes", "No"),
        ("Kitchenette", "Worktops, sink and the front of the fridge",
         "Dry, no crumbs, no mugs left in the sink", 10,
         "Every visit", "Yes", "No"),
        ("Toilets", "Clean and disinfect, restock consumables",
         "Paper, soap and towels topped up", 15, "Every visit", "Yes",
         "No"),
        ("Meeting room", "Table, chairs and whiteboard",
         "Board wiped properly, chairs pushed in", 8, "Every visit",
         "Yes", "No"),
        ("Whole property", "Touch points - handles and switches",
         "Every handle, switch and lift button", 8, "Every visit",
         "Yes", "No"),
        ("Whole property", "Set the alarm and sign out",
         "Alarm armed, badge returned, log signed", 4, "Every visit",
         "Yes", "No"),
    ],
    "Airbnb changeover": [
        ("Whole property", "Strip beds and start the laundry",
         "Linen counted - report anything stained or missing", 10,
         "Every visit", "Yes", "Yes"),
        ("Bedroom", "Make up beds, hotel finish",
         "Crisp, square corners, cushions set out", 15, "Every visit",
         "Yes", "Yes"),
        ("Bathroom", "Full clean and fresh towels",
         "Towels folded the way the host likes them", 15,
         "Every visit", "Yes", "Yes"),
        ("Kitchen", "Empty, clean and restock",
         "Fridge emptied of guest food, dishwasher run and emptied",
         15, "Every visit", "Yes", "No"),
        ("Whole property", "Welcome pack and consumables",
         "Coffee, tea, milk, paper and soap replaced", 8,
         "Every visit", "Yes", "No"),
        ("Whole property", "Check for damage and lost property",
         "Photograph anything at all - before the next guest arrives",
         8, "Every visit", "Yes", "Yes"),
        ("Living room", "Reset the room to the photos",
         "Furniture and cushions exactly as the listing shows", 10,
         "Every visit", "Yes", "Yes"),
        ("Whole property", "Vacuum and mop through",
         "No sand, no crumbs, no hair", 15, "Every visit", "Yes",
         "No"),
        ("Whole property", "Bins out, recycling sorted",
         "Collection day is Tuesday - put them out the night before",
         5, "Every visit", "Yes", "No"),
        ("Whole property", "Photos to the host",
         "Six photos: each room ready for the guest", 5,
         "Every visit", "Yes", "Yes"),
    ],
    "End of tenancy clean": [
        ("Kitchen", "Oven, racks and extractor",
         "Degreased to bare metal - this is what the deposit turns "
         "on", 60, "Every visit", "Yes", "Yes"),
        ("Kitchen", "Inside every cupboard and drawer",
         "Emptied, washed, dried, left open to air", 35,
         "Every visit", "Yes", "No"),
        ("Kitchen", "Fridge and freezer, defrosted",
         "Switched off, propped open, shelves washed", 30,
         "Every visit", "Yes", "Yes"),
        ("Kitchen", "Behind and beside the appliances",
         "Pull them out if you safely can", 20, "Every visit", "Yes",
         "No"),
        ("Bathroom", "Descale everything",
         "Tiles, screen, taps, shower head soaked", 35, "Every visit",
         "Yes", "Yes"),
        ("Whole property", "Interior windows, frames and sills",
         "Including the runners", 30, "Every visit", "Yes", "No"),
        ("Whole property", "Skirtings, doors and frames",
         "Washed, not just dusted", 30, "Every visit", "Yes", "No"),
        ("Whole property", "Light fittings and switches",
         "Shades dusted, switches wiped", 15, "Every visit", "Yes",
         "No"),
        ("Whole property", "Carpets vacuumed and spot treated",
         "Note any stain that will not come out", 25, "Every visit",
         "Yes", "Yes"),
        ("Whole property", "Hard floors washed",
         "Edges and corners by hand", 20, "Every visit", "Yes", "No"),
        ("Whole property", "Cobwebs, vents and extractor covers",
         "Look up - inventory clerks do", 15, "Every visit", "Yes",
         "No"),
        ("Whole property", "Photograph every room finished",
         "Date stamped. It settles deposit arguments", 10,
         "Every visit", "Yes", "Yes"),
    ],
    "Communal area clean": [
        ("Communal area", "Vacuum stairs and landings",
         "Including the edges and the corners", 20, "Every visit",
         "Yes", "No"),
        ("Communal area", "Mop the entrance hall",
         "Wet floor sign out while it dries", 15, "Every visit",
         "Yes", "No"),
        ("Communal area", "Glass in the entrance doors",
         "Both sides, no smears", 10, "Every visit", "Yes", "No"),
        ("Communal area", "Bin store - sweep and deodorise",
         "Bins pushed back, lids closed", 15, "Every visit", "Yes",
         "Yes"),
        ("Communal area", "Dust handrails, sills and notice boards",
         "Old notices taken down", 10, "Every visit", "Yes", "No"),
        ("Communal area", "Report anything broken or dumped",
         "Photograph it and send it to the managing agent", 5,
         "Every visit", "Yes", "Yes"),
    ],
    "Nursery & school clean": [
        ("Whole property", "Sanitise all touch points",
         "Handles, switches, rails, table edges - approved product "
         "only", 20, "Every visit", "Yes", "No"),
        ("Whole property", "Toy surfaces and play mats",
         "Child-safe sanitiser, left to air dry", 25, "Every visit",
         "Yes", "No"),
        ("Toilets", "Child toilets and low basins",
         "Disinfected, steps wiped, paper restocked", 25,
         "Every visit", "Yes", "No"),
        ("Kitchen", "Food prep surfaces",
         "Food-safe sanitiser, separate cloths", 15, "Every visit",
         "Yes", "No"),
        ("Whole property", "Floors vacuumed and mopped",
         "Under the tables, along the skirtings", 30, "Every visit",
         "Yes", "No"),
        ("Whole property", "Bins emptied, nappy bin sealed",
         "Nappy waste goes in the designated bin only", 10,
         "Every visit", "Yes", "No"),
        ("Whole property", "Sign the cleaning record",
         "The inspector will ask for it", 5, "Every visit", "Yes",
         "No"),
    ],
    "Restaurant & kitchen clean": [
        ("Whole property", "Sweep and mop front of house",
         "Chairs up, under the tables, back down", 25, "Every visit",
         "Yes", "No"),
        ("Whole property", "Tables, booths and high chairs",
         "Sanitised, no sticky edges", 15, "Every visit", "Yes",
         "No"),
        ("Kitchen", "Kitchen floor degreased",
         "Including under the pass and behind the fryer", 25,
         "Every visit", "Yes", "No"),
        ("Kitchen", "Stainless steel surfaces polished",
         "With the grain, no smears", 15, "Every visit", "Yes", "No"),
        ("Toilets", "Customer toilets, full clean and restock",
         "Checked twice - this is what reviews mention", 20,
         "Every visit", "Yes", "No"),
        ("Whole property", "Bin area and bottle store",
         "Hosed, deodorised, bins returned", 10, "Every visit", "Yes",
         "No"),
        ("Whole property", "Entrance glass and menu board",
         "First thing a customer sees", 10, "Every visit", "Yes",
         "No"),
    ],
}

# ---------------------------------------------------------------------------
# The cupboard
# ---------------------------------------------------------------------------
# name, category, supplier, unit, pack size, pack cost, stock, reorder,
# order qty, eco, coshh, notes
SUPPLIES = [
    ("Multi-surface cleaner", "Chemicals & detergents", "CleanCo Supplies",
     "5L drum", 5, 18.40, 6, 3, 6, "Yes", "Yes",
     "Diluted 1:40 in the trigger bottles. Data sheet on file."),
    ("Bathroom descaler", "Chemicals & detergents", "CleanCo Supplies",
     "5L drum", 5, 22.80, 3, 2, 4, "No", "Yes",
     "Acidic - never mix with bleach. Gloves and goggles."),
    ("Kitchen degreaser", "Chemicals & detergents", "CleanCo Supplies",
     "5L drum", 5, 24.50, 2, 2, 4, "No", "Yes",
     "The cafe and the oven work eat this."),
    ("Glass cleaner", "Chemicals & detergents", "CleanCo Supplies",
     "Bottle", 1, 3.10, 14, 6, 12, "Yes", "No",
     "Buy the concentrate, it is a third of the price."),
    ("Toilet cleaner", "Chemicals & detergents", "Wholesale Direct",
     "Bottle", 1, 2.40, 11, 6, 12, "No", "Yes", ""),
    ("Disinfectant - nursery safe", "Specialist products",
     "Hygiene Partners", "5L drum", 5, 31.00, 2, 1, 2, "Yes", "Yes",
     "The only product allowed in the baby room."),
    ("Microfibre cloths - colour coded", "Cloths & sponges",
     "Wholesale Direct", "Pack", 10, 9.90, 9, 4, 8, "Yes", "No",
     "Red bathroom, blue general, green kitchen, yellow glass."),
    ("Scourer sponges", "Cloths & sponges", "Wholesale Direct", "Pack",
     10, 4.20, 7, 3, 6, "No", "No", ""),
    ("Mop heads", "Mops & buckets", "CleanCo Supplies", "Each", 1, 6.50,
     5, 4, 8, "No", "No", "Washed at 60 after every commercial job."),
    ("Bin bags - heavy duty", "Bin bags & waste", "Wholesale Direct",
     "Roll", 50, 7.80, 8, 4, 10, "No", "No", ""),
    ("Nitrile gloves", "Gloves & PPE", "Hygiene Partners", "Box", 100,
     11.50, 4, 3, 6, "No", "No", "Powder free. Two sizes in the van."),
    ("Toilet rolls - commercial", "Paper & consumables",
     "Wholesale Direct", "Case", 36, 28.00, 2, 2, 4, "Yes", "No",
     "Recharged to the office clients at cost plus 10%."),
    ("Hand soap refill", "Paper & consumables", "Wholesale Direct",
     "Refill", 1, 4.60, 6, 3, 6, "Yes", "No", ""),
    ("Vacuum bags - Sebo", "Vacuum bags & filters", "Vac Spares Online",
     "Pack", 10, 12.40, 3, 2, 4, "No", "No",
     "Only fits the two uprights, not the backpack."),
    ("Laundry detergent", "Laundry", "Wholesale Direct", "5L drum", 5,
     14.20, 2, 1, 2, "Yes", "No", "For the changeover linen."),
    ("Oven cleaning gel", "Specialist products", "Oven Pro Supplies",
     "Bottle", 1, 16.90, 3, 2, 4, "No", "Yes",
     "Caustic. Full PPE, windows open, never on a hot oven."),
]

# ---------------------------------------------------------------------------
# The kit
# ---------------------------------------------------------------------------
# name, type, make, serial, ownership, status, assigned, location,
# bought, cost, life, warranty, last service, interval, service cost,
# safety test, insured, notes
EQUIPMENT = [
    ("Sebo upright - van 1", "Upright vacuum", "Sebo Automatic X7",
     "SB-774210", "Owned outright", "In service", "Marta Kovacs",
     "Van 1", -980, 385.00, 7, -250, -120, 180, 45.00, -120, "Yes",
     "Workhorse. Belt changed every spring."),
    ("Sebo upright - van 2", "Upright vacuum", "Sebo Automatic X7",
     "SB-774588", "Owned outright", "In service", "Tomas Ruiz",
     "Van 2", -610, 385.00, 7, -245, -200, 180, 45.00, -200, "Yes",
     ""),
    ("Numatic backpack", "Backpack vacuum", "Numatic RSB150",
     "NM-220914", "Owned outright", "In use", "Luke Bennett", "Van 2",
     -420, 299.00, 6, -55, -95, 180, 35.00, -95, "Yes",
     "For the stairwells - saves twenty minutes a block."),
    ("Wet and dry vacuum", "Wet & dry vacuum", "Karcher WD6",
     "KR-556102", "Owned outright", "Spare", "", "Store cupboard",
     -730, 189.00, 6, "", -310, 365, 30.00, -310, "Yes",
     "Comes out for floods and post-construction work."),
    ("Carpet extraction machine", "Carpet machine", "Prochem Steempro",
     "PC-100483", "Financed", "In service", "Priya Raman", "Van 1",
     -520, 1850.00, 8, -155, -40, 180, 120.00, -40, "Yes",
     "On finance until next March. Earns its keep twice over."),
    ("Steam cleaner", "Steam cleaner", "Karcher SC4", "KR-889321",
     "Owned outright", "In repair", "", "Repair shop", -400, 245.00,
     5, "", -210, 365, 55.00, -210, "Yes",
     "Boiler seal gone. Quoted 60 to fix, 245 to replace."),
    ("Pressure washer", "Pressure washer", "Nilfisk C135",
     "NF-443027", "Owned outright", "In service", "Chris Boyd",
     "Van 2", -300, 320.00, 6, -120, -150, 365, 40.00, -150, "Yes",
     "Patio and bin store work only."),
    ("Window pole system", "Window pole system", "Unger HiFlo",
     "UG-770145", "Owned outright", "In service", "Chris Boyd",
     "Van 2", -510, 640.00, 8, "", -180, 365, 25.00, -180, "Yes",
     "Pure water system, three storey reach."),
    ("Step ladder - 5 tread", "Ladder & steps", "Youngman 5",
     "YM-003311", "Owned outright", "In service", "", "Van 1", -900,
     95.00, 10, "", -75, 365, 0.00, -75, "Yes",
     "Inspected quarterly - log it here."),
    ("Cleaning trolley", "Trolley & caddy", "Numatic NX1",
     "NM-551209", "Owned outright", "In service", "Tomas Ruiz",
     "Cedar & Co site", -350, 210.00, 8, "", "", 365, 0.00, "", "Yes",
     "Lives at the office site - saves carrying."),
    ("Van 1 - Transit Connect", "Vehicle", "Ford Transit Connect",
     "VIN-9921", "Financed", "In service", "Marta Kovacs", "Van 1",
     -1100, 18500.00, 8, -370, -85, 180, 320.00, -85, "Yes",
     "Serviced every six months, MOT in the autumn."),
    ("Van 2 - Berlingo", "Vehicle", "Citroen Berlingo", "VIN-7740",
     "Financed", "In service", "Tomas Ruiz", "Van 2", -700, 16200.00,
     8, -340, -170, 180, 295.00, -170, "Yes", ""),
]

# ---------------------------------------------------------------------------
# The paperwork that has to stay in date
# ---------------------------------------------------------------------------
# document, applies to, employee, reference, issuer, issued, expires,
# cost, holder, required, reviewed, notes
SAFETY_DOCS = [
    ("Public liability insurance", "The business", "", "PL-884213",
     "Hiscox", -300, 65, 486.00, "Insurance folder, office", "Yes",
     -300, "2,000,000 cover. Commercial clients ask for this first."),
    ("Employer liability insurance", "The business", "", "EL-884214",
     "Hiscox", -300, 65, 310.00, "Insurance folder, office", "Yes",
     -300, "Legally required once you employ anyone."),
    ("Vehicle insurance", "Van 1", "", "MV-220714", "Admiral Business",
     -200, 165, 1180.00, "Van document wallet", "Yes", -200, ""),
    ("Vehicle insurance", "Van 2", "", "MV-220715", "Admiral Business",
     -200, 165, 1090.00, "Van document wallet", "Yes", -200, ""),
    ("Background check", "Staff", "Amara Diallo", "BC-55210",
     "CheckPoint", -195, 900, 44.00, "Staff file", "Yes", -195,
     "Required for the nursery contract."),
    ("Background check", "Staff", "Marta Kovacs", "BC-41088",
     "CheckPoint", -1450, 180, 44.00, "Staff file", "Yes", -1450,
     "Due for renewal - the nursery audits this every year."),
    ("Background check", "Staff", "Jade Ellis", "BC-47720",
     "CheckPoint", -970, 740, 44.00, "Staff file", "Yes", -970, ""),
    ("Training certificate", "Staff", "Priya Raman", "TR-COSHH-220",
     "BICSc", -610, 455, 120.00, "Staff file", "Yes", -610,
     "COSHH handling. Renew every three years."),
    ("First aid certificate", "Staff", "Marta Kovacs", "FA-77120",
     "Red Cross", -720, 375, 95.00, "Staff file", "No", -720,
     "One trained first aider is enough for a team this size."),
    ("COSHH data sheet", "Oven cleaning gel", "", "SDS-OVEN-04",
     "Oven Pro Supplies", -400, 330, 0.00, "Chemical folder, van 1",
     "Yes", -120, "Must travel with the product, not stay in the "
     "office."),
    ("COSHH data sheet", "Bathroom descaler", "", "SDS-DESC-11",
     "CleanCo Supplies", -380, 350, 0.00, "Chemical folder, van 1",
     "Yes", -120, ""),
    ("Risk assessment", "Working at height", "", "RA-004",
     "Dana Whitfield", -390, 340, 0.00, "Health and safety folder",
     "Yes", -390,
     "Covers steps and pole work. Reviewed annually - this one is "
     "nearly due."),
    ("Risk assessment", "Lone working", "", "RA-007",
     "Dana Whitfield", -210, 520, 0.00, "Health and safety folder",
     "Yes", -210, "Check-in text at the end of every evening job."),
    ("Method statement", "End of tenancy cleaning", "", "MS-002",
     "Dana Whitfield", -260, 470, 0.00, "Health and safety folder",
     "No", -260, "Letting agents ask for this with the quote."),
    ("Electrical safety test", "Equipment", "", "PAT-2026-01",
     "SafeTest Ltd", -85, 280, 140.00, "Equipment folder", "Yes",
     -85, "All twelve items tested. Stickers on each machine."),
    ("Waste carrier licence", "The business", "", "WCL-99120",
     "Environment Agency", -500, 230, 154.00, "Office folder", "Yes",
     -500, "Needed to take trade waste away from commercial jobs."),
]

# ---------------------------------------------------------------------------
# Enquiries in the pipeline
# ---------------------------------------------------------------------------
# days ago, name, phone, email, address, type, source, service,
# frequency, stage, value, quote no, visit offset, contacted,
# next action, next offset, client no, lost reason, notes
LEADS = [
    (-2, "Ellis Shaw", "(503) 555 0612", "ellis.shaw@example.com",
     "31 Fir Street, Portland OR 97215", "Residential", "Google search",
     "Regular home clean", "Fortnightly", "New enquiry", 150, "", "",
     "No", "Call back and book a quote visit", 1, "", "",
     "Rang while we were on a job. Three bed, two bath, one dog."),
    (-3, "Priya Nair", "(503) 555 0633", "priya.nair@example.com",
     "8 Aspen Court, Portland OR 97212", "Residential",
     "Referral - client", "Regular home clean", "Weekly", "Contacted",
     220, "", "", "Yes", "Send the quote", 1, "", "",
     "Referred by the Okafors. Wants the same cleaner every week."),
    (-6, "Marsh & Fletcher Dental", "(503) 555 0655",
     "office@marshfletcher.example",
     "202 Bridge Road, Portland OR 97209", "Medical & dental", "Yelp",
     "Office clean - small", "Weekly", "Visit booked", 480, "", 2,
     "Yes", "Site visit Thursday at 6pm", 2, "", "",
     "Second practice in the area - could lead to both."),
    (-9, "Tom Ridley", "(503) 555 0677", "t.ridley@example.com",
     "77 Chestnut Avenue, Portland OR 97206", "Landlord",
     "Google Business Profile", "End of tenancy clean", "One-off",
     "Quote sent", 395, "Q-0004", "", "Yes", "Chase the quote", 0,
     "", "", "Tenant leaves on the 28th. Needs it done that week."),
    (-14, "Beacon Coworking", "(503) 555 0699",
     "hello@beaconcowork.example",
     "14 Beacon Street, Portland OR 97204", "Commercial office",
     "Nextdoor", "Office clean - large", "Twice a week",
     "Following up", 1240, "Q-0005", -8, "Yes",
     "Second follow-up call", -1, "", "",
     "Biggest opportunity this quarter. Decision maker is Ana."),
    (-21, "Harriet Vance", "(503) 555 0701", "h.vance@example.com",
     "5 Willow Bank, Portland OR 97202", "Residential", "Facebook",
     "First deep clean", "One-off", "Won", 320, "Q-0003", -18, "Yes",
     "Booked in", "", "", "",
     "Became the Byrne-style trial - booked a deep clean first."),
    (-26, "Sunnyside Apartments", "(503) 555 0722",
     "manager@sunnyside.example",
     "Sunnyside Block C, Portland OR 97217", "Property manager",
     "Referral - staff", "Communal area clean", "Monthly", "Lost",
     320, "Q-0006", -22, "Yes", "", "", "", "Price too high",
     "Went with a cheaper outfit. Worth a call in six months."),
    (-32, "Dalton Residence", "(503) 555 0744",
     "dalton.home@example.com",
     "19 Poplar Drive, Portland OR 97203", "Residential",
     "Flyer / leaflet", "Regular home clean", "Monthly", "Lost", 110,
     "", "", "Yes", "", "", "", "Out of our area",
     "Forty minutes each way. Politely declined."),
    (-5, "Craig Mellor", "(503) 555 0766", "craig.mellor@example.com",
     "44 Hazel Road, Portland OR 97213", "Residential", "Instagram",
     "Oven clean", "One-off", "New enquiry", 85, "", "", "No",
     "Price it and text him back", 0, "", "",
     "Saw the before and after photos on Instagram."),
    (-11, "Northgate Physio", "(503) 555 0788",
     "clinic@northgatephysio.example",
     "60 Northgate Way, Portland OR 97211", "Medical & dental",
     "Directory listing", "Office clean - small", "Twice a week",
     "Negotiating", 620, "Q-0007", -7, "Yes",
     "They want 10% off for a 12 month contract", 2, "", "",
     "Worth doing at 5% if they sign for a year."),
]

# ---------------------------------------------------------------------------
# Quotes raised
# ---------------------------------------------------------------------------
# key, date offset, client no, prospect, property no, address, service,
# frequency, status, sent offset, followed up offset, decision offset,
# discount, job key, decline reason, notes, lines
QUOTES = [
    ("Q1", -34, 13, "", 15, "", "Regular home clean", "Fortnightly",
     "Accepted", -34, "", -31, 0.0, "", "",
     "Trial fortnightly visit after the deep clean.",
     [("Regular home clean", 1, None, 0.0, "No",
       "Three bed terrace, fortnightly visit."),
      ("Interior window clean", 1, None, 0.0, "Yes",
       "Optional - they took it on the first visit only.")]),
    ("Q2", -28, 9, "", 9, "", "Office clean - small", "Twice a week",
     "Accepted", -28, "", -24, 0.05, "", "",
     "Studio wanted mirrors and mats done properly.",
     [("Office clean - small", 2, None, 0.0, "No",
       "Two evening visits a week, mirrors and mats included.")]),
    ("Q3", -21, "", "Harriet Vance", "", "5 Willow Bank, Portland OR "
     "97202", "First deep clean", "One-off", "Accepted", -21, -19,
     -18, 0.0, "", "", "Deep clean first, then talk about regular.",
     [("First deep clean", 1, None, 0.0, "No",
       "Whole house deep clean, two cleaners, one day.")]),
    ("Q4", -9, "", "Tom Ridley", "", "77 Chestnut Avenue, Portland OR "
     "97206", "End of tenancy clean", "One-off", "Sent", -9, "", "",
     0.0, "", "",
     "Two bed flat, tenant moving out on the 28th.",
     [("End of tenancy clean", 1, None, 0.0, "No",
       "Full end of tenancy clean to inventory standard."),
      ("Oven clean", 1, None, 0.0, "No",
       "Oven, racks and extractor - dip tank treatment."),
      ("Carpet clean", 3, None, 0.0, "Yes",
       "Three rooms of carpet, hot water extraction.")]),
    ("Q5", -14, "", "Beacon Coworking", "", "14 Beacon Street, "
     "Portland OR 97204", "Office clean - large", "Twice a week",
     "Followed up", -14, -6, "", 0.0, "", "",
     "Coworking space, 4,000 sq ft, two evenings a week.",
     [("Office clean - large", 2, None, 0.0, "No",
       "Two evening visits a week across the whole floor."),
      ("Interior window clean", 1, None, 0.0, "Yes",
       "Monthly internal glass - optional extra.")]),
    ("Q6", -26, "", "Sunnyside Apartments", "", "Sunnyside Block C, "
     "Portland OR 97217", "Communal area clean", "Monthly",
     "Declined", -26, -20, -18, 0.0, "", "Price too high",
     "Lost on price. They took a 55 a visit offer.",
     [("Communal area clean", 4, None, 0.0, "No",
       "Four blocks, monthly visit each.")]),
    ("Q7", -11, "", "Northgate Physio", "", "60 Northgate Way, "
     "Portland OR 97211", "Office clean - small", "Twice a week",
     "Sent", -11, -4, "", 0.05, "", "",
     "They are pushing for 10%. Hold at 5% with a 12 month term.",
     [("Office clean - small", 2, None, 0.0, "No",
       "Two visits a week, clinical waste excluded.")]),
    ("Q8", -45, 5, "", 5, "", "End of tenancy clean", "One-off",
     "Expired", -45, -38, "", 0.0, "", "",
     "Agent never came back on this one - chase the branch.",
     [("End of tenancy clean", 1, 420.00, 0.0, "No",
       "Two bed apartment, end of tenancy to inventory standard.")]),
]

# ---------------------------------------------------------------------------
# Regular rounds: client, property, service, weekday, start, every n days
# ---------------------------------------------------------------------------
ROUND = [
    (1, 1, "Regular home clean", 1, (9, 0), 7, "Marta Kovacs", "", 1,
     "Weekly"),
    (2, 2, "Regular home clean", 3, (14, 0), 14, "Jade Ellis", "", 1,
     "Fortnightly"),
    (3, 3, "Office clean - small", 4, (18, 0), 7, "Tomas Ruiz", "", 1,
     "Weekly"),
    (4, 4, "Office clean - small", 0, (18, 30), 7, "Tomas Ruiz", "", 1,
     "Twice a week"),
    (4, 4, "Office clean - small", 3, (18, 30), 7, "Tomas Ruiz", "", 1,
     "Twice a week"),
    (6, 6, "Restaurant & kitchen clean", 6, (7, 0), 7, "Amara Diallo",
     "", 1, "Weekly"),
    (7, 7, "Nursery & school clean", 4, (17, 30), 7, "Amara Diallo",
     "Luke Bennett", 2, "Weekly"),
    (8, 8, "Regular home clean", 2, (9, 30), 28, "Jade Ellis", "", 1,
     "Monthly"),
    (9, 9, "Office clean - small", 0, (7, 0), 7, "Luke Bennett", "", 1,
     "Twice a week"),
    (9, 9, "Office clean - small", 3, (7, 0), 7, "Luke Bennett", "", 1,
     "Twice a week"),
    (10, 10, "Airbnb changeover", 5, (11, 0), 7, "Jade Ellis", "", 1,
     "Weekly"),
    (10, 11, "Airbnb changeover", 5, (14, 0), 7, "Jade Ellis", "", 1,
     "Weekly"),
    (12, 13, "Communal area clean", 2, (14, 0), 28, "Luke Bennett", "",
     1, "Monthly"),
    (12, 14, "Communal area clean", 2, (15, 30), 28, "Luke Bennett",
     "", 1, "Monthly"),
    (13, 15, "Regular home clean", 1, (14, 0), 14, "Marta Kovacs", "",
     1, "Fortnightly"),
    (14, 16, "Office clean - small", 4, (20, 0), 7, "Tomas Ruiz", "",
     1, "Weekly"),
]

#: client 11 is paused for the summer - jobs stop forty days ago
ROUND_STOPS = {(11, 12): -40}

# ---------------------------------------------------------------------------
# Jobs that are not part of a round
# ---------------------------------------------------------------------------
# key, day offset, client, property, service, type, start, status,
# employee, employee2, cleaners, hours override, price override,
# rating, feedback, notes
ONE_OFFS = [
    ("J_tenancy", -8, 5, 5, "End of tenancy clean", "One-off clean",
     (8, 30), "Completed", "Priya Raman", "Amara Diallo", 2, 6.5,
     395.00, 5, "Agent said it was the best handover all year.",
     "Oven was the worst we have seen. Photographed everything."),
    ("J_deep", -18, 13, 15, "First deep clean", "First deep clean",
     (9, 0), "Completed", "Priya Raman", "Marta Kovacs", 2, 6.0, None,
     5, "Lovely job, booked a fortnightly after this.",
     "Converted straight to a fortnightly visit."),
    ("J_oven", -10, 1, 1, "Oven clean", "One-off clean", (13, 0),
     "Completed", "Priya Raman", "", 1, None, None, 5, "", 
     "Added on at the end of the regular visit."),
    ("J_carpet", 6, 4, 4, "Carpet clean", "One-off clean", (18, 0),
     "Booked", "Priya Raman", "", 1, 3.0, None, "", "",
     "Four offices of carpet, quoted separately from the contract."),
    ("J_spring", 9, 8, 8, "Spring clean", "One-off clean", (9, 0),
     "Booked", "Marta Kovacs", "Jade Ellis", 2, 5.0, None, "", "",
     "Annual going-over. Blinds, windows and cupboards."),
    ("J_noaccess", -13, 8, 8, "Regular home clean", "Recurring visit",
     (9, 30), "No access", "Jade Ellis", "", 1, None, 0.00, "", "",
     "Nobody home and the lockbox was empty. Charged nothing, "
     "rebooked."),
    ("J_cancel", -4, 2, 2, "Regular home clean", "Recurring visit",
     (14, 0), "Cancelled", "Jade Ellis", "", 1, None, None, "", "",
     "Family away. Cancelled the night before, no charge."),
    ("J_reclean", -5, 6, 6, "Restaurant & kitchen clean",
     "Re-clean (free)", (7, 0), "Completed", "Marta Kovacs", "", 1,
     1.5, 0.00, 4, "Sorted, thank you for coming back.",
     "Free re-clean after the complaint about the kitchen floor."),
    ("J_quote", 2, "", "", "", "Quote visit", (17, 0), "Booked",
     "Dana Whitfield", "", 1, 0.5, 0.00, "", "",
     "Site visit at Marsh & Fletcher Dental - take the method "
     "statement."),
]

# ---------------------------------------------------------------------------
# Money going out
# ---------------------------------------------------------------------------
# category, type, description, supplier, amount, tax, method, receipt,
# recurring, deductible, day of month, months back
#: service name -> the hours one cleaner is expected to take, so the
#: demo timesheets can be built from the same figure the price uses
SERVICE_HOURS = {row[0]: row[3] for row in SERVICES}

MONTHLY_COSTS = [
    ("Business insurance", "Fixed monthly", "Public and employer "
     "liability", "Hiscox", 66.00, 0.0, "Direct debit", "Yes", "Yes",
     "Yes", 1),
    ("Vehicle insurance", "Fixed monthly", "Van 1 and van 2 fleet "
     "policy", "Admiral Business", 150.00, 0.0, "Direct debit", "Yes",
     "Yes", "Yes", 1),
    ("Software & apps", "Fixed monthly", "Accounting software and "
     "cloud storage", "Ledgerly", 32.00, 0.0, "Card", "Yes", "Yes",
     "Yes", 3),
    ("Phone & internet", "Fixed monthly", "Two business mobiles",
     "Telecom NW", 58.00, 0.0, "Direct debit", "Yes", "Yes", "Yes", 5),
    ("Marketing & advertising", "Fixed monthly", "Google Ads - local "
     "cleaning searches", "Google", 120.00, 0.0, "Card", "Yes", "Yes",
     "Yes", 7),
    ("Office & rent", "Fixed monthly", "Storage unit for stock and "
     "machines", "SpaceStore", 75.00, 0.0, "Direct debit", "Yes",
     "Yes", "Yes", 2),
    ("Accounting & legal", "Fixed monthly", "Bookkeeping",
     "Cedar & Co Accountants", 85.00, 0.0, "Bank transfer", "Yes",
     "Yes", "Yes", 12),
    ("Vehicle fuel", "Variable", "Van 1 fuel", "Shell", 110.00, 0.0,
     "Card", "Yes", "No", "Yes", 9),
    ("Vehicle fuel", "Variable", "Van 2 fuel", "Shell", 95.00, 0.0,
     "Card", "Yes", "No", "Yes", 16),
    ("Cleaning supplies", "Variable", "Monthly stock order",
     "CleanCo Supplies", 165.00, 0.0, "Card", "Yes", "No", "Yes", 4),
    ("Laundry", "Variable", "Changeover linen service",
     "Quay Laundry", 64.00, 0.0, "Bank transfer", "Yes", "No", "Yes",
     18),
    ("Subcontractors", "Job related", "Window cleaning - Chris Boyd",
     "Clearview Windows", 180.00, 0.0, "Bank transfer", "Yes", "No",
     "Yes", 21),
]

ONE_OFF_COSTS = [
    (-6, "Equipment repair", "One-off", "Steam cleaner boiler seal",
     "Vac Spares Online", 62.00, 0.0, "Card", "Yes", "No", "Yes"),
    (-12, "Cleaning supplies", "Variable", "Emergency degreaser run",
     "Wholesale Direct", 48.90, 0.0, "Cash", "No", "No", "Yes"),
    (-19, "Training & certification", "One-off",
     "COSHH refresher for two cleaners", "BICSc", 190.00, 0.0,
     "Bank transfer", "Yes", "No", "Yes"),
    (-24, "Uniforms", "One-off", "Six polo shirts and two fleeces",
     "WorkWear Direct", 142.00, 0.0, "Card", "Yes", "No", "Yes"),
    (-31, "Equipment purchase", "One-off", "Replacement mop system",
     "CleanCo Supplies", 98.00, 0.0, "Card", "Yes", "No", "Yes"),
    (-38, "Marketing & advertising", "One-off",
     "Van signage refresh", "SignShop PDX", 380.00, 0.0,
     "Bank transfer", "Yes", "No", "Yes"),
    (-44, "Vehicle maintenance", "One-off", "Van 1 service and MOT",
     "Hilltop Garage", 318.00, 0.0, "Card", "Yes", "No", "Yes"),
    (-2, "Bank & card fees", "Fixed monthly", "Card processing fees",
     "SumUp", 18.40, 0.0, "Direct debit", "No", "Yes", "Yes"),
    (-27, "Waste disposal", "One-off", "Trade waste from the end of "
     "tenancy job", "City Waste", 45.00, 0.0, "Card", "Yes", "No",
     "Yes"),
]

# ---------------------------------------------------------------------------
# When it went wrong
# ---------------------------------------------------------------------------
# day offset, client, job key, employee, type, severity, description,
# status, owner, action, resolution, reclean key, cost, resolved
# offset, kept, rating after, prevention
COMPLAINTS = [
    (-7, 6, "", "Amara Diallo", "Missed area", "Medium",
     "Kitchen floor behind the fryer was missed two weeks running.",
     "Resolved", "Dana Whitfield",
     "Went back the same week and did it free.", "Free re-clean",
     "J_reclean", 68.00, -5, "Yes", 4,
     "Added 'behind the fryer' to the cafe checklist as a required "
     "task."),
    (-16, 10, "", "Jade Ellis", "Quality below standard", "Low",
     "Host said the bathroom mirror was smeared on a changeover.",
     "Closed - no action", "Dana Whitfield",
     "Spoke to Jade, sent the photo. Guest never mentioned it.",
     "Apology", "", 0.00, -15, "Yes", 5,
     "Glass cloth now colour coded yellow so it is never the wrong "
     "one."),
    (-3, 12, "", "Luke Bennett", "Arrived late", "Low",
     "Agent rang - we arrived at 3pm instead of 2pm for Block A.",
     "Investigating", "Dana Whitfield",
     "Checking the rota - the two blocks are booked too close "
     "together.", "", "", 0.00, "", "", "",
     "Move Block B to 4pm so there is travel time between them."),
    (-42, 2, "", "Jade Ellis", "Breakage", "Medium",
     "Glass vase knocked off a windowsill while cleaning behind it.",
     "Resolved", "Dana Whitfield",
     "Replaced it the next day and apologised in person.",
     "Item replaced", "", 38.00, -40, "Yes", 5,
     "Ornaments are now moved to the table, not balanced on the "
     "sill."),
]

# ---------------------------------------------------------------------------
# Safety incidents
# ---------------------------------------------------------------------------
# day offset, time, type, employee, job key, property, severity,
# description, injury, first aid, medical, damage, cost, insurance,
# reported to, action, prevention, status, closed offset
INCIDENTS = [
    (-23, (10, 15), "Slip, trip or fall", "Jade Ellis", "", 8, "Low",
     "Slipped on a wet kitchen floor that had just been mopped - no "
     "wet floor sign out.", "Bruised elbow", "Yes", "No", "No", 0.00,
     "No", "Dana Whitfield",
     "First aid applied, finished the shift.",
     "Wet floor signs now in both vans and on every checklist.",
     "Closed", -22),
    (-61, (16, 40), "Chemical splash", "Priya Raman", "", 5, "Medium",
     "Oven gel splashed onto the forearm while lifting a rack out of "
     "the dip tank.", "Minor chemical burn", "Yes", "No", "No", 0.00,
     "No", "Dana Whitfield",
     "Rinsed for fifteen minutes, checked at the pharmacy.",
     "Long gauntlets now mandatory for all dip tank work - written "
     "into the method statement.", "Closed", -58),
    (-11, (19, 5), "Near miss", "Tomas Ruiz", "", 4, "Low",
     "Trailing vacuum cable across a doorway in the office - nobody "
     "tripped, but somebody could have.", "None", "No", "No", "No",
     0.00, "No", "Dana Whitfield",
     "Reminded the team to work from the nearest socket.",
     "Cable routing added to the office checklist.", "Closed", -10),
]


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------

_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg)
    return _CACHE[key]


def _month_key(day: _dt.date) -> str:
    return day.strftime("%B %Y")


def _build(cfg: Config) -> Sample:
    rnd = random.Random(20261018)
    today = cfg.today
    monday = week_monday(today)
    s = Sample()

    # ------------------------------------------------------------ clients --
    for (name, kind, contact, phone, email, address, status, source,
         since, freq, day, slot, cleaner, rate, terms, contract, vip,
         notes) in CLIENTS:
        s.clients.append({
            "name": name, "type": kind, "contact": contact,
            "phone": phone, "email": email, "address": address,
            "status": status, "source": source,
            "since": d(today, since) if since != "" else d(today, -30),
            "frequency": freq, "pref_day": day, "pref_time": slot,
            "pref_cleaner": cleaner, "rate": rate, "terms": terms,
            "contract": contract, "vip": vip, "marketing": "Yes",
            "notes": notes,
        })

    # --------------------------------------------------------- properties --
    for index, (client, name, kind, address, beds, baths, rooms, area,
                condition, access, key, parking, pets, floors, hours,
                notes) in enumerate(PROPERTIES, 1):
        distance, checklist, onsite = PROPERTY_EXTRA.get(
            index, (5.0, "Regular home clean", "No"))
        s.properties.append({
            "client": cid(client), "name": name, "type": kind,
            "address": address, "beds": beds, "baths": baths,
            "rooms": rooms, "area": area, "condition": condition,
            "access": access, "key_status": key, "parking": parking,
            "pets": pets, "floors": floors, "hours": hours,
            "distance": distance, "checklist": checklist,
            "supplies_onsite": onsite, "active": "Yes",
            "entry_notes": ENTRY_NOTES.get(access, ""),
            "notes": notes,
        })

    # ----------------------------------------------------------- services --
    for (name, category, model, hours, price, unit, uplift, supplies,
         popular, description) in SERVICES:
        s.services.append({
            "name": name, "category": category, "model": model,
            "hours": hours, "price": price, "unit": unit,
            "uplift": uplift, "supplies": supplies, "popular": popular,
            "active": "Yes", "description": description,
        })

    # ---------------------------------------------------------- employees --
    for (name, role, kind, phone, email, start, pay, hours, status,
         skills, checked, notes) in EMPLOYEES:
        s.employees.append({
            "name": name, "role": role, "type": kind, "phone": phone,
            "email": email, "start": d(today, start), "pay": pay,
            "hours": hours, "status": status, "skills": skills,
            "checked": checked, "notes": notes,
            "licence": "Yes" if role != "Cleaner" else "No",
            "insured": "Yes" if kind == "Subcontractor" else "N/A",
        })

    # ----------------------------------------------------------- the diary --
    job_index: Dict[str, int] = {}
    rows: List[Dict[str, Any]] = []

    def add_job(**row) -> int:
        rows.append(row)
        return len(rows)

    first_day = monday - _dt.timedelta(days=35)
    last_day = monday + _dt.timedelta(days=27)

    for (client, prop, service, weekday, start, step, cleaner, mate,
         crew, frequency) in ROUND:
        day = first_day + _dt.timedelta(
            days=(weekday - first_day.weekday()) % 7)
        stop = ROUND_STOPS.get((client, prop))
        while day <= last_day:
            if stop is not None and day > d(today, stop):
                break
            if day < today:
                status = "Completed"
            elif day == today:
                status = "In progress"
            elif (day - today).days <= 5:
                status = "Confirmed"
            else:
                status = "Booked"
            rating = ""
            feedback = ""
            actual = None
            if status == "Completed" and rnd.random() < 0.45:
                rating = rnd.choice([5, 5, 5, 4])
            if status == "Completed" and rnd.random() < 0.6:
                actual = None
            distance = PROPERTY_EXTRA.get(prop, (5.0,))[0]
            add_job(
                date=day, start=_dt.time(*start), client=cid(client),
                property=pid(prop), service=service,
                type="Recurring visit", frequency=frequency,
                status=status, priority="Normal", employee=cleaner,
                employee2=mate, cleaners=crew, hours_input=None,
                price_input=None, travel=round(distance * 2, 1),
                recurring="Yes",
                confirmed="Yes" if status in ("Completed", "Confirmed",
                                              "In progress") else "No",
                actual_input=actual, rating=rating, feedback=feedback,
                notes="")
            day += _dt.timedelta(days=step)

    for (tag, offset, client, prop, service, kind, start, status,
         cleaner, mate, crew, hours, price, rating, feedback,
         notes) in ONE_OFFS:
        day = d(today, offset)
        distance = PROPERTY_EXTRA.get(prop, (4.0,))[0] if prop else 3.0
        n = add_job(
            date=day, start=_dt.time(*start),
            client=cid(client) if client else "",
            property=pid(prop) if prop else "", service=service,
            type=kind,
            frequency="One-off" if kind != "Recurring visit"
            else "Weekly",
            status=status,
            priority="High" if kind == "One-off clean" else "Normal",
            employee=cleaner, employee2=mate, cleaners=crew,
            hours_input=hours, price_input=price,
            travel=round(distance * 2, 1), recurring="No",
            confirmed="Yes" if status != "Booked" else "No",
            actual_input=None, rating=rating, feedback=feedback,
            notes=notes)
        job_index[tag] = n

    rows.sort(key=lambda r: (r["date"], str(r["start"])))
    s.jobs = rows
    # the sort moved everything, so rebuild the tag index
    job_index = {}
    for tag, offset, client, prop, service, kind, start, *_ in ONE_OFFS:
        day = d(today, offset)
        for i, row in enumerate(rows, 1):
            if (row["date"] == day and row["start"] == _dt.time(*start)
                    and row["service"] == service
                    and row["type"] == kind):
                job_index[tag] = i
                break

    def job_no(tag: str) -> str:
        n = job_index.get(tag)
        return jid(n) if n else ""

    # -------------------------------------------------------- timesheets --
    for n, row in enumerate(rows, 1):
        if row["status"] not in ("Completed", "In progress"):
            continue
        if not row["employee"]:
            continue
        if row["date"] < monday - _dt.timedelta(days=56):
            # every finished visit in the history gets its hours, so
            # the payroll, the utilisation and the job costing all
            # cover the same period
            continue
        # the hours a visit is planned at - the same number the Jobs
        # tab prices from, so the Timesheets tab cannot quietly
        # contradict it. Real visits run a little over or under.
        prop_hours = 2.5
        if row["property"]:
            prop_hours = PROPERTIES[int(row["property"][3:]) - 1][14]
        planned = (row["hours_input"] or SERVICE_HOURS.get(row["service"])
                   or prop_hours)
        planned = planned * rnd.uniform(0.92, 1.14)
        hours = round(planned / max(1, row["cleaners"]), 2)
        start = row["start"]
        begin = _dt.datetime.combine(row["date"], start)
        wobble = rnd.choice([0, 0, 0, 15, -15, 30])
        finish = begin + _dt.timedelta(hours=hours, minutes=wobble)
        late = rnd.random() < 0.08
        week_old = (today - row["date"]).days
        s.timesheets.append({
            "date": row["date"], "employee": row["employee"],
            "job": jid(n), "activity": "", "start": start,
            "finish": finish.time(),
            "break_min": 30 if hours >= 5 else 0,
            "hours_input": None, "overtime": None,
            "attendance": "Late" if late else "On time",
            "billable": "Yes", "approved": "Yes" if week_old > 7
            else "No", "paid": "Yes" if week_old > 14 else "No",
            "notes": "Traffic on the bridge." if late else "",
        })
        if row["employee2"]:
            s.timesheets.append({
                "date": row["date"], "employee": row["employee2"],
                "job": jid(n), "activity": "", "start": start,
                "finish": finish.time(), "break_min": 0,
                "hours_input": None, "overtime": None,
                "attendance": "On time", "billable": "Yes",
                "approved": "Yes" if week_old > 7 else "No",
                "paid": "Yes" if week_old > 14 else "No", "notes": "",
            })

    # a few non-cleaning rows so the utilisation figure is honest
    for offset, who, activity, start, hours, note in [
            (-9, "Dana Whitfield", "Quote visit", (17, 0), 1.5,
             "Two quote visits and the drive between them."),
            (-6, "Marta Kovacs", "Training", (9, 0), 3.0,
             "COSHH refresher with Amara."),
            (-4, "Luke Bennett", "Supply run", (8, 0), 1.0,
             "Stock collection from CleanCo."),
            (-2, "Dana Whitfield", "Admin errand", (14, 0), 2.0,
             "Invoicing and chasing two overdue accounts.")]:
        begin = _dt.datetime.combine(d(today, offset), _dt.time(*start))
        s.timesheets.append({
            "date": d(today, offset), "employee": who,
            "job": "", "activity": activity, "start": _dt.time(*start),
            "finish": (begin + _dt.timedelta(hours=hours)).time(),
            "break_min": 0, "hours_input": None, "overtime": None,
            "attendance": "On time", "billable": "No",
            "approved": "Yes", "paid": "No", "notes": note,
        })
    s.timesheets.append({
        "date": d(today, -13), "employee": "Amara Diallo", "job": "",
        "activity": "", "start": None, "finish": None,
        "break_min": 0, "hours_input": 0, "overtime": None,
        "attendance": "Sick", "billable": "No", "approved": "Yes",
        "paid": "Yes", "notes": "Covered by Marta at the nursery.",
    })

    # ----------------------------------------------------------- invoices --
    groups: Dict[Any, List[int]] = {}
    for n, row in enumerate(rows, 1):
        if row["status"] != "Completed":
            continue
        if row["price_input"] == 0:
            continue                      # free re-clean, nothing to bill
        if (today - row["date"]).days < 4:
            continue                      # done, not billed yet
        key = (row["client"], row["date"].year, row["date"].month)
        groups.setdefault(key, []).append(n)

    ordered = sorted(groups.items(),
                     key=lambda kv: (kv[0][1], kv[0][2], kv[0][0]))
    overdue_slots = {2, 5}
    part_paid_slot = 8
    for slot, ((client, year, month), job_nos) in enumerate(ordered):
        last_job = max(rows[n - 1]["date"] for n in job_nos)
        raised = min(last_job + _dt.timedelta(days=1), today)
        age = (today - raised).days
        client_no = int(client[3:])
        terms = CLIENTS[client_no - 1][14]
        commercial = CLIENTS[client_no - 1][1] != "Residential"
        if age > terms + 7 and slot in overdue_slots:
            status = "Overdue"
        elif slot == part_paid_slot and age > 5:
            status = "Part paid"
        elif age >= terms:
            status = "Paid"
        elif age >= 1:
            status = "Sent"
        else:
            status = "Draft"
        invoice_no = len(s.invoices) + 1
        s.invoices.append({
            "date": raised, "client": client,
            "period": _dt.date(year, month, 1).strftime("%B %Y"),
            "status": status, "sent_date": raised if status != "Draft"
            else None, "terms_input": None, "discount": None,
            "method": "Bank transfer" if commercial else "Card",
            "po": f"PO-{4000 + slot}" if commercial else "",
            "reminders": 2 if status == "Overdue" else
            (1 if status == "Part paid" else None),
            "last_reminder": d(today, -3) if status == "Overdue"
            else None,
            "notes": "Monthly account." if commercial else "",
        })
        for n in job_nos:
            s.invoice_lines.append({
                "invoice": iid(invoice_no), "job": jid(n),
                "description": "", "qty": 1, "price_input": None,
                "taxable": "Yes", "notes": "",
            })
            rows[n - 1]["status"] = "Invoiced"
        # ------------------------------------------------------ payments --
        if status in ("Paid", "Part paid"):
            paid_on = min(raised + _dt.timedelta(
                days=max(3, terms - rnd.choice([0, 1, 2, 5]))), today)
            s.payments.append({
                "date": paid_on, "invoice": iid(invoice_no),
                "amount": None, "fee": None,
                "method": "Bank transfer" if commercial else "Card",
                "reference": f"BT-{8800 + invoice_no}" if commercial
                else f"SUMUP-{3300 + invoice_no}",
                "deposit": "No", "banked": "Yes",
                "notes": "" if status == "Paid"
                else "Part payment - balance promised next week.",
            })

    # the payment amounts: full for Paid, 60% for Part paid
    for payment in s.payments:
        invoice_no = int(payment["invoice"][4:])
        invoice = s.invoices[invoice_no - 1]
        total = 0.0
        for line in s.invoice_lines:
            if line["invoice"] != payment["invoice"]:
                continue
            job = rows[int(line["job"][4:]) - 1]
            total += _job_price(cfg, job)
        if invoice["status"] == "Part paid":
            payment["amount"] = round(total * 0.6, 2)
        else:
            payment["amount"] = round(total, 2)
        if payment["method"] == "Card":
            payment["fee"] = round(payment["amount"] * 0.0175, 2)

    # a deposit on the end of tenancy job, taken before the work
    if job_no("J_tenancy"):
        s.payments.append({
            "date": d(today, -12), "invoice": "", "amount": 100.00,
            "fee": 1.75, "method": "Card", "reference": "SUMUP-3290",
            "deposit": "Yes", "banked": "Yes",
            "notes": "Deposit to hold the date for the end of tenancy "
                     "clean.",
        })

    # ------------------------------------------------------------- quotes --
    quote_index = {tag: n for n, (tag, *_) in enumerate(QUOTES, 1)}
    for (tag, offset, client, prospect, prop, address, service,
         frequency, status, sent, followed, decided, discount, job_tag,
         decline, notes, lines) in QUOTES:
        s.quotes.append({
            "date": d(today, offset),
            "client": cid(client) if client else "",
            "prospect": prospect,
            "property": pid(prop) if prop else "",
            "address": address, "service": service,
            "frequency": frequency, "status": status,
            "sent_date": d(today, sent) if sent != "" else None,
            "followed_up": d(today, followed) if followed != ""
            else None,
            "decision_date": d(today, decided) if decided != ""
            else None,
            "discount": discount or None, "deposit_input": None,
            "valid_input": None, "job": job_no(job_tag) if job_tag
            else "", "decline_reason": decline, "notes": notes,
            "terms": "",
        })
        for (line_service, qty, price, line_discount, optional,
             wording) in lines:
            s.quote_lines.append({
                "quote": qid(quote_index[tag]), "service": line_service,
                "description": wording, "qty": qty,
                "price_input": price, "discount": line_discount or None,
                "optional": optional, "notes": "",
            })

    # -------------------------------------------------------------- leads --
    for (offset, name, phone, email, address, kind, source, service,
         frequency, stage, value, quote, visit, contacted, action,
         next_offset, client, lost, notes) in LEADS:
        s.leads.append({
            "date": d(today, offset), "name": name, "phone": phone,
            "email": email, "address": address, "type": kind,
            "source": source, "service": service,
            "frequency": frequency, "stage": stage, "value": value,
            "quote": quote, "visit_date": d(today, visit)
            if visit != "" else None, "contacted": contacted,
            "next_action": action,
            "next_date": d(today, next_offset) if next_offset != ""
            else None,
            "client": cid(client) if client else "",
            "lost_reason": lost, "notes": notes,
        })

    # --------------------------------------------------------- checklists --
    for service, tasks in CHECKLIST_TASKS.items():
        for order, (area, task, detail, minutes, frequency, required,
                    photo) in enumerate(tasks, 1):
            s.checklists.append({
                "template": service, "property": "", "area": area,
                "order": order, "task": task, "detail": detail,
                "minutes": minutes, "frequency": frequency,
                "required": required, "photo": photo,
                "products": "", "active": "Yes", "notes": "",
            })
    # one property specific extra
    s.checklists.append({
        "template": "Regular home clean", "property": pid(1),
        "area": "Whole property", "order": 17,
        "task": "Put the recycling box out on a Tuesday",
        "detail": "Box at the end of the drive, lid weighted down",
        "minutes": 3, "frequency": "Weekly", "required": "Yes",
        "photo": "No", "products": "", "active": "Yes",
        "notes": "Only at the Harper house - they asked specially.",
    })

    # --------------------------------------------------------- job tasks --
    recent_home = [n for n, row in enumerate(rows, 1)
                   if row["service"] == "Regular home clean"
                   and row["status"] in ("Completed", "Invoiced")
                   and (today - row["date"]).days <= 9]
    for n in recent_home[:2]:
        row = rows[n - 1]
        for order, (area, task, detail, minutes, frequency, required,
                    photo) in enumerate(
                        CHECKLIST_TASKS["Regular home clean"], 1):
            roll = rnd.random()
            status = "Done"
            quality = "Good"
            issue = ""
            reported = "No"
            if roll > 0.93:
                status = "Issue found"
                quality = "Acceptable"
                issue = "Limescale is building up again - needs a "\
                        "descale next visit."
                reported = "Yes"
            elif roll > 0.88:
                status = "Not needed"
                quality = ""
            s.job_tasks.append({
                "job": jid(n), "area": area, "task": task,
                "status": status, "employee": row["employee"],
                "minutes": minutes + rnd.choice([-2, 0, 0, 0, 3]),
                "quality": quality, "issue": issue,
                "reported": reported, "photo": "", "notes": "",
            })
    office = [n for n, row in enumerate(rows, 1)
              if row["service"] == "Office clean - small"
              and row["status"] in ("Completed", "Invoiced")
              and (today - row["date"]).days <= 6]
    for n in office[:1]:
        row = rows[n - 1]
        for area, task, detail, minutes, frequency, required, photo in \
                CHECKLIST_TASKS["Office clean - small"]:
            s.job_tasks.append({
                "job": jid(n), "area": area, "task": task,
                "status": "Done", "employee": row["employee"],
                "minutes": minutes, "quality": "Excellent",
                "issue": "", "reported": "No", "photo": "",
                "notes": "",
            })

    # ---------------------------------------------------------- supplies --
    for (name, category, supplier, unit, pack_size, pack_cost, stock,
         reorder, order_qty, eco, coshh, notes) in SUPPLIES:
        s.supplies.append({
            "name": name, "category": category, "supplier": supplier,
            "unit": unit, "pack_size": pack_size,
            "pack_cost": pack_cost, "stock": stock,
            "count_date": d(today, -14), "reorder": reorder,
            "order_qty": order_qty,
            "last_order": d(today, -rnd.choice([12, 19, 26, 33])),
            "eco": eco, "coshh": coshh, "notes": notes,
        })

    # ------------------------------------------------------ supply usage --
    usable = [n for n, row in enumerate(rows, 1)
              if row["status"] in ("Completed", "Invoiced")
              and (today - row["date"]).days <= 28]
    picks = ["Multi-surface cleaner", "Microfibre cloths - colour coded",
             "Bin bags - heavy duty", "Toilet cleaner", "Glass cleaner",
             "Nitrile gloves", "Bathroom descaler",
             "Toilet rolls - commercial", "Hand soap refill",
             "Kitchen degreaser"]
    for i, n in enumerate(usable[::3][:22]):
        row = rows[n - 1]
        item = picks[i % len(picks)]
        s.supply_usage.append({
            "date": row["date"], "supply": item,
            "qty": rnd.choice([0.25, 0.5, 1, 1, 2]),
            "job": jid(n), "employee": row["employee"],
            "reason": "Used on the job", "notes": "",
        })
    s.supply_usage.append({
        "date": d(today, -6), "supply": "Oven cleaning gel", "qty": 1,
        "job": job_no("J_tenancy"), "employee": "Priya Raman",
        "reason": "Used on the job",
        "notes": "Whole bottle on that oven. Charged for in the price.",
    })
    s.supply_usage.append({
        "date": d(today, -9), "supply": "Laundry detergent", "qty": 0.5,
        "job": "", "employee": "Jade Ellis", "reason": "Stock take",
        "notes": "Changeover linen for the week.",
    })

    # --------------------------------------------------------- equipment --
    for (name, kind, make, serial, ownership, status, assigned,
         location, bought, cost, life, warranty, last_service,
         interval, service_cost, safety_test, insured,
         notes) in EQUIPMENT:
        s.equipment.append({
            "name": name, "type": kind, "make": make, "serial": serial,
            "ownership": ownership, "status": status,
            "assigned": assigned, "location": location,
            "bought": d(today, bought), "cost": cost,
            "life_years": life,
            "warranty": d(today, warranty) if warranty != "" else None,
            "last_service": d(today, last_service)
            if last_service != "" else None,
            "interval_input": interval, "service_cost": service_cost,
            "safety_test": d(today, safety_test) if safety_test != ""
            else None, "insured": insured, "notes": notes,
        })

    # ---------------------------------------------------------- expenses --
    for back in (2, 1, 0):
        month_day = _dt.date(today.year, today.month, 1)
        for _ in range(back):
            month_day = (month_day - _dt.timedelta(days=1)).replace(day=1)
        for (category, kind, description, supplier, amount, tax,
             method, receipt, recurring, deductible,
             day_of_month) in MONTHLY_COSTS:
            try:
                when = month_day.replace(day=day_of_month)
            except ValueError:
                continue
            if when > today:
                continue
            wobble = 1.0
            if kind == "Variable":
                wobble = rnd.uniform(0.85, 1.2)
            s.expenses.append({
                "date": when, "category": category, "type": kind,
                "description": description, "supplier": supplier,
                "amount": round(amount * wobble, 2), "tax": tax or None,
                "method": method, "job": "", "employee": "",
                "receipt": receipt, "recurring": recurring,
                "deductible": deductible, "notes": "",
            })
    for (offset, category, kind, description, supplier, amount, tax,
         method, receipt, recurring, deductible) in ONE_OFF_COSTS:
        s.expenses.append({
            "date": d(today, offset), "category": category,
            "type": kind, "description": description,
            "supplier": supplier, "amount": amount, "tax": tax or None,
            "method": method, "job": "", "employee": "",
            "receipt": receipt, "recurring": recurring,
            "deductible": deductible, "notes": "",
        })
    # wages, the biggest line of all
    # Wages go out every Friday, and each run is sized from the hours
    # that were really worked that week - so the books and the
    # Timesheets tab tell the same story instead of two different
    # ones.
    cost_an_hour = cfg.pay_rate * (1 + cfg.employer_oncost)
    worked: dict = {}
    for sheet_row in s.timesheets:
        start, finish = sheet_row["start"], sheet_row["finish"]
        if start is None or finish is None:
            # a row clocked by hand, with the hours typed straight in
            typed = sheet_row.get("hours_input") or 0
            if typed:
                worked[sheet_row["date"]] = (
                    worked.get(sheet_row["date"], 0.0) + float(typed))
            continue
        minutes = ((finish.hour * 60 + finish.minute)
                   - (start.hour * 60 + start.minute)
                   - (sheet_row.get("break_min") or 0))
        if minutes > 0:
            worked[sheet_row["date"]] = (worked.get(sheet_row["date"], 0.0)
                                         + minutes / 60.0)
    if worked:
        # the Friday on or before today, then back a fortnight at a time
        payday = today - _dt.timedelta(days=(today.weekday() - 4) % 7)
        earliest = min(worked)
        while payday > earliest:
            opened = payday - _dt.timedelta(days=7)
            hours = sum(value for day, value in worked.items()
                        if opened <= day < payday)
            if hours > 0:
                s.expenses.append({
                    "date": payday, "category": "Wages & payroll",
                    "type": "Fixed monthly",
                    "description": "Weekly payroll run",
                    "supplier": "Payroll",
                    "amount": round(hours * cost_an_hour, 2),
                    "tax": None, "method": "Bank transfer", "job": "",
                    "employee": "", "receipt": "Yes",
                    "recurring": "Yes", "deductible": "Yes",
                    "notes": f"{hours:,.1f} cleaner hours from the "
                             f"Timesheets tab, plus employer "
                             f"on-costs.",
                })
            payday = opened

    # ----------------------------------------------------------- mileage --
    for i, n in enumerate(usable[::4][:16]):
        row = rows[n - 1]
        prop_no = int(row["property"][3:]) if row["property"] else 1
        distance = PROPERTY_EXTRA.get(prop_no, (5.0,))[0]
        s.mileage.append({
            "date": row["date"], "employee": row["employee"],
            "vehicle": "Van 1" if i % 2 == 0 else "Van 2",
            "purpose": "Travel to job", "from_place": "Base",
            "to_place": PROPERTIES[prop_no - 1][1], "job": jid(n),
            "distance": distance, "return_trip": "Yes",
            "rate_input": None, "parking": 0 if i % 5 else 4.50,
            "claimed": "Yes" if (today - row["date"]).days > 14
            else "No", "notes": "",
        })
    for offset, who, purpose, frm, to, distance, parking, note in [
            (-9, "Dana Whitfield", "Quote visit", "Base",
             "Marsh & Fletcher Dental", 7.4, 3.00,
             "Two quotes in one trip."),
            (-4, "Luke Bennett", "Supply run", "Base",
             "CleanCo Supplies", 11.2, 0, ""),
            (-16, "Dana Whitfield", "Bank & post", "Base",
             "Bank and post office", 3.8, 2.00, ""),
            (-6, "Marta Kovacs", "Training", "Base",
             "BICSc training centre", 14.6, 0, "COSHH refresher.")]:
        s.mileage.append({
            "date": d(today, offset), "employee": who,
            "vehicle": "Car", "purpose": purpose, "from_place": frm,
            "to_place": to, "job": "", "distance": distance,
            "return_trip": "Yes", "rate_input": None,
            "parking": parking or None, "claimed": "Yes",
            "notes": note,
        })

    # -------------------------------------------------------- complaints --
    for (offset, client, job_tag, employee, kind, severity,
         description, status, owner, action, resolution, reclean,
         cost, resolved, kept, rating_after, prevent) in COMPLAINTS:
        s.complaints.append({
            "date": d(today, offset), "client": cid(client),
            "job": job_no(job_tag) if job_tag else "",
            "employee": employee, "type": kind, "severity": severity,
            "description": description, "status": status,
            "owner": owner, "action": action, "resolution": resolution,
            "reclean_job": job_no(reclean) if reclean else "",
            "cost": cost or None,
            "resolved_date": d(today, resolved) if resolved != ""
            else None, "kept": kept, "rating_after": rating_after
            or None, "prevent": prevent, "notes": "",
        })

    # --------------------------------------------------------- incidents --
    for (offset, when, kind, employee, job_tag, prop, severity,
         description, injury, first_aid, medical, damage, cost,
         insurance, reported_to, action, prevent, status,
         closed) in INCIDENTS:
        s.incidents.append({
            "date": d(today, offset), "time": _dt.time(*when),
            "type": kind, "employee": employee,
            "job": job_no(job_tag) if job_tag else "",
            "property": pid(prop) if prop else "", "severity": severity,
            "description": description, "injury": injury,
            "first_aid": first_aid, "medical": medical,
            "damage": damage, "cost": cost or None,
            "insurance": insurance, "reported_to": reported_to,
            "reported_date": d(today, offset), "action": action,
            "prevent": prevent, "status": status,
            "closed_date": d(today, closed) if closed != "" else None,
            "notes": "",
        })

    # ------------------------------------------------------------ safety --
    for (document, applies, employee, reference, issuer, issued,
         expires, cost, holder, required, reviewed,
         notes) in SAFETY_DOCS:
        s.safety.append({
            "document": document, "applies_to": applies,
            "employee": employee, "reference": reference,
            "issuer": issuer, "issued": d(today, issued),
            "expires": d(today, expires) if expires != "" else None,
            "cost": cost or None, "holder": holder, "link": "",
            "required": required,
            "reviewed": d(today, reviewed) if reviewed != "" else None,
            "notes": notes,
        })

    return s


def _job_price(cfg: Config, job: Dict[str, Any]) -> float:
    """Mirror the Jobs price formula closely enough for the demo totals."""
    if job.get("price_input") is not None:
        return float(job["price_input"])
    service = job.get("service", "")
    prop = job.get("property", "")
    hours = job.get("hours_input")
    model = price = svc_hours = None
    for row in SERVICES:
        if row[0] == service:
            model, svc_hours, price = row[2], row[3], row[4]
            break
    if hours is None:
        if prop:
            hours = PROPERTIES[int(prop[3:]) - 1][14]
        else:
            hours = svc_hours or 2.0
        if model and model != "Per hour" and svc_hours:
            hours = svc_hours
    if model == "Flat price per visit" and price:
        base = float(price)
    else:
        client_no = int(job["client"][3:]) if job.get("client") else 0
        rate = CLIENTS[client_no - 1][13] if client_no else 45
        base = float(hours) * float(rate)
    factor = {
        "Weekly": 0.85, "Fortnightly": 0.90, "Every 3 weeks": 0.95,
        "Monthly": 1.00, "Quarterly": 1.05, "Twice a week": 0.80,
    }.get(job.get("frequency", "One-off"), 1.00)
    return round(max(cfg.minimum_charge, base * factor), 2)
````

---

## 3. The work registers

One module per data tab. Each returns a `SheetSpec` - a list of columns with their formulas - and the writer turns that into a formatted, validated, protected sheet. These are the chain that turns an enquiry into a priced job and a paid invoice.

### `cleaning_business_spreadsheet/sheets/__init__.py`

*6 lines*

> One module per tab.
>
> Register sheets expose ``NAME`` and ``spec(cfg)`` (plus an optional
> ``decorate``); page sheets expose ``NAME`` and ``build(book, fmts, cfg,
> reg)``. `build.py` wires them together.

````python
"""One module per tab.

Register sheets expose ``NAME`` and ``spec(cfg)`` (plus an optional
``decorate``); page sheets expose ``NAME`` and ``build(book, fmts, cfg,
reg)``. `build.py` wires them together.
"""
````

### `cleaning_business_spreadsheet/sheets/common.py`

*267 lines · 24 functions*

> Small building blocks shared by the register sheets.

| Defines | What it does |
|---------|--------------|
| `id_col(trigger, prefix, digits, width, header, note)` | Auto numbered ID column (blank until the trigger column is filled) |
| `month_col(date_key, key, header, width)` | `yyyy-mm` text column - the key every monthly roll-up groups on |
| `weekday_col(date_key, key, header, width)` |  |
| `week_col(date_key, key, header, width)` | Week number, honouring the week start day on Settings |
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
| `lookup_client(target, key, missing)` | Pull a column of the Clients tab for the client ID in `key` |
| `lookup_property(target, key, missing)` | Pull a column of the Properties tab for the property ID in `key` |
| `lookup_service(target, key, missing)` | Pull a column of the Services catalogue for the service in `key` |
| `lookup_job(target, key, missing)` | Pull a column of the Jobs tab for the job ID in `key` |
| `lookup_employee(target, key, missing)` | Pull a column of the Employees tab for the name in `key` |
| `rank_desc(key, guard, header, width, note)` | Hidden dense rank, 1 = the biggest value in `key` |
| `rank_asc(key, guard, header, width, note)` | Hidden dense rank, 1 = the smallest value in `key` |
| `min_number(value_range, cond, cap)` | Smallest positive number in `value_range` where `cond` holds |

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
    """Week number, honouring the week start day on Settings."""
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'WEEKNUM({{{date_key}}}{{r}},'
                        f'IF(WeekStart="Sunday",1,2)))'),
               note="Week number of the year, using the week start day set "
                    "on the Settings tab.")


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
        cf_text("Inactive", "muted"),
        cf_text("Archived", "muted"),
        cf_text("Not needed", "muted"),
    ]


def rate_cf(target_name: str = "TargetMargin") -> List[Dict[str, Any]]:
    """Green at or above the target, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def bar_cf(color: str = "#2F8F9D") -> List[Dict[str, Any]]:
    return [{"type": "data_bar", "bar_color": color}]


def scale_cf() -> List[Dict[str, Any]]:
    return [{"type": "3_color_scale", "min_color": "#FBD9DD",
             "mid_color": "#FFF0CE", "max_color": "#D7F2E3"}]


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


def lookup_client(target: str, key: str = "client",
                  missing: str = "Not found") -> str:
    """Pull a column of the Clients tab for the client ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Clients.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Clients.id]],0)),"{missing}"))')


def lookup_property(target: str, key: str = "property",
                    missing: str = "Not found") -> str:
    """Pull a column of the Properties tab for the property ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Properties.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Properties.id]],0)),"{missing}"))')


def lookup_service(target: str, key: str = "service",
                   missing: str = "Add to Services") -> str:
    """Pull a column of the Services catalogue for the service in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Services.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Services.name]],0)),"{missing}"))')


def lookup_job(target: str, key: str = "job", missing: str = "") -> str:
    """Pull a column of the Jobs tab for the job ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Jobs.'
            f'{target}]],MATCH({{{key}}}{{r}},[[Jobs.id]],0)),"{missing}"))')


def lookup_employee(target: str, key: str = "employee",
                    missing: str = "") -> str:
    """Pull a column of the Employees tab for the name in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Employees.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Employees.name]],0)),"{missing}"))')


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
````

### `cleaning_business_spreadsheet/sheets/lists.py`

*386 lines · 1 function*

> `Lists` - every drop down in the workbook, plus the pricing factor tables.
>
> Each list becomes a workbook level defined name, so a column can point at
> it with ``validate=list_from("JobStatusList")``. Edit a list here and
> every drop down that uses it follows.
>
> The two little tables at the bottom are the engine behind the quote
> calculator: one turns a cleaning frequency into a price factor and a
> number of days between visits, the other turns "how dirty is it" into a
> time factor. Change a number there and every quote, every recurring
> booking and every price on the workbook re-prices itself.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Lists'), `LAST_COL` (9), `LISTS` (61 items), `FREQUENCY_TABLE` (7 items), `CONDITION_TABLE` (list(...)).

````python
"""`Lists` - every drop down in the workbook, plus the pricing factor tables.

Each list becomes a workbook level defined name, so a column can point at
it with ``validate=list_from("JobStatusList")``. Edit a list here and
every drop down that uses it follows.

The two little tables at the bottom are the engine behind the quote
calculator: one turns a cleaning frequency into a price factor and a
number of days between visits, the other turns "how dirty is it" into a
time factor. Change a number there and every quote, every recurring
booking and every price on the workbook re-prices itself.
"""

from __future__ import annotations

from typing import List, Tuple

from ..config import (CONDITION_FACTORS, FREQUENCIES, FREQUENCY_FACTORS,
                      INVOICE_STATUSES, JOB_STATUSES, QUOTE_STATUSES,
                      Config, Palette)
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Lists"
LAST_COL = 9

#: name -> (header, [values]) - written one list per column block
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("YesNoNAList", ["Yes", "No", "N/A"]),
    ("ClientTypeList", ["Residential", "Commercial office", "Retail",
                        "Restaurant & cafe", "Holiday let / Airbnb",
                        "Landlord", "Letting agent", "Property manager",
                        "Builder / contractor", "Medical & dental",
                        "Gym & studio", "School & nursery", "Church & hall",
                        "Warehouse", "Other"]),
    ("ClientStatusList", ["Prospect", "New", "Active", "Regular",
                          "Occasional", "Paused", "Inactive", "Lost",
                          "Do not serve"]),
    ("PropertyTypeList", ["House", "Apartment", "Townhouse", "Bungalow",
                          "Studio", "Office suite", "Shop", "Restaurant",
                          "Clinic", "Salon", "Gym", "School", "Warehouse",
                          "Communal area", "Holiday let", "Other"]),
    ("AccessList", ["Client at home", "We hold a key", "Lockbox",
                    "Keypad code", "Smart lock", "Concierge / front desk",
                    "Reception sign-in", "Key collected each visit",
                    "Neighbour holds key", "Agent meets us", "Open access"]),
    ("KeyStatusList", ["Held by us", "With the cleaner", "In the key safe",
                       "Returned to client", "Not applicable", "Lost"]),
    ("ParkingList", ["Driveway", "Free street parking", "Metered street",
                     "Permit needed", "Visitor bay", "Car park",
                     "Loading bay", "Underground garage", "Difficult",
                     "Pay and display"]),
    ("ConditionList", list(CONDITION_FACTORS)),
    ("FrequencyList", list(FREQUENCIES)),
    ("ServiceCategoryList", ["Regular domestic clean", "Deep clean",
                             "End of tenancy", "Move in / move out",
                             "Commercial & office", "Spring clean",
                             "Post-construction", "Airbnb changeover",
                             "Carpet & upholstery", "Window cleaning",
                             "Oven cleaning", "Pressure washing",
                             "Add-on / extra", "One-off special"]),
    ("PricingModelList", ["Per hour", "Flat price per visit",
                          "Per area", "Per room", "Per bedroom & bathroom",
                          "Per item", "Monthly contract", "Quote only"]),
    ("JobStatusList", list(JOB_STATUSES)),
    ("JobTypeList", ["Recurring visit", "One-off clean", "First deep clean",
                     "Quote visit", "Re-clean (free)", "Emergency call out",
                     "Contract work"]),
    ("PriorityList", ["Normal", "High", "Urgent", "VIP client", "Flexible"]),
    ("TimeSlotList", ["Early (7-9)", "Morning (9-12)", "Midday (12-2)",
                      "Afternoon (2-5)", "Late (5-8)", "Evening",
                      "Flexible", "Client to confirm"]),
    ("QuoteStatusList", list(QUOTE_STATUSES)),
    ("ContractStatusList", ["No contract", "Draft", "Sent", "Signed",
                            "Active", "Due for renewal", "Renewed",
                            "Expired", "Terminated", "Declined"]),
    ("ContractTermList", ["Rolling monthly", "Trial period", "3 months",
                          "6 months", "12 months", "24 months",
                          "No fixed term"]),
    ("BillingCycleList", ["Pay per visit", "Weekly", "Fortnightly",
                          "Monthly in arrears", "Monthly in advance",
                          "Quarterly", "On completion",
                          "Deposit then balance"]),
    ("InvoiceStatusList", list(INVOICE_STATUSES)),
    ("PaymentMethodList", ["Bank transfer", "Card", "Cash", "Cheque",
                           "Direct debit", "Standing order", "PayPal",
                           "Online payment link", "Mobile wallet",
                           "Account credit", "Other"]),
    ("LeadSourceList", ["Google search", "Google Business Profile",
                        "Facebook", "Instagram", "Nextdoor", "Yelp",
                        "Directory listing", "Referral - client",
                        "Referral - staff", "Flyer / leaflet",
                        "Van signage", "Word of mouth", "Repeat client",
                        "Letting agent", "Walk-up", "Other"]),
    ("LeadStageList", ["New enquiry", "Contacted", "Visit booked",
                       "Quote sent", "Following up", "Negotiating", "Won",
                       "Lost", "On hold", "Not a fit"]),
    ("LostReasonList", ["Price too high", "Chose a competitor",
                        "No response", "Out of our area", "Wrong timing",
                        "Doing it themselves", "Needed a specialist",
                        "Not a fit", "Duplicate enquiry", "Other"]),
    ("TaskAreaList", ["Whole property", "Kitchen", "Bathroom",
                      "Second bathroom", "Bedroom", "Living room",
                      "Dining room", "Hallway & stairs", "Home office",
                      "Utility room", "Conservatory", "Reception",
                      "Open plan office", "Meeting room", "Toilets",
                      "Kitchenette", "Communal area", "Windows",
                      "Exterior", "Garage"]),
    ("TaskStatusList", ["To do", "In progress", "Done", "Skipped",
                        "Not needed", "Issue found", "Needs rework"]),
    ("QualityList", ["Excellent", "Good", "Acceptable", "Needs rework",
                     "Failed"]),
    ("RatingList", ["5 - Delighted", "4 - Happy", "3 - Fine",
                    "2 - Unhappy", "1 - Complaint"]),
    ("EmployeeRoleList", ["Owner", "Manager", "Team leader", "Supervisor",
                          "Cleaner", "Deep clean specialist",
                          "Window cleaner", "Carpet technician", "Driver",
                          "Trainee", "Subcontractor", "Office admin"]),
    ("EmploymentTypeList", ["Full time", "Part time", "Casual",
                            "Zero hours", "Subcontractor", "Agency",
                            "Apprentice", "Owner"]),
    ("EmployeeStatusList", ["Active", "Probation", "Training", "On leave",
                            "Sick", "Working notice", "Left", "Standby"]),
    ("ShiftList", ["Morning", "Afternoon", "Evening", "Night", "All day",
                   "Split shift", "On call", "Day off"]),
    ("AttendanceList", ["On time", "Late", "Left early", "Absent - told us",
                        "Absent - no show", "Holiday", "Sick", "Training",
                        "Day off", "Public holiday"]),
    ("LeaveTypeList", ["Holiday", "Sick", "Unpaid", "Parental",
                       "Bereavement", "Training", "Public holiday",
                       "Jury service"]),
    ("TrainingList", ["Induction", "COSHH & chemicals", "Manual handling",
                      "Working at height", "Infection control",
                      "Customer service", "Deep clean technique",
                      "Carpet machine", "Window technique", "First aid",
                      "Safeguarding", "Driving & vehicle", "Data privacy"]),
    ("SupplyCategoryList", ["Chemicals & detergents", "Cloths & sponges",
                            "Mops & buckets", "Paper & consumables",
                            "Bin bags & waste", "Gloves & PPE",
                            "Vacuum bags & filters", "Laundry",
                            "Specialist products", "Hardware & spares",
                            "Uniform", "Office"]),
    ("SupplyUnitList", ["Each", "Bottle", "Litre", "5L drum", "Gallon",
                        "Box", "Case", "Pack", "Roll", "Pair", "Kg",
                        "Refill", "Sachet"]),
    ("EquipmentTypeList", ["Upright vacuum", "Backpack vacuum",
                           "Wet & dry vacuum", "Carpet machine",
                           "Steam cleaner", "Floor polisher",
                           "Pressure washer", "Window pole system",
                           "Ladder & steps", "Trolley & caddy",
                           "Vehicle", "Power tool", "Other"]),
    ("EquipmentStatusList", ["In service", "In use", "Spare", "In repair",
                             "Awaiting parts", "Out of service", "Retired",
                             "Lost or stolen"]),
    ("OwnershipList", ["Owned outright", "Financed", "Leased", "Rented",
                       "Staff owned", "Client supplied"]),
    ("MaintenanceList", ["Routine service", "Filter change", "Belt change",
                         "Repair", "Deep clean", "Safety test",
                         "Calibration", "Inspection", "Replacement due"]),
    ("ExpenseCategoryList", ["Cleaning supplies", "Equipment purchase",
                             "Equipment repair", "Vehicle fuel",
                             "Vehicle maintenance", "Vehicle insurance",
                             "Business insurance", "Wages & payroll",
                             "Subcontractors", "Marketing & advertising",
                             "Website & hosting", "Software & apps",
                             "Phone & internet", "Office & rent",
                             "Uniforms", "Training & certification",
                             "Licences & memberships",
                             "Accounting & legal", "Bank & card fees",
                             "Laundry", "Waste disposal",
                             "Travel & parking", "Bad debt", "Other"]),
    ("ExpenseTypeList", ["Fixed monthly", "Variable", "One-off",
                         "Job related", "Annual"]),
    ("MileagePurposeList", ["Travel to job", "Between jobs", "Quote visit",
                            "Supply run", "Equipment repair",
                            "Staff pick-up", "Bank & post", "Training",
                            "Admin errand", "Marketing drop", "Other"]),
    ("VehicleList", ["Van 1", "Van 2", "Van 3", "Car", "Staff vehicle",
                     "Hire van", "Bike"]),
    ("ComplaintTypeList", ["Missed area", "Quality below standard",
                           "Arrived late", "Did not turn up",
                           "Damage to property", "Breakage",
                           "Item gone missing", "Staff conduct",
                           "Billing error", "Access problem", "Pet issue",
                           "Communication", "Wrong day", "Other"]),
    ("ComplaintStatusList", ["Logged", "Investigating", "Awaiting client",
                             "Re-clean booked", "Resolved", "Escalated",
                             "Refunded", "Closed - no action"]),
    ("SeverityList", ["Low", "Medium", "High", "Critical"]),
    ("ResolutionList", ["Free re-clean", "Partial refund", "Full refund",
                        "Discount next visit", "Apology", "Staff coaching",
                        "Item replaced", "Insurance claim",
                        "Process changed", "No action needed"]),
    ("IncidentTypeList", ["Slip, trip or fall", "Cut or laceration",
                          "Chemical splash", "Burn or scald",
                          "Manual handling strain", "Property damage",
                          "Equipment failure", "Near miss",
                          "Vehicle incident", "Aggression or abuse",
                          "Lone worker concern", "Illness at work",
                          "Other"]),
    ("SafetyDocList", ["COSHH data sheet", "Risk assessment",
                       "Method statement", "Public liability insurance",
                       "Employer liability insurance",
                       "Background check", "Training certificate",
                       "First aid certificate", "Electrical safety test",
                       "Vehicle insurance", "Driving licence",
                       "Right to work", "Client contract",
                       "Privacy notice", "Business licence"]),
    ("DocStatusList", ["Valid", "Due soon", "Expired", "Missing",
                       "Pending renewal", "Not required"]),
    ("RetentionList", ["Loyal", "Steady", "Slipping", "At risk", "Dormant",
                       "Won back", "Lost"]),
    ("MonthList", ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November",
                   "December"]),
    ("WeekdayList", ["Monday", "Tuesday", "Wednesday", "Thursday",
                     "Friday", "Saturday", "Sunday"]),
    ("WeekStartList", ["Monday", "Sunday"]),
    ("AreaUnitList", ["sq ft", "sq m"]),
    ("DistanceUnitList", ["miles", "km"]),
    ("RoundingList", ["Nearest 1", "Nearest 5", "Nearest 10", "Exact"]),
    ("GoalStatusList", ["Not started", "On track", "Behind", "At risk",
                        "Achieved", "Paused"]),
]

#: frequency -> (days between visits, visits a year, price factor)
#: The factor is why a weekly client pays less per visit than a one-off:
#: a place that is cleaned every week never gets badly dirty.
FREQUENCY_TABLE: List[Tuple[str, int, float, float]] = [
    ("One-off", 0, 1, FREQUENCY_FACTORS["One-off"]),
    ("Weekly", 7, 52, FREQUENCY_FACTORS["Weekly"]),
    ("Fortnightly", 14, 26, FREQUENCY_FACTORS["Fortnightly"]),
    ("Every 3 weeks", 21, 17, FREQUENCY_FACTORS["Every 3 weeks"]),
    ("Monthly", 28, 12, FREQUENCY_FACTORS["Monthly"]),
    ("Quarterly", 91, 4, FREQUENCY_FACTORS["Quarterly"]),
    ("Twice a week", 3, 104, FREQUENCY_FACTORS["Twice a week"]),
]

#: how dirty the place is -> how much longer it takes
CONDITION_TABLE: List[Tuple[str, float]] = list(CONDITION_FACTORS.items())


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.grey_text)

    s.banner("Drop Down Lists & Pricing Factors",
             "Every list the workbook offers you, and the two tables the "
             "quote calculator multiplies by. Edit an entry and the drop "
             "down that uses it changes everywhere.",
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

    # ------------------------------------------------------- frequency table
    row = s.section(row, "How often we clean, and what that does to the "
                         "price", LAST_COL)
    row = s.note(
        row,
        "Regular work is worth a discount - a weekly house never gets as "
        "dirty as a one-off, so it takes less time. The factor below is "
        "multiplied by the standard price on every quote and every "
        "recurring booking. 'Days between visits' is what the scheduler "
        "adds to today's date to work out the next clean.",
        LAST_COL, height=44)
    row += 1

    s.table_head(row, 0, ["Frequency", "Days between visits",
                          "Visits a year", "Price factor"],
                 [20, 20, 14, 14])
    first_freq = row + 1
    for index, (freq, days, per_year, factor) in enumerate(FREQUENCY_TABLE):
        r = first_freq + index
        s.write(r, 0, freq, fmts.input("text"))
        s.write(r, 1, days, fmts.input("int"))
        s.write(r, 2, per_year, fmts.input("int"))
        s.write(r, 3, factor, fmts.input("number"))
    spare_freq = 5
    for index in range(len(FREQUENCY_TABLE),
                       len(FREQUENCY_TABLE) + spare_freq):
        r = first_freq + index
        for c, style in enumerate(["text", "int", "int", "number"]):
            s.write(r, c, None, fmts.input(style))
    last_freq = first_freq + len(FREQUENCY_TABLE) + spare_freq

    sheet_ref = quote_sheet(NAME)
    book.define_name("FreqNames",
                     f"={sheet_ref}!$A${first_freq + 1}:$A${last_freq}")
    book.define_name("FreqDays",
                     f"={sheet_ref}!$B${first_freq + 1}:$B${last_freq}")
    book.define_name("FreqPerYear",
                     f"={sheet_ref}!$C${first_freq + 1}:$C${last_freq}")
    book.define_name("FreqFactors",
                     f"={sheet_ref}!$D${first_freq + 1}:$D${last_freq}")
    row = last_freq + 2

    # ------------------------------------------------------- condition table
    row = s.section(row, "How dirty is it, and what that does to the time",
                    LAST_COL)
    row = s.note(
        row,
        "The honest part of pricing a clean. The same three bedroom house "
        "can take three hours or six depending on what state it is in, so "
        "the calculator multiplies the estimated hours by the factor "
        "below. Walk the property, pick a line, and the quote follows.",
        LAST_COL, height=38)
    row += 1

    s.table_head(row, 0, ["Condition", "Time factor", "What it means"],
                 [26, 14, 52])
    meanings = [
        "Kept on top of between visits - a quick, tidy run through.",
        "A normal lived-in home or a tidy office.",
        "Behind on the housework - expect extra scrubbing.",
        "Neglected for months: grease, limescale, heavy build-up.",
        "Specialist job. Price it after a site visit, never over "
        "the phone.",
    ]
    first_cond = row + 1
    for index, (condition, factor) in enumerate(CONDITION_TABLE):
        r = first_cond + index
        s.write(r, 0, condition, fmts.input("text"))
        s.write(r, 1, factor, fmts.input("number"))
        s.write(r, 2,
                meanings[index] if index < len(meanings) else None,
                fmts.input("text"))
    spare_cond = 4
    for index in range(len(CONDITION_TABLE),
                       len(CONDITION_TABLE) + spare_cond):
        r = first_cond + index
        for c, style in enumerate(["text", "number", "text"]):
            s.write(r, c, None, fmts.input(style))
    last_cond = first_cond + len(CONDITION_TABLE) + spare_cond

    book.define_name("ConditionNames",
                     f"={sheet_ref}!$A${first_cond + 1}:$A${last_cond}")
    book.define_name("ConditionFactors",
                     f"={sheet_ref}!$B${first_cond + 1}:$B${last_cond}")
    row = last_cond + 2

    row = s.note(
        row,
        "A word on the factors: they are a starting point from real "
        "cleaning businesses, not a law. Track your actual hours on the "
        "Job Costing tab for a month, compare them with what you quoted, "
        "then nudge these numbers until the quotes match reality.",
        LAST_COL, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/settings.py`

*342 lines · 1 function*

> The `Settings` tab - business profile, pricing, pay rates, targets, alerts.
>
> Everything the rest of the workbook needs to know about your cleaning
> business is entered once here and published as a defined name (BizName,
> HourlyRate, MinimumCharge, PayRate, TaxRate, PaymentTerms ...) so the
> formulas stay readable and there is exactly one place to change a price.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `SHEET` ('Settings'), `LAST_COL` (8).

````python
"""
The `Settings` tab - business profile, pricing, pay rates, targets, alerts.

Everything the rest of the workbook needs to know about your cleaning
business is entered once here and published as a defined name (BizName,
HourlyRate, MinimumCharge, PayRate, TaxRate, PaymentTerms ...) so the
formulas stay readable and there is exactly one place to change a price.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

SHEET = "Settings"
LAST_COL = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, SHEET, tab_color=Palette.accent)
    b = cfg.business

    s.banner("Business Settings",
             "Fill in the white cells once. Shaded cells are locked "
             "formulas. Every price, every invoice and every alert in "
             "the workbook reads these values.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 34, fmts.chrome("label"))
    s.set_col(2, 2, 14, fmts.chrome("label"))
    s.set_col(3, 3, 24)
    s.set_col(4, 4, 18)
    s.set_col(5, 8, 17, fmts.chrome("note"))

    r = 4  # 0 based row index -> Excel row 5

    # -------------------------------------------------------- 1. the business
    r = s.section(r, "1. Your business", LAST_COL)
    r = s.field(r, "Business name", b.name, name="BizName", value_end=4,
                help_text="Shown on every quote, invoice and printable "
                          "page.")
    r = s.field(r, "Owner / manager", b.owner, name="BizOwner", value_end=4)
    r = s.field(r, "Email", b.email, name="BizEmail", value_end=4)
    r = s.field(r, "Phone", b.phone, name="BizPhone", value_end=4)
    r = s.field(r, "Address", b.address, name="BizAddress", value_end=4)
    r = s.field(r, "Website", b.website, name="BizWebsite", value_end=4)
    r = s.field(r, "Area you cover", b.area, name="BizArea", value_end=4,
                help_text="Used on quotes and to sanity check how far a "
                          "job is.")
    r = s.field(r, "Tagline", b.tagline, name="BizTagline", value_end=4)
    r = s.field(r, "Business / tax number", b.reg_number, name="BizRegNo",
                value_end=4)
    r = s.field(r, "Insurance policy", b.insurance, name="BizInsurance",
                value_end=4,
                help_text="Clients ask for this more often than you think "
                          "- keep it here so it is easy to quote.")

    # ---------------------------------------------------------- 2. the diary
    r += 1
    r = s.section(r, "2. Your working week", LAST_COL)
    r = s.field(r, "Year", cfg.year, style="int", name="PlanYear",
                validate=number_check(2000, 2100, decimal=False),
                help_text="The year the calendar, the planner and the "
                          "reports open on.")
    r = s.field(r, "Week starts on", cfg.week_start, name="WeekStart",
                validate=list_from("WeekStartList"),
                help_text="Used by the calendar grid, the staff rota and "
                          "the week numbers.")
    r = s.field(r, "Billable hours in a working day", cfg.productive_hours,
                style="number1", name="ProductiveHours",
                validate=number_check(1, 16),
                help_text="Hours one cleaner can actually charge for "
                          "after travel, setup and breaks. The capacity "
                          "and utilisation figures use this.")
    r = s.field(r, "Cleaners on a standard job", cfg.cleaners_per_job,
                style="int", name="CleanersPerJob",
                validate=number_check(1, 20, decimal=False),
                help_text="Two cleaners for three hours is six paid "
                          "hours - the costing tab works that out for "
                          "you.")
    r = s.field(r, "Travel allowance between jobs (minutes)",
                cfg.travel_minutes, style="int", name="TravelMinutes",
                validate=number_check(0, 240, decimal=False),
                help_text="Added to each job when the scheduler checks "
                          "whether two bookings clash.")
    r = s.field(r, "Area measured in", cfg.area_unit, name="AreaUnit",
                validate=list_from("AreaUnitList"),
                help_text="Square feet or square metres - the price per "
                          "unit below should match.")
    r = s.field(r, "Distance measured in", cfg.distance_unit,
                name="DistanceUnit", validate=list_from("DistanceUnitList"))
    r = s.field(r, "Today", None, style="date", mode="calc",
                name="TodayDate", formula="=TODAY()",
                help_text="Every 'due', 'overdue' and 'coming up' flag in "
                          "the workbook counts from this cell, so they all "
                          "move on by themselves overnight.")
    r = s.field(r, "This month starts", None, style="date", mode="calc",
                name="MonthStart",
                formula="=DATE(YEAR(TodayDate),MONTH(TodayDate),1)")
    r = s.field(r, "This month ends", None, style="date", mode="calc",
                name="MonthEnd",
                formula="=DATE(YEAR(TodayDate),MONTH(TodayDate)+1,1)-1")
    r = s.field(r, "This week starts", None, style="date", mode="calc",
                name="WeekStartDate",
                formula="=TodayDate-WEEKDAY(TodayDate,"
                        "IF(WeekStart=\"Sunday\",1,2))+1")

    # --------------------------------------------------------- 3. your prices
    r += 1
    r = s.section(r, "3. What you charge", LAST_COL)
    r = s.note(r, "These are the numbers the quote calculator multiplies. "
                  "Set them once from your own costs and local market - "
                  "the figures shipped with the workbook are a realistic "
                  "starting point, not advice on what to charge.",
               LAST_COL, first_col=1, height=30)
    r = s.field(r, "Standard hourly rate", cfg.hourly_rate, style="money",
                name="HourlyRate", validate=number_check(0, 10000),
                help_text="What one cleaner's hour costs the client.")
    r = s.field(r, "Minimum charge for any job", cfg.minimum_charge,
                style="money", name="MinimumCharge",
                validate=number_check(0, 10000),
                help_text="No quote leaves below this, however small the "
                          "job. It covers getting there and back.")
    r = s.field(r, f"Price per {cfg.area_unit}", cfg.rate_per_area,
                style="money", name="RatePerArea",
                validate=number_check(0, 100),
                help_text="Used by the area based pricing method - "
                          "common for offices and end of tenancy work.")
    r = s.field(r, "Price per bedroom", cfg.rate_per_bedroom, style="money",
                name="RatePerBedroom", validate=number_check(0, 10000))
    r = s.field(r, "Price per bathroom", cfg.rate_per_bathroom,
                style="money", name="RatePerBathroom",
                validate=number_check(0, 10000),
                help_text="Bathrooms and kitchens take the longest - "
                          "price them higher than a bedroom.")
    r = s.field(r, "Price per other room", cfg.rate_per_room, style="money",
                name="RatePerRoom", validate=number_check(0, 10000))
    r = s.field(r, "Call out / travel fee", cfg.callout_fee, style="money",
                name="CalloutFee", validate=number_check(0, 10000))
    r = s.field(r, "Deep clean uplift", cfg.deep_clean_uplift,
                style="percent0", name="DeepCleanUplift",
                validate=number_check(0, 5),
                help_text="A deep clean takes far longer than a weekly "
                          "visit. 45% means a $200 clean quotes at $290.")
    r = s.field(r, "Move in / move out uplift", cfg.move_clean_uplift,
                style="percent0", name="MoveCleanUplift",
                validate=number_check(0, 5))
    r = s.field(r, "Pet surcharge", cfg.pet_surcharge, style="money",
                name="PetSurcharge", validate=number_check(0, 10000),
                help_text="Hair and paw prints are real extra minutes. "
                          "Set it to 0 if you would rather absorb it.")
    r = s.field(r, "Deposit to take up front", cfg.deposit_percent,
                style="percent0", name="DepositPercent",
                validate=number_check(0, 1),
                help_text="Shown on quotes and contracts for larger jobs.")
    r = s.field(r, "Price rounding", "Nearest 5", name="PriceRounding",
                validate=list_from("RoundingList"),
                help_text="Quotes look more professional rounded. "
                          "'Nearest 5' turns 187.40 into 185.")

    # ----------------------------------------------------- 4. what it costs
    r += 1
    r = s.section(r, "4. What the work costs you", LAST_COL)
    r = s.field(r, "Cleaner pay rate an hour", cfg.pay_rate, style="money",
                name="PayRate", validate=number_check(0, 10000),
                help_text="Gross pay before your on-costs.")
    r = s.field(r, "Overtime multiplier", cfg.overtime_multiplier,
                style="number1", name="OvertimeRate",
                validate=number_check(1, 5),
                help_text="1.5 = time and a half on overtime hours.")
    r = s.field(r, "Employer on-costs", cfg.employer_oncost,
                style="percent0", name="EmployerOncost",
                validate=number_check(0, 2),
                help_text="Payroll tax, insurance, holiday pay, sick pay "
                          "- the part of an employee most owners forget "
                          "to price in.")
    r = s.field(r, "True cost of one cleaner hour", None, style="money",
                mode="calc", name="CostPerHour",
                formula="=PayRate*(1+EmployerOncost)",
                help_text="Pay plus on-costs. Every job costing and "
                          "profit figure in the workbook uses this.")
    r = s.field(r, "Supplies used per hour", cfg.supplies_per_hour,
                style="money", name="SuppliesPerHour",
                validate=number_check(0, 1000),
                help_text="Chemicals, cloths, bags - the consumables a "
                          "cleaning hour burns through.")
    r = s.field(r, f"Mileage rate per {cfg.distance_unit[:-1]}",
                cfg.travel_rate, style="money", name="MileageRate",
                validate=number_check(0, 100),
                help_text="What you reimburse (or claim) for travel.")
    r = s.field(r, "Target gross margin", cfg.target_margin,
                style="percent0", name="TargetMargin",
                validate=number_check(0, 1),
                help_text="What a healthy job should make after wages "
                          "and supplies. The quote calculator shows you "
                          "the price that hits it.")
    r = s.field(r, "Flag a job below this margin", cfg.margin_floor,
                style="percent0", name="MarginFloor",
                validate=number_check(0, 1),
                help_text="Anything thinner than this turns red on the "
                          "costing tab and the dashboard.")

    # --------------------------------------------------------- 5. invoicing
    r += 1
    r = s.section(r, "5. Invoicing and payment", LAST_COL)
    r = s.field(r, "Sales tax / VAT rate", cfg.tax_rate, style="percent0",
                name="TaxRate", validate=number_check(0, 1),
                help_text="Leave at 0 if you are not registered. Every "
                          "invoice total adds this to the net amount.")
    r = s.field(r, "What you call the tax", cfg.tax_label, name="TaxLabel",
                help_text="VAT, GST, Sales tax - whatever the invoice "
                          "should say.")
    r = s.field(r, "Payment terms (days)", cfg.payment_terms_days,
                style="int", name="PaymentTerms",
                validate=number_check(0, 180, decimal=False),
                help_text="Due date = invoice date plus this. The overdue "
                          "flags count from the due date.")
    r = s.field(r, "Late fee", cfg.late_fee_percent, style="percent0",
                name="LateFeePercent", validate=number_check(0, 1),
                help_text="Only charge it if your contract says you can.")
    r = s.field(r, "Grace days before 'overdue'", cfg.overdue_grace_days,
                style="int", name="OverdueGrace",
                validate=number_check(0, 60, decimal=False))
    r = s.field(r, "A quote stays valid for (days)", cfg.quote_valid_days,
                style="int", name="QuoteValidDays",
                validate=number_check(1, 365, decimal=False))
    r = s.field(r, "Bank / payment details", "Bright & Tidy Ltd - "
                                             "00-00-00 12345678",
                name="BizPayDetails", value_end=4,
                help_text="Printed at the bottom of every invoice.")

    # ------------------------------------------------------- 6. the targets
    r += 1
    r = s.section(r, "6. Targets the dashboard measures you against",
                  LAST_COL)
    r = s.field(r, "Revenue target a month", cfg.revenue_target_month,
                style="money0", name="MonthlyRevenueTarget",
                validate=number_check(0, 100000000))
    r = s.field(r, "Revenue target a year", cfg.revenue_target_year,
                style="money0", name="AnnualRevenueTarget",
                validate=number_check(0, 1000000000))
    r = s.field(r, "Expense budget a month", cfg.expense_budget_month,
                style="money0", name="MonthlyExpenseBudget",
                validate=number_check(0, 100000000))
    r = s.field(r, "Utilisation target", cfg.utilisation_target,
                style="percent0", name="UtilisationTarget",
                validate=number_check(0, 1),
                help_text="Share of the hours you pay for that you "
                          "actually bill. 75% is good for a small team.")
    r = s.field(r, "Client retention target", cfg.retention_target,
                style="percent0", name="RetentionTarget",
                validate=number_check(0, 1))
    r = s.field(r, "Quote win rate target", cfg.quote_win_target,
                style="percent0", name="QuoteWinTarget",
                validate=number_check(0, 1),
                help_text="Winning every quote usually means you are too "
                          "cheap. Half is a healthy place to be.")
    r = s.field(r, "Satisfaction target (out of 5)",
                cfg.satisfaction_target, style="number1",
                name="SatisfactionTarget", validate=number_check(0, 5))

    # ------------------------------------------------------- 7. the alerts
    r += 1
    r = s.section(r, "7. When the workbook should nudge you", LAST_COL)
    r = s.field(r, "Remind me about a job this many days early",
                cfg.reminder_days, style="int", name="ReminderDays",
                validate=number_check(0, 60, decimal=False),
                help_text="Drives the 'coming up' list on the dashboard "
                          "and the confirmation chase list.")
    r = s.field(r, "Chase an unanswered quote after (days)",
                cfg.quote_followup_days, style="int",
                name="QuoteFollowupDays",
                validate=number_check(0, 120, decimal=False),
                help_text="The single highest value number on this tab. "
                          "Most lost quotes were never followed up.")
    r = s.field(r, "A client is 'going quiet' after (days)",
                cfg.inactive_client_days, style="int",
                name="InactiveClientDays",
                validate=number_check(0, 999, decimal=False),
                help_text="Used by the retention tab to spot clients "
                          "drifting away before they are gone.")
    r = s.field(r, "Days of stock that counts as low", cfg.low_stock_days,
                style="int", name="LowStockDays",
                validate=number_check(0, 365, decimal=False))
    r = s.field(r, "Service equipment every (days)", cfg.maintenance_days,
                style="int", name="MaintenanceDays",
                validate=number_check(0, 3650, decimal=False))
    r = s.field(r, "Warn before a certificate expires (days)",
                cfg.certificate_warning_days, style="int",
                name="CertWarningDays",
                validate=number_check(0, 365, decimal=False),
                help_text="Insurance, background checks, training - the "
                          "Safety tab turns amber this far ahead.")
    r = s.field(r, "Warn before a contract renews (days)",
                cfg.renewal_warning_days, style="int",
                name="RenewalWarningDays",
                validate=number_check(0, 365, decimal=False))

    # ------------------------------------------------------- 8. ID prefixes
    r += 1
    r = s.section(r, "8. ID prefixes", LAST_COL)
    r = s.note(r, "The letters in front of the automatic numbers. Change "
                  "them before you start entering data, not after.",
               LAST_COL, first_col=1, height=24)
    for label, value, name in [
            ("Client IDs", "CL-", "ClientPrefix"),
            ("Property IDs", "PR-", "PropertyPrefix"),
            ("Lead IDs", "LD-", "LeadPrefix"),
            ("Service IDs", "SV-", "ServicePrefix"),
            ("Quote IDs", "Q-", "QuotePrefix"),
            ("Job IDs", "JOB-", "JobPrefix"),
            ("Invoice IDs", "INV-", "InvoicePrefix"),
            ("Payment IDs", "PAY-", "PaymentPrefix"),
            ("Employee IDs", "EMP-", "EmployeePrefix"),
            ("Supply IDs", "SUP-", "SupplyPrefix"),
            ("Equipment IDs", "EQ-", "EquipmentPrefix"),
            ("Expense IDs", "EXP-", "ExpensePrefix"),
            ("Complaint IDs", "CMP-", "ComplaintPrefix")]:
        r = s.field(r, label, value, name=name)

    # ---------------------------------------------------------- 9. about it
    r += 1
    r = s.section(r, "9. About this workbook", LAST_COL)
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
           "Changed a price here? Everything updates straight away - the "
           "quote calculator, every open quote, the job costing and the "
           "dashboard all read these cells rather than keeping their own "
           "copy. Invoices you have already sent keep the amount you "
           "typed on them.", LAST_COL, first_col=1, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/clients.py`

*255 lines · 1 function*

> `Clients` - the customer database every other tab points at.
>
> One row per client. You type who they are and how they like to be
> cleaned for; the workbook adds up what they have spent, when you last
> saw them, what their work actually earns you and whether they are
> quietly drifting away.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Clients').

````python
"""`Clients` - the customer database every other tab points at.

One row per client. You type who they are and how they like to be
cleaned for; the workbook adds up what they have spent, when you last
saw them, what their work actually earns you and whether they are
quietly drifting away.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if, min_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, rank_desc,
                     status_cf)

NAME = "Clients"

_JOBS = "[[Jobs.client]]"
_INV = "[[Invoices.client]]"

#: the next visit in the diary for this client
_NEXT = min_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
               "[[Jobs.date]]>=TodayDate",
               '[[Jobs.status]]<>"Cancelled"')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ClientPrefix", digits=4, width=10,
               note="The ID the properties, jobs, quotes and invoices all "
                    "refer to. Generated for you - never type over it."),
        Col("name", "Client", 26, "text",
            note="A household name or a company name. One row per "
                 "paying customer; their addresses go on the Properties "
                 "tab."),
        Col("type", "Type", 16, "text", validate=list_from("ClientTypeList"),
            note="Residential and commercial work behave very "
                 "differently - the reports split on this."),
        Col("contact", "Main contact", 18, "text"),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 26, "text"),
        Col("address", "Billing address", 34, "wrap",
            note="Where the invoice goes. The place you clean goes on "
                 "the Properties tab - often not the same address."),
        Col("status", "Status", 13, "status",
            validate=list_from("ClientStatusList"), cf=status_cf(),
            note="Prospect until they book. Mark them Paused rather "
                 "than deleting them - the history is worth keeping."),
        Col("source", "Came from", 16, "text",
            validate=list_from("LeadSourceList"),
            note="Where they heard about you. Fill this in every time "
                 "and the Reports tab tells you where to spend your "
                 "marketing money."),
        Col("since", "Client since", 12, "date", validate=date_check()),
        Col("frequency", "Usual frequency", 14, "text",
            validate=list_from("FrequencyList"),
            note="Drives the recurring booking dates and the value of "
                 "the client over a year."),
        Col("pref_day", "Preferred day", 12, "text",
            validate=list_from("WeekdayList")),
        Col("pref_time", "Preferred time", 14, "text",
            validate=list_from("TimeSlotList")),
        Col("pref_cleaner", "Preferred cleaner", 16, "text",
            validate=list_from("EmployeeNames"),
            note="The scheduler flags a job given to anybody else, so "
                 "you can at least tell the client."),
        Col("rate", "Agreed hourly rate", 13, "money",
            validate=number_check(0, 10000),
            note="Leave blank to use the standard rate from Settings. "
                 "Fill it in for a client on a special price."),
        Col("terms", "Payment terms (days)", 12, "int",
            validate=number_check(0, 180, decimal=False),
            note="Blank uses the default terms from the Settings tab."),
        Col("payment_method", "Pays by", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("contract", "Contract", 14, "status",
            validate=list_from("ContractTermList"), cf=status_cf()),
        Col("contract_end", "Contract ends", 12, "date",
            validate=date_check(),
            note="The renewal alert counts down to this date."),
        Col("vip", "VIP?", 7, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("marketing", "Happy to be contacted?", 11, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Keep a record of consent before you add anyone to a "
                 "mailing list."),
        Col("notes", "Notes", 40, "wrap",
            note="Allergies, pets, gate codes, the mug they do not want "
                 "moved. The things that keep a client."),
        # ------------------------------------------------- what they do --
        Col("properties", "Properties", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Properties.client]],'
                    '{id}{r}))',
            note="Counted from the Properties tab."),
        Col("jobs_booked", "Jobs booked", 11, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _JOBS + ',{id}{r}))'),
        Col("jobs_done", "Jobs done", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _JOBS + ',{id}{r},'
                    '[[Jobs.status]],"Completed")+COUNTIFS(' + _JOBS +
                    ',{id}{r},[[Jobs.status]],"Invoiced"))'),
        Col("cancelled", "Cancelled", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _JOBS + ',{id}{r},'
                    '[[Jobs.status]],"Cancelled")+COUNTIFS(' + _JOBS +
                    ',{id}{r},[[Jobs.status]],"No access"))',
            cf=[{"type": "cell", "criteria": ">=", "value": 3,
                 "fmt": "warn"}]),
        Col("last_job", "Last clean", 12, "date",
            formula='=IF({id}{r}="","",IF(' +
                    max_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
                           '([[Jobs.status]]="Completed")+'
                           '([[Jobs.status]]="Invoiced")') +
                    '=0,"",' +
                    max_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
                           '([[Jobs.status]]="Completed")+'
                           '([[Jobs.status]]="Invoiced")') + '))',
            note="The last visit that actually happened."),
        Col("days_since", "Days since", 10, "int",
            formula='=IF({last_job}{r}="","",TodayDate-{last_job}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": "=InactiveClientDays",
                 "fmt": "bad"}]),
        Col("next_job", "Next booking", 12, "date",
            formula='=IF({id}{r}="","",IF(' + _NEXT + '=2958465,"",' +
                    _NEXT + '))',
            note="The next visit in the diary. Empty means nothing is "
                 "booked - that is the list to work from."),
        # ------------------------------------------------- what they pay --
        Col("invoiced", "Invoiced to date", 14, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Invoices.total]],' + _INV +
                    ',{id}{r}))',
            note="Everything you have billed them, lifetime."),
        Col("received", "Paid to date", 13, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Invoices.paid]],' + _INV +
                    ',{id}{r}))'),
        Col("outstanding", "Owing now", 12, "money",
            formula='=IF({id}{r}="","",{invoiced}{r}-{received}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("avg_job", "Average job", 12, "money",
            formula='=IF(OR({id}{r}="",{jobs_done}{r}=0),"",'
                    '{invoiced}{r}/{jobs_done}{r})'),
        Col("profit", "Profit to date", 13, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Job Costing.profit]],'
                    '[[Job Costing.client]],{id}{r}))',
            note="Revenue less wages, supplies and travel, from the Job "
                 "Costing tab. The number that tells you whether a "
                 "client is worth keeping.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({id}{r}="",{invoiced}{r}=0),"",'
                    '{profit}{r}/{invoiced}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}]),
        Col("year_value", "Value a year", 13, "money",
            formula='=IF(OR({id}{r}="",{frequency}{r}="",{avg_job}{r}=""),'
                    '"",{avg_job}{r}*IFERROR(INDEX(FreqPerYear,'
                    'MATCH({frequency}{r},FreqNames,0)),1))',
            note="What this client is worth over twelve months at their "
                 "usual frequency. Losing one of these is worth a phone "
                 "call to keep."),
        Col("rating", "Average rating", 11, "number1",
            formula='=IF({id}{r}="","",IFERROR(AVERAGEIFS('
                    '[[Jobs.rating]],' + _JOBS + ',{id}{r},'
                    '[[Jobs.rating]],">0"),""))',
            cf=[{"type": "cell", "criteria": "<", "value": 3.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 4.5,
                 "fmt": "good"}]),
        Col("complaints", "Complaints", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Complaints.client]],'
                    '{id}{r}))',
            cf=[{"type": "cell", "criteria": ">=", "value": 2,
                 "fmt": "warn"}]),
        rank_desc("invoiced", guard='{name}{r}=""',
                  note="1 = your biggest client by money billed."),
        rank_desc("profit", guard='{name}{r}=""',
                  note="1 = your most profitable client."),
        rank_desc("year_value", guard='{name}{r}=""',
                  note="1 = the client worth most over a year."),
        Col("retention", "Standing", 12, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Lost","Lost",'
                'IF({status}{r}="Inactive","Dormant",'
                'IF({last_job}{r}="","Dormant",'
                'IF({days_since}{r}>InactiveClientDays*2,"Dormant",'
                'IF({days_since}{r}>InactiveClientDays,"At risk",'
                'IF({next_job}{r}="","Slipping",'
                'IF({jobs_done}{r}>=10,"Loyal","Steady")))))))'),
            cf=status_cf(),
            note="Worked out from how long it has been and whether "
                 "anything is in the diary. 'Slipping' means they are "
                 "still a client but nothing is booked."),
        Col("risk", "Risk score", 9, "number1",
            formula='=IF(OR({name}{r}="",{retention}{r}="Lost",'
                    '{retention}{r}="Dormant"),"",'
                    'IF(OR({retention}{r}="At risk",'
                    '{retention}{r}="Slipping"),'
                    'IF({days_since}{r}="",999,{days_since}{r}),""))',
            hidden=True,
            note="Days since the last clean, but only for clients "
                 "who are drifting. The Retention tab ranks on it."),
        Col("risk_rank", "Risk order", 9, "int",
            formula='=IF({risk}{r}="","",COUNTIFS({risk_first}:'
                    '{risk_last},">"&{risk}{r})+COUNTIFS('
                    '{risk_first}:{risk}{r},{risk}{r}))',
            hidden=True, note="1 = the client you should ring first."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({outstanding}{r}>0.5,"Chase - "&TEXT({outstanding}{r},'
                '"0")&" owing",'
                'IF(AND({status}{r}<>"Lost",{status}{r}<>"Inactive",'
                '{status}{r}<>"Paused",{next_job}{r}="",'
                '{frequency}{r}<>"One-off"),"Book - nothing in the diary",'
                'IF(AND({contract_end}{r}<>"",{contract_end}{r}-TodayDate'
                '<=RenewalWarningDays,{contract_end}{r}>=TodayDate),'
                '"Due - contract renews soon",'
                'IF({retention}{r}="At risk","Call - going quiet",'
                'IF({complaints}{r}>=2,"Check - complaints logged",'
                'IF(AND({margin}{r}<>"",{margin}{r}<MarginFloor,'
                '{jobs_done}{r}>0),"Check - thin margin","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Client Database",
        subtitle="Every customer, what they have spent, what they earn "
                 "you and when you last saw them.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).clients,
        help_text="Add the client here first, then their address on the "
                  "Properties tab, then book the work on the Jobs tab. "
                  "Everything downstream keys off the client ID in "
                  "column A.",
        extra_notes=[
            "The grey columns are all worked out for you - jobs done, "
            "money owing, profit, average rating and whether the "
            "client is drifting. There is nothing to update by hand.",
            "'Value a year' multiplies the average job by how often "
            "they book. Sort on it and you can see at a glance which "
            "five clients your business actually rests on.",
            "Client records hold names, addresses, phone numbers and "
            "sometimes keys or alarm codes. Keep the file somewhere "
            "sensible, and delete what you no longer need.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/properties.py`

*217 lines · 1 function*

> `Properties` - the places you actually clean.
>
> A client can have one address or twenty. This tab holds the detail that
> decides how long a job takes and how much it should cost: size, number
> of bathrooms, what state it is in, how you get in and where to park.
> The estimated hours and the suggested price are worked out from those,
> using the rates on the Settings tab.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Properties').

````python
"""`Properties` - the places you actually clean.

A client can have one address or twenty. This tab holds the detail that
decides how long a job takes and how much it should cost: size, number
of bathrooms, what state it is in, how you get in and where to park.
The estimated hours and the suggested price are worked out from those,
using the rates on the Settings tab.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_client,
                     status_cf)

NAME = "Properties"

_JOBS = "[[Jobs.property]]"


def spec(cfg: Config) -> SheetSpec:
    area_header = f"Area ({cfg.area_unit})"
    columns = [
        id_col("name", "PropertyPrefix", digits=4, width=10,
               note="The ID the jobs, quotes and checklists refer to."),
        Col("name", "Property", 28, "text",
            note="A name you will recognise on a job sheet: "
                 "'Harper - family home', 'Willow Park - Block A'."),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientIDs"),
            note="Pick the client from the drop down. Everything else "
                 "about them is pulled through."),
        Col("client_name", "Client", 22, "text",
            formula=lookup_client("name"),
            note="Automatic, from the Clients tab."),
        Col("type", "Property type", 15, "text",
            validate=list_from("PropertyTypeList")),
        Col("address", "Address", 36, "wrap"),
        Col("area", area_header, 11, "number",
            validate=number_check(0, 1000000),
            note="Rough floor area. Used by the area based pricing "
                 "method - pace it out once and write it down."),
        Col("beds", "Bedrooms", 9, "int",
            validate=number_check(0, 50, decimal=False)),
        Col("baths", "Bathrooms", 10, "int",
            validate=number_check(0, 50, decimal=False),
            note="Bathrooms and kitchens are where the time goes."),
        Col("rooms", "Other rooms", 11, "int",
            validate=number_check(0, 200, decimal=False)),
        Col("floors", "Floor types", 14, "text",
            note="Carpet, hardwood, tile, mixed - it changes what you "
                 "take with you."),
        Col("condition", "Condition", 20, "text",
            validate=list_from("ConditionList"),
            note="The honest answer. It multiplies the estimated hours "
                 "- the factors are on the Lists tab."),
        Col("pets", "Pets", 12, "text",
            note="Pet hair is real extra time, and the cleaner needs "
                 "to know before they open the door."),
        Col("access", "How we get in", 20, "text",
            validate=list_from("AccessList")),
        Col("key_status", "Key", 14, "status",
            validate=list_from("KeyStatusList"), cf=status_cf(),
            note="If you hold keys, keep this up to date. It is the "
                 "first thing anyone asks when a key goes missing."),
        Col("entry_notes", "Entry instructions", 32, "wrap",
            note="Lockbox position, alarm panel, which door sticks. "
                 "Treat codes as sensitive - see the note below the "
                 "table."),
        Col("parking", "Parking", 16, "text",
            validate=list_from("ParkingList"),
            note="Nothing wastes a morning like a permit you did not "
                 "know about."),
        Col("distance", f"Distance ({cfg.distance_unit})", 10, "number1",
            validate=number_check(0, 10000),
            note="One way, from your base. The travel cost on a quote "
                 "doubles it for the round trip."),
        Col("checklist", "Checklist", 18, "text",
            validate=list_from("ServiceNames"),
            note="Which checklist the cleaner should work through at "
                 "this property. Set up on the Checklists tab."),
        Col("supplies_onsite", "Client supplies?", 11, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Yes if the client provides the products - it comes "
                 "off the job cost."),
        Col("active", "Active?", 8, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("notes", "Notes", 38, "wrap"),
        # -------------------------------------------------- the estimate --
        Col("size_hours", "Size hours", 10, "number1",
            formula=blank(
                "name",
                'MAX(IF({beds}{r}="",0,{beds}{r})*0.45+'
                'IF({baths}{r}="",0,{baths}{r})*0.6+'
                'IF({rooms}{r}="",0,{rooms}{r})*0.3,'
                'IF({area}{r}="",0,{area}{r})/650)'),
            hidden=True,
            note="Rooms and bathrooms, or floor area - whichever "
                 "suggests the bigger job."),
        Col("est_hours", "Estimated hours", 13, "number1",
            formula=blank(
                "name",
                'ROUND(MAX(1,{size_hours}{r})*IFERROR(INDEX('
                'ConditionFactors,MATCH({condition}{r},ConditionNames,0)),'
                '1),2)'),
            note="How long one cleaner would take. Split it between "
                 "two cleaners and you have the time on site.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("visit_hours", "Hours on site", 11, "number1",
            formula=blank(
                "name",
                'ROUND({est_hours}{r}/MAX(1,CleanersPerJob),2)'),
            note="Estimated hours shared between the standard number "
                 "of cleaners from the Settings tab."),
        Col("suggested", "Suggested price", 13, "money",
            formula=blank(
                "name",
                'MAX(MinimumCharge,ROUND({est_hours}{r}*'
                'IF(IFERROR(INDEX([[Clients.rate]],MATCH({client}{r},'
                '[[Clients.id]],0)),0)>0,INDEX([[Clients.rate]],'
                'MATCH({client}{r},[[Clients.id]],0)),HourlyRate)+'
                'IF({pets}{r}="",0,IF({pets}{r}="None",0,PetSurcharge)),'
                '2))'),
            note="Hours times the rate, with the pet surcharge and the "
                 "minimum charge applied. A starting point for the "
                 "quote, not a contract.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("labour_cost", "Labour cost", 11, "money",
            formula=blank("name", 'ROUND({est_hours}{r}*CostPerHour,2)'),
            hidden=True),
        Col("est_margin", "Estimated margin", 12, "percent0",
            formula=blank(
                "name",
                'IF({suggested}{r}=0,"",({suggested}{r}-{labour_cost}{r}-'
                '{est_hours}{r}*SuppliesPerHour)/{suggested}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}]),
        # ------------------------------------------------- the real world --
        Col("jobs", "Jobs here", 9, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _JOBS + ',{id}{r}))'),
        Col("last_visit", "Last visit", 12, "date",
            formula='=IF({id}{r}="","",IF(' +
                    max_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
                           '([[Jobs.status]]="Completed")+'
                           '([[Jobs.status]]="Invoiced")') +
                    '=0,"",' +
                    max_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
                           '([[Jobs.status]]="Completed")+'
                           '([[Jobs.status]]="Invoiced")') + '))'),
        Col("actual_hours", "Average actual hours", 13, "number1",
            formula='=IF({id}{r}="","",IFERROR(AVERAGEIFS('
                    '[[Jobs.actual_hours]],' + _JOBS + ',{id}{r},'
                    '[[Jobs.actual_hours]],">0"),""))',
            note="What the job really takes, averaged over every visit. "
                 "Compare it with the estimate - that gap is your "
                 "profit or your loss."),
        Col("variance", "Estimate vs actual", 12, "number1",
            formula='=IF(OR({actual_hours}{r}="",{est_hours}{r}=""),"",'
                    '{actual_hours}{r}-{est_hours}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 0.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}],
            note="Positive means the job takes longer than you price "
                 "for. Re-quote it or re-scope it."),
        Col("revenue", "Revenue here", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Jobs.price]],' + _JOBS +
                    ',{id}{r},[[Jobs.status]],"Invoiced")+SUMIFS('
                    '[[Jobs.price]],' + _JOBS + ',{id}{r},'
                    '[[Jobs.status]],"Completed"))'),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({client}{r}="","Add - no client linked",'
                'IF(IFERROR(INDEX([[Clients.name]],MATCH({client}{r},'
                '[[Clients.id]],0)),"")="","Check - client ID not found",'
                'IF(AND({area}{r}="",{beds}{r}=""),'
                '"Add - size needed to price it",'
                'IF(AND({variance}{r}<>"",{variance}{r}>0.5),'
                '"Check - runs over the estimate",'
                'IF({access}{r}="","Add - how do we get in?","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Property & Site Database",
        subtitle="The places you clean, how long they take and what "
                 "they should be worth.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).properties,
        help_text="One row per address. Fill in the size, the condition "
                  "and how you get in, and the workbook estimates the "
                  "hours and suggests a price straight away.",
        extra_notes=[
            "The estimate is deliberately conservative: it takes the "
            "larger of 'rooms and bathrooms' or 'floor area divided by "
            "650', then multiplies by the condition factor. Compare it "
            "with the actual hours column after a month and adjust the "
            "factors on the Lists tab until they match your reality.",
            "Entry instructions, alarm codes and key locations are "
            "sensitive. Anyone you send this file to can read them - "
            "share printouts of a single job sheet instead, and keep "
            "codes out of the file altogether if you can.",
            "A property that keeps running over its estimate is not a "
            "bad cleaner, it is a bad price. The variance column tells "
            "you which ones to re-quote.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/leads.py`

*148 lines · 1 function*

> `Leads` - every enquiry, from first phone call to won or lost.
>
> The pipeline tab. Most cleaning businesses lose more work to a quote
> nobody chased than to a price that was too high, so this tab counts the
> days an enquiry has been sitting and tells you who to ring today.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Leads').

````python
"""`Leads` - every enquiry, from first phone call to won or lost.

The pipeline tab. Most cleaning businesses lose more work to a quote
nobody chased than to a price that was too high, so this tab counts the
days an enquiry has been sitting and tells you who to ring today.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, status_cf

NAME = "Leads"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "LeadPrefix", digits=4, width=10),
        Col("date", "Enquiry date", 12, "date", validate=date_check(),
            note="The day they got in touch. The clock starts here."),
        Col("name", "Name / company", 24, "text"),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 24, "text"),
        Col("address", "Address or area", 28, "wrap"),
        Col("type", "Client type", 16, "text",
            validate=list_from("ClientTypeList")),
        Col("source", "Came from", 17, "text",
            validate=list_from("LeadSourceList"),
            note="Fill this in every single time. It is the only way "
                 "to find out which advertising actually works."),
        Col("service", "Interested in", 22, "text",
            validate=list_from("ServiceNames")),
        Col("frequency", "Frequency wanted", 14, "text",
            validate=list_from("FrequencyList")),
        Col("stage", "Stage", 14, "status",
            validate=list_from("LeadStageList"), cf=status_cf(),
            note="Move it along as you go. Won creates a client; Lost "
                 "wants a reason in the next column."),
        Col("value", "Estimated value a year", 13, "money",
            validate=number_check(0, 1000000),
            note="Rough annual value. Sort on it to decide who to "
                 "chase first."),
        Col("quote", "Quote sent", 11, "id", validate=list_from("QuoteIDs"),
            note="Link the quote once you have raised it on the "
                 "Quotes tab."),
        Col("visit_date", "Visit booked", 11, "date",
            validate=date_check()),
        Col("contacted", "Last contacted", 12, "date",
            validate=date_check(),
            note="Update it every time you speak to them. The chase "
                 "list counts from this date."),
        Col("next_action", "Next action", 24, "text",
            note="'Call Tuesday after 6', 'send revised price'. "
                 "Specific beats vague."),
        Col("next_date", "Next action due", 12, "date",
            validate=date_check()),
        Col("client", "Became client", 11, "id",
            validate=list_from("ClientIDs"),
            note="Once they book, add them to the Clients tab and put "
                 "their ID here."),
        Col("lost_reason", "If lost, why", 20, "text",
            validate=list_from("LostReasonList"),
            note="Be honest with yourself here. Three 'price too "
                 "high' in a row is information."),
        Col("notes", "Notes", 36, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("days_open", "Days open", 10, "int",
            formula='=IF(OR({name}{r}="",{date}{r}=""),"",'
                    'IF(OR({stage}{r}="Won",{stage}{r}="Lost",'
                    '{stage}{r}="Not a fit"),"",TodayDate-{date}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 14,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 7,
                 "fmt": "warn"}]),
        Col("days_quiet", "Days since anything happened", 11, "int",
            formula='=IF({name}{r}="","",TodayDate-MAX('
                    'IF({date}{r}="",0,{date}{r}),'
                    'IF({visit_date}{r}="",0,{visit_date}{r}),'
                    'IF({next_date}{r}="",0,IF({next_date}{r}>TodayDate,'
                    '0,{next_date}{r})),'
                    'IF({quote}{r}="",0,IFERROR(INDEX('
                    '[[Quotes.sent_date]],MATCH({quote}{r},'
                    '[[Quotes.id]],0)),0))))',
            hidden=True,
            note="Days since the enquiry, the visit, the quote going "
                 "out or the last action - whichever is latest."),
        Col("quote_value", "Quote total", 12, "money",
            formula='=IF({quote}{r}="","",IFERROR(INDEX('
                    '[[Quotes.total]],MATCH({quote}{r},[[Quotes.id]],0)),'
                    '""))'),
        Col("quote_status", "Quote status", 13, "status",
            formula='=IF({quote}{r}="","",IFERROR(INDEX('
                    '[[Quotes.status]],MATCH({quote}{r},[[Quotes.id]],0)),'
                    '""))',
            cf=status_cf()),
        Col("won", "Won?", 8, "flag",
            formula=blank("name",
                          'IF({stage}{r}="Won","Yes",'
                          'IF({stage}{r}="Lost","No",""))'),
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({stage}{r}="Won","OK - won",'
                'IF(OR({stage}{r}="Lost",{stage}{r}="Not a fit"),'
                '"Closed",'
                'IF(AND({next_date}{r}<>"",{next_date}{r}<TodayDate),'
                '"Overdue - action was due "&TEXT({next_date}{r},'
                '"d mmm"),'
                'IF(AND({next_date}{r}<>"",{next_date}{r}=TodayDate),'
                '"Today - "&{next_action}{r},'
                'IF(AND({quote}{r}<>"",{days_quiet}{r}<>"",'
                '{days_quiet}{r}>=QuoteFollowupDays),'
                '"Chase - quote sent, no answer",'
                'IF(AND({quote}{r}="",{days_open}{r}<>"",'
                '{days_open}{r}>=2),"Action - no quote sent yet",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Leads & Sales Pipeline",
        subtitle="Every enquiry, how long it has been waiting and who "
                 "needs a phone call today.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).leads,
        help_text="Log the enquiry the day it arrives, even the ones "
                  "you are not sure about. Set a next action and a "
                  "date, and the Action column becomes your call list.",
        extra_notes=[
            "The win rate on the Reports tab is counted from the "
            "Stage column, so keep it honest - a lead parked at "
            "'Following up' for three months is a Lost.",
            "'Came from' is worth more than it looks. After fifty "
            "enquiries you will know exactly which of your adverts "
            "is paying for itself.",
            "Nothing here emails or texts the client. The workbook "
            "tells you who to contact; the phone call is still "
            "yours to make.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/services.py`

*182 lines · 1 function*

> `Services` - the price list.
>
> Every job you sell, what it costs you to deliver and what it earns.
> The quote calculator, the quote lines and the invoice lines all read
> their prices from here, so putting your rates up is a one-cell job.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Services').

````python
"""`Services` - the price list.

Every job you sell, what it costs you to deliver and what it earns.
The quote calculator, the quote lines and the invoice lines all read
their prices from here, so putting your rates up is a one-cell job.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, rank_desc)

NAME = "Services"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ServicePrefix", digits=3, width=9,
               note="Reference only - quotes and jobs pick the service "
                    "by name."),
        Col("name", "Service", 26, "text",
            note="What the client sees on the quote. Keep the names "
                 "short and obvious."),
        Col("category", "Category", 20, "text",
            validate=list_from("ServiceCategoryList")),
        Col("model", "Priced by", 18, "text",
            validate=list_from("PricingModelList"),
            note="How this service is priced. The quote calculator "
                 "asks for a different number for each model."),
        Col("unit", "Unit", 10, "text",
            note="hour, visit, room, oven, sq ft - whatever the price "
                 "below is per."),
        Col("price", "Price per unit", 12, "money",
            validate=number_check(0, 1000000),
            note="Leave blank on an hourly service to use the standard "
                 "rate from the Settings tab."),
        Col("hours", "Hours (one cleaner)", 12, "number1",
            validate=number_check(0, 200),
            note="Typical time for one cleaner. Used to cost the job "
                 "and to block the diary out."),
        Col("uplift", "Uplift", 9, "percent0",
            validate=number_check(0, 5),
            note="Extra percentage for the harder jobs - deep cleans "
                 "and end of tenancy work."),
        Col("supplies", "Supplies cost", 11, "money",
            validate=number_check(0, 100000),
            note="What the products for this service cost you. Blank "
                 "falls back to the per hour figure on Settings."),
        Col("minimum", "Minimum charge", 12, "money",
            validate=number_check(0, 100000),
            note="Blank uses the business minimum from Settings."),
        Col("popular", "Popular?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Shows up first on the quote calculator's list of "
                 "ideas."),
        Col("active", "Offered?", 9, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf(),
            note="Set to No to retire a service without losing its "
                 "history."),
        Col("skills", "Who can do it", 20, "text",
            note="Training or kit the job needs. The scheduler cannot "
                 "enforce it, but you can read it before you assign."),
        Col("equipment", "Equipment needed", 20, "text"),
        Col("description", "What it includes", 46, "wrap",
            note="Paste this straight onto a quote. Being specific "
                 "about what is and is not included prevents most "
                 "arguments."),
        # ------------------------------------------------- the economics --
        Col("standard_price", "Standard price", 13, "money",
            formula=blank(
                "name",
                'ROUND(IF({model}{r}="Per hour",'
                'IF({price}{r}="",HourlyRate,{price}{r})*'
                'IF({hours}{r}="",1,{hours}{r}),'
                'IF({price}{r}="",0,{price}{r}))*'
                '(1+IF({uplift}{r}="",0,{uplift}{r})),2)'),
            note="What a standard booking of this service comes to "
                 "before any property specific adjustment.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("labour_cost", "Labour cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({hours}{r}="",0,{hours}{r})*CostPerHour,2)'),
            note="Hours times the true cost of a cleaner hour (pay "
                 "plus on-costs) from the Settings tab."),
        Col("supply_cost", "Supply cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({supplies}{r}="",IF({hours}{r}="",0,'
                '{hours}{r})*SuppliesPerHour,{supplies}{r}),2)'),
            hidden=True),
        Col("total_cost", "Cost to deliver", 12, "money",
            formula=blank("name",
                          '{labour_cost}{r}+{supply_cost}{r}')),
        Col("profit", "Profit", 11, "money",
            formula=blank("name",
                          '{standard_price}{r}-{total_cost}{r}'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}]),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({name}{r}="",{standard_price}{r}=0),"",'
                    '{profit}{r}/{standard_price}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}],
            note="The honest test of a price. Anything under the floor "
                 "on the Settings tab is work you are doing for "
                 "nothing."),
        Col("hourly_yield", "Earns an hour", 11, "money",
            formula='=IF(OR({name}{r}="",{hours}{r}="",{hours}{r}=0),"",'
                    '{profit}{r}/{hours}{r})',
            note="Profit per hour on site. Sort on this column and you "
                 "know which work to chase and which to drop.",
            cf=[{"type": "3_color_scale", "min_color": "#FBD9DD",
                 "mid_color": "#FFF0CE", "max_color": "#D7F2E3"}]),
        Col("target_price", "Price for target margin", 13, "money",
            formula='=IF(OR({name}{r}="",TargetMargin>=1),"",'
                    'ROUND({total_cost}{r}/(1-TargetMargin),2))',
            note="What you would have to charge to hit the target "
                 "margin from the Settings tab."),
        # --------------------------------------------------- how it sells --
        Col("times_quoted", "Times quoted", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Quote Lines.service]],{name}{r}))'),
        Col("times_booked", "Times booked", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS([[Jobs.service]],'
                    '{name}{r}))'),
        Col("revenue", "Revenue", 12, "money",
            formula='=IF({name}{r}="","",SUMIFS([[Jobs.price]],'
                    '[[Jobs.service]],{name}{r},[[Jobs.status]],'
                    '"Completed")+SUMIFS([[Jobs.price]],[[Jobs.service]],'
                    '{name}{r},[[Jobs.status]],"Invoiced"))'),
        rank_desc("revenue", guard='{name}{r}=""',
                  note="1 = your biggest earning service."),
        rank_desc("hourly_yield", guard='{name}{r}=""',
                  note="1 = the most profitable use of an hour."),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "name",
                'IF(AND({model}{r}<>"Per hour",{price}{r}=""),'
                '"Add - no price set",'
                'IF({hours}{r}="","Add - how long does it take?",'
                'IF(AND({margin}{r}<>"",{margin}{r}<MarginFloor),'
                '"Below - margin under the floor",'
                'IF(AND({margin}{r}<>"",{margin}{r}<TargetMargin),'
                '"Check - under target margin",'
                'IF({active}{r}="No","Not offered","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Cleaning Service Catalogue",
        subtitle="Your price list, with the margin on every service "
                 "worked out for you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).services,
        help_text="Set a price and the hours it takes, and the workbook "
                  "tells you what the job costs you, what it makes and "
                  "what you would need to charge to hit your target "
                  "margin.",
        extra_notes=[
            "'Earns an hour' is the column that changes how people "
            "run a cleaning business. A $85 oven clean that takes 90 "
            "minutes beats a $150 half day every time.",
            "Prices here are the default. A client on a special rate "
            "overrides it on the Clients tab, and any quote line can "
            "be typed over by hand.",
            "Retire a service with 'Offered? = No' rather than "
            "deleting the row, or the history on old quotes and jobs "
            "loses its name.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/quotes.py`

*185 lines · 1 function*

> `Quotes` - one row per quote, totalled from its lines.
>
> You write the lines on the Quote Lines tab; this tab adds them up,
> applies the discount and the tax, works out the deposit, counts the
> days since you sent it and tells you when to follow up. The win rate
> on the Reports tab is counted from the Status column here.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Quotes').

````python
"""`Quotes` - one row per quote, totalled from its lines.

You write the lines on the Quote Lines tab; this tab adds them up,
applies the discount and the tax, works out the deposit, counts the
days since you sent it and tells you when to follow up. The win rate
on the Reports tab is counted from the Status column here.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, month_col,
                     status_cf)

NAME = "Quotes"

_LINES = "[[Quote Lines.quote]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "QuotePrefix", digits=4, width=10,
               note="Quote number. Put it on the Quote Lines tab to "
                    "add lines to this quote."),
        Col("date", "Quote date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("client", "Client ID", 11, "id", validate=list_from("ClientIDs"),
            note="An existing client. For a brand new enquiry, leave "
                 "it blank and type the name in the next column."),
        Col("client_name", "Client", 22, "text",
            formula='=IF({client}{r}<>"",IFERROR(INDEX([[Clients.name]],'
                    'MATCH({client}{r},[[Clients.id]],0)),'
                    '{prospect}{r}),{prospect}{r})',
            note="The client's name, or the prospect name if they are "
                 "not on the books yet."),
        Col("prospect", "Prospect name", 20, "text",
            note="For a quote to somebody who is not a client yet."),
        Col("property", "Property ID", 11, "id",
            validate=list_from("PropertyIDs")),
        Col("address", "Address", 30, "wrap"),
        Col("service", "Main service", 22, "text",
            validate=list_from("ServiceNames")),
        Col("frequency", "Frequency", 14, "text",
            validate=list_from("FrequencyList")),
        Col("status", "Status", 13, "status",
            validate=list_from("QuoteStatusList"), cf=status_cf(),
            note="Draft, Sent, Followed up, Accepted, Declined. The "
                 "win rate counts Accepted against everything sent."),
        Col("sent_date", "Sent on", 11, "date", validate=date_check(),
            note="The follow-up clock starts here, not on the quote "
                 "date."),
        Col("followed_up", "Last followed up", 12, "date",
            validate=date_check()),
        Col("decision_date", "Accepted / declined on", 12, "date",
            validate=date_check()),
        Col("discount", "Discount", 10, "percent0",
            validate=number_check(0, 1),
            note="A percentage off the lines. First clean discounts "
                 "win a lot of recurring work."),
        Col("deposit_input", "Deposit asked for", 11, "money",
            validate=number_check(0, 1000000),
            note="Blank uses the deposit percentage from Settings."),
        Col("valid_input", "Valid until", 11, "date",
            validate=date_check(),
            note="Blank adds the standard validity from Settings to "
                 "the quote date."),
        Col("job", "Job booked", 11, "id", validate=list_from("JobIDs"),
            note="Once they say yes, book the job and put its ID "
                 "here."),
        Col("decline_reason", "If declined, why", 18, "text",
            validate=list_from("LostReasonList")),
        Col("notes", "Notes", 34, "wrap"),
        Col("terms", "Terms shown", 34, "wrap",
            note="What is included, what is not, cancellation terms. "
                 "This is what prints on the quote."),
        # ------------------------------------------------- the automatics --
        Col("lines", "Lines", 8, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _LINES + ',{id}{r}))',
            hidden=True),
        Col("hours", "Hours quoted", 11, "number1",
            formula='=IF({id}{r}="","",SUMIFS([[Quote Lines.hours]],' +
                    _LINES + ',{id}{r}))'),
        Col("subtotal", "Subtotal", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Quote Lines.total]],' +
                    _LINES + ',{id}{r}))',
            note="Added up from the Quote Lines tab."),
        Col("discount_value", "Less discount", 11, "money",
            formula='=IF({id}{r}="","",ROUND({subtotal}{r}*'
                    'IF({discount}{r}="",0,{discount}{r}),2))'),
        Col("net", "Net", 12, "money",
            formula='=IF({id}{r}="","",{subtotal}{r}-'
                    '{discount_value}{r})'),
        Col("tax", "Tax", 10, "money",
            formula='=IF({id}{r}="","",ROUND({net}{r}*TaxRate,2))'),
        Col("total", "Total", 12, "money",
            formula='=IF({id}{r}="","",{net}{r}+{tax}{r})',
            note="What the client sees. Price the lines, not this "
                 "cell.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("deposit", "Deposit", 11, "money",
            formula='=IF({id}{r}="","",IF({deposit_input}{r}<>"",'
                    '{deposit_input}{r},ROUND({total}{r}*'
                    'DepositPercent,2)))'),
        Col("cost", "Cost to deliver", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Quote Lines.cost]],' +
                    _LINES + ',{id}{r}))',
            hidden=True),
        Col("profit", "Profit", 11, "money",
            formula='=IF({id}{r}="","",{net}{r}-{cost}{r})'),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({id}{r}="",{net}{r}=0),"",'
                    '{profit}{r}/{net}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}],
            note="Check this before you send it. A discount that "
                 "drops the margin under the floor is work you "
                 "cannot afford to win."),
        Col("valid_until", "Valid to", 11, "date",
            formula='=IF(OR({id}{r}="",{date}{r}=""),"",'
                    'IF({valid_input}{r}<>"",{valid_input}{r},'
                    '{date}{r}+QuoteValidDays))'),
        Col("days_waiting", "Days waiting", 11, "int",
            formula='=IF(OR({id}{r}="",{sent_date}{r}=""),"",'
                    'IF(OR({status}{r}="Accepted",{status}{r}="Declined"),'
                    '"",TodayDate-{sent_date}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 7,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 3,
                 "fmt": "warn"}]),
        Col("expired", "Expired?", 9, "flag",
            formula='=IF(OR({id}{r}="",{valid_until}{r}=""),"",'
                    'IF(AND({valid_until}{r}<TodayDate,{status}{r}<>'
                    '"Accepted",{status}{r}<>"Declined"),"Yes","No"))',
            cf=flag_yes_no_cf(), hidden=True),
        Col("won", "Won?", 8, "flag",
            formula=blank("date",
                          'IF({status}{r}="Accepted","Yes",'
                          'IF({status}{r}="Declined","No",""))'),
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({status}{r}="Accepted",IF({job}{r}="",'
                '"Action - accepted, book the job","OK - won"),'
                'IF({status}{r}="Declined","Closed",'
                'IF({lines}{r}=0,"Add - no lines on this quote yet",'
                'IF({status}{r}="Draft","Draft - not sent",'
                'IF({expired}{r}="Yes","Expired - re-quote or close",'
                'IF(AND({days_waiting}{r}<>"",'
                '{days_waiting}{r}>=QuoteFollowupDays),'
                '"Chase - "&{days_waiting}{r}&" days with no answer",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Quotes & Estimates",
        subtitle="Every quote you have sent, what it is worth, what "
                 "it makes you and when to chase it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).quotes,
        help_text="Raise the quote here, then list what is included on "
                  "the Quote Lines tab using this quote's ID. The "
                  "totals, the deposit and the margin all appear by "
                  "themselves.",
        extra_notes=[
            "Use the Pricing & Quote Calculator tab to work out the "
            "price first, then write the quote here. The calculator "
            "shows you the margin before you commit to a number.",
            "A quote with no lines totals zero - that is what the "
            "'Add - no lines' warning means.",
            "A spreadsheet cannot collect a signature. Print the "
            "quote, email it, and record the acceptance date here "
            "when they reply; the email itself is your evidence.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/quote_lines.py`

*140 lines · 1 function*

> `Quote Lines` - what is actually on each quote.
>
> One row per line. Pick the quote, pick the service and type a quantity;
> the price, the hours and the cost come from the Services tab, and the
> line total feeds straight back into the quote.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Quote Lines').

````python
"""`Quote Lines` - what is actually on each quote.

One row per line. Pick the quote, pick the service and type a quantity;
the price, the hours and the cost come from the Services tab, and the
line total feeds straight back into the quote.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, lookup_service

NAME = "Quote Lines"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("quote", '"QL-"', digits=4, width=10),
        Col("quote", "Quote ID", 11, "id", validate=list_from("QuoteIDs"),
            note="Which quote this line belongs to. Pick it from the "
                 "drop down."),
        Col("service", "Service", 24, "text",
            validate=list_from("ServiceNames"),
            note="From the Services tab. The unit price and the hours "
                 "come with it."),
        Col("description", "Description on the quote", 40, "wrap",
            note="Leave blank to use the standard wording from the "
                 "Services tab."),
        Col("qty", "Quantity", 10, "number1",
            validate=number_check(0, 100000),
            note="Hours, rooms, ovens, visits - whatever the service "
                 "is priced per."),
        Col("price_input", "Price each (override)", 12, "money",
            validate=number_check(0, 1000000),
            note="Blank uses the catalogue price."),
        Col("discount", "Discount", 9, "percent0",
            validate=number_check(0, 1)),
        Col("optional", "Optional extra?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Shown on the quote but not counted in the headline "
                 "total until the client says yes. Optional lines "
                 "still add to the subtotal here - take them off the "
                 "quote if they decline."),
        Col("notes", "Line notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("catalogue_price", "Catalogue price", 12, "money",
            formula=lookup_service("standard_price", missing="0"),
            hidden=True),
        Col("unit", "Per", 9, "text",
            formula=lookup_service("unit", missing="")),
        Col("price_used", "Price each", 11, "money",
            formula=blank("quote",
                          'IF({price_input}{r}<>"",{price_input}{r},'
                          'IFERROR({catalogue_price}{r}*1,0))')),
        Col("wording", "Wording used", 40, "wrap",
            formula=blank("quote",
                          'IF({description}{r}<>"",{description}{r},'
                          'IFERROR(INDEX([[Services.description]],'
                          'MATCH({service}{r},[[Services.name]],0)),""))'),
            hidden=True,
            note="What the printed quote says for this line."),
        Col("hours", "Hours", 9, "number1",
            formula=blank(
                "quote",
                'IFERROR(INDEX([[Services.hours]],MATCH({service}{r},'
                '[[Services.name]],0)),0)*IF({qty}{r}="",1,'
                'IF(IFERROR(INDEX([[Services.model]],MATCH({service}{r},'
                '[[Services.name]],0)),"")="Per hour",1,{qty}{r}))'),
            note="Hours for the service, multiplied by the quantity "
                 "unless it is already priced by the hour."),
        Col("gross", "Line value", 12, "money",
            formula=blank("quote",
                          'ROUND(IF({qty}{r}="",1,{qty}{r})*'
                          '{price_used}{r},2)'),
            hidden=True),
        Col("total", "Line total", 12, "money",
            formula=blank("quote",
                          'ROUND({gross}{r}*(1-IF({discount}{r}="",0,'
                          '{discount}{r})),2)'),
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("cost", "Cost", 11, "money",
            formula=blank(
                "quote",
                'ROUND({hours}{r}*CostPerHour+IF({qty}{r}="",1,'
                '{qty}{r})*IFERROR(INDEX([[Services.supply_cost]],'
                'MATCH({service}{r},[[Services.name]],0)),0),2)'),
            hidden=True),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({quote}{r}="",{total}{r}=0),"",'
                    '({total}{r}-{cost}{r})/{total}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}]),
        Col("seq", "Line number", 9, "int",
            formula='=IF(OR({quote}{r}="",{service}{r}=""),"",'
                    'COUNTIFS({quote_first}:{quote}{r},'
                    '{quote}{r}))',
            hidden=True,
            note="1, 2, 3 within each quote. The printable form "
                 "reads it."),
        Col("key", "Lookup key", 12, "text",
            formula='=IF({seq}{r}="","",{quote}{r}&"|"&{seq}{r})',
            hidden=True),
        Col("alert", "Check", 24, "status",
            formula=blank(
                "quote",
                'IF({service}{r}="","Add - pick a service",'
                'IF({qty}{r}="","Add - how many?",'
                'IF({price_used}{r}=0,"Add - no price",'
                'IF(AND({margin}{r}<>"",{margin}{r}<MarginFloor),'
                '"Below - thin margin","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Quote Lines",
        subtitle="The detail behind every quote - one row per thing "
                 "you are charging for.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).quote_lines,
        help_text="Pick the quote ID, pick a service, type a quantity. "
                  "Everything else is filled in from the Services tab "
                  "and totalled back onto the quote.",
        extra_notes=[
            "Quoting a job that is not in your catalogue? Leave the "
            "service blank, write your own description and type the "
            "price - the line still totals correctly.",
            "The margin on each line is worth a glance. A quote can "
            "look healthy overall while one line of it is being done "
            "at a loss.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/jobs.py`

*315 lines · 1 function*

> `Jobs` - the job scheduler, and the spine of the whole workbook.
>
> One row per visit. Pick the client, the property and the service and
> the rest fills itself in: how long it should take, what it should cost
> the client, who is free, when the next visit in the series falls and
> whether you have double-booked anybody.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Jobs').

````python
"""`Jobs` - the job scheduler, and the spine of the whole workbook.

One row per visit. Pick the client, the property and the service and
the rest fills itself in: how long it should take, what it should cost
the client, who is free, when the next visit in the series falls and
whether you have double-booked anybody.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_client,
                     lookup_property, month_col, score_validation,
                     status_cf, week_col, weekday_col)

NAME = "Jobs"

# ---------------------------------------------------------------- lookups --
_SVC_HOURS = ('IFERROR(INDEX([[Services.hours]],MATCH({service}{r},'
              '[[Services.name]],0)),0)')
_SVC_MODEL = ('IFERROR(INDEX([[Services.model]],MATCH({service}{r},'
              '[[Services.name]],0)),"")')
_SVC_PRICE = ('IFERROR(INDEX([[Services.standard_price]],MATCH('
              '{service}{r},[[Services.name]],0)),0)')
_PROP_HOURS = ('IFERROR(INDEX([[Properties.est_hours]],MATCH('
               '{property}{r},[[Properties.id]],0)),0)')
_CLIENT_RATE = ('IFERROR(INDEX([[Clients.rate]],MATCH({client}{r},'
                '[[Clients.id]],0)),0)')
_FREQ_FACTOR = ('IFERROR(INDEX(FreqFactors,MATCH({frequency}{r},'
                'FreqNames,0)),1)')
_FREQ_DAYS = ('IFERROR(INDEX(FreqDays,MATCH({frequency}{r},FreqNames,0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("client", "JobPrefix", digits=4, width=11,
               note="The job reference. Checklists, timesheets, "
                    "invoices and the costing tab all key off it."),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day of the visit. Everything in the calendar, "
                 "the rota and the reminders reads this column."),
        weekday_col("date"),
        week_col("date"),
        month_col("date"),
        Col("start", "Start", 9, "time", validate=time_check(),
            note="Type 9:00 AM or 14:30 - either works."),
        Col("client", "Client ID", 11, "id", validate=list_from("ClientIDs"),
            note="Pick from the drop down. The client's name, their "
                 "rate and their terms all follow."),
        Col("client_name", "Client", 22, "text",
            formula=lookup_client("name")),
        Col("client_type", "Client type", 16, "text",
            formula=lookup_client("type", missing=""), hidden=True,
            note="Residential or commercial, pulled from the client. "
                 "The Reports tab splits the money by it."),
        Col("property", "Property ID", 11, "id",
            validate=list_from("PropertyIDs"),
            note="Which address. Leave blank for a client with only "
                 "one property and the price still works."),
        Col("property_name", "Property", 24, "text",
            formula=lookup_property("name", missing="")),
        Col("address", "Address", 30, "text",
            formula=lookup_property("address", missing=""),
            note="Pulled through for the job sheet and the cleaner's "
                 "list."),
        Col("service", "Service", 22, "text",
            validate=list_from("ServiceNames"),
            note="From the Services tab. The price and the time come "
                 "from there unless you override them."),
        Col("type", "Job type", 16, "text",
            validate=list_from("JobTypeList")),
        Col("frequency", "Frequency", 14, "text",
            validate=list_from("FrequencyList"),
            note="One-off, weekly, fortnightly. Drives the discount "
                 "and the next visit date."),
        Col("status", "Status", 13, "status",
            validate=list_from("JobStatusList"), cf=status_cf(),
            note="Move it along as the day goes: Booked, Confirmed, "
                 "In progress, Completed. Invoiced happens when you "
                 "raise the invoice."),
        Col("priority", "Priority", 11, "text",
            validate=list_from("PriorityList")),
        Col("employee", "Cleaner", 18, "text",
            validate=list_from("EmployeeNames"),
            note="The lead cleaner. A second name goes in the next "
                 "column; anyone else goes on the timesheet."),
        Col("employee2", "Second cleaner", 18, "text",
            validate=list_from("EmployeeNames")),
        Col("cleaners", "Cleaners", 9, "int",
            validate=number_check(1, 20, decimal=False),
            note="Blank uses the standard crew size from the Settings "
                 "tab."),
        Col("hours_input", "Planned hours (override)", 12, "number1",
            validate=number_check(0, 100),
            note="Total cleaner-hours. Leave blank and the workbook "
                 "uses the service or the property estimate."),
        Col("price_input", "Price (override)", 12, "money",
            validate=number_check(0, 1000000),
            note="Leave blank to let the calculator price it. Type a "
                 "number to fix the price for this visit only."),
        Col("travel", f"Travel ({cfg.distance_unit}, round trip)", 11,
            "number1", validate=number_check(0, 10000)),
        Col("recurring", "Recurring?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Yes means the workbook should suggest the next "
                 "visit date when this one is done."),
        Col("confirmed", "Confirmed with client?", 11, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("actual_input", "Actual hours (if no timesheet)", 12,
            "number1", validate=number_check(0, 100),
            note="Only needed if you are not filling in timesheets. "
                 "Timesheet hours always win."),
        Col("rating", "Client rating", 10, "rating",
            validate=score_validation(5),
            note="1 to 5, filled in after the visit. Feeds the client "
                 "and employee averages."),
        Col("feedback", "Feedback", 30, "wrap"),
        Col("notes", "Job notes", 34, "wrap",
            note="What is different about this visit: the spare room "
                 "is out of bounds, the dog is at the vet."),
        # ------------------------------------------------- the automatics --
        Col("est_hours", "Hours (total)", 11, "number1",
            formula=blank(
                "client",
                'IF({hours_input}{r}<>"",{hours_input}{r},'
                'IF(AND(' + _SVC_MODEL + '="Per hour",' + _PROP_HOURS +
                '>0),' + _PROP_HOURS + ',IF(' + _SVC_HOURS + '>0,' +
                _SVC_HOURS + ',' + _PROP_HOURS + ')))'),
            note="Total cleaner-hours the visit should take. Hourly "
                 "services use the property estimate; fixed price "
                 "services use the time on the Services tab."),
        Col("crew", "Crew used", 8, "int",
            formula=blank("client",
                          'IF({cleaners}{r}="",CleanersPerJob,'
                          'MAX(1,{cleaners}{r}))'),
            hidden=True),
        Col("hours_each", "Hours on site", 11, "number1",
            formula=blank("client",
                          'ROUND({est_hours}{r}/{crew}{r},2)'),
            note="How long the visit lasts: total hours shared "
                 "between the crew. This is what blocks out the "
                 "diary.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("finish", "Finish", 9, "time",
            formula='=IF(OR({client}{r}="",{start}{r}="",'
                    '{hours_each}{r}=""),"",{start}{r}+'
                    '{hours_each}{r}/24)',
            note="Start plus the time on site. Add the travel "
                 "allowance from Settings before you book the next "
                 "one."),
        Col("rate_used", "Rate used", 10, "money",
            formula=blank("client",
                          'IF(' + _CLIENT_RATE + '>0,' + _CLIENT_RATE +
                          ',HourlyRate)'),
            hidden=True,
            note="The client's agreed rate if they have one, "
                 "otherwise the standard rate."),
        Col("base_price", "Base price", 11, "money",
            formula=blank(
                "client",
                'IF(AND({service}{r}<>"",' + _SVC_PRICE + '>0,' +
                _SVC_MODEL + '<>"Per hour"),' + _SVC_PRICE + ','
                '{est_hours}{r}*{rate_used}{r})'),
            hidden=True),
        Col("price", "Price", 12, "money",
            formula=blank(
                "client",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF({status}{r}="Cancelled",0,'
                'MAX(MinimumCharge,ROUND({base_price}{r}*' +
                _FREQ_FACTOR + ',2))))'),
            note="What the client pays. Worked out from the service, "
                 "the hours, the client's rate and the frequency "
                 "discount - unless you typed an override.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("actual_hours", "Actual hours", 11, "number1",
            formula=blank(
                "client",
                'IF(SUMIFS([[Timesheets.hours]],[[Timesheets.job]],'
                '{id}{r})>0,SUMIFS([[Timesheets.hours]],'
                '[[Timesheets.job]],{id}{r}),'
                'IF({actual_input}{r}="","",{actual_input}{r}))'),
            note="Timesheet hours for this job if there are any, "
                 "otherwise what you typed in the override column."),
        Col("over_run", "Over / under", 10, "number1",
            formula='=IF(OR({actual_hours}{r}="",{est_hours}{r}=""),"",'
                    '{actual_hours}{r}-{est_hours}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 0.25,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}],
            note="Hours over the estimate. Consistently positive "
                 "means the price is wrong, not the cleaner."),
        Col("tasks", "Checklist", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Job Tasks.job]],'
                    '{id}{r}))',
            hidden=True),
        Col("tasks_done", "Tasks done", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Job Tasks.job]],'
                    '{id}{r},[[Job Tasks.status]],"Done"))',
            hidden=True),
        Col("checklist_pct", "Checklist complete", 11, "percent0",
            formula='=IF(OR({id}{r}="",{tasks}{r}=0),"",'
                    '{tasks_done}{r}/{tasks}{r})',
            cf=[{"type": "data_bar", "bar_color": "#4E8A63"}],
            note="Filled in from the Job Tasks tab as the cleaner "
                 "ticks things off."),
        Col("invoice", "Invoice", 11, "id",
            formula='=IF({id}{r}="","",IFERROR(INDEX('
                    '[[Invoice Lines.invoice]],MATCH({id}{r},'
                    '[[Invoice Lines.job]],0)),""))',
            note="Found automatically once the job appears on an "
                 "invoice line."),
        Col("next_due", "Next visit due", 12, "date",
            formula='=IF(OR({client}{r}="",{date}{r}="",'
                    '{recurring}{r}<>"Yes"),"",IF(' + _FREQ_DAYS +
                    '=0,"",{date}{r}+' + _FREQ_DAYS + '))',
            note="Date plus the gap for this frequency. Copy the row "
                 "down and paste this date in to book the next one.",
            cf=[{"type": "cell", "criteria": "<", "value": "=TodayDate",
                 "fmt": "warn"}]),
        Col("clash", "Diary", 16, "status",
            formula=blank(
                "client",
                'IF({employee}{r}="","Unassigned",'
                'IF(COUNTIFS({employee_first}:{employee_last},'
                '{employee}{r},{date_first}:{date_last},{date}{r},'
                '{start_first}:{start_last},{start}{r})>1,"Clash",'
                'IF(SUMIFS({hours_each_first}:{hours_each_last},'
                '{employee_first}:{employee_last},{employee}{r},'
                '{date_first}:{date_last},{date}{r})>ProductiveHours,'
                '"Long day","OK")))'),
            cf=alert_cf(),
            note="Two jobs at the same time for the same cleaner is a "
                 "Clash. More than a full day's work is a Long day."),
        Col("future", "Upcoming date", 11, "date",
            formula='=IF(OR({client}{r}="",{date}{r}=""),"",'
                    'IF(AND({date}{r}>=TodayDate,{status}{r}<>'
                    '"Cancelled"),{date}{r},""))',
            hidden=True,
            note="The date, but only for visits still to come. The "
                 "calendar and the dashboard read it."),
        Col("future_rank", "Diary order", 9, "int",
            formula='=IF({future}{r}="","",COUNTIFS({future_first}:'
                    '{future_last},"<"&{future}{r})+COUNTIFS('
                    '{future_first}:{future}{r},{future}{r}))',
            hidden=True,
            note="1 = the next job in the diary. No array formulas "
                 "anywhere in this workbook."),
        Col("unassigned", "Needs a cleaner", 11, "date",
            formula='=IF(OR({client}{r}="",{date}{r}="",'
                    '{employee}{r}<>""),"",IF(AND({date}{r}>=TodayDate,'
                    '{status}{r}<>"Cancelled"),{date}{r},""))',
            hidden=True),
        Col("unassigned_rank", "Assign order", 9, "int",
            formula='=IF({unassigned}{r}="","",COUNTIFS('
                    '{unassigned_first}:{unassigned_last},"<"&'
                    '{unassigned}{r})+COUNTIFS({unassigned_first}:'
                    '{unassigned}{r},{unassigned}{r}))',
            hidden=True,
            note="1 = the soonest job with nobody on it."),
        Col("days_away", "Days away", 9, "int",
            formula='=IF(OR({client}{r}="",{date}{r}=""),"",'
                    '{date}{r}-TodayDate)',
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "client",
                'IF({status}{r}="Cancelled","Cancelled",'
                'IF({status}{r}="No access","Action - no access, rebook",'
                'IF(AND({date}{r}<>"",{days_away}{r}<0,'
                'OR({status}{r}="Booked",{status}{r}="Confirmed",'
                '{status}{r}="In progress")),"Late - mark it completed",'
                'IF(AND({status}{r}="Completed",{invoice}{r}=""),'
                '"Action - not invoiced yet",'
                'IF(AND({days_away}{r}>=0,{days_away}{r}<=ReminderDays,'
                '{confirmed}{r}<>"Yes"),"Confirm - visit coming up",'
                'IF({employee}{r}="","Assign - no cleaner booked",'
                'IF(AND({status}{r}="Completed",{checklist_pct}{r}<>"",'
                '{checklist_pct}{r}<1),"Check - checklist unfinished",'
                '"OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Job Scheduler",
        subtitle="Every visit you have booked, priced automatically "
                 "and checked for clashes.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).jobs,
        help_text="Pick a client, a property and a service, give it a "
                  "date and a cleaner, and the hours, the finish time "
                  "and the price work themselves out. The Calendar tab "
                  "shows the same rows as a diary.",
        extra_notes=[
            "To book a recurring client, fill one row in, set "
            "'Recurring? = Yes' and read the next visit date in the "
            "last columns - copy the row down and paste that date in. "
            "A spreadsheet cannot create rows on its own, so this is "
            "the honest way to do it.",
            "The Diary column is a real double-booking check: it "
            "looks for the same cleaner, the same date and the same "
            "start time anywhere on this tab.",
            "A job only becomes revenue when its status is Completed "
            "or Invoiced. Cancelled jobs are priced at zero and drop "
            "out of every total.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/job_tasks.py`

*120 lines · 1 function*

> `Job Tasks` - the checklist as it was actually done, job by job.
>
> Copy the tasks for the service onto a job, tick them off as you go,
> and the job's completion percentage, the quality score and the list of
> issues found all fill themselves in. This is the tab that turns "it was
> cleaned" into something you can show a client.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Job Tasks').

````python
"""`Job Tasks` - the checklist as it was actually done, job by job.

Copy the tasks for the service onto a job, tick them off as you go,
and the job's completion percentage, the quality score and the list of
issues found all fill themselves in. This is the tab that turns "it was
cleaned" into something you can show a client.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_job,
                     status_cf)

NAME = "Job Tasks"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("job", '"JT-"', digits=4, width=10),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Which visit. Everything about the job is pulled "
                 "through from the Jobs tab."),
        Col("date", "Date", 11, "date",
            formula=lookup_job("date", missing="")),
        Col("client_name", "Client", 20, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))'),
        Col("property_name", "Property", 22, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.property_name]],MATCH({job}{r},[[Jobs.id]],'
                    '0)),""))'),
        Col("area", "Area", 16, "text", validate=list_from("TaskAreaList")),
        Col("task", "Task", 38, "text",
            note="Copy these from the Checklists tab for the service "
                 "you are doing - filter, copy, paste."),
        Col("status", "Status", 13, "status",
            validate=list_from("TaskStatusList"), cf=status_cf(),
            note="Done, Skipped, Issue found. The job's completion "
                 "percentage counts the Done ones."),
        Col("employee", "Done by", 16, "text",
            validate=list_from("EmployeeNames"),
            note="Blank means the lead cleaner on the job."),
        Col("minutes", "Minutes taken", 10, "int",
            validate=number_check(0, 600, decimal=False)),
        Col("quality", "Quality", 13, "status",
            validate=list_from("QualityList"), cf=status_cf(),
            note="Filled in on a spot check, not on every task. Only "
                 "the ones you look at."),
        Col("issue", "Issue found", 32, "wrap",
            note="A dripping tap, a stain that will not shift, a "
                 "light bulb gone. Tell the client before they find "
                 "it."),
        Col("reported", "Told the client?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("photo", "Photo link", 20, "link",
            note="A spreadsheet cannot hold the photo itself tidily - "
                 "paste a cloud link."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("required", "Required?", 9, "flag",
            formula='=IF({task}{r}="","",IFERROR(INDEX('
                    '[[Checklists.required]],MATCH({task}{r},'
                    '[[Checklists.task]],0)),"Yes"))',
            hidden=True),
        Col("planned", "Planned minutes", 10, "int",
            formula='=IF({task}{r}="","",IFERROR(INDEX('
                    '[[Checklists.minutes]],MATCH({task}{r},'
                    '[[Checklists.task]],0)),""))',
            hidden=True),
        Col("over", "Over / under (min)", 10, "int",
            formula='=IF(OR({minutes}{r}="",{planned}{r}=""),"",'
                    '{minutes}{r}-{planned}{r})',
            hidden=True),
        Col("done", "Complete?", 9, "flag",
            formula=blank("job",
                          'IF({status}{r}="Done","Yes","No")'),
            cf=flag_yes_no_cf(), hidden=True),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "job",
                'IF({task}{r}="","Add - what was the task?",'
                'IF({status}{r}="Issue found",IF({reported}{r}="Yes",'
                '"OK - issue reported","Action - tell the client"),'
                'IF({status}{r}="Needs rework","Action - rework needed",'
                'IF(AND({status}{r}="Skipped",{required}{r}="Yes"),'
                '"Check - required task skipped",'
                'IF({status}{r}="Done","OK","Due")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Job Tasks & Completion",
        subtitle="The checklist as it was really done - ticked off, "
                 "timed and with the problems written down.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).job_tasks,
        help_text="Filter the Checklists tab to the service you are "
                  "doing, copy the tasks, paste them here against the "
                  "job ID, and tick them off. The job's completion "
                  "percentage updates as you go.",
        extra_notes=[
            "You do not have to use this tab for every visit. Most "
            "businesses use it for deep cleans, end of tenancy work "
            "and anywhere a client has complained before.",
            "'Issue found' plus a photo link is the cheapest "
            "insurance there is. It proves the damage was already "
            "there.",
            "Minutes taken against planned minutes is how you find "
            "out that the oven really takes 90 minutes, not 60.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/checklists.py`

*153 lines · 1 function*

> `Checklists` - the master task lists behind every clean.
>
> Build one checklist per service (and, where it matters, per property)
> and you have a standard. The cleaner works through it on the Job Tasks
> tab; the minutes here add up to a realistic time for the visit, which
> is what turns a guess into a price.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Checklists').

````python
"""`Checklists` - the master task lists behind every clean.

Build one checklist per service (and, where it matters, per property)
and you have a standard. The cleaner works through it on the Job Tasks
tab; the minutes here add up to a realistic time for the visit, which
is what turns a guess into a price.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Checklists"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("task", '"CHK-"', digits=4, width=10),
        Col("template", "Checklist for", 24, "text",
            validate=list_from("ServiceNames"),
            note="Which service this task belongs to. The job sheet "
                 "prints every task for the service you booked."),
        Col("property", "Only at this property", 13, "id",
            validate=list_from("PropertyIDs"),
            note="Leave blank for a task that applies everywhere. "
                 "Fill it in for the one client who wants the bins "
                 "put out."),
        Col("area", "Area", 18, "text", validate=list_from("TaskAreaList"),
            note="Group the tasks by room so the cleaner works "
                 "through the property in a sensible order."),
        Col("order", "Order", 7, "int",
            validate=number_check(1, 999, decimal=False),
            note="Sort on Area then Order to print the checklist in "
                 "the order you want it done."),
        Col("task", "Task", 40, "text",
            note="One job, written as an instruction: 'Descale the "
                 "shower head', not 'bathroom'."),
        Col("detail", "How it should look", 36, "wrap",
            note="The standard. This is what settles an argument "
                 "about whether something was done properly."),
        Col("minutes", "Minutes", 9, "int",
            validate=number_check(0, 600, decimal=False),
            note="Honest minutes for one cleaner. The total at the "
                 "bottom of this tab is how long the service really "
                 "takes."),
        Col("frequency", "How often", 14, "text",
            validate=list_from("FrequencyList"),
            note="Every visit, or only weekly / monthly / quarterly. "
                 "Deep tasks do not need doing every time."),
        Col("required", "Required?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Required tasks are the ones that make the job "
                 "'complete'. Extras are nice to have."),
        Col("photo", "Photo needed?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Useful for end of tenancy and changeover work - "
                 "proof is cheaper than an argument."),
        Col("products", "Products / kit", 22, "text"),
        Col("active", "In use?", 8, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("hours", "Hours", 8, "number1",
            formula=blank("task",
                          'ROUND(IF({minutes}{r}="",0,{minutes}{r})/60,2)'),
            hidden=True),
        Col("service_total", "Total for this checklist", 12, "number1",
            formula='=IF({template}{r}="","",ROUND(SUMIFS('
                    '{minutes_first}:{minutes_last},'
                    '{template_first}:{template_last},{template}{r},'
                    '{active_first}:{active_last},"Yes")/60,2))',
            note="Every active task for this service, in hours. "
                 "Compare it with the hours on the Services tab - if "
                 "they disagree, one of them is wrong.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("service_hours", "Service allows", 11, "number1",
            formula='=IF({template}{r}="","",IFERROR(INDEX('
                    '[[Services.hours]],MATCH({template}{r},'
                    '[[Services.name]],0)),""))',
            hidden=True),
        Col("times_done", "Times ticked off", 11, "int",
            formula='=IF({task}{r}="","",COUNTIFS([[Job Tasks.task]],'
                    '{task}{r},[[Job Tasks.status]],"Done"))'),
        Col("times_skipped", "Times skipped", 11, "int",
            formula='=IF({task}{r}="","",COUNTIFS([[Job Tasks.task]],'
                    '{task}{r},[[Job Tasks.status]],"Skipped"))',
            cf=[{"type": "cell", "criteria": ">=", "value": 3,
                 "fmt": "warn"}],
            note="A task skipped again and again is either not "
                 "needed or not possible in the time you allow."),
        Col("issues", "Issues found", 10, "int",
            formula='=IF({task}{r}="","",COUNTIFS([[Job Tasks.task]],'
                    '{task}{r},[[Job Tasks.status]],"Issue found"))'),
        Col("status", "Standard", 13, "status",
            formula=blank(
                "task",
                'IF({times_done}{r}+{times_skipped}{r}=0,"New",'
                'IF({times_skipped}{r}/({times_done}{r}+'
                '{times_skipped}{r})>0.25,"Often skipped","Working"))'),
            cf=status_cf(), hidden=True),
        Col("seq", "Line number", 9, "int",
            formula='=IF(OR({template}{r}="",{task}{r}=""),"",'
                    'COUNTIFS({template_first}:{template}{r},'
                    '{template}{r}))',
            hidden=True,
            note="1, 2, 3 within each checklist. The printable form "
                 "reads it."),
        Col("key", "Lookup key", 12, "text",
            formula='=IF({seq}{r}="","",{template}{r}&"|"&{seq}{r})',
            hidden=True),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "task",
                'IF({template}{r}="","Add - which service?",'
                'IF({minutes}{r}="","Add - how many minutes?",'
                'IF(AND({service_hours}{r}<>"",{service_total}{r}>'
                '{service_hours}{r}+0.25),'
                '"Check - checklist is longer than the service allows",'
                'IF({times_skipped}{r}>=3,"Check - often skipped",'
                '"OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Cleaning Checklists",
        subtitle="The standard for every service, room by room, with "
                 "honest minutes against each task.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).checklists,
        help_text="Write the checklist once per service. Add a "
                  "property ID only for the tasks that apply to one "
                  "address. The Print Job Sheet tab turns it into "
                  "something the cleaner can carry.",
        extra_notes=[
            "Adding up the minutes is the single most useful thing "
            "on this tab. Most underpriced cleans are underpriced "
            "because nobody ever counted the tasks.",
            "The 'Checklist is longer than the service allows' "
            "warning means your price and your standard disagree. "
            "Change one of them.",
            "Tasks ticked off on the Job Tasks tab feed back here, "
            "so you can see which parts of your standard actually "
            "get done.",
        ],
    )
````

---

## 4. The people and money registers

Who did the work, what it cost and what came back in: staff, timesheets, job costing, invoices and their lines, payments, supplies, equipment, expenses, mileage, complaints, incidents and compliance documents.

### `cleaning_business_spreadsheet/sheets/employees.py`

*208 lines · 1 function*

> `Employees` - the team.
>
> Who works for you, what they cost, what they are trained on and how
> much of the work they are carrying. The hours, the pay and the revenue
> each cleaner brings in are counted from the Timesheets and Jobs tabs,
> so this tab is where you find out - without arguing about it - who is
> overloaded and who is standing still.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Employees').

````python
"""`Employees` - the team.

Who works for you, what they cost, what they are trained on and how
much of the work they are carrying. The hours, the pay and the revenue
each cleaner brings in are counted from the Timesheets and Jobs tabs,
so this tab is where you find out - without arguing about it - who is
overloaded and who is standing still.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, rank_desc,
                     status_cf)

NAME = "Employees"

_T = "[[Timesheets.employee]]"
_J = "[[Jobs.employee]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "EmployeePrefix", digits=3, width=9),
        Col("name", "Name", 20, "text",
            note="The name that appears in every job and rota drop "
                 "down. Keep it exactly the same everywhere."),
        Col("role", "Role", 18, "text",
            validate=list_from("EmployeeRoleList")),
        Col("type", "Employment", 14, "text",
            validate=list_from("EmploymentTypeList"),
            note="Subcontractors are costed differently - they invoice "
                 "you rather than appearing on payroll."),
        Col("status", "Status", 12, "status",
            validate=list_from("EmployeeStatusList"), cf=status_cf()),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 26, "text"),
        Col("emergency", "Emergency contact", 22, "text",
            note="Worth having for anyone working alone in someone "
                 "else's house."),
        Col("start", "Started", 11, "date", validate=date_check()),
        Col("pay", "Pay an hour", 11, "money",
            validate=number_check(0, 1000),
            note="Blank uses the standard pay rate from Settings. The "
                 "true cost adds your on-costs on top."),
        Col("hours", "Contracted hours a week", 12, "number1",
            validate=number_check(0, 100),
            note="Used to work out how full their week is."),
        Col("skills", "Trained on", 24, "text",
            note="Deep cleans, carpet machine, out of hours, driving. "
                 "Check it before you assign the job."),
        Col("areas", "Areas covered", 18, "text"),
        Col("vehicle", "Vehicle", 12, "text",
            validate=list_from("VehicleList")),
        Col("licence", "Driver?", 8, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("checked", "Background check?", 11, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Essential for schools, nurseries and vulnerable "
                 "clients. The expiry goes on the Safety tab."),
        Col("insured", "Own insurance?", 11, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Subcontractors should have their own. Keep a copy."),
        Col("uniform", "Uniform issued", 11, "text"),
        Col("notes", "Notes", 34, "wrap"),
        # ------------------------------------------------ what they cost --
        Col("pay_rate", "Pay used", 11, "money",
            formula=blank("name", 'IF({pay}{r}="",PayRate,{pay}{r})'),
            hidden=True),
        Col("cost_rate", "True cost an hour", 12, "money",
            formula=blank(
                "name",
                'ROUND({pay_rate}{r}*IF({type}{r}="Subcontractor",1,'
                '1+EmployerOncost),2)'),
            note="Pay plus your on-costs - tax, insurance, holiday "
                 "pay. A subcontractor invoices you, so no on-costs."),
        # ------------------------------------------------- what they do --
        Col("jobs_assigned", "Jobs assigned", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS(' + _J + ',{name}{r}))'),
        Col("jobs_done", "Jobs done", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS(' + _J + ',{name}{r},'
                    '[[Jobs.status]],"Completed")+COUNTIFS(' + _J +
                    ',{name}{r},[[Jobs.status]],"Invoiced"))'),
        Col("upcoming", "Jobs ahead", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS(' + _J + ',{name}{r},'
                    '[[Jobs.date]],">="&TodayDate,[[Jobs.status]],'
                    '"<>Cancelled"))',
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("week_hours", "Hours booked this week", 12, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Jobs.hours_each]],' + _J +
                    ',{name}{r},[[Jobs.date]],">="&WeekStartDate,'
                    '[[Jobs.date]],"<="&WeekStartDate+6,[[Jobs.status]],'
                    '"<>Cancelled"))',
            note="Straight from the diary. Compare it with the "
                 "contracted hours in the next column.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("load", "Week full", 9, "percent0",
            formula='=IF(OR({name}{r}="",{hours}{r}="",{hours}{r}=0),"",'
                    '{week_hours}{r}/{hours}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "warn"}],
            note="Over 100% means you have booked them for more than "
                 "they are contracted to work."),
        Col("paid_hours", "Hours paid (all time)", 12, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Timesheets.hours]],' +
                    _T + ',{name}{r}))'),
        Col("billed_hours", "Hours billed", 11, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Timesheets.hours]],' +
                    _T + ',{name}{r},[[Timesheets.billable]],"Yes"))'),
        Col("utilisation", "Utilisation", 10, "percent0",
            formula='=IF(OR({name}{r}="",{paid_hours}{r}=0),"",'
                    '{billed_hours}{r}/{paid_hours}{r})',
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=UtilisationTarget", "fmt": "good"},
                {"type": "cell", "criteria": "<",
                 "value": "=UtilisationTarget", "fmt": "warn"}],
            note="The share of the hours you pay for that a client "
                 "actually pays for. Travel, setup and waiting are the "
                 "usual culprits."),
        Col("wage_cost", "Wages paid", 12, "money",
            formula='=IF({name}{r}="","",SUMIFS([[Timesheets.cost]],' + _T +
                    ',{name}{r}))'),
        Col("revenue", "Revenue brought in", 13, "money",
            formula='=IF({name}{r}="","",SUMIFS([[Jobs.price]],' + _J +
                    ',{name}{r},[[Jobs.status]],"Completed")+SUMIFS('
                    '[[Jobs.price]],' + _J + ',{name}{r},[[Jobs.status]],'
                    '"Invoiced"))'),
        Col("contribution", "Contribution", 12, "money",
            formula='=IF({name}{r}="","",{revenue}{r}-{wage_cost}{r})',
            note="Revenue from their jobs less their wages. Not a "
                 "complete picture - overheads sit on top - but it is "
                 "the right shape.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("rating", "Average rating", 11, "number1",
            formula='=IF({name}{r}="","",IFERROR(AVERAGEIFS('
                    '[[Jobs.rating]],' + _J + ',{name}{r},[[Jobs.rating]],'
                    '">0"),""))',
            cf=[{"type": "cell", "criteria": "<", "value": 3.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 4.5,
                 "fmt": "good"}]),
        Col("complaints", "Complaints", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Complaints.employee]],{name}{r}))',
            cf=[{"type": "cell", "criteria": ">=", "value": 2,
                 "fmt": "warn"}]),
        Col("late", "Late or absent", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Timesheets.attendance]],"Late",' + _T + ',{name}{r})'
                    '+COUNTIFS([[Timesheets.attendance]],'
                    '"Absent - no show",' + _T + ',{name}{r}))',
            cf=[{"type": "cell", "criteria": ">=", "value": 3,
                 "fmt": "warn"}]),
        Col("name_rank", "List order", 8, "int",
            formula='=IF({name}{r}="","",COUNTIF({name_first}:'
                    '{name}{r},"?*"))',
            hidden=True,
            note="1, 2, 3 down the filled rows. The rota reads it."),
        rank_desc("revenue", guard='{name}{r}=""',
                  note="1 = brings in the most revenue."),
        rank_desc("week_hours", guard='{name}{r}=""',
                  note="1 = the fullest week. Used by the rota."),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Left","Archived",'
                'IF(AND({load}{r}<>"",{load}{r}>1),'
                '"Overloaded - over contracted hours",'
                'IF({complaints}{r}>=2,"Check - complaints logged",'
                'IF(AND({rating}{r}<>"",{rating}{r}<3.5),'
                '"Check - low client ratings",'
                'IF({checked}{r}="No","Action - no background check",'
                'IF(AND({upcoming}{r}=0,{status}{r}="Active"),'
                '"Idle - nothing booked","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Employee Database",
        subtitle="Your team, what they cost, what they bring in and "
                 "how full their week is.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).employees,
        help_text="Add everyone who cleans for you, including yourself "
                  "and any subcontractors. Their names fill the drop "
                  "downs on the Jobs, rota and timesheet tabs.",
        extra_notes=[
            "'True cost an hour' is the number to quote from, not the "
            "pay rate. An $18 cleaner on 18% on-costs really costs you "
            "$21.24 before they have put a cloth in a bucket.",
            "Utilisation is the quiet killer in a cleaning business. "
            "Paying for eight hours and billing five is a 37% hole in "
            "your margin, and it never shows up on a bank statement.",
            "Employee records are personal data. Only keep what you "
            "actually need, and keep the file somewhere only you can "
            "get at.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/timesheets.py`

*144 lines · 1 function*

> `Timesheets` - hours worked, what they cost and whether they were billed.
>
> One row per cleaner per job (or per shift). The hours come from the
> start and finish times less the break, the cost comes from that
> person's true hourly cost on the Employees tab, and the billable flag
> is what makes the utilisation figure on the dashboard mean something.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Timesheets').

````python
"""`Timesheets` - hours worked, what they cost and whether they were billed.

One row per cleaner per job (or per shift). The hours come from the
start and finish times less the break, the cost comes from that
person's true hourly cost on the Employees tab, and the billable flag
is what makes the utilisation figure on the dashboard mean something.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import (date_check, list_from, number_check, time_check)
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, month_col,
                     status_cf, week_col, weekday_col)

NAME = "Timesheets"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("employee", '"TS-"', digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        weekday_col("date"),
        week_col("date"),
        month_col("date"),
        Col("employee", "Cleaner", 18, "text",
            validate=list_from("EmployeeNames")),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Leave blank for time that is not against a job - "
                 "training, admin, a supply run."),
        Col("client_name", "Client", 20, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))'),
        Col("activity", "Activity", 16, "text",
            validate=list_from("MileagePurposeList"),
            note="Only needed for non-job time. Blank means cleaning."),
        Col("start", "Start", 9, "time", validate=time_check()),
        Col("finish", "Finish", 9, "time", validate=time_check()),
        Col("break_min", "Break (min)", 10, "int",
            validate=number_check(0, 480, decimal=False)),
        Col("hours_input", "Hours (override)", 11, "number1",
            validate=number_check(0, 24),
            note="Only if you are not recording start and finish "
                 "times. Otherwise leave it blank."),
        Col("overtime", "Of which overtime", 10, "number1",
            validate=number_check(0, 24),
            note="Hours paid at the overtime multiplier from the "
                 "Settings tab."),
        Col("attendance", "Attendance", 16, "status",
            validate=list_from("AttendanceList"), cf=status_cf()),
        Col("billable", "Billable?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Yes for time a client pays for. Travel, training "
                 "and admin are No - and that gap is your "
                 "utilisation."),
        Col("approved", "Approved?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("paid", "Paid?", 8, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("clock_hours", "Clocked", 9, "number1",
            formula='=IF(OR({start}{r}="",{finish}{r}=""),"",'
                    'ROUND(({finish}{r}-{start}{r})*24-'
                    'IF({break_min}{r}="",0,{break_min}{r})/60,2))',
            hidden=True,
            note="Finish less start, less the break."),
        Col("hours", "Hours", 9, "number1",
            formula=blank("employee",
                          'IF({hours_input}{r}<>"",{hours_input}{r},'
                          'IF({clock_hours}{r}="",0,{clock_hours}{r}))'),
            note="The hours everything else counts: pay, job cost, "
                 "utilisation.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("rate", "Cost an hour", 10, "money",
            formula=blank(
                "employee",
                'IFERROR(INDEX([[Employees.cost_rate]],MATCH('
                '{employee}{r},[[Employees.name]],0)),'
                'PayRate*(1+EmployerOncost))'),
            hidden=True,
            note="That person's true cost from the Employees tab."),
        Col("basic_hours", "Basic hours", 9, "number1",
            formula=blank("employee",
                          'MAX(0,{hours}{r}-IF({overtime}{r}="",0,'
                          '{overtime}{r}))'),
            hidden=True),
        Col("cost", "Cost to you", 11, "money",
            formula=blank(
                "employee",
                'ROUND({basic_hours}{r}*{rate}{r}+IF({overtime}{r}="",0,'
                '{overtime}{r})*{rate}{r}*OvertimeRate,2)'),
            note="Basic hours plus overtime at the multiplier, at "
                 "that cleaner's true cost per hour."),
        Col("billed_value", "Billed value", 11, "money",
            formula=blank(
                "employee",
                'IF({billable}{r}<>"Yes",0,ROUND({hours}{r}*'
                'IF(IFERROR(INDEX([[Clients.rate]],MATCH(IFERROR(INDEX('
                '[[Jobs.client]],MATCH({job}{r},[[Jobs.id]],0)),""),'
                '[[Clients.id]],0)),0)>0,INDEX([[Clients.rate]],'
                'MATCH(INDEX([[Jobs.client]],MATCH({job}{r},[[Jobs.id]],'
                '0)),[[Clients.id]],0)),HourlyRate),2))'),
            hidden=True,
            note="What these hours are worth at the client's rate."),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "employee",
                'IF({hours}{r}=0,"Add - no hours recorded",'
                'IF({attendance}{r}="Absent - no show","Action - no show",'
                'IF({hours}{r}>10,"Check - over ten hours in a day",'
                'IF({approved}{r}<>"Yes","Due - needs approving",'
                'IF({paid}{r}<>"Yes","Due - not paid yet","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Timesheets & Attendance",
        subtitle="Hours worked, what they cost you and how much of "
                 "them a client actually paid for.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).timesheets,
        help_text="One row per cleaner per job. Record the start and "
                  "finish times and the hours, the cost and the job "
                  "costing all follow.",
        extra_notes=[
            "Hours entered here override the estimate on the job, so "
            "the profit figures are based on what really happened "
            "rather than what you hoped.",
            "Mark travel, training and admin as not billable. The "
            "difference between paid hours and billed hours is the "
            "single biggest hidden cost in a cleaning business.",
            "This is a record of hours, not a payroll system. It "
            "will not calculate tax, and it is not a substitute for "
            "whatever your country requires you to keep.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/job_costing.py`

*166 lines · 2 functions*

> `Job Costing` - what every job really made, line by line.
>
> There is nothing to type on this tab. Each row mirrors the job on the
> same row of the Jobs tab and works out the wages, the supplies, the
> travel and the profit. It is the tab that answers the question most
> cleaning businesses never ask: which of these jobs is actually worth
> doing?

| Defines | What it does |
|---------|--------------|
| `_from_job(key, fallback)` |  |
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Job Costing').

````python
"""`Job Costing` - what every job really made, line by line.

There is nothing to type on this tab. Each row mirrors the job on the
same row of the Jobs tab and works out the wages, the supplies, the
travel and the profit. It is the tab that answers the question most
cleaning businesses never ask: which of these jobs is actually worth
doing?
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..writer import number_check
from .common import alert_cf, blank, month_col, rank_desc, status_cf

NAME = "Job Costing"

#: the job on the same row of the Jobs tab
_MIRROR = 'IFERROR(INDEX([[Jobs.id]],ROW()-4),"")'


def _from_job(key: str, fallback: str = '""') -> str:
    return ('=IF({job}{r}="","",IFERROR(INDEX([[Jobs.' + key + ']],'
            'MATCH({job}{r},[[Jobs.id]],0)),' + fallback + '))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("job", "Job ID", 11, "id",
            formula='=IF(' + _MIRROR + '="","",' + _MIRROR + ')',
            note="Mirrors the Jobs tab row for row. Nothing to type."),
        Col("date", "Date", 11, "date", formula=_from_job("date")),
        month_col("date"),
        Col("client", "Client ID", 11, "id", formula=_from_job("client")),
        Col("client_name", "Client", 20, "text",
            formula=_from_job("client_name")),
        Col("property_name", "Property", 20, "text",
            formula=_from_job("property_name")),
        Col("service", "Service", 20, "text", formula=_from_job("service")),
        Col("employee", "Cleaner", 16, "text",
            formula=_from_job("employee")),
        Col("status", "Status", 13, "status", formula=_from_job("status"),
            cf=status_cf()),
        # ----------------------------------------------------- the money --
        Col("revenue", "Revenue", 12, "money",
            formula='=IF({job}{r}="","",IF(OR({status}{r}="Completed",'
                    '{status}{r}="Invoiced"),IFERROR(INDEX('
                    '[[Jobs.price]],MATCH({job}{r},[[Jobs.id]],0)),0),0))',
            note="Only counted once the job is Completed or Invoiced "
                 "- a booking is not money.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("est_hours", "Hours quoted", 10, "number1",
            formula=_from_job("est_hours", "0"), hidden=True),
        Col("actual_hours", "Hours worked", 10, "number1",
            formula=_from_job("actual_hours", "0"),
            note="From the timesheets if you keep them, otherwise "
                 "the figure typed on the job."),
        Col("hours_used", "Hours costed", 10, "number1",
            formula='=IF({job}{r}="","",IF(IFERROR({actual_hours}{r}*1,0)'
                    '>0,{actual_hours}{r},IFERROR({est_hours}{r}*1,0)))',
            hidden=True),
        Col("timesheet_cost", "Wages from timesheets", 11, "money",
            formula='=IF({job}{r}="","",SUMIFS([[Timesheets.cost]],'
                    '[[Timesheets.job]],{job}{r}))',
            hidden=True),
        Col("labour", "Wages", 11, "money",
            formula='=IF({job}{r}="","",IF({timesheet_cost}{r}>0,'
                    '{timesheet_cost}{r},ROUND({hours_used}{r}*'
                    'CostPerHour,2)))',
            note="Real timesheet cost where you have one, otherwise "
                 "the hours times the true cost of a cleaner hour."),
        Col("supply_logged", "Supplies logged", 11, "money",
            formula='=IF({job}{r}="","",SUMIFS([[Supply Usage.cost]],'
                    '[[Supply Usage.job]],{job}{r}))',
            hidden=True),
        Col("supplies", "Supplies", 11, "money",
            formula='=IF({job}{r}="","",IF({supply_logged}{r}>0,'
                    '{supply_logged}{r},ROUND({hours_used}{r}*'
                    'SuppliesPerHour,2)))',
            note="What you actually used if you logged it on the "
                 "Supply Usage tab, otherwise the per hour allowance "
                 "from Settings."),
        Col("travel", "Travel", 10, "money",
            formula='=IF({job}{r}="","",ROUND(IFERROR(INDEX('
                    '[[Jobs.travel]],MATCH({job}{r},[[Jobs.id]],0)),0)*'
                    'MileageRate,2))'),
        Col("other", "Other cost", 10, "money",
            validate=number_check(0, 1000000),
            note="The only cell you ever type on this tab: parking, "
                 "a skip, a replacement item, a subcontractor's "
                 "invoice."),
        Col("cost", "Total cost", 12, "money",
            formula=blank("job",
                          '{labour}{r}+{supplies}{r}+{travel}{r}+'
                          'IF({other}{r}="",0,{other}{r})')),
        Col("profit", "Profit", 12, "money",
            formula=blank("job", '{revenue}{r}-{cost}{r}'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}]),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({job}{r}="",{revenue}{r}=0),"",'
                    '{profit}{r}/{revenue}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}]),
        Col("per_hour", "Profit an hour", 11, "money",
            formula='=IF(OR({job}{r}="",{hours_used}{r}=0,'
                    '{revenue}{r}=0),"",{profit}{r}/{hours_used}{r})',
            note="The fairest comparison between a two hour office "
                 "and a six hour deep clean.",
            cf=[{"type": "3_color_scale", "min_color": "#FBD9DD",
                 "mid_color": "#FFF0CE", "max_color": "#D7F2E3"}]),
        Col("hours_over", "Hours over estimate", 10, "number1",
            formula='=IF(OR({job}{r}="",{actual_hours}{r}="",'
                    '{actual_hours}{r}=0,{est_hours}{r}=""),"",'
                    '{actual_hours}{r}-{est_hours}{r})',
            hidden=True),
        rank_desc("profit", guard='{job}{r}=""',
                  note="1 = the most profitable job."),
        rank_desc("per_hour", guard='{job}{r}=""',
                  note="1 = the best use of an hour."),
        Col("alert", "Verdict", 28, "status",
            formula=blank(
                "job",
                'IF({revenue}{r}=0,"",'
                'IF({profit}{r}<0,"Loss - this job cost you money",'
                'IF({margin}{r}<MarginFloor,"Below - margin under the '
                'floor",IF({margin}{r}<TargetMargin,'
                '"Check - under target margin",'
                'IF(AND({hours_over}{r}<>"",{hours_over}{r}>0.5),'
                '"Check - ran over the estimate","OK")))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap",
            note="Why this one went the way it did. Worth writing "
                 "down while you remember."),
    ]

    return SheetSpec(
        name=NAME,
        title="Job Costing & Profitability",
        subtitle="Every job, every cost, every margin - worked out "
                 "for you. Only two columns here accept typing.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        help_text="This tab mirrors the Jobs tab row for row. Add any "
                  "extra cost in the 'Other cost' column and read the "
                  "profit, the margin and the profit per hour.",
        extra_notes=[
            "Wages come from the timesheets when you keep them and "
            "from the estimate when you do not - so the numbers are "
            "useful from day one and get more accurate as you fill "
            "the workbook in.",
            "Sort by 'Profit an hour' and you will see something "
            "uncomfortable and useful: the jobs that feel biggest "
            "are rarely the ones that pay best.",
            "A single loss-making job is nothing. The same client "
            "appearing three times in the red is a conversation "
            "about price.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/invoices.py`

*175 lines · 1 function*

> `Invoices` - what you have billed and what is still owed.
>
> One row per invoice, totalled from its lines, with the due date worked
> out from the client's terms and the amount received pulled from the
> Payments tab. The overdue flag is the one that gets you paid.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Invoices').

````python
"""`Invoices` - what you have billed and what is still owed.

One row per invoice, totalled from its lines, with the due date worked
out from the client's terms and the amount received pulled from the
Payments tab. The overdue flag is the one that gets you paid.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_client,
                     month_col, status_cf)

NAME = "Invoices"

_LINES = "[[Invoice Lines.invoice]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "InvoicePrefix", digits=4, width=11,
               note="The invoice number. Put it on the Invoice Lines "
                    "tab to add lines, and on any payment you "
                    "receive."),
        Col("date", "Invoice date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("client", "Client ID", 11, "id", validate=list_from("ClientIDs"),
            note="Everything else about them - name, address, terms - "
                 "comes from the Clients tab."),
        Col("client_name", "Client", 22, "text",
            formula=lookup_client("name")),
        Col("address", "Billing address", 30, "text",
            formula=lookup_client("address", missing=""), hidden=True),
        Col("period", "Covers", 20, "text",
            note="'Visits in March', 'End of tenancy, 12B Lakeview'. "
                 "What the client will recognise."),
        Col("status", "Status", 13, "status",
            validate=list_from("InvoiceStatusList"), cf=status_cf(),
            note="Draft until you send it. Paid is worked out for "
                 "you from the payments, but you can override it."),
        Col("sent_date", "Sent on", 11, "date", validate=date_check(),
            note="The due date counts from the invoice date, but it "
                 "is worth recording when it really went out."),
        Col("terms_input", "Terms (days)", 10, "int",
            validate=number_check(0, 180, decimal=False),
            note="Blank uses the client's terms, or the business "
                 "default from Settings."),
        Col("discount", "Discount", 9, "percent0",
            validate=number_check(0, 1)),
        Col("method", "Pay by", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("po", "PO / reference", 14, "text",
            note="Commercial clients often will not pay without "
                 "their own reference on the invoice."),
        Col("reminders", "Reminders sent", 10, "int",
            validate=number_check(0, 50, decimal=False)),
        Col("last_reminder", "Last reminder", 11, "date",
            validate=date_check()),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("lines", "Lines", 7, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _LINES + ',{id}{r}))',
            hidden=True),
        Col("subtotal", "Subtotal", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Invoice Lines.total]],' +
                    _LINES + ',{id}{r}))'),
        Col("discount_value", "Less discount", 11, "money",
            formula='=IF({id}{r}="","",ROUND({subtotal}{r}*'
                    'IF({discount}{r}="",0,{discount}{r}),2))',
            hidden=True),
        Col("net", "Net", 12, "money",
            formula='=IF({id}{r}="","",{subtotal}{r}-'
                    '{discount_value}{r})'),
        Col("tax", "Tax", 10, "money",
            formula='=IF({id}{r}="","",ROUND({net}{r}*TaxRate,2))'),
        Col("total", "Total", 12, "money",
            formula='=IF({id}{r}="","",{net}{r}+{tax}{r})',
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}],
            note="What the client owes. Edit the lines, not this "
                 "cell."),
        Col("paid", "Received", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Payments.amount]],'
                    '[[Payments.invoice]],{id}{r}))',
            note="Added up from the Payments tab."),
        Col("outstanding", "Still owed", 12, "money",
            formula='=IF({id}{r}="","",ROUND({total}{r}-{paid}{r},2))',
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}]),
        Col("terms_used", "Terms used", 9, "int",
            formula=blank(
                "date",
                'IF({terms_input}{r}<>"",{terms_input}{r},'
                'IF(IFERROR(INDEX([[Clients.terms]],MATCH({client}{r},'
                '[[Clients.id]],0)),"")<>"",INDEX([[Clients.terms]],'
                'MATCH({client}{r},[[Clients.id]],0)),PaymentTerms))'),
            hidden=True),
        Col("due_date", "Due", 11, "date",
            formula='=IF(OR({id}{r}="",{date}{r}=""),"",'
                    '{date}{r}+{terms_used}{r})'),
        Col("days_late", "Days late", 9, "int",
            formula='=IF(OR({id}{r}="",{due_date}{r}=""),"",'
                    'IF({outstanding}{r}<=0.004,"",'
                    'MAX(0,TodayDate-{due_date}{r}-OverdueGrace)))',
            cf=[{"type": "cell", "criteria": ">", "value": 30,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("age", "Age bracket", 12, "status",
            formula=blank(
                "date",
                'IF({outstanding}{r}<=0.004,"Paid",'
                'IF({days_late}{r}="","Not due yet",'
                'IF({days_late}{r}=0,"Due now",'
                'IF({days_late}{r}<=30,"1-30 days",'
                'IF({days_late}{r}<=60,"31-60 days",'
                'IF({days_late}{r}<=90,"61-90 days","90+ days"))))))'),
            cf=status_cf(),
            note="The ageing buckets the Reports tab counts."),
        Col("late_fee", "Late fee due", 10, "money",
            formula='=IF(OR({id}{r}="",{days_late}{r}=""),"",'
                    'IF({days_late}{r}<=0,0,ROUND({outstanding}{r}*'
                    'LateFeePercent,2)))',
            hidden=True,
            note="Only charge it if your terms say you can."),
        Col("fully_paid", "Settled?", 9, "flag",
            formula=blank("date",
                          'IF({outstanding}{r}<=0.004,"Yes","No")'),
            cf=flag_yes_no_cf(), hidden=True),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "date",
                'IF({lines}{r}=0,"Add - no lines on this invoice yet",'
                'IF({status}{r}="Written off","Written off",'
                'IF({fully_paid}{r}="Yes","OK - paid",'
                'IF({status}{r}="Draft","Draft - not sent yet",'
                'IF(AND({days_late}{r}<>"",{days_late}{r}>30),'
                '"Overdue - "&{days_late}{r}&" days, call them",'
                'IF(AND({days_late}{r}<>"",{days_late}{r}>0),'
                '"Overdue - send a reminder",'
                'IF({paid}{r}>0,"Part paid","Due - waiting for '
                'payment")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Invoices",
        subtitle="What you have billed, what has come in and what is "
                 "late.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).invoices,
        help_text="Raise the invoice here, list the jobs on the "
                  "Invoice Lines tab, and record the money on the "
                  "Payments tab. The due date, the balance and the "
                  "chase list look after themselves.",
        extra_notes=[
            "The fastest way to invoice a month of visits: filter "
            "the Jobs tab to one client and status Completed, then "
            "add one invoice line per job with the job ID.",
            "Once a job appears on an invoice line, the Jobs tab "
            "shows the invoice number against it automatically - so "
            "nothing gets billed twice or forgotten.",
            "The workbook cannot take a payment or email a "
            "reminder. It tells you exactly who to chase and for "
            "how much; the sending is still yours.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/invoice_lines.py`

*115 lines · 1 function*

> `Invoice Lines` - the detail on each invoice.
>
> Usually one line per job: pick the invoice, pick the job, and the
> date, the description and the price come through from the scheduler.
> Add a line without a job for anything else you are charging for.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Invoice Lines').

````python
"""`Invoice Lines` - the detail on each invoice.

Usually one line per job: pick the invoice, pick the job, and the
date, the description and the price come through from the scheduler.
Add a line without a job for anything else you are charging for.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, id_col, lookup_job

NAME = "Invoice Lines"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("invoice", '"IL-"', digits=4, width=10),
        Col("invoice", "Invoice ID", 11, "id",
            validate=list_from("InvoiceIDs"),
            note="Which invoice this line belongs to."),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Link the visit you are billing. Leave blank for a "
                 "line that is not a job - a late fee, materials, a "
                 "deposit."),
        Col("date", "Job date", 11, "date",
            formula=lookup_job("date", missing="")),
        Col("description", "Description", 40, "wrap",
            note="Blank uses the service and the date from the job."),
        Col("qty", "Quantity", 9, "number1",
            validate=number_check(0, 100000),
            note="Blank counts as one."),
        Col("price_input", "Price each (override)", 12, "money",
            validate=number_check(0, 1000000),
            note="Blank uses the price on the job."),
        Col("taxable", "Taxable?", 9, "flag",
            validate=list_from("YesNoList"),
            note="Only matters if you charge tax. The invoice total "
                 "applies the rate from Settings to the whole net "
                 "amount."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("job_price", "Job price", 11, "money",
            formula=lookup_job("price", missing="0"), hidden=True),
        Col("job_service", "Service", 20, "text",
            formula=lookup_job("service", missing=""), hidden=True),
        Col("wording", "Wording used", 40, "wrap",
            formula=blank(
                "invoice",
                'IF({description}{r}<>"",{description}{r},'
                'IF({job}{r}="","",{job_service}{r}&" - "&'
                'TEXT({date}{r},"d mmm yyyy")))'),
            hidden=True,
            note="What the printed invoice says for this line."),
        Col("price_used", "Price each", 11, "money",
            formula=blank("invoice",
                          'IF({price_input}{r}<>"",{price_input}{r},'
                          'IFERROR({job_price}{r}*1,0))')),
        Col("total", "Line total", 12, "money",
            formula=blank("invoice",
                          'ROUND(IF({qty}{r}="",1,{qty}{r})*'
                          '{price_used}{r},2)'),
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("duplicate", "Billed twice?", 11, "status",
            formula=blank(
                "invoice",
                'IF({job}{r}="","",IF(COUNTIFS({job_first}:{job_last},'
                '{job}{r})>1,"Duplicate","OK"))'),
            cf=alert_cf(),
            note="Catches the same job appearing on two invoice "
                 "lines - the most common billing mistake there is."),
        Col("seq", "Line number", 9, "int",
            formula='=IF(OR({invoice}{r}="",{description}{r}=""),"",'
                    'COUNTIFS({invoice_first}:{invoice}{r},'
                    '{invoice}{r}))',
            hidden=True,
            note="1, 2, 3 within each invoice. The printable form "
                 "reads it."),
        Col("key", "Lookup key", 12, "text",
            formula='=IF({seq}{r}="","",{invoice}{r}&"|"&{seq}{r})',
            hidden=True),
        Col("alert", "Check", 24, "status",
            formula=blank(
                "invoice",
                'IF({duplicate}{r}="Duplicate",'
                '"Duplicate - job billed twice",'
                'IF({price_used}{r}=0,"Add - no price on this line",'
                'IF(AND({job}{r}="",{description}{r}=""),'
                '"Add - what is this line for?","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Invoice Lines",
        subtitle="One row per job (or per extra) on each invoice.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).invoice_lines,
        help_text="Pick the invoice, pick the job. The description "
                  "and the price come from the scheduler, and the "
                  "total goes straight back onto the invoice.",
        extra_notes=[
            "The 'Billed twice?' column is worth its place on its "
            "own. Billing the same visit on two invoices is both "
            "embarrassing and expensive.",
            "Charging a late fee or a deposit? Leave the job blank, "
            "write your own description and type the amount.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/payments.py`

*116 lines · 1 function*

> `Payments` - the money that actually arrived.
>
> Every payment against an invoice. The invoice balance, the client's
> account and the cash figures on the dashboard all come from here, so
> one row per payment - including part payments - keeps everything
> honest.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Payments').

````python
"""`Payments` - the money that actually arrived.

Every payment against an invoice. The invoice balance, the client's
account and the cash figures on the dashboard all come from here, so
one row per payment - including part payments - keeps everything
honest.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, month_col

NAME = "Payments"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "PaymentPrefix", digits=4, width=11),
        Col("date", "Date received", 12, "date", validate=date_check(),
            note="The day the money landed, not the day you invoiced."),
        month_col("date"),
        Col("invoice", "Invoice ID", 11, "id",
            validate=list_from("InvoiceIDs"),
            note="Which invoice this pays. Part payments get their "
                 "own row against the same invoice."),
        Col("client", "Client ID", 11, "id",
            formula='=IF({invoice}{r}="","",IFERROR(INDEX('
                    '[[Invoices.client]],MATCH({invoice}{r},'
                    '[[Invoices.id]],0)),""))',
            note="Pulled from the invoice."),
        Col("client_name", "Client", 22, "text",
            formula='=IF({invoice}{r}="","",IFERROR(INDEX('
                    '[[Invoices.client_name]],MATCH({invoice}{r},'
                    '[[Invoices.id]],0)),""))'),
        Col("amount", "Amount", 12, "money",
            validate=number_check(0, 10000000),
            note="What arrived in the bank. If a card processor took "
                 "a fee, put the fee in the next column and this "
                 "stays the gross amount."),
        Col("fee", "Processing fee", 10, "money",
            validate=number_check(0, 100000),
            note="Card or platform fee. It is a real cost - the "
                 "Expenses tab has a category for it too."),
        Col("method", "Method", 15, "text",
            validate=list_from("PaymentMethodList")),
        Col("reference", "Reference", 18, "text",
            note="Bank reference, cheque number, transaction ID. "
                 "Worth having when a client says they paid."),
        Col("deposit", "Is this a deposit?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("banked", "Banked?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Cash in a jacket pocket is not banked. This column "
                 "is how you find it again."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("net", "Net received", 11, "money",
            formula=blank("date",
                          'IF({amount}{r}="",0,{amount}{r})-'
                          'IF({fee}{r}="",0,{fee}{r})'),
            note="After the processing fee."),
        Col("invoice_total", "Invoice total", 11, "money",
            formula='=IF({invoice}{r}="","",IFERROR(INDEX('
                    '[[Invoices.total]],MATCH({invoice}{r},'
                    '[[Invoices.id]],0)),""))',
            hidden=True),
        Col("invoice_paid", "Paid on invoice", 11, "money",
            formula='=IF({invoice}{r}="","",IFERROR(INDEX('
                    '[[Invoices.paid]],MATCH({invoice}{r},'
                    '[[Invoices.id]],0)),""))',
            hidden=True),
        Col("balance", "Invoice balance after", 12, "money",
            formula='=IF(OR({invoice}{r}="",{invoice_total}{r}=""),"",'
                    'ROUND({invoice_total}{r}-{invoice_paid}{r},2))',
            note="What is still owed on that invoice once every "
                 "payment against it is counted.",
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": -0.004,
                 "fmt": "bad"}]),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "date",
                'IF({invoice}{r}="","Add - which invoice?",'
                'IF(OR({amount}{r}="",{amount}{r}=0),'
                '"Add - how much came in?",'
                'IF({balance}{r}<-0.004,"Check - overpaid",'
                'IF({banked}{r}="No","Due - not banked yet",'
                'IF({balance}{r}>0.004,"Part paid","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Payment Tracking",
        subtitle="Every payment received, matched to its invoice.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).payments,
        help_text="One row per payment. Part payments get a row each "
                  "against the same invoice, and the invoice balance "
                  "updates itself.",
        extra_notes=[
            "Record the gross amount and the fee separately. A 2.9% "
            "card fee on $60,000 of turnover is $1,740 you would "
            "otherwise never see.",
            "The 'Banked?' column exists because cash jobs are how "
            "small cleaning businesses lose money without noticing.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/supplies.py`

*158 lines · 1 function*

> `Supplies` - the cupboard, and what it is costing you.
>
> Count the stock once, log what you use on the Supply Usage tab, and
> this tab works out what is left, how fast it is going and when to
> reorder. The cost per use is what makes a job costing honest.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Supplies').

````python
"""`Supplies` - the cupboard, and what it is costing you.

Count the stock once, log what you use on the Supply Usage tab, and
this tab works out what is left, how fast it is going and when to
reorder. The cost per use is what makes a job costing honest.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Supplies"

_USE = "[[Supply Usage.supply]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "SupplyPrefix", digits=4, width=10),
        Col("name", "Item", 26, "text",
            note="What you call it on the shopping list. The Supply "
                 "Usage tab picks from these names."),
        Col("category", "Category", 20, "text",
            validate=list_from("SupplyCategoryList")),
        Col("supplier", "Supplier", 18, "text"),
        Col("unit", "Unit", 11, "text",
            validate=list_from("SupplyUnitList"),
            note="The unit you use it in - a bottle, a litre, a "
                 "roll. Buy in packs, use in units."),
        Col("pack_size", "Units in a pack", 10, "number1",
            validate=number_check(0, 100000),
            note="Six bottles in a case, 24 rolls in a pack."),
        Col("pack_cost", "Pack cost", 11, "money",
            validate=number_check(0, 100000)),
        Col("stock", "Counted stock", 11, "number1",
            validate=number_check(0, 1000000),
            note="How many units you had when you last counted. "
                 "Count the cupboard once a month and type it here."),
        Col("count_date", "Counted on", 11, "date", validate=date_check(),
            note="Usage logged after this date is taken off the "
                 "counted stock."),
        Col("reorder", "Reorder at", 10, "number1",
            validate=number_check(0, 1000000),
            note="The level that triggers a 'low stock' flag."),
        Col("order_qty", "Order quantity", 10, "number1",
            validate=number_check(0, 1000000)),
        Col("last_order", "Last ordered", 11, "date",
            validate=date_check()),
        Col("eco", "Eco / low tox?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Worth tracking - more and more clients ask."),
        Col("coshh", "Hazard sheet?", 10, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Any chemical your staff use should have its safety "
                 "data sheet on file. The Safety tab tracks them."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("unit_cost", "Cost a unit", 10, "money",
            formula='=IF(OR({name}{r}="",{pack_cost}{r}="",'
                    '{pack_size}{r}="",{pack_size}{r}=0),"",'
                    'ROUND({pack_cost}{r}/{pack_size}{r},4))',
            note="What one use of it costs. This is the figure the "
                 "Supply Usage tab multiplies."),
        Col("used_total", "Used (all time)", 11, "number1",
            formula='=IF({name}{r}="","",SUMIFS('
                    '[[Supply Usage.qty]],' + _USE + ',{name}{r}))'),
        Col("used_since", "Used since the count", 11, "number1",
            formula='=IF(OR({name}{r}="",{count_date}{r}=""),"",'
                    'SUMIFS([[Supply Usage.qty]],' + _USE + ',{name}{r},'
                    '[[Supply Usage.date]],">"&{count_date}{r}))',
            hidden=True),
        Col("on_hand", "On hand", 10, "number1",
            formula=blank(
                "name",
                'IF({stock}{r}="","",{stock}{r}-IF({used_since}{r}="",0,'
                '{used_since}{r}))'),
            cf=[{"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "bad"}],
            note="Counted stock less everything logged since the "
                 "count."),
        Col("used_28", "Used in 28 days", 10, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Supply Usage.qty]],' +
                    _USE + ',{name}{r},[[Supply Usage.date]],'
                    '">="&TodayDate-28))',
            hidden=True),
        Col("daily_use", "Used a day", 9, "number1",
            formula='=IF({name}{r}="","",ROUND({used_28}{r}/28,3))',
            hidden=True),
        Col("days_cover", "Days of cover", 10, "int",
            formula='=IF(OR({name}{r}="",{on_hand}{r}="",'
                    '{daily_use}{r}=0),"",ROUND({on_hand}{r}/'
                    '{daily_use}{r},0))',
            cf=[{"type": "cell", "criteria": "<", "value": "=LowStockDays",
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "bad"}],
            note="At the rate of the last four weeks. Blank means "
                 "you have not used any recently."),
        Col("stock_value", "Stock value", 11, "money",
            formula='=IF(OR({name}{r}="",{on_hand}{r}="",'
                    '{unit_cost}{r}=""),"",ROUND({on_hand}{r}*'
                    '{unit_cost}{r},2))',
            note="What is sitting in the cupboard. Add the column up "
                 "and most owners get a surprise."),
        Col("spend", "Spent on it", 11, "money",
            formula='=IF({name}{r}="","",SUMIFS('
                    '[[Supply Usage.cost]],' + _USE + ',{name}{r}))'),
        Col("status", "Stock", 13, "status",
            formula=blank(
                "name",
                'IF({on_hand}{r}="","Not counted",'
                'IF({on_hand}{r}<=0,"Out of stock",'
                'IF(AND({reorder}{r}<>"",{on_hand}{r}<={reorder}{r}),'
                '"Low stock",IF(AND({days_cover}{r}<>"",'
                '{days_cover}{r}<LowStockDays),"Low stock",'
                '"In stock"))))'),
            cf=status_cf()),
        Col("alert", "Action", 26, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Out of stock","Order - none left",'
                'IF({status}{r}="Low stock","Order - running low",'
                'IF({status}{r}="Not counted","Add - count the stock",'
                'IF(AND({coshh}{r}="No",{category}{r}='
                '"Chemicals & detergents"),'
                '"Action - no safety data sheet","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Cleaning Supplies & Inventory",
        subtitle="What you hold, what you use and what it is really "
                 "costing per job.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).supplies,
        help_text="Count the cupboard once and put the date against "
                  "it. From then on, log what you use on the Supply "
                  "Usage tab and the stock level follows by itself.",
        extra_notes=[
            "Cost per unit is the number that matters. A $14 bottle "
            "of degreaser that does eight ovens costs $1.75 a job - "
            "that belongs in your price.",
            "Days of cover beats a reorder level on its own, "
            "because it notices when a new contract doubles your "
            "usage.",
            "Any chemical your staff use should have a safety data "
            "sheet. Keep them with the file - the Safety tab tracks "
            "which ones you hold.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/supply_usage.py`

*85 lines · 1 function*

> `Supply Usage` - what got used, where.
>
> A short tab with a long reach: it takes stock off the cupboard, puts a
> real materials cost onto the job costing, and shows you which client
> quietly gets through twice as much as everybody else.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Supply Usage').

````python
"""`Supply Usage` - what got used, where.

A short tab with a long reach: it takes stock off the cupboard, puts a
real materials cost onto the job costing, and shows you which client
quietly gets through twice as much as everybody else.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, month_col

NAME = "Supply Usage"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("supply", '"SU-"', digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("supply", "Item", 24, "text", validate=list_from("SupplyNames"),
            note="From the Supplies tab. The cost comes with it."),
        Col("qty", "Quantity used", 11, "number1",
            validate=number_check(0, 100000),
            note="In the unit on the Supplies tab - bottles, rolls, "
                 "litres."),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Link it to a job and the cost lands on that job's "
                 "profit. Leave blank for stock used round the "
                 "office."),
        Col("employee", "Used by", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("reason", "Why", 20, "text",
            note="'Weekly visit', 'restocked the van', 'spillage'."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("client_name", "Client", 20, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))'),
        Col("unit_cost", "Cost a unit", 10, "money",
            formula='=IF({supply}{r}="","",IFERROR(INDEX('
                    '[[Supplies.unit_cost]],MATCH({supply}{r},'
                    '[[Supplies.name]],0)),""))',
            hidden=True),
        Col("cost", "Cost", 11, "money",
            formula=blank(
                "supply",
                'ROUND(IF({qty}{r}="",0,{qty}{r})*IFERROR('
                '{unit_cost}{r}*1,0),2)'),
            note="Quantity times the cost per unit. This is what "
                 "lands on the job costing.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("alert", "Check", 22, "status",
            formula=blank(
                "supply",
                'IF({qty}{r}="","Add - how much?",'
                'IF({unit_cost}{r}="","Check - item not in Supplies",'
                'IF({job}{r}="","Not on a job","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Supply Usage",
        subtitle="What you used, on which job, and what it cost.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_soft,
        freeze_cols=2,
        samples=sample(cfg).supply_usage,
        help_text="Log the big items - the chemicals, the pads, the "
                  "bags. Nobody needs to log a single cloth, and a "
                  "tab nobody fills in is worse than no tab at all.",
        extra_notes=[
            "Every row here comes off the stock on the Supplies tab "
            "and goes onto the cost of the job you linked it to.",
            "If you do not log usage, the job costing falls back to "
            "the per-hour supplies allowance on the Settings tab - "
            "so your numbers still work, they are just rougher.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/equipment.py`

*138 lines · 1 function*

> `Equipment` - vacuums, machines, vans and ladders.
>
> What you own, what it cost, who has it and when it was last serviced.
> A vacuum that dies on a Monday morning costs you a day's work, so the
> next service date on this tab is worth more than it looks.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Equipment').

````python
"""`Equipment` - vacuums, machines, vans and ladders.

What you own, what it cost, who has it and when it was last serviced.
A vacuum that dies on a Monday morning costs you a day's work, so the
next service date on this tab is worth more than it looks.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Equipment"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "EquipmentPrefix", digits=3, width=10),
        Col("name", "Equipment", 26, "text",
            note="Something you would recognise on a van check: "
                 "'Numatic Henry #2', not 'vacuum'."),
        Col("type", "Type", 18, "text",
            validate=list_from("EquipmentTypeList")),
        Col("make", "Make & model", 20, "text"),
        Col("serial", "Serial / registration", 16, "text",
            note="Worth recording. Stolen equipment without a serial "
                 "number is gone."),
        Col("ownership", "Ownership", 14, "text",
            validate=list_from("OwnershipList")),
        Col("status", "Status", 14, "status",
            validate=list_from("EquipmentStatusList"), cf=status_cf()),
        Col("assigned", "Assigned to", 16, "text",
            validate=list_from("EmployeeNames"),
            note="Who has it. The commonest cause of a 'missing' "
                 "machine is that somebody took it home."),
        Col("location", "Kept at", 16, "text",
            validate=list_from("VehicleList")),
        Col("bought", "Bought on", 11, "date", validate=date_check()),
        Col("cost", "Cost", 11, "money", validate=number_check(0, 1000000)),
        Col("life_years", "Expected life (years)", 10, "number1",
            validate=number_check(0, 50),
            note="Used to work out what it costs you a year, and "
                 "when to start saving for its replacement."),
        Col("warranty", "Warranty to", 11, "date", validate=date_check()),
        Col("last_service", "Last serviced", 11, "date",
            validate=date_check()),
        Col("interval_input", "Service every (days)", 10, "int",
            validate=number_check(0, 3650, decimal=False),
            note="Blank uses the standard interval from the Settings "
                 "tab."),
        Col("service_cost", "Spent on repairs", 11, "money",
            validate=number_check(0, 1000000),
            note="Running total of what you have spent keeping it "
                 "going. When it passes the replacement cost, stop."),
        Col("safety_test", "Safety tested on", 11, "date",
            validate=date_check(),
            note="Electrical safety testing, where your country "
                 "requires it."),
        Col("insured", "Insured?", 9, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf()),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("age_years", "Age (years)", 10, "number1",
            formula='=IF(OR({name}{r}="",{bought}{r}=""),"",'
                    'ROUND((TodayDate-{bought}{r})/365.25,1))'),
        Col("annual_cost", "Costs a year", 11, "money",
            formula='=IF(OR({name}{r}="",{cost}{r}="",'
                    '{life_years}{r}="",{life_years}{r}=0),"",'
                    'ROUND({cost}{r}/{life_years}{r}+IF('
                    '{service_cost}{r}="",0,{service_cost}{r})/'
                    'MAX(1,{age_years}{r}),2))',
            note="Purchase spread over its life, plus what you have "
                 "spent repairing it. The real cost of owning it."),
        Col("interval", "Interval used", 9, "int",
            formula=blank("name",
                          'IF({interval_input}{r}="",MaintenanceDays,'
                          '{interval_input}{r})'),
            hidden=True),
        Col("next_service", "Next service due", 12, "date",
            formula='=IF(OR({name}{r}="",{last_service}{r}=""),"",'
                    '{last_service}{r}+{interval}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=TodayDate",
                 "fmt": "bad"}]),
        Col("days_to_service", "Days to service", 10, "int",
            formula='=IF({next_service}{r}="","",'
                    '{next_service}{r}-TodayDate)',
            hidden=True),
        Col("replace_year", "Replace by", 10, "int",
            formula='=IF(OR({name}{r}="",{bought}{r}="",'
                    '{life_years}{r}=""),"",YEAR({bought}{r})+'
                    '{life_years}{r})',
            hidden=True),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Retired","Archived",'
                'IF({status}{r}="Lost or stolen","Action - report it",'
                'IF({status}{r}="In repair","Check - in for repair",'
                'IF(AND({days_to_service}{r}<>"",'
                '{days_to_service}{r}<0),"Overdue - service it",'
                'IF(AND({days_to_service}{r}<>"",'
                '{days_to_service}{r}<=14),"Due - service coming up",'
                'IF(AND({last_service}{r}="",{bought}{r}<>""),'
                '"Add - never serviced",'
                'IF(AND({age_years}{r}<>"",{life_years}{r}<>"",'
                '{age_years}{r}>{life_years}{r}),'
                '"Check - past its expected life","OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Equipment & Assets",
        subtitle="What you own, what it costs to keep and when it "
                 "next needs looking at.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).equipment,
        help_text="Add every machine worth more than a week's "
                  "takings, plus the vans. Record the service date "
                  "each time and the next one works itself out.",
        extra_notes=[
            "'Costs a year' spreads the purchase over the expected "
            "life and adds the repairs. It is the number to compare "
            "with the price of a new one.",
            "A machine that has cost more in repairs than it did to "
            "buy is telling you something. Replace it before it "
            "breaks down in a client's hallway.",
            "Electrical safety testing is a legal requirement in "
            "some countries and not in others. Check what applies "
            "where you are - the column is here either way.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/expenses.py`

*116 lines · 1 function*

> `Expenses` - everything the business spends.
>
> Wages, products, fuel, insurance, advertising, software. One row per
> payment out. The dashboard's profit figure is this tab subtracted from
> the invoices, so what is missing here is money you think you have.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Expenses').

````python
"""`Expenses` - everything the business spends.

Wages, products, fuel, insurance, advertising, software. One row per
payment out. The dashboard's profit figure is this tab subtracted from
the invoices, so what is missing here is money you think you have.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, month_col

NAME = "Expenses"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "ExpensePrefix", digits=4, width=11),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("category", "Category", 22, "text",
            validate=list_from("ExpenseCategoryList"),
            note="Keep to the list. Consistent categories are what "
                 "make the yearly comparison worth reading."),
        Col("type", "Type", 13, "text", validate=list_from("ExpenseTypeList"),
            note="Fixed costs happen whether you work or not. "
                 "Variable costs follow the work. Knowing which is "
                 "which is most of financial planning."),
        Col("description", "What it was", 32, "text"),
        Col("supplier", "Paid to", 20, "text"),
        Col("amount", "Amount (net)", 12, "money",
            validate=number_check(0, 10000000),
            note="Before tax, if you can reclaim it. Otherwise just "
                 "put the amount you paid."),
        Col("tax", "Tax", 10, "money", validate=number_check(0, 1000000)),
        Col("method", "Paid by", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Only for a cost that belongs to one job - a skip, "
                 "a specialist product, a subcontractor."),
        Col("employee", "Who spent it", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("receipt", "Receipt?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="No receipt usually means no deduction. Photograph "
                 "them in the van and put the link in the notes."),
        Col("recurring", "Recurring?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("deductible", "Claimable?", 9, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Your accountant decides this, not the "
                 "spreadsheet. The column is here so you can mark "
                 "what you have asked about."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("total", "Total paid", 12, "money",
            formula=blank("date",
                          'IF({amount}{r}="",0,{amount}{r})+'
                          'IF({tax}{r}="",0,{tax}{r})'),
            cf=[{"type": "data_bar", "bar_color": "#C86B6B"}]),
        Col("month_total", "This category, this month", 12, "money",
            formula='=IF(OR({date}{r}="",{category}{r}=""),"",'
                    'SUMIFS({total_first}:{total_last},'
                    '{category_first}:{category_last},{category}{r},'
                    '{month_first}:{month_last},{month}{r}))',
            hidden=True),
        Col("share", "Share of the month", 10, "percent0",
            formula='=IF(OR({date}{r}="",{month}{r}=""),"",'
                    'IFERROR({month_total}{r}/SUMIFS('
                    '{total_first}:{total_last},'
                    '{month_first}:{month_last},{month}{r}),""))',
            hidden=True),
        Col("client_name", "On behalf of", 18, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))',
            hidden=True),
        Col("alert", "Check", 24, "status",
            formula=blank(
                "date",
                'IF({amount}{r}="","Add - how much?",'
                'IF({category}{r}="","Add - which category?",'
                'IF({receipt}{r}="No","Action - no receipt",'
                'IF(AND({amount}{r}>500,{description}{r}=""),'
                '"Add - describe a large payment","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Business Expenses",
        subtitle="Everything that goes out, grouped the way your "
                 "accountant wants to see it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).expenses,
        help_text="One row per payment. Use the same categories every "
                  "time - that is what makes the monthly comparison "
                  "on the Reports tab worth reading.",
        extra_notes=[
            "Wages belong here too, even if you also keep "
            "timesheets. The dashboard's profit is invoices less "
            "this tab, so leaving wages out flatters you badly.",
            "Mileage has its own tab because it is usually claimed "
            "at a rate per mile rather than as a receipt. Do not "
            "enter it twice.",
            "This is a record of what you spent, not tax advice. "
            "What is and is not claimable depends on where you are "
            "and what you do - ask an accountant.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/mileage.py`

*108 lines · 1 function*

> `Mileage & Travel` - the miles between the jobs.
>
> Cleaning is a driving business. Logged properly, travel is both a real
> cost to put against a job and, in most countries, a deduction. Logged
> nowhere, it is the quietest hole in a cleaning company's margin.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Mileage').

````python
"""`Mileage & Travel` - the miles between the jobs.

Cleaning is a driving business. Logged properly, travel is both a real
cost to put against a job and, in most countries, a deduction. Logged
nowhere, it is the quietest hole in a cleaning company's margin.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, month_col

NAME = "Mileage"


def spec(cfg: Config) -> SheetSpec:
    unit = cfg.distance_unit
    columns = [
        id_col("date", '"MIL-"', digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("employee", "Driver", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("vehicle", "Vehicle", 13, "text",
            validate=list_from("VehicleList")),
        Col("purpose", "Purpose", 18, "text",
            validate=list_from("MileagePurposeList")),
        Col("from_place", "From", 22, "text"),
        Col("to_place", "To", 22, "text"),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Link it to a job and the travel cost lands on that "
                 "job's profit."),
        Col("distance", f"Distance ({unit})", 11, "number1",
            validate=number_check(0, 10000)),
        Col("return_trip", "Return trip?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Yes doubles the distance. Most job trips are "
                 "returns."),
        Col("rate_input", "Rate (override)", 10, "money",
            validate=number_check(0, 100),
            note=f"Blank uses the rate per {unit[:-1]} from the "
                 f"Settings tab."),
        Col("parking", "Parking & tolls", 10, "money",
            validate=number_check(0, 10000)),
        Col("claimed", "Claimed?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("client_name", "Client", 18, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))',
            hidden=True),
        Col("total_distance", f"Total {unit}", 10, "number1",
            formula=blank("date",
                          'IF({distance}{r}="",0,{distance}{r})*'
                          'IF({return_trip}{r}="Yes",2,1)'),
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("rate_used", "Rate", 9, "money",
            formula=blank("date",
                          'IF({rate_input}{r}<>"",{rate_input}{r},'
                          'MileageRate)'),
            hidden=True),
        Col("amount", "Travel cost", 11, "money",
            formula=blank("date",
                          'ROUND({total_distance}{r}*{rate_used}{r}+'
                          'IF({parking}{r}="",0,{parking}{r}),2)'),
            note="Distance times the rate, plus parking. This is "
                 "what you claim and what the job carries."),
        Col("alert", "Check", 22, "status",
            formula=blank(
                "date",
                'IF({distance}{r}="","Add - how far?",'
                'IF({employee}{r}="","Add - who drove?",'
                'IF({claimed}{r}="No","Due - not claimed yet",'
                'IF({total_distance}{r}>200,"Check - a long day",'
                '"OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Mileage & Travel",
        subtitle="Every trip, what it cost and which job should "
                 "carry it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).mileage,
        help_text="Log the trip on the day - a week later nobody "
                  "remembers. Link it to a job and the travel cost "
                  "appears in that job's profit.",
        extra_notes=[
            "Travel between two jobs is not billable time, but it "
            "is paid time. Twenty minutes between every visit is "
            "most of an hour a day you are paying for and nobody is "
            "buying.",
            "The rate per mile is on the Settings tab. Check what "
            "your tax authority allows - it changes, and it is "
            "rarely what you guess.",
            "A mileage log is only evidence if it is contemporary. "
            "From, to, purpose and distance, written the same week.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/complaints.py`

*127 lines · 1 function*

> `Complaints` - what went wrong and what you did about it.
>
> Nobody enjoys this tab. It is also the one that keeps clients: a
> complaint handled in two days is a loyal customer, and a complaint
> nobody wrote down happens again next month.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Complaints').

````python
"""`Complaints` - what went wrong and what you did about it.

Nobody enjoys this tab. It is also the one that keeps clients: a
complaint handled in two days is a loyal customer, and a complaint
nobody wrote down happens again next month.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_client,
                     month_col, score_validation, status_cf)

NAME = "Complaints"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "ComplaintPrefix", digits=4, width=11),
        Col("date", "Date raised", 11, "date", validate=date_check()),
        month_col("date"),
        Col("client", "Client ID", 11, "id", validate=list_from("ClientIDs")),
        Col("client_name", "Client", 20, "text",
            formula=lookup_client("name")),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Which visit it relates to, if you know."),
        Col("employee", "Cleaner on the job", 16, "text",
            validate=list_from("EmployeeNames"),
            note="Record it, but treat a single complaint as "
                 "information rather than evidence."),
        Col("type", "What about", 20, "text",
            validate=list_from("ComplaintTypeList")),
        Col("severity", "Severity", 11, "status",
            validate=list_from("SeverityList"), cf=status_cf()),
        Col("description", "What they said", 38, "wrap",
            note="In their words where you can. It reads very "
                 "differently from yours."),
        Col("status", "Status", 15, "status",
            validate=list_from("ComplaintStatusList"), cf=status_cf()),
        Col("owner", "Being handled by", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("action", "What we did", 36, "wrap"),
        Col("resolution", "Resolution", 18, "text",
            validate=list_from("ResolutionList")),
        Col("reclean_job", "Re-clean job", 11, "id",
            validate=list_from("JobIDs"),
            note="If you went back, book it as a job with type "
                 "'Re-clean (free)' and link it here."),
        Col("cost", "Cost to put right", 11, "money",
            validate=number_check(0, 1000000),
            note="Refund, free re-clean, replacement. Worth knowing "
                 "- complaints are a real line in your costs."),
        Col("resolved_date", "Resolved on", 11, "date",
            validate=date_check()),
        Col("kept", "Client kept?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("rating_after", "Rating afterwards", 10, "rating",
            validate=score_validation(5)),
        Col("prevent", "How we stop it happening again", 34, "wrap",
            note="The most valuable column on the tab, and the one "
                 "everybody leaves blank."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("days_open", "Days open", 10, "int",
            formula='=IF(OR({date}{r}="",{client}{r}=""),"",'
                    'IF({resolved_date}{r}<>"",'
                    '{resolved_date}{r}-{date}{r},TodayDate-{date}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 7,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 2,
                 "fmt": "warn"}]),
        Col("open", "Still open?", 9, "flag",
            formula=blank("date",
                          'IF(OR({status}{r}="Resolved",'
                          '{status}{r}="Closed - no action",'
                          '{status}{r}="Refunded"),"No","Yes")'),
            cf=flag_yes_no_cf(), hidden=True),
        Col("repeat", "Repeat issue?", 10, "flag",
            formula=blank(
                "date",
                'IF(COUNTIFS({client_first}:{client_last},{client}{r},'
                '{type_first}:{type_last},{type}{r})>1,"Yes","No")'),
            cf=flag_yes_no_cf(),
            note="The same client complaining about the same thing "
                 "twice. Fix the cause, not the visit."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({open}{r}="No","OK - closed",'
                'IF({days_open}{r}>7,"Overdue - open "&{days_open}{r}&'
                '" days",'
                'IF({severity}{r}="Critical","Action - critical, today",'
                'IF({owner}{r}="","Assign - nobody owns this",'
                'IF({repeat}{r}="Yes","Check - repeat issue",'
                '"Action - open")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Complaints & Quality Control",
        subtitle="What went wrong, how fast you fixed it and whether "
                 "you kept the client.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).complaints,
        help_text="Log every complaint, even the small ones you "
                  "sorted out on the phone. The pattern is the point, "
                  "not the individual row.",
        extra_notes=[
            "Days open is the number that matters. Research after "
            "research says the same thing: it is not the mistake "
            "that loses the client, it is the silence afterwards.",
            "A free re-clean has a real cost - the wages, the "
            "products and the hour you could have sold to somebody "
            "else. Put it in the cost column and the picture gets "
            "clearer.",
            "Treat the cleaner column carefully. One complaint is "
            "noise; three in a month against the same person is a "
            "training conversation.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/incidents.py`

*96 lines · 1 function*

> `Incidents` - accidents, near misses and damage.
>
> Short, dull and important. Cleaning involves wet floors, chemicals,
> ladders and other people's property, so things occasionally go wrong.
> A written record protects your staff, your client and you.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Incidents').

````python
"""`Incidents` - accidents, near misses and damage.

Short, dull and important. Cleaning involves wet floors, chemicals,
ladders and other people's property, so things occasionally go wrong.
A written record protects your staff, your client and you.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Incidents"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"INC-"', digits=3, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("time", "Time", 9, "text"),
        Col("type", "What happened", 22, "text",
            validate=list_from("IncidentTypeList")),
        Col("employee", "Person involved", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs")),
        Col("property", "Property ID", 11, "id",
            validate=list_from("PropertyIDs")),
        Col("severity", "Severity", 11, "status",
            validate=list_from("SeverityList"), cf=status_cf()),
        Col("description", "What happened, in full", 40, "wrap",
            note="Write it the same day, in plain language, "
                 "including what was being done at the time."),
        Col("injury", "Injury?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("first_aid", "First aid given?", 10, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf()),
        Col("medical", "Medical help needed?", 10, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf()),
        Col("damage", "Property damaged", 26, "wrap"),
        Col("cost", "Cost", 11, "money", validate=number_check(0, 1000000)),
        Col("insurance", "Insurance claim?", 10, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf()),
        Col("reported_to", "Reported to", 18, "text",
            note="Client, insurer, the authority your country "
                 "requires. Write down who and when."),
        Col("reported_date", "Reported on", 11, "date",
            validate=date_check()),
        Col("action", "Action taken", 34, "wrap"),
        Col("prevent", "How we prevent a repeat", 32, "wrap"),
        Col("status", "Status", 13, "status",
            validate=list_from("ComplaintStatusList"), cf=status_cf()),
        Col("closed_date", "Closed on", 11, "date", validate=date_check()),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("days_open", "Days open", 10, "int",
            formula='=IF({date}{r}="","",IF({closed_date}{r}<>"",'
                    '{closed_date}{r}-{date}{r},TodayDate-{date}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 14,
                 "fmt": "bad"}]),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "date",
                'IF({closed_date}{r}<>"","OK - closed",'
                'IF(AND({injury}{r}="Yes",{reported_date}{r}=""),'
                '"Action - injury not reported yet",'
                'IF({severity}{r}="Critical","Action - critical",'
                'IF({prevent}{r}="","Add - what stops it happening '
                'again?","Action - open"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Incidents & Near Misses",
        subtitle="Accidents, damage and the things that almost went "
                 "badly wrong.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.grey_text,
        freeze_cols=2,
        samples=sample(cfg).incidents,
        help_text="Record it the same day, including near misses. "
                  "The near misses are the free lessons.",
        extra_notes=[
            "Reporting duties differ from country to country and "
            "some injuries must be reported to an authority within "
            "days. Find out what applies where you are - this "
            "workbook keeps the record, it does not know your law.",
            "Damage to a client's property is better reported by "
            "you, with a photo, than discovered by them a week "
            "later.",
        ],
    )
````

### `cleaning_business_spreadsheet/sheets/safety.py`

*113 lines · 1 function*

> `Safety & Compliance` - the paperwork that has to be in date.
>
> Insurance certificates, background checks, training, risk assessments,
> chemical data sheets. Each one has an expiry date and a quiet habit of
> lapsing in the week a new commercial client asks to see it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Safety').

````python
"""`Safety & Compliance` - the paperwork that has to be in date.

Insurance certificates, background checks, training, risk assessments,
chemical data sheets. Each one has an expiry date and a quiet habit of
lapsing in the week a new commercial client asks to see it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Safety"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("document", '"DOC-"', digits=3, width=10),
        Col("document", "Document", 26, "text",
            validate=list_from("SafetyDocList"),
            note="Pick the type, or type your own. One row per "
                 "document per person."),
        Col("applies_to", "Applies to", 18, "text",
            note="A person, a property, a chemical, or the business "
                 "as a whole."),
        Col("employee", "Person", 16, "text",
            validate=list_from("EmployeeNames"),
            note="Leave blank for a business-wide document."),
        Col("reference", "Reference / number", 20, "text"),
        Col("issuer", "Issued by", 20, "text"),
        Col("issued", "Issued on", 11, "date", validate=date_check()),
        Col("expires", "Expires", 11, "date", validate=date_check(),
            note="The whole point of the tab. Leave blank for a "
                 "document that does not expire."),
        Col("cost", "Cost to renew", 10, "money",
            validate=number_check(0, 1000000)),
        Col("holder", "Where it is kept", 22, "text",
            note="Folder, drive, filing cabinet. 'Somewhere in the "
                 "van' is not an answer."),
        Col("link", "Link", 24, "link",
            note="A link to the scan beats a photocopy in a drawer."),
        Col("required", "Required?", 9, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Some documents are legally required, some are "
                 "asked for by clients, some are neither."),
        Col("reviewed", "Last reviewed", 11, "date",
            validate=date_check(),
            note="Risk assessments and method statements should be "
                 "looked at again every year, not just filed."),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("days_left", "Days to expiry", 11, "int",
            formula='=IF(OR({document}{r}="",{expires}{r}=""),"",'
                    '{expires}{r}-TodayDate)',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=CertWarningDays", "fmt": "warn"}]),
        Col("status", "Status", 14, "status",
            formula=blank(
                "document",
                'IF({expires}{r}="","Valid",'
                'IF({days_left}{r}<0,"Expired",'
                'IF({days_left}{r}<=CertWarningDays,"Due soon",'
                '"Valid")))'),
            cf=status_cf()),
        Col("renew_by", "Start renewing", 11, "date",
            formula='=IF({expires}{r}="","",'
                    '{expires}{r}-CertWarningDays)',
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "document",
                'IF({status}{r}="Expired","Expired - renew it now",'
                'IF({status}{r}="Due soon","Due - renew within "&'
                '{days_left}{r}&" days",'
                'IF(AND({required}{r}="Yes",{link}{r}="",'
                '{holder}{r}=""),"Add - where is this kept?",'
                'IF(AND({reviewed}{r}<>"",TodayDate-{reviewed}{r}>365),'
                '"Due - review it, over a year old","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Safety & Compliance",
        subtitle="Certificates, checks and assessments, with the "
                 "expiry dates watched for you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.grey_text,
        freeze_cols=2,
        samples=sample(cfg).safety,
        help_text="One row per document. Put the expiry date in and "
                  "the workbook warns you well before a commercial "
                  "client asks to see it.",
        extra_notes=[
            "What you are legally required to hold depends "
            "entirely on where you are and who you employ - "
            "insurance, background checks for work around "
            "children, chemical data sheets, electrical testing. "
            "This tab keeps the record; it does not know your "
            "country's rules.",
            "The warning window is on the Settings tab. Six weeks "
            "is usually enough to renew anything without "
            "panicking.",
            "Risk assessments are not a one-off. The review column "
            "flags anything you have not looked at for a year.",
        ],
    )
````

---

## 5. The pages

The tabs that read the registers: the welcome page, the dashboard, the quote calculator, the month calendar, the staff rota, retention, the reports, the financial plan, the three printables and the help tab.

### `cleaning_business_spreadsheet/sheets/start_here.py`

*294 lines · 1 function*

> `Start Here` - the first thing anyone sees when they open the file.
>
> Twenty minutes of setup, then the workbook runs the admin side of a
> cleaning business for you. This page explains the order to do things
> in, what fills itself in, and - just as important - what a spreadsheet
> honestly cannot do.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Start Here'), `LAST_COL` (9).

````python
"""`Start Here` - the first thing anyone sees when they open the file.

Twenty minutes of setup, then the workbook runs the admin side of a
cleaning business for you. This page explains the order to do things
in, what fills itself in, and - just as important - what a spreadsheet
honestly cannot do.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Start Here"
LAST_COL = 9


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)

    s.banner(cfg.product_name,
             "A complete back office for a cleaning business - "
             "clients, quotes, jobs, staff, invoices and profit.",
             LAST_COL)

    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 18)
    s.set_col(3, LAST_COL, 15)

    row = 4
    row = s.note(row,
                 "Welcome. This workbook replaces the diary, the "
                 "quote pad, the invoice book and the shoebox of "
                 "receipts. Everything is a formula you can read and "
                 "change - no macros, no add-ins, no subscription.",
                 LAST_COL, height=34)
    row += 1

    # -------------------------------------------------- what you get ---
    row = s.section(row, "What is in the file", LAST_COL)
    row = s.table_head(row, 1, ["Tab", "What it does", "Your job"],
                       [30, 52, 26])
    tabs = [
        ("Business Settings", "Your name, rates, tax, targets and "
                              "every reminder window",
         "Fill in once, revisit twice a year"),
        ("Dashboard", "Twelve live numbers, the jobs for the week and "
                      "everything that needs doing",
         "Read it every morning"),
        ("Pricing & Quote Calculator",
         "Prices a clean four different ways and shows the margin on "
         "each", "Use it before you quote"),
        ("Client CRM", "Every client, what they are worth and whether "
                       "they are drifting away", "Add clients"),
        ("Property & Site Database",
         "Addresses, access, keys, pets, parking and the hours each "
         "property needs", "Add properties"),
        ("Leads & Sales Pipeline",
         "Enquiries from first call to won or lost",
         "Log every enquiry"),
        ("Service Catalog", "What you offer, what it costs you and "
                            "what it earns an hour", "Set up once"),
        ("Quotes & Estimates", "Quotes, their lines, follow-ups and "
                               "acceptance", "Raise quotes"),
        ("Job Scheduler", "Every visit: who, where, when, how long, "
                          "how much", "The heart of the workbook"),
        ("Calendar", "The month at a glance and the next thirty "
                     "visits in order", "Read only"),
        ("Cleaning Checklists",
         "Room-by-room task templates, and the tasks ticked off on "
         "the day", "Build your own"),
        ("Employee Database", "Your team, their pay, skills, checks "
                              "and performance", "Add staff"),
        ("Employee Schedule & Attendance",
         "The week's rota, who is overloaded and who has room",
         "Read, then fix on Jobs"),
        ("Job Costing & Profitability",
         "Wages, supplies and travel against the price, job by job",
         "Automatic"),
        ("Invoices & Payments",
         "Invoices, lines, payments, ageing and chasing",
         "Raise and record"),
        ("Supplies & Equipment",
         "Stock, usage, reorder points, machines and their servicing",
         "Count monthly"),
        ("Expenses & Mileage", "Every cost, with tax and categories, "
                               "plus business travel", "Log weekly"),
        ("Client Retention", "Who is slipping away and what they are "
                             "worth", "Act on it"),
        ("Complaints & Quality",
         "Complaints, what they cost and whether you kept the client",
         "Log honestly"),
        ("Safety & Compliance",
         "Certificates, checks and assessments with expiry warnings",
         "Keep in date"),
        ("Financial Plan", "Break-even, targets and what if you grew",
         "Plan with it"),
        ("Reports & Analytics", "Where the money comes from and where "
                                "it goes", "Read monthly"),
        ("Printable Forms", "Quote, invoice and job sheet, ready for "
                            "A4", "Print as needed"),
        ("Settings & Dropdown Lists",
         "Every drop-down list in the workbook", "Edit to suit you"),
    ]
    for tab, what, job in tabs:
        s.write(row, 1, tab, fmts.chrome("label_bold"))
        s.merge(row, 2, 6, what, fmts.chrome("label"))
        s.merge(row, 7, LAST_COL, job, fmts.chrome("note"))
        row += 1
    row += 1

    # ------------------------------------------------- the setup list --
    row = s.section(row, "Twenty minutes to set it up", LAST_COL)
    row = s.table_head(row, 1, ["Step", "Where", "What to do"],
                       [8, 30, 60])
    steps = [
        ("1", "Business Settings",
         "Your business name, address, phone and email. They print on "
         "every quote and invoice."),
        ("2", "Business Settings",
         "Your hourly rate, minimum charge and what a cleaner costs "
         "you an hour. Nothing else works until these are right."),
        ("3", "Business Settings",
         "Tax rate and payment terms. Set the tax rate to 0 if you "
         "are not registered."),
        ("4", "Service Catalog",
         "Edit the services to match what you actually sell. Delete "
         "the ones you do not offer."),
        ("5", "Cleaning Checklists",
         "Adjust the task templates. These become the job sheets your "
         "cleaners carry."),
        ("6", "Employee Database",
         "Add your team - even if the team is just you. Put your own "
         "pay rate in; it is what makes the profit figures real."),
        ("7", "Client CRM and Property & Site Database",
         "Add your clients, then their properties. The property is "
         "where the hours and the access details live."),
        ("8", "Job Scheduler",
         "Put next week's visits in. The price, the finish time and "
         "the invoice link all work themselves out."),
        ("9", "Dashboard",
         "Open it every morning. If a number looks wrong, the tab it "
         "came from is named right beside it."),
    ]
    for num, where, what in steps:
        s.write(row, 1, num, fmts.chrome("kpi_label"))
        s.merge(row, 2, 3, where, fmts.chrome("label_bold"))
        s.merge(row, 4, LAST_COL, what, fmts.chrome("label"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ------------------------------------------------ the daily rhythm -
    row = s.section(row, "The rhythm that keeps it accurate", LAST_COL)
    row = s.table_head(row, 1, ["When", "What", "How long"],
                       [20, 62, 18])
    rhythm = [
        ("Every morning", "Open the Dashboard. Confirm tomorrow's "
                          "visits. Assign anything with no cleaner.",
         "2 minutes"),
        ("After each job", "Mark it Completed and put the actual "
                           "hours in.", "30 seconds"),
        ("Every Friday", "Approve timesheets, raise the week's "
                         "invoices, log receipts.", "20 minutes"),
        ("Every Monday", "Chase unpaid invoices and unanswered "
                         "quotes. This is the hour that pays for "
                         "itself.", "30 minutes"),
        ("Monthly", "Count supplies, check equipment servicing, read "
                    "the Reports tab.", "1 hour"),
        ("Quarterly", "Review prices against the Job Costing tab. Any "
                      "client under your margin floor needs a "
                      "conversation.", "1 hour"),
    ]
    for when, what, how_long in rhythm:
        s.write(row, 1, when, fmts.chrome("label_bold"))
        s.merge(row, 2, 7, what, fmts.chrome("label"))
        s.merge(row, 8, LAST_COL, how_long, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ------------------------------------------------ what it does ----
    row = s.section(row, "What the workbook does by itself", LAST_COL)
    row = s.bullets(row, [
        "Prices a clean from the floor area, the room count, the "
        "service or your target margin - and shows all four.",
        "Works out finish times, crew size and hours per cleaner.",
        "Applies the frequency discount, the pet surcharge, the "
        "minimum charge and your rounding rule.",
        "Flags double bookings, long days and jobs with nobody on "
        "them.",
        "Totals quotes from their lines, dates the expiry and tells "
        "you when to chase.",
        "Builds invoices from the lines, ages them, adds the late fee "
        "and tracks the balance after every payment.",
        "Costs every job with real wages from the timesheets, real "
        "supply usage and the travel, then shows the margin.",
        "Tracks stock down to days of cover and tells you what to "
        "reorder.",
        "Warns you before certificates expire and before machines are "
        "due a service.",
        "Spots clients going quiet long before they tell you they are "
        "leaving.",
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
        "Send reminders by itself. It flags what needs chasing - you "
        "or your phone still send the message.",
        "Sync with Google Calendar, Outlook or a phone diary. Export "
        "the Jobs tab to CSV if you need that.",
        "Take card payments or connect to your bank. Record the "
        "payment here once it lands.",
        "Collect a legally binding electronic signature. Print the "
        "quote, or use a proper e-signature service.",
        "Track your cleaners' location, or clock them in by GPS.",
        "File your tax return. It gets the numbers ready for someone "
        "who can.",
        "Replace advice from an accountant or an employment lawyer, "
        "and it does not know your local rules.",
    ], LAST_COL, first_col=1, height=16)
    row += 1

    # ------------------------------------------------ privacy ---------
    row = s.section(row, "A word about privacy", LAST_COL)
    row = s.note(row,
                 "This file will end up holding client addresses, "
                 "alarm codes, where the keys are kept, staff pay and "
                 "background checks. That is genuinely sensitive. "
                 "Keep it somewhere access-controlled, do not email "
                 "the whole workbook to anyone, give cleaners the "
                 "printed job sheet rather than the file, and delete "
                 "entry details when a client leaves. The sheet "
                 "password protects formulas from accidents - it is "
                 "not encryption.",
                 LAST_COL, first_col=1, height=56)
    row += 1

    # ------------------------------------------------ housekeeping ----
    row = s.section(row, "Housekeeping", LAST_COL)
    row = s.table_head(row, 1, ["Question", "Answer"], [30, 68])
    facts = [
        ("Is there example data in it?",
         "Both versions are in your download: one filled with a "
         "worked example, one empty and ready for your own."),
        ("How do I clear the example?",
         "Select the data rows on a tab and press Delete. The "
         "formula columns refill themselves - they are not deleted."),
        ("Why are some cells locked?",
         f"Every calculated cell is protected so a stray paste "
         f"cannot break it. The password is '{cfg.password}'. "
         f"Review, Unprotect Sheet if you want to change a formula."),
        ("Which cells do I type in?",
         "The white ones. Grey cells calculate themselves."),
        ("Does it work in Google Sheets?",
         "Yes. File, Import, Upload. Drop-downs and conditional "
         "colours come across; sheet protection behaves a little "
         "differently."),
        ("Which Excel versions?",
         "Excel 2010 and newer, Excel for Mac, Excel online, "
         "LibreOffice Calc and Google Sheets. No macros, so nothing "
         "to enable."),
        ("Can I add more rows?",
         "Yes - insert rows inside the existing block and copy the "
         "formulas down. Never insert above the header row."),
        ("Can I add my logo?",
         "Insert, Picture on the printable forms. Everything else is "
         "typed on the Settings tab."),
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

### `cleaning_business_spreadsheet/sheets/dashboard.py`

*375 lines · 1 function*

> `Dashboard` - the one page that answers 'how is the business doing?'
>
> Twelve numbers across the top, a list of things that need doing today,
> the next seven days of work, your biggest clients and twelve months of
> revenue, cost and profit. Everything reads straight from the registers,
> so it is never out of date.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dashboard'), `LAST_COL` (11), `MONTHS` (12).

````python
"""`Dashboard` - the one page that answers 'how is the business doing?'

Twelve numbers across the top, a list of things that need doing today,
the next seven days of work, your biggest clients and twelve months of
revenue, cost and profit. Everything reads straight from the registers,
so it is never out of date.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 11
MONTHS = 12

#: revenue counted the same way everywhere: completed or invoiced work
_DONE = ('SUMIFS([[Jobs.price]],[[Jobs.month]],{key},[[Jobs.status]],'
         '"Completed")+SUMIFS([[Jobs.price]],[[Jobs.month]],{key},'
         '[[Jobs.status]],"Invoiced")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Business Dashboard",
             "Today, this week and the year so far - all worked out "
             "from the jobs, invoices and expenses you have already "
             "entered.", LAST_COL)
    s.set_col(0, 0, 22)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ------------------------------------------------------- the header --
    row = 4
    s.merge(row, 0, 0, "Business", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    ws.write_formula(row, 1, "=BizName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    ws.write_formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")',
                     fmts.calc("text"))
    book.define_name("DashMonth", f"={quote_sheet(NAME)}!$F${row + 1}")
    s.merge(row, 6, 6, "Month key", fmts.chrome("label_right"))
    ws.write_formula(row, 7, '=TEXT(TodayDate,"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$H${row + 1}")
    s.merge(row, 8, 9, "Today", fmts.chrome("label_right"))
    s.merge(row, 10, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 10, "=TodayDate", fmts.calc("date"))
    row += 2

    # ---------------------------------------------------------- the KPIs --
    money = cfg.currency_symbol
    tiles_1 = [
        ("Revenue this month",
         '=' + _DONE.format(key="DashKey"), "money0",
         f"target {money}{cfg.revenue_target_month:,.0f}"),
        ("Invoiced this month",
         '=SUMIFS([[Invoices.total]],[[Invoices.month]],DashKey)',
         "money0", "what you have billed"),
        ("Collected this month",
         '=SUMIFS([[Payments.amount]],[[Payments.month]],DashKey)',
         "money0", "money actually in"),
        ("Owed to you right now",
         '=SUM([[Invoices.outstanding]])', "money0",
         "every unpaid invoice"),
    ]
    tiles_2 = [
        ("Costs this month",
         '=SUMIFS([[Expenses.total]],[[Expenses.month]],DashKey)',
         "money0", f"budget {money}{cfg.expense_budget_month:,.0f}"),
        ("Profit this month",
         '=' + _DONE.format(key="DashKey") +
         '-SUMIFS([[Expenses.total]],[[Expenses.month]],DashKey)',
         "money0", "revenue less expenses"),
        ("Margin this month",
         '=IFERROR((' + _DONE.format(key="DashKey") +
         '-SUMIFS([[Expenses.total]],[[Expenses.month]],DashKey))/(' +
         _DONE.format(key="DashKey") + '),"")', "percent0",
         "after wages and overheads"),
        ("Average job value",
         # the sum is two SUMIFS added together, so it has to be
         # bracketed before it is divided - a+b/c is not (a+b)/c
         '=IFERROR((' + _DONE.format(key="DashKey") +
         ')/(COUNTIFS([[Jobs.month]],DashKey,[[Jobs.status]],'
         '"Completed")+COUNTIFS([[Jobs.month]],DashKey,'
         '[[Jobs.status]],"Invoiced")),"")', "money",
         "completed work this month"),
    ]
    tiles_3 = [
        ("Jobs this week",
         '=COUNTIFS([[Jobs.date]],">="&WeekStartDate,[[Jobs.date]],'
         '"<="&WeekStartDate+6,[[Jobs.status]],"<>Cancelled")', "int",
         "booked in the next seven days"),
        ("Hours booked this week",
         '=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],">="&WeekStartDate,'
         '[[Jobs.date]],"<="&WeekStartDate+6,[[Jobs.status]],'
         '"<>Cancelled")', "number1", "across the whole team"),
        ("Quotes waiting",
         '=COUNTIFS([[Quotes.status]],"Sent")+'
         'COUNTIFS([[Quotes.status]],"Followed up")', "int",
         "worth chasing today"),
        ("Active clients",
         '=COUNTIFS([[Clients.status]],"Active")+'
         'COUNTIFS([[Clients.status]],"Regular")', "int",
         "on the books right now"),
    ]
    for tiles in (tiles_1, tiles_2, tiles_3):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, s.reg.render(formula), style=style,
                  span=3, note=note)
        row += 4

    # ----------------------------------------------- needs your attention
    row = s.section(row, "Needs your attention", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [26, 11, 46, 20])
    alerts = [
        ("Jobs with no cleaner",
         '=COUNTIFS([[Jobs.employee]],"",[[Jobs.client]],"<>",'
         '[[Jobs.date]],">="&TodayDate)',
         '="Booked work nobody is assigned to yet"', "Jobs"),
        ("Visits to confirm",
         '=COUNTIF([[Jobs.alert]],"Confirm*")',
         '="Coming up within "&ReminderDays&" days and not confirmed '
         'with the client"', "Jobs"),
        ("Double bookings",
         '=COUNTIF([[Jobs.clash]],"Clash")',
         '="The same cleaner booked twice at the same time"', "Jobs"),
        ("Jobs done but not invoiced",
         '=COUNTIFS([[Jobs.status]],"Completed",[[Jobs.invoice]],"")',
         '="Work you have done and not yet billed - worth "&'
         'TEXT(SUMIFS([[Jobs.price]],[[Jobs.status]],"Completed",'
         '[[Jobs.invoice]],""),"#,##0")', "Invoices"),
        ("Overdue invoices",
         '=COUNTIF([[Invoices.alert]],"Overdue*")',
         '="Past their due date, totalling "&'
         'TEXT(SUMIFS([[Invoices.outstanding]],[[Invoices.age]],'
         '"1-30 days")+SUMIFS([[Invoices.outstanding]],'
         '[[Invoices.age]],"31-60 days")+SUMIFS('
         '[[Invoices.outstanding]],[[Invoices.age]],"61-90 days")+'
         'SUMIFS([[Invoices.outstanding]],[[Invoices.age]],'
         '"90+ days"),"#,##0")', "Invoices"),
        ("Quotes to chase",
         '=COUNTIF([[Quotes.alert]],"Chase*")',
         '="Sent more than "&QuoteFollowupDays&" days ago with no '
         'answer - this is the fastest money in the workbook"',
         "Quotes"),
        ("Leads with no quote",
         '=COUNTIF([[Leads.alert]],"Action - no quote*")',
         '="Enquiries that have not been priced yet"', "Leads"),
        ("Clients going quiet",
         '=COUNTIF([[Clients.retention]],"At risk")+'
         'COUNTIF([[Clients.retention]],"Slipping")',
         '="No visit booked, or nothing for "&InactiveClientDays&'
         '" days"', "Client Retention"),
        ("Supplies to reorder",
         '=COUNTIF([[Supplies.status]],"Low stock")+'
         'COUNTIF([[Supplies.status]],"Out of stock")',
         '="Below the reorder level or less than "&LowStockDays&'
         '" days of cover"', "Supplies"),
        ("Equipment due a service",
         '=COUNTIF([[Equipment.alert]],"Overdue*")+'
         'COUNTIF([[Equipment.alert]],"Due*")',
         '="A machine that dies mid-job costs you the day"',
         "Equipment"),
        ("Certificates expiring",
         '=COUNTIF([[Safety.status]],"Due soon")+'
         'COUNTIF([[Safety.status]],"Expired")',
         '="Insurance, checks and training within "&CertWarningDays&'
         '" days of running out"', "Safety"),
        ("Open complaints",
         '=COUNTIF([[Complaints.open]],"Yes")',
         '="Every day one of these stays open costs you a client"',
         "Complaints"),
        ("Loss-making jobs",
         '=COUNTIF([[Job Costing.alert]],"Loss*")',
         '="Jobs that cost more than they earned"', "Job Costing"),
        ("Timesheets to approve",
         '=COUNTIF([[Timesheets.alert]],"Due - needs approving")',
         '="Hours entered but not signed off"', "Timesheets"),
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

    # ------------------------------------------------------ the next week
    row = s.section(row, "The next seven days", LAST_COL)
    s.write(row, 0, "Day", fmts.chrome("header_left"))
    for i in range(7):
        ws.write_formula(row, i + 1, f"=TodayDate+{i}",
                         fmts.chrome("header", num_format="ddd d mmm"))
    s.write(row, 8, "Week", fmts.chrome("header"))
    day_row = row
    row += 1
    week = ('[[Jobs.date]],">="&TodayDate,[[Jobs.date]],'
            '"<="&TodayDate+6,[[Jobs.status]],"<>Cancelled")')
    for label, fn, value_range, style in [
            ("Jobs booked", "COUNTIFS", "", "int"),
            ("Hours on site", "SUMIFS", "[[Jobs.hours_each]],",
             "number1"),
            ("Value", "SUMIFS", "[[Jobs.price]],", "money0"),
            ("Cleaners needed", "SUMIFS", "[[Jobs.crew]],", "int")]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        for i in range(7):
            cell = f"{chr(66 + i)}${day_row + 1}"
            day = (f'={fn}({value_range}[[Jobs.date]],{cell},'
                   f'[[Jobs.status]],"<>Cancelled")')
            ws.write_formula(row, i + 1, s.reg.render(day),
                             fmts.calc(style))
        ws.write_formula(row, 8,
                         s.reg.render(f"={fn}({value_range}{week}"),
                         fmts.chrome("total_value",
                                     num_format=s.num_format(style)))
        row += 1
    ws.conditional_format(day_row + 1, 1, day_row + 1, 7,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    row += 1
    s.merge(row, 0, LAST_COL,
            "Nothing in a column means a free day - the best day to "
            "fit in a quote visit, a deep clean or some time off.",
            fmts.chrome("note"))
    row += 2

    # ---------------------------------------------------- biggest clients
    row = s.section(row, "Your five biggest clients", LAST_COL)
    row = s.table_head(row, 0, ["Client", "Type", "Jobs done",
                                "Invoiced", "Profit", "Margin",
                                "Worth a year", "Last clean",
                                "Standing"],
                       [24, 14, 10, 13, 12, 10, 13, 12, 13])
    top = max(3, cfg.top_rows)
    first_client = row
    for n in range(1, top + 1):
        pick = ('MATCH({n},[[Clients.invoiced_rank]],0)'
                .replace("{n}", str(n)))
        s.write(row, 0, None, fmts.calc("text"))
        for col, (key, style) in enumerate([
                ("name", "text"), ("type", "text"),
                ("jobs_done", "int"), ("invoiced", "money"),
                ("profit", "money"), ("margin", "percent0"),
                ("year_value", "money"), ("last_job", "date"),
                ("retention", "status")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Clients.{key}]],{pick}),'
                             f'"")'),
                fmts.calc(style))
        row += 1
    ws.conditional_format(first_client, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    ws.conditional_format(first_client, 4, row - 1, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row += 1

    # ------------------------------------------------------ the year table
    row = s.section(row, "The year so far", LAST_COL)
    row = s.table_head(row, 0, ["Month", "Jobs", "Revenue", "Invoiced",
                                "Collected", "Costs", "Profit", "Margin",
                                "New clients", "Trend"],
                       [10, 9, 13, 13, 13, 12, 12, 9, 11, 14])
    first_month = row
    for m in range(MONTHS):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        ws.write_formula(row, 0, f'=TEXT(DATE(PlanYear,{m + 1},1),"mmm")',
                         fmts.calc("text"))
        revenue = _DONE.format(key=key)
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Jobs.month]],{key},[[Jobs.status]],'
            f'"<>Cancelled")',
            f'={revenue}',
            f'=SUMIFS([[Invoices.total]],[[Invoices.month]],{key})',
            f'=SUMIFS([[Payments.amount]],[[Payments.month]],{key})',
            f'=SUMIFS([[Expenses.total]],[[Expenses.month]],{key})',
            f'={revenue}-SUMIFS([[Expenses.total]],[[Expenses.month]],'
            f'{key})',
            f'=IFERROR(({revenue}-SUMIFS([[Expenses.total]],'
            f'[[Expenses.month]],{key}))/({revenue}),"")',
            f'=COUNTIFS([[Clients.since]],">="&DATE(PlanYear,{m + 1},1),'
            f'[[Clients.since]],"<"&DATE(PlanYear,{m + 2},1))',
        ], ["int", "money0", "money0", "money0", "money0", "money0",
            "percent0", "int"])
        row += 1
    last_month = row - 1
    ws.conditional_format(first_month, 2, last_month, 2,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    ws.conditional_format(first_month, 6, last_month, 6,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_month, 7, last_month, 7,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetMargin",
                           "format": fmts.cf("good")})

    for idx, (src, colour) in enumerate(
            [(2, "#14616B"), (5, "#C86B6B"), (6, "#4E8A63")]):
        letter = chr(65 + src)
        try:
            ws.add_sparkline(first_month + idx, 9, {
                "range": (f"{quote_sheet(NAME)}!{letter}"
                          f"{first_month + 1}:{letter}{last_month + 1}"),
                "type": "column",
                "series_color": colour,
            })
        except Exception:        # pragma: no cover - library differences
            pass
    row += 1
    s.merge(row, 0, LAST_COL,
            "Revenue counts jobs marked Completed or Invoiced. "
            "Collected is money actually received. The little charts "
            "on the right are, top to bottom: revenue, costs, profit.",
            fmts.chrome("note"))
    row += 2

    # ------------------------------------------------------- the pipeline
    row = s.section(row, "Sales pipeline", LAST_COL)
    row = s.table_head(row, 0, ["Stage", "How many", "Value", "Notes"],
                       [20, 12, 14, 50])
    stages = [
        ("New enquiries", '=COUNTIF([[Leads.stage]],"New enquiry")',
         '=SUMIFS([[Leads.value]],[[Leads.stage]],"New enquiry")',
         "Call them today - speed wins more cleaning work than price"),
        ("Quotes out",
         '=COUNTIF([[Quotes.status]],"Sent")+'
         'COUNTIF([[Quotes.status]],"Followed up")',
         '=SUMIFS([[Quotes.total]],[[Quotes.status]],"Sent")+'
         'SUMIFS([[Quotes.total]],[[Quotes.status]],"Followed up")',
         "Money sitting in somebody else's inbox"),
        ("Accepted, not booked",
         '=COUNTIFS([[Quotes.status]],"Accepted",[[Quotes.job]],"")',
         '=SUMIFS([[Quotes.total]],[[Quotes.status]],"Accepted",'
         '[[Quotes.job]],"")',
         "They said yes. Put it in the diary before they forget"),
        ("Won this year",
         '=COUNTIF([[Quotes.won]],"Yes")',
         '=SUMIFS([[Quotes.total]],[[Quotes.won]],"Yes")',
         "Quotes accepted"),
        ("Lost this year",
         '=COUNTIF([[Quotes.won]],"No")',
         '=SUMIFS([[Quotes.total]],[[Quotes.won]],"No")',
         "Worth reading the reasons on the Leads tab"),
    ]
    for label, count, value, note in stages:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(count), fmts.calc("int"))
        ws.write_formula(row, 2, s.reg.render(value), fmts.calc("money0"))
        s.merge(row, 3, LAST_COL, note, fmts.chrome("note"))
        row += 1
    row += 1

    s.note(row,
           "Every figure on this page is calculated. If a number looks "
           "wrong, the fix is on the tab it came from - usually a job "
           "left at the wrong status, or an invoice with no lines.",
           LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/quote_calc.py`

*424 lines · 2 functions*

> `Quote Calculator` - price a clean in thirty seconds, honestly.
>
> Answer eight questions about the property and the workbook prices the
> job four different ways, shows you what it will cost you to deliver,
> what margin each price leaves and what the same client is worth over a
> year at each frequency. Nothing here is hidden: every number is a
> formula you can read.

| Defines | What it does |
|---------|--------------|
| `_round(expr)` | Apply the rounding rule from the Settings tab |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Quote Calculator'), `LAST_COL` (11), `EXTRAS` (4).

````python
"""`Quote Calculator` - price a clean in thirty seconds, honestly.

Answer eight questions about the property and the workbook prices the
job four different ways, shows you what it will cost you to deliver,
what margin each price leaves and what the same client is worth over a
year at each frequency. Nothing here is hidden: every number is a
formula you can read.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import (SheetWriter, list_from, number_check)

NAME = "Quote Calculator"
LAST_COL = 11
EXTRAS = 4

#: how much floor area one cleaner gets through in an hour
_AREA_PER_HOUR = 'IF(AreaUnit="sq m",60,650)'


def _round(expr: str) -> str:
    """Apply the rounding rule from the Settings tab."""
    return (f'IF(PriceRounding="Nearest 10",ROUND(({expr})/10,0)*10,'
            f'IF(PriceRounding="Nearest 5",ROUND(({expr})/5,0)*5,'
            f'IF(PriceRounding="Nearest 1",ROUND({expr},0),'
            f'ROUND({expr},2))))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Pricing & Quote Calculator",
             "Fill in the white cells. The price, the cost, the margin "
             "and the yearly value all work themselves out.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 34, fmts.chrome("label"))
    s.set_col(2, 2, 10, fmts.chrome("label"))
    s.set_col(3, 3, 16)
    s.set_col(4, 4, 14)
    s.set_col(5, LAST_COL, 15, fmts.chrome("note"))

    r = 4

    # ------------------------------------------------------- 1. the job --
    r = s.section(r, "1. What are we cleaning?", LAST_COL)
    r = s.note(r, "You can start from a property you already have on "
                  "file, or just type the numbers in. Nothing here "
                  "changes any of your saved data.",
               LAST_COL, first_col=1, height=24)
    r = s.field(r, "Service", "Regular home clean", name="QcService",
                validate=list_from("ServiceNames"), value_end=4,
                help_text="From the Services tab. Its hours, price and "
                          "uplift are used below.")
    r = s.field(r, "Client type", "Residential", name="QcType",
                validate=list_from("ClientTypeList"), value_end=4)
    r = s.field(r, f"Floor area ({cfg.area_unit})", 1450, style="number",
                name="QcArea", validate=number_check(0, 1000000),
                help_text="Leave at 0 if you do not know it - the room "
                          "count will carry the estimate.")
    r = s.field(r, "Bedrooms", 3, style="int", name="QcBeds",
                validate=number_check(0, 50, decimal=False))
    r = s.field(r, "Bathrooms", 2, style="int", name="QcBaths",
                validate=number_check(0, 50, decimal=False),
                help_text="Bathrooms and kitchens are where the time "
                          "really goes.")
    r = s.field(r, "Other rooms", 4, style="int", name="QcRooms",
                validate=number_check(0, 200, decimal=False))
    r = s.field(r, "Condition", "Normal", name="QcCondition",
                validate=list_from("ConditionList"), value_end=4,
                help_text="Be honest. This multiplies the hours - the "
                          "factors are on the Lists tab.")
    r = s.field(r, "How often", "Weekly", name="QcFrequency",
                validate=list_from("FrequencyList"), value_end=4,
                help_text="Regular work takes less time per visit, so "
                          "it is priced lower.")
    r = s.field(r, "Pets at the property?", "Yes", name="QcPets",
                validate=list_from("YesNoList"))
    r = s.field(r, f"Distance one way ({cfg.distance_unit})", 6,
                style="number1", name="QcDistance",
                validate=number_check(0, 1000))
    r = s.field(r, "Add the call-out fee?", "No", name="QcCallout",
                validate=list_from("YesNoList"),
                help_text=f"Adds the {cfg.currency_symbol}"
                          f"{cfg.callout_fee:,.0f} call-out fee from "
                          f"Settings.")
    r = s.field(r, "Cleaners on the job", cfg.cleaners_per_job,
                style="int", name="QcCleaners",
                validate=number_check(1, 20, decimal=False))
    r = s.field(r, "Hours (override the estimate)", 0, style="number1",
                name="QcHoursOverride", validate=number_check(0, 200),
                help_text="Leave at 0 to let the calculator estimate. "
                          "Type a number if you have walked the job.")
    r = s.field(r, "Hourly rate (override)", 0, style="money",
                name="QcRateOverride", validate=number_check(0, 10000),
                help_text="0 uses the standard rate from Settings.")
    r = s.field(r, "Discount", 0, style="percent0", name="QcDiscount",
                validate=number_check(0, 1),
                help_text="A first-clean discount is the cheapest way "
                          "to win recurring work - but watch the margin "
                          "below.")

    # ----------------------------------------------------- 2. the extras --
    r += 1
    r = s.section(r, "2. Extras", LAST_COL)
    r = s.note(r, "Add-ons priced straight from the Services tab. "
                  "Ovens, interior windows, inside the fridge - the "
                  "jobs that turn an average clean into a good one.",
               LAST_COL, first_col=1, height=24)
    s.write(r, 1, "Extra", fmts.chrome("header_left"))
    s.write(r, 2, "Qty", fmts.chrome("header"))
    s.write(r, 3, "Each", fmts.chrome("header"))
    s.write(r, 4, "Total", fmts.chrome("header"))
    s.merge(r, 5, LAST_COL, "Hours it adds", fmts.chrome("header"))
    r += 1
    extras_top = r
    for i in range(EXTRAS):
        s.write(r, 1, None if i else "Oven clean", fmts.input("text"))
        ws.data_validation(r, 1, r, 1, list_from("ServiceNames"))
        s.write(r, 2, 1 if not i else None, fmts.input("int"))
        ws.data_validation(r, 2, r, 2, number_check(0, 1000))
        ws.write_formula(
            r, 3,
            s.reg.render(f'=IF(B{r + 1}="","",IFERROR(INDEX('
                         f'[[Services.standard_price]],MATCH(B{r + 1},'
                         f'[[Services.name]],0)),0))'),
            fmts.calc("money"))
        ws.write_formula(
            r, 4,
            f'=IF(B{r + 1}="","",IF(C{r + 1}="",0,C{r + 1})*D{r + 1})',
            fmts.calc("money"))
        ws.write_formula(
            r, 5,
            s.reg.render(f'=IF(B{r + 1}="","",IFERROR(INDEX('
                         f'[[Services.hours]],MATCH(B{r + 1},'
                         f'[[Services.name]],0)),0)*IF(C{r + 1}="",0,'
                         f'C{r + 1}))'),
            fmts.calc("number1"))
        s.merge(r, 6, LAST_COL, None, fmts.chrome("note"))
        r += 1
    extras_bottom = r - 1
    extras_sum = f"SUM($E${extras_top + 1}:$E${extras_bottom + 1})"
    extras_hours = f"SUM($F${extras_top + 1}:$F${extras_bottom + 1})"
    s.write(r, 1, "Extras total", fmts.chrome("label_bold"))
    ws.write_formula(r, 4, f"={extras_sum}",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    ws.write_formula(r, 5, f"={extras_hours}",
                     fmts.chrome("total_value", num_format="0.0"))
    book.define_name("QcExtras", f"={quote_sheet(NAME)}!$E${r + 1}")
    book.define_name("QcExtraHours", f"={quote_sheet(NAME)}!$F${r + 1}")
    r += 2

    # ------------------------------------------------- 3. how long it takes
    r = s.section(r, "3. How long it should take", LAST_COL)
    r = s.field(r, "Hours from the room count", None, style="number1",
                mode="calc", name="QcRoomHours",
                formula="=QcBeds*0.45+QcBaths*0.6+QcRooms*0.3",
                help_text="Bedrooms at 27 minutes, bathrooms at 36, "
                          "everything else at 18. Change the weights "
                          "here if your own figures differ.")
    r = s.field(r, "Hours from the floor area", None, style="number1",
                mode="calc", name="QcAreaHours",
                formula=f"=IF(QcArea=0,0,QcArea/{_AREA_PER_HOUR})",
                help_text="One cleaner covers about 650 sq ft "
                          "(60 sq m) an hour on a standard clean.")
    r = s.field(r, "Hours the service allows", None, style="number1",
                mode="calc", name="QcServiceHours",
                formula=s.reg.render(
                    '=IFERROR(INDEX([[Services.hours]],MATCH(QcService,'
                    '[[Services.name]],0)),0)'))
    r = s.field(r, "Base hours used", None, style="number1", mode="calc",
                name="QcBaseHours",
                formula=s.reg.render(
                    '=IF(QcHoursOverride>0,QcHoursOverride,'
                    'IF(AND(IFERROR(INDEX([[Services.model]],'
                    'MATCH(QcService,[[Services.name]],0)),"")'
                    '<>"Per hour",QcServiceHours>0),QcServiceHours,'
                    'MAX(QcRoomHours,QcAreaHours)))'),
                help_text="Your override if you gave one, the service "
                          "time for a fixed-price job, otherwise the "
                          "larger of the room and area estimates.")
    r = s.field(r, "Condition factor", None, style="number1", mode="calc",
                name="QcConditionFactor",
                formula='=IFERROR(INDEX(ConditionFactors,'
                        'MATCH(QcCondition,ConditionNames,0)),1)')
    r = s.field(r, "Hours for this job", None, style="number1",
                mode="calc", name="QcHours",
                formula="=ROUND(QcBaseHours*QcConditionFactor+"
                        "QcExtraHours,2)",
                help_text="Total cleaner-hours, including the extras.")
    r = s.field(r, "Time on site", None, style="number1", mode="calc",
                name="QcVisitHours",
                formula="=ROUND(QcHours/MAX(1,QcCleaners),2)",
                help_text="How long the visit lasts with your crew on "
                          "it. This is what you tell the client.")

    # ---------------------------------------------- 4. the four prices --
    r += 1
    r = s.section(r, "4. Four ways to price it", LAST_COL)
    r = s.note(r, "The same job, priced four different ways. They "
                  "rarely agree - that is the point. Pick the method "
                  "that suits the client and sanity-check it against "
                  "the others.",
               LAST_COL, first_col=1, height=26)
    r = s.field(r, "Rate used", None, style="money", mode="calc",
                name="QcRate",
                formula="=IF(QcRateOverride>0,QcRateOverride,HourlyRate)")
    r = s.field(r, "Service uplift", None, style="percent0", mode="calc",
                name="QcUplift",
                formula=s.reg.render(
                    '=IFERROR(INDEX([[Services.uplift]],MATCH(QcService,'
                    '[[Services.name]],0)),0)'),
                help_text="Deep cleans and end of tenancy work carry "
                          "an uplift - it is set on the Services tab.")
    r = s.field(r, "A. By the hour", None, style="money", mode="calc",
                name="QcPriceHour",
                formula="=ROUND(QcHours*QcRate*(1+QcUplift),2)")
    r = s.field(r, "B. By floor area", None, style="money", mode="calc",
                name="QcPriceArea",
                formula="=ROUND(QcArea*RatePerArea*(1+QcUplift),2)",
                help_text="The usual way to price offices and end of "
                          "tenancy work.")
    r = s.field(r, "C. By the room", None, style="money", mode="calc",
                name="QcPriceRoom",
                formula="=ROUND((QcBeds*RatePerBedroom+QcBaths*"
                        "RatePerBathroom+QcRooms*RatePerRoom)*"
                        "(1+QcUplift),2)")
    r = s.field(r, "D. The catalogue price", None, style="money",
                mode="calc", name="QcPriceService",
                formula=s.reg.render(
                    '=IFERROR(INDEX([[Services.standard_price]],'
                    'MATCH(QcService,[[Services.name]],0)),0)'))
    r = s.field(r, "Price for your target margin", None, style="money",
                mode="calc", name="QcPriceTarget",
                formula="=IF(TargetMargin>=1,0,ROUND(QcCost/"
                        "(1-TargetMargin),2))",
                help_text="What you would have to charge to hit the "
                          "margin on your Settings tab.")
    r = s.field(r, "Price using", "A - by the hour", name="QcMethod",
                value_end=4,
                validate={"validate": "list",
                          "source": ["A - by the hour",
                                     "B - by floor area",
                                     "C - by the room",
                                     "D - catalogue price",
                                     "E - target margin",
                                     "F - my own price"]},
                help_text="Choose the method this quote uses.")
    r = s.field(r, "My own price", 0, style="money", name="QcManual",
                validate=number_check(0, 1000000),
                help_text="Only used when the method above is F.")

    # ------------------------------------------------- 5. what it costs --
    r += 1
    r = s.section(r, "5. What the job costs you", LAST_COL)
    r = s.field(r, "Wages", None, style="money", mode="calc",
                name="QcLabour",
                formula="=ROUND(QcHours*CostPerHour,2)",
                help_text="Hours times the true cost of a cleaner "
                          "hour - pay plus your on-costs.")
    r = s.field(r, "Supplies", None, style="money", mode="calc",
                name="QcSupplies",
                formula="=ROUND(QcHours*SuppliesPerHour,2)")
    r = s.field(r, "Travel", None, style="money", mode="calc",
                name="QcTravel",
                formula="=ROUND(QcDistance*2*MileageRate,2)",
                help_text="There and back, at the mileage rate on the "
                          "Settings tab.")
    r = s.field(r, "Total cost to deliver", None, style="money",
                mode="calc", name="QcCost",
                formula="=QcLabour+QcSupplies+QcTravel")

    # ---------------------------------------------------- 6. the answer --
    r += 1
    r = s.section(r, "6. The quote", LAST_COL)
    r = s.field(r, "Chosen price", None, style="money", mode="calc",
                name="QcChosen",
                formula='=IF(LEFT(QcMethod,1)="A",QcPriceHour,'
                        'IF(LEFT(QcMethod,1)="B",QcPriceArea,'
                        'IF(LEFT(QcMethod,1)="C",QcPriceRoom,'
                        'IF(LEFT(QcMethod,1)="D",QcPriceService,'
                        'IF(LEFT(QcMethod,1)="E",QcPriceTarget,'
                        'QcManual)))))')
    r = s.field(r, "Frequency factor", None, style="number1", mode="calc",
                name="QcFreqFactor",
                formula='=IFERROR(INDEX(FreqFactors,MATCH(QcFrequency,'
                        'FreqNames,0)),1)',
                help_text="Weekly work is priced below a one-off "
                          "because it takes less time. The factors are "
                          "on the Lists tab.")
    r = s.field(r, "After the frequency discount", None, style="money",
                mode="calc", name="QcAfterFreq",
                formula="=ROUND(QcChosen*QcFreqFactor,2)")
    r = s.field(r, "Pet surcharge", None, style="money", mode="calc",
                name="QcPetFee",
                formula='=IF(QcPets="Yes",PetSurcharge,0)')
    r = s.field(r, "Call-out fee", None, style="money", mode="calc",
                name="QcCalloutFee",
                formula='=IF(QcCallout="Yes",CalloutFee,0)')
    r = s.field(r, "Extras", None, style="money", mode="calc",
                name="QcExtrasTotal", formula="=QcExtras")
    r = s.field(r, "Discount", None, style="money", mode="calc",
                name="QcDiscountValue",
                formula="=ROUND((QcAfterFreq+QcPetFee+QcCalloutFee+"
                        "QcExtras)*QcDiscount,2)")
    r = s.field(r, "Price before rounding", None, style="money",
                mode="calc", name="QcRaw",
                formula="=MAX(MinimumCharge,QcAfterFreq+QcPetFee+"
                        "QcCalloutFee+QcExtras-QcDiscountValue)",
                help_text=f"The minimum charge of "
                          f"{cfg.currency_symbol}"
                          f"{cfg.minimum_charge:,.0f} applies to every "
                          f"job.")
    r = s.field(r, "Quote this price", None, style="money", mode="calc",
                name="QcPrice", formula="=" + _round("QcRaw"),
                help_text="Rounded the way you asked on the Settings "
                          "tab. This is the number you give the "
                          "client.")
    r = s.field(r, f"{cfg.tax_label}", None, style="money", mode="calc",
                name="QcTax", formula="=ROUND(QcPrice*TaxRate,2)")
    r = s.field(r, "Total including tax", None, style="money",
                mode="calc", name="QcTotal", formula="=QcPrice+QcTax")
    r = s.field(r, "Deposit to ask for", None, style="money", mode="calc",
                name="QcDeposit",
                formula="=ROUND(QcTotal*DepositPercent,2)")

    # ---------------------------------------------------- 7. the verdict --
    r += 1
    r = s.section(r, "7. Is it worth doing?", LAST_COL)
    r = s.field(r, "Profit on the visit", None, style="money", mode="calc",
                name="QcProfit", formula="=QcPrice-QcCost")
    r = s.field(r, "Margin", None, style="percent0", mode="calc",
                name="QcMargin",
                formula='=IF(QcPrice=0,"",QcProfit/QcPrice)')
    r = s.field(r, "You earn an hour", None, style="money", mode="calc",
                name="QcPerHour",
                formula='=IF(QcHours=0,"",QcProfit/QcHours)',
                help_text="Profit per cleaner-hour. Compare it with "
                          "your other work before you say yes.")
    r = s.field(r, "Price per sq unit", None, style="money", mode="calc",
                name="QcPerArea",
                formula='=IF(QcArea=0,"",QcPrice/QcArea)')
    verdict_row = r
    r = s.field(r, "Verdict", None, mode="calc", value_end=6,
                name="QcVerdict",
                formula='=IF(QcPrice=0,"Fill in the job above",'
                        'IF(QcProfit<0,"No - this job loses money",'
                        'IF(QcMargin<MarginFloor,'
                        '"Careful - the margin is under your floor",'
                        'IF(QcMargin<TargetMargin,'
                        '"Acceptable - but under your target margin",'
                        '"Good - this one is worth having"))))')
    ws.conditional_format(verdict_row, 3, verdict_row, 6, {
        "type": "text", "criteria": "containing", "value": "Good",
        "format": fmts.cf("good")})
    ws.conditional_format(verdict_row, 3, verdict_row, 6, {
        "type": "text", "criteria": "containing", "value": "Careful",
        "format": fmts.cf("warn")})
    ws.conditional_format(verdict_row, 3, verdict_row, 6, {
        "type": "text", "criteria": "containing", "value": "No -",
        "format": fmts.cf("bad")})
    r += 1

    # ------------------------------------------- 8. the year, at each gap
    r = s.section(r, "8. The same client at every frequency", LAST_COL)
    r = s.note(r, "What this job is worth over twelve months if they "
                  "book weekly, fortnightly or monthly. It is the "
                  "argument for offering the regular discount - and "
                  "the reason a one-off deep clean is worth less than "
                  "it looks.",
               LAST_COL, first_col=1, height=30)
    r = s.table_head(r, 1, ["Frequency", "Factor", "Price a visit",
                            "Visits a year", "Value a year",
                            "Profit a year", "Margin"],
                     [22, 9, 14, 12, 14, 14, 10])
    first_freq = r
    for i in range(7):
        n = i + 1
        name = f'IFERROR(INDEX(FreqNames,{n}),"")'
        factor = f'IFERROR(INDEX(FreqFactors,{n}),"")'
        visits = f'IFERROR(INDEX(FreqPerYear,{n}),"")'
        body = _round(f"MAX(MinimumCharge,(QcChosen*{factor}+QcPetFee+"
                      f"QcCalloutFee+QcExtras)*(1-QcDiscount))")
        ws.write_formula(r, 1, f'={name}', fmts.calc("text"))
        ws.write_formula(r, 2, f'={factor}', fmts.calc("number1"))
        ws.write_formula(r, 3, f'=IF({name}="","",{body})',
                         fmts.calc("money"))
        ws.write_formula(r, 4, f'={visits}', fmts.calc("int"))
        ws.write_formula(r, 5, f'=IF(D{r + 1}="","",D{r + 1}*E{r + 1})',
                         fmts.calc("money0"))
        ws.write_formula(r, 6,
                         f'=IF(D{r + 1}="","",(D{r + 1}-QcCost)*'
                         f'E{r + 1})', fmts.calc("money0"))
        ws.write_formula(r, 7,
                         f'=IF(OR(D{r + 1}="",D{r + 1}=0),"",'
                         f'(D{r + 1}-QcCost)/D{r + 1})',
                         fmts.calc("percent0"))
        r += 1
    ws.conditional_format(first_freq, 5, r - 1, 5,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    ws.conditional_format(first_freq, 7, r - 1, 7,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginFloor",
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_freq, 7, r - 1, 7,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetMargin",
                           "format": fmts.cf("good")})
    r += 1

    s.note(r,
           "Happy with the number? Raise it on the Quotes tab, add the "
           "lines on Quote Lines, and the follow-up reminders start by "
           "themselves. Nothing you do on this page is saved - it is a "
           "calculator, not a record.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/calendar_page.py`

*176 lines · 1 function*

> `Calendar` - the month at a glance, and the next jobs in order.
>
> A real diary built from the Jobs tab: a month grid showing how busy
> each day is, and a running list of the next thirty visits with the
> client, the address, the cleaner and the price. Change the month in
> one cell and the whole page moves.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Calendar'), `LAST_COL` (10), `WEEKS` (6), `AGENDA` (30).

````python
"""`Calendar` - the month at a glance, and the next jobs in order.

A real diary built from the Jobs tab: a month grid showing how busy
each day is, and a running list of the next thirty visits with the
client, the address, the cleaner and the price. Change the month in
one cell and the whole page moves.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Calendar"
LAST_COL = 10
WEEKS = 6
AGENDA = 30


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Job Calendar",
             "Every booked visit, laid out as a month and as a list. "
             "It is the Jobs tab seen sideways - change a job there "
             "and it moves here.", LAST_COL)

    s.set_col(0, 0, 16)
    for c in range(1, 8):
        s.set_col(c, c, 15)
    s.set_col(8, LAST_COL, 14)

    # ------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Month starting", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=MonthStart", fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("CalStart", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "Showing", fmts.chrome("label_right"))
    s.merge(row, 3, 4, None, fmts.calc("text"))
    ws.write_formula(row, 3, '=TEXT(CalStart,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 5, 5, "Jobs", fmts.chrome("label_right"))
    ws.write_formula(
        row, 6,
        s.reg.render('=COUNTIFS([[Jobs.date]],">="&CalStart,'
                     '[[Jobs.date]],"<"&DATE(YEAR(CalStart),'
                     'MONTH(CalStart)+1,1),[[Jobs.status]],'
                     '"<>Cancelled")'), fmts.calc("int"))
    s.merge(row, 7, 7, "Value", fmts.chrome("label_right"))
    ws.write_formula(
        row, 8,
        s.reg.render('=SUMIFS([[Jobs.price]],[[Jobs.date]],">="&CalStart,'
                     '[[Jobs.date]],"<"&DATE(YEAR(CalStart),'
                     'MONTH(CalStart)+1,1),[[Jobs.status]],'
                     '"<>Cancelled")'), fmts.calc("money0"))
    row += 1
    row = s.note(row,
                 "Type any date in the white cell - the grid always "
                 "starts on the week containing it and runs for six "
                 "weeks.", LAST_COL, height=22)
    row += 1

    # --------------------------------------------------- the month grid --
    row = s.section(row, "The month", LAST_COL)
    first_cell = ('CalStart-WEEKDAY(CalStart,IF(WeekStart="Sunday",1,2))'
                  '+1')
    for i in range(7):
        ws.write_formula(
            row, i + 1,
            f'=TEXT({first_cell}+{i},"dddd")',
            fmts.chrome("header"))
    s.write(row, 0, "Week", fmts.chrome("header_left"))
    row += 1

    grid_top = row
    for w in range(WEEKS):
        offset = w * 7
        # week label
        s.merge(row, 0, 0, None, fmts.calc("date"))
        ws.write_formula(row, 0, f'={first_cell}+{offset}',
                         fmts.chrome("label_bold", num_format="d mmm"))
        for i in range(7):
            ws.write_formula(
                row, i + 1, f'={first_cell}+{offset + i}',
                fmts.chrome("kpi_label", num_format="ddd d"))
        row += 1
        s.write(row, 0, "Jobs", fmts.chrome("label"))
        for i in range(7):
            cell = f"{chr(66 + i)}${row}"
            ws.write_formula(
                row, i + 1,
                s.reg.render(f'=IF(COUNTIFS([[Jobs.date]],{cell},'
                             f'[[Jobs.status]],"<>Cancelled")=0,"",'
                             f'COUNTIFS([[Jobs.date]],{cell},'
                             f'[[Jobs.status]],"<>Cancelled")&" jobs, "&'
                             f'TEXT(SUMIFS([[Jobs.hours_each]],'
                             f'[[Jobs.date]],{cell},[[Jobs.status]],'
                             f'"<>Cancelled"),"0.0")&"h")'),
                fmts.cell("text", "calc"))
        row += 1
        s.write(row, 0, "Value", fmts.chrome("label"))
        for i in range(7):
            cell = f"{chr(66 + i)}${row - 1}"
            ws.write_formula(
                row, i + 1,
                s.reg.render(f'=IF(SUMIFS([[Jobs.price]],[[Jobs.date]],'
                             f'{cell},[[Jobs.status]],"<>Cancelled")=0,'
                             f'"",SUMIFS([[Jobs.price]],[[Jobs.date]],'
                             f'{cell},[[Jobs.status]],"<>Cancelled"))'),
                fmts.calc("money0"))
            ws.conditional_format(row, i + 1, row, i + 1, {
                "type": "cell", "criteria": ">", "value": 0,
                "format": fmts.cf("good")})
        s.row_height(row, 16)
        row += 1
    grid_bottom = row - 1
    ws.conditional_format(grid_top, 1, grid_bottom, 7, {
        "type": "text", "criteria": "containing", "value": "jobs",
        "format": fmts.cf("info")})
    row += 1

    # ------------------------------------------------------- the agenda --
    row = s.section(row, "The next thirty visits", LAST_COL)
    row = s.table_head(row, 0, ["Date", "Day", "Start", "Client",
                                "Property", "Service", "Cleaner",
                                "Hours", "Price", "Status", "Action"],
                       [11, 8, 8, 20, 22, 20, 16, 8, 11, 13, 26])
    first_row = row
    for n in range(1, AGENDA + 1):
        pick = f'MATCH({n},[[Jobs.future_rank]],0)'
        for col, (key, style) in enumerate([
                ("date", "date"), ("weekday", "text"), ("start", "time"),
                ("client_name", "text"), ("property_name", "text"),
                ("service", "text"), ("employee", "text"),
                ("hours_each", "number1"), ("price", "money"),
                ("status", "status"), ("alert", "status")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Jobs.{key}]],{pick}),"")'),
                fmts.calc(style))
        row += 1
    last_row = row - 1
    for col, text, fmt in [(9, "Completed", "good"), (9, "Confirmed", "info"),
                           (9, "Booked", "info"), (9, "Enquiry", "warn"),
                           (10, "Assign", "warn"), (10, "Confirm", "warn"),
                           (10, "Overdue", "bad"), (10, "Late", "bad"),
                           (10, "OK", "good")]:
        ws.conditional_format(first_row, col, last_row, col, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(fmt)})
    ws.conditional_format(first_row, 8, last_row, 8,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    row += 1

    s.note(row,
           "The list is in date order, nearest first, and it only "
           "shows visits from today onwards. An empty row means you "
           "have run out of bookings - which is either a quiet week "
           "or a sales problem.",
           LAST_COL, height=30)
    row += 2

    s.note(row,
           "A note on calendar sync: a spreadsheet cannot push these "
           "visits into Google Calendar, Outlook or a phone by "
           "itself. What it can do is print - set the print area to "
           "the agenda and the week goes on the van dashboard. If you "
           "need live sync, export the Jobs tab to CSV and import it "
           "into your calendar app.",
           LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/staff_schedule.py`

*305 lines · 1 function*

> `Staff Schedule` - the week's rota, and who has room for one more.
>
> Pick a week and every cleaner's day-by-day load appears: hours booked,
> jobs, how full their week is against their contracted hours and
> whether anything clashes. Underneath it, the attendance for the week
> and the jobs still waiting for a name.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Staff Schedule'), `LAST_COL` (12), `STAFF_ROWS` (12), `UNASSIGNED` (8).

````python
"""`Staff Schedule` - the week's rota, and who has room for one more.

Pick a week and every cleaner's day-by-day load appears: hours booked,
jobs, how full their week is against their contracted hours and
whether anything clashes. Underneath it, the attendance for the week
and the jobs still waiting for a name.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Staff Schedule"
LAST_COL = 12
STAFF_ROWS = 12
UNASSIGNED = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Employee Schedule & Attendance",
             "Who is working when, how full their week is and who has "
             "capacity for one more job.", LAST_COL)

    s.set_col(0, 0, 20)
    for c in range(1, 8):
        s.set_col(c, c, 11)
    for c in range(8, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Week starting", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=WeekStartDate", fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("RotaStart", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 3, None, fmts.calc("text"))
    ws.write_formula(row, 2,
                     '="Week "&TEXT(RotaStart,"d mmm")&" to "&'
                     'TEXT(RotaStart+6,"d mmm yyyy")', fmts.calc("text"))
    s.merge(row, 4, 4, "Jobs", fmts.chrome("label_right"))
    ws.write_formula(
        row, 5,
        s.reg.render('=COUNTIFS([[Jobs.date]],">="&RotaStart,'
                     '[[Jobs.date]],"<="&RotaStart+6,[[Jobs.status]],'
                     '"<>Cancelled")'), fmts.calc("int"))
    s.merge(row, 6, 6, "Hours", fmts.chrome("label_right"))
    ws.write_formula(
        row, 7,
        s.reg.render('=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],'
                     '">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6,'
                     '[[Jobs.status]],"<>Cancelled")'),
        fmts.calc("number1"))
    s.merge(row, 8, 9, "Value of the week", fmts.chrome("label_right"))
    ws.write_formula(
        row, 10,
        s.reg.render('=SUMIFS([[Jobs.price]],[[Jobs.date]],'
                     '">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6,'
                     '[[Jobs.status]],"<>Cancelled")'),
        fmts.calc("money0"))
    row += 2

    # ---------------------------------------------------------- the rota --
    row = s.section(row, "Hours booked, by cleaner", LAST_COL)
    s.write(row, 0, "Cleaner", fmts.chrome("header_left"))
    for i in range(7):
        ws.write_formula(row, i + 1, f"=RotaStart+{i}",
                         fmts.chrome("header", num_format="ddd d"))
    for i, head in enumerate(["Week hours", "Contracted", "How full",
                              "Jobs", "Check"]):
        s.write(row, 8 + i, head, fmts.chrome("header"))
    day_row = row
    row += 1

    first_staff = row
    for n in range(1, STAFF_ROWS + 1):
        pick = f'MATCH({n},[[Employees.name_rank]],0)'
        name_cell = f"$A${row + 1}"
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX([[Employees.name]],{pick}),"")'),
            fmts.chrome("label_bold"))
        for i in range(7):
            cell = f"{chr(66 + i)}${day_row + 1}"
            ws.write_formula(
                row, i + 1,
                s.reg.render(
                    f'=IF({name_cell}="","",IF(SUMIFS('
                    f'[[Jobs.hours_each]],[[Jobs.employee]],{name_cell},'
                    f'[[Jobs.date]],{cell},[[Jobs.status]],'
                    f'"<>Cancelled")=0,"",SUMIFS([[Jobs.hours_each]],'
                    f'[[Jobs.employee]],{name_cell},[[Jobs.date]],'
                    f'{cell},[[Jobs.status]],"<>Cancelled")))'),
                fmts.calc("number1"))
        # week totals
        ws.write_formula(
            row, 8,
            s.reg.render(f'=IF({name_cell}="","",SUMIFS('
                         f'[[Jobs.hours_each]],[[Jobs.employee]],'
                         f'{name_cell},[[Jobs.date]],">="&RotaStart,'
                         f'[[Jobs.date]],"<="&RotaStart+6,'
                         f'[[Jobs.status]],"<>Cancelled"))'),
            fmts.calc("number1"))
        ws.write_formula(
            row, 9,
            s.reg.render(f'=IF({name_cell}="","",IFERROR(INDEX('
                         f'[[Employees.hours]],MATCH({name_cell},'
                         f'[[Employees.name]],0)),""))'),
            fmts.calc("number1"))
        ws.write_formula(
            row, 10,
            f'=IF(OR({name_cell}="",J{row + 1}="",J{row + 1}=0),"",'
            f'I{row + 1}/J{row + 1})', fmts.calc("percent0"))
        ws.write_formula(
            row, 11,
            s.reg.render(f'=IF({name_cell}="","",COUNTIFS('
                         f'[[Jobs.employee]],{name_cell},[[Jobs.date]],'
                         f'">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6,'
                         f'[[Jobs.status]],"<>Cancelled"))'),
            fmts.calc("int"))
        ws.write_formula(
            row, 12,
            s.reg.render(
                f'=IF({name_cell}="","",IF(COUNTIFS([[Jobs.clash]],'
                f'"Clash",[[Jobs.employee]],{name_cell},[[Jobs.date]],'
                f'">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6)>0,'
                f'"Clash",IF(K{row + 1}="","",IF(K{row + 1}>1,'
                f'"Overloaded",IF(K{row + 1}<0.6,"Has room","OK")))))'),
            fmts.calc("status"))
        row += 1
    last_staff = row - 1
    ws.conditional_format(first_staff, 1, last_staff, 7,
                          {"type": "3_color_scale",
                           "min_color": "#FFFFFF",
                           "mid_color": "#DCEEF0",
                           "max_color": "#2F8F9D"})
    ws.conditional_format(first_staff, 10, last_staff, 10,
                          {"type": "cell", "criteria": ">", "value": 1,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_staff, 10, last_staff, 10,
                          {"type": "cell", "criteria": "<", "value": 0.6,
                           "format": fmts.cf("warn")})
    for text, fmt in [("Clash", "bad"), ("Overloaded", "bad"),
                      ("Has room", "info"), ("OK", "good")]:
        ws.conditional_format(first_staff, 12, last_staff, 12, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(fmt)})
    row += 1

    # --------------------------------------------------- the day totals --
    s.write(row, 0, "Team total", fmts.chrome("label_bold"))
    for i in range(7):
        cell = f"{chr(66 + i)}${day_row + 1}"
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],'
                         f'{cell},[[Jobs.status]],"<>Cancelled")'),
            fmts.chrome("total_value", num_format="0.0"))
    ws.write_formula(
        row, 8,
        s.reg.render('=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],'
                     '">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6,'
                     '[[Jobs.status]],"<>Cancelled")'),
        fmts.chrome("total_value", num_format="0.0"))
    s.merge(row, 9, LAST_COL,
            "Hours on site, not hours paid - add travel between jobs.",
            fmts.chrome("note"))
    row += 2

    # ----------------------------------------------------- capacity check
    row = s.section(row, "Can you take more work this week?", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "This week", "What it means"],
                       [26, 14, 60])
    checks = [
        ("Cleaners available",
         '=COUNTIF([[Employees.status]],"Active")',
         '="Everyone not on leave or off sick"'),
        ("Hours you could sell",
         '=COUNTIF([[Employees.status]],"Active")*ProductiveHours*5',
         '="Active cleaners times "&TEXT(ProductiveHours,"0.0")&'
         '" billable hours over five days"'),
        ("Hours booked",
         '=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],">="&RotaStart,'
         '[[Jobs.date]],"<="&RotaStart+6,[[Jobs.status]],'
         '"<>Cancelled")', '="What is in the diary"'),
        ("Spare capacity",
         '=COUNTIF([[Employees.status]],"Active")*ProductiveHours*5-'
         'SUMIFS([[Jobs.hours_each]],[[Jobs.date]],">="&RotaStart,'
         '[[Jobs.date]],"<="&RotaStart+6,[[Jobs.status]],'
         '"<>Cancelled")',
         '="Hours you could still sell this week. Negative means you '
         'need help"'),
        ("Jobs with nobody on them",
         '=COUNTIFS([[Jobs.employee]],"",[[Jobs.client]],"<>",'
         '[[Jobs.date]],">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6)',
         '="Assign these before the week starts"'),
        ("Double bookings",
         '=COUNTIFS([[Jobs.clash]],"Clash",[[Jobs.date]],'
         '">="&RotaStart,[[Jobs.date]],"<="&RotaStart+6)',
         '="Same cleaner, same time, two places"'),
    ]
    first_check = row
    for label, formula, meaning in checks:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula),
                         fmts.calc("number1"))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(meaning), fmts.calc("text"))
        row += 1
    ws.conditional_format(first_check + 3, 1, first_check + 3, 1,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_check + 4, 1, first_check + 5, 1,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # ------------------------------------------------- unassigned jobs --
    row = s.section(row, "Jobs still waiting for a cleaner", LAST_COL)
    row = s.table_head(row, 0, ["Date", "Client", "Property", "Service",
                                "Hours", "Price", "Priority"],
                       [12, 22, 24, 20, 9, 11, 12])
    for n in range(1, UNASSIGNED + 1):
        pick = f'MATCH({n},[[Jobs.unassigned_rank]],0)'
        for col, (key, style) in enumerate([
                ("date", "date"), ("client_name", "text"),
                ("property_name", "text"), ("service", "text"),
                ("hours_each", "number1"), ("price", "money"),
                ("priority", "text")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Jobs.{key}]],{pick}),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    # ------------------------------------------------ attendance summary
    row = s.section(row, "Attendance this week", LAST_COL)
    row = s.table_head(row, 0, ["Cleaner", "Hours paid", "Hours billed",
                                "Utilisation", "Late", "Absent",
                                "Cost", "Approved?"],
                       [20, 12, 12, 12, 9, 9, 12, 11])
    first_att = row
    for n in range(1, STAFF_ROWS + 1):
        pick = f'MATCH({n},[[Employees.name_rank]],0)'
        name_cell = f"$A${row + 1}"
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX([[Employees.name]],{pick}),"")'),
            fmts.chrome("label_bold"))
        week = (f'[[Timesheets.employee]],{name_cell},'
                f'[[Timesheets.date]],">="&RotaStart,'
                f'[[Timesheets.date]],"<="&RotaStart+6')
        cells = [
            (f'=IF({name_cell}="","",SUMIFS([[Timesheets.hours]],{week}))',
             "number1"),
            (f'=IF({name_cell}="","",SUMIFS([[Timesheets.hours]],{week},'
             f'[[Timesheets.billable]],"Yes"))', "number1"),
            (f'=IF(OR({name_cell}="",B{row + 1}=0),"",C{row + 1}/'
             f'B{row + 1})', "percent0"),
            (f'=IF({name_cell}="","",COUNTIFS([[Timesheets.attendance]],'
             f'"Late",{week}))', "int"),
            (f'=IF({name_cell}="","",COUNTIFS([[Timesheets.attendance]],'
             f'"Absent - no show",{week})+COUNTIFS('
             f'[[Timesheets.attendance]],"Absent - told us",{week}))',
             "int"),
            (f'=IF({name_cell}="","",SUMIFS([[Timesheets.cost]],{week}))',
             "money"),
            (f'=IF({name_cell}="","",COUNTIFS([[Timesheets.approved]],'
             f'"Yes",{week})&" of "&COUNTIFS([[Timesheets.employee]],'
             f'{name_cell},[[Timesheets.date]],">="&RotaStart,'
             f'[[Timesheets.date]],"<="&RotaStart+6))', "text"),
        ]
        for col, (formula, style) in enumerate(cells, start=1):
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    last_att = row - 1
    ws.conditional_format(first_att, 3, last_att, 3,
                          {"type": "cell", "criteria": ">=",
                           "value": "=UtilisationTarget",
                           "format": fmts.cf("good")})
    ws.conditional_format(first_att, 3, last_att, 3,
                          {"type": "cell", "criteria": "<",
                           "value": "=UtilisationTarget",
                           "format": fmts.cf("warn")})
    ws.conditional_format(first_att, 4, last_att, 5,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    s.note(row,
           "The rota reads the Jobs tab - it does not change it. To "
           "move a job, change the cleaner or the date on the Jobs "
           "tab and this page follows. Holidays and sickness come "
           "from the Timesheets tab.",
           LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/retention.py`

*233 lines · 1 function*

> `Client Retention` - keeping the clients you already have.
>
> Winning a new cleaning client costs several times what it costs to
> keep one. This page shows how much of your income is recurring, which
> clients are quietly drifting away, what each one is worth over a year
> and who to ring this week.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Client Retention'), `LAST_COL` (11), `LIST_ROWS` (8).

````python
"""`Client Retention` - keeping the clients you already have.

Winning a new cleaning client costs several times what it costs to
keep one. This page shows how much of your income is recurring, which
clients are quietly drifting away, what each one is worth over a year
and who to ring this week.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Client Retention"
LAST_COL = 11
LIST_ROWS = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Client Retention & Recurring Services",
             "How much of your work comes back every week, who is "
             "slipping away and what it is worth to keep them.",
             LAST_COL)

    s.set_col(0, 0, 24)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.note(row,
                 "Everything on this page reads the Clients and Jobs "
                 "tabs. Nothing here needs filling in - but the phone "
                 "calls it suggests are the most profitable half hour "
                 "in your week.", LAST_COL, height=26)
    row += 1

    # ------------------------------------------------------------ KPIs --
    active = ('COUNTIF([[Clients.status]],"Active")+'
              'COUNTIF([[Clients.status]],"Regular")')
    tiles_1 = [
        ("Active clients", f"={active}", "int", "Active or Regular"),
        ("Recurring clients",
         '=COUNTIFS([[Clients.frequency]],"<>One-off",'
         '[[Clients.frequency]],"<>",[[Clients.next_job]],"<>")',
         "int", "with a booking in the diary"),
        ("Recurring value a year",
         '=SUMIFS([[Clients.year_value]],[[Clients.frequency]],'
         '"<>One-off",[[Clients.frequency]],"<>")', "money0",
         "at today's frequencies"),
        ("Average client value a year",
         '=IFERROR(AVERAGEIFS([[Clients.year_value]],'
         '[[Clients.year_value]],">0"),"")', "money0",
         "across everyone on the books"),
    ]
    tiles_2 = [
        ("Clients at risk",
         '=COUNTIF([[Clients.retention]],"At risk")', "int",
         f"nothing for {cfg.inactive_client_days}+ days"),
        ("Slipping away",
         '=COUNTIF([[Clients.retention]],"Slipping")', "int",
         "no next visit booked"),
        ("Dormant or lost",
         '=COUNTIF([[Clients.retention]],"Dormant")+'
         'COUNTIF([[Clients.retention]],"Lost")', "int",
         "gone, for now"),
        ("Retention rate",
         f'=IFERROR(({active})/({active}+COUNTIF('
         f'[[Clients.retention]],"Dormant")+COUNTIF('
         f'[[Clients.retention]],"Lost")),"")', "percent0",
         f"target {cfg.retention_target:.0%}"),
    ]
    for tiles in (tiles_1, tiles_2):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, s.reg.render(formula), style=style,
                  span=3, note=note)
        row += 4

    # ------------------------------------------------- recurring mix ----
    row = s.section(row, "How much of the work comes back", LAST_COL)
    row = s.table_head(row, 0, ["Frequency", "Clients", "Value a year",
                                "Share of the year", "Jobs booked"],
                       [22, 11, 15, 15, 13])
    first_freq = row
    for i in range(7):
        n = i + 1
        name = f'IFERROR(INDEX(FreqNames,{n}),"")'
        cell = f"$A${row + 1}"
        ws.write_formula(row, 0, f'={name}', fmts.calc("text"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IF({cell}="","",COUNTIFS('
                         f'[[Clients.frequency]],{cell}))'),
            fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=IF({cell}="","",SUMIFS('
                         f'[[Clients.year_value]],[[Clients.frequency]],'
                         f'{cell}))'), fmts.calc("money0"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF({cell}="","",IFERROR(C{row + 1}/'
                         f'SUM([[Clients.year_value]]),""))'),
            fmts.calc("percent0"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IF({cell}="","",COUNTIFS('
                         f'[[Jobs.frequency]],{cell}))'),
            fmts.calc("int"))
        row += 1
    ws.conditional_format(first_freq, 2, row - 1, 2,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    row += 1
    s.merge(row, 0, LAST_COL,
            "A business built on weekly and fortnightly work survives "
            "a quiet January. One built on one-off deep cleans starts "
            "every month from zero.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------------- ring these ----
    row = s.section(row, "Ring these clients this week", LAST_COL)
    row = s.table_head(row, 0, ["Client", "Phone", "Standing",
                                "Last clean", "Days", "Next booked",
                                "Worth a year", "Jobs done",
                                "What to say"],
                       [22, 15, 12, 12, 8, 12, 13, 10, 30])
    first_risk = row
    for n in range(1, LIST_ROWS + 1):
        pick = f'MATCH({n},[[Clients.risk_rank]],0)'
        for col, (key, style) in enumerate([
                ("name", "text"), ("phone", "text"),
                ("retention", "status"), ("last_job", "date"),
                ("days_since", "int"), ("next_job", "date"),
                ("year_value", "money"), ("jobs_done", "int")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Clients.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 8,
            f'=IF($A${row + 1}="","",IF($F${row + 1}="",'
            f'"Nothing booked - offer them a slot",'
            f'"Check they are happy with the service"))',
            fmts.calc("text"))
        row += 1
    ws.conditional_format(first_risk, 4, row - 1, 4,
                          {"type": "cell", "criteria": ">",
                           "value": "=InactiveClientDays",
                           "format": fmts.cf("bad")})
    row += 1

    # ------------------------------------------------- most valuable ----
    row = s.section(row, "The clients worth keeping most", LAST_COL)
    row = s.table_head(row, 0, ["Client", "Type", "Frequency",
                                "Worth a year", "Profit to date",
                                "Margin", "Rating", "Complaints",
                                "Standing"],
                       [22, 15, 14, 13, 13, 10, 10, 11, 12])
    first_top = row
    for n in range(1, LIST_ROWS + 1):
        pick = f'MATCH({n},[[Clients.year_value_rank]],0)'
        for col, (key, style) in enumerate([
                ("name", "text"), ("type", "text"),
                ("frequency", "text"), ("year_value", "money"),
                ("profit", "money"), ("margin", "percent0"),
                ("rating", "number1"), ("complaints", "int"),
                ("retention", "status")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Clients.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        row += 1
    ws.conditional_format(first_top, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    ws.conditional_format(first_top, 5, row - 1, 5,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginFloor",
                           "format": fmts.cf("bad")})
    row += 1

    # ------------------------------------------------------ contracts ---
    row = s.section(row, "Contracts and renewals", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "How many", "What it means"],
                       [28, 12, 60])
    rows = [
        ("Clients on a contract",
         '=COUNTIFS([[Clients.contract]],"<>No contract",'
         '[[Clients.contract]],"<>")',
         '="Rolling or fixed term - the work you can count on"'),
        ("Renewing soon",
         '=COUNTIFS([[Clients.contract_end]],">="&TodayDate,'
         '[[Clients.contract_end]],"<="&TodayDate+RenewalWarningDays)',
         '="Within "&RenewalWarningDays&" days. Ask early and you '
         'usually keep them"'),
        ("Contracts already expired",
         '=COUNTIFS([[Clients.contract_end]],"<"&TodayDate,'
         '[[Clients.contract_end]],"<>")',
         '="Still cleaning, no agreement in place"'),
        ("Clients with no next visit",
         '=COUNTIFS([[Clients.next_job]],"",[[Clients.status]],'
         '"Active")+COUNTIFS([[Clients.next_job]],"",'
         '[[Clients.status]],"Regular")',
         '="Active clients with an empty diary - the quickest money '
         'in the workbook"'),
        ("Won back this year",
         '=COUNTIF([[Clients.retention]],"Won back")',
         '="Mark a returning client as Won back on the Clients tab"'),
    ]
    first_c = row
    for label, formula, meaning in rows:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula), fmts.calc("int"))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(meaning), fmts.calc("text"))
        row += 1
    ws.conditional_format(first_c + 1, 1, first_c + 3, 1,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    s.note(row,
           "Three numbers worth remembering: a cleaning client who "
           "stays two years is worth roughly double one who stays "
           "one; most clients who leave never complained first; and "
           "the single best predictor of leaving is an empty diary. "
           "That is exactly what the first table on this page "
           "lists.", LAST_COL, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/reports.py`

*379 lines · 1 function*

> `Reports & Analytics` - the questions a cleaning business should ask.
>
> Where the money comes from, which services earn their keep, which
> advertising works, how many quotes turn into work, what the team
> costs and who owes you money. Every table reads the registers, so the
> answers change the moment you enter a job.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Reports'), `LAST_COL` (11), `TOP` (8).

````python
"""`Reports & Analytics` - the questions a cleaning business should ask.

Where the money comes from, which services earn their keep, which
advertising works, how many quotes turn into work, what the team
costs and who owes you money. Every table reads the registers, so the
answers change the moment you enter a job.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Reports"
LAST_COL = 11
TOP = 8

_REV = ('SUMIFS([[Jobs.price]],{cond},[[Jobs.status]],"Completed")+'
        'SUMIFS([[Jobs.price]],{cond},[[Jobs.status]],"Invoiced")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Reports & Business Analytics",
             "Where the work comes from, what it earns and what it "
             "costs - all from the tabs you have already filled in.",
             LAST_COL)

    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.note(row,
                 "Every figure is live. Filter or sort the registers "
                 "as much as you like - these tables add up the whole "
                 "column either way.", LAST_COL, height=22)
    row += 1

    # ------------------------------------------------- by service -------
    row = s.section(row, "Which services earn their keep", LAST_COL)
    row = s.table_head(row, 0, ["Service", "Times booked", "Revenue",
                                "Hours", "Earns an hour", "Margin",
                                "Times quoted", "Conversion"],
                       [26, 12, 13, 10, 13, 10, 12, 12])
    first_svc = row
    for n in range(1, TOP + 1):
        pick = f'MATCH({n},[[Services.revenue_rank]],0)'
        cell = f"$A${row + 1}"
        for col, (key, style) in enumerate([
                ("name", "text"), ("times_booked", "int"),
                ("revenue", "money0")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Services.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF({cell}="","",SUMIFS([[Jobs.est_hours]],'
                         f'[[Jobs.service]],{cell},[[Jobs.status]],'
                         f'"Completed")+SUMIFS([[Jobs.est_hours]],'
                         f'[[Jobs.service]],{cell},[[Jobs.status]],'
                         f'"Invoiced"))'), fmts.calc("number1"))
        for col, key, style in [(4, "hourly_yield", "money"),
                                (5, "margin", "percent0"),
                                (6, "times_quoted", "int")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Services.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 7,
            f'=IF(OR({cell}="",G{row + 1}=0),"",B{row + 1}/G{row + 1})',
            fmts.calc("percent0"))
        row += 1
    ws.conditional_format(first_svc, 2, row - 1, 2,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    ws.conditional_format(first_svc, 4, row - 1, 4,
                          {"type": "3_color_scale",
                           "min_color": "#FBD9DD",
                           "mid_color": "#FFF0CE",
                           "max_color": "#D7F2E3"})
    row += 2

    # ------------------------------------------------- by client type ---
    row = s.section(row, "Residential or commercial?", LAST_COL)
    row = s.table_head(row, 0, ["Client type", "Clients", "Jobs",
                                "Revenue", "Average job",
                                "Share of revenue"],
                       [24, 11, 10, 14, 13, 15])
    types = ["Residential", "Commercial office", "Retail",
             "Restaurant & cafe", "Holiday let / Airbnb", "Landlord",
             "Letting agent", "Property manager", "Medical & dental",
             "School & nursery", "Gym & studio"]
    first_type = row
    for kind in types:
        cond = f'[[Jobs.client_type]],"{kind}"'
        s.write(row, 0, kind, fmts.chrome("label_bold"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=COUNTIFS([[Clients.type]],"{kind}")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=COUNTIFS({cond},[[Jobs.status]],'
                         f'"<>Cancelled")'), fmts.calc("int"))
        ws.write_formula(row, 3,
                         s.reg.render("=" + _REV.format(cond=cond)),
                         fmts.calc("money0"))
        ws.write_formula(
            row, 4, f'=IF(C{row + 1}=0,"",D{row + 1}/C{row + 1})',
            fmts.calc("money"))
        ws.write_formula(
            row, 5,
            s.reg.render('=IFERROR(D{r}/('
                         + _REV.format(cond='[[Jobs.client]],"<>"')
                         + '),"")').replace("{r}", str(row + 1)),
            fmts.calc("percent0"))
        row += 1
    ws.conditional_format(first_type, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    row += 2

    # ---------------------------------------------- marketing ----------
    row = s.section(row, "Which advertising actually works", LAST_COL)
    row = s.table_head(row, 0, ["Came from", "Enquiries", "Won", "Lost",
                                "Win rate", "Clients now",
                                "Revenue from them"],
                       [24, 11, 9, 9, 10, 12, 16])
    sources = ["Google search", "Google Business Profile", "Facebook",
               "Instagram", "Nextdoor", "Yelp", "Directory listing",
               "Referral - client", "Referral - staff",
               "Flyer / leaflet", "Van signage", "Word of mouth",
               "Letting agent"]
    first_src = row
    for src in sources:
        s.write(row, 0, src, fmts.chrome("label_bold"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=COUNTIFS([[Leads.source]],"{src}")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=COUNTIFS([[Leads.source]],"{src}",'
                         f'[[Leads.stage]],"Won")'), fmts.calc("int"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=COUNTIFS([[Leads.source]],"{src}",'
                         f'[[Leads.stage]],"Lost")'), fmts.calc("int"))
        ws.write_formula(
            row, 4,
            f'=IF(C{row + 1}+D{row + 1}=0,"",C{row + 1}/'
            f'(C{row + 1}+D{row + 1}))', fmts.calc("percent0"))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=COUNTIFS([[Clients.source]],"{src}")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=SUMIFS([[Clients.invoiced]],'
                         f'[[Clients.source]],"{src}")'),
            fmts.calc("money0"))
        row += 1
    ws.conditional_format(first_src, 6, row - 1, 6,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    ws.conditional_format(first_src, 4, row - 1, 4,
                          {"type": "cell", "criteria": ">=", "value": 0.5,
                           "format": fmts.cf("good")})
    row += 1
    s.merge(row, 0, LAST_COL,
            "Referrals almost always win this table. If an advert has "
            "produced ten enquiries and no clients, it is not an "
            "advert, it is a subscription.", fmts.chrome("note"))
    row += 2

    # ---------------------------------------------- quotes -------------
    row = s.section(row, "Quoting and converting", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "Value", "What it means"],
                       [26, 14, 58])
    quote_rows = [
        ("Quotes sent", '=COUNTIF([[Quotes.won]],"Yes")+'
                        'COUNTIF([[Quotes.won]],"No")+'
                        'COUNTIF([[Quotes.status]],"Sent")+'
                        'COUNTIF([[Quotes.status]],"Followed up")',
         "int", '="Everything that left your desk with a price on it"'),
        ("Quotes won", '=COUNTIF([[Quotes.won]],"Yes")', "int",
         '="Accepted by the client"'),
        ("Win rate",
         '=IFERROR(COUNTIF([[Quotes.won]],"Yes")/(COUNTIF('
         '[[Quotes.won]],"Yes")+COUNTIF([[Quotes.won]],"No")),"")',
         "percent0",
         '="Target "&TEXT(QuoteWinTarget,"0%")&". Winning every quote '
         'means you are too cheap"'),
        ("Average quote",
         '=IFERROR(AVERAGEIFS([[Quotes.total]],[[Quotes.total]],">0"),'
         '"")', "money", '="Across every quote you have raised"'),
        ("Average won quote",
         '=IFERROR(AVERAGEIFS([[Quotes.total]],[[Quotes.won]],"Yes"),'
         '"")', "money",
         '="Compare it with the line above - if won quotes are much '
         'smaller, you are losing the big ones"'),
        ("Value still out there",
         '=SUMIFS([[Quotes.total]],[[Quotes.status]],"Sent")+'
         'SUMIFS([[Quotes.total]],[[Quotes.status]],"Followed up")',
         "money0", '="Quotes with no answer yet"'),
        ("Average days to answer",
         '=IFERROR(AVERAGEIFS([[Quotes.days_waiting]],'
         '[[Quotes.days_waiting]],">0"),"")', "number1",
         '="How long clients take to come back to you"'),
        ("Expired, never chased",
         '=COUNTIF([[Quotes.expired]],"Yes")', "int",
         '="Quotes that ran out of time. Every one is money you '
         'already did the work for"'),
    ]
    for label, formula, style, meaning in quote_rows:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula), fmts.calc(style))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(meaning), fmts.calc("text"))
        row += 1
    row += 1

    # ---------------------------------------------- operations ---------
    row = s.section(row, "How the work goes", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "Value", "What it means"],
                       [26, 14, 58])
    ops_rows = [
        ("Jobs completed",
         '=COUNTIF([[Jobs.status]],"Completed")+'
         'COUNTIF([[Jobs.status]],"Invoiced")', "int",
         '="Visits that happened"'),
        ("Cancelled",
         '=COUNTIF([[Jobs.status]],"Cancelled")', "int",
         '="Every cancellation is an empty slot you could have sold"'),
        ("No access",
         '=COUNTIF([[Jobs.status]],"No access")', "int",
         '="Turned up, could not get in. Charge for these or they '
         'repeat"'),
        ("Average rating",
         '=IFERROR(AVERAGEIFS([[Jobs.rating]],[[Jobs.rating]],">0"),"")',
         "number1", '="Target "&TEXT(SatisfactionTarget,"0.0")&" out '
                    'of 5"'),
        ("Checklists finished",
         '=IFERROR(AVERAGEIFS([[Jobs.checklist_pct]],'
         '[[Jobs.checklist_pct]],">0"),"")', "percent0",
         '="Across every job with a checklist against it"'),
        ("Jobs that ran over",
         '=COUNTIF([[Job Costing.alert]],"Check - ran over*")', "int",
         '="More hours than the estimate. A pricing problem, not a '
         'people problem"'),
        ("Loss-making jobs",
         '=COUNTIF([[Job Costing.alert]],"Loss*")', "int",
         '="Cost more than they earned"'),
        ("Average profit a job",
         '=IFERROR(AVERAGEIFS([[Job Costing.profit]],'
         '[[Job Costing.revenue]],">0"),"")', "money",
         '="After wages, supplies and travel"'),
        ("Average profit an hour",
         '=IFERROR(AVERAGEIFS([[Job Costing.per_hour]],'
         '[[Job Costing.per_hour]],">0"),"")', "money",
         '="The number to compare every job against"'),
        ("Team utilisation",
         '=IFERROR(SUMIFS([[Timesheets.hours]],[[Timesheets.billable]],'
         '"Yes")/SUM([[Timesheets.hours]]),"")', "percent0",
         '="Billed hours as a share of paid hours. Target "&'
         'TEXT(UtilisationTarget,"0%")'),
        ("Complaints per hundred jobs",
         '=IFERROR(COUNTIF([[Complaints.date]],">0")/(COUNTIF('
         '[[Jobs.status]],"Completed")+COUNTIF([[Jobs.status]],'
         '"Invoiced"))*100,"")', "number1",
         '="Under two is good. Under one is exceptional"'),
    ]
    for label, formula, style, meaning in ops_rows:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula), fmts.calc(style))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(meaning), fmts.calc("text"))
        row += 1
    row += 1

    # ---------------------------------------------- expenses -----------
    row = s.section(row, "Where the money goes", LAST_COL)
    row = s.table_head(row, 0, ["Category", "This month", "This year",
                                "Share", "Monthly average"],
                       [26, 14, 14, 10, 15])
    categories = ["Cleaning supplies", "Equipment purchase",
                  "Equipment repair", "Vehicle fuel",
                  "Vehicle maintenance", "Business insurance",
                  "Wages & payroll", "Subcontractors",
                  "Marketing & advertising", "Software & apps",
                  "Phone & internet", "Office & rent", "Uniforms",
                  "Training & certification", "Accounting & legal",
                  "Bank & card fees", "Laundry", "Waste disposal",
                  "Travel & parking", "Other"]
    first_cat = row
    for cat in categories:
        s.write(row, 0, cat, fmts.chrome("label_bold"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=SUMIFS([[Expenses.total]],'
                         f'[[Expenses.category]],"{cat}",'
                         f'[[Expenses.month]],DashKey)'),
            fmts.calc("money"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=SUMIFS([[Expenses.total]],'
                         f'[[Expenses.category]],"{cat}",'
                         f'[[Expenses.date]],">="&DATE(PlanYear,1,1),'
                         f'[[Expenses.date]],"<="&DATE(PlanYear,12,31))'),
            fmts.calc("money0"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IFERROR(C{row + 1}/SUMIFS('
                         f'[[Expenses.total]],[[Expenses.date]],'
                         f'">="&DATE(PlanYear,1,1),[[Expenses.date]],'
                         f'"<="&DATE(PlanYear,12,31)),"")'),
            fmts.calc("percent0"))
        ws.write_formula(row, 4, f'=C{row + 1}/12', fmts.calc("money"))
        row += 1
    last_cat = row - 1
    ws.conditional_format(first_cat, 2, last_cat, 2,
                          {"type": "data_bar", "bar_color": "#C86B6B"})
    s.write(row, 0, "Total", fmts.chrome("label_bold"))
    for col, letter in [(1, "B"), (2, "C")]:
        ws.write_formula(
            row, col,
            f"=SUM({letter}{first_cat + 1}:{letter}{last_cat + 1})",
            fmts.chrome("total_value", num_format=cfg.money_format_0))
    row += 2

    # ---------------------------------------------- debtors ------------
    row = s.section(row, "Who owes you money", LAST_COL)
    row = s.table_head(row, 0, ["Age", "Invoices", "Amount",
                                "What to do"],
                       [18, 11, 14, 56])
    buckets = [
        ("Not due yet", "Keep an eye on it"),
        ("Due now", "Today is the day to send the gentle reminder"),
        ("1-30 days", "Phone, do not email. It works far better"),
        ("31-60 days", "Stop work for this client until they pay"),
        ("61-90 days", "Final demand, in writing, with your terms"),
        ("90+ days", "Decide: formal recovery, or write it off and "
                     "learn from it"),
    ]
    first_b = row
    for label, advice in buckets:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=COUNTIF([[Invoices.age]],"{label}")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=SUMIFS([[Invoices.outstanding]],'
                         f'[[Invoices.age]],"{label}")'),
            fmts.calc("money"))
        s.merge(row, 3, LAST_COL, advice, fmts.chrome("note"))
        row += 1
    ws.conditional_format(first_b + 2, 2, row - 1, 2,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("bad")})
    s.write(row, 0, "Total owed", fmts.chrome("label_bold"))
    ws.write_formula(row, 2, f"=SUM(C{first_b + 1}:C{row})",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    s.note(row,
           "Three tables on this page pay for the workbook on their "
           "own: the advertising table (stop paying for what does not "
           "work), the service table (do more of what earns most an "
           "hour) and the debtors table (chase at thirty days, not "
           "ninety).", LAST_COL, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/financial.py`

*341 lines · 1 function*

> `Financial Plan` - targets, break-even and what if you grew.
>
> Three honest questions: how much work do you have to do to cover the
> bills, are you on track against your own targets, and what happens if
> you put the rate up or take on another cleaner. The white cells are
> yours to play with; nothing here touches your records.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Financial Plan'), `LAST_COL` (11), `MONTHS` (12).

````python
"""`Financial Plan` - targets, break-even and what if you grew.

Three honest questions: how much work do you have to do to cover the
bills, are you on track against your own targets, and what happens if
you put the rate up or take on another cleaner. The white cells are
yours to play with; nothing here touches your records.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, number_check

NAME = "Financial Plan"
LAST_COL = 11
MONTHS = 12

_REV = ('SUMIFS([[Jobs.price]],[[Jobs.month]],{key},[[Jobs.status]],'
        '"Completed")+SUMIFS([[Jobs.price]],[[Jobs.month]],{key},'
        '[[Jobs.status]],"Invoiced")')

#: the same revenue, with no month filter on it
_REV_ALL = ('SUMIFS([[Jobs.price]],[[Jobs.status]],"Completed")+'
            'SUMIFS([[Jobs.price]],[[Jobs.status]],"Invoiced")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Financial Planning & Forecasting",
             "What you must earn to break even, how the year is "
             "tracking and what growing would actually be worth.",
             LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 32, fmts.chrome("label"))
    s.set_col(2, 2, 10, fmts.chrome("label"))
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 14)
    s.set_col(5, LAST_COL, 14)

    r = 4

    # ----------------------------------------------------- break-even ---
    r = s.section(r, "1. Break-even: the number you must hit", LAST_COL)
    r = s.note(r, "Fixed costs are the ones that arrive whether you "
                  "clean anything or not - insurance, the van, "
                  "software, your phone. Fill those in and the "
                  "workbook tells you how many jobs a month keep the "
                  "lights on.", LAST_COL, first_col=1, height=30)
    r = s.field(r, "Fixed costs a month", 1800, style="money",
                name="FpFixed", validate=number_check(0, 1000000),
                help_text="Insurance, vehicle, software, phone, "
                          "accountant, rent, your own basic draw.")
    r = s.field(r, "Average job value", None, style="money", mode="calc",
                name="FpAvgJob",
                formula=s.reg.render(
                    '=IFERROR((' + _REV_ALL + ')/(COUNTIF('
                    '[[Jobs.status]],"Completed")+COUNTIF('
                    '[[Jobs.status]],"Invoiced")),MinimumCharge)'),
                help_text="Taken from the work you have finished or "
                          "billed. If you have none yet it falls back "
                          "to the minimum charge.")
    r = s.field(r, "Average hours a job", None, style="number1",
                mode="calc", name="FpAvgHours",
                formula=s.reg.render(
                    '=IFERROR(AVERAGEIFS([[Jobs.est_hours]],'
                    '[[Jobs.est_hours]],">0"),2)'))
    r = s.field(r, "Variable cost a job", None, style="money",
                mode="calc", name="FpVarCost",
                formula="=ROUND(FpAvgHours*(CostPerHour+"
                        "SuppliesPerHour),2)",
                help_text="Wages and supplies for an average job. "
                          "Travel sits in fixed costs unless you drive "
                          "a lot.")
    r = s.field(r, "Contribution a job", None, style="money",
                mode="calc", name="FpContribution",
                formula="=FpAvgJob-FpVarCost",
                help_text="What each job leaves towards the fixed "
                          "costs. If this is near zero, no amount of "
                          "work will save the business.")
    r = s.field(r, "Jobs a month to break even", None, style="int",
                mode="calc", name="FpBreakEvenJobs",
                formula='=IF(FpContribution<=0,"Never - the price is '
                        'too low",ROUNDUP(FpFixed/FpContribution,0))')
    r = s.field(r, "Revenue a month to break even", None, style="money0",
                mode="calc", name="FpBreakEvenRevenue",
                formula='=IF(FpContribution<=0,"",ROUNDUP(FpFixed/'
                        'FpContribution,0)*FpAvgJob)')
    r = s.field(r, "Hours a month to break even", None, style="number1",
                mode="calc", name="FpBreakEvenHours",
                formula='=IF(FpContribution<=0,"",ROUNDUP(FpFixed/'
                        'FpContribution,0)*FpAvgHours)')
    r = s.field(r, "Jobs a month you are doing now", None, style="int",
                mode="calc", name="FpJobsNow",
                formula=s.reg.render(
                    '=ROUND((COUNTIFS([[Jobs.date]],">="&TodayDate-90,'
                    '[[Jobs.date]],"<="&TodayDate,[[Jobs.status]],'
                    '"Completed")+COUNTIFS([[Jobs.date]],">="&'
                    'TodayDate-90,[[Jobs.date]],"<="&TodayDate,'
                    '[[Jobs.status]],"Invoiced"))/3,0)'),
                help_text="The work you actually finished or billed in "
                          "the last three months, as a monthly "
                          "average.")
    verdict = r
    r = s.field(r, "Where that leaves you", None, mode="calc",
                value_end=7, name="FpBreakEvenVerdict",
                formula='=IF(FpContribution<=0,"Put the price up - '
                        'every job loses money",IF(FpJobsNow>='
                        'FpBreakEvenJobs,"Covered - you are "&'
                        '(FpJobsNow-FpBreakEvenJobs)&" jobs a month '
                        'past break-even","Short by "&'
                        '(FpBreakEvenJobs-FpJobsNow)&" jobs a month"))')
    ws.conditional_format(verdict, 3, verdict, 7, {
        "type": "text", "criteria": "containing", "value": "Covered",
        "format": fmts.cf("good")})
    ws.conditional_format(verdict, 3, verdict, 7, {
        "type": "text", "criteria": "containing", "value": "Short",
        "format": fmts.cf("warn")})
    ws.conditional_format(verdict, 3, verdict, 7, {
        "type": "text", "criteria": "containing", "value": "Put the",
        "format": fmts.cf("bad")})
    r += 2

    # -------------------------------------------------- the year plan ---
    r = s.section(r, "2. The year against your targets", LAST_COL)
    r = s.table_head(r, 1, ["Month", "Target", "Revenue", "Difference",
                            "Costs", "Profit", "Running profit",
                            "Hit target?"],
                     [10, 13, 13, 13, 12, 12, 14, 12])
    first_month = r
    for m in range(MONTHS):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        revenue = _REV.format(key=key)
        ws.write_formula(r, 1, f'=TEXT(DATE(PlanYear,{m + 1},1),"mmm")',
                         fmts.calc("text"))
        ws.write_formula(r, 2, "=MonthlyRevenueTarget",
                         fmts.calc("money0"))
        ws.write_formula(r, 3, s.reg.render(f"={revenue}"),
                         fmts.calc("money0"))
        ws.write_formula(r, 4, f"=D{r + 1}-C{r + 1}", fmts.calc("money0"))
        ws.write_formula(
            r, 5,
            s.reg.render(f'=SUMIFS([[Expenses.total]],'
                         f'[[Expenses.month]],{key})'),
            fmts.calc("money0"))
        ws.write_formula(r, 6, f"=D{r + 1}-F{r + 1}", fmts.calc("money0"))
        if m == 0:
            ws.write_formula(r, 7, f"=G{r + 1}", fmts.calc("money0"))
        else:
            ws.write_formula(r, 7, f"=H{r}+G{r + 1}", fmts.calc("money0"))
        ws.write_formula(
            r, 8,
            f'=IF(D{r + 1}=0,"",IF(D{r + 1}>=C{r + 1},"Yes","No"))',
            fmts.calc("status"))
        r += 1
    last_month = r - 1
    ws.conditional_format(first_month, 3, last_month, 3,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    ws.conditional_format(first_month, 4, last_month, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_month, 6, last_month, 6,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    for text, fmt in [("Yes", "good"), ("No", "warn")]:
        ws.conditional_format(first_month, 8, last_month, 8, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(fmt)})
    # totals
    ws.write(r, 1, "Year", fmts.chrome("label_bold"))
    for col, letter in [(2, "C"), (3, "D"), (4, "E"), (5, "F"),
                        (6, "G")]:
        ws.write_formula(
            r, col,
            f"=SUM({letter}{first_month + 1}:{letter}{last_month + 1})",
            fmts.chrome("total_value", num_format=cfg.money_format_0))
    ws.write_formula(r, 7, f"=H{last_month + 1}",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format_0))
    ws.write_formula(
        r, 8,
        f'=IF(D{r + 1}=0,"",TEXT(G{r + 1}/D{r + 1},"0%")&" margin")',
        fmts.chrome("total_value"))
    r += 2

    # ------------------------------------------------- what if ----------
    r = s.section(r, "3. What if the business grew?", LAST_COL)
    r = s.note(r, "Change the three white cells and watch the table "
                  "below. It is the quickest way to see whether you "
                  "need more clients, better prices or another pair "
                  "of hands.", LAST_COL, first_col=1, height=26)
    r = s.field(r, "Put the hourly rate up by", 0.05, style="percent0",
                name="FpRateRise", validate=number_check(-0.5, 2))
    r = s.field(r, "Extra regular clients a month", 2, style="int",
                name="FpNewClients", validate=number_check(0, 500,
                                                           decimal=False))
    r = s.field(r, "Extra cleaners", 1, style="int", name="FpNewStaff",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "Jobs a new client books a month", 4, style="number1",
                name="FpClientJobs", validate=number_check(0, 60))
    r += 1

    r = s.table_head(r, 1, ["Scenario", "Revenue a month",
                            "Costs a month", "Profit a month",
                            "Profit a year", "Change"],
                     [30, 15, 15, 15, 15, 14])
    base_rev = "FpJobsNow*FpAvgJob"
    base_cost = "FpFixed+FpJobsNow*FpVarCost"
    scenarios = [
        ("As you are now", base_rev, base_cost),
        ("With the rate rise",
         f"({base_rev})*(1+FpRateRise)", base_cost),
        ("With the extra clients",
         f"({base_rev})+FpNewClients*FpClientJobs*FpAvgJob",
         f"({base_cost})+FpNewClients*FpClientJobs*FpVarCost"),
        ("With another cleaner",
         f"({base_rev})+FpNewStaff*ProductiveHours*21*HourlyRate*"
         f"UtilisationTarget",
         f"({base_cost})+FpNewStaff*ProductiveHours*21*CostPerHour"),
        ("All three together",
         f"(({base_rev})+FpNewClients*FpClientJobs*FpAvgJob+"
         f"FpNewStaff*ProductiveHours*21*HourlyRate*UtilisationTarget)"
         f"*(1+FpRateRise)",
         f"({base_cost})+FpNewClients*FpClientJobs*FpVarCost+"
         f"FpNewStaff*ProductiveHours*21*CostPerHour"),
    ]
    first_s = r
    for label, rev, cost in scenarios:
        ws.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, f"={rev}", fmts.calc("money0"))
        ws.write_formula(r, 3, f"={cost}", fmts.calc("money0"))
        ws.write_formula(r, 4, f"=C{r + 1}-D{r + 1}", fmts.calc("money0"))
        ws.write_formula(r, 5, f"=E{r + 1}*12", fmts.calc("money0"))
        if r == first_s:
            ws.write(r, 6, "baseline", fmts.chrome("note"))
        else:
            ws.write_formula(
                r, 6, f"=E{r + 1}-E${first_s + 1}", fmts.calc("money0"))
        r += 1
    last_s = r - 1
    ws.conditional_format(first_s, 4, last_s, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_s + 1, 6, last_s, 6,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    r += 1
    s.merge(r, 1, LAST_COL,
            "The cleaner scenario assumes 21 working days a month at "
            "your utilisation target - be careful, a new cleaner "
            "costs money from day one and earns it from week three.",
            fmts.chrome("note"))
    r += 2

    # ------------------------------------------------- cash position ----
    r = s.section(r, "4. Cash position today", LAST_COL)
    r = s.table_head(r, 1, ["Measure", "Amount", "What it means"],
                     [28, 15, 54])
    cash = [
        ("Invoiced, not yet paid", '=SUM([[Invoices.outstanding]])',
         '="Money you have earned and not received"'),
        ("Overdue right now",
         '=SUMIFS([[Invoices.outstanding]],[[Invoices.age]],'
         '"1-30 days")+SUMIFS([[Invoices.outstanding]],'
         '[[Invoices.age]],"31-60 days")+SUMIFS('
         '[[Invoices.outstanding]],[[Invoices.age]],"61-90 days")+'
         'SUMIFS([[Invoices.outstanding]],[[Invoices.age]],"90+ days")',
         '="Past the due date. Chase today"'),
        ("Work done, not invoiced",
         '=SUMIFS([[Jobs.price]],[[Jobs.status]],"Completed",'
         '[[Jobs.invoice]],"")',
         '="Raise these invoices and the cash follows"'),
        ("Deposits taken",
         '=SUMIFS([[Payments.amount]],[[Payments.deposit]],"Yes")',
         '="Money in advance - lovely, but it is not profit yet"'),
        ("Booked work, not yet done",
         '=SUMIFS([[Jobs.price]],[[Jobs.date]],">"&TodayDate,'
         '[[Jobs.status]],"<>Cancelled")',
         '="The diary, in money"'),
        ("Quotes awaiting an answer",
         '=SUMIFS([[Quotes.total]],[[Quotes.status]],"Sent")+'
         'SUMIFS([[Quotes.total]],[[Quotes.status]],"Followed up")',
         '="Possible, not certain"'),
        ("Expenses this month",
         '=SUMIFS([[Expenses.total]],[[Expenses.month]],DashKey)',
         '="Against a budget of "&TEXT(MonthlyExpenseBudget,"#,##0")'),
        ("Stock on the shelf",
         '=SUM([[Supplies.stock_value]])',
         '="Cash you have already spent, sitting in the cupboard"'),
    ]
    first_cash = r
    for label, formula, meaning in cash:
        ws.write(r, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(r, 2, s.reg.render(formula), fmts.calc("money0"))
        s.merge(r, 3, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(r, 3, s.reg.render(meaning), fmts.calc("text"))
        r += 1
    ws.conditional_format(first_cash + 1, 2, first_cash + 1, 2,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_cash + 2, 2, first_cash + 2, 2,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    r += 1

    # ------------------------------------------------- the tax reminder -
    r = s.section(r, "5. Money you should not spend", LAST_COL)
    r = s.field(r, "Set aside for tax", 0.25, style="percent0",
                name="FpTaxSet", validate=number_check(0, 1),
                help_text="A guess, not advice. Ask an accountant "
                          "what your rate really is.")
    r = s.field(r, "Profit so far this year", None, style="money0",
                mode="calc", name="FpYearProfit",
                formula=s.reg.render(
                    '=SUMIFS([[Jobs.price]],[[Jobs.status]],"Completed",'
                    '[[Jobs.date]],">="&DATE(PlanYear,1,1))+SUMIFS('
                    '[[Jobs.price]],[[Jobs.status]],"Invoiced",'
                    '[[Jobs.date]],">="&DATE(PlanYear,1,1))-SUMIFS('
                    '[[Expenses.total]],[[Expenses.date]],">="&'
                    'DATE(PlanYear,1,1))'))
    r = s.field(r, "Put this much aside", None, style="money0",
                mode="calc", name="FpTaxPot",
                formula="=MAX(0,ROUND(FpYearProfit*FpTaxSet,0))",
                help_text="Move it to a separate account the day the "
                          "money lands. It is the single habit that "
                          "keeps small businesses out of trouble.")
    r += 1

    s.note(r,
           "This page is a planning tool, not an accounts package. It "
           "does not know about loans, drawings, depreciation or your "
           "local tax rules, and it is not financial advice. Use it "
           "to decide what to charge and when to hire - then show the "
           "real numbers to a real accountant.",
           LAST_COL, first_col=1, height=40)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/printables.py`

*538 lines · 5 functions*

> The three printable forms.
>
> * `Print Quote` - a quote on one side of A4 with your letterhead, the
>   lines, the total and an acceptance box the client can sign.
> * `Print Invoice` - the same for an invoice, with payment details, the
>   due date and what is still owed.
> * `Print Job Sheet` - the sheet the cleaner takes to the property:
>   access details, the checklist, time in and out, and a sign-off.
>
> Pick an ID at the top of each page and the whole form fills itself in
> from the registers. Nothing is typed twice.

| Defines | What it does |
|---------|--------------|
| `_letterhead(s, fmts, row, last_col, heading)` | The business name, address and contact block |
| `_pick(s, fmts, book, row, label, default, list_name, define, help_text)` | The 'which one are we printing' picker |
| `build_quote(book, fmts, cfg, reg)` |  |
| `build_invoice(book, fmts, cfg, reg)` |  |
| `build_job(book, fmts, cfg, reg)` |  |

**Module constants:** `QUOTE_NAME` ('Print Quote'), `INVOICE_NAME` ('Print Invoice'), `JOB_NAME` ('Print Job Sheet'), `QUOTE_LINES` (12), `INVOICE_LINES` (12), `TASK_LINES` (18).

````python
"""The three printable forms.

* `Print Quote` - a quote on one side of A4 with your letterhead, the
  lines, the total and an acceptance box the client can sign.
* `Print Invoice` - the same for an invoice, with payment details, the
  due date and what is still owed.
* `Print Job Sheet` - the sheet the cleaner takes to the property:
  access details, the checklist, time in and out, and a sign-off.

Pick an ID at the top of each page and the whole form fills itself in
from the registers. Nothing is typed twice.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from

QUOTE_NAME = "Print Quote"
INVOICE_NAME = "Print Invoice"
JOB_NAME = "Print Job Sheet"

QUOTE_LINES = 12
INVOICE_LINES = 12
TASK_LINES = 18


def _letterhead(s: SheetWriter, fmts, row: int, last_col: int,
                heading: str) -> int:
    """The business name, address and contact block."""
    ws = s.ws
    s.merge(row, 0, 3, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0, "=BizName", fmts.chrome("print_title"))
    s.merge(row, 4, last_col, heading, fmts.chrome("print_title"))
    row += 1
    s.merge(row, 0, 3, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0,
                     '=BizAddress&IF(BizArea=""," "," - "&BizArea)',
                     fmts.chrome("print_sub"))
    s.merge(row, 4, last_col, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 4,
                     '=BizPhone&"   "&BizEmail&IF(BizWebsite="","",'
                     '"   "&BizWebsite)', fmts.chrome("print_sub"))
    row += 1
    s.merge(row, 0, last_col, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0,
                     '=IF(BizRegNo="","","Registration "&BizRegNo&'
                     '"   ")&IF(BizInsurance="","","Insurance: "&'
                     'BizInsurance)', fmts.chrome("print_sub"))
    return row + 2


def _pick(s: SheetWriter, fmts, book, row: int, label: str,
          default: str, list_name: str, define: str,
          help_text: str) -> int:
    """The 'which one are we printing' picker."""
    ws = s.ws
    s.write(row, 0, label, fmts.chrome("label_bold"))
    s.write(row, 1, default, fmts.input("id"))
    ws.data_validation(row, 1, row, 1, list_from(list_name))
    book.define_name(define,
                     f"={quote_sheet(s.name)}!$B${row + 1}")
    s.merge(row, 2, 7, help_text, fmts.chrome("note"))
    return row + 2


# ---------------------------------------------------------------------------
# Print Quote
# ---------------------------------------------------------------------------

def build_quote(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = QUOTE_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Quote", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 16)
    s.set_col(1, 1, 24)
    s.set_col(2, 2, 26)
    s.set_col(3, 3, 10)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 12)
    s.set_col(7, 7, 14)

    row = 4
    row = _pick(s, fmts, book, row, "Quote to print", "Q-0001",
                "QuoteIDs", "PrintQuote",
                "Pick a quote. Everything below fills itself in. "
                "Print with File, Print - the page is already sized "
                "for A4.")
    row = _letterhead(s, fmts, row, LAST_COL, "QUOTE")

    def q(key: str) -> str:
        return s.reg.render(f'=IFERROR(INDEX([[Quotes.{key}]],'
                            f'MATCH(PrintQuote,[[Quotes.id]],0)),"")')

    # --------------------------------------------------- the two blocks --
    s.write(row, 0, "Quote for", fmts.chrome("label_bold"))
    s.write(row, 4, "Quote number", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, "=PrintQuote", fmts.cell("id", "plain"))
    row += 1
    s.merge(row, 0, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 0,
                     s.reg.render('=IFERROR(INDEX([[Quotes.client_name]],'
                                  'MATCH(PrintQuote,[[Quotes.id]],0)),"")'
                                  ), fmts.cell("text", "plain"))
    s.write(row, 4, "Date", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, q("date"), fmts.cell("date", "plain"))
    row += 1
    s.merge(row, 0, 3, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0, q("address"), fmts.cell("wrap", "plain"))
    s.write(row, 4, "Valid until", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, q("valid_until"), fmts.cell("date", "plain"))
    s.row_height(row, 32)
    row += 1
    s.write(row, 0, "Service", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 1, q("service"), fmts.cell("text", "plain"))
    s.write(row, 4, "How often", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, q("frequency"), fmts.cell("text", "plain"))
    row += 2

    # ------------------------------------------------------- the lines --
    row = s.table_head(row, 0, ["Item", "What is included", "Qty",
                                "Price", "Line total", "Optional?"],
                       [18, 36, 8, 12, 13, 11])
    first_line = row
    for n in range(1, QUOTE_LINES + 1):
        key = f'PrintQuote&"|{n}"'
        for col, (field, style) in enumerate([
                ("service", "text"), ("wording", "wrap"),
                ("qty", "number1"), ("price_used", "money"),
                ("total", "money"), ("optional", "text")]):
            if col == 1:
                s.merge(row, 1, 2, None, fmts.cell("wrap", "plain"))
            target = col if col == 0 else col + 1
            ws.write_formula(
                row, target,
                s.reg.render(f'=IFERROR(INDEX([[Quote Lines.{field}]],'
                             f'MATCH({key},[[Quote Lines.key]],0)),"")'),
                fmts.cell(style, "plain"))
        row += 1
    last_line = row - 1
    ws.conditional_format(first_line, 6, last_line, 6, {
        "type": "text", "criteria": "containing", "value": "Yes",
        "format": fmts.cf("info")})
    row += 1

    # ------------------------------------------------------ the totals --
    totals = [("Subtotal", q("subtotal"), "money"),
              ("Discount", q("discount_value"), "money"),
              (cfg.tax_label, q("tax"), "money"),
              ("Total", q("total"), "money"),
              ("Deposit to confirm", q("deposit"), "money")]
    for i, (label, formula, style) in enumerate(totals):
        s.merge(row, 3, 4, label,
                fmts.chrome("total_label" if i == 3 else "label_right"))
        ws.write_formula(
            row, 5, formula,
            fmts.chrome("total_value", num_format=cfg.money_format)
            if i == 3 else fmts.cell(style, "plain"))
        if i == 0:
            s.merge(row, 0, 2,
                    "Prices hold until the valid-until date above.",
                    fmts.chrome("note"))
        if i == 4:
            s.merge(row, 0, 2,
                    "The balance is due on the terms below.",
                    fmts.chrome("note"))
        row += 1
    row += 1

    # -------------------------------------------------- what's included --
    row = s.section(row, "What is included", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0, q("notes"), fmts.cell("wrap", "plain"))
    s.row_height(row, 46)
    row += 1
    row = s.section(row, "Terms", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        s.reg.render('=IF(IFERROR(INDEX([[Quotes.terms]],MATCH('
                     'PrintQuote,[[Quotes.id]],0)),"")="",'
                     '"Payment within "&PaymentTerms&" days of '
                     'invoice. Cancellations inside 24 hours may be '
                     'charged. We bring our own supplies and '
                     'equipment unless agreed otherwise.",'
                     'INDEX([[Quotes.terms]],MATCH(PrintQuote,'
                     '[[Quotes.id]],0)))'),
        fmts.cell("wrap", "plain"))
    s.row_height(row, 46)
    row += 2

    # -------------------------------------------------- the acceptance --
    row = s.section(row, "Acceptance", LAST_COL)
    s.merge(row, 0, LAST_COL,
            "I accept this quote and agree to the terms above.",
            fmts.chrome("note"))
    row += 1
    for label, span in [("Signed", 3), ("Print name", 3)]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, span, None, fmts.chrome("form_line"))
        s.write(row, 4, "Date", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 26)
        row += 1
    row += 1
    s.merge(row, 0, LAST_COL,
            "A signature on a printed page is not an electronic "
            "signature. If you need one that stands up legally, use "
            "a proper e-signature service - then record the date "
            "here.", fmts.chrome("note"))
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(row, 0,
                     '="Quote produced by "&BizName&" on "&'
                     'TEXT(TodayDate,"d mmmm yyyy")',
                     fmts.chrome("note"))

    s.finish(freeze_row=None, print_area=f"A1:H{row + 2}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Invoice
# ---------------------------------------------------------------------------

def build_invoice(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = INVOICE_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Invoice", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 16)
    s.set_col(1, 1, 24)
    s.set_col(2, 2, 26)
    s.set_col(3, 3, 10)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 12)
    s.set_col(7, 7, 14)

    row = 4
    row = _pick(s, fmts, book, row, "Invoice to print", "INV-0001",
                "InvoiceIDs", "PrintInvoice",
                "Pick an invoice. Check the totals against the "
                "Invoices tab before you send it.")
    row = _letterhead(s, fmts, row, LAST_COL, "INVOICE")

    def inv(key: str) -> str:
        return s.reg.render(f'=IFERROR(INDEX([[Invoices.{key}]],'
                            f'MATCH(PrintInvoice,[[Invoices.id]],0)),"")')

    s.write(row, 0, "Invoice to", fmts.chrome("label_bold"))
    s.write(row, 4, "Invoice number", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, "=PrintInvoice", fmts.cell("id", "plain"))
    row += 1
    s.merge(row, 0, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 0, inv("client_name"),
                     fmts.cell("text", "plain"))
    s.write(row, 4, "Date", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, inv("date"), fmts.cell("date", "plain"))
    row += 1
    s.merge(row, 0, 3, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0, inv("address"), fmts.cell("wrap", "plain"))
    s.write(row, 4, "Due date", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, inv("due_date"), fmts.cell("date", "plain"))
    s.row_height(row, 32)
    row += 1
    s.write(row, 0, "Period", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 1, inv("period"), fmts.cell("text", "plain"))
    s.write(row, 4, "Your reference", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, inv("po"), fmts.cell("text", "plain"))
    row += 2

    row = s.table_head(row, 0, ["Date", "What we did", "Qty", "Price",
                                "Line total", "Tax?"],
                       [12, 42, 8, 12, 13, 9])
    first_line = row
    for n in range(1, INVOICE_LINES + 1):
        key = f'PrintInvoice&"|{n}"'
        fields = [("date", "date"), ("wording", "wrap"),
                  ("qty", "number1"), ("price_used", "money"),
                  ("total", "money"), ("taxable", "text")]
        for col, (field, style) in enumerate(fields):
            if col == 1:
                s.merge(row, 1, 2, None, fmts.cell("wrap", "plain"))
            target = col if col == 0 else col + 1
            ws.write_formula(
                row, target,
                s.reg.render(f'=IFERROR(INDEX([[Invoice Lines.{field}]],'
                             f'MATCH({key},[[Invoice Lines.key]],0)),"")'),
                fmts.cell(style, "plain"))
        row += 1
    last_line = row - 1
    ws.conditional_format(first_line, 5, last_line, 5,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("info")})
    row += 1

    totals = [("Subtotal", inv("subtotal")),
              ("Discount", inv("discount_value")),
              (cfg.tax_label, inv("tax")),
              ("Invoice total", inv("total")),
              ("Paid so far", inv("paid")),
              ("Amount due", inv("outstanding"))]
    for i, (label, formula) in enumerate(totals):
        big = label in ("Invoice total", "Amount due")
        s.merge(row, 3, 4, label,
                fmts.chrome("total_label" if big else "label_right"))
        ws.write_formula(
            row, 5, formula,
            fmts.chrome("total_value", num_format=cfg.money_format)
            if big else fmts.cell("money", "plain"))
        row += 1
    row += 1

    row = s.section(row, "How to pay", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        '="Payment within "&PaymentTerms&" days. "&'
        'IF(BizPayDetails="","Bank details on request.",BizPayDetails)&'
        'IF(LateFeePercent=0,"","  Late payments may carry a fee of "&'
        'TEXT(LateFeePercent,"0%")&".")',
        fmts.cell("wrap", "plain"))
    s.row_height(row, 40)
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0, inv("notes"), fmts.cell("wrap", "plain"))
    s.row_height(row, 32)
    row += 2

    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(
        row, 0,
        '="Thank you for your business. Any questions, call "&'
        'BizPhone&" or email "&BizEmail&"."', fmts.chrome("note"))
    row += 1
    s.merge(row, 0, LAST_COL,
            "A spreadsheet cannot take card payments. Put a payment "
            "link in your payment details on the Settings tab and it "
            "prints here automatically.", fmts.chrome("note"))

    s.finish(freeze_row=None, print_area=f"A1:H{row + 2}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Job Sheet
# ---------------------------------------------------------------------------

def build_job(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = JOB_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Job Sheet", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 16)
    s.set_col(1, 1, 22)
    s.set_col(2, 2, 30)
    s.set_col(3, 3, 10)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 10)
    s.set_col(7, 7, 14)

    row = 4
    row = _pick(s, fmts, book, row, "Job to print", "JOB-0001",
                "JobIDs", "PrintJob",
                "Pick a job. Print it for the cleaner, or print "
                "several by changing the ID and printing again.")
    # Which checklist to print: the service booked, falling back to the
    # one set against the property.
    helper = row - 1
    ws.write_formula(
        helper, 0,
        s.reg.render('=IFERROR(IF(COUNTIF([[Checklists.template]],'
                     'INDEX([[Jobs.service]],MATCH(PrintJob,'
                     '[[Jobs.id]],0)))>0,INDEX([[Jobs.service]],'
                     'MATCH(PrintJob,[[Jobs.id]],0)),INDEX('
                     '[[Properties.checklist]],MATCH(INDEX('
                     '[[Jobs.property]],MATCH(PrintJob,[[Jobs.id]],0)),'
                     '[[Properties.id]],0))),"")'),
        fmts.chrome("note"))
    book.define_name("PrintJobChecklist",
                     f"={quote_sheet(NAME)}!$A${helper + 1}")
    row = _letterhead(s, fmts, row, LAST_COL, "JOB SHEET")

    def job(key: str) -> str:
        return s.reg.render(f'=IFERROR(INDEX([[Jobs.{key}]],'
                            f'MATCH(PrintJob,[[Jobs.id]],0)),"")')

    def prop(key: str) -> str:
        return s.reg.render(
            f'=IFERROR(INDEX([[Properties.{key}]],MATCH(INDEX('
            f'[[Jobs.property]],MATCH(PrintJob,[[Jobs.id]],0)),'
            f'[[Properties.id]],0)),"")')

    pairs = [
        ("Job number", "=PrintJob", "id", "Date", job("date"), "date"),
        ("Client", job("client_name"), "text", "Start", job("start"),
         "time"),
        ("Property", job("property_name"), "text", "Finish",
         job("finish"), "time"),
        ("Service", job("service"), "text", "Hours on site",
         job("hours_each"), "number1"),
        ("Cleaner", job("employee"), "text", "Second cleaner",
         job("employee2"), "text"),
        ("Frequency", job("frequency"), "text", "Cleaners",
         job("crew"), "int"),
    ]
    for left, lf, lstyle, right, rf, rstyle in pairs:
        s.write(row, 0, left, fmts.chrome("label_bold"))
        s.merge(row, 1, 2, None, fmts.cell(lstyle, "plain"))
        ws.write_formula(row, 1, lf, fmts.cell(lstyle, "plain"))
        s.write(row, 4, right, fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.cell(rstyle, "plain"))
        ws.write_formula(row, 5, rf, fmts.cell(rstyle, "plain"))
        row += 1
    row += 1

    # ---------------------------------------------------- getting in ----
    row = s.section(row, "Getting in", LAST_COL)
    for label, formula, height in [
            ("Address", prop("address"), 30),
            ("Access", prop("access"), 18),
            ("Keys", prop("key_status"), 18),
            ("Entry notes", prop("entry_notes"), 40),
            ("Parking", prop("parking"), 24),
            ("Pets", prop("pets"), 18)]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, None, fmts.cell("wrap", "plain"))
        ws.write_formula(row, 1, formula, fmts.cell("wrap", "plain"))
        s.row_height(row, height)
        row += 1
    row += 1

    # ------------------------------------------------- the checklist ----
    row = s.section(row, "The clean", LAST_COL)
    s.merge(row, 0, 5, None, fmts.chrome("note"))
    ws.write_formula(
        row, 0,
        '="Checklist: "&IF(PrintJobChecklist="","none set yet - build '
        'one on the Checklists tab",PrintJobChecklist)&"   -   tick '
        'each line as you finish it"', fmts.chrome("note"))
    s.merge(row, 6, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(
        row, 6,
        s.reg.render('=IF(PrintJobChecklist="","",COUNTIF('
                     '[[Checklists.template]],PrintJobChecklist)&'
                     '" tasks")'), fmts.chrome("note"))
    row += 1
    row = s.table_head(row, 0, ["Area", "Task", "What good looks like",
                                "Mins", "Done", "Issue?"],
                       [16, 22, 30, 8, 9, 10])
    template = "PrintJobChecklist"
    first_task = row
    for n in range(1, TASK_LINES + 1):
        key = f'{template}&"|{n}"'
        for col, (field, style) in enumerate([
                ("area", "text"), ("task", "text"), ("detail", "wrap"),
                ("minutes", "int")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Checklists.{field}]],'
                             f'MATCH({key},[[Checklists.key]],0)),"")'),
                fmts.cell(style, "plain"))
        s.write(row, 4, None, fmts.chrome("form_box"))
        s.merge(row, 5, 6, None, fmts.chrome("form_box"))
        s.row_height(row, 20)
        row += 1
    last_task = row - 1
    ws.conditional_format(first_task, 0, last_task, 0, {
        "type": "text", "criteria": "containing", "value": "Kitchen",
        "format": fmts.cf("info")})
    s.merge(row, 0, 2, "Planned minutes", fmts.chrome("total_label"))
    ws.write_formula(row, 3, f"=SUM(D{first_task + 1}:D{last_task + 1})",
                     fmts.chrome("total_value", num_format="0"))
    s.merge(row, 4, LAST_COL,
            "If the clean took much longer than this, say so below - "
            "it usually means the price is wrong, not the cleaner.",
            fmts.chrome("note"))
    row += 2

    # --------------------------------------------------- supplies -------
    row = s.section(row, "Supplies and equipment", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        s.reg.render('="Supplies kept on site: "&IFERROR(INDEX('
                     '[[Properties.supplies_onsite]],MATCH(INDEX('
                     '[[Jobs.property]],MATCH(PrintJob,[[Jobs.id]],0)),'
                     '[[Properties.id]],0)),"unknown")&"      Bring '
                     'anything else you need and log it on the Supply '
                     'Usage tab."'), fmts.cell("wrap", "plain"))
    s.row_height(row, 30)
    row += 2

    # ---------------------------------------------------- sign off ------
    row = s.section(row, "Sign off", LAST_COL)
    for label in ["Time in", "Time out", "Total hours"]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 2, None, fmts.chrome("form_line"))
        s.write(row, 4, "Notes", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 24)
        row += 1
    row += 1
    for label in ["Cleaner signature", "Client signature"]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 3, None, fmts.chrome("form_line"))
        s.write(row, 4, "Date", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 28)
        row += 1
    row += 1
    s.merge(row, 0, LAST_COL,
            "Anything broken, missing or not right? Write it here, "
            "photograph it, and tell the office the same day.",
            fmts.chrome("note"))
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("form_box"))
    s.row_height(row, 60)
    row += 1

    s.finish(freeze_row=None, print_area=f"A1:H{row + 1}",
             gridlines=False, fit_width=True)
    return s
````

### `cleaning_business_spreadsheet/sheets/help_tab.py`

*255 lines · 1 function*

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
             "What the colours mean, what to do when a cell looks "
             "wrong, and where every feature lives.", LAST_COL)

    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, LAST_COL, 14)

    row = 4

    # ----------------------------------------------- the colour code ---
    row = s.section(row, "The colour code", LAST_COL)
    row = s.table_head(row, 1, ["Looks like", "Means"], [30, 68])
    colours = [
        ("White cell with a blue edge",
         "Type here. These are the only cells you need to fill in."),
        ("Pale grey cell", "A formula. It is locked so you cannot "
                           "overwrite it by accident."),
        ("Green text or fill", "Good - on target, paid, in date, "
                               "profitable."),
        ("Amber fill", "Needs attention soon - due, low, under "
                       "target."),
        ("Red fill", "Act today - overdue, out of stock, expired, "
                     "losing money."),
        ("Blue fill", "For information - booked, sent, in progress."),
        ("A coloured bar inside a cell",
         "A data bar: the longer the bar, the bigger the number "
         "compared with the rest of the column."),
        ("A tiny chart",
         "A sparkline - twelve months of that number in one cell."),
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
                       [24, 36, 38])
    problems = [
        ("#N/A in a name column",
         "The ID you typed is not on the tab it points at",
         "Pick the ID from the drop-down instead of typing it"),
        ("'Not found' instead of a client",
         "The client ID has been deleted or renumbered",
         "Re-pick the client on the row"),
        ("#VALUE! in a total",
         "Text has been typed into a number cell - '2 hours' "
         "instead of 2",
         "Clear the cell and type the number only"),
        ("#REF! anywhere",
         "A row or column the formula needed has been deleted",
         "Undo the deletion, or copy the formula down from the row "
         "above"),
        ("#DIV/0! in a margin or rate",
         "Dividing by a zero - usually a price or an hours figure "
         "that is still blank",
         "Fill in the missing number; most of these are already "
         "guarded"),
        ("A price of 0",
         "No rate, no service price and no override",
         "Check the Services tab and the hourly rate on Settings"),
        ("The next visit date is empty",
         "The client has no future job in the diary, or the "
         "frequency is One-off",
         "Book the next visit on the Job Scheduler"),
        ("Hours look far too low",
         "The property has no area and no room counts",
         "Fill in the property, or type an hours override on the "
         "job"),
        ("Profit is negative on every job",
         "The cost per hour on Settings is wrong - it should be pay "
         "plus on-costs, not what you charge",
         "Check Pay rate and Employer on-cost on the Settings tab"),
        ("A formula cell will not let me type",
         "The sheet is protected so formulas survive a stray paste",
         f"Review, Unprotect Sheet, password '{cfg.password}'"),
        ("A drop-down has the wrong options",
         "The list behind it lives on the Settings & Dropdown Lists "
         "tab",
         "Edit the list there and every drop-down updates"),
        ("Numbers have not changed after an edit",
         "Calculation has been switched to manual",
         "Formulas, Calculation Options, Automatic - or press F9"),
        ("New rows have no formulas",
         "Rows added below the block do not inherit anything",
         "Copy the last row of the block down into the new rows"),
        ("The dashboard shows nothing",
         "Jobs are still at Booked, not Completed",
         "Revenue counts Completed and Invoiced work only"),
    ]
    for what, why, fix in problems:
        s.write(row, 1, what, fmts.chrome("label_bold"))
        s.merge(row, 2, 5, why, fmts.chrome("label"))
        s.merge(row, 6, LAST_COL, fix, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ----------------------------------------------- the glossary ------
    row = s.section(row, "The words this workbook uses", LAST_COL)
    row = s.table_head(row, 1, ["Term", "What it means here"],
                       [26, 72])
    glossary = [
        ("Cost per hour", "What one cleaner-hour really costs you: "
                          "pay plus your on-costs. Not what you "
                          "charge."),
        ("Margin", "Profit as a percentage of the price. 40% margin "
                   "on a $100 clean means $40 left after wages, "
                   "supplies and travel."),
        ("Margin floor", "The lowest margin you will accept. Jobs "
                         "under it are flagged red all over the "
                         "workbook."),
        ("Contribution", "Price less the costs that only happen "
                         "because you did the job. What is left to "
                         "pay the fixed bills."),
        ("Break-even", "The number of jobs a month that covers your "
                       "fixed costs. Below it you are paying to "
                       "work."),
        ("Utilisation", "Billed hours as a share of paid hours. "
                        "Travel, waiting and admin are the gap."),
        ("Frequency factor", "The discount a regular client gets. "
                             "Weekly work takes less time per visit "
                             "because less dirt builds up."),
        ("Condition factor", "How much longer a dirty property takes. "
                             "1.0 is normal, 1.8 is a hoarder clean."),
        ("Days of cover", "How many days of stock you have at your "
                          "recent rate of use."),
        ("Ageing", "How long an invoice has been overdue, in "
                   "buckets. Anything past 60 days rarely gets paid "
                   "in full."),
        ("Recurring revenue", "What your regular clients are worth "
                              "over a year at today's frequencies. "
                              "The number a buyer would pay for."),
        ("Churn", "Clients who have stopped booking. Every cleaning "
                  "business has some; the question is how fast you "
                  "replace them."),
    ]
    for term, meaning in glossary:
        s.write(row, 1, term, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, meaning, fmts.chrome("label"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ----------------------------------------------- feature map -------
    row = s.section(row, "Where every feature lives", LAST_COL)
    row = s.table_head(row, 1, ["If you want to...", "Go to"],
                       [50, 48])
    where = [
        ("Price a clean before you quote",
         "Pricing & Quote Calculator"),
        ("See what a job will really cost you",
         "Pricing & Quote Calculator, section 5"),
        ("Compare weekly, fortnightly and monthly over a year",
         "Pricing & Quote Calculator, section 8"),
        ("Keep alarm codes, key numbers and parking notes",
         "Property & Site Database"),
        ("Know which clients are drifting away", "Client Retention"),
        ("Find out which advertising actually works",
         "Reports & Analytics"),
        ("See who is free on Thursday",
         "Employee Schedule & Attendance"),
        ("Check nobody is double booked",
         "Job Scheduler, Check column"),
        ("Build a room-by-room task list", "Cleaning Checklists"),
        ("Give the cleaner a sheet to take with them",
         "Print Job Sheet"),
        ("Work out what a job actually earned",
         "Job Costing & Profitability"),
        ("Chase an unpaid invoice", "Invoices, Action column"),
        ("Know what to reorder", "Supplies Inventory"),
        ("Remember to service the carpet machine",
         "Equipment & Assets"),
        ("Claim your mileage", "Mileage & Travel"),
        ("Keep insurance and background checks in date",
         "Safety & Compliance"),
        ("Work out how many jobs cover the bills", "Financial Plan"),
        ("Decide whether to hire",
         "Financial Plan, section 3"),
        ("Put your prices up", "Business Settings, then re-read "
                               "Job Costing"),
        ("Change a drop-down list", "Settings & Dropdown Lists"),
    ]
    for want, go in where:
        s.merge(row, 1, 5, want, fmts.chrome("label"))
        s.merge(row, 6, LAST_COL, go, fmts.chrome("label_bold"))
        s.row_height(row, 20)
        row += 1
    row += 1

    # ----------------------------------------------- questions ---------
    row = s.section(row, "Questions people ask", LAST_COL)
    row = s.table_head(row, 1, ["Question", "Answer"], [30, 68])
    faqs = [
        ("Can two people use it at once?",
         "Not safely in Excel on a shared drive. Upload it to Google "
         "Sheets or OneDrive if more than one person needs it open."),
        ("Will it work on a phone?",
         "It opens, and the dashboard reads well enough. Typing is "
         "painful on anything this wide - use the printed job "
         "sheet."),
        ("How many clients can it hold?",
         "The blank workbook is sized for a few hundred clients and "
         "several thousand jobs. Build it with more rows if you need "
         "them - the generator takes a row count."),
        ("Can I change the colours?",
         "Yes. They are ordinary cell formats - select and restyle."),
        ("Can I delete tabs I do not use?",
         "You can, but other tabs may reference them. Hiding is "
         "safer: right-click the tab, Hide."),
        ("Does it handle more than one currency?",
         "One at a time. The symbol is on the Settings tab."),
        ("Is the sheet password secure?",
         "No. Sheet protection stops accidents, not people. Protect "
         "the file itself if the data is sensitive - and it is."),
        ("Can I use it for a window cleaning or gardening round?",
         "Yes, with almost no changes. Rename the services and the "
         "checklist areas and the rest follows."),
    ]
    for question, answer in faqs:
        s.write(row, 1, question, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, answer, fmts.chrome("label"))
        s.row_height(row, 26)
        row += 1
    row += 1

    s.note(row,
           f"{cfg.product_name}, version {cfg.version}. Built by "
           f"{cfg.author} with formulas only - no macros, no add-ins, "
           f"nothing to install and nothing that phones home. Every "
           f"calculation on every tab can be read, checked and "
           f"changed by you.",
           LAST_COL, first_col=1, height=34)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

---

## 6. Quality checks (optional)

Not needed to build the workbook - these are the checks that keep it honest. The first two need `openpyxl`, the third needs `formulas`.

### `tools/validate_cleaning_workbook.py`

*537 lines · 9 functions*

> Static checks on a generated Cleaning Business workbook.
>
> Run it with any Python that has openpyxl installed::
>
>     python tools/validate_cleaning_workbook.py output/crm.xlsx
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
| `_specs()` | Every register specification, keyed by sheet name |
| `check_chain(wb, problems, notes)` | The quote-to-job-to-invoice chain only holds while these do |
| `check(path)` |  |

**Module constants:** `FUNCTION_RE` (re.compile(...)), `NAME_RE` (re.compile(...)), `BARE_COL_RE` (re.compile(...)), `KNOWN_FUNCTIONS` ({'IF', 'IFERROR', 'SUM', 'SUMIF', 'SUMIFS', 'COUNT', 'COU...), `MODERN_ONLY` ({'XLOOKUP', 'IFS', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCA...), `PRECEDENCE_OK` ({'F#+AF#/#', '(K#-J#)*#-IF(L#="",#,L#)/#', 'K#/L#+IF(P#="...), `SEEDED_INPUTS` (2 entries), `CHAIN_CHECKS` (35 items).

````python
"""Static checks on a generated Cleaning Business workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_cleaning_workbook.py output/crm.xlsx

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
from cleaning_business_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402

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
    # finish time = start + hours/24 (a whole day is 1.0 in Excel)
    "F#+AF#/#",
    # hours worked = (finish-start)*24 less the break in minutes
    '(K#-J#)*#-IF(L#="",#,L#)/#',
    # cost a year = purchase/life + repairs spread over its age
    'K#/L#+IF(P#="",#,P#)/MAX(#,T#)',
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
    "Calendar!B5": "the month picker starts on this month",
    "Staff Schedule!B5": "the rota picker starts on this week",
}

#: the automation chain, link by link: (sheet, column key, what the
#: formula in that column has to still be pointing at)
CHAIN_CHECKS = [
    ("Jobs", "est_hours", ("sheet", "Services")),
    ("Jobs", "hours_each", ("column", "Jobs", "est_hours")),
    ("Jobs", "price", ("column", "Jobs", "base_price")),
    ("Jobs", "actual_hours", ("sheet", "Timesheets")),
    ("Jobs", "invoice", ("sheet", "Invoice Lines")),
    ("Jobs", "tasks_done", ("sheet", "Job Tasks")),
    ("Jobs", "checklist_pct", ("column", "Jobs", "tasks_done")),
    ("Job Costing", "revenue", ("sheet", "Jobs")),
    ("Job Costing", "labour", ("column", "Job Costing", "hours_used")),
    ("Job Costing", "supply_logged", ("sheet", "Supply Usage")),
    ("Job Costing", "supplies",
     ("column", "Job Costing", "supply_logged")),
    ("Job Costing", "profit", ("column", "Job Costing", "cost")),
    ("Quotes", "subtotal", ("sheet", "Quote Lines")),
    ("Quotes", "total", ("column", "Quotes", "net")),
    ("Quote Lines", "gross",
     ("column", "Quote Lines", "price_used")),
    ("Quote Lines", "total", ("column", "Quote Lines", "gross")),
    ("Invoices", "subtotal", ("sheet", "Invoice Lines")),
    ("Invoices", "paid", ("sheet", "Payments")),
    ("Invoices", "outstanding", ("column", "Invoices", "total")),
    ("Invoices", "days_late", ("column", "Invoices", "due_date")),
    ("Invoice Lines", "price_used", ("column", "Invoice Lines",
                                     "job_price")),
    ("Payments", "invoice_total", ("sheet", "Invoices")),
    ("Payments", "balance",
     ("column", "Payments", "invoice_total")),
    ("Clients", "invoiced", ("sheet", "Invoices")),
    ("Clients", "profit", ("sheet", "Job Costing")),
    ("Clients", "next_job", ("sheet", "Jobs")),
    ("Clients", "retention", ("column", "Clients", "days_since")),
    ("Properties", "est_hours", ("column", "Properties", "size_hours")),
    ("Properties", "visit_hours",
     ("column", "Properties", "est_hours")),
    ("Properties", "suggested",
     ("column", "Properties", "est_hours")),
    ("Employees", "utilisation", ("column", "Employees", "paid_hours")),
    ("Timesheets", "cost", ("column", "Timesheets", "rate")),
    ("Supplies", "days_cover", ("column", "Supplies", "daily_use")),
    ("Supply Usage", "cost", ("column", "Supply Usage", "unit_cost")),
    ("Services", "hourly_yield", ("column", "Services", "profit")),
]


def _specs():
    """Every register specification, keyed by sheet name."""
    from cleaning_business_spreadsheet.build import REGISTER_MODULES
    from cleaning_business_spreadsheet.config import Config
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
              else "output/Cleaning Business CRM & Pricing "
                   "Calculator.xlsx")
    raise SystemExit(check(target))
````

### `tools/check_cleaning_sample.py`

*309 lines · 5 functions*

> Cross-check the example records in a generated cleaning workbook.
>
> The demo data has to join up, otherwise the dashboard opens full of
> zeros and every lookup column fills with "Not found". This reads the
> literal values out of the finished file and makes sure every job
> points at a real client and a real property, every invoice line
> points at a real job, every payment at a real invoice, every
> timesheet at a real employee - and that there is enough of all of it
> to make the dashboard look like a business.
>
>     python tools/check_cleaning_sample.py output/test.xlsx

| Defines | What it does |
|---------|--------------|
| `data_rows()` | Last data row of every register, straight from its specification |
| `headers(ws, header_row)` |  |
| `column_values(ws, header, header_row)` | Every non-empty literal value under the given header |
| `check(path)` |  |
| `main(argv)` |  |

**Module constants:** `HEADER_ROW` (4), `LAST_ROW` (0 entries).

````python
"""Cross-check the example records in a generated cleaning workbook.

The demo data has to join up, otherwise the dashboard opens full of
zeros and every lookup column fills with "Not found". This reads the
literal values out of the finished file and makes sure every job
points at a real client and a real property, every invoice line
points at a real job, every payment at a real invoice, every
timesheet at a real employee - and that there is enough of all of it
to make the dashboard look like a business.

    python tools/check_cleaning_sample.py output/test.xlsx
"""

from __future__ import annotations

import os
import sys
from collections import Counter

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

HEADER_ROW = 4


def data_rows() -> dict:
    """Last data row of every register, straight from its specification.

    Without this the checks read the summary strip underneath the grid
    and report labels like 'Total this month' as a missing client.
    """
    from cleaning_business_spreadsheet.build import REGISTER_MODULES
    from cleaning_business_spreadsheet.config import Config
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
        # headers carry the unit in brackets - 'Distance (miles)' only
        # matches a build in miles, so fall back to a prefix match
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

    clients = sheet("Clients")
    properties = sheet("Properties")
    jobs = sheet("Jobs")
    services = sheet("Services")
    employees = sheet("Employees")
    if None in (clients, properties, jobs, services, employees):
        for problem in problems:
            print(f"    - {problem}")
        return 1

    client_names = column_values(clients, "Client")
    property_names = column_values(properties, "Property")
    service_names = column_values(services, "Service")
    staff_names = column_values(employees, "Name")
    notes.append(f"{len(client_names)} clients, {len(property_names)} "
                 f"properties, {len(service_names)} services, "
                 f"{len(staff_names)} staff")
    for label, values in (("client", client_names),
                          ("property", property_names),
                          ("service", service_names),
                          ("employee", staff_names)):
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

    # the spine: client -> property -> job -> invoice -> payment
    id_link("Properties", "Client ID", "Clients", "Client")
    id_link("Jobs", "Client ID", "Clients", "Client")
    id_link("Jobs", "Property ID", "Properties", "Property")
    id_link("Invoices", "Client ID", "Clients", "Client")
    id_link("Invoice Lines", "Invoice ID", "Invoices", "Invoice date")
    id_link("Invoice Lines", "Job ID", "Jobs", "Date")
    id_link("Payments", "Invoice ID", "Invoices", "Invoice date")
    id_link("Timesheets", "Job ID", "Jobs", "Date")
    id_link("Quote Lines", "Quote ID", "Quotes", "Quote date")
    id_link("Supply Usage", "Job ID", "Jobs", "Date")
    id_link("Complaints", "Client ID", "Clients", "Client")

    link("Jobs", "Service", "Services", "Service")
    link("Jobs", "Cleaner", "Employees", "Name")
    link("Properties", "Checklist", "Services", "Service")
    link("Checklists", "Checklist for", "Services", "Service")
    link("Timesheets", "Cleaner", "Employees", "Name")
    link("Supply Usage", "Item", "Supplies", "Item")
    link("Mileage", "Driver", "Employees", "Name")
    link("Quote Lines", "Service", "Services", "Service")
    link("Complaints", "Cleaner on the job", "Employees", "Name")

    # every quote needs lines, or the totals are all zero
    quote_ids = set(str(v) for v in column_values(sheet("Quote Lines"),
                                                  "Quote ID"))
    quote_count = len(column_values(sheet("Quotes"), "Quote date"))
    if len(quote_ids) < quote_count:
        problems.append(
            f"only {len(quote_ids)} of {quote_count} quotes have lines - "
            f"the rest total zero")
    else:
        notes.append(f"{len(column_values(sheet('Quote Lines'), 'Service'))}"
                     f" quote lines across {len(quote_ids)} quotes")

    # every invoice needs lines too
    invoiced = set(str(v) for v in column_values(sheet("Invoice Lines"),
                                                 "Invoice ID"))
    invoice_count = len(column_values(sheet("Invoices"),
                                  "Invoice date"))
    if len(invoiced) < invoice_count:
        problems.append(
            f"only {len(invoiced)} of {invoice_count} invoices have "
            f"lines - the rest total zero")
    else:
        notes.append(f"{invoice_count} invoices, all with lines")

    # the diary: enough jobs, spread over time, mostly with a cleaner
    dates = [v for v in column_values(jobs, "Date") if hasattr(v, "year")]
    if len(dates) < 40:
        problems.append(f"only {len(dates)} example jobs - the calendar "
                        f"and the dashboard will look empty")
    else:
        span = (max(dates) - min(dates)).days + 1
        notes.append(f"{len(dates)} jobs over {span} days "
                     f"({min(dates):%d %b} to {max(dates):%d %b})")
    statuses = Counter(column_values(jobs, "Status"))
    done = statuses.get("Completed", 0) + statuses.get("Invoiced", 0)
    if done < 20:
        problems.append(f"only {done} completed jobs - every revenue "
                        f"figure in the workbook reads zero")
    else:
        notes.append("jobs by status: " + ", ".join(
            f"{k} {v}" for k, v in statuses.most_common(5)))
    future = [day for day in dates if day >= min(dates)]
    assigned = len(column_values(jobs, "Cleaner"))
    if assigned < len(future) * 0.8:
        problems.append(f"only {assigned} of {len(dates)} jobs have a "
                        f"cleaner against them")
    else:
        notes.append(f"{assigned} jobs have a cleaner")

    # prices: a service with no price silently costs nothing.
    # 'Standard price' is a formula, so read the typed-in rate instead
    priced = [v for v in column_values(services, "Price per unit")
              if isinstance(v, (int, float)) and v > 0]
    if len(priced) < len(service_names) * 0.9:
        problems.append(f"only {len(priced)} of {len(service_names)} "
                        f"services have a price")
    else:
        notes.append(f"{len(priced)} services priced")

    # the rest of the dashboard tiles
    for name, header, least, label in (
            ("Payments", "Amount", 6, "payments"),
            ("Expenses", "Amount (net)", 25, "expenses"),
            ("Supplies", "Counted stock", 10, "stock counts"),
            ("Supply Usage", "Quantity used", 10, "supply usages"),
            ("Equipment", "Cost", 6, "pieces of equipment"),
            ("Mileage", "Distance", 10, "mileage trips"),
            ("Timesheets", "Break (min)", 20, "timesheet rows"),
            ("Job Tasks", "Minutes taken", 15, "checklist ticks"),
            ("Checklists", "Minutes", 30, "checklist tasks")):
        ws = sheet(name)
        if ws is None:
            continue
        try:
            values = [v for v in column_values(ws, header)
                      if isinstance(v, (int, float))]
        except KeyError as error:
            problems.append(str(error))
            continue
        if len(values) < least:
            problems.append(f"only {len(values)} {label} - expected at "
                            f"least {least}")
        else:
            notes.append(f"{len(values)} {label}")

    # leads and complaints give the pipeline and the quality tabs life
    for name, header, least, label in (
            ("Leads", "Name / company", 6, "leads"),
            ("Quotes", "Quote date", 5, "quotes"),
            ("Complaints", "What they said", 3, "complaints"),
            ("Incidents", "What happened, in full", 2, "incidents"),
            ("Safety", "Document", 8, "compliance documents")):
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
    targets = argv[1:] or ["output/Cleaning Business CRM & Pricing "
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

*1,076 lines · 10 functions*

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

**Module constants:** `ROOT` (os.path.dirname(...)), `SHARED_TOOLS` (2 items), `PET_SECTIONS` (5 items), `PET` (9 entries), `SOCIAL_SECTIONS` (7 items), `SOCIAL` (9 entries), `MEAL_SECTIONS` (6 items), `MEAL` (11 entries), `CLEANING_SECTIONS` (6 items), `CLEANING` (11 entries), `PRODUCTS` (4 entries), `ARCHITECTURE` ('Everything here is plain Python writing a `.xlsx` file t...), `DEFAULT_CLOSING` ('**Hide the `Calc` engine**, activate the first tab and c...), `CALC_ENGINE` ('### The hidden `Calc` tab\n\nDashboards do not aggregate...) and 1 more.

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

PRODUCTS = {"pet": PET, "social": SOCIAL, "meal": MEAL,
            "cleaning": CLEANING}


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

#### Clients

Every customer, what they have spent, what they earn you and when you last saw them.

45 columns, 200 rows of capacity, 21 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Client | `name` | text | you type it |
| C | Type | `type` | text | drop down `ClientTypeList` |
| D | Main contact | `contact` | text | you type it |
| E | Phone | `phone` | text | you type it |
| F | Email | `email` | text | you type it |
| G | Billing address | `address` | wrap | you type it |
| H | Status | `status` | status | drop down `ClientStatusList` |
| I | Came from | `source` | text | drop down `LeadSourceList` |
| J | Client since | `since` | date | you type it (checked) |
| K | Usual frequency | `frequency` | text | drop down `FrequencyList` |
| L | Preferred day | `pref_day` | text | drop down `WeekdayList` |
| M | Preferred time | `pref_time` | text | drop down `TimeSlotList` |
| N | Preferred cleaner | `pref_cleaner` | text | drop down `EmployeeNames` |
| O | Agreed hourly rate | `rate` | money | you type it (checked) |
| P | Payment terms (days) | `terms` | int | you type it (checked) |
| Q | Pays by | `payment_method` | text | drop down `PaymentMethodList` |
| R | Contract | `contract` | status | drop down `ContractTermList` |
| S | Contract ends | `contract_end` | date | you type it (checked) |
| T | VIP? | `vip` | flag | drop down `YesNoList` |
| U | Happy to be contacted? | `marketing` | flag | drop down `YesNoList` |
| V | Notes | `notes` | wrap | you type it |
| W | Properties | `properties` | int | formula |
| X | Jobs booked | `jobs_booked` | int | formula |
| Y | Jobs done | `jobs_done` | int | formula |
| Z | Cancelled | `cancelled` | int | formula |
| AA | Last clean | `last_job` | date | formula |
| AB | Days since | `days_since` | int | formula |
| AC | Next booking | `next_job` | date | formula |
| AD | Invoiced to date | `invoiced` | money | formula |
| AE | Paid to date | `received` | money | formula |
| AF | Owing now | `outstanding` | money | formula |
| AG | Average job | `avg_job` | money | formula |
| AH | Profit to date | `profit` | money | formula |
| AI | Margin | `margin` | percent0 | formula |
| AJ | Value a year | `year_value` | money | formula |
| AK | Average rating | `rating` | number1 | formula |
| AL | Complaints | `complaints` | int | formula |
| AM | Rank | `invoiced_rank` | int | formula, hidden |
| AN | Rank | `profit_rank` | int | formula, hidden |
| AO | Rank | `year_value_rank` | int | formula, hidden |
| AP | Standing | `retention` | status | formula |
| AQ | Risk score | `risk` | number1 | formula, hidden |
| AR | Risk order | `risk_rank` | int | formula, hidden |
| AS | Action | `alert` | status | formula |

#### Properties

The places you clean, how long they take and what they should be worth.

34 columns, 250 rows of capacity, 20 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Property | `name` | text | you type it |
| C | Client ID | `client` | id | drop down `ClientIDs` |
| D | Client | `client_name` | text | formula |
| E | Property type | `type` | text | drop down `PropertyTypeList` |
| F | Address | `address` | wrap | you type it |
| G | Area (sq ft) | `area` | number | you type it (checked) |
| H | Bedrooms | `beds` | int | you type it (checked) |
| I | Bathrooms | `baths` | int | you type it (checked) |
| J | Other rooms | `rooms` | int | you type it (checked) |
| K | Floor types | `floors` | text | you type it |
| L | Condition | `condition` | text | drop down `ConditionList` |
| M | Pets | `pets` | text | you type it |
| N | How we get in | `access` | text | drop down `AccessList` |
| O | Key | `key_status` | status | drop down `KeyStatusList` |
| P | Entry instructions | `entry_notes` | wrap | you type it |
| Q | Parking | `parking` | text | drop down `ParkingList` |
| R | Distance (miles) | `distance` | number1 | you type it (checked) |
| S | Checklist | `checklist` | text | drop down `ServiceNames` |
| T | Client supplies? | `supplies_onsite` | flag | drop down `YesNoList` |
| U | Active? | `active` | flag | drop down `YesNoList` |
| V | Notes | `notes` | wrap | you type it |
| W | Size hours | `size_hours` | number1 | formula, hidden |
| X | Estimated hours | `est_hours` | number1 | formula |
| Y | Hours on site | `visit_hours` | number1 | formula |
| Z | Suggested price | `suggested` | money | formula |
| AA | Labour cost | `labour_cost` | money | formula, hidden |
| AB | Estimated margin | `est_margin` | percent0 | formula |
| AC | Jobs here | `jobs` | int | formula |
| AD | Last visit | `last_visit` | date | formula |
| AE | Average actual hours | `actual_hours` | number1 | formula |
| AF | Estimate vs actual | `variance` | number1 | formula |
| AG | Revenue here | `revenue` | money | formula |
| AH | Action | `alert` | status | formula |

#### Leads

Every enquiry, how long it has been waiting and who needs a phone call today.

26 columns, 150 rows of capacity, 19 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Enquiry date | `date` | date | you type it (checked) |
| C | Name / company | `name` | text | you type it |
| D | Phone | `phone` | text | you type it |
| E | Email | `email` | text | you type it |
| F | Address or area | `address` | wrap | you type it |
| G | Client type | `type` | text | drop down `ClientTypeList` |
| H | Came from | `source` | text | drop down `LeadSourceList` |
| I | Interested in | `service` | text | drop down `ServiceNames` |
| J | Frequency wanted | `frequency` | text | drop down `FrequencyList` |
| K | Stage | `stage` | status | drop down `LeadStageList` |
| L | Estimated value a year | `value` | money | you type it (checked) |
| M | Quote sent | `quote` | id | drop down `QuoteIDs` |
| N | Visit booked | `visit_date` | date | you type it (checked) |
| O | Last contacted | `contacted` | date | you type it (checked) |
| P | Next action | `next_action` | text | you type it |
| Q | Next action due | `next_date` | date | you type it (checked) |
| R | Became client | `client` | id | drop down `ClientIDs` |
| S | If lost, why | `lost_reason` | text | drop down `LostReasonList` |
| T | Notes | `notes` | wrap | you type it |
| U | Days open | `days_open` | int | formula |
| V | Days since anything happened | `days_quiet` | int | formula, hidden |
| W | Quote total | `quote_value` | money | formula |
| X | Quote status | `quote_status` | status | formula |
| Y | Won? | `won` | flag | formula, hidden |
| Z | Action | `alert` | status | formula |

#### Services

Your price list, with the margin on every service worked out for you.

29 columns, 60 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Service | `name` | text | you type it |
| C | Category | `category` | text | drop down `ServiceCategoryList` |
| D | Priced by | `model` | text | drop down `PricingModelList` |
| E | Unit | `unit` | text | you type it |
| F | Price per unit | `price` | money | you type it (checked) |
| G | Hours (one cleaner) | `hours` | number1 | you type it (checked) |
| H | Uplift | `uplift` | percent0 | you type it (checked) |
| I | Supplies cost | `supplies` | money | you type it (checked) |
| J | Minimum charge | `minimum` | money | you type it (checked) |
| K | Popular? | `popular` | flag | drop down `YesNoList` |
| L | Offered? | `active` | flag | drop down `YesNoList` |
| M | Who can do it | `skills` | text | you type it |
| N | Equipment needed | `equipment` | text | you type it |
| O | What it includes | `description` | wrap | you type it |
| P | Standard price | `standard_price` | money | formula |
| Q | Labour cost | `labour_cost` | money | formula |
| R | Supply cost | `supply_cost` | money | formula, hidden |
| S | Cost to deliver | `total_cost` | money | formula |
| T | Profit | `profit` | money | formula |
| U | Margin | `margin` | percent0 | formula |
| V | Earns an hour | `hourly_yield` | money | formula |
| W | Price for target margin | `target_price` | money | formula |
| X | Times quoted | `times_quoted` | int | formula |
| Y | Times booked | `times_booked` | int | formula |
| Z | Revenue | `revenue` | money | formula |
| AA | Rank | `revenue_rank` | int | formula, hidden |
| AB | Rank | `hourly_yield_rank` | int | formula, hidden |
| AC | Check | `alert` | status | formula |

#### Quotes

Every quote you have sent, what it is worth, what it makes you and when to chase it.

37 columns, 200 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Quote date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula |
| D | Client ID | `client` | id | drop down `ClientIDs` |
| E | Client | `client_name` | text | formula |
| F | Prospect name | `prospect` | text | you type it |
| G | Property ID | `property` | id | drop down `PropertyIDs` |
| H | Address | `address` | wrap | you type it |
| I | Main service | `service` | text | drop down `ServiceNames` |
| J | Frequency | `frequency` | text | drop down `FrequencyList` |
| K | Status | `status` | status | drop down `QuoteStatusList` |
| L | Sent on | `sent_date` | date | you type it (checked) |
| M | Last followed up | `followed_up` | date | you type it (checked) |
| N | Accepted / declined on | `decision_date` | date | you type it (checked) |
| O | Discount | `discount` | percent0 | you type it (checked) |
| P | Deposit asked for | `deposit_input` | money | you type it (checked) |
| Q | Valid until | `valid_input` | date | you type it (checked) |
| R | Job booked | `job` | id | drop down `JobIDs` |
| S | If declined, why | `decline_reason` | text | drop down `LostReasonList` |
| T | Notes | `notes` | wrap | you type it |
| U | Terms shown | `terms` | wrap | you type it |
| V | Lines | `lines` | int | formula, hidden |
| W | Hours quoted | `hours` | number1 | formula |
| X | Subtotal | `subtotal` | money | formula |
| Y | Less discount | `discount_value` | money | formula |
| Z | Net | `net` | money | formula |
| AA | Tax | `tax` | money | formula |
| AB | Total | `total` | money | formula |
| AC | Deposit | `deposit` | money | formula |
| AD | Cost to deliver | `cost` | money | formula, hidden |
| AE | Profit | `profit` | money | formula |
| AF | Margin | `margin` | percent0 | formula |
| AG | Valid to | `valid_until` | date | formula |
| AH | Days waiting | `days_waiting` | int | formula |
| AI | Expired? | `expired` | flag | formula, hidden |
| AJ | Won? | `won` | flag | formula, hidden |
| AK | Action | `alert` | status | formula |

#### Quote Lines

The detail behind every quote - one row per thing you are charging for.

21 columns, 600 rows of capacity, 8 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Quote ID | `quote` | id | drop down `QuoteIDs` |
| C | Service | `service` | text | drop down `ServiceNames` |
| D | Description on the quote | `description` | wrap | you type it |
| E | Quantity | `qty` | number1 | you type it (checked) |
| F | Price each (override) | `price_input` | money | you type it (checked) |
| G | Discount | `discount` | percent0 | you type it (checked) |
| H | Optional extra? | `optional` | flag | drop down `YesNoList` |
| I | Line notes | `notes` | wrap | you type it |
| J | Catalogue price | `catalogue_price` | money | formula, hidden |
| K | Per | `unit` | text | formula |
| L | Price each | `price_used` | money | formula |
| M | Wording used | `wording` | wrap | formula, hidden |
| N | Hours | `hours` | number1 | formula |
| O | Line value | `gross` | money | formula, hidden |
| P | Line total | `total` | money | formula |
| Q | Cost | `cost` | money | formula, hidden |
| R | Margin | `margin` | percent0 | formula |
| S | Line number | `seq` | int | formula, hidden |
| T | Lookup key | `key` | text | formula, hidden |
| U | Check | `alert` | status | formula |

#### Jobs

Every visit you have booked, priced automatically and checked for clashes.

50 columns, 600 rows of capacity, 21 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Day | `weekday` | text | formula |
| D | Week | `week` | int | formula |
| E | Month | `month` | text | formula |
| F | Start | `start` | time | you type it (checked) |
| G | Client ID | `client` | id | drop down `ClientIDs` |
| H | Client | `client_name` | text | formula |
| I | Client type | `client_type` | text | formula, hidden |
| J | Property ID | `property` | id | drop down `PropertyIDs` |
| K | Property | `property_name` | text | formula |
| L | Address | `address` | text | formula |
| M | Service | `service` | text | drop down `ServiceNames` |
| N | Job type | `type` | text | drop down `JobTypeList` |
| O | Frequency | `frequency` | text | drop down `FrequencyList` |
| P | Status | `status` | status | drop down `JobStatusList` |
| Q | Priority | `priority` | text | drop down `PriorityList` |
| R | Cleaner | `employee` | text | drop down `EmployeeNames` |
| S | Second cleaner | `employee2` | text | drop down `EmployeeNames` |
| T | Cleaners | `cleaners` | int | you type it (checked) |
| U | Planned hours (override) | `hours_input` | number1 | you type it (checked) |
| V | Price (override) | `price_input` | money | you type it (checked) |
| W | Travel (miles, round trip) | `travel` | number1 | you type it (checked) |
| X | Recurring? | `recurring` | flag | drop down `YesNoList` |
| Y | Confirmed with client? | `confirmed` | flag | drop down `YesNoList` |
| Z | Actual hours (if no timesheet) | `actual_input` | number1 | you type it (checked) |
| AA | Client rating | `rating` | rating | you type it (checked) |
| AB | Feedback | `feedback` | wrap | you type it |
| AC | Job notes | `notes` | wrap | you type it |
| AD | Hours (total) | `est_hours` | number1 | formula |
| AE | Crew used | `crew` | int | formula, hidden |
| AF | Hours on site | `hours_each` | number1 | formula |
| AG | Finish | `finish` | time | formula |
| AH | Rate used | `rate_used` | money | formula, hidden |
| AI | Base price | `base_price` | money | formula, hidden |
| AJ | Price | `price` | money | formula |
| AK | Actual hours | `actual_hours` | number1 | formula |
| AL | Over / under | `over_run` | number1 | formula |
| AM | Checklist | `tasks` | int | formula, hidden |
| AN | Tasks done | `tasks_done` | int | formula, hidden |
| AO | Checklist complete | `checklist_pct` | percent0 | formula |
| AP | Invoice | `invoice` | id | formula |
| AQ | Next visit due | `next_due` | date | formula |
| AR | Diary | `clash` | status | formula |
| AS | Upcoming date | `future` | date | formula, hidden |
| AT | Diary order | `future_rank` | int | formula, hidden |
| AU | Needs a cleaner | `unassigned` | date | formula, hidden |
| AV | Assign order | `unassigned_rank` | int | formula, hidden |
| AW | Days away | `days_away` | int | formula, hidden |
| AX | Action | `alert` | status | formula |

#### Job Tasks

The checklist as it was really done - ticked off, timed and with the problems written down.

20 columns, 900 rows of capacity, 11 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Job ID | `job` | id | drop down `JobIDs` |
| C | Date | `date` | date | formula |
| D | Client | `client_name` | text | formula |
| E | Property | `property_name` | text | formula |
| F | Area | `area` | text | drop down `TaskAreaList` |
| G | Task | `task` | text | you type it |
| H | Status | `status` | status | drop down `TaskStatusList` |
| I | Done by | `employee` | text | drop down `EmployeeNames` |
| J | Minutes taken | `minutes` | int | you type it (checked) |
| K | Quality | `quality` | status | drop down `QualityList` |
| L | Issue found | `issue` | wrap | you type it |
| M | Told the client? | `reported` | flag | drop down `YesNoList` |
| N | Photo link | `photo` | link | you type it |
| O | Notes | `notes` | wrap | you type it |
| P | Required? | `required` | flag | formula, hidden |
| Q | Planned minutes | `planned` | int | formula, hidden |
| R | Over / under (min) | `over` | int | formula, hidden |
| S | Complete? | `done` | flag | formula, hidden |
| T | Check | `alert` | status | formula |

#### Checklists

The standard for every service, room by room, with honest minutes against each task.

24 columns, 400 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Checklist for | `template` | text | drop down `ServiceNames` |
| C | Only at this property | `property` | id | drop down `PropertyIDs` |
| D | Area | `area` | text | drop down `TaskAreaList` |
| E | Order | `order` | int | you type it (checked) |
| F | Task | `task` | text | you type it |
| G | How it should look | `detail` | wrap | you type it |
| H | Minutes | `minutes` | int | you type it (checked) |
| I | How often | `frequency` | text | drop down `FrequencyList` |
| J | Required? | `required` | flag | drop down `YesNoList` |
| K | Photo needed? | `photo` | flag | drop down `YesNoList` |
| L | Products / kit | `products` | text | you type it |
| M | In use? | `active` | flag | drop down `YesNoList` |
| N | Notes | `notes` | wrap | you type it |
| O | Hours | `hours` | number1 | formula, hidden |
| P | Total for this checklist | `service_total` | number1 | formula |
| Q | Service allows | `service_hours` | number1 | formula, hidden |
| R | Times ticked off | `times_done` | int | formula |
| S | Times skipped | `times_skipped` | int | formula |
| T | Issues found | `issues` | int | formula |
| U | Standard | `status` | status | formula, hidden |
| V | Line number | `seq` | int | formula, hidden |
| W | Lookup key | `key` | text | formula, hidden |
| X | Check | `alert` | status | formula |

#### Employees

Your team, what they cost, what they bring in and how full their week is.

39 columns, 40 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Name | `name` | text | you type it |
| C | Role | `role` | text | drop down `EmployeeRoleList` |
| D | Employment | `type` | text | drop down `EmploymentTypeList` |
| E | Status | `status` | status | drop down `EmployeeStatusList` |
| F | Phone | `phone` | text | you type it |
| G | Email | `email` | text | you type it |
| H | Emergency contact | `emergency` | text | you type it |
| I | Started | `start` | date | you type it (checked) |
| J | Pay an hour | `pay` | money | you type it (checked) |
| K | Contracted hours a week | `hours` | number1 | you type it (checked) |
| L | Trained on | `skills` | text | you type it |
| M | Areas covered | `areas` | text | you type it |
| N | Vehicle | `vehicle` | text | drop down `VehicleList` |
| O | Driver? | `licence` | flag | drop down `YesNoList` |
| P | Background check? | `checked` | flag | drop down `YesNoNAList` |
| Q | Own insurance? | `insured` | flag | drop down `YesNoNAList` |
| R | Uniform issued | `uniform` | text | you type it |
| S | Notes | `notes` | wrap | you type it |
| T | Pay used | `pay_rate` | money | formula, hidden |
| U | True cost an hour | `cost_rate` | money | formula |
| V | Jobs assigned | `jobs_assigned` | int | formula |
| W | Jobs done | `jobs_done` | int | formula |
| X | Jobs ahead | `upcoming` | int | formula |
| Y | Hours booked this week | `week_hours` | number1 | formula |
| Z | Week full | `load` | percent0 | formula |
| AA | Hours paid (all time) | `paid_hours` | number1 | formula |
| AB | Hours billed | `billed_hours` | number1 | formula |
| AC | Utilisation | `utilisation` | percent0 | formula |
| AD | Wages paid | `wage_cost` | money | formula |
| AE | Revenue brought in | `revenue` | money | formula |
| AF | Contribution | `contribution` | money | formula |
| AG | Average rating | `rating` | number1 | formula |
| AH | Complaints | `complaints` | int | formula |
| AI | Late or absent | `late` | int | formula |
| AJ | List order | `name_rank` | int | formula, hidden |
| AK | Rank | `revenue_rank` | int | formula, hidden |
| AL | Rank | `week_hours_rank` | int | formula, hidden |
| AM | Check | `alert` | status | formula |

#### Timesheets

Hours worked, what they cost you and how much of them a client actually paid for.

26 columns, 800 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Day | `weekday` | text | formula |
| D | Week | `week` | int | formula |
| E | Month | `month` | text | formula |
| F | Cleaner | `employee` | text | drop down `EmployeeNames` |
| G | Job ID | `job` | id | drop down `JobIDs` |
| H | Client | `client_name` | text | formula |
| I | Activity | `activity` | text | drop down `MileagePurposeList` |
| J | Start | `start` | time | you type it (checked) |
| K | Finish | `finish` | time | you type it (checked) |
| L | Break (min) | `break_min` | int | you type it (checked) |
| M | Hours (override) | `hours_input` | number1 | you type it (checked) |
| N | Of which overtime | `overtime` | number1 | you type it (checked) |
| O | Attendance | `attendance` | status | drop down `AttendanceList` |
| P | Billable? | `billable` | flag | drop down `YesNoList` |
| Q | Approved? | `approved` | flag | drop down `YesNoList` |
| R | Paid? | `paid` | flag | drop down `YesNoList` |
| S | Notes | `notes` | wrap | you type it |
| T | Clocked | `clock_hours` | number1 | formula, hidden |
| U | Hours | `hours` | number1 | formula |
| V | Cost an hour | `rate` | money | formula, hidden |
| W | Basic hours | `basic_hours` | number1 | formula, hidden |
| X | Cost to you | `cost` | money | formula |
| Y | Billed value | `billed_value` | money | formula, hidden |
| Z | Check | `alert` | status | formula |

#### Job Costing

Every job, every cost, every margin - worked out for you. Only two columns here accept typing.

28 columns, 600 rows of capacity, 2 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Job ID | `job` | id | formula |
| B | Date | `date` | date | formula |
| C | Month | `month` | text | formula |
| D | Client ID | `client` | id | formula |
| E | Client | `client_name` | text | formula |
| F | Property | `property_name` | text | formula |
| G | Service | `service` | text | formula |
| H | Cleaner | `employee` | text | formula |
| I | Status | `status` | status | formula |
| J | Revenue | `revenue` | money | formula |
| K | Hours quoted | `est_hours` | number1 | formula, hidden |
| L | Hours worked | `actual_hours` | number1 | formula |
| M | Hours costed | `hours_used` | number1 | formula, hidden |
| N | Wages from timesheets | `timesheet_cost` | money | formula, hidden |
| O | Wages | `labour` | money | formula |
| P | Supplies logged | `supply_logged` | money | formula, hidden |
| Q | Supplies | `supplies` | money | formula |
| R | Travel | `travel` | money | formula |
| S | Other cost | `other` | money | you type it (checked) |
| T | Total cost | `cost` | money | formula |
| U | Profit | `profit` | money | formula |
| V | Margin | `margin` | percent0 | formula |
| W | Profit an hour | `per_hour` | money | formula |
| X | Hours over estimate | `hours_over` | number1 | formula, hidden |
| Y | Rank | `profit_rank` | int | formula, hidden |
| Z | Rank | `per_hour_rank` | int | formula, hidden |
| AA | Verdict | `alert` | status | formula |
| AB | Notes | `notes` | wrap | you type it |

#### Invoices

What you have billed, what has come in and what is late.

31 columns, 400 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Invoice date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula |
| D | Client ID | `client` | id | drop down `ClientIDs` |
| E | Client | `client_name` | text | formula |
| F | Billing address | `address` | text | formula, hidden |
| G | Covers | `period` | text | you type it |
| H | Status | `status` | status | drop down `InvoiceStatusList` |
| I | Sent on | `sent_date` | date | you type it (checked) |
| J | Terms (days) | `terms_input` | int | you type it (checked) |
| K | Discount | `discount` | percent0 | you type it (checked) |
| L | Pay by | `method` | text | drop down `PaymentMethodList` |
| M | PO / reference | `po` | text | you type it |
| N | Reminders sent | `reminders` | int | you type it (checked) |
| O | Last reminder | `last_reminder` | date | you type it (checked) |
| P | Notes | `notes` | wrap | you type it |
| Q | Lines | `lines` | int | formula, hidden |
| R | Subtotal | `subtotal` | money | formula |
| S | Less discount | `discount_value` | money | formula, hidden |
| T | Net | `net` | money | formula |
| U | Tax | `tax` | money | formula |
| V | Total | `total` | money | formula |
| W | Received | `paid` | money | formula |
| X | Still owed | `outstanding` | money | formula |
| Y | Terms used | `terms_used` | int | formula, hidden |
| Z | Due | `due_date` | date | formula |
| AA | Days late | `days_late` | int | formula |
| AB | Age bracket | `age` | status | formula |
| AC | Late fee due | `late_fee` | money | formula, hidden |
| AD | Settled? | `fully_paid` | flag | formula, hidden |
| AE | Action | `alert` | status | formula |

#### Invoice Lines

One row per job (or per extra) on each invoice.

18 columns, 800 rows of capacity, 7 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Invoice ID | `invoice` | id | drop down `InvoiceIDs` |
| C | Job ID | `job` | id | drop down `JobIDs` |
| D | Job date | `date` | date | formula |
| E | Description | `description` | wrap | you type it |
| F | Quantity | `qty` | number1 | you type it (checked) |
| G | Price each (override) | `price_input` | money | you type it (checked) |
| H | Taxable? | `taxable` | flag | drop down `YesNoList` |
| I | Notes | `notes` | wrap | you type it |
| J | Job price | `job_price` | money | formula, hidden |
| K | Service | `job_service` | text | formula, hidden |
| L | Wording used | `wording` | wrap | formula, hidden |
| M | Price each | `price_used` | money | formula |
| N | Line total | `total` | money | formula |
| O | Billed twice? | `duplicate` | status | formula |
| P | Line number | `seq` | int | formula, hidden |
| Q | Lookup key | `key` | text | formula, hidden |
| R | Check | `alert` | status | formula |

#### Payments

Every payment received, matched to its invoice.

18 columns, 400 rows of capacity, 9 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date received | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula |
| D | Invoice ID | `invoice` | id | drop down `InvoiceIDs` |
| E | Client ID | `client` | id | formula |
| F | Client | `client_name` | text | formula |
| G | Amount | `amount` | money | you type it (checked) |
| H | Processing fee | `fee` | money | you type it (checked) |
| I | Method | `method` | text | drop down `PaymentMethodList` |
| J | Reference | `reference` | text | you type it |
| K | Is this a deposit? | `deposit` | flag | drop down `YesNoList` |
| L | Banked? | `banked` | flag | drop down `YesNoList` |
| M | Notes | `notes` | wrap | you type it |
| N | Net received | `net` | money | formula |
| O | Invoice total | `invoice_total` | money | formula, hidden |
| P | Paid on invoice | `invoice_paid` | money | formula, hidden |
| Q | Invoice balance after | `balance` | money | formula |
| R | Check | `alert` | status | formula |

#### Supplies

What you hold, what you use and what it is really costing per job.

26 columns, 150 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Item | `name` | text | you type it |
| C | Category | `category` | text | drop down `SupplyCategoryList` |
| D | Supplier | `supplier` | text | you type it |
| E | Unit | `unit` | text | drop down `SupplyUnitList` |
| F | Units in a pack | `pack_size` | number1 | you type it (checked) |
| G | Pack cost | `pack_cost` | money | you type it (checked) |
| H | Counted stock | `stock` | number1 | you type it (checked) |
| I | Counted on | `count_date` | date | you type it (checked) |
| J | Reorder at | `reorder` | number1 | you type it (checked) |
| K | Order quantity | `order_qty` | number1 | you type it (checked) |
| L | Last ordered | `last_order` | date | you type it (checked) |
| M | Eco / low tox? | `eco` | flag | drop down `YesNoList` |
| N | Hazard sheet? | `coshh` | flag | drop down `YesNoNAList` |
| O | Notes | `notes` | wrap | you type it |
| P | Cost a unit | `unit_cost` | money | formula |
| Q | Used (all time) | `used_total` | number1 | formula |
| R | Used since the count | `used_since` | number1 | formula, hidden |
| S | On hand | `on_hand` | number1 | formula |
| T | Used in 28 days | `used_28` | number1 | formula, hidden |
| U | Used a day | `daily_use` | number1 | formula, hidden |
| V | Days of cover | `days_cover` | int | formula |
| W | Stock value | `stock_value` | money | formula |
| X | Spent on it | `spend` | money | formula |
| Y | Stock | `status` | status | formula |
| Z | Action | `alert` | status | formula |

#### Supply Usage

What you used, on which job, and what it cost.

13 columns, 300 rows of capacity, 7 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula |
| D | Item | `supply` | text | drop down `SupplyNames` |
| E | Quantity used | `qty` | number1 | you type it (checked) |
| F | Job ID | `job` | id | drop down `JobIDs` |
| G | Used by | `employee` | text | drop down `EmployeeNames` |
| H | Why | `reason` | text | you type it |
| I | Notes | `notes` | wrap | you type it |
| J | Client | `client_name` | text | formula |
| K | Cost a unit | `unit_cost` | money | formula, hidden |
| L | Cost | `cost` | money | formula |
| M | Check | `alert` | status | formula |

#### Equipment

What you own, what it costs to keep and when it next needs looking at.

26 columns, 80 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Equipment | `name` | text | you type it |
| C | Type | `type` | text | drop down `EquipmentTypeList` |
| D | Make & model | `make` | text | you type it |
| E | Serial / registration | `serial` | text | you type it |
| F | Ownership | `ownership` | text | drop down `OwnershipList` |
| G | Status | `status` | status | drop down `EquipmentStatusList` |
| H | Assigned to | `assigned` | text | drop down `EmployeeNames` |
| I | Kept at | `location` | text | drop down `VehicleList` |
| J | Bought on | `bought` | date | you type it (checked) |
| K | Cost | `cost` | money | you type it (checked) |
| L | Expected life (years) | `life_years` | number1 | you type it (checked) |
| M | Warranty to | `warranty` | date | you type it (checked) |
| N | Last serviced | `last_service` | date | you type it (checked) |
| O | Service every (days) | `interval_input` | int | you type it (checked) |
| P | Spent on repairs | `service_cost` | money | you type it (checked) |
| Q | Safety tested on | `safety_test` | date | you type it (checked) |
| R | Insured? | `insured` | flag | drop down `YesNoNAList` |
| S | Notes | `notes` | wrap | you type it |
| T | Age (years) | `age_years` | number1 | formula |
| U | Costs a year | `annual_cost` | money | formula |
| V | Interval used | `interval` | int | formula, hidden |
| W | Next service due | `next_service` | date | formula |
| X | Days to service | `days_to_service` | int | formula, hidden |
| Y | Replace by | `replace_year` | int | formula, hidden |
| Z | Action | `alert` | status | formula |

#### Expenses

Everything that goes out, grouped the way your accountant wants to see it.

21 columns, 400 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula |
| D | Category | `category` | text | drop down `ExpenseCategoryList` |
| E | Type | `type` | text | drop down `ExpenseTypeList` |
| F | What it was | `description` | text | you type it |
| G | Paid to | `supplier` | text | you type it |
| H | Amount (net) | `amount` | money | you type it (checked) |
| I | Tax | `tax` | money | you type it (checked) |
| J | Paid by | `method` | text | drop down `PaymentMethodList` |
| K | Job ID | `job` | id | drop down `JobIDs` |
| L | Who spent it | `employee` | text | drop down `EmployeeNames` |
| M | Receipt? | `receipt` | flag | drop down `YesNoList` |
| N | Recurring? | `recurring` | flag | drop down `YesNoList` |
| O | Claimable? | `deductible` | flag | drop down `YesNoNAList` |
| P | Notes | `notes` | wrap | you type it |
| Q | Total paid | `total` | money | formula |
| R | This category, this month | `month_total` | money | formula, hidden |
| S | Share of the month | `share` | percent0 | formula, hidden |
| T | On behalf of | `client_name` | text | formula, hidden |
| U | Check | `alert` | status | formula |

#### Mileage

Every trip, what it cost and which job should carry it.

20 columns, 400 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula |
| D | Driver | `employee` | text | drop down `EmployeeNames` |
| E | Vehicle | `vehicle` | text | drop down `VehicleList` |
| F | Purpose | `purpose` | text | drop down `MileagePurposeList` |
| G | From | `from_place` | text | you type it |
| H | To | `to_place` | text | you type it |
| I | Job ID | `job` | id | drop down `JobIDs` |
| J | Distance (miles) | `distance` | number1 | you type it (checked) |
| K | Return trip? | `return_trip` | flag | drop down `YesNoList` |
| L | Rate (override) | `rate_input` | money | you type it (checked) |
| M | Parking & tolls | `parking` | money | you type it (checked) |
| N | Claimed? | `claimed` | flag | drop down `YesNoList` |
| O | Notes | `notes` | wrap | you type it |
| P | Client | `client_name` | text | formula, hidden |
| Q | Total miles | `total_distance` | number1 | formula |
| R | Rate | `rate_used` | money | formula, hidden |
| S | Travel cost | `amount` | money | formula |
| T | Check | `alert` | status | formula |

#### Complaints

What went wrong, how fast you fixed it and whether you kept the client.

25 columns, 80 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date raised | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula |
| D | Client ID | `client` | id | drop down `ClientIDs` |
| E | Client | `client_name` | text | formula |
| F | Job ID | `job` | id | drop down `JobIDs` |
| G | Cleaner on the job | `employee` | text | drop down `EmployeeNames` |
| H | What about | `type` | text | drop down `ComplaintTypeList` |
| I | Severity | `severity` | status | drop down `SeverityList` |
| J | What they said | `description` | wrap | you type it |
| K | Status | `status` | status | drop down `ComplaintStatusList` |
| L | Being handled by | `owner` | text | drop down `EmployeeNames` |
| M | What we did | `action` | wrap | you type it |
| N | Resolution | `resolution` | text | drop down `ResolutionList` |
| O | Re-clean job | `reclean_job` | id | drop down `JobIDs` |
| P | Cost to put right | `cost` | money | you type it (checked) |
| Q | Resolved on | `resolved_date` | date | you type it (checked) |
| R | Client kept? | `kept` | flag | drop down `YesNoList` |
| S | Rating afterwards | `rating_after` | rating | you type it (checked) |
| T | How we stop it happening again | `prevent` | wrap | you type it |
| U | Notes | `notes` | wrap | you type it |
| V | Days open | `days_open` | int | formula |
| W | Still open? | `open` | flag | formula, hidden |
| X | Repeat issue? | `repeat` | flag | formula |
| Y | Action | `alert` | status | formula |

#### Incidents

Accidents, damage and the things that almost went badly wrong.

24 columns, 60 rows of capacity, 21 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Time | `time` | text | you type it |
| D | What happened | `type` | text | drop down `IncidentTypeList` |
| E | Person involved | `employee` | text | drop down `EmployeeNames` |
| F | Job ID | `job` | id | drop down `JobIDs` |
| G | Property ID | `property` | id | drop down `PropertyIDs` |
| H | Severity | `severity` | status | drop down `SeverityList` |
| I | What happened, in full | `description` | wrap | you type it |
| J | Injury? | `injury` | flag | drop down `YesNoList` |
| K | First aid given? | `first_aid` | flag | drop down `YesNoNAList` |
| L | Medical help needed? | `medical` | flag | drop down `YesNoNAList` |
| M | Property damaged | `damage` | wrap | you type it |
| N | Cost | `cost` | money | you type it (checked) |
| O | Insurance claim? | `insurance` | flag | drop down `YesNoNAList` |
| P | Reported to | `reported_to` | text | you type it |
| Q | Reported on | `reported_date` | date | you type it (checked) |
| R | Action taken | `action` | wrap | you type it |
| S | How we prevent a repeat | `prevent` | wrap | you type it |
| T | Status | `status` | status | drop down `ComplaintStatusList` |
| U | Closed on | `closed_date` | date | you type it (checked) |
| V | Notes | `notes` | wrap | you type it |
| W | Days open | `days_open` | int | formula |
| X | Action | `alert` | status | formula |

#### Safety

Certificates, checks and assessments, with the expiry dates watched for you.

18 columns, 80 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Document | `document` | text | drop down `SafetyDocList` |
| C | Applies to | `applies_to` | text | you type it |
| D | Person | `employee` | text | drop down `EmployeeNames` |
| E | Reference / number | `reference` | text | you type it |
| F | Issued by | `issuer` | text | you type it |
| G | Issued on | `issued` | date | you type it (checked) |
| H | Expires | `expires` | date | you type it (checked) |
| I | Cost to renew | `cost` | money | you type it (checked) |
| J | Where it is kept | `holder` | text | you type it |
| K | Link | `link` | link | you type it |
| L | Required? | `required` | flag | drop down `YesNoNAList` |
| M | Last reviewed | `reviewed` | date | you type it (checked) |
| N | Notes | `notes` | wrap | you type it |
| O | Days to expiry | `days_left` | int | formula |
| P | Status | `status` | status | formula |
| Q | Start renewing | `renew_by` | date | formula, hidden |
| R | Action | `alert` | status | formula |

---

*13,817 lines of Python in 53 files.*
