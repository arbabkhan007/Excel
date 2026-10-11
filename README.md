# Pet Sitting & Dog Walking Business Spreadsheet — generator

> **This repository holds two spreadsheet products** built on the same
> engine, with the same conventions, the same author (Novality Store) and
> the same `premium` sheet protection:
>
> * **Pet Sitting & Dog Walking Business Spreadsheet** — this README,
>   `build_pet_sitting_workbook.py`, 27 tabs.
> * **Social Media Content Calendar & Analytics Planner** —
>   [README_SOCIAL.md](README_SOCIAL.md),
>   `build_social_media_workbook.py`, 43 tabs.

Python code that builds a finished, ready-to-sell Excel workbook for pet
sitters and dog walkers.

> **Listing title:** Pet Sitting Business Spreadsheet, Dog Walker Client
> Booking, Payments, Mileage and Profit Tracker
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_pet_sitting_workbook.py
# -> output/Pet Sitting Business Spreadsheet.xlsx   (27 tabs, ~860 KB)
```

The repository also contains the XlsxWriter library itself (`xlsxwriter/`),
so the command above runs with **no installation and no dependencies** from
the repository root. Outside this repo, `pip install xlsxwriter` first — the
generator is tested against the bundled 0.3.0 and against XlsxWriter 3.2.

---

## 1. What gets built

27 tabs, 50,000+ formulas, all joined up and all driven by drop downs.

| # | Tab | What it does |
|---|-----|--------------|
| 1 | **Start Here** | Welcome page, 10-step quick start, clickable map of every tab, setup checklist with a progress bar, colour key |
| 2 | **Dashboard** | Pick a year and month (or "All year") and 24 KPI tiles, a 12-month table with sparklines, revenue by service, expenses by category, top 10 clients and an alert snapshot all follow the picker |
| 3 | **Settings** | Business details, currency, tax, invoice numbering, surcharges, mileage and travel rates, reminder windows, public holidays — ~60 named settings that drive every formula |
| 4 | **Lists** | 47 editable drop down lists; add an option and every sheet picks it up |
| 5 | **Services** | Price list: billing unit, rate, duration, extra pet fee, staff cost, margin, bookings and revenue per service |
| 6 | **Clients** | Client database: contacts, keys and access codes, contracts, vet consent, lifetime revenue, outstanding balance, loyalty tier, duplicate and "going quiet" alerts |
| 7 | **Pets** | Pet profiles: breed, age, weight, microchip, temperament, allergies, medication, vet and emergency clinic, vaccination countdown, feeding and walking routine, risk flags |
| 8 | **Bookings** | The schedule. Prices itself from Services and applies extra pet, weekend, holiday, travel and last-minute rules, flags double bookings, tracks confirmations and invoicing |
| 9 | **Visit Log** | Proof of service: arrival/departure, on-time check, walk distance, feeding, water, medication, potty, mood, appetite, health notes, photo link, client update sent |
| 10 | **Health Log** | Medication administration, vaccinations and vet appointments with refill and appointment countdowns |
| 11 | **Incidents** | Emergency contact panel plus a full incident register (severity, vet, owner notified, insurance, days open) |
| 12 | **Messages** | Communication log, enquiry source, feedback, ratings, complaints and follow-up chasers |
| 13 | **Invoices** | Automatic invoice numbers, due dates, tax, deposits, payments applied, balance, ageing buckets, suggested late fee |
| 14 | **Payments** | Cash in: method, deposits, tips, refunds, late fees, bank reconciliation |
| 15 | **Expenses** | Every cost split into tax / net / claimable, business vs personal, recurring with next due date, receipt chaser |
| 16 | **Mileage** | Odometer trip log, round trips, business classification, mileage claim, parking and tolls |
| 17 | **Staff** | Team directory, certificates with expiry alerts, availability, visits and hours worked, revenue generated, wages due |
| 18 | **Inventory** | Supplies, stock levels, reorder points, expiry alerts, value on hand |
| 19 | **Reminders** | 24 automatic alerts (overdue invoices, unconfirmed bookings, missed meds, expired vaccinations, low stock …) with jump links, renewal dates and your own task list |
| 20 | **Reports** | Set a From/To date and get bookings, profit & loss, revenue by service, top clients, ageing, payment methods, expenses by category, staff performance, retention, referrals and cancellation analysis |
| 21 | **Budget** | Targets, monthly budget vs actual with sparklines, break-even, forecast, cash flow, owner drawings, savings and year-on-year comparison |
| 22 | **Print Invoice** | Pick an invoice number → a client-ready A4 invoice with your branding, the visits covered and payment details |
| 23 | **Print Pet Care** | Pick a pet → a one page care sheet for a sitter: routine, medication, vet, risks, recent visits |
| 24 | **Print Visit** | Pick a visit → a visit report to hand or send to the client |
| 25 | **Print Forms** | 10 blank forms (registration, pet profile, service agreement & vet consent, key handover, incident report, daily checklist, timesheet, mileage log, feedback, meet & greet) each on its own page |
| 26 | **Help** | How it works, "how do I …", what every formula does, troubleshooting, protection, data privacy, compatibility, version history |
| 27 | *Calc* (hidden) | The aggregation engine: period-aware client, service, expense, staff, month, referral, cancellation and payment-method blocks that keep the visible formulas short |

### Product rules baked in

* **Formulas are locked, inputs are not.** Grey cells are locked formulas,
  cream cells are unlocked inputs; every sheet is protected with the
  password `premium`, so a stray keystroke cannot destroy the maths.
  Filtering, row insertion and formatting stay available.
* **Document author is `Novality Store`** (`dc:creator`, last modified by,
  company and manager) on every generated file.
* **No hard coded column letters anywhere.** Formulas are written as
  templates (`[[Invoices.balance]]`, `{client}{r}`) and resolved from the
  sheet specification, so a column can be inserted or moved without
  breaking a single formula.
* **Excel 2010+ safe.** No dynamic arrays, no `XLOOKUP`, no `IFS`, no
  `AGGREGATE`, no array entry required — min/max-if style maths uses
  `SUMPRODUCT`. Everything also survives a trip through Google Sheets.
* **Example data included** (9 clients, 12 pets, 12 services, 4 staff, 32
  bookings, visits, 23 invoices, 15 payments, 22 expenses, trips,
  incidents, messages, inventory and tasks), all cross-referenced so the
  dashboard is alive on first open — a profitable little demo business
  with a couple of overdue invoices left in on purpose so the alerts and
  the ageing report have something to show. `--blank` ships the empty
  version.

---

## 1b. Download it as one file

Everything needed to build the workbook is packed into

    download/pet-sitting-spreadsheet-source.zip   (220 KB)

[Download the zip](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/pet-sitting-spreadsheet-source.zip)
— unzip it anywhere and run:

```bash
pip install -r requirements.txt
python build_pet_sitting_workbook.py
```

It holds the generator, the quality tools, this README, `SOURCE_CODE.md`,
a `requirements.txt` and a plain `HOW TO RUN.txt`. The vendored copy of
XlsxWriter in this repository is left out on purpose — the zip runs
against a normal `pip install XlsxWriter` (tested on 3.2.0). Rebuild it
after a code change with `python tools/make_download_zip.py`.

---

## 2. Command line

```bash
python build_pet_sitting_workbook.py [options]
python -m pet_sitting_spreadsheet  [options]      # same thing
```

| Option | Default | Notes |
|--------|---------|-------|
| `-o`, `--output` | `output/Pet Sitting Business Spreadsheet.xlsx` | where to write the file |
| `--blank` | off | no example records — the version a buyer starts from |
| `--no-protect` | off | skip sheet protection (development only) |
| `--password` | `premium` | sheet protection password |
| `--author` | `Novality Store` | document creator |
| `--business` | `Happy Paws Pet Care` | pre-fills the business name on Settings |
| `--currency` / `--currency-code` | `$` / `USD` | e.g. `--currency "£" --currency-code GBP` |
| `--unit` | `km` | `km` or `miles` — retitles the mileage columns |
| `--rows` | 400 | capacity of the Bookings register; every other register scales with it |
| `--today` | today | treat this ISO date as "today" when generating the example data |
| `--quiet` | off | print only the output path |

Examples:

```bash
python build_pet_sitting_workbook.py --blank -o "dist/Pet Sitting (blank).xlsx"
python build_pet_sitting_workbook.py --currency "£" --currency-code GBP --unit miles
python build_pet_sitting_workbook.py --rows 1000 --business "Happy Paws"
```

From Python:

```python
from pet_sitting_spreadsheet import Config, build_workbook

cfg = Config(include_sample_data=False, currency_symbol="€",
             currency_code="EUR", tax_rate=0.21)
build_workbook("Pet Sitting (EU).xlsx", cfg)
```

---

## 3. How the code is organised

```
pet_sitting_spreadsheet/
    config.py        branding, palette, row counts, pricing defaults, Config
    layout.py        Col / SheetSpec / Registry + the formula template engine
    formats.py       every cell format (input = unlocked cream, calc = locked grey)
    writer.py        SheetWriter: banner, nav, sections, KPI tiles, grid(), protect()
    sample_data.py   the cross-referenced demo dataset
    names.py         live defined names (ClientNames, PetLabels, InvoiceNos …)
    build.py         tab order, builds every sheet, sets document properties
    cli.py           command line front end
    sheets/
        lists.py settings.py services.py clients.py pets.py bookings.py
        visits.py health.py incidents.py messages.py invoices.py payments.py
        expenses.py mileage.py staff.py inventory.py reminders.py   <- registers
        calc.py dashboard.py reports.py budget.py start_here.py
        printables.py help_tab.py                                   <- pages
tools/
    validate_workbook.py    static QA on a generated file
    check_sample_data.py    makes sure the demo records join up
    recalc_check.py         calculates every formula and checks the answer
    export_source_md.py     writes the whole source into SOURCE_CODE.md
    make_download_zip.py    packs download/…-source.zip
build_pet_sitting_workbook.py
SOURCE_CODE.md              every file above in one readable document
download/
    pet-sitting-spreadsheet-source.zip   ready to download, unzip and run
```

**Registers** (the data tabs) expose `spec(cfg) -> SheetSpec` — a list of
`Col` objects with the header, width, number style, validation, conditional
formats and, for calculated columns, a formula template. `build.py` turns
each spec into a sheet with `SheetWriter.grid()`: banner, nav strip, header
row, N pre-built rows, drop downs, header comments, conditional formats,
freeze panes, autofilter, print setup and protection. A register may also
export `decorate(sheet, fmts, cfg[, reg])` to add a panel above the grid
(Incidents and Reminders do).

**Pages** (Dashboard, Reports, Budget, Start Here, the print sheets, Help)
expose `build(book, fmts, cfg, reg)` and lay themselves out freely.

### Formula templates

Formulas never contain a hard coded column letter:

| Token | Becomes |
|-------|---------|
| `{r}` / `{r_prev}` | the current / previous Excel row number |
| `{client}` | the column letter of `client` on the same sheet |
| `{first}` / `{last}` | first / last data row of the sheet |
| `[[Payments.net]]` | `Payments!$K$5:$K$304` |
| `[[Invoices.total#12]]` | `Invoices!$O$12` |
| `[[Calc.m_rev#]]` | the same row of `Calc.m_rev` (used by the engine) |

```python
Col("balance", "Balance due", 12, "money",
    formula=blank("client", 'ROUND(IF({status}{r}="Cancelled",0,'
                            '{total}{r}-{paid}{r}),2)'))
```

### Adding something

* **A drop down option** — edit `sheets/lists.py` (or just edit the Lists tab
  in the finished file; the names are dynamic ranges).
* **A column** — add a `Col(...)` to that sheet's `spec()`. Nothing else
  changes: other formulas follow the new letters automatically.
* **A tab** — write `sheets/mytab.py` with `spec()` or `build()`, then add it
  to `REGISTER_MODULES` / `PAGE_BUILDERS` and `SHEET_ORDER` in `build.py`.
* **A setting** — add one `s.field(..., name="MyName")` line in
  `sheets/settings.py` and use `MyName` in any formula.

---

## 4. Quality checks

```bash
python build_pet_sitting_workbook.py -o output/test.xlsx
python tools/validate_workbook.py   output/test.xlsx      # needs openpyxl
python tools/check_sample_data.py   output/test.xlsx

python build_pet_sitting_workbook.py --rows 120 -o /tmp/small.xlsx
python tools/recalc_check.py /tmp/small.xlsx              # needs formulas
```

`validate_workbook.py` reports: tab order and the hidden Calc sheet,
document author, sheet protection and passwords, overlapping merges or
hyperlinks trapped inside merges (Excel calls these corrupt), unresolved
template tokens, `#REF!`, unbalanced brackets or quotes, empty arguments,
unknown or too-modern functions, formulas pointing at a missing sheet,
defined names and drop down sources that do not exist, and formula cells
that are not locked.

`check_sample_data.py` re-reads the finished file and confirms every
example booking, visit, invoice, payment, pet and task points at a record
that exists — that payments sit on their own client's invoice, that a
visit matches the client and pet on its booking, and that an invoice
raised from a booking is addressed to that booking's client.

`recalc_check.py` goes further and actually *runs* the workbook: it loads
it into the `formulas` calculation engine, evaluates the whole dependency
graph, then reads the answer out of every formula cell — so a formula that
parses perfectly but divides by zero on an empty row is still caught.
Check a `--rows 120` build rather than a full-size one: the formulas are
identical, there are just fewer rows of them, and 50,000 of them need
several GB of memory (the tool says so and stops unless you pass
`--force`). The engine cannot resolve `OFFSET()` defined names, so a
throwaway copy with those flattened is made first; your file is not
touched.

Current status: all three clean and `pyflakes` clean. Every formula in the
workbook has been calculated and lands on a real answer — checked on the
sample build, on `--blank`, and with prices-include-tax and
not-tax-registered settings, so the money adds up in each tax mode.

---

## 5. What is deliberately *not* in version 1

The brief listed 500 features; v1 ships the ones a working pet sitter uses
every day. These are the natural premium / v2 additions:

* staff scheduling grid, shift swaps, payroll export and commission splits;
* purchase orders, supplier price lists and stock movements;
* multi-year forecasting, scenario planning and loan/asset registers;
* automated email/SMS merge files, review requests and marketing campaigns;
* package and gift card balances, subscription billing and auto-renewals;
* a Google Sheets edition with Apps Script automation.

---

© Novality Store. Generated workbooks are licensed for personal use and for
a single business.
