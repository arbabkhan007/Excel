# Cleaning Business CRM & Pricing Calculator — generator

> **This repository holds four spreadsheet products** built on the same
> engine, with the same conventions, the same author (Novality Store) and
> the same `premium` sheet protection:
>
> * **Pet Sitting & Dog Walking Business Spreadsheet** —
>   [README.md](README.md), `build_pet_sitting_workbook.py`, 27 tabs.
> * **Social Media Content Calendar & Analytics Planner** —
>   [README_SOCIAL.md](README_SOCIAL.md),
>   `build_social_media_workbook.py`, 43 tabs.
> * **Automated Meal Planner & Grocery List Spreadsheet** —
>   [README_MEAL.md](README_MEAL.md),
>   `build_meal_planner_workbook.py`, 36 tabs.
> * **Cleaning Business CRM & Pricing Calculator** — this README,
>   `build_cleaning_workbook.py`, 37 tabs.

Python code that builds a finished, ready-to-sell Excel workbook for a
cleaning company: keep the clients and their properties, price a job in
seconds, put it in the diary, send the cleaner with the right
checklist, bill it, chase it, and see what each client is actually
worth.

> **Listing title:** Cleaning Business Spreadsheet, Cleaning Service
> CRM, Job Scheduler, Quote Calculator, Employee Planner, Invoices,
> Expenses and Profit Dashboard
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_cleaning_workbook.py
# -> output/Cleaning Business CRM & Pricing Calculator.xlsx
#    (37 tabs, ~1.7 MB)
```

The repository also contains the XlsxWriter library itself
(`xlsxwriter/`), so the command above runs with **no installation and no
dependencies** from the repository root. Outside this repo,
`pip install xlsxwriter` first — the generator is tested against the
bundled 0.3.0 and against XlsxWriter 3.2.

---

## 1. What gets built

37 tabs, 110,800+ formulas, all joined up and all driven by drop downs.

| # | Tab | What it does |
|---|-----|--------------|
| 1 | **Start Here** | Welcome page, the thirty minute set-up in order, the daily and weekly rhythm, a map of every tab, the colour key and an honest note on what a spreadsheet can and cannot do |
| 2 | **Settings** | The business name and contact details, the year, pricing (hourly rate, minimum charge, per-room and per-area rates, uplifts, deposit, rounding), costs (pay rate, overtime, on-costs, supplies an hour, mileage rate, target margin), tax, payment terms, targets and the ID prefixes — around 70 named settings the rest of the workbook reads |
| 3 | **Dashboard** | The business in one screen: KPI tiles for revenue, profit, margin, jobs, hours, utilisation and money owed, a "needs your attention" list, this week's diary, the top clients, and a 12 month table with sparklines |
| 4 | **Quote Calculator** | **The headline feature.** Pick a client and property (or type a new one), choose the service and frequency, and it sizes the job from bedrooms, bathrooms, rooms or floor area, applies the condition and frequency factors, adds pets, call-out and travel, takes the minimum charge into account and shows the price, the cost, the profit and the margin — with a what-if row that tells you the price you need for your target margin |
| 5 | **Clients** | The CRM. One row per customer with contact details, terms, contract dates and preferences — and jobs done, last clean, next booking, invoiced, paid, owing, profit, margin, value a year, rating, complaints, a standing band and a churn-risk score worked out for you |
| 6 | **Properties** | The places you clean: type, size, bedrooms, bathrooms, floors, condition, pets, access and key details, parking, distance, and which checklist applies — with estimated hours, suggested price, labour cost, estimated margin, jobs done here and estimate-versus-actual hours |
| 7 | **Leads** | The sales pipeline: every enquiry, where it came from, what it is worth a year, the stage, the quote, the visit, the next action — with days open, days since anything happened and a "chase this today" verdict |
| 8 | **Services** | The price list: priced per hour, per visit, per room or per area, with hours, uplift, supplies cost and minimum charge — and standard price, labour cost, cost to deliver, profit, margin, what it earns an hour, the price needed for your target margin, and how often it is quoted and booked |
| 9 | **Quotes** | Every quote with its client, property, status, dates, discount, deposit and validity — totals pulled from the lines, tax, profit, margin, days waiting, expiry and a chase-or-close action |
| 10 | **Quote Lines** | One row per thing you are charging for. Catalogue price in, override allowed, discount, hours, line value, cost and margin |
| 11 | **Jobs** | **The scheduler.** One row per visit: date, start, client, property, service, type, frequency, status, priority, up to two cleaners, travel — and hours, crew, finish time, rate used, base price, price (minimum charge and frequency factor applied), actual hours from the timesheets, checklist progress, the invoice it ended up on, the next visit due and a plain-English action |
| 12 | **Job Tasks** | The checklist as it was really done: ticked off, timed, graded, with the problems and photo links written down |
| 13 | **Calendar** | Any month as a grid from one date, each day showing the jobs, the hours and the value, with today highlighted and the week totals down the side |
| 14 | **Checklists** | The standard for every service, area by area, with honest minutes, "required" and "photo needed" flags, the products and kit — and how often each task actually gets ticked, skipped or fails |
| 15 | **Employees** | The team: role, pay, contracted hours, start date, skills, DBS/background check, licences — with hours worked, cost, revenue brought in, utilisation, jobs done, average rating and complaints |
| 16 | **Staff Schedule** | The rota. Pick a week and see every cleaner's day-by-day jobs, hours and value, who is overloaded, who has gaps, and which jobs still have nobody on them |
| 17 | **Timesheets** | Hours worked with start, finish and break, attendance, overtime, approval and payment — what they cost you and how much of them a client actually paid for |
| 18 | **Job Costing** | Every job, every cost, every margin: revenue, hours used, labour at the real rate with on-costs, supplies logged, travel, other costs, profit, margin, profit an hour and a verdict |
| 19 | **Invoices** | What you have billed: client, period covered, status, terms, discount, PO reference, reminders — with totals from the lines, tax, received, still owed, due date, days late, ageing bracket and the late fee due |
| 20 | **Invoice Lines** | One row per job (or per extra) on each invoice, priced from the job and flagged if a job is billed twice |
| 21 | **Payments** | Every payment received, matched to its invoice, with method, fee, net received and the balance left on that invoice |
| 22 | **Supplies** | Stock: pack size, pack cost, cost a unit, counted stock, reorder level — with what has been used, what is on hand, usage a day, days of cover, stock value and an order-now flag |
| 23 | **Supply Usage** | What was used, on which job, by whom and what it cost — the link that puts product cost on the right job |
| 24 | **Equipment & Assets** | What you own: purchase, cost, expected life, warranty, service interval, safety test, insurance — with age, cost a year, next service due, days to go and a replace-by date |
| 25 | **Expenses** | Everything that goes out, grouped the way an accountant wants it, with tax, who spent it, receipts, recurring flags and the job it belongs to |
| 26 | **Mileage & Travel** | Every trip with from, to, distance, return flag, rate, parking and tolls — and the travel cost for the job |
| 27 | **Client Retention** | Who is slipping away: days since the last clean against their usual frequency, the at-risk list, lapsed clients, recurring contracts and what each client is worth a year |
| 28 | **Complaints** | What went wrong, how fast you fixed it, what it cost to put right, whether the client stayed and how you stop it happening again |
| 29 | **Incidents** | Accidents, damage and near misses, with injury, first aid, insurance, who it was reported to and the preventive action |
| 30 | **Safety & Compliance** | Insurance, certificates, risk assessments, background checks and training with expiry dates watched for you and a start-renewing date |
| 31 | **Financial Plan** | Twelve months of planned revenue, costs and profit against actual, break-even jobs a month, the hourly rate you need, and a what-if on rate, hours and utilisation |
| 32 | **Reports** | Pick a month: revenue and profit by client, by service and by cleaner, jobs and hours, the quote win rate, the debtor list and the year in one table |
| 33 | **Print Quote** | Any quote on one page for the client, with the lines, totals, deposit and terms |
| 34 | **Print Invoice** | Any invoice on one page with the lines, the total, what is paid and the payment details |
| 35 | **Print Job Sheet** | The job sheet the cleaner takes with them: address, access and key details, pets, parking, the checklist for that property, the time allowed and a sign-off box |
| 36 | **Lists** | Every drop down in the workbook — around 60 of them — each a named range the registers point at |
| 37 | **Help** | Set-up in order, what every tab is for, the automation chain in plain English, troubleshooting, the honest limits and the shortcuts |

### Product rules baked in

* **Nothing is typed twice.** A client is entered once, a property is
  measured once, a service is priced once. Quotes, jobs, invoices,
  costing, the dashboard and the reports all read those three.
* **The price is the product.** Settings holds your rates → Properties
  turns bedrooms, bathrooms, rooms or square feet into hours →
  Services says what an hour of that work is worth → the Quote
  Calculator applies condition, frequency, pets, call-out and travel →
  Jobs prices every visit with the minimum charge respected → Job
  Costing puts the real labour, supplies and travel against it. Plain
  formulas all the way down, no macros.
* **Every formula cell is locked**, every sheet is protected with
  `premium`, and the cream coloured cells are the only ones left open —
  8,800 of them, so there is plenty of room to type.
* **Excel 2010 safe.** `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `SMALL`,
  `LARGE`, `EOMONTH`, `SUMPRODUCT` — no dynamic arrays, no `XLOOKUP`,
  no `IFS`, nothing that needs Ctrl+Shift+Enter. It opens in Google
  Sheets, LibreOffice and Numbers.
* **No charts, by design** — the vendored XlsxWriter cannot write them.
  Trends use sparklines, data bars and colour scales, which survive a
  Google Sheets import better anyway.
* **"Top five" lists without array formulas.** Every ranking is a
  hidden rank column on the register itself (`profit_rank`,
  `invoiced_rank`, `risk_rank`, `unassigned_rank`, `future_rank`,
  `name_rank`…) and the page just does `MATCH(n, rank, 0)`.
* **A quote is never silently wrong.** The minimum charge, the
  frequency factor and the rounding rule are settings, not magic
  numbers buried in a formula, and the calculator shows the cost and
  the margin next to the price so nothing goes out below the floor.

---

## 1b. Download it

Everything is committed to the repository, so these links work straight
from GitHub — right click, *Save link as*.

| What | Size | Link |
|------|-----:|------|
| **The finished workbook, with the example company filled in** | 1,742 KB | [cleaning-business-crm.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/cleaning-business-crm.xlsx) |
| **The finished workbook, empty and ready to sell** | 1,685 KB | [cleaning-business-crm-blank.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/cleaning-business-crm-blank.xlsx) |
| **The generator source, zipped** | 374 KB | [cleaning-business-crm-source.zip](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/cleaning-business-crm-source.zip) |

The two `.xlsx` files are the product itself: 37 tabs, every formula
cell locked, every sheet protected with `premium`, author Novality
Store. Open the blank one, fill in **Settings**, then **Clients** and
their **Properties**, and work down the **Start Here** tab.

The sample file is a small cleaning company with a full quarter behind
it — 15 clients, 16 properties, 21 services, 8 cleaners, 127 jobs, 25
invoices, 95 timesheets and the supplies, mileage, complaints and
compliance records that go with them — so every tile on the dashboard
has a real number in it.

The zip is for anyone who wants to rebuild or change it — unzip
anywhere and run:

```bash
pip install -r requirements.txt
python build_cleaning_workbook.py
```

It holds the generator, the quality tools, this README,
`SOURCE_CODE_CLEANING.md`, a `requirements.txt` and a plain
`HOW TO RUN.txt`. The vendored copy of XlsxWriter in this repository is
left out on purpose — the zip runs against a normal
`pip install XlsxWriter` (tested on 3.2.0). Rebuild everything after a
code change with:

```bash
python build_cleaning_workbook.py -o download/cleaning-business-crm.xlsx
python build_cleaning_workbook.py --blank -o download/cleaning-business-crm-blank.xlsx
python tools/export_source_md.py --product cleaning
python tools/make_download_zip.py --product cleaning
```

---

## 2. Command line

```bash
python build_cleaning_workbook.py        [options]
python -m cleaning_business_spreadsheet  [options]      # same thing
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Cleaning Business CRM & Pricing Calculator.xlsx` | where to write the file |
| `--blank` | off | no example data, just the empty workbook |
| `--no-protect` | off | leave the sheets unlocked (for development) |
| `--password TEXT` | `premium` | sheet protection password |
| `--author TEXT` | `Novality Store` | document author |
| `--business TEXT` | `Bright & Tidy Cleaning Co.` | business name on the cover and the printables |
| `--owner TEXT` | `Dana Whitfield` | the owner / manager |
| `--currency SYMBOL` | `$` | currency symbol, e.g. `£`, `€`, `A$` |
| `--currency-code XXX` | `USD` | three letter code used in headings |
| `--year YYYY` | this year | the year the planner and reports open on |
| `--week-start DAY` | `Monday` | `Monday` or `Sunday` — drives the rota and the calendar |
| `--rate AMOUNT` | 45 | hourly rate charged to the client |
| `--pay AMOUNT` | 18 | hourly pay rate of a cleaner |
| `--minimum AMOUNT` | 90 | minimum charge for any job |
| `--tax PERCENT` | 0 | sales tax / VAT rate, e.g. `20` for 20% |
| `--terms DAYS` | 14 | payment terms |
| `--units {sq ft,sq m}` | `sq ft` | the area unit used for pricing |
| `--distance {miles,km}` | `miles` | the distance unit used for mileage |
| `--rows N` | 600 | rows of capacity in Jobs; every other register scales with it |
| `--today YYYY-MM-DD` | today | pretend it is this date (keeps the demo data reproducible) |
| `--quiet` | off | print only the output path |

```bash
# a blank British workbook for 2027, VAT registered, metric
python build_cleaning_workbook.py --blank --year 2027 \
    --currency "£" --currency-code GBP --tax 20 --rate 38 --pay 13.50 \
    --minimum 75 --units "sq m" --distance km \
    -o "output/Cleaning Business 2027.xlsx"
```

---

## 3. How the code is organised

```
build_cleaning_workbook.py          one line wrapper
cleaning_business_spreadsheet/
    config.py       every default: palette, row counts, rates, factors
    layout.py       SheetSpec / Col and the formula template engine
    formats.py      the whole format book (styles x input/calc/plain)
    writer.py       SheetWriter - banners, nav, sections, KPI tiles, grids
    names.py        the cross sheet defined names
    sample_data.py  one coherent demo company, from a fixed seed
    build.py        tab order, page builders, the workbook assembly
    sheets/
        common.py settings.py lists.py            the setup tabs
        clients.py properties.py leads.py         the work registers
        services.py quotes.py quote_lines.py
        jobs.py job_tasks.py checklists.py
        employees.py timesheets.py job_costing.py the people and money
        invoices.py invoice_lines.py payments.py
        supplies.py supply_usage.py equipment.py
        expenses.py mileage.py complaints.py
        incidents.py safety.py
        start_here.py dashboard.py quote_calc.py  the pages
        calendar_page.py staff_schedule.py
        retention.py reports.py financial.py
        printables.py help_tab.py
tools/              the quality checks
```

A **register** is a data tab. Its module exposes `NAME`, a `spec(cfg)`
that returns a `SheetSpec` (a list of `Col`s with their formulas,
widths, drop downs, comments and conditional formats) and an optional
`decorate(sheet, fmts, cfg, reg)` for the summary strip underneath the
grid. A **page** is a hand-built tab — the dashboard, the quote
calculator, the rota, the printables — written directly with
`SheetWriter`.

### Formula templates

Columns are written once, for one row, with placeholders:

```python
Col("margin", "Margin", 9, "percent",
    formula='=IF(OR({job}{r}="",{revenue}{r}=0),"",'
            '{profit}{r}/{revenue}{r})')
```

* `{r}` — the current row, `{r_prev}` — the row above
* `{key}` — the column letter of another column on the same sheet
* `{first}` / `{last}` — the first and last data row
* `{key_first}` / `{key_last}` — a whole column as `$D$5:$D$604`
* `[[Sheet.key]]` — an absolute range on another sheet,
  `[[Sheet.key#12]]` one cell of it
* helpers for the patterns Excel 2010 makes awkward: `max_if`, `min_if`,
  `blank_if`, `safe`, `lookup`, `min_number`

The registry renders them per row, so a column's logic lives in exactly
one place — and a typo in a sheet or column name is a build error, not
a `#REF!` a customer finds.

### The automation chain

The one thing worth understanding before changing anything:

| Step | Where | What happens |
|------|-------|--------------|
| 1 | `Settings` | your rates, minimum charge, uplifts and margin targets |
| 2 | `Properties.size_hours` → `est_hours` | bedrooms, bathrooms, rooms or area become hours, adjusted for condition |
| 3 | `Services.standard_price` | what that work is worth, and what it costs to deliver |
| 4 | `Quote Calculator` / `Quotes.total` | the quote: lines, discount, deposit, tax, margin |
| 5 | `Jobs.base_price` → `price` | the visit, priced with the frequency factor and the minimum charge |
| 6 | `Timesheets` → `Jobs.actual_hours` | what it really took |
| 7 | `Supply Usage` → `Job Costing.supplies` | what it really used |
| 8 | `Job Costing.profit` → `margin` | revenue minus labour, supplies, travel and extras |
| 9 | `Invoice Lines.price_used` → `Invoices.total` | the bill, built from the jobs |
| 10 | `Payments` → `Invoices.outstanding` | what came back, and what is still late |
| 11 | `Clients.profit` / `value_year` / `risk` | what each client is actually worth |

`tools/validate_cleaning_workbook.py` checks all thirty-five links of
that chain on every build, so a well-meaning edit cannot quietly break
the headline feature.

### Adding something

* **A column** — add a `Col` to that sheet's `spec()`. Widths, formats,
  validation, comments and conditional formatting come with it.
* **A drop down** — add the list to `sheets/lists.py`; it becomes a
  named range automatically and any column can point at it with
  `validate=list_from("MyList")`.
* **A tab** — write the module, add it to `sheets/__init__.py`, then to
  `REGISTER_MODULES` (or `PAGE_BUILDERS`) and `SHEET_ORDER` in
  `build.py`.
* **A "top five" list** — add a hidden rank column to the register, then
  `MATCH(n, rank, 0)` on the page. Never an array formula.
* **A KPI tile** — one `sheet.kpi(row, col, label, formula, style)` call.
* **A line that belongs to a parent** (a quote line, an invoice line, a
  checklist task) — give it the hidden `seq` / `key` pair and the
  printables can pull line *n* with
  `MATCH(ParentID&"|n", key, 0)`.

---

## 4. Quality checks

```bash
python build_cleaning_workbook.py -o output/test.xlsx
python tools/validate_cleaning_workbook.py output/test.xlsx  # needs openpyxl
python tools/check_cleaning_sample.py      output/test.xlsx

python build_cleaning_workbook.py --rows 130 -o /tmp/small.xlsx
python tools/recalc_check.py /tmp/small.xlsx --force          # needs formulas
```

`validate_cleaning_workbook.py` reports: tab order, document author,
sheet protection and passwords, overlapping merges or hyperlinks
trapped inside merges (Excel calls these corrupt), unresolved template
tokens, `#REF!`, **unbalanced brackets** (quote-aware — a doubled `""`
inside a string is not the end of the string), odd quote counts, empty
arguments, unknown or too-modern functions, formulas pointing at a
missing sheet, defined names and drop down sources that do not exist,
column letters left without a row number, formula cells that are not
locked, **`a+b/c` where `(a+b)/c` was meant** (Excel divides before it
adds, so a total built from two `SUMIFS` has to be bracketed before it
is divided - the file opens, the brackets balance, the answer is just
wrong), and the thirty-five links of the automation chain.

The bracket check earns its place: a deeply nested `IF()` built by
string concatenation can lose a closing bracket, and both Excel and
openpyxl will accept the file without a word.

`check_cleaning_sample.py` re-reads the finished file and confirms the
example data joins up — every property points at a real client, every
job at a real client, property, service and cleaner, every invoice line
at a real invoice and a real job, every payment at a real invoice,
every timesheet, mileage trip and complaint at a real employee, every
supply usage at a real item — and that there is enough of it (15
clients, 16 properties, 127 jobs over nine weeks, 25 invoices, 73
invoice lines, 95 timesheets, 68 checklist tasks) for the dashboard to
show something.

`recalc_check.py` goes further and actually *runs* the workbook: it
loads it into the `formulas` calculation engine, evaluates the whole
dependency graph, then reads the answer out of every formula cell — so
a formula that parses perfectly but subtracts a blank on an empty row
is still caught. Check a `--rows 130` build rather than a full-size
one: it still holds the whole example company, the formulas are
identical, there are just fewer spare rows of them, and 110,000 of them
need far more memory than a laptop will give you (the tool says so and
stops unless you pass `--force`). Add
`--show "Dashboard!A8,Dashboard!J12"` to print particular cells - which
is how the average-job-value tile above was caught reading `a+b/c`.

Current status: all three clean and `pyflakes` clean. All 25,652
formulas of the `--rows 130` build were calculated and every one lands
on a real answer - and the answers were read back and sanity checked
against the example company, not just scanned for errors.

---

## 5. What a spreadsheet cannot do

Worth saying plainly, in the listing as well as on the Help tab:

* **it cannot send the reminder for you.** The workbook will tell you,
  the moment you open it, which invoices are overdue, which quotes need
  chasing and which clients have gone quiet — but the email or the text
  message is still you, or an app with a server behind it;
* **the calendar does not sync.** The Job Scheduler and the Calendar
  tab are the diary; Google Calendar, Outlook and the cleaners' phones
  need an integration (or a weekly export) to see the same thing;
* **no e-signatures and no card payments.** A quote can be marked
  accepted and an invoice marked paid, but the signature and the money
  happen somewhere else — Stripe, Square, your bank — and get typed or
  pasted back in;
* **recurring jobs are generated, not watched.** The workbook works out
  the next visit due from the frequency; nobody books it until a person
  (or the fill-down) puts the row in;
* **estimates are estimates.** Hours from bedrooms and square feet are
  a sensible starting point, not a promise — that is exactly why the
  Job Costing tab compares them with the timesheets;
* **this holds personal data.** Client addresses, entry instructions,
  alarm and key details, employee records and financial history all sit
  in one file. Keep it somewhere access-controlled, protect the
  workbook with a real password (not `premium`), do not email it
  around, and delete what you no longer need — your local privacy rules
  apply to a spreadsheet exactly as they do to an app.

What it does give you is a price in under a minute, a diary that costs
itself, invoices that build themselves from the work, and an honest
answer to "which of these clients is actually making me money".

---

## 6. What is deliberately *not* in version 1

The brief listed 720 features across 28 categories; v1 ships the ones a
cleaning business touches every day. These are the natural premium / v2
additions:

* a phone-friendly companion for the cleaners (Google Sheets edition
  with Apps Script, or a Form that appends to Job Tasks with photos);
* macro-driven features — one-click "roll the recurring jobs forward",
  bulk invoice generation, a proper route optimiser;
* live integrations: calendar sync, email and SMS reminders,
  e-signature on the quote, card payment links on the invoice,
  bookkeeping export to Xero or QuickBooks;
* payroll proper — tax tables, holiday accrual and payslips, rather
  than hours, overtime and on-costs;
* GPS or geofenced clock-in instead of typed start and finish times;
* a customer portal and automatic review requests after a clean.

---

© Novality Store. Generated workbooks are licensed for personal use and
for a single business.
