# Ultimate Landscaping & Lawn Care Business Manager — generator

> **This repository holds eleven spreadsheet products** built on the same
> engine, with the same conventions, the same author (Novality Store)
> and the same `premium` sheet protection:
>
> * **Pet Sitting & Dog Walking Business Spreadsheet** —
>   [README.md](README.md), `build_pet_sitting_workbook.py`, 27 tabs.
> * **Social Media Content Calendar & Analytics Planner** —
>   [README_SOCIAL.md](README_SOCIAL.md),
>   `build_social_media_workbook.py`, 43 tabs.
> * **Automated Meal Planner & Grocery List Spreadsheet** —
>   [README_MEAL.md](README_MEAL.md),
>   `build_meal_planner_workbook.py`, 36 tabs.
> * **Cleaning Business CRM & Pricing Calculator** —
>   [README_CLEANING.md](README_CLEANING.md),
>   `build_cleaning_workbook.py`, 37 tabs.
> * **Handmade Product Pricing & Profit Calculator** —
>   [README_HANDMADE.md](README_HANDMADE.md),
>   `build_handmade_workbook.py`, 40 tabs.
> * **Salon Business Management Spreadsheet** —
>   [README_SALON.md](README_SALON.md), `build_salon_workbook.py`,
>   20 tabs.
> * **Daycare & Childcare Business Management Spreadsheet** —
>   [README_DAYCARE.md](README_DAYCARE.md),
>   `build_daycare_workbook.py`, 23 tabs.
> * **Ultimate Airbnb & Short-Term Rental Management Spreadsheet** —
>   [README_AIRBNB.md](README_AIRBNB.md),
>   `build_airbnb_workbook.py`, 20 tabs.
> * **Wedding Photographer Business Manager** —
>   [README_WEDDING.md](README_WEDDING.md),
>   `build_wedding_workbook.py`, 23 tabs.
> * **Ultimate Inventory & Wholesale Order Manager** —
>   [README_INVENTORY.md](README_INVENTORY.md),
>   `build_inventory_workbook.py`, 23 tabs.
> * **Ultimate Landscaping & Lawn Care Business Manager** —
>   this README, `build_landscaping_workbook.py`, 23 tabs.
>
> Every download is listed in one place in
> [DOWNLOADS.md](DOWNLOADS.md).

Python code that builds a finished, ready-to-sell Excel workbook for
anybody who cuts grass for a living: the customers, the properties and
what each one is actually like to work on, the services and what they
really take, the quotes, the work orders, the round, the crew, the
machines, the trucks, the materials, the invoices — and the question
most lawn care spreadsheets cannot answer, **which jobs made money and
which ones only looked like they did**.

Every tab is joined to every other tab by four generated IDs —
**Customer ID, Property ID, Job ID, Invoice ID** — so one visit is one
record from the quote that priced it to the profit line that judges
it.

> **Listing title:** Landscaping Business Spreadsheet | Lawn Care Job
> Scheduler, Quote Calculator, Customer CRM, Crew & Equipment Tracker
> & Profit Dashboard
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_landscaping_workbook.py
# -> output/Ultimate Landscaping & Lawn Care Business Manager.xlsx
#    (23 tabs, 172,546 formulas, ~2.7 MB)
```

The repository also contains the XlsxWriter library itself
(`xlsxwriter/`), so the command above runs with **no installation and
no dependencies** from the repository root. Outside this repo,
`pip install xlsxwriter` first — the generator is tested against the
bundled 0.3.0 and against XlsxWriter 3.2.

---

## 1. What gets built

23 tabs, in this order. Seven are **pages** — designed screens you read
and print. Sixteen are **registers** — a title, a frozen header row,
and hundreds of rows of typed columns and locked formulas.

| # | Tab | Kind | Formulas | What it is for |
|---|-----|------|---------:|----------------|
| 1 | **Dashboard** | page | 201 | The centrepiece. Today's round, the week ahead, the work that slipped, what you are owed, machines due a service, materials running out — and a matrix that counts every alert on sixteen registers. |
| 2 | **Getting Started** | page | 0 | The one thing to understand (the four IDs), the first hour in six steps, the eight workflows, and an honest list of what this is *not*. |
| 3 | **Settings & Lists** | page | 4 | Your business, your labour and charge-out rates, overhead per crew hour, target and floor margins, minimum visit, travel rate, service intervals, season, targets, and the 44 drop-down lists everything else reads. |
| 4 | **Customer CRM** | register | 6,250 | Who they are, how they pay, what they have spent, what they owe, what they are worth over a lifetime — and who has gone quiet. |
| 5 | **Property Database** | register | 6,000 | Where the work happens: lawn area, hedge length, terrain, access, gate code, the dog, drive time and miles, and a difficulty factor every estimate reads. |
| 6 | **Service Catalog** | register | 1,520 | 30 services priced per 1,000 sq ft, per visit, per hour, per linear foot, per cubic yard or per plant, with the minutes each unit really takes and the margin it earns. |
| 7 | **Quote Calculator** | register | 12,000 | Lawn size + labour + materials + equipment + two-way travel + overhead → the price that hits the margin you asked for, with your minimum visit as the floor. |
| 8 | **Job & Work Orders** | register | 73,500 | The spine. Every visit: what was planned, what happened, what it was charged at, what it cost in crew hours and materials, and the gap between estimate and actual. |
| 9 | **Route Planner** | page | 168 | Pick a day and a crew: the stops, gate codes, dog warnings, times and charges on one printable sheet — plus a printable quote and a printable invoice. |
| 10 | **Recurring Contracts** | register | 6,250 | The standing round. Frequency, visit price, annual value, the next visit worked out from the last one, and a renewal warning before it lapses. |
| 11 | **Crew Database** | register | 600 | Pay rates, licences and their expiry, what each person has worked and what they have earned you per hour. |
| 12 | **Crew Schedule** | register | 17,100 | Planned hours against capacity, day by day and person by person, with over-booked days and idle days both flagged. |
| 13 | **Timesheets & Payroll** | register | 17,100 | Hours against a Job ID, overtime split out on a weekly running total, every hour costed at that person's rate — the source of every actual labour figure. |
| 14 | **Equipment Inventory** | register | 2,280 | Mowers, blowers, trimmers: engine hours, service intervals in *hours and months*, maintenance spend, cost per hour and a replacement warning. |
| 15 | **Vehicle & Fuel Tracker** | register | 4,800 | Fuel, repairs, tax and insurance per vehicle, with miles per gallon worked out from the odometer rather than asked for. |
| 16 | **Materials Inventory** | register | 3,600 | Seed, fertiliser, mulch, fuel: stock derived from purchases and usage, reorder triggers, weeks of cover and what a top-up would cost. |
| 17 | **Invoice & Payments** | register | 12,600 | What was billed, what arrived, what is still out — aged from the **due date**, with part payments and chase tracking. |
| 18 | **Expense Tracker** | register | 6,000 | Every cost, split between what belongs to one visit and what belongs to the month, with the material purchases and usage that drive stock. |
| 19 | **Customer Feedback** | register | 1,800 | Compliments, complaints and quality issues scored and attributed to a crew, left open until somebody closes them. |
| 20 | **Seasonal Planner** | page | 13 | The year in four blocks of work, with the lawn maintenance checklist and the property inspection checklist the crew actually uses. |
| 21 | **Profit Analysis** | register | 350 | Fourteen months: revenue, labour, materials, fuel, overheads, profit, margin, profit per crew hour and billable percentage. |
| 22 | **Monthly Reports** | page | 153 | Pick a month. Money in, work done, where the hours went, the best and worst services, against target. |
| 23 | **Annual Reports** | page | 257 | Twelve months in a grid, in-season against out-of-season, the services and customers that earned, and the questions a year of records can answer. |

**172,546 formulas** in total, every one of them locked, on a 2.7 MB
file that opens in Excel 2010.

### The eight workflows it was built around

1. **Customer → property → job.** Add a customer once; their
   properties hang off that Customer ID, and every quote, visit,
   invoice and complaint picks both from a drop-down.
2. **Property → work order.** The lawn area, the difficulty factor,
   the drive time and the gate code travel with the property into
   every estimate and onto the crew's printed sheet.
3. **Quote accepted → work order.** An accepted quote carries its
   price straight through to the job that does the work, and the job
   keeps the Quote ID as the audit trail from price to paid.
4. **Contract → upcoming schedule.** A recurring agreement works out
   its own next visit date from the frequency and the last visit, and
   says so when nothing is booked.
5. **Job completed → crew performance and actual cost.** The hours the
   crew logged and the fuel, materials and disposal booked against the
   Job ID become that visit's real cost, that person's billable
   percentage and that service's average time.
6. **Invoice paid → outstanding balances and revenue.** Payments land
   on the invoice, the customer's balance and the month's collected
   figure at the same moment; nothing is counted twice and nothing is
   typed twice.
7. **Fuel and materials purchased → expense reports.** A purchase is
   one row: it costs the month, it stocks the yard, and — if it
   carries a Job ID — it costs that visit too.
8. **Equipment maintenance due → dashboard alert.** Engine hours since
   the last service and months since the last service, whichever lands
   first, with a warning window before each.

### The four IDs

One decision holds the whole product together: **Customer ID, Property
ID, Job ID and Invoice ID are generated by the workbook and never
typed.** They are minted once — in the Customer CRM, the Property
Database, the Job & Work Orders register and the Invoice & Payments
register — and chosen from a drop-down everywhere else.
`tools/validate_landscaping_workbook.py` checks those joins in both
directions on every build, because the moment an ID can be mistyped
the quote, the visit, the timesheet, the materials and the invoice
stop being the same piece of work.

### The seven things this product is not allowed to get wrong

Each of these is a way a lawn care spreadsheet quietly starts lying to
its owner, and each one is checked against the real formulas on every
build:

1. **A quote is built, not guessed.** Lawn size from the property,
   minutes from the catalogue, labour at your rate and crew size,
   materials at your markup, equipment, travel both ways, overhead per
   crew hour — then the price worked back off the margin you asked
   for, never under your minimum visit.
2. **Every job is costed twice.** Estimated from the catalogue and the
   property's difficulty; actual from the timesheets and the expenses.
   Both are shown, and so is the gap.
3. **Recurring contracts schedule themselves.** The next visit follows
   from the frequency and the last visit, and a contract with nothing
   booked says so.
4. **Equipment warns on hours *and* on months.** A mower that sat all
   winter is as due a service as one that ran all summer.
5. **Material stock is derived, never typed.** Opening stock, plus
   what was purchased, minus what was used on jobs — every term a row
   somebody entered on the Expense Tracker.
6. **Job money is derived, never typed.** The charge comes from the
   quote, the contract or the catalogue, in that order; the cost comes
   from the hours and the receipts.
7. **The route planner reads the job register.** It is a view of one
   crew's day, not a second diary that can disagree with the first.

---

## 1b. Download it

You do not have to run anything. The finished workbooks are in
[`download/`](download/) and indexed in [DOWNLOADS.md](DOWNLOADS.md):

| File | What it is |
|------|------------|
| `landscaping-lawn-care-business-manager.xlsx` | Filled in with a worked example — one grounds care business, mid-season, mid-mess. |
| `landscaping-lawn-care-business-manager-blank.xlsx` | Empty, ready for your own round. |
| `landscaping-lawn-care-business-manager-source.zip` | The generator source. |

### The example business

The sample file is Thornhill Grounds Care of Granton, Ohio: Marcus
Thornhill, eleven people on the books, three crews, 42 customers, 58
properties and six months of history — 791 visits, 256 invoices, 448
timesheet lines and 38 standing contracts. The season it is having is
deliberately not a tidy one:

* a fall cleanup **quoted at four hours that took eleven and a half**,
  because the estimate was made from a photograph — the job-level
  variance column exists for exactly this;
* **26 invoices overdue** and one part paid, aged from the due date,
  with the oldest past sixty days and flagged as a decision rather
  than a chase;
* **12 visits rained off and 7 needing a revisit**, waiting to be
  caught up — the workbook has no weather feed and does not pretend to;
* a mower **out of service** and another **past its service hours**,
  both of them on the dashboard before a crew finds out the hard way;
* **four materials at or below their reorder trigger**, with the
  top-up value worked out;
* a customer who has **gone dormant** after three years of regular
  work, and a lead that was quoted and never followed up;
* **13 quotes expired** without a decision and ten lost, with the
  reasons recorded, which is worth more after six months than any
  advice about pricing;
* five complaints and five quality issues against 45 compliments, each
  scored and attributed to a crew;
* a contract **due renewal** inside the warning window, and one on a
  seasonal break.

A sample company where everything reconciles teaches nobody anything.
This one teaches you what each alert column is for.

---

## 2. Command line

```bash
python build_landscaping_workbook.py --help
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Ultimate Landscaping & Lawn Care Business Manager.xlsx` | Where to write the file. |
| `--blank` | off | No example data — just the empty workbook. |
| `--rows N` | scaled | Rows of capacity in the big registers. |
| `--today YYYY-MM-DD` | today | Pretend it is this date, for the example. |
| `--quiet` | off | No progress output. |

```bash
# the empty workbook, ready to sell
python build_landscaping_workbook.py --blank \
  -o "download/landscaping-lawn-care-business-manager-blank.xlsx"

# a small file for testing - every register scaled down
python build_landscaping_workbook.py --rows 20 -o /tmp/small.xlsx
```

Everything else is a setting *inside* the workbook, not a command line
flag: the business name and licence number, the labour rate, the
charge-out rate, overhead per crew hour, the target and minimum
margins, the minimum visit charge, the material markup, the travel
rate per mile, the drive time between stops, mowing minutes per 1,000
sq ft, quote validity, the service intervals in hours and in months,
the season months, the dormant-customer window, the tax rate and all
44 drop-down lists are on **Settings & Lists**. You should not need to
edit the Python to make this fit a business.

---

## 3. How the code is organised

```
build_landscaping_workbook.py        the one-liner entry point
landscaping_spreadsheet/
    cli.py           argument parsing, the summary it prints
    config.py        every rate, status set, threshold and disclaimer
    layout.py        Col / SheetSpec / the {token} template engine
    formats.py       every cell format, locked and unlocked
    writer.py        SheetWriter: banners, KPIs, grids, protection
    names.py         551 dynamic named ranges that grow with the data
    sample_data.py   the worked example, checked twice before it ships
    build.py         the 23 tab order and the seven page builders
    sheets/
        common.py        the column factories every register shares
        customers.py     ... and the sixteen registers
        properties.py
        ...
        dashboard.py     ... and the seven pages
        getting_started.py
        settings.py
        route.py
        seasonal.py
        monthly.py
        annual.py
```

A **register** module is a `NAME`, a `spec(cfg)` returning a
`SheetSpec`, and optionally a `decorate(...)`. A **page** module is a
`build(book, fmts, cfg, reg)`. Nothing else — the engine does the rest.

### Two rules the engine enforces

* **Formula cells are locked, input cells are not.** A column with a
  `formula=` is written with the locked format; a column without one is
  written unlocked. There is no third case, and the validator fails the
  build if a formula cell is ever writable.
* **Cross-sheet references go through `[[Sheet.column]]`.** No formula
  anywhere names a cell range by hand. Adding a column in the middle of
  a register cannot break another tab, because the reference is
  resolved from the specs at build time — and a reference to a column
  that does not exist is a build error, not a `#REF!` a buyer finds.

---

## 4. Quality checks

Nothing ships until all six pass:

```bash
pyflakes landscaping_spreadsheet/ tools/
python tools/check_landscaping_specs.py        # before a workbook exists
python tools/check_landscaping_sample.py       # the worked example
python build_landscaping_workbook.py -o /tmp/land.xlsx
python tools/check_parens.py /tmp/land.xlsx
python tools/validate_landscaping_workbook.py /tmp/land.xlsx
python tools/recalc_check.py /tmp/land-small.xlsx   # really evaluates it
```

**`check_landscaping_specs.py`** reads the 16 register specifications
without building anything and catches, in under a second, the mistakes
that are expensive to find later: brackets that do not balance, a
`{token}` naming no column on that sheet, a `[[Sheet.column]]`
pointing at a column that does not exist, `a+b/c` where `(a+b)/c` was
meant, and a COUNTIFS criteria range carrying its own test. It reports
611 columns, 345 formula columns and 270 cross-sheet references.

**`check_landscaping_sample.py`** checks the worked example against the
specs: every key is an *input* column (a value written into a formula
column is dropped silently and takes the formula with it), every
drop-down value exists in its list, every ID resolves to a real record
on the register that mints it — and then it replays the story the data
tells. Stock never goes negative at any point in the timeline,
odometers only ever go up, no invoice is paid for more than it was
raised for, every completed job has hours against it, there is a round
booked for today, and the overdue invoices, the machine needing a
service and the dormant customer are all really there.

**`validate_landscaping_workbook.py`** opens the finished file and
checks 100 links in the automation chain, the Customer, Property, Job
and Invoice joins in both directions, the seven technical decisions
above, sheet protection and cell locking, every defined name a
drop-down reads, and that no cell anywhere claims GPS navigation,
route optimisation or a connection to accounting, payroll or booking
software.

**`recalc_check.py`** really evaluates every formula with the
`formulas` engine and fails on any `#VALUE!`, `#REF!` or `#DIV/0!`. It
runs on the sample *and* the blank file, because a formula that works
with data in it and breaks without is still broken.

---

## 5. What a spreadsheet cannot do

Said plainly on the Getting Started tab, and enforced by the validator:

* **The route planner is not GPS navigation and does not optimise a
  route.** It puts the stops *you* have chosen, in the order *you* have
  chosen them, onto one printable sheet with the gate codes and the
  dog warnings. It will not tell you that Mill Lane before Oak Avenue
  saves eleven minutes. You change the order by changing the start
  times on the job rows.
* **It does not sync with accounting, payroll or booking software.**
  Not QuickBooks, not Xero, not Jobber, not Yardbook. You enter what
  happened and the formulas do the arithmetic.
* **It cannot take a card payment or send an invoice.** It produces a
  printable invoice; sending and collecting are yours.
* **Timesheets are not a payroll filing.** They work out gross labour
  cost so you can price work properly. Tax, national insurance,
  workers' compensation and overtime law vary by state and country —
  give the totals to whoever runs your payroll.
* **There is no weather feed.** You mark a job rained off; the
  workbook then tells you what has to be caught up.
* **Photos are links, not attachments.** A spreadsheet holding
  hundreds of site photographs becomes a spreadsheet nobody can open.
* **It is not multi-user in real time.** One person has the file open
  at a time, or it lives on a drive that syncs.
* **It is not accounting or tax advice.** The tax category column
  organises costs for whoever does your books.

---

## 6. What is deliberately *not* in version 1

* **Automatic route optimisation.** Genuinely useful, genuinely not a
  spreadsheet feature: it needs a distance matrix and a solver, and
  anything short of that is a guess dressed up as an answer.
* **Per-visit GPS or crew clock-in from a phone.** Timesheets are
  entered in the evening or copied off the printed day sheet.
* **A customer portal.** No logins, no online booking, no payment
  links.
* **Equipment depreciation schedules for the accounts.** There is a
  straight-line life and a cost per hour for pricing decisions, which
  is not the same thing as a tax depreciation method.
* **Chemical application records to a regulator's format.** The
  licence and the treatment are recorded; a state-specific pesticide
  log is a different document.

---

## 7. The other landscaping products

The same engine, re-pointed, gives six closely related products that
share this one's spine: a **Lawn Care Business Spreadsheet** trimmed
to mowing rounds, a **Landscaping Contractor Project Manager**
weighted towards installs and phases, a **Gardening Business
Tracker**, a **Snow Removal Business Manager** with per-storm
callouts, a **Tree Service Business Manager** with crew certification
and equipment hire, and a **Pool Cleaning Service Manager** with
chemical readings. Each is the registers in `sheets/` with different
emphasis, not a rewrite.
