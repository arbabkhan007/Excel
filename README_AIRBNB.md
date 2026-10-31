# Ultimate Airbnb & Short-Term Rental Management Spreadsheet — generator

> **This repository holds eight spreadsheet products** built on the same
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
>   this README, `build_airbnb_workbook.py`, 20 tabs.
>
> Every download is listed in one place in
> [DOWNLOADS.md](DOWNLOADS.md).

Python code that builds a finished, ready-to-sell Excel workbook for an
Airbnb, Vrbo or Booking.com host: the bookings, the calendar, the
turnovers, the money and the question nobody can usually answer —
**which property is actually making anything**. Multiple properties and
multiple platforms from the first row. Every tab is joined to every
other tab by a **Property ID** and a **Booking ID**, so adding a fifth
property is typing a row, not rebuilding the file.

> **Listing title:** Airbnb Host Spreadsheet | Short-Term Rental
> Management | Booking Calendar, Cleaning, Guest Tracker, Expenses &
> Profit Dashboard
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_airbnb_workbook.py
# -> output/Airbnb & Short-Term Rental Management Spreadsheet.xlsx
#    (20 tabs, ~2.0 MB)
```

The repository also contains the XlsxWriter library itself
(`xlsxwriter/`), so the command above runs with **no installation and
no dependencies** from the repository root. Outside this repo,
`pip install xlsxwriter` first — the generator is tested against the
bundled 0.3.0 and against XlsxWriter 3.2.

---

## 1. What gets built

20 tabs, 138,600+ formulas, all joined up and all driven by drop downs.

| # | Tab | What it does |
|---|-----|--------------|
| 1 | **Dashboard** | The whole portfolio on one screen. A dated header, then **tonight and the next few days** — guests in tonight, arriving today, leaving today, same-day turnarounds, properties empty, cleans outstanding — **this month** — nights booked, occupancy against target, accommodation revenue, ADR, RevPAR, platform fees, costs, profit, margin — **money still moving** — payouts expected, late, unpaid balances, deposits held — and **what needs doing**, a list that counts the real problems rather than describing features: double bookings, unassigned cleans, arrivals with no access code sent, supplies out of stock, urgent repairs, reviews worth chasing. Below that, **every property side by side** on thirteen measures, and **the last twelve months** of revenue, costs and profit with sparklines |
| 2 | **Getting Started** | The eleven step set-up in order, the six ideas the workbook is built on (Property ID, Booking ID, status, accommodation vs cleaning, blocked nights, counting a cost once), a map of all 20 tabs, eight live counters that prove the file is working, the colour key, the privacy note and the tax disclaimer |
| 3 | **Settings & Lists** | Your business name and contact details; **the platform fee table — platform, commission, payment fee, is it an OTA — eight rows and two spares, every cell editable**, because 3% is Airbnb's host-only split and almost nobody reading this pays exactly that; pricing defaults — nightly rate, cleaning fee, what a turnover actually costs you, extra guest threshold and fee, pet fee, deposit, minimum stay, co-host share; lodging tax, its name and **who remits it**; turnover costs — laundry, consumables, welcome basket, cleaner rate, what counts as a tight turnaround; two switches that stop a cost being counted twice; eight targets; six alert windows; twelve ID prefixes; 30 drop-down lists; and 24 months of month keys |
| 4 | **Property Database** | 63 columns, one property a row. Automatic Property ID, nickname, address, type, bedrooms, bathrooms, beds, maximum guests, status, owner, manager, **ownership — owned, mortgaged, rental arbitrage, co-hosted, head lease** — main platform, listing URL, licence number, check-in and check-out times, minimum stay, Wi-Fi, access method, and the money: nightly rate, cleaning fee charged, cleaning cost paid, deposit, monthly rent, monthly mortgage, purchase price. Then everything the other tabs know about it — bookings, nights available, nights booked, nights blocked, **occupancy, ADR, RevPAR, average length of stay**, review score, accommodation revenue, gross, platform fees, net revenue, cleaning spend, maintenance spend, other expenses, total costs, **profit, margin, profit per night, yield on purchase price, and its rank in the portfolio** |
| 5 | **Booking Tracker** | 60 columns, 900 rows, the spine of the file. Property, guest, platform, reference, status, booked date, check-in, check-out, adults, children, pets, the rate and fees if you are overriding them, discount, payment status, requests. Then the automation: **nights**, month and season, lead time, total guests, **the nightly rate actually used** (yours if typed, the property's if not), accommodation revenue, cleaning fee, extra guest fee, gross, lodging tax, **the commission and payment-fee percentages looked up from the platform table**, platform fees, **net revenue**, net per night, length-of-stay band, whether it holds the calendar, whether it earns, **a double-booking flag that reads every other row on the same property**, a same-day turnaround flag, where the stay is up to today, its place in the arrivals queue, and an action column |
| 6 | **Booking Calendar** | Pick a property and a month from two drop downs and the whole page redraws. Eight tiles — nights booked, free, blocked, occupancy, revenue, arrivals, departures, turnarounds — then a **six-week month grid**, one cell per day, each showing the date, the guest and what is happening, colour coded for occupied, arriving, departing, same-day turnaround, maintenance and owner block. Below it a **fourteen day strip across up to twelve properties at once** — one letter per property per day, O A D T M B — which is the view that catches the gap nobody noticed. Then the **Turnover Command Centre**: today's check-outs, tomorrow's arrivals, who is cleaning what, which turnovers are unfinished and which properties are guest-ready right now |
| 7 | **Guest Database** | One row per guest, not per booking. Name, email, phone, country, type, trip purpose, the platform they first came through, marketing consent, review status, preferences, requests, feedback, your private note. Then their history: stays, nights, **repeat guest or not**, first and last stay, gross and net revenue, average stay, average rating, cancellations, and what they are worth to you. **No ID documents, no passport numbers, no card details** — the workbook has nowhere to put them on purpose |
| 8 | **Check-in & Check-out** | One row per arrival. Arrival and departure times, late check-out, early check-in, how the keys change hands, then the nine things that have to happen — **directions sent, access code sent, Wi-Fi sent, house rules sent, welcome message sent, check-out instructions sent, checked out, inspected, photographed** — with the property and dates pulled from the Booking ID, a pre-arrival readiness percentage, a steps-done count, what is still outstanding and where the stay has got to |
| 9 | **Cleaning Schedule** | 45 columns. **A job is generated from a check-out, not typed from memory** — give it the Booking ID and the property, date, guest and the next arrival all arrive with it. Cleaner, start and end time, status, then the **nine room tasks** — bedrooms, bathrooms, kitchen, living areas, floors, laundry, rubbish, restock, final inspection — hours worked, completion percentage, **the gap to the next guest in days and in hours**, a same-day flag, whether the property is guest-ready, cleaner cost, laundry, consumables, **total cost against the cleaning fee you charged** and the difference, paid or not, and an action column |
| 10 | **Turnover Checklist** | The printable one. Type a cleaning job ID at the top and the card fills in — property, guest, check-out, next arrival, cleaner, and a red banner if it is a same-day turnaround. Then **five grouped checklists** — bedroom, bathroom, kitchen, living areas, final inspection — each item with a tick box, a live completion percentage that turns green at 100%, nine report-back flags that write straight back to the schedule, and a sign-off line. One page, fits on A4, designed to be printed and handed over |
| 11 | **Platform Revenue & Fees** | One row per payout. Booking, payout date, amount received, the fee the platform actually took, refunds, adjustments, status, reference. Against that the workbook puts **what it expected** — expected gross, expected fee, expected net from the fee table — and then the number that matters: **the variance, in money and in percent, and the effective fee rate the platform really charged you**. Days to pay, and a chase alert when a payout is later than your own threshold |
| 12 | **Expense Tracker** | 28 categories including **mortgage interest and rent**, so an arbitrage operator and an outright owner can both use it honestly. Date, property, category, vendor, description, amount, tax amount, method, receipt, **recurring frequency**, the Booking ID if it belongs to one stay, and a tax category. Then: monthly equivalent of any recurring cost, whether it is cleaning or maintenance, whether it is capital rather than an expense, **whether it belongs in the P&L at all** — the two switches on Settings stop a cleaning or maintenance cost being counted twice — and the amount that actually reaches the profit figure |
| 13 | **Supplies Inventory** | Per property or shared. Item, category, unit, supplier, SKU, on hand, minimum level, last counted, last restocked, **how much is used per turnover**, unit cost, reorder quantity. Then: how old the count is, **turnovers this month from the cleaning schedule**, usage, weeks of cover left, stock value, what an order would cost, cost per turnover, whether it needs ordering and a status — out of stock, order now, low, in stock, overstocked |
| 14 | **Maintenance Log** | Report date, property, category, description, **priority**, who reported it, the Booking ID if a guest did, contractor, scheduled and completed dates, status, **whether it blocks the calendar**, estimated cost, actual cost, invoice reference, paid, **warranty expiry**. Then: the cost in force, the overrun against the estimate, whether it is still open, days open, days until it is due, and an alert that escalates with the priority |
| 15 | **Reviews Tracker** | One row per review with **seven sub-ratings** — overall, cleanliness, accuracy, check-in, communication, location, value — plus requested date, status, the public review, your private note, whether you responded, when, and what you changed as a result. Then: days to review, the detail average, **your weakest category on that stay**, whether it is below your target, and a chase alert. Averages roll up to the property and to the platform |
| 16 | **Staff & Vendors** | Cleaners, contractors, laundry, managers, gardeners, pest control. Role, company, phone, email, which properties, availability, **pay basis — per turnover, per hour, per job, retainer, per property per month, percentage of revenue** — rate, contract status, start date, insured, **insurance expiry**, keyholder. Then: cover days, cleans assigned, cleans completed, **their average completion score and average hours**, open jobs, total paid, what is owed, and their rank by spend |
| 17 | **Tax Summary** | Type a year. Income — nights, accommodation, cleaning fees collected, extras, gross, platform fees withheld, **net taxable income** — then costs **by tax heading** (operating, capital, financing, owner drawing, not deductible, review with accountant) and **by all 28 categories**, the bottom line, and the same thing per property. Carries a disclaimer in writing: **what is deductible, how a cost is classified and whether a capital item is depreciated depend entirely on your jurisdiction and your circumstances** |
| 18 | **Occupancy & Performance** | One row per property per month, 400 rows. Nights available, booked, blocked, **occupancy, ADR, RevPAR, average length of stay, cancellation rate, cleaning cost per booking, operating cost per night**, bookings, guests, accommodation, gross, net, cleaning, maintenance, expenses, total cost, profit, margin and the review score — the tab an investor asks for |
| 19 | **Monthly Reports** | Pick a month key. Ten KPI tiles, then **the same fifteen measures against last month and against the same month last year** with the change in both, where the bookings came from by platform, every property on ten measures, what the month cost by heading, and how the month ran — arrivals, turnovers, tight turnarounds, average gap, cleaning completion, repairs raised and closed, reviews and the average score |
| 20 | **Annual Reports** | Pick a year. Fourteen KPI tiles, **twelve months across and twelve measures down** with a year column and sparklines, the four seasons (winter wraps December into January properly), eleven measures against last year, every property on eleven measures, and fourteen questions worth asking in January — your best and worst month, your best property, the month occupancy peaked, what the platforms took, what cleaning cost you, and which property you should think hardest about |

### The sums that matter

Every one of these is written out on the Getting Started tab too,
because a buyer should be able to check the arithmetic rather than
trust it.

| Figure | How it is worked out |
|--------|----------------------|
| Nights | check-out − check-in (the night of departure is not a night) |
| Accommodation revenue | nightly rate × nights, the rate you typed or the property's default |
| Extra guest fee | guests over the threshold × the fee × nights |
| Gross | accommodation + cleaning fee + extra guest fee + pet fee − discount |
| Platform fee | gross × the commission for that platform, **read from the editable table** |
| Payment fee | gross × the payment-processing rate for that platform |
| Net revenue | gross − platform fee − payment fee − lodging tax |
| **ADR** | **accommodation revenue ÷ nights booked — cleaning fees excluded** |
| **RevPAR** | **accommodation revenue ÷ nights available — cleaning fees excluded** |
| Occupancy | nights booked ÷ nights available, blocked nights removed from both |
| ALOS | nights ÷ bookings |
| Payout variance | what the platform paid − what the fee table said it would |
| Effective fee rate | the fee the platform actually took ÷ gross |
| True profit per booking | net revenue − cleaning − laundry − supplies − the booking's share of fixed costs |
| Property profit | net revenue − cleaning − maintenance − expenses allocated to it |
| Monthly equivalent | a recurring cost converted: annual ÷ 12, quarterly ÷ 3, weekly × 52 ÷ 12 |
| Weeks of cover | stock on hand ÷ (usage a month ÷ 4.33) |
| Days until a repair is due | scheduled date − today, escalated by priority |

### Rules baked in

* **Property ID and Booking ID join everything.** Not property names,
  which get retyped, abbreviated and misspelled. Every lookup in all
  twenty tabs goes through an ID, which is why the file works the same
  with one property or twelve.
* **The platform fee table is a setting, not a rule.** Eight platforms
  ship filled in — Airbnb, Vrbo, Booking.com, Expedia, direct,
  returning guest, corporate, other — each with a commission and a
  payment-processing rate, and two spare rows. Every one of those
  numbers is a cell you type over, because commission varies by
  platform, by market, by account type and by whatever you negotiated.
  Nothing anywhere is hard coded to 3%.
* **ADR and RevPAR exclude the cleaning fee.** This is the most common
  error in rental spreadsheets sold online, and it flatters both
  numbers. Accommodation revenue is kept as its own column from the
  very first calculation so that it cannot leak in later. The
  validator fails the build if either formula so much as mentions
  cleaning.
* **Blocked nights are bookings.** An owner stay or a maintenance
  closure is a row on the Booking Tracker with the status `Owner
  block` or `Maintenance block`. It holds the calendar, it shows on
  the grid, it earns nothing, and it comes out of the occupancy
  denominator — which is the honest way to do it, and it means there
  is no second calendar to keep in step.
* **A cleaning job comes from a check-out.** Type the Booking ID and
  the property, the date, the guest and the next arrival all follow.
  There is no integration claimed and none needed.
* **A cost is counted once.** Cleaning spend comes from the cleaning
  schedule and maintenance from the maintenance log. The same cost
  typed into Expenses is flagged and excluded — unless you flip
  `CleaningInExpenses` or `MaintenanceInExpenses` to Yes on Settings,
  in which case the other tab stops feeding the profit instead.
* **Revenue and cash are never mixed.** Net revenue is what the
  booking earned. A payout is what arrived in the bank. Both are
  tracked, neither is derived from the other, and the gap between them
  is a column called variance.
* **No guest identity documents and no card data.** There is nowhere
  in the Guest Database to put a passport number or a card, by design,
  and the workbook says so in writing on Getting Started.
* **A blank row costs nothing.** Every formula starts with a guard on
  its anchor column, every division has a floor, every lookup is
  coerced, and the whole workbook was recalculated empty as well as
  full to prove it.
* **No charts, no macros, no array entry.** Sparklines and conditional
  formatting only, so the file opens and recalculates everywhere,
  including Google Sheets.

---

## 1b. Download it

Everything is committed to the repository, so these links work straight
from GitHub — right click, *Save link as*.

| What | Size | Link |
|------|-----:|------|
| **The finished workbook, with the example portfolio filled in** | 1,988 KB | [airbnb-rental-spreadsheet.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/airbnb-rental-spreadsheet.xlsx) |
| **The finished workbook, empty and ready to sell** | 1,963 KB | [airbnb-rental-spreadsheet-blank.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/airbnb-rental-spreadsheet-blank.xlsx) |
| **The generator source, zipped** | 332 KB | [airbnb-rental-spreadsheet-source.zip](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/airbnb-rental-spreadsheet-source.zip) |

The example host is **Harbour & Pine Stays** of Portland, Oregon, with
five properties: Harbour View Loft, a one-bedroom owned outright that
turns over constantly; Pine Ridge Cabin, mortgaged, sleeps eight, and
**its boiler failed six days ago** — there is a maintenance block on
the calendar, a $1,180 invoice, a plumber, and a partial refund of
$395 to the guest whose stay it ruined; Old Mill Townhouse, run on
**rental arbitrage** at $1,950 a month, which is why the rent category
exists; Creekside Bungalow, the newest; and Fir Street Apartment,
status **Being set up** — it has $4,530 of furniture and appliances
against it, no bookings, and it earns nothing, which is exactly what a
property in that state should look like on the dashboard.

Forty-six bookings run from four months ago to a month out across
Airbnb, Vrbo, Booking.com, Expedia, direct and corporate — with two
cancellations, two live inquiries, two guests checked in right now,
an owner block and the six-night maintenance block. Twelve guests,
four of them repeat. Thirty-three cleans, eight vendors, thirty
payouts, seventy-five expense rows over four months, twenty supply
lines, twelve repairs and twenty-five reviews.

The demo is deliberately imperfect, because a tidy demo hides the
features: a vacuum broke and **two reviews mention cleanliness**;
coffee pods are at **zero on hand** and the restock was never done;
**seven guests arrive tomorrow and one has not been sent the access
code**; one payout is late enough to be chased; two Expedia payouts
came in under the fee table and the variance column says so; and one
contractor's **insurance expires in 24 days**. Every alert column has
something real in it on first open. `--blank` strips all of it out and
leaves the formulas.

Rebuild the three files with:

```bash
python build_airbnb_workbook.py -o download/airbnb-rental-spreadsheet.xlsx
python build_airbnb_workbook.py --blank -o download/airbnb-rental-spreadsheet-blank.xlsx
python tools/export_source_md.py --product airbnb
python tools/make_download_zip.py --product airbnb
```

---

## 2. Command line

```bash
python build_airbnb_workbook.py        [options]
python -m airbnb_rental_spreadsheet    [options]      # same thing
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Airbnb & Short-Term Rental Management Spreadsheet.xlsx` | where to write the file |
| `--blank` | off | no example data, just the empty workbook |
| `--no-protect` | off | leave the sheets unlocked (for development) |
| `--password TEXT` | `premium` | sheet protection password |
| `--author TEXT` | `Novality Store` | document author |
| `--business TEXT` | `Harbour & Pine Stays` | your business name on the cover and the reports |
| `--owner TEXT` | `Nadia Whitlock` | the host or owner |
| `--currency SYMBOL` | `$` | currency symbol, e.g. `£`, `€`, `A$` |
| `--currency-code XXX` | `USD` | three letter code used in headings |
| `--year YYYY` | this year | the year the reports open on |
| `--nightly AMOUNT` | 185 | the default nightly rate |
| `--cleaning-fee AMOUNT` | 95 | the cleaning fee you charge a guest |
| `--cleaning-cost AMOUNT` | 65 | what a turnover actually costs you |
| `--deposit AMOUNT` | 300 | the security deposit |
| `--min-nights N` | 2 | the default minimum stay |
| `--tax PERCENT` | 0 | lodging / occupancy tax rate |
| `--tax-self` | off | you remit the lodging tax, not the platform |
| `--occupancy-target PERCENT` | 72 | the occupancy you are aiming at |
| `--adr-target AMOUNT` | 195 | the average daily rate you are aiming at |
| `--revenue-target AMOUNT` | 14,000 | monthly revenue target |
| `--properties N` | 4 | properties the targets are scaled for |
| `--rows N` | 900 | rows of capacity in the Booking Tracker; every other register scales with it |
| `--today YYYY-MM-DD` | today | pretend it is this date (keeps the demo data reproducible) |
| `--quiet` | off | print only the output path |

```bash
# a blank UK portfolio: eight properties, VAT-registered, pounds
python build_airbnb_workbook.py --blank --currency "£" --currency-code GBP \
    --business "Kestrel Coast Lets" --owner "Rhian Morgan" \
    --properties 8 --nightly 145 --cleaning-fee 75 --cleaning-cost 55 \
    --min-nights 3 --tax 20 --tax-self --occupancy-target 68 \
    --adr-target 160 --revenue-target 22000 --rows 1200 \
    -o "output/Kestrel Coast GBP.xlsx"
```

---

## 3. How the code is organised

```
build_airbnb_workbook.py           the launcher
airbnb_rental_spreadsheet/
    __init__.py  __main__.py  cli.py     entry points and the options
    config.py                            every default, list, platform
                                         rate and palette
    layout.py                            Col / SheetSpec / Registry, the
                                         formula template engine
    formats.py                           every cell style, once
    writer.py                            SheetWriter: banners, sections,
                                         KPI tiles, tables, grids
    names.py                             the 42 live OFFSET lookup names
    build.py                             tab order and the build pipeline
    sample_data.py                       Harbour & Pine Stays, four
                                         months of it
    sheets/
        common.py                        shared columns, conditional
                                         formats, formula fragments
        settings.py                      the settings and lists tab
        properties.py  bookings.py  guests.py  arrivals.py
        cleaning.py  payouts.py  expenses.py  supplies.py
        maintenance.py  reviews.py  vendors.py  performance.py
                                         the twelve registers
        dashboard.py  getting_started.py  calendar.py  turnover.py
        tax.py  monthly.py  annual.py    the summary pages
```

A **register** is a table: one `spec(cfg)` function returning a
`SheetSpec` of `Col` objects, and the engine draws it, validates it,
formats it, protects it and fills it with example data. A **page** is
laid out cell by cell with `SheetWriter`. That is the whole
architecture, and it is why a twenty-tab workbook is about 11,000
lines rather than 60,000.

### Formula templates

A column's formula is written once, as a template, against named
columns rather than letters:

```python
Col("net_revenue", "Net revenue", style="money",
    formula='=IF({property}{r}="","",'
            '{gross}{r}-{platform_fee_total}{r}-{tax}{r})')
```

`{gross}` becomes the column letter, `{r}` becomes the row. Insert a
column anywhere and every formula that mentions it follows, because
nothing refers to `K` — it refers to `gross`. Across sheets the token
is `[[Booking Tracker.net_revenue]]`, rendered to
`'Booking Tracker'!$AI$5:$AI$904`.

Page modules must write formulas through `s.formula(row, col, text)`,
never `ws.write_formula` — `s.formula` is what renders the `[[...]]`
tokens, and a raw call writes the literal token text into the file.

### The chain

```
Settings & Lists
   platform fee table, pricing defaults, targets, alert windows
        |
Property Database  --- Property ID --->  every other tab
        |
Booking Tracker  --- Booking ID --->  arrivals, cleaning, payouts,
        |                             reviews, expenses, maintenance
        |-- nights, accommodation, gross, platform fee, net revenue
        |-- double-booking flag, same-day flag, arrival queue
        |
        +--> Booking Calendar      the month grid and the 12-property strip
        +--> Check-in & Check-out  the nine steps before a guest arrives
        +--> Cleaning Schedule --> Turnover Checklist (printable)
        +--> Platform Revenue & Fees   expected vs actual payout
        |
Expense Tracker + Cleaning + Maintenance
        |
        +--> Property Database profit columns
        +--> Occupancy & Performance   property x month
        +--> Tax Summary / Monthly Reports / Annual Reports / Dashboard
```

### Adding something

A new column on an existing register is one `Col(...)` in that
module's `spec()`. A new register is a new module in `sheets/` with a
`NAME` and a `spec(cfg)`, added to `REGISTER_MODULES` and
`SHEET_ORDER` in `build.py`. A new drop-down list is a list in
`config.py` and one line in `settings.LISTS`. Nothing else has to
change — the names, the validation, the protection and the sample data
hook themselves up.

---

## 4. Quality checks

```bash
python build_airbnb_workbook.py -o output/test.xlsx
python tools/check_parens.py              output/test.xlsx
python tools/validate_airbnb_workbook.py  output/test.xlsx  # needs openpyxl
python tools/check_airbnb_sample.py

python build_airbnb_workbook.py --rows 130 -o /tmp/small.xlsx
python build_airbnb_workbook.py --rows 130 --blank -o /tmp/blank.xlsx
python tools/recalc_check.py /tmp/small.xlsx    # needs formulas
python tools/recalc_check.py /tmp/blank.xlsx
```

`check_parens.py` walks the saved XML and balances the brackets and
quotes of all 138,641 formulas — concatenated nested `IF`s are the one
mistake a generator makes that no amount of reading catches — and
fails on any template placeholder or sheet token left behind.

`validate_airbnb_workbook.py` reports: tab order, document author,
sheet protection and passwords, unresolved tokens, `#REF!`, unbalanced
brackets, odd quote counts, empty arguments, unknown or too-modern
functions (`XLOOKUP`, `IFS`, `MAXIFS`, dynamic arrays — none are used),
formulas pointing at a missing sheet, defined names and drop-down
sources that do not exist, column letters left without a row number,
formula cells that are not locked, **`a+b/c` where `(a+b)/c` was
meant**, **`COUNTIFS` and `SUMIFS` whose criteria ranges are different
sizes** (Excel answers `#VALUE!` and nothing else warns you), the 81
links of the automation chain, the 55 named cells the pages depend on,
the 36 drop-down lists that must exist, and three promises the
listing makes in writing: that **every cell of the platform fee table
is still unlocked**, that **neither ADR nor RevPAR mentions cleaning**,
and that **nothing in the file claims to synchronise with a booking
platform**.

`check_airbnb_sample.py` reads the example data before it is written
and fails if any value is sent to a column that holds a formula — it
would be silently dropped — or if any value that feeds a drop down is
not in that drop down's list, which is how a `SUMIFS` ends up
returning zero for a reason nobody can see.

`recalc_check.py` loads the workbook into the `formulas` engine and
evaluates the whole dependency graph, then checks every formula cell
for an error value. It is run on **both** the sample and the blank
build, because the empty workbook is the one that exposes a division
by zero and the half-full one is where a cancelled booking with no
check-out date turns a nights column to `#VALUE!` — a buyer's first
ten minutes are spent in exactly those files.

Current state: 20 tabs, **138,641 formulas**, 147 defined names,
3,898 unlocked cells in the sample build and 414 in the blank one, no
problems from any of the four tools, and both recalculation runs —
sample and blank, 22,315 formulas each — finish with zero error
cells.

---

## 5. What a spreadsheet cannot do

**It does not connect to Airbnb, Vrbo, Booking.com or any other
platform.** There is no sync, no iCal import, no API and no live
pricing feed. Bookings are typed or pasted in. This is said plainly on
the Getting Started tab as well, because a host who expects the
calendar to fill itself will be disappointed on day one, and because
several products sold with this promise cannot actually keep it.

It will not message a guest, send a check-in code, chase a review,
place a supply order or pay a cleaner. It has no app, no guest portal
and no smart-lock integration. It does not take payments and it stores
no card details. It holds no passports, no ID scans and no identity
documents — there is nowhere to put them on purpose.

It is not accounting or tax software, and nothing in it is accounting,
legal or tax advice. **How a cost is classified, what is deductible,
whether a capital item is depreciated and how short-term rental income
is taxed at all depend entirely on your jurisdiction and your
circumstances** — several countries treat furnished holiday lets as a
different thing again. The tax tab organises your year so that the
conversation with an accountant takes an hour instead of a weekend. It
does not replace them.

The occupancy, ADR and RevPAR figures are only as good as the nights
you have entered, and a property that was not listed for the whole
month will read low until you block the nights it was off the market.

What it does give you is every booking across every platform in one
place, a calendar that shows twelve properties at once, a double
booking caught before the guest finds it, a cleaning job for every
check-out with a printable checklist, the real fee each platform took
against the one it quoted, and a profit figure per property that
includes the mortgage, the rent, the laundry and the supplies.

---

## 6. What is deliberately *not* in version 1

* iCal, API or channel-manager synchronisation of any kind;
* dynamic pricing, competitor rate scraping and market data;
* guest messaging, automated check-in codes and smart-lock control;
* payment processing, deposit holds and damage claims;
* macro-driven features — one-click invoicing, bulk messaging, QR
  codes;
* depreciation schedules and jurisdiction-specific tax returns;
* multi-currency portfolios in one file, which need a rate table and a
  policy decision about when the rate is struck.

---

© Novality Store. Generated workbooks are licensed for personal use and
for a single business. Not affiliated with, endorsed by or sponsored by
Airbnb, Vrbo or Booking.com.
