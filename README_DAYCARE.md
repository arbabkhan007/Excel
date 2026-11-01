# Daycare & Childcare Business Management Spreadsheet — generator

> **This repository holds nine spreadsheet products** built on the same
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
> * **Daycare & Childcare Business Management Spreadsheet** — this
>   README, `build_daycare_workbook.py`, 23 tabs.
> * **Ultimate Airbnb & Short-Term Rental Management Spreadsheet** —
>   [README_AIRBNB.md](README_AIRBNB.md),
>   `build_airbnb_workbook.py`, 20 tabs.
> * **Wedding Photographer Business Manager** —
>   [README_WEDDING.md](README_WEDDING.md),
>   `build_wedding_workbook.py`, 23 tabs.
>
> Every download is listed in one place in
> [DOWNLOADS.md](DOWNLOADS.md).

Python code that builds a finished, ready-to-sell Excel workbook for a
nursery, daycare, preschool or home childcare setting: enrolment, the
daily register, the money, the rota, the rooms and every renewal date
in the building. One family, one balance, however many children. One
charge ledger that adds up into a printable invoice. One ratio table
you type your own licensing into, and every room alert in the workbook
follows it.

> **Listing title:** Daycare Business Spreadsheet, Childcare Management
> System, Enrollment Attendance Tuition Invoice Staff and Parent
> Tracker, Home Daycare and Preschool Excel Template
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_daycare_workbook.py
# -> output/Daycare Business Management Spreadsheet.xlsx
#    (23 tabs, ~2.5 MB)
```

The repository also contains the XlsxWriter library itself
(`xlsxwriter/`), so the command above runs with **no installation and
no dependencies** from the repository root. Outside this repo,
`pip install xlsxwriter` first — the generator is tested against the
bundled 0.3.0 and against XlsxWriter 3.2.

---

## 1. What gets built

23 tabs, 143,400+ formulas, all joined up and all driven by drop downs.

| # | Tab | What it does |
|---|-----|--------------|
| 1 | **Dashboard** | Today and this month in one screen: a dated header with children in today and the live occupancy, **twenty KPI tiles** in five bands — children on the books, occupancy, places free, waiting for a place · here today, attendance this month, absences, still to check out · billed, against target, collected, outstanding · operating profit, margin, wage cost of income, income per child · staff on the rota today, rooms over ratio, documents expired, jobs waiting on you — a "needs doing today" list that counts the real problems — missing check-outs, unauthorised collections, rooms over ratio, overdue families, lapsed certificates, expiring documents, low stock, unanswered messages — every room with its occupancy and ratio, the families who owe you money, birthdays coming up, and twelve months of billed, collected and profit with sparklines |
| 2 | **Getting Started** | The seven step set-up in order, a map of all 23 tabs, the house rules, the headline sums written out in plain English, the colour key, and an honest list of what the workbook will *not* do |
| 3 | **Settings & Lists** | Your setting's name, hours, trading weeks and licensed places; **the age band table — band, from and to in months, children per adult and maximum group size** — which is what makes the ratio alerts match your own licensing; fees, sibling and staff discounts, registration and deposit; payment terms, invoice day, late payment and late pickup charges with a grace period; tax rate and two switches; payroll defaults, overtime threshold and employer costs; eight targets; alert windows for documents, birthdays, absences and reordering; nineteen ID prefixes; 43 drop-down lists; the expense budget; and the month keys — around 70 named settings the rest of the workbook reads |
| 4 | **Child Enrollment** | 73 columns, one child a row. Automatic child ID, **age in months and the age band that follows from it**, family link and sibling number, program, room, teacher, schedule, days and hours a week, start and leaving dates — then the commercial side: six billing bases (weekly, fortnightly, monthly, term, daily or per-session, hourly) all converted to **one weekly figure**, the automatic sibling discount, subsidy, weekly, monthly and annual net fee, what has actually been billed, the family balance, sessions and attendance percentage, late arrivals and late pickups, allergy level, documents required against documents on file, the next birthday and the day they move up a room, and an action column |
| 5 | **Parent Database** | One row per **family**, not per child. Both parents, employer and work number, address, emergency contacts, doctor, the authorised pickup list, custody notes, preferred contact method and language — then the children linked to it by name, weekly and monthly fees for the whole family, invoiced, paid, **balance, overdue amount and the oldest unpaid due date**, days since they last paid, whether a subsidy is involved, open messages, and what to do about it |
| 6 | **Waitlist** | Every enquiry with the room and days they want, priority, sibling flag, source, tour and offer dates — then the child's age at the date they want to start, the band that puts them in, **their position in the queue for that room**, days waiting, days since you last spoke, seats free in that room right now, and what the place would be worth a week |
| 7 | **Attendance** | 2,000 rows of register. Date, child, status, in and out times, who dropped off, **who collected** — then the room and age band pulled from the child, hours, the expected times from their contract, late arrival in minutes, late departure, **the late pickup charge worked out from your grace period and per-minute rate**, whether the person who collected them was on the authorised list, a running absence count, and a missing check-out alert |
| 8 | **Tuition & Fees** | The charge ledger: one row per fee or credit. Tuition, registration, deposit, meals, transport, late pickup, extra sessions, trips, plus the credits — sibling, staff, promotional, subsidy, scholarship — all typed as positive numbers and **negated automatically**. Quantity × rate, discount, tax by your own switches, net, which invoice it is on, and a line key so a printable invoice can pull line 1, 2, 3 without an array formula |
| 9 | **Invoices & Receipts** | One ledger for **invoices, payments, credit notes, refunds and write-offs**. An invoice totals the charge lines pointed at it (or the amount you type if there are none); payments match an invoice by reference; the due date comes from your terms; and the status — Draft, Paid, Part paid, Due, Overdue — falls out of the balance and the date. **Billed** and **collected** are separate columns, so revenue and cash never get mixed up |
| 10 | **Daily Reports** | The report a parent reads at pickup: mood, behaviour, activities, learning, outdoor minutes, two naps, nappies and toileting, water and bottles, the teacher's observations and a space for the parent — plus hours from the register, nap hours, how many meals were logged and how they ate, incidents that day, and **a written summary assembled from the columns beside it** |
| 11 | **Meal Tracker** | Every plate: sitting, child, food, portion eaten, alternative offered, drink, water, own food brought in — with **the child's dietary requirement and allergy level pulled onto the same line as the food**, so an allergy is in front of whoever is serving |
| 12 | **Allergy & Medical** | Conditions, allergies with severity and reaction, action plans, medications with dose and storage, who is trained, consent and consent date, prescriber, immunisations, review and expiry dates — **and the medication dose log in the same register**, so each dose records the date, time, who gave it and who witnessed it, counted back against the medication record it belongs to |
| 13 | **Incident Reports** | What happened, where, category and severity, injury, first aid given and by whom, staff and witness, **when the parent was told, how, and by whom — with the minutes from the incident to the notification worked out for you**, the signature, follow-up action and whether it was done, and whether it had to be reported outside |
| 14 | **Staff Database** | The team: role, room, contact, hire date, employment status, pay type (salary, hourly, daily, contractor, volunteer) and rate, contracted hours — then **first aid, CPR, safeguarding and background check expiry dates rolled into one countdown**, documents missing, hours this week and this month, over or under contract, pay this month, how many children are in their room, and an action |
| 15 | **Staff Schedule** | The rota and the clock. Date, staff, room, shift type, leave, start, finish, break, **clocked in and clocked out**, substitute — then rota'd hours, actual hours (falling back to the rota when nobody clocked), variance, hours so far that week, **overtime split off hour by hour at your weekly threshold**, basic hours, shift cost, and a double-shift alarm. The "counts towards ratio" flag is what the classroom tab reads |
| 16 | **Payroll Tracker** | One row per person per pay run. Hours come from the rota between the two dates, the rate from the staff tab, overtime at your multiplier — then gross pay, **employer costs at your percentage (zero for contractors and volunteers)**, total cost to the setting, net pay after deductions, and cost per hour |
| 17 | **Classroom Management** | One row per room: capacity, age band, lead teacher and assistant — then children enrolled, active, starting and leaving, **seats free and occupancy**, how many are waiting for it, **children present today, staff in the room today, staff needed at your ratio, and whether the ratio is met right now**, weekly revenue, revenue per seat and **what the empty places are costing you a week** |
| 18 | **Expense Tracker** | Every cost with a category, vendor, tax, method, receipt reference and recurring frequency — turned into a **monthly equivalent** so a yearly insurance bill and a weekly fruit delivery become one honest overhead, measured against the budget for that category, and expressed as a cost per child. **Wages are counted once**: a wage row is flagged and kept out of the profit figure unless you tell Settings you are not using the payroll tab |
| 19 | **Supply Inventory** | Nappies, wipes, craft materials, food, cleaning: **opening + received − used − waste + adjustments**, stock value, weekly usage, weeks of cover, how much to order and what it will cost, a status, and the date you last counted |
| 20 | **Parent Communication** | Every conversation: date, time, family, child, direction, method, category, subject, what was said, who dealt with it, the follow-up date and action, status — then days open, whether a follow-up is overdue, and an alert |
| 21 | **Document Tracker** | Children's forms, staff certificates and the setting's own licences and policies in one place: who it belongs to, type, whether it is required, status, requested, received and expiry dates, renewal interval, owner, where the original lives — then **days left and a state: Missing, Expired, Expiring, On file, Not required** |
| 22 | **Monthly Reports** | Pick a month and everything follows it. Nine headline numbers against last month, children in and out, **room by room**, money in, charges by fee type, money out, expenses by category against budget, care, safety and paperwork — and a **printable invoice or receipt**: choose a reference from the drop-down and the document builds itself from the ledger and the charge lines, fourteen lines, totals, terms, ready to print |
| 23 | **Annual Reports** | Pick a year. Eight headline numbers, **twelve months down the page against twelve measures, with sparklines**, this year against last, the fee mix, costs by category against a twelve-month budget, who pays the bills, and the year in people — enrolments, leavers, enquiries, conversion, staff hours, incidents, doses and conversations |

### The sums that matter

Every one of these is written out on the Getting Started tab too,
because a buyer should be able to check the arithmetic rather than
trust it.

| Figure | How it is worked out |
|--------|----------------------|
| Weekly fee | the rate converted by the billing basis — monthly × 12 ÷ 52, term ÷ term weeks, daily × days a week, hourly × hours × days |
| Net weekly fee | weekly fee × (1 − discount) − subsidy |
| Sibling discount | applied automatically from the second child in a family onwards |
| Invoice total | the sum of the charge lines pointed at it, or the amount typed if there are none |
| Billed | invoices − credit notes − write-offs |
| Collected | payments − refunds |
| Family balance | billed − collected, for the whole family |
| Attendance % | days present ÷ sessions that were not a closure day |
| Late pickup charge | base fee once past the grace period, plus a per-minute rate after the first block |
| Staff needed | children present today ÷ children per adult for that room's band, rounded up |
| Occupancy | children enrolled ÷ places in that room |
| Empty seat cost | seats free × (room revenue ÷ children enrolled) |
| Total cost of a pay run | gross pay + employer costs at your percentage |
| Operating profit | billed − payroll total cost − other expenses |
| Days left on a document | expiry date − today |

### Rules baked in

* **The ratio table is a setting, not a rule.** Six age bands ship
  filled in — infant, young toddler, toddler, preschool, pre-K, school
  age — each with the months it covers, the children one adult may
  supervise and the maximum group size. Every one of those numbers is
  a cell you type over. A setting in Ohio, Ontario or Oxfordshire
  types in its own figures and every room alert in the workbook
  follows, because nothing anywhere is hard coded to one country's
  licensing.
* **The family is the billing unit.** Siblings share a row on Parent
  Database, share one balance and appear on one invoice. The sibling
  discount is applied from the second child onwards without anybody
  remembering to.
* **Charges and invoices are separate things.** A fee is a line on
  Tuition & Fees. An invoice is a reference those lines point at. That
  is what makes a printable invoice possible, what makes a part
  payment work, and what stops a credit note from quietly deleting the
  history of what was charged.
* **Revenue and cash are never mixed.** Billed is invoices less
  credits and write-offs. Collected is payments less refunds. Both
  live in the same ledger and neither is ever derived from the other.
* **Wages are counted once.** The profit figure uses the payroll tab's
  total cost. A wage row typed into Expenses is flagged and excluded —
  unless you switch `PayrollInExpenses` to Yes on Settings, in which
  case the payroll tab stops feeding the profit instead.
* **An overdue invoice is one you actually sent.** Nothing is chased
  until the `Sent` flag says Yes.
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
| **The finished workbook, with the example nursery filled in** | 2,471 KB | [daycare-business-spreadsheet.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/daycare-business-spreadsheet.xlsx) |
| **The finished workbook, empty and ready to sell** | 2,254 KB | [daycare-business-spreadsheet-blank.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/daycare-business-spreadsheet-blank.xlsx) |
| **The generator source, zipped** | 380 KB | [daycare-business-spreadsheet-source.zip](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/daycare-business-spreadsheet-source.zip) |

The example setting is **Little Acorns Childcare** of Columbus, Ohio:
four rooms — Acorns 8 places, Saplings 10, Oaks 12, Willows 12 — 42
licensed places, 34 children across 20 families (13 of them with more
than one child, which is the whole point), and twelve staff including
one whose first aid certificate has already lapsed and one whose
expires next month. Ten weeks of register — 1,371 rows at 91% present,
with a closure day, two missing check-outs and one collection by
somebody who was not on the list — 200 charge lines, 157 ledger rows
across four months of invoices with part payments, three late
families, a deposit, a credit note, a refund and a write-off, 278
shifts, 170 daily reports, 243 meals, 18 incidents, 26 parent
conversations and 197 documents of which 28 need renewing. Every tab
means something the moment it opens. `--blank` strips all of it out
and leaves the formulas.

Rebuild the three files with:

```bash
python build_daycare_workbook.py -o download/daycare-business-spreadsheet.xlsx
python build_daycare_workbook.py --blank -o download/daycare-business-spreadsheet-blank.xlsx
python tools/export_source_md.py --product daycare
python tools/make_download_zip.py --product daycare
```

---

## 2. Command line

```bash
python build_daycare_workbook.py          [options]
python -m daycare_business_spreadsheet    [options]      # same thing
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Daycare Business Management Spreadsheet.xlsx` | where to write the file |
| `--blank` | off | no example data, just the empty workbook |
| `--no-protect` | off | leave the sheets unlocked (for development) |
| `--password TEXT` | `premium` | sheet protection password |
| `--author TEXT` | `Novality Store` | document author |
| `--centre TEXT` | `Little Acorns Childcare` | your setting's name on the cover and the reports |
| `--owner TEXT` | `Marianne Okafor` | the owner or director |
| `--currency SYMBOL` | `$` | currency symbol, e.g. `£`, `€`, `A$` |
| `--currency-code XXX` | `USD` | three letter code used in headings |
| `--year YYYY` | this year | the year the reports open on |
| `--tax PERCENT` | 0 | tax rate — childcare is exempt in many places, so it ships at zero |
| `--tax-on WHAT` | `neither` | `both`, `tuition`, `extras` or `neither` |
| `--capacity N` | 42 | licensed places for the whole setting |
| `--open-days N` | 5 | days open a week |
| `--open-hour N` | 7 | opening time as a 24 hour number |
| `--close-hour N` | 18 | closing time as a 24 hour number |
| `--weeks-open N` | 50 | trading weeks a year |
| `--full-time-week AMOUNT` | 285 | the standard full-time weekly fee |
| `--basis TEXT` | `Weekly rate` | the billing basis new children start on |
| `--sibling-discount PERCENT` | 10 | sibling discount from the second child |
| `--hourly AMOUNT` | 16.50 | default staff hourly rate |
| `--burden PERCENT` | 14 | employer costs on top of gross pay |
| `--overtime-after N` | 40 | hours a week before overtime starts |
| `--target AMOUNT` | 42,000 | monthly revenue target |
| `--rows N` | 2000 | rows of capacity in Attendance; every other register scales with it |
| `--today YYYY-MM-DD` | today | pretend it is this date (keeps the demo data reproducible) |
| `--quiet` | off | print only the output path |

```bash
# a blank British nursery: 60 places, six days, term-time weeks, VAT off
python build_daycare_workbook.py --blank --currency "£" --currency-code GBP \
    --centre "Oakfield Day Nursery" --capacity 60 --open-days 5 \
    --open-hour 7.5 --close-hour 18 --weeks-open 51 --full-time-week 320 \
    --sibling-discount 15 --hourly 12.5 --burden 18 --target 58000 \
    -o "output/Nursery GBP.xlsx"
```

---

## 3. How the code is organised

```
build_daycare_workbook.py          the launcher
daycare_business_spreadsheet/
    __init__.py  __main__.py  cli.py     entry points and the options
    config.py                            every default, list, age band
                                         and palette
    layout.py                            Col / SheetSpec / Registry, the
                                         formula template engine
    formats.py                           every cell style, once
    writer.py                            SheetWriter: banners, sections,
                                         KPI tiles, tables, grids
    names.py                             the live OFFSET lookup names
    build.py                             tab order and the build pipeline
    sample_data.py                       Little Acorns, ten weeks of it
    sheets/
        common.py                        shared column helpers
        classrooms.py children.py parents.py waitlist.py
        attendance.py tuition.py invoices.py
        daily_reports.py meals.py medical.py incidents.py
        staff.py schedule.py payroll.py
        expenses.py supplies.py communication.py documents.py
                                         the eighteen registers
        dashboard.py getting_started.py settings.py
        monthly.py annual.py             the five hand-built pages
tools/
    validate_daycare_workbook.py         static checks
    check_parens.py                      brackets, quotes, placeholders
    check_daycare_sample.py              does the demo nursery add up
    recalc_check.py                      really evaluate every formula
```

### Formula templates

A column is declared once, for one row, with placeholders:

```python
Col("weekly_net", "Weekly fee (net)", 12, "money",
    formula=blank("name",
        'ROUND(MAX(0,{gross_week}{r}*(1-{discount_used}{r})'
        '-{subsidy_week}{r}),2)'))
```

`{r}` is the row, `{key}` is another column on the same sheet,
`[[Sheet.key]]` is a whole column on another sheet and
`[[Sheet.key#12]]` is one cell of it. The registry resolves all of it
to `$Y$5`-style references at build time, so the workbook a buyer opens
contains ordinary Excel formulas with no names to maintain — and moving
a column never breaks a reference, because nothing refers to a letter.

### The chain

```
Settings (age bands, ratios, fees, terms, targets)
    │
    ├─> Classroom Management ──> capacity, occupancy, ratio needed
    │            ▲         ▲
    │            │         └── Staff Schedule ──> who is in the room now
    │            │                     │
    │            │                     ├─> overtime ──> Payroll ──> cost
    │            │                     └─> Staff Database ──> hours, certs
    │            │
    ├─> Parent Database ──> one family, one balance
    │            ▲
    │            └── Child Enrollment ──> age band, weekly fee, siblings
    │                        │
    │                        ├─> Attendance ──> hours, late pickup charge,
    │                        │          │       authorised collection
    │                        │          └─> Daily Reports, Meal Tracker
    │                        │
    │                        └─> Tuition & Fees ──> charge lines
    │                                     │
    │                                     └─> Invoices & Receipts
    │                                              │
    │                                              ├─> family balance
    │                                              └─> printable invoice
    │
    ├─> Expense Tracker ──> monthly overhead ─┐
    └─> Document Tracker ──> renewals ────────┴─> Dashboard, Monthly,
                                                  Annual
```

### Adding something

A new register is one module with `NAME`, `spec(cfg)` and an optional
`decorate(sheet, fmts, cfg, reg)`, added to `REGISTER_MODULES` and
`SHEET_ORDER` in `build.py`. A new page is a module with
`build(book, fmts, cfg, reg)` added to `PAGE_BUILDERS`. A new setting is
one line in `config.py` and one `s.field(...)` on the settings tab,
which gives it a defined name every other formula can use.

---

## 4. Quality checks

```bash
python build_daycare_workbook.py -o output/test.xlsx
python tools/check_parens.py             output/test.xlsx
python tools/validate_daycare_workbook.py output/test.xlsx  # needs openpyxl
python tools/check_daycare_sample.py      output/test.xlsx

python build_daycare_workbook.py --rows 130 -o /tmp/small.xlsx
python build_daycare_workbook.py --rows 130 --blank -o /tmp/blank.xlsx
python tools/recalc_check.py /tmp/small.xlsx    # needs formulas
python tools/recalc_check.py /tmp/blank.xlsx
```

`check_parens.py` walks the saved XML and balances the brackets and
quotes of all 143,419 formulas — concatenated nested `IF`s are the one
mistake a generator makes that no amount of reading catches — and
fails on any template placeholder or sheet token left behind.

`validate_daycare_workbook.py` reports: tab order, document author,
sheet protection and passwords, unresolved tokens, `#REF!`, unbalanced
brackets, odd quote counts, empty arguments, unknown or too-modern
functions (`XLOOKUP`, `IFS`, `MAXIFS`, dynamic arrays — none are used),
formulas pointing at a missing sheet, defined names and drop-down
sources that do not exist, column letters left without a row number,
formula cells that are not locked, **`a+b/c` where `(a+b)/c` was
meant**, **`COUNTIFS` and `SUMIFS` whose criteria ranges are different
sizes** (Excel answers `#VALUE!` and nothing else warns you), the 99
links of the automation chain, the 45 named cells the pages depend on,
the 22 drop-down lists that must exist, and that **every cell of the
staff-to-child ratio table is still unlocked** — a buyer who cannot
edit the ratios cannot use the workbook.

`check_daycare_sample.py` ignores every formula and reads only the
cells a user would have typed, then rebuilds the arithmetic in Python:
it fails if a child is in a room that does not exist, if a room holds
more children than it has places, if a child is marked present before
they started or after they left, if a child is marked twice on one
day, if a child is collected before they arrived, if an invoice's
typed total disagrees with its charge lines, if a payment is bigger
than the invoice it is against, if a payment names no invoice (a
deposit may), if two shifts for the same person overlap, if a shift
ends before it starts, if any supply's stock works out negative, if
the demo months are so far from the target that the dashboard would
look wrong, or if the demo is so tidy that the alert columns have
nothing to show — no missing check-out, no lapsed certificate, no
expired document and no siblings would each be a failure, because the
features would be invisible on first open.

`recalc_check.py` loads the workbook into the `formulas` engine and
evaluates the whole dependency graph, then checks every formula cell
for an error value. It is run at two sizes on **both** the sample and
the blank build, because the empty workbook is the one that exposes a
division by zero and the half-full one is where a leave day with no
shift times turns an hours column to `#VALUE!` — a buyer's first ten
minutes are spent in exactly those files.

Current state: 23 tabs, **143,419 formulas**, 178 defined names, 38,247
unlocked cells in the sample build and 668 in the blank one, no
problems from any of the four tools, and all four recalculation runs
finish with zero error cells.

---

## 5. What a spreadsheet cannot do

It will not send a message, an invoice, a reminder or a daily report to
a parent, it has no app, no parent portal and no sign-in kiosk, and it
does not sync with a childcare platform, a payment processor or a
government subsidy system. Everything in it is typed or pasted in by
the setting. It does not take payments and it stores no card details.

It is not accounting, payroll or tax software, and nothing in it is
accounting, payroll, legal or tax advice. It does not calculate income
tax, national insurance or social security — it shows what a pay run
costs the setting, which is the number an accountant actually asks for.

**It is a business management workbook, not a compliance system.** It
is not a child development or observation record, it is not a medical
record, and it does not certify anything to a licensing body. The
ratio table helps you *watch* your ratios; meeting them is your job,
and the figures in it are the ones you type. Your own regulator's
rules always win.

What it does give you is one balance per family however many children
they have, a bill that assembles itself from the charges you logged, a
register that tells you who has not been collected and who collected
somebody they should not have, a room view that says you are one adult
short right now, and a list of every certificate and form in the
building with the number of days left on it.

---

## 6. What is deliberately *not* in version 1

* a parent app, portal, sign-in kiosk or automated daily report email;
* payment links, direct debit and card processing;
* macro-driven features — one-click invoicing, bulk messaging, QR
  code check-in;
* child development tracking, observations, learning journeys and
  curriculum frameworks;
* government subsidy and funding claim returns, which differ in every
  jurisdiction;
* multi-site consolidation and inter-site staff sharing.

---

© Novality Store. Generated workbooks are licensed for personal use and
for a single business.
