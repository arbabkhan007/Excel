# Wedding Photographer Business Manager — generator

> **This repository holds ten spreadsheet products** built on the same
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
>   this README, `build_wedding_workbook.py`, 23 tabs.
> * **Ultimate Inventory & Wholesale Order Manager** —
>   [README_INVENTORY.md](README_INVENTORY.md),
>   `build_inventory_workbook.py`, 23 tabs.
>
> Every download is listed in one place in
> [DOWNLOADS.md](DOWNLOADS.md).

Python code that builds a finished, ready-to-sell Excel workbook for a
wedding photographer: the enquiries, the dates, the contracts, the
money, the running order of the day itself, the shot list that stops
you missing the grandmother who will not be there next year, the
editing backlog, the gallery, and the question most photographers
cannot answer — **which weddings actually made money, and which ones
you were paying to attend**.

Every tab is joined to every other tab by a **Booking ID** (`WED-0001`),
so a wedding is one record from first enquiry to final album, not nine
unrelated lists in the same file.

> **Listing title:** Wedding Photographer Business Spreadsheet |
> Photography CRM, Booking Calendar, Pricing, Invoices & Client
> Management
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_wedding_workbook.py
# -> output/Wedding Photographer Business Manager.xlsx
#    (23 tabs, 114,996 formulas, ~1.7 MB)
```

The repository also contains the XlsxWriter library itself
(`xlsxwriter/`), so the command above runs with **no installation and
no dependencies** from the repository root. Outside this repo,
`pip install xlsxwriter` first — the generator is tested against the
bundled 0.3.0 and against XlsxWriter 3.2.

---

## 1. What gets built

23 tabs, in this order. Six are **pages** — designed screens you read
and print. Seventeen are **registers** — a title, a frozen header row,
and hundreds of rows of typed columns and locked formulas.

| # | Tab | Kind | Formulas | What it is for |
|---|-----|------|---------:|----------------|
| 1 | **Dashboard** | page | 440 | The centrepiece. Money, this month, the next fourteen days, and a "what needs doing" matrix that counts every alert on eleven registers. |
| 2 | **Getting Started** | page | 0 | The one thing to understand (the Booking ID), six steps for the first hour, the five workflows, and an honest list of what this is *not*. |
| 3 | **Settings & Lists** | page | 6 | Your studio, your prices, your turnaround promises, your targets, and the 48 drop-down lists everything else reads. |
| 4 | **Client CRM** | register | 4,500 | Couples: contact details, venue, style, how they found you, what they are worth. |
| 5 | **Lead Tracker** | register | 6,800 | Enquiries before they are bookings. Reply times, consultations, proposals, win rate, and why you lost. |
| 6 | **Booking Calendar** | page | 1,188 | Twelve months of date squares. Green means held, red means two couples want the same Saturday. |
| 7 | **Wedding Bookings** | register | 11,400 | The hub. One row per wedding, 87 columns, every other tab joins here. |
| 8 | **Packages & Pricing** | register | 720 | What each package costs you to deliver, what it earns per hour, and how often it converts. |
| 9 | **Contracts & Documents** | register | 2,000 | What was sent, what came back signed, what expires, where the file lives. |
| 10 | **Payments & Invoices** | register | 9,600 | Deposit, instalments, balance, add-ons, refunds. Who owes you and for how long. |
| 11 | **Wedding Timeline** | register | 18,900 | The running order of the day, built from offsets against the ceremony time, with planned against actual. |
| 12 | **Shot List Manager** | register | 19,600 | Every shot, who is in it, who asked for it, whether it is a must-have, whether you got it. |
| 13 | **Photographer & Staff** | register | 760 | You and your second shooters. Rates, skills, insurance expiry. |
| 14 | **Editing Workflow** | register | 4,400 | Ten steps per wedding, a due date from your own promise, and an honest cost per hour. |
| 15 | **Gallery & Delivery** | register | 3,000 | Where the gallery lives, when it was delivered, when it expires, what it sold after the fact. |
| 16 | **Equipment Inventory** | register | 1,950 | Kit, value, shutter counts, warranty, what each item is a backup for. |
| 17 | **Travel & Mileage** | register | 4,500 | Mileage, tolls, parking, accommodation — and whether the travel fee covered any of it. |
| 18 | **Marketing Calendar** | register | 6,000 | Posts, platforms, spend and the enquiries each one actually produced. |
| 19 | **Expense Tracker** | register | 7,200 | Every cost, tagged to a wedding or to the business, with a tax category. |
| 20 | **Profit Analysis** | register | 8,600 | One line per wedding: revenue in, every cost out, margin, and profit per hour of your life. |
| 21 | **Client Questionnaire** | page | 278 | A printable sheet for one wedding, pulling the timeline and the must-have shots the couple already told you about. |
| 22 | **Reviews & Referrals** | register | 2,800 | Who was asked, who replied, what they said, who they sent you. |
| 23 | **Monthly & Annual Reports** | page | 354 | A month against the same month last year, then twelve months in a grid. |

**114,996 formulas** in total, every one of them locked, on a 1.7 MB
file that opens in Excel 2010.

### The five workflows it was built around

1. **Enquiry → confirmed booking.** Lead Tracker records the enquiry
   and chases it on your own schedule; when it converts you pick the
   Client ID and the wedding appears on Wedding Bookings with a Booking
   ID, a contract row and a payment schedule behind it.
2. **Package → payment collection.** Packages & Pricing sets the price;
   Wedding Bookings works out the real total with extras, travel and
   discount; Payments & Invoices tracks deposit, instalments and
   balance and tells you who is late and by how many days.
3. **Wedding day: timeline + shots.** Wedding Timeline builds the
   running order from offsets against the ceremony time, so moving the
   ceremony moves the whole day. Shot List Manager holds the must-have
   shots and the family groups, and the Client Questionnaire prints
   them.
4. **Editing → final gallery delivery.** Editing Workflow runs ten
   steps against a due date derived from the turnaround you promised on
   Settings; Gallery & Delivery takes over at upload and tracks views,
   downloads, album orders and the expiry date.
5. **Revenue, expenses and profit per wedding.** Profit Analysis pulls
   the money in from the booking, the costs out from staff, editing,
   travel, albums and the expense tracker, and reports margin and
   **profit per hour**.

### The Booking ID

One decision holds the whole product together: **every wedding gets a
unique Booking ID** (`WED-0001`), generated once on Wedding Bookings and
chosen from a drop-down everywhere else. Eleven registers join back
through it. `tools/validate_wedding_workbook.py` checks that join in
both directions on every build, because the moment a Booking ID can be
mistyped the payments, the timeline, the shot list, the editing queue
and the gallery stop being the same wedding.

---

## 1b. Download it

You do not have to run anything. The finished workbooks are in
[`download/`](download/) and indexed in [DOWNLOADS.md](DOWNLOADS.md):

| File | What it is |
|------|------------|
| `wedding-photographer-business-manager.xlsx` | Filled in with a worked example — one studio, fifteen weddings. |
| `wedding-photographer-business-manager-blank.xlsx` | Empty, ready for your own weddings. |
| `wedding-photographer-business-manager-source.zip` | The generator source. |

### The example studio

The sample file is Rowan & Fern Photography, and the year it is having
is deliberately not a flattering one:

* a wedding from three months ago **still in editing**, 92 days against
  a 42-day promise, parked at "selection complete" because the owner
  kept meaning to do it herself — while an editor who charges $28 an
  hour sits idle;
* a couple whose **balance is 43 days overdue** after three emails and
  no phone call;
* **the same Saturday promised to two couples**, one confirmed and one
  pencilled in, lighting up red on the calendar;
* an **elopement priced in two minutes on the phone** that works out at
  the worst hourly rate on the sheet;
* a second shooter whose **insurance expired twelve days ago**;
* a camera body at **93% of its rated shutter life**;
* a delivered wedding **nobody ever asked for a review**.

A sample studio where everything runs smoothly teaches nobody anything.
This one teaches you what each alert column is for.

---

## 2. Command line

```bash
python build_wedding_workbook.py [options]
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Wedding Photographer Business Manager.xlsx` | Where to write it |
| `--blank` | off | Leave the registers empty |
| `--no-protect` | off | Do not protect the sheets |
| `--password TEXT` | `premium` | Sheet protection password |
| `--studio TEXT` | Rowan & Fern Photography | Pre-fill the studio name |
| `--owner TEXT` | Marisa Calloway | Pre-fill the photographer |
| `--currency TEXT` | `$` | Currency symbol |
| `--currency-code TEXT` | `USD` | Three-letter code |
| `--deposit N` | `30` | Deposit percentage that holds a date |
| `--final-days N` | `14` | Days before the wedding the balance falls due |
| `--extra-hour N` | `250` | Charge for an extra hour on the day |
| `--second-shooter N` | `350` | What a second shooter costs you |
| `--mileage N` | `0.67` | Mileage rate |
| `--delivery-days N` | `42` | The gallery turnaround you promise |
| `--editing-hours N` | `18` | Honest editing hours per wedding |
| `--tax N` | `0` | Sales tax / VAT percentage on packages |
| `--weddings-target N` | `24` | Weddings a year you want |
| `--revenue-target N` | `72000` | Revenue target for the year |
| `--margin-target N` | `55` | Profit margin target, percent |
| `--rows N` | — | Scale every register against this many wedding rows |
| `--today ISO` | today | Treat this date as today in the example data |
| `--quiet` | off | Print only the output path |

```bash
# a blank UK studio, VAT registered, pounds, 20 weddings a year
python build_wedding_workbook.py --blank \
    --studio "Ashgrove Photography" --owner "Ellis Ashgrove" \
    --currency "£" --currency-code GBP --tax 20 \
    --weddings-target 20 --revenue-target 54000 \
    -o "output/Ashgrove.xlsx"

# a small file for testing - every register scaled down
python build_wedding_workbook.py --rows 18 -o /tmp/small.xlsx
```

---

## 3. How the code is organised

```
wedding_photography_spreadsheet/
├── config.py        every number, name, list and disclaimer
├── layout.py        Col / SheetSpec / Registry, formula helpers
├── formats.py       the warm-neutral palette and every cell style
├── writer.py        SheetWriter: banner, section, field, kpi, grid, cf
├── names.py         the five dynamic OFFSET lists (BookingList …)
├── sample_data.py   the worked example, 517 rows
├── build.py         registration order, sheet order, page builders
├── cli.py           argument parsing
└── sheets/
    ├── common.py            shared column factories and CF rules
    ├── settings.py          page: 48 lists, every setting
    ├── dashboard.py         page: the centrepiece
    ├── getting_started.py   page: how to use it, and what it is not
    ├── calendar.py          page: twelve months of date squares
    ├── questionnaire.py     page: printable, one wedding
    ├── reports.py           page: monthly and annual
    └── (17 register modules)
```

A **register module** exports `NAME`, a `spec(cfg)` returning a
`SheetSpec`, and optionally `decorate(...)`. A **page module** exports
`build(book, fmts, cfg, reg)`. `build.py` registers all 17 register
specs *first*, so a page written early can still point at a register
written late, then writes the tabs in the order the listing promises.

### Two rules the engine enforces

* **Registers never hand-write a cell reference.** A column's formula
  uses `{key}{r}` for its own row and `[[Sheet.key]]` for another tab's
  column; `Registry.render` turns those into real references at write
  time. Reordering columns cannot break a formula.
* **Pages must go through `SheetWriter.formula()`**, never
  `ws.write_formula`, or a `[[Sheet.key]]` token ships unrendered. The
  validator fails the build if one ever does.

---

## 4. Quality checks

```bash
pip install pyflakes openpyxl formulas

pyflakes wedding_photography_spreadsheet/ tools/
python3 tools/check_wedding_specs.py             # specs, before a build
python3 tools/check_wedding_sample.py            # the example data
python3 build_wedding_workbook.py -o /tmp/w.xlsx
python3 tools/check_parens.py /tmp/w.xlsx
python3 tools/validate_wedding_workbook.py /tmp/w.xlsx
python3 tools/recalc_check.py /tmp/w.xlsx        # use --rows 18 first
```

What each one is actually for:

* **`check_wedding_sample.py`** — every key in a sample row must be an
  *input* column (a value written into a formula column is silently
  dropped and takes the formula with it), every drop-down value must
  exist in its list, and every Booking ID must point at a real wedding.
  An earlier product in this repo shipped 199 bad drop-down values
  before this check existed.
* **`validate_wedding_workbook.py`** — 79 automation-chain links, the
  Booking ID join in both directions, overlapping merges, unresolved
  tokens, undefined names, unlocked formula cells, `a+b/c` where
  `(a+b)/c` was meant, collapsed `COUNTIFS` criteria, ragged criteria
  ranges, and functions that need a modern Excel.
* **`recalc_check.py`** — actually evaluates every formula and fails on
  any cell that ends up `#VALUE!`, `#DIV/0!` or `#REF!`. Run it on the
  **blank** file too: empty rows are where division by zero lives.

Current state on a full build, sample and blank:

```
  . automation chain: 79 links verified
  . summary pages: 7 live formulas, 31 typeable inputs verified
  . booking ID join: 11 registers joined through BookingList
  . product promises: 0 false sync claims, 0 binding-contract claims
  . 262 defined names
  . 114,996 formula cells
  OK - no problems found
```

---

## 5. What a spreadsheet cannot do

This is a business organisation tool. It is **not** a booking
integration and **not** a contract-signing platform, and it says so on
the Getting Started tab rather than burying it in a footnote.

* **It does not connect to anything.** No live calendar sync, no
  studio-management API, no card processor. You type what happened;
  the formulas do the arithmetic. Every date square on the Booking
  Calendar is drawn from the rows *you* entered.
* **It tracks documents, it does not create them.** Contracts &
  Documents records that a contract was sent on a date and came back
  signed on another, and where the PDF lives. It does not draft a
  contract, it does not sign one, and nothing it produces is a legal
  instrument. Have a real contract, written by someone qualified,
  stored somewhere you can find it.
* **It is not an accounting package.** Profit Analysis and the reports
  are management figures for making decisions. Tax treatment depends
  on your jurisdiction and your circumstances — the workbook says so
  and points you at an accountant.
* **It holds no sensitive personal data by design.** No identity
  documents, no card numbers. Contact details and a venue are enough to
  run a wedding business.
* **It cannot tell you whether the photographs are any good.** It can
  tell you that the ones from July are 50 days late.

---

## 6. What is deliberately *not* in version 1

* **Charts.** Every product in this repo avoids them: they bloat the
  file, they break on older Excel and they are the first thing to go
  wrong when a buyer adds a row. The dashboard uses conditional
  formatting and bar-style cells instead.
* **Macros.** A macro-free `.xlsx` opens everywhere, including Google
  Sheets and Excel on the web, with no security warning.
* **A separate album-design or print-sales module.** Album orders and
  after-sales live on Gallery & Delivery, which is where you are
  already looking.
* **Multi-currency.** One currency per workbook. Build a second file if
  you shoot destination weddings billed in euros.
* **Per-image culling or catalogue management.** That is Lightroom's
  job. This tracks the *counts* and the *hours*, which is what costs
  money.

---

## 7. The other photography products

The CRM, booking, payment and dashboard core here is deliberately
reusable. Portrait, product and event photography versions are the same
machine with different registers bolted on — the Booking ID join, the
package costing, the payment schedule, the editing workflow and the
profit-per-hour maths do not change.
