# Salon Business Management Spreadsheet — generator

> **This repository holds six spreadsheet products** built on the same
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
> * **Salon Business Management Spreadsheet** — this README,
>   `build_salon_workbook.py`, 20 tabs.

Python code that builds a finished, ready-to-sell Excel workbook for a
salon: the diary, the client book, the menu priced by the chair hour,
the till, the deposits, eight ways of paying a stylist, the stock
room, and a profit and loss that agrees with all of it. Not an income
and expense tracker with a hairdryer on the cover — a business system
for someone who owns four chairs and does not own a back office.

> **Listing title:** Salon Business Management Spreadsheet, Appointment
> and Client Tracker, Staff Commission Calculator, Salon Income Expense
> and Profit Template
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_salon_workbook.py
# -> output/Salon Business Management Spreadsheet.xlsx
#    (20 tabs, ~2.0 MB)
```

The repository also contains the XlsxWriter library itself
(`xlsxwriter/`), so the command above runs with **no installation and
no dependencies** from the repository root. Outside this repo,
`pip install xlsxwriter` first — the generator is tested against the
bundled 0.3.0 and against XlsxWriter 3.2.

---

## 1. What gets built

20 tabs, 122,600+ formulas, all joined up and all driven by drop downs.

| # | Tab | What it does |
|---|-----|--------------|
| 1 | **Start Here** | The six step set-up in order, a map of every tab, the seven house rules, the eleven headline sums written out in plain English, and an honest list of what the workbook will *not* do |
| 2 | **Dashboard** | The month in one screen: sixteen KPI tiles (net sales against target, operating profit, margin, clients served, average ticket, retail share, tips, completion, cancellation and no-show rates, chair utilisation, rebooking, repeat clients, money owed, stock on the shelf), a sixteen-line "needs doing today" list that counts the real problems, every stylist's month side by side, the best services and clients, twelve months of sales and profit with sparklines, and the month in one line down to break-even |
| 3 | **Appointment Planner** | The diary. Date, time, client, service, stylist, status, source, deposit, notes — and the finish time, the chair hours, the price, the deposit wanted against the deposit paid, what was actually billed, whether they rebooked, the lead time, and a **double-booking alarm** that checks every row against every other row for that stylist |
| 4 | **Client CRM** | Contact details, how they found you, who referred them, their usual stylist, birthday, colour formula, allergies, patch test date, marketing consent — and visits, last visit, days since, the gap they usually leave, when they are due, what is already booked, spend this year and lifetime, average ticket, outstanding balance, loyalty points, a segment (VIP, Regular, Active, New, Lapsing, Inactive) and an action for the front desk |
| 5 | **Services & Pricing** | The menu, with the minutes and the turnaround time, the backbar cost, the commission it earns and therefore the **gross profit and the profit an hour** — plus the price your chair rate target implies, the gap to it, how many times it has been booked, what it has earned, and how long since the price was reviewed |
| 6 | **Sales & Transactions** | One row per thing sold. Service, retail, package, gift certificate — quantity, unit price, discount as a percentage or an amount, promo code, refund, then **net = gross − discounts − refunds**, tax by your own rules, cost, profit, what has been paid against it (deposit included, once) and the balance |
| 7 | **Payments & Deposits** | Money in and out: deposits, balances, full payments, retail, tips, refunds, gift certificate and package sales. Knows which kinds count as revenue, which are the stylist's (tips), which reverse it (refunds), what the card machine takes, and what actually reaches the bank |
| 8 | **Staff Management** | The team: role, level, employment type, days, contracted hours — **and their own commission method and rate**, with appointments, completions, no-shows, clients, hours booked against hours available, utilisation, service and retail revenue, this month against target, average ticket, retail share, revenue an hour, rebooking rate, tips, commission, what they cost and what they leave behind |
| 9 | **Staff Commissions** | One row per stylist per month, and **eight supported methods**: none, percent of service, percent of service and retail, percent after product cost, tiered on monthly service revenue, flat amount per service, hourly only, rent a chair. The eligible base respects one switch — net of discounts or gross before them — and the tiered ladder is an editable table on Settings |
| 10 | **Payroll & Tips** | Pay runs: hours at a rate, the commission released this run, the tips being passed on, bonuses, holiday pay, deductions — then gross pay, employer costs at your percentage (zero for the self-employed and chair renters) and **the total cost to the salon**, which is the figure the profit and loss uses. A running "commission left" column stops a monthly commission being paid twice across weekly runs |
| 11 | **Expenses** | Every cost, categorised the way a profit and loss wants it, with a due date, a paid flag and an overdue alert — and a **monthly equivalent** column that turns a yearly insurance bill and a weekly laundry run into one honest fixed overhead |
| 12 | **Profit & Loss** | Twelve months across, the lines down: service, retail, packages and other income; discounts, refunds, tax and tips as memos; product used and stock sold; gross profit and margin; hourly pay, commission, employer costs and the staff cost share of service sales; every expense category; operating profit, operating margin, the target and the variance — with a sparkline of the year |
| 13 | **Product Inventory** | Retail and backbar in one place: cost, retail price, margin, **opening + received − used − sold − waste + adjustments**, weeks of cover at the current rate, a reorder point it works out when you have not set one, how much to order, stock value at cost and at retail, expiry dates, and the date you last counted |
| 14 | **Suppliers & Purchases** | Orders placed and received, short deliveries, invoices paid — and a **cost moved** column that shows a wholesaler's quiet price rise, line by line. The editable supplier book underneath feeds the supplier drop-downs on this tab and on the inventory tab |
| 15 | **Marketing & Promotions** | Campaigns with a promo code that links them to real sales: times used, revenue traced, **the discount given away counted as part of the cost**, new clients won through that channel, return, and the cost of winning a new client against your average ticket |
| 16 | **Client Loyalty** | Points earned from spend already logged, points redeemed, the balance, how many points to the next reward, what the outstanding rewards would cost you if everybody cashed in tomorrow, referrals counted from the Clients tab and the referral credit still owed |
| 17 | **Goals & Budget** | Nine goals for the year with the actual beside them, a month-by-month revenue plan with actual, variance, clients and profit, a **budget for every expense category** (which the expense tab reads back), and a break-even block: fixed costs plus the pay run over the margin you are really achieving |
| 18 | **Business Analytics** | Revenue by service category and revenue an hour for each; which days pay and which days cancel; **which hours are dead**; where the clients come from and what each channel is worth; new against returning month by month; and the client book split by value |
| 19 | **Weekly & Monthly Reports** | Two pickers at the top. Twelve measures for the week against last week and the four week average, the week day by day, then sixteen for the month against the month before with shares, the month by stylist, the month's best sellers, and the year to date |
| 20 | **Settings & Lists** | Salon details, opening hours and capacity, **tax rate with separate switches for services and retail**, card fees, deposit rules, cancellation policy, rounding rule, chair rate target, commission defaults and the tiered ladder, employer costs, nine targets, stock and reorder rules, loyalty and referral values, thirteen ID prefixes, twenty-four drop-down lists and the month keys — around 60 named settings the rest of the workbook reads |

### The sums that matter

Every one of these is written out on the Start Here tab too, because a
buyer should be able to check the arithmetic rather than trust it.

| Figure | How it is worked out |
|--------|----------------------|
| Net sales | gross − discounts − refunds |
| Operating profit | net sales − cost of sales − total staff cost − expenses |
| Average ticket | net sales ÷ completed appointments that were paid for |
| Completion rate | completed ÷ appointments with a final status (completed, cancelled, no show) |
| Cancellation rate | cancelled ÷ appointments with a final status |
| Client balance | service total − discounts − payments applied |
| Commission | eligible revenue × rate, by the method that stylist is on |
| Service gross profit | price − direct costs (product used + commission earned) |
| Stock on hand | opening + receipts − usage − sales − waste + adjustments |
| Budget variance | actual − budget (positive means over) |
| Repeat client rate | clients who came back ÷ clients served |

### Rules baked in

* **A deposit is never revenue twice.** It is logged on Payments
  against the *appointment*, not the sale. When the visit is rung up,
  only the first sale line of that appointment absorbs it — the
  `COUNTIFS(first:here, appointment) = 1` trick — so a client who has a
  colour and buys shampoo does not have their deposit counted against
  both lines.
* **Tips are not income.** They are logged as payments, flagged as not
  counting, kept out of net sales, shown on the payroll tab as money
  passing through, and listed as a memo on the profit and loss.
* **Refunds reverse, they do not disappear.** A refund reduces the net
  sale on the line it belongs to and appears as a negative payment.
* **Tax is a setting, not an assumption.** One rate, two switches
  (services, retail). Set both to No and the workbook simply never
  mentions tax again. No country is baked in anywhere.
* **The commission method belongs to the person.** Eight of them are
  supported and the workbook never assumes one. The basis — net of
  discounts or gross before them — is a single switch, applied to
  everybody, so nobody argues about which figure was used.
* **Wages come from the payroll tab only.** They are never an expense
  category as well, so they cannot be counted twice.
* **Stock is a cost when it sells**, not when it arrives. A big
  delivery does not make a month look like a disaster.
* **Blank rows cost nothing.** Every formula starts with a guard on its
  anchor column, every division has a floor, and the whole workbook was
  recalculated empty as well as full to prove it.

---

## 1b. Download it

Everything is committed to the repository, so these links work straight
from GitHub — right click, *Save link as*.

| What | Size | Link |
|------|-----:|------|
| **The finished workbook, with the example salon filled in** | 1,965 KB | [salon-business-spreadsheet.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/salon-business-spreadsheet.xlsx) |
| **The finished workbook, empty and ready to sell** | 1,819 KB | [salon-business-spreadsheet-blank.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/salon-business-spreadsheet-blank.xlsx) |
| **The generator source, zipped** | 288 KB | [salon-business-spreadsheet-source.zip](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/salon-business-spreadsheet-source.zip) |

The example salon is **Gloss & Grace** of Austin: four chairs, eight
people — an owner who still works the floor, a commission-only
colourist on a tiered ladder, a stylist on service and retail, a
self-employed nail and lash tech paid after product cost, an apprentice
on a flat amount per service, a front desk on hourly only, a part-timer
on a straight percentage and a chair renter who bills her own clients.
Ten weeks of diary — 731 bookings, 648 sales, 898 payments, 132
clients, 34 services, 32 products, 31 purchase orders, eight campaigns
and 46 loyalty members — so every tab means something the moment it
opens. `--blank` strips all of it out and leaves the formulas.

Rebuild the three files with:

```bash
python build_salon_workbook.py -o download/salon-business-spreadsheet.xlsx
python build_salon_workbook.py --blank -o download/salon-business-spreadsheet-blank.xlsx
python tools/export_source_md.py --product salon
python tools/make_download_zip.py --product salon
```

---

## 2. Command line

```bash
python build_salon_workbook.py          [options]
python -m salon_business_spreadsheet    [options]      # same thing
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Salon Business Management Spreadsheet.xlsx` | where to write the file |
| `--blank` | off | no example data, just the empty workbook |
| `--no-protect` | off | leave the sheets unlocked (for development) |
| `--password TEXT` | `premium` | sheet protection password |
| `--author TEXT` | `Novality Store` | document author |
| `--salon TEXT` | `Gloss & Grace Salon` | salon name on the cover and the reports |
| `--owner TEXT` | `Nadia Ellison` | the owner |
| `--currency SYMBOL` | `$` | currency symbol, e.g. `£`, `€`, `A$` |
| `--currency-code XXX` | `USD` | three letter code used in headings |
| `--year YYYY` | this year | the year the reports open on |
| `--tax PERCENT` | 8.25 | sales tax / VAT rate |
| `--tax-on WHAT` | `retail` | `both`, `retail`, `services` or `neither` |
| `--commission PERCENT` | 40 | default service commission |
| `--commission-method TEXT` | `Percent of service` | the method new stylists start on |
| `--hourly AMOUNT` | 15 | default hourly rate |
| `--burden PERCENT` | 12 | employer costs on top of gross pay |
| `--deposit PERCENT` | 25 | deposit as a share of the price |
| `--rounding RULE` | `Nearest 1` | how a calculated price is rounded |
| `--target AMOUNT` | 18,000 | monthly revenue target |
| `--chairs N` | 4 | chairs or stations, for the utilisation figures |
| `--open-days N` | 5 | trading days a week |
| `--rows N` | 1200 | rows of capacity in Appointments; every other register scales with it |
| `--today YYYY-MM-DD` | today | pretend it is this date (keeps the demo data reproducible) |
| `--quiet` | off | print only the output path |

```bash
# a blank British workbook: VAT on everything, 45% commission, six chairs
python build_salon_workbook.py --blank --currency "£" --currency-code GBP \
    --tax 20 --tax-on both --commission 45 --burden 15 --chairs 6 \
    --open-days 6 --target 32000 -o "output/Salon GBP.xlsx"
```

---

## 3. How the code is organised

```
build_salon_workbook.py            the launcher
salon_business_spreadsheet/
    __init__.py  __main__.py  cli.py     entry points and the options
    config.py                            every default, list and palette
    layout.py                            Col / SheetSpec / Registry, the
                                         formula template engine
    formats.py                           every cell style, once
    writer.py                            SheetWriter: banners, sections,
                                         KPI tiles, tables, grids
    names.py                             the live OFFSET lookup names
    build.py                             tab order and the build pipeline
    sample_data.py                       Gloss & Grace, ten weeks of it
    sheets/
        common.py                        shared column helpers
        services.py clients.py appointments.py
        sales.py payments.py
        staff.py commissions.py payroll.py
        expenses.py inventory.py purchases.py
        marketing.py loyalty.py          the thirteen registers
        start_here.py dashboard.py profit_loss.py
        goals.py analytics.py reports.py settings.py
                                         the seven hand-built pages
tools/
    validate_salon_workbook.py           static checks
    check_parens.py                      brackets, quotes, placeholders
    check_salon_sample.py                does the demo salon add up
    recalc_check.py                      really evaluate every formula
```

### Formula templates

A column is declared once, for one row, with placeholders:

```python
Col("net", "Net sale", 11, "money",
    formula=blank("date",
        'ROUND({gross}{r}-{discount}{r}-IF({refund}{r}="",0,'
        '{refund}{r}),2)'))
```

`{r}` is the row, `{key}` is another column on the same sheet,
`[[Sheet.key]]` is a whole column on another sheet and
`[[Sheet.key#12]]` is one cell of it. The registry resolves all of it
to `$Y$5`-style references at build time, so the workbook a buyer opens
contains ordinary Excel formulas with no names to maintain — and moving
a column never breaks a reference, because nothing refers to a letter.

### The chain

```
Services (price, minutes, product cost)
    │
    ├─> Appointments ──> clash check, deposit wanted, chair hours
    │        │
    │        └─> Sales ──> net, tax, cost, profit, balance
    │                 │
    │                 ├─> Payments ──> deposits applied once, tips apart,
    │                 │                card fees, cash banked
    │                 ├─> Clients  ──> visits, spend, segment, due date
    │                 ├─> Staff    ──> revenue, utilisation, retail share
    │                 │        └─> Commissions ──> eight methods
    │                 │                    └─> Payroll ──> cost to salon
    │                 ├─> Inventory ──> sold, stock on hand, reorder
    │                 └─> Marketing ──> traced revenue, discount given
    │
    └─> Expenses ──> monthly overhead ──> Profit & Loss ──> Dashboard
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
python build_salon_workbook.py -o output/test.xlsx
python tools/check_parens.py           output/test.xlsx
python tools/validate_salon_workbook.py output/test.xlsx  # needs openpyxl
python tools/check_salon_sample.py      output/test.xlsx

python build_salon_workbook.py --rows 70 -o /tmp/small.xlsx
python build_salon_workbook.py --rows 70 --blank -o /tmp/blank.xlsx
python tools/recalc_check.py /tmp/small.xlsx    # needs formulas
python tools/recalc_check.py /tmp/blank.xlsx
```

`check_parens.py` walks the saved XML and balances the brackets and
quotes of all 122,614 formulas — concatenated nested `IF`s are the one
mistake a generator makes that no amount of reading catches — and
fails on any template placeholder or sheet token left behind.

`validate_salon_workbook.py` reports: tab order, document author, sheet
protection and passwords, unresolved tokens, `#REF!`, unbalanced
brackets, odd quote counts, empty arguments, unknown or too-modern
functions (`XLOOKUP`, `IFS`, `MAXIFS`, dynamic arrays — none are used),
formulas pointing at a missing sheet, defined names and drop-down
sources that do not exist, column letters left without a row number,
formula cells that are not locked, **`a+b/c` where `(a+b)/c` was
meant**, the 57 links of the automation chain, and the 29 named cells
the pages depend on.

`check_salon_sample.py` ignores every formula and reads only the cells
a user would have typed, then rebuilds the arithmetic in Python: it
fails if the demo diary contains a double booking, if a completed visit
has nothing rung up against it, if the payments do not match the till,
if payroll releases commission that was never earned, if any product's
stock works out negative, if a sale uses a promo code that is not on
the marketing tab, if a client appears twice on the loyalty scheme, or
if the demo months are so far from the target that the dashboard would
look wrong.

`recalc_check.py` loads the workbook into the `formulas` engine and
evaluates the whole dependency graph, then checks every formula cell
for an error value. It is run on **both** the sample and the blank
build, because the empty workbook is the one that exposes a division by
zero, and a buyer's first ten minutes are spent in exactly that file.

Current state: 20 tabs, **122,614 formulas**, 137 defined names, 28,831
unlocked cells, no problems from any of the four tools, and both small
builds recalculate with zero error cells.

---

## 5. What a spreadsheet cannot do

It will not send a confirmation, a reminder, a text or an email, and it
does not sync with a booking platform, a till or a card reader.
Everything in it is typed or pasted in by the owner. It does not take
payments and it stores no card details.

It is not accounting, payroll or tax software, and nothing in it is
accounting, payroll or tax advice. It does not calculate income tax,
national insurance or social security — it shows what a pay run costs
the salon, which is the number an accountant actually asks for.

What it does give you is a diary that cannot double book, a client book
that tells you who is drifting away, prices measured against what an
hour of chair time has to earn, commission that is calculated the way
you agreed it rather than the way somebody remembers it, and a monthly
profit figure you can act on in the month rather than finding out about
in March.

---

## 6. What is deliberately *not* in version 1

* online booking, reminders and two-way sync with a booking platform;
* a card reader or payment link integration;
* macro-driven features — one-click rebooking, bulk SMS, a client app;
* rota and shift planning with holiday accrual and clocking in;
* a full stock engine with batch tracking and supplier EDI;
* multi-site consolidation and inter-salon transfers.

---

© Novality Store. Generated workbooks are licensed for personal use and
for a single business.
