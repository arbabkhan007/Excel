# Ultimate Inventory & Wholesale Order Manager — generator

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
>   this README, `build_inventory_workbook.py`, 23 tabs.
> * **Ultimate Landscaping & Lawn Care Business Manager** —
>   [README_LANDSCAPING.md](README_LANDSCAPING.md),
>   `build_landscaping_workbook.py`, 23 tabs.
>
> Every download is listed in one place in
> [DOWNLOADS.md](DOWNLOADS.md).

Python code that builds a finished, ready-to-sell Excel workbook for
anybody who buys stock and sells it: the products, the SKUs, the
suppliers, the purchase orders, the stock that is actually on the
shelf, the wholesale customers and the price breaks you give them,
the orders you have promised, the sales across every channel — and
the question most stock spreadsheets cannot answer, **which products
are making you money and which ones are just sitting there**.

Every tab is joined to every other tab by a **SKU**, so a product is
one record from the purchase order that brought it in to the profit
line that says whether it was worth it.

> **Listing title:** Inventory Management Spreadsheet | SKU Stock
> Tracker, Supplier Database, Purchase Orders, Wholesale Orders &
> Profit Dashboard
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_inventory_workbook.py
# -> output/Ultimate Inventory & Wholesale Order Manager.xlsx
#    (23 tabs, 136,202 formulas, ~2.1 MB)
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
| 1 | **Dashboard** | page | 393 | The centrepiece. What the stock is worth, what is out of stock, what needs ordering today, what is reserved and unpicked, what is late — and a matrix that counts every alert on seventeen registers. |
| 2 | **Getting Started** | page | 0 | The one thing to understand (the SKU), the first hour in six steps, the five workflows, and an honest list of what this is *not*. |
| 3 | **Settings & Lists** | page | 9 | Your business, your markup and margin floor, your lead times and safety stock, your selling fees, your targets, and the 40 drop-down lists everything else reads. |
| 4 | **Product & SKU Database** | register | 4,200 | The catalogue. One row per separately stocked variant, with its own SKU, its cost, its three prices and a duplicate-SKU alarm. |
| 5 | **Inventory Tracker** | register | 9,900 | What you actually have. **No cell on this tab can be typed in** — every figure is derived from the transaction registers. |
| 6 | **Stock Movement Log** | register | 14,400 | Damage, loss, samples, production use, transfers and corrections. Everything that moves stock and is not a purchase, a sale or a return. |
| 7 | **Supplier Database** | register | 960 | Who you buy from, what they promise, and the lead time they actually deliver. |
| 8 | **Purchase Orders** | register | 15,000 | Line by line, with freight landed into the unit cost, part deliveries, invoice checks and a duplicate-receiving guard. |
| 9 | **Wholesale Customer CRM** | register | 2,160 | Trade accounts: tier, terms, credit limit, what they owe and what they are worth. |
| 10 | **Wholesale Pricing** | register | 5,700 | Your price breaks, checked against weighted average cost and against your own floor price. |
| 11 | **Wholesale Orders** | register | 21,000 | The biggest register. Confirming a line **reserves** the stock, and the sheet tells you what you are short before you promise it. |
| 12 | **Retail Sales Tracker** | register | 18,000 | Every sale across ten channels, with the marketplace and payment fees taken out so the profit column is money that arrived. |
| 13 | **Reorder Planner** | register | 9,900 | Sales rate × real lead time + safety stock = reorder point, a suggested quantity rounded to the case, and the date you have to order by. |
| 14 | **Inventory Valuation & COGS** | register | 8,400 | One costing method, weighted average, and the ABC classification that says where your money is. |
| 15 | **Returns & Damaged Goods** | register | 3,000 | What came back, why, and whether its condition puts it back on the shelf or writes it off. |
| 16 | **Warehouse Locations** | register | 2,200 | One SKU across several bins without losing the total, with a gap column when they disagree. |
| 17 | **Stock Audit** | register | 4,500 | A count never overwrites stock. It raises a variance, and you post the adjustment yourself. |
| 18 | **Product Performance** | register | 10,500 | The catalogue ranked by profit, by return on the money tied up in it, and by how long it has been sitting there. |
| 19 | **Expense Tracker** | register | 4,800 | Every cost — with stock purchases deliberately kept out of the operating total. |
| 20 | **Profit Analysis** | register | 364 | Twelve months: revenue both ways, COGS, fees, expenses, net profit, and the cash effect of what you bought. |
| 21 | **Monthly Reports** | page | 178 | Pick a month. Revenue, margin, top products, channel split, stock movement, against target. |
| 22 | **Annual Reports** | page | 317 | Twelve months in a grid, the year's best and worst products, and the categories behind them. |
| 23 | **Printable Orders & Invoices** | page | 321 | Type a PO number or an order number and get a purchase order, an invoice and a picking list laid out to print. |

**136,202 formulas** in total, every one of them locked, on a 2.1 MB
file that opens in Excel 2010.

### The five workflows it was built around

1. **Product → SKU → stock tracking.** The catalogue mints a Product ID
   and gives every separately stocked variant its own SKU; the
   Inventory Tracker then derives what you have from what actually
   happened, and keeps reserved stock apart from available stock.
2. **Supplier → purchase order → receiving.** Suppliers carry a quoted
   lead time; purchase orders record what was received and when, land
   the freight into the unit cost, compare the invoice to the order,
   and flag the same line twice before a delivery can be counted
   twice. The supplier's real lead time is then measured against the
   one they quoted.
3. **Wholesale customer → tiered pricing → fulfilment.** The CRM holds
   the tier and the terms; Wholesale Pricing holds the breaks you set;
   an order line picks up the right price automatically, reserves the
   stock on confirmation, warns you what you are short, and ages the
   invoice afterwards.
4. **Stock movement → reorder alert → replenishment.** Sales, damage
   and adjustments move the stock figure; the Reorder Planner turns the
   resulting sales rate and the supplier's *actual* lead time into a
   reorder point, a suggested order quantity and an order-by date.
5. **Sales → COGS → product-level profitability.** Retail and wholesale
   revenue meet weighted average cost on the valuation tab, and Product
   Performance reports profit, margin, return on stock and stock turns
   for every SKU you sell.

### The SKU

One decision holds the whole product together: **every separately
stocked variant gets its own SKU** — size, colour, packaging, each one
its own row. It is minted once on the Product & SKU Database and
chosen from a drop-down everywhere else. Sixteen registers join back
through it. `tools/validate_inventory_workbook.py` checks that join in
both directions on every build, because the moment a SKU can be
mistyped the purchase order, the stock figure, the wholesale line and
the profit ranking stop being the same product.

### The five things this product is not allowed to get wrong

Each of these is a way an inventory spreadsheet quietly starts lying to
its owner, and each one is checked against the real formulas on every
build:

1. **Stock is derived, never typed.** The Inventory Tracker has no
   input columns at all. To correct a count you post an adjustment on
   the Stock Movement Log, so the number always has a reason behind it.
2. **Reserved is not the same as shipped.** A confirmed order takes
   units out of *available* without pretending they have gone, so the
   same ninety candles cannot be promised to two shops.
3. **The IDs connect.** Product ID, SKU, Supplier ID, Purchase Order
   ID and Order ID are generated or chosen from a list, never typed.
4. **Receiving twice cannot double-count.** A purchase order line that
   appears twice for the same SKU is flagged before the delivery is
   counted.
5. **One costing method.** Weighted average, everywhere — the
   valuation, the retail margin, the wholesale line profit and the
   profit-by-product ranking all read the same number.

---

## 1b. Download it

You do not have to run anything. The finished workbooks are in
[`download/`](download/) and indexed in [DOWNLOADS.md](DOWNLOADS.md):

| File | What it is |
|------|------------|
| `inventory-wholesale-order-manager.xlsx` | Filled in with a worked example — one supply business, mid-year, mid-mess. |
| `inventory-wholesale-order-manager-blank.xlsx` | Empty, ready for your own stock. |
| `inventory-wholesale-order-manager-source.zip` | The generator source. |

### The example business

The sample file is Thornfield Supply Co., which buys and makes
homeware and sells it three ways — a marketplace shop, a handful of
wholesale stockists and the occasional market stall. The year it is
having is deliberately not a tidy one:

* a bud vase **out of stock for nine days**, cleared out by a trade
  order, while the purchase order that would refill it **sits unsent**;
* a supplier whose deliveries are reliably about a week late and whose
  **quoted lead time still says seven days**;
* a wholesale order **confirmed three weeks ago and never picked**,
  quietly reserving ninety units the whole time;
* another order **confirmed for more units than exist**, with the
  shortfall counted to the unit;
* a delivery where the **invoice and the order value disagree**;
* a pallet where **more arrived than was ordered**;
* a clearance line whose **hundred-plus price has slipped below its own
  floor**;
* a stock count that found **eleven candles missing** with no
  explanation, posted as an adjustment so the shelf and the system
  agree again;
* two more counts **still unposted**, which is what the dashboard
  counts;
* and a stock purchase sitting in the expense tracker, **deliberately
  excluded from the operating total**.

A sample company where everything reconciles teaches nobody anything.
This one teaches you what each alert column is for.

---

## 2. Command line

```bash
python build_inventory_workbook.py --help
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Ultimate Inventory & Wholesale Order Manager.xlsx` | Where to write the file. |
| `--blank` | off | No example data — just the empty workbook. |
| `--no-protect` | off | Leave the sheets unlocked. |
| `--password TEXT` | `premium` | Sheet protection password. |
| `--author TEXT` | `Novality Store` | Document author. |
| `--business TEXT` | `Thornfield Supply Co.` | Your business name on the cover. |
| `--owner TEXT` | `Rosa Thornfield` | The owner. |
| `--currency SYMBOL` | `$` | Currency symbol. |
| `--currency-code CODE` | `USD` | Three letter code. |
| `--markup NUMBER` | `2.5` | The markup a new product starts at. |
| `--min-margin PERCENT` | `25` | The margin you refuse to go under. |
| `--lead-time DAYS` | `14` | Lead time when a supplier has none set. |
| `--safety-days N` | `10` | Days of safety stock in the reorder point. |
| `--sales-window N` | `90` | Days the sales rate averages over. |
| `--marketplace-fee P` | `6.5` | Marketplace commission percentage. |
| `--payment-fee P` | `2.9` | Payment processing percentage. |
| `--tax PERCENT` | `0` | Sales tax or VAT on sales. |
| `--revenue-target AMT` | `216000` | Revenue target for the year. |
| `--margin-target P` | `55` | The gross margin you are aiming at. |
| `--rows N` | scaled | Rows of capacity in the big registers. |
| `--today YYYY-MM-DD` | today | Pretend it is this date, for the example. |

```bash
# a blank UK boutique, VAT registered, pounds, tighter margins
python build_inventory_workbook.py --blank \
  --business "Harbour Lane Home" --owner "Nia Ellis" \
  --currency "£" --currency-code GBP \
  --tax 20 --min-margin 30 --margin-target 60 \
  -o "output/harbour-lane.xlsx"

# a small file for testing - every register scaled down
python build_inventory_workbook.py --rows 20 -o /tmp/small.xlsx
```

---

## 3. How the code is organised

```
build_inventory_workbook.py          the one-liner entry point
inventory_wholesale_spreadsheet/
    cli.py           argument parsing, the summary it prints
    config.py        every price, status set, threshold and disclaimer
    layout.py        Col / SheetSpec / the {token} template engine
    formats.py       every cell format, locked and unlocked
    writer.py        SheetWriter: banners, KPIs, grids, protection
    names.py         227 dynamic named ranges that grow with the data
    sample_data.py   the worked example, checked twice before it ships
    build.py         the 23 tab order and the six page builders
    sheets/
        common.py        the column factories every register shares
        products.py      ... and the seventeen registers
        inventory.py
        ...
        dashboard.py     ... and the six pages
        getting_started.py
        settings.py
        monthly.py
        annual.py
        printables.py
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
pyflakes inventory_wholesale_spreadsheet/ tools/
python tools/check_inventory_specs.py        # before a workbook exists
python tools/check_inventory_sample.py       # the worked example
python build_inventory_workbook.py -o /tmp/inv.xlsx
python tools/check_parens.py /tmp/inv.xlsx
python tools/validate_inventory_workbook.py /tmp/inv.xlsx
python tools/recalc_check.py /tmp/inv-small.xlsx   # really evaluates it
```

**`check_inventory_specs.py`** reads the 17 register specifications
without building anything and catches, in under a second, the five
mistakes that are expensive to find later: brackets that do not
balance, a `{token}` naming no column on that sheet, a
`[[Sheet.column]]` pointing at a column that does not exist, `a+b/c`
where `(a+b)/c` was meant, a COUNTIFS criteria range carrying its own
test, and a numeric function applied to a possibly-blank column from
inside `AND()` — which evaluates every argument and turns the whole
cell into `#VALUE!`. It reports 569 columns, 376 formula columns and
427 cross-sheet references.

**`check_inventory_sample.py`** checks the worked example against the
specs: every key is an *input* column (a value written into a formula
column is dropped silently and takes the formula with it), every
drop-down value exists in its list, and **every SKU on every register
resolves to a real catalogue product**.

**`validate_inventory_workbook.py`** opens the finished file and checks
104 links in the automation chain, the SKU join in both directions, the
five technical decisions above, sheet protection and cell locking, every
defined name a drop-down reads, and that no cell anywhere claims a
marketplace connection the product does not have.

**`recalc_check.py`** really evaluates every formula with the
`formulas` engine and fails on any `#VALUE!`, `#REF!` or `#DIV/0!`. It
runs on the sample *and* the blank file, because a formula that works
with data in it and breaks without is still broken.

---

## 5. What a spreadsheet cannot do

Said plainly on the Getting Started tab, and enforced by the validator:

* **It does not connect to Etsy, Shopify, Amazon, eBay or anything
  else.** You enter your orders or paste them in from an export. There
  is no live sync, there never was one in a `.xlsx` file, and any
  listing that claims otherwise is describing software you have to log
  into.
* **It does not scan barcodes or talk to a till.** There is a barcode
  column so you can store the number; reading one needs hardware.
* **It is not multi-user in real time.** One person has the file open
  at a time, or it lives on a drive that syncs.
* **It is not accounting software.** Weighted average costing is a
  defensible method and the one used throughout, but your jurisdiction
  may require something else for your accounts. Check with an
  accountant.
* **It does not file your tax return.** The tax category column
  organises your costs for whoever does.

---

## 6. What is deliberately *not* in version 1

* **Serial and batch numbers.** A batch-traced product needs a row per
  batch with its own expiry, which is a different data model.
* **Multi-currency purchasing.** You set one currency. Buying in two
  needs a rate per purchase order and a revaluation policy.
* **Assemblies and bills of materials.** The movement log has a "used
  in production" type so components can be consumed honestly, but there
  is no recipe that explodes one build into its parts.
* **Consignment and drop-ship stock you do not own.** It would need a
  separate ownership flag everywhere stock is counted.

---

## 7. The other inventory products

The same engine, re-pointed, gives four closely related products that
share this one's spine: a **Retail Store Inventory Manager**, a
**Handmade Seller Stock & Materials Tracker**, a **Wholesale Business
Manager** weighted towards trade accounts, and an **E-commerce
Multi-Channel Inventory Manager**. Each is the registers in `sheets/`
with different emphasis, not a rewrite.
