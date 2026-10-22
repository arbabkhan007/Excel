# Handmade Product Pricing & Profit Calculator — generator

> **This repository holds six spreadsheet products** built on the same
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
> * **Cleaning Business CRM & Pricing Calculator** —
>   [README_CLEANING.md](README_CLEANING.md),
>   `build_cleaning_workbook.py`, 37 tabs.
> * **Handmade Product Pricing & Profit Calculator** — this README,
>   `build_handmade_workbook.py`, 40 tabs.
> * **Salon Business Management Spreadsheet** —
>   [README_SALON.md](README_SALON.md),
>   `build_salon_workbook.py`, 20 tabs.

Python code that builds a finished, ready-to-sell Excel workbook for a
maker: what the materials really cost, what the hours are really worth,
what the marketplace really takes, and therefore what the thing has to
sell for. Six genuinely different craft calculators, a wholesale price
that still leaves a profit, and a dashboard that says which products
are quietly losing money.

> **Listing title:** Handmade Product Pricing Calculator, Material and
> Labor Cost Tracker, Profit Margin and Wholesale Price Spreadsheet
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_handmade_workbook.py
# -> output/Handmade Product Pricing & Profit Calculator.xlsx
#    (40 tabs, ~1.7 MB)
```

The repository also contains the XlsxWriter library itself
(`xlsxwriter/`), so the command above runs with **no installation and no
dependencies** from the repository root. Outside this repo,
`pip install xlsxwriter` first — the generator is tested against the
bundled 0.3.0 and against XlsxWriter 3.2.

---

## 1. What gets built

40 tabs, 97,600+ formulas, all joined up and all driven by drop downs.

| # | Tab | What it does |
|---|-----|--------------|
| 1 | **Start Here** | The five minute set-up in order, the pricing formula written out with a worked example, a map of every tab, the weekly and monthly rhythm, the colour key, and an honest note on what the workbook can and cannot do |
| 2 | **Settings** | Your shop details, the year, your hourly rate and what an employee really costs, productive hours a month, waste and defect allowances, packaging and postage defaults, the overhead method, **every marketplace fee as an editable number**, target margins, the rounding rule, your targets and the ID prefixes — around 70 named settings the rest of the workbook reads |
| 3 | **Dashboard** | The business in one screen: twelve KPI tiles (revenue, profit, margin, what your hours earned, fees paid, postage absorbed), a fourteen-line "needs your attention" list that counts the real problems, best and worst by four different measures, profit by category, twelve months with sparklines, the channel mix, and one sentence telling you the single most useful thing to do this week |
| 4 | **Products** | The catalog, and the spine of the file. Name, SKU (built for you), category, status, season, difficulty, channel, packaging profile, postage band, weight, your price — and materials, labour, packaging, postage absorbed, overhead, total cost, fees, the suggested price at your target margin, profit, margin, markup, profit an hour, what you are really paying yourself, the wholesale price, the competitor gap, units sold, revenue, actual profit, days since it was last priced and a plain-English action |
| 5 | **Materials** | Every supply with its pack size, pack price, delivery, trade discount and waste allowance — and the cost a unit, the cost per gram/millilitre/centimetre so two suppliers with different pack sizes can be compared honestly, stock value, days of cover, the price change since last time, and how much of it your products actually consume |
| 6 | **Material Usage** | The bill of materials: product, material, quantity, waste — line cost, share of the product, and the parent/line key the printable costing sheet reads |
| 7 | **Labor** | Minutes, not guesses. Stage, who does it, rate override, minutes, per item or per batch, batch size — the rate with on-costs, hours, stage cost, cost a unit and its share of the product |
| 8 | **Packaging** | Mailer, tissue, card, sticker, ribbon, filler — pack size, pack price, how many per order and the weight each, with four profiles (Budget, Standard, Gift, Premium) that the catalog picks from by name |
| 9 | **Shipping** | Postage bands: service, zone, weight limit, what the label costs, extra packaging, what the buyer is charged — and what you absorb on every order, per gram, and how many products use it |
| 10 | **Overhead** | Rent, insurance, subscriptions, software, tools, the lot — monthly and yearly, with the share of your total, per productive hour, per unit made, and how many units it takes to cover it |
| 11 | **Channels** | **Editable marketplace fees.** Listing fee and how many sales a listing makes, transaction %, processing % + fixed, offsite ads and how many orders really come through them, regulatory and currency fees, anything else — rolled into the two numbers the pricing formula needs (`r` and `F`), with the effective take on a $20 sale, the fees you actually paid, and the date you last checked the official schedule |
| 12 | **Retail Calculator** | **The headline feature.** Pick a product and a channel, see the cost built up line by line, then the price at your target margin — `Price = (Cost + Fixed fees) / (1 − margin − fee rate)` — with the rounding rule applied, a margin ladder from 30% to 70%, a price ladder showing the margin and hourly return at each price point, a comparison against the keystone and hourly-rate methods, and a verdict on the price you are charging today |
| 13 | **Wholesale Calculator** | The trade price from the other end: your retail, their discount, the quantity, the carriage — does the wholesale price still clear your cost? With the minimum retail that makes the discount work, the order value, the margin at that price, the minimum order check and a line-by-line "what a stockist costs you" |
| 14 | **Profit & Break-Even** | Contribution a unit, fixed costs a month, **units to break even (rounded up, and it says so when the contribution is zero or negative)**, the hours that takes, the revenue it needs, every product ranked by contribution, and where you actually are this month |
| 15 | **Crochet** | Yarn by the gram with waste, a second yarn, notions, stuffing, the pattern cost spread over the pieces it makes, hours hooking plus finishing minutes — skeins needed, leftovers, cost per gram and the price it suggests |
| 16 | **Jewelry** | Metal by the gram at today's price, scrap and how much of it you recover, stones, chain and wire by the centimetre, findings, solder, plating, hallmarking, and design/bench/polishing minutes separately |
| 17 | **Candles** | Vessel size × fill × density = wax, plus melt loss, the fragrance load as a percentage of the wax, wicks, vessel, dye, label — poured in batches, so the batch minutes are spread over the batch, with cost per 100 g and the burn hours the buyer gets |
| 18 | **Clothing** | Fabric by the centimetre with layout waste, lining, interfacing, thread, notions, labels — scaled by a size factor, with cutting, sewing, finishing, pressing and fitting minutes kept apart |
| 19 | **Woodworking** | Timber by the centimetre with the offcut waste nobody admits to, a second timber, hardware, glue, finish by the millilitre, abrasives, and machine time charged at a machine rate on top of your own |
| 20 | **Digital** | Design, writing, testing and photography hours spread over the sales you expect, software and tester costs, support minutes a sale, yearly updates — recovery so far, the sales it still needs to break even and the margin once it has |
| 21 | **Bundles** | Gift sets and multi-buys: the parts, the assembly minutes, the discount off buying separately — what the set costs, what it should sell for, what the customer saves and how much margin the bundle gives away |
| 22 | **Bundle Lines** | What is in each set, with the per-line cost and price, and the key the bundle totals read |
| 23 | **Variations** | Sizes, colours, personalisation and rush orders: the extra materials and minutes against the base product, the size and rush factors, the price it justifies and the uplift you are actually charging |
| 24 | **Batches** | Planned against made, seconds and what they sold for, materials and hours actually used — cost a good unit, the variance against the estimate, and the time variance that tells you the costing is optimistic |
| 25 | **Sales** | Every order: product, quantity, channel, price, discount, promo code, postage charged and paid, the fee the marketplace really took, refunds — with revenue, cost of goods, postage absorbed, estimated fees, the gap between estimated and actual, profit, margin and profit an hour |
| 26 | **Promotions** | Codes and sales with their value, scope, dates, minimum spend — the margin after the discount, the share of profit given away, **the extra sales volume needed to stand still**, and what it actually did |
| 27 | **Price History** | Every price change with the reason, the cost at the time, the margin before and after, and the units sold in the thirty days either side — so "did the price rise cost me sales?" has an answer |
| 28 | **Competitors** | What other people charge, delivered price included, with quality and whether it is really handmade — your gap in money and percent, and a position verdict |
| 29 | **Waste Log** | Offcuts, miscuts, failed pours, breakages — what it cost, what share of your material spend it is, and how much of it was avoidable |
| 30 | **Development** | What a new design cost to develop — design, sample and photography hours, sample materials — against what it has recovered, how many more it needs to sell, and the per-unit charge if you want to recover it in the price |
| 31 | **Events** | Craft fairs: stall, travel, accommodation, prep and stall hours — the break-even sales, what you took, profit, return an hour and a verdict on going back |
| 32 | **Wholesale Customers** | Stockists with their terms, discount and minimum — orders, units, value, profit, margin, average order, days since the last one, what is outstanding and whether they are worth keeping |
| 33 | **Wholesale Orders** | Line sheet orders: quantity, the agreed price (or the discount applied to retail), carriage charged and paid, due and despatched dates — order value, cost, profit, margin, the hours it takes, whether it meets the minimum, and the lead time |
| 34 | **Scenarios** | Four what-ifs that write nothing back: every product priced at five market positions with a verdict on where yours sits, one product across every channel, a materials cost-shock table at 0/5/10/15/25/40%, and a batch-size and hourly-wage table |
| 35 | **Reports** | Six numbered reports: profit by product, material spend, what your hours are worth, a month-by-month profit and loss with overhead charged flat, your prices against the market, and the year against your targets |
| 36 | **Print Costing** | A one-page costing sheet for any product — the bill of materials, the labour stages, the cost tail and the price block — the thing to put in the folder or send to a stockist |
| 37 | **Print Quote** | A custom-order quote on one page: typed customer, eight product lines, deposit at 50% and the terms |
| 38 | **Print Order** | A wholesale line sheet priced at that customer's discount, with a blank quantity column, the minimum order check and a signature box |
| 39 | **Lists** | Every drop down in the workbook — 50 of them — plus the unit conversion, difficulty, rush, size and pricing-scenario factor tables, each a named range the rest of the file points at |
| 40 | **Help** | The colour code, fourteen troubleshooting rows, a fourteen-term glossary, nine real questions answered, a feature map and the small print |

### Product rules baked in

* **Price is not cost plus a bit.** The formula in the file, on the
  Start Here tab and in the listing is
  `Price = (Cost + Fixed fees) / (1 − target margin − fee rate)`, where
  `Cost = Materials + Labour + Packaging + Shipping + Overhead`.
  Dividing is the whole point: adding 30% to a cost and then letting
  Etsy take 12.8% of the *price* leaves nothing like 30%.
* **Marketplace fees are data, never constants.** Every rate lives on
  the Channels tab, every one is editable, each row carries the link to
  the official schedule and the date you last read it, and the
  dashboard nags when that date is older than the review period on
  Settings. The workbook ships with plausible 2026 rates and says, in
  several places, to verify them against the current official fee
  schedule for your country.
* **Two ways to cost a product, never both at once.** Either the
  generic bill of materials plus labour rows, or one of the six
  specialist tabs. Each specialist tab has a check column that goes red
  if the same product also has bill-of-materials lines, because that
  would double its cost.
* **Minutes, not vibes.** Labour is `minutes / 60 × rate`, batch stages
  are divided by the batch, and the rate carries on-costs for employees
  and contractors. Every calculator reports what you are actually
  paying yourself an hour — the number most makers have never worked
  out.
* **Every formula cell is locked**, every sheet is protected with
  `premium`, and the cream coloured cells are the only ones left open —
  3,485 of them in the filled copy, so there is plenty of room to type.
* **Excel 2010 safe.** `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `CEILING`,
  `ROUNDUP`, `EOMONTH`, `SUMPRODUCT` — no dynamic arrays, no `XLOOKUP`,
  no `IFS`, nothing that needs Ctrl+Shift+Enter. It opens in Google
  Sheets, LibreOffice and Numbers.
* **No charts, by design** — the vendored XlsxWriter cannot write them.
  Trends use sparklines, data bars and colour scales, which survive a
  Google Sheets import better anyway.
* **"Top five" lists without array formulas.** Every ranking is a
  hidden rank column on the register itself (`profit_rank`,
  `margin_rank`, `revenue_rank`, `profit_hour_rank`, `cost_used_rank`)
  and the page just does `MATCH(n, rank, 0)`.

---

## 1b. Download it

Everything is committed to the repository, so these links work straight
from GitHub — right click, *Save link as*.

| What | Size | Link |
|------|-----:|------|
| **The finished workbook, with the example shop filled in** | 1,702 KB | [handmade-pricing-calculator.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/handmade-pricing-calculator.xlsx) |
| **The finished workbook, empty and ready to sell** | 1,677 KB | [handmade-pricing-calculator-blank.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/handmade-pricing-calculator-blank.xlsx) |
| **The generator source, zipped** | 390 KB | [handmade-pricing-calculator-source.zip](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/handmade-pricing-calculator-source.zip) |

The example shop is **Willow & Thread Handmade** of Asheville: fourteen
products across crochet, jewelry, candles, sewing, woodworking and
digital patterns, 24 materials, 58 sales over five months, three craft
fairs, five stockists and ten wholesale orders — enough that every tab
means something the moment it opens. `--blank` strips all of it out and
leaves the formulas.

Rebuild the three files with:

```bash
python build_handmade_workbook.py -o download/handmade-pricing-calculator.xlsx
python build_handmade_workbook.py --blank -o download/handmade-pricing-calculator-blank.xlsx
python tools/export_source_md.py --product handmade
python tools/make_download_zip.py --product handmade
```

---

## 2. Command line

```bash
python build_handmade_workbook.py        [options]
python -m handmade_pricing_spreadsheet   [options]      # same thing
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Handmade Product Pricing & Profit Calculator.xlsx` | where to write the file |
| `--blank` | off | no example data, just the empty workbook |
| `--no-protect` | off | leave the sheets unlocked (for development) |
| `--password TEXT` | `premium` | sheet protection password |
| `--author TEXT` | `Novality Store` | document author |
| `--shop TEXT` | `Willow & Thread Handmade` | shop name on the cover and the printables |
| `--owner TEXT` | `Maya Albright` | the maker |
| `--currency SYMBOL` | `$` | currency symbol, e.g. `£`, `€`, `A$` |
| `--currency-code XXX` | `USD` | three letter code used in headings |
| `--year YYYY` | this year | the year the reports open on |
| `--rate AMOUNT` | 25 | what you pay yourself an hour |
| `--margin PERCENT` | 30 | target retail margin |
| `--wholesale PERCENT` | 40 | the discount a stockist gets off retail |
| `--overhead AMOUNT` | monthly total | overhead a month, if you would rather type it than list it |
| `--hours N` | 70 | productive hours a month — the hours you can actually sell |
| `--rounding RULE` | `Nearest 0.99` | how a calculated price is rounded |
| `--rows N` | 300 | rows of capacity in Products; every other register scales with it |
| `--today YYYY-MM-DD` | today | pretend it is this date (keeps the demo data reproducible) |
| `--quiet` | off | print only the output path |

```bash
# a blank British workbook, £ at 60% margin, rounding to .95
python build_handmade_workbook.py --blank --currency "£" \
    --currency-code GBP --rate 22 --margin 60 --rounding "Nearest 0.95" \
    -o "output/Handmade Pricing GBP.xlsx"
```

---

## 3. How the code is organised

```
build_handmade_workbook.py          one line wrapper
handmade_pricing_spreadsheet/
    config.py       every default: palette, row counts, rates, fees, factors
    layout.py       SheetSpec / Col and the formula template engine
    formats.py      the whole format book (styles x input/calc/plain)
    writer.py       SheetWriter - banners, nav, sections, KPI tiles, grids
    names.py        the cross sheet defined names
    sample_data.py  one coherent demo shop, dated from the build date
    build.py        tab order, page builders, the workbook assembly
    sheets/
        common.py settings.py lists.py            the setup tabs
        products.py materials.py material_usage.py
        labor.py packaging.py shipping.py         the cost registers
        overhead.py channels.py
        crochet.py jewelry.py candles.py          the six calculators
        clothing.py woodworking.py digital.py
        bundles.py bundle_lines.py variations.py  the selling registers
        batches.py sales.py promotions.py
        price_history.py competitors.py waste.py
        development.py events.py
        wholesale_customers.py wholesale_orders.py
        start_here.py dashboard.py                the pages
        retail_calc.py wholesale_calc.py
        break_even.py scenarios.py reports.py
        printables.py help_tab.py
tools/              the quality checks
```

A **register** is a data tab. Its module exposes `NAME`, a `spec(cfg)`
that returns a `SheetSpec` (a list of `Col`s with their formulas,
widths, drop downs, comments and conditional formats) and an optional
`decorate(sheet, fmts, cfg, reg)` for the summary strip underneath the
grid. A **page** is a hand-built tab — the dashboard, the calculators,
the printables — written directly with `SheetWriter`.

### Formula templates

Columns are written once, for one row, with placeholders:

```python
Col("margin", "Margin", 9, "percent",
    formula='=IF(OR({name}{r}="",{price}{r}=0),"",'
            '{profit}{r}/{price}{r})')
```

* `{r}` — the current row, `{r_prev}` — the row above
* `{key}` — the column letter of another column on the same sheet
* `{first}` / `{last}` — the first and last data row
* `{key_first}` / `{key_last}` — a whole column as `$D$5:$D$304`
* `[[Sheet.key]]` — an absolute range on another sheet,
  `[[Sheet.key#12]]` one cell of it
* helpers for the patterns Excel 2010 makes awkward: `max_if`, `min_if`,
  `blank_if`, `safe`, `lookup`, `min_number`, `round_price`

The registry renders them per row, so a column's logic lives in exactly
one place — and a typo in a sheet or column name is a build error, not
a `#REF!` a customer finds.

### The pricing chain

The one thing worth understanding before changing anything:

| Step | Where | What happens |
|------|-------|--------------|
| 1 | `Materials.unit_cost` | pack price less trade discount, plus delivery, over the pack quantity — then per gram, millilitre or centimetre |
| 2 | `Material Usage.line_cost` | quantity × (1 + waste) × cost a unit |
| 3 | `Labor.cost` | minutes ÷ 60 × rate, batch stages divided by the batch |
| 4 | the six craft tabs | the same two steps, done the way that craft actually works |
| 5 | `Products.mat_cost` / `labor_cost` | whichever of the two routes the product uses |
| 6 | `Packaging` / `Shipping` / `Overhead` | the per-order costs most makers forget |
| 7 | `Products.cost_base` | `C = M + L + P + S`, with overhead added in `total_cost` |
| 8 | `Channels.pct_total` / `fixed_total` | every marketplace fee rolled into `r` and `F` |
| 9 | `Products.suggested` | `(C + F) / (1 − r − margin)`, then the rounding rule |
| 10 | `Products.profit` / `margin` / `markup` | against the price you actually charge |
| 11 | `Products.wholesale` | retail less the trade discount, checked against cost |
| 12 | `Sales` → `Products.actual_profit` | what really happened, fees from the statement |
| 13 | `Dashboard` / `Reports` | the whole thing, read back |

`tools/validate_handmade_workbook.py` checks all eighty-two links of
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
* **A seventh craft calculator** — copy the closest of the six, keep
  the shared tail (`craft_tail()` gives you cost, fees, suggested
  price, profit, margin and hourly wage for free) and add the check
  column with `craft_check()`.
* **A "top five" list** — add a hidden rank column to the register, then
  `MATCH(n, rank, 0)` on the page. Never an array formula.
* **A KPI tile** — one `sheet.kpi(row, col, label, formula, style)` call.

---

## 4. Quality checks

```bash
python build_handmade_workbook.py -o output/test.xlsx
python tools/validate_handmade_workbook.py output/test.xlsx  # needs openpyxl
python tools/check_handmade_sample.py      output/test.xlsx

python build_handmade_workbook.py --rows 24 -o /tmp/small.xlsx
python tools/recalc_check.py /tmp/small.xlsx                 # needs formulas
```

`validate_handmade_workbook.py` reports: tab order, document author,
sheet protection and passwords, overlapping merges or hyperlinks
trapped inside merges (Excel calls these corrupt), unresolved template
tokens, `#REF!`, **unbalanced brackets** (quote-aware — a doubled `""`
inside a string is not the end of the string), odd quote counts, empty
arguments, unknown or too-modern functions, formulas pointing at a
missing sheet, defined names and drop down sources that do not exist,
column letters left without a row number, formula cells that are not
locked, **`a+b/c` where `(a+b)/c` was meant** (Excel divides before it
adds, so a total built from two `SUMIFS` has to be bracketed before it
is divided — the file opens, the brackets balance, the answer is just
wrong), the eighty-two links of the pricing chain, and the eleven live
formulas and ten typed inputs that make the calculator pages work.

`check_handmade_sample.py` re-reads the finished file and confirms the
example data joins up — every bill-of-materials line points at a real
product and a real material, every labour row at a real product, every
sale at a real product, channel, promo code and event, every wholesale
order at a real stockist, every postage option at a real band, every
material has a pack price — and that **nothing is costed twice**, once
through the bill of materials and again on a specialist tab, which
would silently double its cost.

`recalc_check.py` goes further and actually *runs* the workbook: it
loads it into the `formulas` calculation engine, evaluates the whole
dependency graph, then reads the answer out of every formula cell — so
a formula that parses perfectly but subtracts a blank on an empty row
is still caught. That is exactly what it found here: `INDEX()` on an
empty row hands back a zero, not a blank, so `1 − margin − ""` on the
Scenarios tab was a `#VALUE!` in the blank workbook. Both the filled
and the blank `--rows 24` builds now calculate clean, start to finish.

Current status: all three clean and `pyflakes` clean on the package and
the tools.

---

## 5. What a spreadsheet cannot do

Worth saying plainly, in the listing as well as on the Help tab:

* **it does not know this week's Etsy fees.** The rates ship as
  editable numbers with the official link next to them and the date
  they were last checked; when a marketplace changes its fees, you
  change one cell and the whole workbook re-prices. Nothing in here
  reads a website;
* **it cannot tell you what the market will pay.** It tells you what
  the piece costs, what it has to sell for to make your margin, and
  what your competitors charge. The decision in between is yours;
* **it does not track stock in real time.** Stock on hand is a number
  you type after a count; the workbook tells you when the days of
  cover run low;
* **it is not your bookkeeping.** It is costing and pricing. There is
  no tax return in here, no VAT return, no payroll — and every maker
  should check their own country's rules on what a hobby business must
  declare;
* **estimates are estimates.** Minutes from a costing are a starting
  point, not a promise — which is exactly why the Batches tab compares
  them with the hours you really spent, and usually finds the costing
  optimistic;
* **nothing here is financial advice.** It is arithmetic, done
  consistently, on numbers you supply.

What it does give you is a price you can defend, a wholesale price that
still pays you, an honest hourly wage for every product you make, and
an early warning on the ones that are quietly losing money.

---

## 6. What is deliberately *not* in version 1

The brief listed 1,080 features across 35 categories; v1 ships the ones
a maker touches every week. These are the natural premium / v2
additions:

* a live Etsy or Shopify import — orders, fees and listings pulled in
  rather than typed or pasted;
* macro-driven features: one-click "re-price everything after a
  material rise", bulk listing export, a photo-led catalog sheet;
* a proper stock engine — reservations, work in progress, components
  made from other components more than one level deep;
* multi-currency with live rates, and a sales-tax/VAT module for
  cross-border selling;
* a phone-friendly companion for craft fairs (Google Sheets edition
  with Apps Script, or a Form that appends to Sales);
* subscription-box and made-to-order workflow with deposits, due dates
  and customer messaging.

---

© Novality Store. Generated workbooks are licensed for personal use and
for a single business.
