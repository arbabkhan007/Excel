# Automated Meal Planner & Grocery List Spreadsheet — generator

> **This repository holds three spreadsheet products** built on the same
> engine, with the same conventions, the same author (Novality Store) and
> the same `premium` sheet protection:
>
> * **Pet Sitting & Dog Walking Business Spreadsheet** —
>   [README.md](README.md), `build_pet_sitting_workbook.py`, 27 tabs.
> * **Social Media Content Calendar & Analytics Planner** —
>   [README_SOCIAL.md](README_SOCIAL.md),
>   `build_social_media_workbook.py`, 43 tabs.
> * **Automated Meal Planner & Grocery List Spreadsheet** — this README,
>   `build_meal_planner_workbook.py`, 36 tabs.

Python code that builds a finished, ready-to-sell Excel workbook for
anyone who feeds a household: plan the week, keep the recipes, let the
shopping list write itself, and watch what the food actually costs.

> **Listing title:** Automated Meal Planner Spreadsheet, Weekly and
> Monthly Meal Planner, Recipe Organizer, Grocery List, Pantry
> Inventory, Food Budget and Nutrition Tracker
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_meal_planner_workbook.py
# -> output/Meal Planner & Grocery List.xlsx   (36 tabs, ~1.1 MB)
```

The repository also contains the XlsxWriter library itself
(`xlsxwriter/`), so the command above runs with **no installation and no
dependencies** from the repository root. Outside this repo,
`pip install xlsxwriter` first — the generator is tested against the
bundled 0.3.0 and against XlsxWriter 3.2.

The companion products in this repository are the
[Pet Sitting & Dog Walking Business Spreadsheet](README.md) and the
[Social Media Content Calendar & Analytics Planner](README_SOCIAL.md);
all three share the same engine, the same conventions and the same
quality tools.

---

## 1. What gets built

36 tabs, 72,800+ formulas, all joined up and all driven by drop downs.

| # | Tab | What it does |
|---|-----|--------------|
| 1 | **Start Here** | Welcome page, the ten minute set-up, the weekly rhythm, a map of every tab, the colour key and an honest note on what a spreadsheet can and cannot do |
| 2 | **Dashboard** | The week in one screen: 8 KPI tiles, a nine-row "needs your attention" list, this week's dinners, a 12-month table with sparklines and the year on one line |
| 3 | **Weekly Planner** | Seven days across, seven meal slots down. Pick any week with one date; cost, calories and the shopping total follow |
| 4 | **Monthly Planner** | The same plan as a six-week calendar grid, dinners in the day cells, today highlighted |
| 5 | **Settings** | Household, week start, shopping window, budgets, nutrition targets, expiry and rotation windows — every number the rest of the workbook reads |
| 6 | **Lists** | Every drop down in the workbook, plus the unit conversion table (unit → base unit → factor) |
| 7 | **Household** | One row per person: portion factor, diet, allergies, dislikes, who cooks, who shops |
| 8 | **Recipes** | The recipe book. Servings, timings, method, tags — and cost, calories and macros handed back from the ingredient lines |
| 9 | **Recipe Ingredients** | **The engine room.** One row per ingredient per recipe, converted to base units and multiplied by what the plan needs |
| 10 | **Ingredients** | The price book: pack size, pack price, cost per unit, nutrition per 100, allergens, storage, what is in the pantry and what is still needed |
| 11 | **Meal Plan** | One row per planned meal — the single source of truth behind both planners |
| 12 | **Grocery List** | **Writes itself.** What the plan needs, minus the pantry, rounded to whole packs, priced at the cheapest store, grouped by aisle |
| 13 | **Pantry** | Cupboard, fridge and freezer stock with value, expiry and a days-left countdown |
| 14 | **Staples** | The things you always keep in, with a par level and a "below par" flag |
| 15 | **Leftovers** | Portions in the fridge with a use-by, a value and a queue position |
| 16 | **Food Waste** | What got binned, why, what it cost and what would have stopped it |
| 17 | **Meal Prep** | Batch cooking sessions: batches, portions, minutes, cost per portion and which meals they cover |
| 18 | **Dietary Filters** | Six drop downs — diet, category, method, time, cost, allergen — and the recipes that survive all six, plus who in the house can eat what |
| 19 | **Nutrition** | Daily calories and macros against the targets, a comparison table, calories by meal slot and the top recipes for protein and fibre |
| 20 | **Budget** | Twelve months of planned against actual, split by spend type and by store, with a savings tracker |
| 21 | **Recipe Costs** | Cheapest meals first, cost per serving against a takeaway, the ingredients the money goes on, and cost by category |
| 22 | **Converter** | Unit conversion both ways, a quick reference, and a recipe scaler that rewrites every quantity for any number of people |
| 23 | **Rotation** | Variety: what you cook too often, what you have forgotten, the longest-untouched list and a suggested swap |
| 24 | **Food Safety** | A 14-allergen by household-member grid, recipes to check, fridge and freezer storage times, cooking temperatures and the leftovers rules |
| 25 | **Fridge & Freezer** | The cold storage list sorted by what needs eating first, the leftovers queue and a defrost planner |
| 26 | **Grocery Spending** | Every receipt. This is what the budget compares against |
| 27 | **Store Prices** | The same item at different shops, the cheapest one, and what switching would save |
| 28 | **Stores** | Where you shop: travel time, delivery, price level, what each is best for |
| 29 | **Events** | Dinners, parties and holidays with guests, portions and their own budget |
| 30 | **Goals** | A handful of food habits with a measure, a target and a progress bar |
| 31 | **Sustainability** | Why food gets binned, what it costs a year, the things you bin most and seven habits that change it |
| 32 | **Reports** | Pick a month: spend by type, the month in the kitchen, week by week, the year in one table and the shops ranked |
| 33 | **Print Week** | The week's meals on one sheet of A4 for the fridge door, with the prep jobs and the leftovers to eat |
| 34 | **Print Shopping** | The shopping list with tick boxes, grouped by aisle, with a running total |
| 35 | **Print Recipe** | Any recipe as a card, scaled to any number of people |
| 36 | **Help** | Set-up in order, what every tab is for, the automation chain in plain English, troubleshooting, the honest limits and the shortcuts |

### Product rules baked in

* **Nothing is typed twice.** An ingredient is priced once, a recipe is
  written once, a meal is planned once. The shopping list, the costs,
  the calories and the budget all read those three.
* **The shopping list is the product.** Settings sets a window →
  Meal Plan says what is being cooked → Recipes works out the batches →
  Recipe Ingredients converts to base units → Ingredients subtracts the
  pantry → Grocery List rounds to whole packs and prices it. Nine
  columns of ordinary formulas, no macros.
* **Every formula cell is locked**, every sheet is protected with
  `premium`, and the cream coloured cells are the only ones left open.
* **Excel 2010 safe.** `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `SMALL`,
  `LARGE`, `EOMONTH`, `SUMPRODUCT` — no dynamic arrays, no `XLOOKUP`,
  no `IFS`, nothing that needs Ctrl+Shift+Enter. It opens in Google
  Sheets, LibreOffice and Numbers.
* **No charts, by design** — the vendored XlsxWriter cannot write them.
  Trends use sparklines, data bars and colour scales, which survive a
  Google Sheets import better anyway.
* **"Top five" lists without array formulas.** Every ranking is a
  hidden rank column on the register itself (`need_rank`, `buy_rank`,
  `cost_rank`, `rotation_rank`, `queue_rank`, `todo_rank`,
  `fridge_rank`…) and the page just does `MATCH(n, rank, 0)`.
* **Weight and volume are never silently converted.** A spreadsheet
  cannot know the density of your flour; grams stay grams and
  millilitres stay millilitres.

---

## 1b. Download it

Everything is committed to the repository, so these links work straight
from GitHub — right click, *Save link as*.

| What | Size | Link |
|------|-----:|------|
| **The finished workbook, with the example household filled in** | 1,090 KB | [meal-planner-grocery-list.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/meal-planner-grocery-list.xlsx) |
| **The finished workbook, empty and ready to sell** | 1,063 KB | [meal-planner-grocery-list-blank.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/meal-planner-grocery-list-blank.xlsx) |
| **The generator source, zipped** | 306 KB | [meal-planner-grocery-list-source.zip](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/meal-planner-grocery-list-source.zip) |

The two `.xlsx` files are the product itself: 36 tabs, every formula
cell locked, every sheet protected with `premium`, author Novality
Store. Open the blank one, fill in **Settings** and **Household**, and
work down the **Start Here** tab.

The zip is for anyone who wants to rebuild or change it — unzip
anywhere and run:

```bash
pip install -r requirements.txt
python build_meal_planner_workbook.py
```

It holds the generator, the quality tools, this README,
`SOURCE_CODE_MEAL.md`, a `requirements.txt` and a plain
`HOW TO RUN.txt`. The vendored copy of XlsxWriter in this repository is
left out on purpose — the zip runs against a normal
`pip install XlsxWriter` (tested on 3.2.0). Rebuild everything after a
code change with:

```bash
python build_meal_planner_workbook.py -o download/meal-planner-grocery-list.xlsx
python build_meal_planner_workbook.py --blank -o download/meal-planner-grocery-list-blank.xlsx
python tools/export_source_md.py --product meal
python tools/make_download_zip.py --product meal
```

---

## 2. Command line

```bash
python build_meal_planner_workbook.py [options]
python -m meal_planner_spreadsheet    [options]      # same thing
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Meal Planner & Grocery List.xlsx` | where to write the file |
| `--blank` | off | no example data, just the empty workbook |
| `--no-protect` | off | leave the sheets unlocked (for development) |
| `--password TEXT` | `premium` | sheet protection password |
| `--author TEXT` | `Novality Store` | document author |
| `--household TEXT` | `The Harper Household` | name on the cover and the printables |
| `--cook TEXT` | `Maya Harper` | the main cook / planner |
| `--currency SYMBOL` | `$` | currency symbol, e.g. `£`, `€`, `A$` |
| `--currency-code XXX` | `USD` | three letter code used in headings |
| `--year YYYY` | this year | the year the planner is built for |
| `--week-start DAY` | `Monday` | `Monday` or `Sunday` — drives both planners |
| `--servings N` | 4 | default number of people a meal is cooked for |
| `--budget AMOUNT` | 150 | weekly food budget |
| `--units {Metric,Imperial}` | `Metric` | which unit set the drop downs lead with |
| `--rows N` | 100 | rows of capacity in the Meal Plan; every other register scales with it |
| `--today YYYY-MM-DD` | today | pretend it is this date (keeps the demo data reproducible) |
| `--quiet` | off | print only the output path |

```bash
# a blank British workbook for 2027, six people, £200 a week
python build_meal_planner_workbook.py --blank --year 2027 \
    --currency "£" --currency-code GBP --servings 6 --budget 200 \
    -o "output/Meal Planner 2027.xlsx"
```

---

## 3. How the code is organised

```
build_meal_planner_workbook.py      one line wrapper
meal_planner_spreadsheet/
    config.py       every default: palette, row counts, targets, budgets
    layout.py       SheetSpec / Col and the formula template engine
    formats.py      the whole format book (styles x input/calc/plain)
    writer.py       SheetWriter - banners, nav, sections, KPI tiles, grids
    names.py        the cross sheet defined names
    sample_data.py  one coherent demo household, from a fixed seed
    build.py        tab order, page builders, the workbook assembly
    sheets/
        common.py settings.py lists.py            the setup tabs
        members.py recipes.py recipe_lines.py     the food registers
        ingredients.py meal_plan.py grocery.py
        pantry.py staples.py leftovers.py waste.py
        prep.py spending.py store_prices.py stores.py
        events.py goals.py
        start_here.py dashboard.py weekly.py monthly.py
        filters.py nutrition.py budget.py recipe_costs.py
        converter.py rotation.py allergens.py fridge.py
        sustainability.py reports.py printables.py help_tab.py
tools/              the quality checks
```

A **register** is a data tab. Its module exposes `NAME`, a `spec(cfg)`
that returns a `SheetSpec` (a list of `Col`s with their formulas,
widths, drop downs, comments and conditional formats) and an optional
`decorate(sheet, fmts, cfg, reg)` for the summary strip underneath the
grid. A **page** is a hand-built tab — the dashboard, the planners, the
printables — written directly with `SheetWriter`.

### Formula templates

Columns are written once, for one row, with placeholders:

```python
Col("cost_serving", "Cost a serving", 13, "money",
    formula='=IF({servings}{r}=0,"",{cost}{r}/{servings}{r})')
```

* `{r}` — the current row, `{r_prev}` — the row above
* `{key}` — the column letter of another column on the same sheet
* `{first}` / `{last}` — the first and last data row
* `{key_first}` / `{key_last}` — a whole column as `$D$5:$D$104`
* `[[Sheet.key]]` — an absolute range on another sheet,
  `[[Sheet.key#12]]` one cell of it
* helpers for the patterns Excel 2010 makes awkward: `max_if`, `min_if`,
  `blank_if`, `safe`, `lookup`, `dynamic_name`, `min_number`

The registry renders them per row, so a column's logic lives in exactly
one place.

### The automation chain

The one thing worth understanding before changing anything:

| Step | Where | What happens |
|------|-------|--------------|
| 1 | `Settings.ShopFrom` / `ShopTo` | you choose the window |
| 2 | `Recipes.plan_scale` | how many batches of each recipe that window needs |
| 3 | `Recipe Ingredients.need_base` | every line × those batches, in grams / ml / each |
| 4 | `Ingredients.need_now` | total needed, minus `in_pantry` |
| 5 | `Ingredients.need_rank` | 1, 2, 3 … for everything still short |
| 6 | `Grocery List.ingredient` | `INDEX(names, MATCH(ROW()-4, need_rank, 0))` |
| 7 | `Grocery List.packs` → `est_cost` | round up to whole packs, price at the best store |

`tools/validate_meal_workbook.py` checks all seventeen links of that
chain on every build, so a well-meaning edit cannot quietly break the
headline feature.

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

---

## 4. Quality checks

```bash
python build_meal_planner_workbook.py -o output/test.xlsx
python tools/validate_meal_workbook.py output/test.xlsx   # needs openpyxl
python tools/check_meal_sample.py      output/test.xlsx

python build_meal_planner_workbook.py --rows 40 -o /tmp/small.xlsx
python tools/recalc_check.py /tmp/small.xlsx              # needs formulas
```

`validate_meal_workbook.py` reports: tab order, document author, sheet
protection and passwords, overlapping merges or hyperlinks trapped
inside merges (Excel calls these corrupt), unresolved template tokens,
`#REF!`, **unbalanced brackets** (quote-aware — a doubled `""` inside a
string is not the end of the string), odd quote counts, empty
arguments, unknown or too-modern functions, formulas pointing at a
missing sheet, defined names and drop down sources that do not exist,
column letters left without a row number, formula cells that are not
locked, and the seventeen links of the automation chain.

The bracket check earns its place: a deeply nested `IF()` built by
string concatenation can lose a closing bracket, and both Excel and
openpyxl will accept the file without a word.

`check_meal_sample.py` re-reads the finished file and confirms the
example data joins up — every recipe line points at a real recipe and a
real ingredient, every planned meal and prep job at a real recipe,
every pantry item, store price, staple and binned item at a real
ingredient, every store and cook at a real record, and that there is
enough of it (18 recipes, 98 lines, 80 planned meals over 28 days, 30
pantry items, 36 store prices) for the dashboard to show something.

`recalc_check.py` goes further and actually *runs* the workbook: it
loads it into the `formulas` calculation engine, evaluates the whole
dependency graph, then reads the answer out of every formula cell — so
a formula that parses perfectly but divides by zero on an empty row is
still caught. Check a `--rows 40` build rather than a full-size one:
the formulas are identical, there are just fewer rows of them, and
72,000 of them need far more memory than a laptop will give you (the
tool says so and stops unless you pass `--force`). Add
`--show "Dashboard!B17,Grocery List!O5"` to print particular cells.

Current status: all three clean and `pyflakes` clean. 17,154 formulas
in the `--rows 40` build were calculated and every one lands on a real
answer.

---

## 5. What a spreadsheet cannot do

Worth saying plainly, in the listing as well as on the Help tab:

* **automatic shopping lists need structured ingredients** — "a splash
  of oil" cannot be added up, "15 ml oil" can. The workbook makes that
  easy with drop downs and a unit table, but it cannot read your mind;
* **nutrition is only as good as the per-100 figures you type in**, and
  none of it is medical or dietary advice;
* **allergen and storage columns are memory aids, not guarantees** —
  always read the actual label;
* it **cannot send reminders, email you, or order the groceries** —
  that needs an app with a server behind it. It works when you open it;
* prices go stale; the store price tab is a snapshot of what you last
  saw, not a live feed.

What it does give you is a week of meals planned in ten minutes, a
shopping list that writes itself from that plan, and an honest number
for what feeding the household actually costs.

---

## 6. What is deliberately *not* in version 1

The brief listed 640 features across 28 categories; v1 ships the ones a
household uses every week. These are the natural premium / v2
additions:

* barcode or receipt scanning, and supermarket API price feeds;
* a shared phone-friendly companion (Google Sheets edition with Apps
  Script, or a Form that appends to the Grocery List);
* macro-driven features — one-click "roll this week forward", a proper
  recipe import, printable recipe books;
* density-aware weight ↔ volume conversion with a per-ingredient table;
* a nutrition database lookup instead of hand-typed per-100 figures;
* multi-week rolling plans and automatic plan generation from the
  rotation gap.

---

© Novality Store. Generated workbooks are licensed for personal use and
for a single business.
