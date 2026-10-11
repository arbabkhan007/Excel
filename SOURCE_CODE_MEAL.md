# Automated Meal Planner & Grocery List Spreadsheet — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 36 tab `.xlsx` workbook with over 72,000 formulas: a weekly and a monthly meal planner, a recipe book with costed ingredient lines, a grocery list that builds itself from the plan and deducts the pantry, fridge and freezer tracking, leftovers and waste, a food budget, recipe costing and store price comparison, unit conversion and recipe scaling, dietary filters, calories and macros, batch cooking, meal rotation, food safety, a dashboard, reports and three printable pages. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_meal_planner_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --household TEXT      household name on the cover and printables
  --cook TEXT           the main cook / planner
  --currency SYMBOL     currency symbol  (default: $)
  --year YYYY           the year the planner is built for
  --week-start DAY      Monday or Sunday
  --servings N          default number of people a meal feeds
  --budget AMOUNT       weekly food budget
  --units {Metric,Imperial}   which unit set the drop downs lead with
  --rows N              rows of capacity in the big registers
  --today YYYY-MM-DD    pretend it is this date (for the demo data)
```

---

## Contents

* [0. How it is put together](#0-how-it-is-put-together)
  * [The build pipeline](#the-build-pipeline)
  * [Describing a column](#describing-a-column)
  * [The rules every formula obeys](#the-rules-every-formula-obeys)
  * [Ranking columns instead of a hidden engine](#ranking-columns-instead-of-a-hidden-engine)
  * [Every tab at a glance](#every-tab-at-a-glance)
  * [Every file at a glance](#every-file-at-a-glance)
* [1. Run it](#1-run-it)
  * [`build_meal_planner_workbook.py`](#build_meal_planner_workbookpy)
* [2. The package](#2-the-package)
  * [`meal_planner_spreadsheet/__init__.py`](#meal_planner_spreadsheet__init__py)
  * [`meal_planner_spreadsheet/__main__.py`](#meal_planner_spreadsheet__main__py)
  * [`meal_planner_spreadsheet/cli.py`](#meal_planner_spreadsheetclipy)
  * [`meal_planner_spreadsheet/config.py`](#meal_planner_spreadsheetconfigpy)
  * [`meal_planner_spreadsheet/layout.py`](#meal_planner_spreadsheetlayoutpy)
  * [`meal_planner_spreadsheet/formats.py`](#meal_planner_spreadsheetformatspy)
  * [`meal_planner_spreadsheet/writer.py`](#meal_planner_spreadsheetwriterpy)
  * [`meal_planner_spreadsheet/names.py`](#meal_planner_spreadsheetnamespy)
  * [`meal_planner_spreadsheet/build.py`](#meal_planner_spreadsheetbuildpy)
  * [`meal_planner_spreadsheet/sample_data.py`](#meal_planner_spreadsheetsample_datapy)
* [3. The food registers](#3-the-food-registers)
  * [`meal_planner_spreadsheet/sheets/__init__.py`](#meal_planner_spreadsheetsheets__init__py)
  * [`meal_planner_spreadsheet/sheets/common.py`](#meal_planner_spreadsheetsheetscommonpy)
  * [`meal_planner_spreadsheet/sheets/lists.py`](#meal_planner_spreadsheetsheetslistspy)
  * [`meal_planner_spreadsheet/sheets/settings.py`](#meal_planner_spreadsheetsheetssettingspy)
  * [`meal_planner_spreadsheet/sheets/members.py`](#meal_planner_spreadsheetsheetsmemberspy)
  * [`meal_planner_spreadsheet/sheets/recipes.py`](#meal_planner_spreadsheetsheetsrecipespy)
  * [`meal_planner_spreadsheet/sheets/recipe_lines.py`](#meal_planner_spreadsheetsheetsrecipe_linespy)
  * [`meal_planner_spreadsheet/sheets/ingredients.py`](#meal_planner_spreadsheetsheetsingredientspy)
  * [`meal_planner_spreadsheet/sheets/meal_plan.py`](#meal_planner_spreadsheetsheetsmeal_planpy)
  * [`meal_planner_spreadsheet/sheets/grocery.py`](#meal_planner_spreadsheetsheetsgrocerypy)
* [4. The kitchen registers](#4-the-kitchen-registers)
  * [`meal_planner_spreadsheet/sheets/pantry.py`](#meal_planner_spreadsheetsheetspantrypy)
  * [`meal_planner_spreadsheet/sheets/staples.py`](#meal_planner_spreadsheetsheetsstaplespy)
  * [`meal_planner_spreadsheet/sheets/leftovers.py`](#meal_planner_spreadsheetsheetsleftoverspy)
  * [`meal_planner_spreadsheet/sheets/waste.py`](#meal_planner_spreadsheetsheetswastepy)
  * [`meal_planner_spreadsheet/sheets/prep.py`](#meal_planner_spreadsheetsheetspreppy)
  * [`meal_planner_spreadsheet/sheets/spending.py`](#meal_planner_spreadsheetsheetsspendingpy)
  * [`meal_planner_spreadsheet/sheets/store_prices.py`](#meal_planner_spreadsheetsheetsstore_pricespy)
  * [`meal_planner_spreadsheet/sheets/stores.py`](#meal_planner_spreadsheetsheetsstorespy)
  * [`meal_planner_spreadsheet/sheets/events.py`](#meal_planner_spreadsheetsheetseventspy)
  * [`meal_planner_spreadsheet/sheets/goals.py`](#meal_planner_spreadsheetsheetsgoalspy)
* [5. The pages](#5-the-pages)
  * [`meal_planner_spreadsheet/sheets/start_here.py`](#meal_planner_spreadsheetsheetsstart_herepy)
  * [`meal_planner_spreadsheet/sheets/dashboard.py`](#meal_planner_spreadsheetsheetsdashboardpy)
  * [`meal_planner_spreadsheet/sheets/weekly.py`](#meal_planner_spreadsheetsheetsweeklypy)
  * [`meal_planner_spreadsheet/sheets/monthly.py`](#meal_planner_spreadsheetsheetsmonthlypy)
  * [`meal_planner_spreadsheet/sheets/filters.py`](#meal_planner_spreadsheetsheetsfilterspy)
  * [`meal_planner_spreadsheet/sheets/nutrition.py`](#meal_planner_spreadsheetsheetsnutritionpy)
  * [`meal_planner_spreadsheet/sheets/budget.py`](#meal_planner_spreadsheetsheetsbudgetpy)
  * [`meal_planner_spreadsheet/sheets/recipe_costs.py`](#meal_planner_spreadsheetsheetsrecipe_costspy)
  * [`meal_planner_spreadsheet/sheets/converter.py`](#meal_planner_spreadsheetsheetsconverterpy)
  * [`meal_planner_spreadsheet/sheets/rotation.py`](#meal_planner_spreadsheetsheetsrotationpy)
  * [`meal_planner_spreadsheet/sheets/allergens.py`](#meal_planner_spreadsheetsheetsallergenspy)
  * [`meal_planner_spreadsheet/sheets/fridge.py`](#meal_planner_spreadsheetsheetsfridgepy)
  * [`meal_planner_spreadsheet/sheets/sustainability.py`](#meal_planner_spreadsheetsheetssustainabilitypy)
  * [`meal_planner_spreadsheet/sheets/reports.py`](#meal_planner_spreadsheetsheetsreportspy)
  * [`meal_planner_spreadsheet/sheets/printables.py`](#meal_planner_spreadsheetsheetsprintablespy)
  * [`meal_planner_spreadsheet/sheets/help_tab.py`](#meal_planner_spreadsheetsheetshelp_tabpy)
* [6. Quality checks (optional)](#6-quality-checks-optional)
  * [`tools/validate_meal_workbook.py`](#toolsvalidate_meal_workbookpy)
  * [`tools/check_meal_sample.py`](#toolscheck_meal_samplepy)
  * [`tools/recalc_check.py`](#toolsrecalc_checkpy)
  * [`tools/export_source_md.py`](#toolsexport_source_mdpy)
  * [`tools/make_download_zip.py`](#toolsmake_download_zippy)
* [Appendix A. Data dictionary](#appendix-a-data-dictionary)

---

## 0. How it is put together

Everything here is plain Python writing a `.xlsx` file through XlsxWriter.
There is no template workbook, no copy of an existing file and no manual
step: the spreadsheet is described in code, and the code is the only
source of truth for what ends up in a cell.

Two kinds of tab exist.

* A **register** is a data tab — a header row and hundreds of rows under
  it. Its module exposes `NAME`, a `spec(cfg)` that returns a
  `SheetSpec` (a list of `Col` objects) and, optionally, a
  `decorate(sheet, fmts, cfg)` that adds totals or notes underneath the
  grid. Nothing in a register module writes cells directly.
* A **page** is a hand-built tab — the dashboards, the reports, the
  printables, the welcome and help pages. Its module exposes
  `build(book, fmts, cfg, reg)` and writes with `SheetWriter` helpers:
  `banner`, `section`, `kpi`, `table_head`, `calc_row`, `note`,
  `bullets`, `field`, `merge`, `finish`.

### The build pipeline

`build_meal_planner_workbook.py` parses the command line into a `Config`, then
`meal_planner_spreadsheet.build.build_workbook(path, cfg)` does six things in order:

1. **Open the workbook** and stamp the document properties — title,
   author, company, keywords and a comment naming the protection
   password. The author is the same on every generated file.
2. **Build the `FormatBook`.** Every visual style the product uses is
   created once, up front: a style (text, date, money, percent, status,
   link, …) crossed with a mode (`INPUT` cream and unlocked, `CALC` grey
   and locked, `PLAIN`), plus the page chrome (titles, section bars, KPI
   tiles, notes) and the conditional formatting palette.
3. **Register every `SheetSpec` first**, before a single tab is written.
   This is what lets any sheet refer to any other sheet's columns by
   name regardless of which tab happens to be written first.
4. **Write the tabs in `SHEET_ORDER`.** Registers go through
   `build_register()` — `SheetWriter.grid(spec)` writes widths, formats,
   headers, formulas, sample rows, drop downs, header comments,
   conditional formats, the autofilter, frozen panes and the print
   setup — then the module's `decorate()` runs and the sheet is
   protected. Pages call their own `build()`.
5. **Define the live lookup names** (`ContentIDs`, `TeamNames`, …) from
   the registry. Each is an `OFFSET` range that grows and shrinks with
   the number of filled rows, so a drop down never shows blanks.
6. **Activate the first tab** and close the file - there is no hidden sheet to hide in this one.

### Describing a column

A column is declared once, for one row, with placeholders:

```python
Col("er", "Engagement rate", 11, "percent",
    formula='=IFERROR({engagements}{r}/{reach}{r},"")',
    cf=[{"type": "cell", "criteria": ">=", "value": "=TargetEngagement",
         "fmt": "good"}])
```

`Col` carries the key, the header text, the width, the style, the
formula template, the data validation, the header comment, the
conditional formats, whether the column is hidden and which outline
group it belongs to. `Registry.render()` expands the template per row:

| Token | Becomes |
|-------|---------|
| `{r}` / `{r_prev}` | this row number / the row above |
| `{key}` | the column letter of another column on the same sheet |
| `{first}` / `{last}` | the first and last data row of the sheet |
| `[[Sheet.key]]` | an absolute range on another sheet, e.g. `'Content Calendar'!$G$5:$G$604` |
| `[[Sheet.key#12]]` | one cell of that column |
| `[[Sheet.key#]]` | the same row as the formula being written |

Helpers cover the patterns Excel 2010 makes awkward — `max_if`,
`min_if`, `blank_if`, `safe`, `lookup`, `dynamic_name`, `quote_sheet`.

### The rules every formula obeys

* **Excel 2010 safe.** `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `SMALL`,
  `LARGE`, `RANK`, `EDATE`, `EOMONTH`, `SUMPRODUCT(MAX(...))`. No
  dynamic arrays, no `XLOOKUP`, no `IFS`, no `AGGREGATE`, nothing that
  needs Ctrl+Shift+Enter. The file opens in Excel 2010 and later, Google
  Sheets, LibreOffice and Numbers.
* **No charts.** The vendored XlsxWriter in this repository cannot write
  them, so trends are drawn with `add_sparkline`, data bars and
  three-colour scales — which also survive a Google Sheets import.
* **Nth-row lookups without array formulas.** Hidden composite key
  columns (`date + ROW()/100000`) plus `SMALL` and `INDEX/MATCH` answer
  "the third thing on 14 March" and "the next eight due" with ordinary
  formulas.
* **One number, one home.** Nothing is typed into two tabs. Every other
  appearance of a figure is a formula pointing back at the one place it
  was entered.
* **Locked where it matters.** Formula cells are locked, input cells are
  not, and every sheet is protected with the same password — so the
  maths cannot be typed over by accident but the cream cells stay open.

### Ranking columns instead of a hidden engine

There is no hidden `Calc` tab in this product. Anything a page needs to
rank - the next thing to buy, the recipe you have left alone longest,
the leftovers to eat first, the fridge item closest to its date - is a
**hidden rank column on the register itself**:

```python
Col("rotation_rank", "Rotation rank", 9, "int",
    formula='=IF({{name}}{{r}}="","",COUNTIFS({{sort_key_first}}:'
            '{{sort_key_last}},">"&{{sort_key}}{{r}})+'
            'COUNTIFS({{sort_key_first}}:{{sort_key}}{{r}},'
            '{{sort_key}}{{r}}))',
    hidden=True)
```

`COUNTIFS(whole column, ">" & this row) + COUNTIFS(up to this row, this
row)` is a dense rank that breaks ties by row order, so the ranks come
out 1, 2, 3 ... with no gaps and no duplicates. The page then asks for
the nth item with an ordinary `MATCH(n, rank_range, 0)` - a plain range
lookup, which is why there is **no array formula and nothing needing
Ctrl+Shift+Enter anywhere in the workbook**. The tempting alternative,
`LARGE(IF(...))` over a computed array, reads better in Python and
quietly returns the wrong answer in Excel 2010.

The ranks in use: `Ingredients.need_rank` (what to buy next),
`Grocery List.buy_rank` (position on the printed list),
`Ingredients.spend_rank` and `waste_rank`, `Recipes.cost_rank`,
`match_rank`, `rotation_rank`, `times_planned_rank`, `calories_rank`,
`protein_rank` and `fibre_rank`, `Pantry.fridge_rank` and
`freezer_rank`, `Leftovers.queue_rank`, `Meal Prep.todo_rank` and
`Stores.spend_rank`.


### Every tab at a glance

36 tabs in the order they appear in the finished workbook. "Rows" is the capacity of a register at the default `--rows` setting.

| # | Tab | Kind | Written by | Columns | Rows |
|---|-----|------|------------|--------:|-----:|
| 1 | **Start Here** | page | `meal_planner_spreadsheet/sheets/start_here.py` | - | - |
| 2 | **Dashboard** | page | `meal_planner_spreadsheet/sheets/dashboard.py` | - | - |
| 3 | **Weekly Planner** | page | `meal_planner_spreadsheet/sheets/weekly.py` | - | - |
| 4 | **Monthly Planner** | page | `meal_planner_spreadsheet/sheets/monthly.py` | - | - |
| 5 | **Settings** | page | `meal_planner_spreadsheet/sheets/settings.py` | - | - |
| 6 | **Lists** | page | `meal_planner_spreadsheet/sheets/lists.py` | - | - |
| 7 | **Household** | register | `meal_planner_spreadsheet/sheets/members.py` | 25 | 12 |
| 8 | **Recipes** | register | `meal_planner_spreadsheet/sheets/recipes.py` | 58 | 200 |
| 9 | **Recipe Ingredients** | register | `meal_planner_spreadsheet/sheets/recipe_lines.py` | 28 | 1200 |
| 10 | **Ingredients** | register | `meal_planner_spreadsheet/sheets/ingredients.py` | 41 | 300 |
| 11 | **Meal Plan** | register | `meal_planner_spreadsheet/sheets/meal_plan.py` | 33 | 500 |
| 12 | **Grocery List** | register | `meal_planner_spreadsheet/sheets/grocery.py` | 26 | 300 |
| 13 | **Pantry** | register | `meal_planner_spreadsheet/sheets/pantry.py` | 25 | 300 |
| 14 | **Staples** | register | `meal_planner_spreadsheet/sheets/staples.py` | 19 | 100 |
| 15 | **Leftovers** | register | `meal_planner_spreadsheet/sheets/leftovers.py` | 19 | 120 |
| 16 | **Food Waste** | register | `meal_planner_spreadsheet/sheets/waste.py` | 20 | 150 |
| 17 | **Meal Prep** | register | `meal_planner_spreadsheet/sheets/prep.py` | 25 | 120 |
| 18 | **Dietary Filters** | page | `meal_planner_spreadsheet/sheets/filters.py` | - | - |
| 19 | **Nutrition** | page | `meal_planner_spreadsheet/sheets/nutrition.py` | - | - |
| 20 | **Budget** | page | `meal_planner_spreadsheet/sheets/budget.py` | - | - |
| 21 | **Recipe Costs** | page | `meal_planner_spreadsheet/sheets/recipe_costs.py` | - | - |
| 22 | **Converter** | page | `meal_planner_spreadsheet/sheets/converter.py` | - | - |
| 23 | **Rotation** | page | `meal_planner_spreadsheet/sheets/rotation.py` | - | - |
| 24 | **Food Safety** | page | `meal_planner_spreadsheet/sheets/allergens.py` | - | - |
| 25 | **Fridge & Freezer** | page | `meal_planner_spreadsheet/sheets/fridge.py` | - | - |
| 26 | **Grocery Spending** | register | `meal_planner_spreadsheet/sheets/spending.py` | 23 | 400 |
| 27 | **Store Prices** | register | `meal_planner_spreadsheet/sheets/store_prices.py` | 17 | 300 |
| 28 | **Stores** | register | `meal_planner_spreadsheet/sheets/stores.py` | 24 | 40 |
| 29 | **Events** | register | `meal_planner_spreadsheet/sheets/events.py` | 23 | 40 |
| 30 | **Goals** | register | `meal_planner_spreadsheet/sheets/goals.py` | 18 | 40 |
| 31 | **Sustainability** | page | `meal_planner_spreadsheet/sheets/sustainability.py` | - | - |
| 32 | **Reports** | page | `meal_planner_spreadsheet/sheets/reports.py` | - | - |
| 33 | **Print Week** | page | `meal_planner_spreadsheet/sheets/printables.py` | - | - |
| 34 | **Print Shopping** | page | `meal_planner_spreadsheet/sheets/printables.py` | - | - |
| 35 | **Print Recipe** | page | `meal_planner_spreadsheet/sheets/printables.py` | - | - |
| 36 | **Help** | page | `meal_planner_spreadsheet/sheets/help_tab.py` | - | - |

### Every file at a glance

| File | Lines | What it is |
|------|------:|------------|
| [`build_meal_planner_workbook.py`](#build_meal_planner_workbookpy) | 19 | Build the Automated Meal Planner & Grocery List Spreadsheet |
| [`meal_planner_spreadsheet/__init__.py`](#meal_planner_spreadsheet__init__py) | 19 | Automated Meal Planner & Grocery List Spreadsheet |
| [`meal_planner_spreadsheet/__main__.py`](#meal_planner_spreadsheet__main__py) | 6 | ``python -m social_media_spreadsheet`` |
| [`meal_planner_spreadsheet/cli.py`](#meal_planner_spreadsheetclipy) | 139 | Command line front end |
| [`meal_planner_spreadsheet/config.py`](#meal_planner_spreadsheetconfigpy) | 241 | Configuration for the Automated Meal Planner & Grocery List Spreadsheet |
| [`meal_planner_spreadsheet/layout.py`](#meal_planner_spreadsheetlayoutpy) | 253 | Layout engine |
| [`meal_planner_spreadsheet/formats.py`](#meal_planner_spreadsheetformatspy) | 200 | Cell format factory |
| [`meal_planner_spreadsheet/writer.py`](#meal_planner_spreadsheetwriterpy) | 444 | Sheet writing helpers shared by every tab of the workbook |
| [`meal_planner_spreadsheet/names.py`](#meal_planner_spreadsheetnamespy) | 50 | Workbook level defined names that point at the register sheets |
| [`meal_planner_spreadsheet/build.py`](#meal_planner_spreadsheetbuildpy) | 161 | Workbook builder: puts every tab together in the right order |
| [`meal_planner_spreadsheet/sample_data.py`](#meal_planner_spreadsheetsample_datapy) | 855 | Demo data: four weeks in the life of the Harper household |
| [`meal_planner_spreadsheet/sheets/__init__.py`](#meal_planner_spreadsheetsheets__init__py) | 6 | One module per tab |
| [`meal_planner_spreadsheet/sheets/common.py`](#meal_planner_spreadsheetsheetscommonpy) | 206 | Small building blocks shared by the register sheets |
| [`meal_planner_spreadsheet/sheets/lists.py`](#meal_planner_spreadsheetsheetslistspy) | 286 | `Lists` - every drop down in the workbook, plus the unit table |
| [`meal_planner_spreadsheet/sheets/settings.py`](#meal_planner_spreadsheetsheetssettingspy) | 239 | The `Settings` tab - household setup, planning window, budget and targets |
| [`meal_planner_spreadsheet/sheets/members.py`](#meal_planner_spreadsheetsheetsmemberspy) | 132 | `Household` - who you are cooking for |
| [`meal_planner_spreadsheet/sheets/recipes.py`](#meal_planner_spreadsheetsheetsrecipespy) | 294 | `Recipes` - the recipe database |
| [`meal_planner_spreadsheet/sheets/recipe_lines.py`](#meal_planner_spreadsheetsheetsrecipe_linespy) | 195 | `Recipe Ingredients` - one line per ingredient in a recipe |
| [`meal_planner_spreadsheet/sheets/ingredients.py`](#meal_planner_spreadsheetsheetsingredientspy) | 230 | `Ingredients` - the price book every cost and shopping line reads |
| [`meal_planner_spreadsheet/sheets/meal_plan.py`](#meal_planner_spreadsheetsheetsmeal_planpy) | 177 | `Meal Plan` - every planned meal, one row at a time |
| [`meal_planner_spreadsheet/sheets/grocery.py`](#meal_planner_spreadsheetsheetsgrocerypy) | 285 | `Grocery List` - the shopping list that writes itself |
| [`meal_planner_spreadsheet/sheets/pantry.py`](#meal_planner_spreadsheetsheetspantrypy) | 188 | `Pantry` - what is actually in the house |
| [`meal_planner_spreadsheet/sheets/staples.py`](#meal_planner_spreadsheetsheetsstaplespy) | 123 | `Staples` - the things you never want to run out of |
| [`meal_planner_spreadsheet/sheets/leftovers.py`](#meal_planner_spreadsheetsheetsleftoverspy) | 144 | `Leftovers` - the half a lasagne in the fridge |
| [`meal_planner_spreadsheet/sheets/waste.py`](#meal_planner_spreadsheetsheetswastepy) | 131 | `Food Waste` - the uncomfortable tab |
| [`meal_planner_spreadsheet/sheets/prep.py`](#meal_planner_spreadsheetsheetspreppy) | 151 | `Meal Prep` - the Sunday afternoon plan |
| [`meal_planner_spreadsheet/sheets/spending.py`](#meal_planner_spreadsheetsheetsspendingpy) | 139 | `Grocery Spending` - what you actually spent |
| [`meal_planner_spreadsheet/sheets/store_prices.py`](#meal_planner_spreadsheetsheetsstore_pricespy) | 131 | `Store Prices` - the same thing, priced in three shops |
| [`meal_planner_spreadsheet/sheets/stores.py`](#meal_planner_spreadsheetsheetsstorespy) | 95 | `Stores` - where you shop |
| [`meal_planner_spreadsheet/sheets/events.py`](#meal_planner_spreadsheetsheetseventspy) | 113 | `Events` - birthdays, Christmas, the in-laws on Sunday |
| [`meal_planner_spreadsheet/sheets/goals.py`](#meal_planner_spreadsheetsheetsgoalspy) | 153 | `Goals` - what you are trying to change, and whether it is working |
| [`meal_planner_spreadsheet/sheets/start_here.py`](#meal_planner_spreadsheetsheetsstart_herepy) | 189 | `Start Here` - the welcome page and the six step set-up |
| [`meal_planner_spreadsheet/sheets/dashboard.py`](#meal_planner_spreadsheetsheetsdashboardpy) | 292 | `Dashboard` - the one page that answers 'are we alright?' Eight numbers across the top, a list of things... |
| [`meal_planner_spreadsheet/sheets/weekly.py`](#meal_planner_spreadsheetsheetsweeklypy) | 255 | `Weekly Planner` - the week on one page |
| [`meal_planner_spreadsheet/sheets/monthly.py`](#meal_planner_spreadsheetsheetsmonthlypy) | 156 | `Monthly Planner` - the month as a calendar |
| [`meal_planner_spreadsheet/sheets/filters.py`](#meal_planner_spreadsheetsheetsfilterspy) | 148 | `Dietary Filters` - 'what can I actually make tonight?' Six drop-downs at the top: diet, category, method,... |
| [`meal_planner_spreadsheet/sheets/nutrition.py`](#meal_planner_spreadsheetsheetsnutritionpy) | 208 | `Nutrition` - calories and macros, without the guilt |
| [`meal_planner_spreadsheet/sheets/budget.py`](#meal_planner_spreadsheetsheetsbudgetpy) | 250 | `Budget` - where the food money actually goes |
| [`meal_planner_spreadsheet/sheets/recipe_costs.py`](#meal_planner_spreadsheetsheetsrecipe_costspy) | 181 | `Recipe Costs` - what every meal actually costs to put on the table |
| [`meal_planner_spreadsheet/sheets/converter.py`](#meal_planner_spreadsheetsheetsconverterpy) | 205 | `Converter` - unit conversion and recipe scaling, on one page |
| [`meal_planner_spreadsheet/sheets/rotation.py`](#meal_planner_spreadsheetsheetsrotationpy) | 201 | `Rotation` - stop eating the same six dinners |
| [`meal_planner_spreadsheet/sheets/allergens.py`](#meal_planner_spreadsheetsheetsallergenspy) | 223 | `Food Safety` - allergens, storage and the things worth checking |
| [`meal_planner_spreadsheet/sheets/fridge.py`](#meal_planner_spreadsheetsheetsfridgepy) | 162 | `Fridge & Freezer` - what is in the cold bits, and how long you have |
| [`meal_planner_spreadsheet/sheets/sustainability.py`](#meal_planner_spreadsheetsheetssustainabilitypy) | 197 | `Sustainability` - the waste report, and what to do about it |
| [`meal_planner_spreadsheet/sheets/reports.py`](#meal_planner_spreadsheetsheetsreportspy) | 294 | `Reports` - the month in one page, and the year in one table |
| [`meal_planner_spreadsheet/sheets/printables.py`](#meal_planner_spreadsheetsheetsprintablespy) | 359 | The three printable pages |
| [`meal_planner_spreadsheet/sheets/help_tab.py`](#meal_planner_spreadsheetsheetshelp_tabpy) | 282 | `Help` - how the thing works, in the order you will need it |
| [`tools/validate_meal_workbook.py`](#toolsvalidate_meal_workbookpy) | 414 | Static checks on a generated Meal Planner workbook |
| [`tools/check_meal_sample.py`](#toolscheck_meal_samplepy) | 244 | Cross-check the example records in a generated meal planner workbook |
| [`tools/recalc_check.py`](#toolsrecalc_checkpy) | 216 | Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error... |
| [`tools/export_source_md.py`](#toolsexport_source_mdpy) | 922 | Write every source file of a generator into one markdown document |
| [`tools/make_download_zip.py`](#toolsmake_download_zippy) | 232 | Pack a generator into a single zip people can download and run |

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_meal_planner_workbook.py
```

The workbook lands in `output/Meal Planner & Grocery List.xlsx`.

### `build_meal_planner_workbook.py`

*19 lines*

> Build the Automated Meal Planner & Grocery List Spreadsheet.
>
>     python build_meal_planner_workbook.py
>     python build_meal_planner_workbook.py --blank -o "output/Blank.xlsx"
>     python build_meal_planner_workbook.py --rows 400 --currency "£"
>
> Everything lives in the `meal_planner_spreadsheet` package; this file is
> just a friendly front door so you do not have to remember `-m`.

````python
#!/usr/bin/env python3
"""Build the Automated Meal Planner & Grocery List Spreadsheet.

    python build_meal_planner_workbook.py
    python build_meal_planner_workbook.py --blank -o "output/Blank.xlsx"
    python build_meal_planner_workbook.py --rows 400 --currency "£"

Everything lives in the `meal_planner_spreadsheet` package; this file is
just a friendly front door so you do not have to remember `-m`.
"""

from __future__ import annotations

import sys

from meal_planner_spreadsheet.cli import main

if __name__ == "__main__":
    sys.exit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer, the cross sheet names and the demo household. These are the parts every tab leans on.

### `meal_planner_spreadsheet/__init__.py`

*19 lines*

> Automated Meal Planner & Grocery List Spreadsheet.
>
> Build the workbook with::
>
>     python -m meal_planner_spreadsheet -o "output/Meal Planner.xlsx"
>
> or from Python::
>
>     from meal_planner_spreadsheet import build_workbook
>     build_workbook("Meal Planner.xlsx")

````python
"""Automated Meal Planner & Grocery List Spreadsheet.

Build the workbook with::

    python -m meal_planner_spreadsheet -o "output/Meal Planner.xlsx"

or from Python::

    from meal_planner_spreadsheet import build_workbook
    build_workbook("Meal Planner.xlsx")
"""

from .build import build_workbook
from .config import (AUTHOR, PRODUCT_LISTING_TITLE, PRODUCT_NAME,
                     PRODUCT_SUBTITLE, PRODUCT_VERSION, Config)

__all__ = ["build_workbook", "Config", "AUTHOR", "PRODUCT_NAME",
           "PRODUCT_SUBTITLE", "PRODUCT_LISTING_TITLE", "PRODUCT_VERSION"]
__version__ = PRODUCT_VERSION
````

### `meal_planner_spreadsheet/__main__.py`

*6 lines*

> ``python -m social_media_spreadsheet``.

````python
"""``python -m social_media_spreadsheet``."""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
````

### `meal_planner_spreadsheet/cli.py`

*139 lines · 3 functions*

> Command line front end.
>
> python -m meal_planner_spreadsheet
> python -m meal_planner_spreadsheet -o "output/Meal Planner.xlsx" --blank
> python -m meal_planner_spreadsheet --currency "$" --year 2027 --rows 300

| Defines | What it does |
|---------|--------------|
| `build_parser()` |  |
| `config_from_args(args)` |  |
| `main(argv)` |  |

**Module constants:** `DEFAULT_OUTPUT` (os.path.join(...)).

````python
"""Command line front end.

    python -m meal_planner_spreadsheet
    python -m meal_planner_spreadsheet -o "output/Meal Planner.xlsx" --blank
    python -m meal_planner_spreadsheet --currency "$" --year 2027 --rows 300
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import AUTHOR, PRODUCT_LISTING_TITLE, PRODUCT_NAME, Config

DEFAULT_OUTPUT = os.path.join(
    "output", "Meal Planner & Grocery List.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="meal_planner_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty "
                             "(no example records)")
    parser.add_argument("--no-protect", action="store_true",
                        help="do not protect the sheets (for development)")
    parser.add_argument("--password", default=None,
                        help="sheet protection password")
    parser.add_argument("--author", default=AUTHOR,
                        help="document author / creator")
    parser.add_argument("--household", default=None,
                        help="pre-fill the household name on Settings")
    parser.add_argument("--cook", default=None,
                        help="pre-fill the main cook's name on Settings")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="planning year the planner opens on")
    parser.add_argument("--week-start", choices=["Monday", "Sunday"],
                        default=None, help="first day of the week")
    parser.add_argument("--servings", type=int, default=None,
                        help="default number of people a meal feeds")
    parser.add_argument("--budget", type=float, default=None,
                        help="weekly grocery budget")
    parser.add_argument("--units", choices=["Metric", "Imperial"],
                        default=None,
                        help="which measurement system to lead with")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the default row "
                             "counts are multiplied to fit this many "
                             "recipe rows)")
    parser.add_argument("--today", default=None,
                        help="treat this ISO date as 'today' when "
                             "generating the example data")
    parser.add_argument("--quiet", action="store_true",
                        help="only print the output path")
    return parser


def config_from_args(args: argparse.Namespace) -> Config:
    kwargs = {}
    if args.blank:
        kwargs["include_sample_data"] = False
    if args.no_protect:
        kwargs["protect_sheets"] = False
    if args.password:
        kwargs["password"] = args.password
    if args.author:
        kwargs["author"] = args.author
    if args.currency:
        kwargs["currency_symbol"] = args.currency
    if args.currency_code:
        kwargs["currency_code"] = args.currency_code
    if args.week_start:
        kwargs["week_start"] = args.week_start
    if args.year:
        kwargs["plan_year"] = args.year
    if args.servings:
        kwargs["default_servings"] = args.servings
    if args.units:
        kwargs["measurement_system"] = args.units
    if args.budget:
        kwargs["weekly_food_budget"] = args.budget
        kwargs["monthly_food_budget"] = round(args.budget * 52 / 12, 2)
        kwargs["annual_food_budget"] = round(args.budget * 52, 2)
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.household:
        cfg.household.name = args.household
    if args.cook:
        cfg.household.cook = args.cook
    if args.rows:
        cfg.scale_rows(args.rows)
    return cfg


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        cfg = config_from_args(args)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    path = build_workbook(args.output, cfg)
    size = os.path.getsize(path) / 1024.0

    if args.quiet:
        print(path)
        return 0

    print(f"{cfg.product_name}")
    print(f"  listing title : {PRODUCT_LISTING_TITLE}")
    print(f"  author        : {cfg.author}")
    print(f"  sheets        : {len(SHEET_ORDER)}")
    print(f"  planning year : {cfg.year}")
    print(f"  household     : {cfg.household.name} "
          f"({cfg.members} people)")
    print(f"  example data  : "
          f"{'no' if not cfg.include_sample_data else 'yes'}")
    print(f"  protection    : "
          f"{'off' if not cfg.protect_sheets else 'on, password ' + cfg.password}")
    print(f"  written to    : {os.path.abspath(path)}  ({size:,.0f} KB)")
    return 0


if __name__ == "__main__":        # pragma: no cover
    raise SystemExit(main())
````

### `meal_planner_spreadsheet/config.py`

*241 lines · 3 classes*

> Configuration for the Automated Meal Planner & Grocery List Spreadsheet.
>
> Everything a shop owner may want to re-brand lives here, so the workbook can
> be regenerated for a different household, currency, measurement system or
> country without touching any of the sheet code.
>
> Author of the generated workbooks: Novality Store.

| Defines | What it does |
|---------|--------------|
| `class Palette` |  |
| `class HouseholdDefaults` | Sample values written into the Household Profile (Settings) sheet |
| `class Config` | Everything the generator needs to know to write one workbook |

**Module constants:** `AUTHOR` ('Novality Store'), `PRODUCT_NAME` ('Automated Meal Planner & Grocery List Spreadsheet'), `PRODUCT_SUBTITLE` ('Weekly & Monthly Meal Planner, Recipe Organizer, Automat...), `PRODUCT_VERSION` ('1.0.0'), `PRODUCT_LISTING_TITLE` ('Automated Meal Planner Spreadsheet, Weekly and Monthly M...), `FORMULA_PASSWORD` ('premium'), `MEAL_SLOTS` (7 items), `MAIN_SLOTS` (3 items), `DEFAULT_ROWS` (29 entries).

````python
"""
Configuration for the Automated Meal Planner & Grocery List Spreadsheet.

Everything a shop owner may want to re-brand lives here, so the workbook can
be regenerated for a different household, currency, measurement system or
country without touching any of the sheet code.

Author of the generated workbooks: Novality Store.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import date


# ---------------------------------------------------------------------------
# Product / branding constants
# ---------------------------------------------------------------------------

AUTHOR = "Novality Store"
PRODUCT_NAME = "Automated Meal Planner & Grocery List Spreadsheet"
PRODUCT_SUBTITLE = (
    "Weekly & Monthly Meal Planner, Recipe Organizer, Automatic Grocery "
    "List, Pantry Inventory, Food Budget & Nutrition Tracker"
)
PRODUCT_VERSION = "1.0.0"

#: Title suggested for the shop listing.
PRODUCT_LISTING_TITLE = (
    "Automated Meal Planner Spreadsheet, Weekly and Monthly Meal Planner, "
    "Recipe Organizer, Grocery List, Pantry Inventory, Food Budget and "
    "Nutrition Tracker"
)

#: Password that protects every sheet so formula cells cannot be edited by
#: accident.  Input cells stay unlocked, formula cells are locked.
FORMULA_PASSWORD = "premium"

#: The meal slots the planner covers out of the box.  Edit them on the Lists
#: tab and every drop down, planner grid and nutrition total follows.
MEAL_SLOTS = ["Breakfast", "Morning snack", "Lunch", "Afternoon snack",
              "Dinner", "Evening snack", "Dessert"]

#: The three meal slots the dashboards and the printable day plan lead with.
MAIN_SLOTS = ["Breakfast", "Lunch", "Dinner"]


# ---------------------------------------------------------------------------
# Colours - a warm kitchen palette: herb green, terracotta, cream
# ---------------------------------------------------------------------------

class Palette:
    ink = "#1E2C22"           # darkest green - titles
    primary = "#2F5E43"       # headers
    primary_mid = "#4E8A63"   # section bands
    primary_soft = "#E2EFE6"  # light band / alternating rows
    accent = "#D9703C"        # terracotta accent
    accent_soft = "#FBE9DC"
    input_bg = "#FFFDF7"      # editable cell background (warm cream)
    input_border = "#E8DEC6"
    calc_bg = "#EFF4F0"       # locked formula cell background (soft sage)
    calc_border = "#D2E0D6"
    white = "#FFFFFF"
    grey_text = "#5A6560"
    line = "#C2D4C7"
    good_bg = "#D9F2E4"
    good_text = "#0E6B45"
    warn_bg = "#FFF1CF"
    warn_text = "#8A5A00"
    bad_bg = "#FBD9DD"
    bad_text = "#9C1C2E"
    info_bg = "#DEE8FA"
    info_text = "#2A4A8B"
    note_bg = "#F7FAF8"


# ---------------------------------------------------------------------------
# Row capacity of every register (how many pre-formatted / pre-formulated rows)
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Settings": 0,
    "Lists": 40,
    "Household": 12,
    "Food Preferences": 40,
    "Recipes": 200,
    "Recipe Ingredients": 1200,
    "Ingredients": 300,
    "Meal Plan": 500,
    "Grocery List": 300,
    "Pantry": 300,
    "Fridge & Freezer": 150,
    "Leftovers": 120,
    "Budget": 60,
    "Grocery Spending": 400,
    "Recipe Costs": 200,
    "Converter": 40,
    "Nutrition": 120,
    "Meal Prep": 120,
    "Staples": 100,
    "Stores": 40,
    "Store Prices": 300,
    "Rotation": 200,
    "Household Tasks": 120,
    "Food Safety": 60,
    "Events": 40,
    "Event Menu": 150,
    "Goals": 40,
    "Habits": 120,
    "Food Waste": 150,
}


@dataclass
class HouseholdDefaults:
    """Sample values written into the Household Profile (Settings) sheet."""

    name: str = "The Harper Household"
    cook: str = "Maya Harper"
    email: str = "maya@harperhome.com"
    members: int = 4
    town: str = "Portland, OR"
    store: str = "Greenfield Market"
    notes: str = "Two adults, two children, one nut allergy in the house"


# ---------------------------------------------------------------------------
# The configuration object every sheet module reads
# ---------------------------------------------------------------------------

@dataclass
class Config:
    """Everything the generator needs to know to write one workbook."""

    author: str = AUTHOR
    product_name: str = PRODUCT_NAME
    product_subtitle: str = PRODUCT_SUBTITLE
    listing_title: str = PRODUCT_LISTING_TITLE
    version: str = PRODUCT_VERSION
    password: str = FORMULA_PASSWORD
    protect_sheets: bool = True
    include_sample_data: bool = True

    # Formatting ------------------------------------------------------------
    currency_symbol: str = "$"
    currency_code: str = "USD"
    date_format: str = "dd-mmm-yyyy"
    time_format: str = "h:mm AM/PM"

    # Household / planning defaults pre-filled on the Settings tab ----------
    household: HouseholdDefaults = field(default_factory=HouseholdDefaults)
    plan_year: int = 0                # 0 = the year of `today`
    week_start: str = "Monday"
    measurement_system: str = "Metric"   # Metric | Imperial
    default_servings: int = 4
    members: int = 4

    # Budget ----------------------------------------------------------------
    weekly_food_budget: float = 180.0
    monthly_food_budget: float = 780.0
    annual_food_budget: float = 9360.0
    eating_out_budget: float = 120.0
    waste_cost_target: float = 15.0      # money a month written off as waste

    # Nutrition targets (per person, per day) -------------------------------
    calorie_target: int = 2000
    protein_target: int = 90
    carb_target: int = 230
    fat_target: int = 70
    fibre_target: int = 30
    sugar_limit: int = 50
    sodium_limit: int = 2300             # mg
    water_target: int = 8                # glasses a day

    # Reminder windows ------------------------------------------------------
    expiry_warning_days: int = 5         # "use it or lose it" window
    leftover_safe_days: int = 3          # fridge life of a cooked leftover
    freezer_review_days: int = 90        # how long before a freezer item is old
    low_stock_days: int = 7              # days of stock that counts as low
    prep_reminder_days: int = 2          # how early prep tasks are flagged
    rotation_gap_days: int = 14          # minimum days before repeating a meal
    shopping_day: str = "Saturday"
    prep_day: str = "Sunday"

    # Targets the habit tracker measures ------------------------------------
    home_cooked_target: int = 18         # home cooked meals a week
    new_recipe_target: int = 2           # new recipes a month
    takeaway_limit: int = 4              # takeaways a month
    packed_lunch_target: int = 5         # packed lunches a week
    waste_reduction_target: float = 0.05  # share of food bought that is binned

    # Behaviour -------------------------------------------------------------
    rounding: str = "Whole packs"
    top_rows: int = 5                    # size of the "top / best" tables
    rows: dict = field(default_factory=lambda: dict(DEFAULT_ROWS))
    today: date = field(default_factory=date.today)

    # ---------------------------------------------------------------- helpers
    def row_count(self, sheet: str) -> int:
        return int(self.rows.get(sheet, 100))

    @property
    def year(self) -> int:
        return self.plan_year or self.today.year

    @property
    def metric(self) -> bool:
        return self.measurement_system.lower().startswith("m")

    def scale_rows(self, recipe_rows: int) -> None:
        """Scale every register in place so Recipes holds N rows."""
        base = DEFAULT_ROWS["Recipes"]
        factor = max(0.02, float(recipe_rows) / float(base))
        self.rows = {k: (max(5, int(round(v * factor))) if v else v)
                     for k, v in self.rows.items()}

    def scaled(self, factor: float) -> "Config":
        """Return a copy with every register scaled."""
        rows = {k: max(5, int(round(v * factor))) if v else v
                for k, v in self.rows.items()}
        return replace(self, rows=rows)

    # Number format strings derived from the chosen currency ------------------
    @property
    def money_format(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0.00;[Red]-{s}#,##0.00'

    @property
    def money_format_0(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0;[Red]-{s}#,##0'

    @property
    def start_date(self) -> date:
        """The Monday (or Sunday) the example plan starts from."""
        day = self.today
        offset = day.weekday() if self.week_start == "Monday" \
            else (day.weekday() + 1) % 7
        return date.fromordinal(day.toordinal() - offset)
````

### `meal_planner_spreadsheet/layout.py`

*253 lines · 3 classes · 10 functions*

> Layout engine.
>
> The workbook is described declaratively: every register sheet (Recipes,
> Recipe Ingredients, Meal Plan, Pantry, ...) is a :class:`SheetSpec` of :class:`Col`
> objects.  Formulas
> are written as templates that never contain hard coded column letters, so
> columns can be reordered or inserted without breaking a single formula:
>
>     "=IF({title}{r}=\"\",\"\",SUMIFS([[Performance.reach]],[[Performance.content]],{id}{r}))"
>
>     {r}            -> the current Excel row number (1 based)
>     {key}          -> column letter of the column `key` on the *same* sheet
>     [[Sheet.key]]  -> absolute range of that column's data rows, e.g.
>                       'Recipe Ingredients'!$K$5:$K$1204
>     [[Sheet.key#r]]-> absolute single cell of that column on row r

| Defines | What it does |
|---------|--------------|
| `col_letter(index)` | 0 based column index -> Excel column letter |
| `quote_sheet(name)` | Quote a sheet name for use inside a formula when it needs it |
| `cell_ref(col, row, absolute)` |  |
| `class Col` | One column of a register sheet |
| `class SheetSpec` | A register sheet: title band, header row and N pre-built data rows |
| `class Registry` | Lookup table of every registered :class:`SheetSpec` |
| `max_if(value_range, *conditions)` | MAX of `value_range` where all conditions hold (no CSE required) |
| `min_if(value_range, *conditions, *, sentinel)` | MIN of `value_range` where all conditions hold (no CSE required) |
| `blank_if(test, result)` |  |
| `safe(expr, fallback)` |  |
| `lookup(value, table_range, index_range, fallback)` | INDEX/MATCH lookup (works left or right, unlike VLOOKUP) |
| `dynamic_name(sheet, letter, first, last, count_text)` | OFFSET based defined name that grows with the populated rows |
| `join_nonblank(parts)` | Build `a & b & c` style concatenation |

**Module constants:** `TITLE_ROW` (1), `SUBTITLE_ROW` (2), `NAV_ROW` (3), `HEADER_ROW` (4), `DATA_ROW` (5).

````python
"""
Layout engine.

The workbook is described declaratively: every register sheet (Recipes,
Recipe Ingredients, Meal Plan, Pantry, ...) is a :class:`SheetSpec` of :class:`Col`
objects.  Formulas
are written as templates that never contain hard coded column letters, so
columns can be reordered or inserted without breaking a single formula:

    "=IF({title}{r}=\\"\\",\\"\\",SUMIFS([[Performance.reach]],[[Performance.content]],{id}{r}))"

    {r}            -> the current Excel row number (1 based)
    {key}          -> column letter of the column `key` on the *same* sheet
    [[Sheet.key]]  -> absolute range of that column's data rows, e.g.
                      'Recipe Ingredients'!$K$5:$K$1204
    [[Sheet.key#r]]-> absolute single cell of that column on row r
"""

from __future__ import annotations

import re
import string
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence

# Row 1 title, row 2 strapline, row 3 navigation, row 4 headers, row 5+ data.
TITLE_ROW = 1
SUBTITLE_ROW = 2
NAV_ROW = 3
HEADER_ROW = 4
DATA_ROW = 5


# ---------------------------------------------------------------------------
# A1 helpers
# ---------------------------------------------------------------------------

def col_letter(index: int) -> str:
    """0 based column index -> Excel column letter."""
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = string.ascii_uppercase[rem] + letters
    return letters


def quote_sheet(name: str) -> str:
    """Quote a sheet name for use inside a formula when it needs it."""
    if re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def cell_ref(col: int, row: int, absolute: bool = True) -> str:
    dollar = "$" if absolute else ""
    return f"{dollar}{col_letter(col)}{dollar}{row}"


# ---------------------------------------------------------------------------
# Column / sheet specifications
# ---------------------------------------------------------------------------

@dataclass
class Col:
    """One column of a register sheet."""

    key: str
    header: str
    width: float = 14.0
    style: str = "text"              # key into the format book (see formats.py)
    formula: Optional[str] = None    # template -> locked, calculated column
    validate: Optional[Dict[str, Any]] = None
    note: Optional[str] = None       # help text shown when the cell is selected
    cf: List[Dict[str, Any]] = field(default_factory=list)
    hidden: bool = False
    group: str = ""                  # section label (documentation only)

    @property
    def is_formula(self) -> bool:
        return self.formula is not None


@dataclass
class SheetSpec:
    """A register sheet: title band, header row and N pre-built data rows."""

    name: str
    title: str
    subtitle: str = ""
    columns: List[Col] = field(default_factory=list)
    rows: int = 100
    header_row: int = HEADER_ROW        # 1 based Excel row of the header
    tab_color: Optional[str] = None
    freeze_cols: int = 2
    freeze_at: Optional[int] = None   # 1 based row to freeze below
    samples: List[Dict[str, Any]] = field(default_factory=list)
    help_text: str = ""
    extra_notes: List[str] = field(default_factory=list)

    # -------------------------------------------------------------- geometry
    @property
    def first_data_row(self) -> int:
        return self.header_row + 1

    @property
    def last_data_row(self) -> int:
        return self.header_row + self.rows

    def index(self, key: str) -> int:
        for i, col in enumerate(self.columns):
            if col.key == key:
                return i
        raise KeyError(f"{self.name}: unknown column '{key}'")

    def letter(self, key: str) -> str:
        return col_letter(self.index(key))

    def col(self, key: str) -> Col:
        return self.columns[self.index(key)]

    @property
    def last_col_index(self) -> int:
        return len(self.columns) - 1


class Registry:
    """Lookup table of every registered :class:`SheetSpec`."""

    def __init__(self) -> None:
        self._specs: Dict[str, SheetSpec] = {}

    def add(self, spec: SheetSpec) -> SheetSpec:
        self._specs[spec.name] = spec
        return spec

    def __contains__(self, name: str) -> bool:
        return name in self._specs

    def __getitem__(self, name: str) -> SheetSpec:
        return self._specs[name]

    def names(self) -> Sequence[str]:
        return list(self._specs)

    # ------------------------------------------------------------ references
    def rng(self, sheet: str, key: str, first: Optional[int] = None,
            last: Optional[int] = None) -> str:
        spec = self._specs[sheet]
        letter = spec.letter(key)
        first = first or spec.first_data_row
        last = last or spec.last_data_row
        return f"{quote_sheet(sheet)}!${letter}${first}:${letter}${last}"

    def cell(self, sheet: str, key: str, row: int) -> str:
        spec = self._specs[sheet]
        return f"{quote_sheet(sheet)}!${spec.letter(key)}${row}"

    def letter(self, sheet: str, key: str) -> str:
        return self._specs[sheet].letter(key)

    # -------------------------------------------------------------- renderer
    _REF = re.compile(r"\[\[([^\].#]+)\.([A-Za-z0-9_]+)(?:#(\d*))?\]\]")

    def render(self, template: str, row: Optional[int] = None,
               spec: Optional[SheetSpec] = None) -> str:
        """Expand [[Sheet.key]] and {key}/{r} tokens into a real formula."""

        def _ref(match: "re.Match[str]") -> str:
            sheet, key, single = match.group(1), match.group(2), match.group(3)
            if single:
                return self.cell(sheet, key, int(single))
            if single == "":
                # [[Sheet.key#]] -> same row as the formula being written
                if row is None:
                    raise ValueError(
                        f"[[{sheet}.{key}#]] needs a row context")
                return self.cell(sheet, key, row)
            return self.rng(sheet, key)

        text = self._REF.sub(_ref, template)

        if spec is not None:
            mapping: Dict[str, Any] = {c.key: spec.letter(c.key)
                                       for c in spec.columns}
            # `{key_first}` / `{key_last}` give the absolute top and bottom
            # of a column on the same sheet, so an in-sheet COUNTIFS reads
            # `{key_first}:{key_last}` instead of a hand built $D$5:$D$104.
            for col in spec.columns:
                letter = spec.letter(col.key)
                mapping[f"{col.key}_first"] = (
                    f"${letter}${spec.first_data_row}")
                mapping[f"{col.key}_last"] = (
                    f"${letter}${spec.last_data_row}")
            if row is not None:
                mapping["r"] = row
                mapping["r_prev"] = row - 1
            mapping["first"] = spec.first_data_row
            mapping["last"] = spec.last_data_row
            text = text.format(**mapping)
        elif row is not None:
            text = text.format(r=row)
        return text


# ---------------------------------------------------------------------------
# Reusable formula fragments
# ---------------------------------------------------------------------------

def max_if(value_range: str, *conditions: str) -> str:
    """MAX of `value_range` where all conditions hold (no CSE required)."""
    cond = "*".join(f"({c})" for c in conditions)
    return f"SUMPRODUCT(MAX({cond}*{value_range}))"


def min_if(value_range: str, *conditions: str, sentinel: int = 2958465) -> str:
    """MIN of `value_range` where all conditions hold (no CSE required).

    Non matching rows are swapped for a sentinel (31-Dec-9999 as a serial) so
    that MIN ignores them.  Returns the sentinel when nothing matches, so the
    caller should wrap the result in an IF test.
    """
    cond = "*".join(f"({c})" for c in conditions)
    return (f"SUMPRODUCT(MIN({cond}*{value_range}"
            f"+(1-{cond})*{sentinel}))")


def blank_if(test: str, result: str) -> str:
    return f'IF({test},"",{result})'


def safe(expr: str, fallback: str = '""') -> str:
    return f"IFERROR({expr},{fallback})"


def lookup(value: str, table_range: str, index_range: str,
           fallback: str = '""') -> str:
    """INDEX/MATCH lookup (works left or right, unlike VLOOKUP)."""
    return f"IFERROR(INDEX({index_range},MATCH({value},{table_range},0)),{fallback})"


def dynamic_name(sheet: str, letter: str, first: int, last: int,
                 count_text: bool = True) -> str:
    """OFFSET based defined name that grows with the populated rows."""
    rng = f"{quote_sheet(sheet)}!${letter}${first}:${letter}${last}"
    counter = f'COUNTIF({rng},"?*")' if count_text else f"COUNTA({rng})"
    return (f"=OFFSET({quote_sheet(sheet)}!${letter}${first},0,0,"
            f"MAX(1,{counter}),1)")


def join_nonblank(parts: Sequence[str]) -> str:
    """Build `a & b & c` style concatenation."""
    return "&".join(parts)
````

### `meal_planner_spreadsheet/formats.py`

*200 lines · 1 class*

> Cell format factory.
>
> Two things matter here:
>
> * **Input cells** are unlocked (``locked: False``) and painted cream, so the
>   user can type in them while the sheet is protected.
> * **Formula cells** are locked (``locked: True``) and painted cool grey.  With
>   the sheet protection password in place ("premium" by default) they cannot be
>   overwritten by accident.

| Defines | What it does |
|---------|--------------|
| `class FormatBook` | Creates (and caches) every :class:`xlsxwriter.format.Format` used |

**Module constants:** `P` (Palette), `INPUT` ('input'), `CALC` ('calc'), `PLAIN` ('plain').

````python
"""
Cell format factory.

Two things matter here:

* **Input cells** are unlocked (``locked: False``) and painted cream, so the
  user can type in them while the sheet is protected.
* **Formula cells** are locked (``locked: True``) and painted cool grey.  With
  the sheet protection password in place ("premium" by default) they cannot be
  overwritten by accident.
"""

from __future__ import annotations

from typing import Any, Dict

from .config import Config, Palette

P = Palette

INPUT = "input"
CALC = "calc"
PLAIN = "plain"


class FormatBook:
    """Creates (and caches) every :class:`xlsxwriter.format.Format` used."""

    def __init__(self, book, config: Config) -> None:
        self.book = book
        self.cfg = config
        self._cache: Dict[Any, Any] = {}

        money = config.money_format
        money0 = config.money_format_0
        self._styles: Dict[str, Dict[str, Any]] = {
            "text": {"align": "left"},
            "id": {"align": "left", "bold": True, "font_color": P.primary},
            "wrap": {"align": "left", "valign": "top", "text_wrap": True},
            "date": {"num_format": config.date_format, "align": "center"},
            "time": {"num_format": config.time_format, "align": "center"},
            "money": {"num_format": money, "align": "right"},
            "money0": {"num_format": money0, "align": "right"},
            "number": {"num_format": "#,##0.00", "align": "right"},
            "number1": {"num_format": "#,##0.0", "align": "right"},
            "int": {"num_format": "#,##0", "align": "right"},
            "percent": {"num_format": "0.0%", "align": "center"},
            "percent0": {"num_format": "0%", "align": "center"},
            "rating": {"num_format": "0", "align": "center"},
            "flag": {"align": "center"},
            "status": {"align": "center", "bold": True},
            "link": {"align": "left", "font_color": P.info_text,
                     "underline": 1},
        }

        self._modes: Dict[str, Dict[str, Any]] = {
            INPUT: {"locked": False, "bg_color": P.input_bg,
                    "border": 1, "border_color": P.input_border},
            CALC: {"locked": True, "bg_color": P.calc_bg,
                   "border": 1, "border_color": P.calc_border,
                   "font_color": P.ink},
            PLAIN: {"locked": True},
        }

    # ------------------------------------------------------------------ core
    def _make(self, props: Dict[str, Any]):
        key = tuple(sorted((k, str(v)) for k, v in props.items()))
        if key not in self._cache:
            self._cache[key] = self.book.add_format(props)
        return self._cache[key]

    def cell(self, style: str = "text", mode: str = INPUT, **extra):
        """Data cell format: `style` = value type, `mode` = input/calc/plain."""
        props: Dict[str, Any] = {"font_name": "Calibri", "font_size": 10,
                                 "valign": "vcenter"}
        props.update(self._modes[mode])
        props.update(self._styles.get(style, {}))
        props.update(extra)
        return self._make(props)

    def input(self, style: str = "text", **extra):
        return self.cell(style, INPUT, **extra)

    def calc(self, style: str = "text", **extra):
        return self.cell(style, CALC, **extra)

    # ---------------------------------------------------------------- chrome
    def chrome(self, name: str, **extra):
        base: Dict[str, Any] = {"font_name": "Calibri", "font_size": 10,
                                "locked": True, "valign": "vcenter"}
        styles: Dict[str, Dict[str, Any]] = {
            "title": {"font_size": 18, "bold": True, "font_color": P.white,
                      "bg_color": P.primary, "align": "left", "indent": 1},
            "title_right": {"font_size": 10, "bold": True,
                            "font_color": P.white, "bg_color": P.primary,
                            "align": "right", "indent": 1},
            "subtitle": {"font_size": 10, "italic": True,
                         "font_color": P.grey_text, "bg_color": P.note_bg,
                         "align": "left", "indent": 1},
            "nav": {"font_size": 9, "bold": True, "font_color": P.info_text,
                    "bg_color": P.note_bg, "underline": 1, "align": "left",
                    "indent": 1},
            "nav_plain": {"font_size": 9, "font_color": P.grey_text,
                          "bg_color": P.note_bg, "align": "left"},
            "section": {"font_size": 11, "bold": True, "font_color": P.white,
                        "bg_color": P.primary_mid, "align": "left",
                        "indent": 1},
            "subsection": {"font_size": 10, "bold": True,
                           "font_color": P.primary, "bg_color": P.primary_soft,
                           "align": "left", "indent": 1},
            "header": {"font_size": 9, "bold": True, "font_color": P.white,
                       "bg_color": P.primary, "align": "center",
                       "valign": "vcenter", "text_wrap": True, "border": 1,
                       "border_color": P.primary},
            "header_left": {"font_size": 9, "bold": True,
                            "font_color": P.white, "bg_color": P.primary,
                            "align": "left", "valign": "vcenter",
                            "text_wrap": True, "border": 1,
                            "border_color": P.primary, "indent": 1},
            "label": {"font_size": 10, "align": "left", "indent": 1,
                      "font_color": P.ink},
            "label_bold": {"font_size": 10, "bold": True, "align": "left",
                           "indent": 1, "font_color": P.ink},
            "label_right": {"font_size": 10, "align": "right",
                            "font_color": P.ink},
            "note": {"font_size": 9, "italic": True,
                     "font_color": P.grey_text, "align": "left",
                     "text_wrap": True, "valign": "top"},
            "note_box": {"font_size": 9, "font_color": P.ink,
                         "bg_color": P.accent_soft, "align": "left",
                         "valign": "top", "text_wrap": True, "border": 1,
                         "border_color": P.accent, "indent": 1},
            "kpi_label": {"font_size": 9, "bold": True,
                          "font_color": P.white, "bg_color": P.primary_mid,
                          "align": "left", "indent": 1},
            "kpi_value": {"font_size": 16, "bold": True,
                          "font_color": P.ink, "bg_color": P.white,
                          "align": "center", "valign": "vcenter",
                          "border": 1, "border_color": P.line},
            "kpi_note": {"font_size": 8, "italic": True,
                         "font_color": P.grey_text, "bg_color": P.white,
                         "align": "center", "border": 1,
                         "border_color": P.line},
            "total_label": {"font_size": 10, "bold": True,
                            "font_color": P.white, "bg_color": P.primary,
                            "align": "left", "indent": 1},
            "total_value": {"font_size": 10, "bold": True,
                            "font_color": P.white, "bg_color": P.primary,
                            "align": "right"},
            "form_label": {"font_size": 10, "bold": True, "align": "left",
                           "font_color": P.primary, "bottom": 1,
                           "bottom_color": P.line, "indent": 1},
            "form_line": {"font_size": 10, "align": "left", "bottom": 1,
                          "bottom_color": P.ink, "locked": False},
            "form_box": {"font_size": 10, "align": "left", "valign": "top",
                         "text_wrap": True, "border": 1,
                         "border_color": P.ink, "locked": False},
            "print_title": {"font_size": 16, "bold": True,
                            "font_color": P.primary, "align": "left"},
            "print_sub": {"font_size": 10, "italic": True,
                          "font_color": P.grey_text, "align": "left"},
            "legend_input": {"font_size": 9, "bg_color": P.input_bg,
                             "border": 1, "border_color": P.input_border,
                             "align": "center"},
            "legend_calc": {"font_size": 9, "bg_color": P.calc_bg,
                            "border": 1, "border_color": P.calc_border,
                            "align": "center"},
            "band": {"bg_color": P.primary_soft},
            "rule": {"bottom": 2, "bottom_color": P.primary},
            "legend_label": {"font_size": 9, "bold": True,
                             "font_color": P.grey_text, "align": "right"},
            "day_number": {"font_size": 11, "bold": True,
                           "font_color": P.primary,
                           "bg_color": P.primary_soft, "align": "left",
                           "indent": 1, "num_format": "d ddd",
                           "border": 1, "border_color": P.line},
            "day_item": {"font_size": 8, "font_color": P.ink,
                         "align": "left", "indent": 1, "border": 1,
                         "border_color": P.line, "bg_color": P.white},
        }
        props = dict(base)
        props.update(styles.get(name, {}))
        props.update(extra)
        return self._make(props)

    # -------------------------------------------------- conditional formats
    def cf(self, kind: str):
        """Formats used by conditional formatting rules."""
        palettes = {
            "good": {"bg_color": P.good_bg, "font_color": P.good_text},
            "warn": {"bg_color": P.warn_bg, "font_color": P.warn_text},
            "bad": {"bg_color": P.bad_bg, "font_color": P.bad_text},
            "info": {"bg_color": P.info_bg, "font_color": P.info_text},
            "muted": {"font_color": P.grey_text, "italic": True},
            "bold_bad": {"bg_color": P.bad_bg, "font_color": P.bad_text,
                         "bold": True},
            "bold_good": {"bg_color": P.good_bg, "font_color": P.good_text,
                          "bold": True},
        }
        return self._make(palettes[kind])
````

### `meal_planner_spreadsheet/writer.py`

*444 lines · 1 class · 7 functions*

> Sheet writing helpers shared by every tab of the workbook.
>
> :class:`SheetWriter` wraps an XlsxWriter worksheet and adds the house style:
> title band, navigation strip, section headers, KPI tiles, label/value settings
> rows and - most importantly - :meth:`SheetWriter.grid`, which turns a
> :class:`~meal_planner_spreadsheet.layout.SheetSpec` into a finished register
> with headers, drop downs, conditional formats, sample rows, frozen panes and
> locked formula cells.

| Defines | What it does |
|---------|--------------|
| `class SheetWriter` | Convenience wrapper around a worksheet |
| `yes_no(validate_note)` |  |
| `list_from(name)` | Drop down fed by a defined name (see lists.py) |
| `date_check()` |  |
| `number_check(minimum, maximum, decimal)` |  |
| `cf_text(text, fmt, criteria)` |  |
| `cf_formula(formula, fmt)` |  |
| `cf_bar(color)` |  |

**Module constants:** `P` (Palette), `NAV_LINKS` (9 items), `PROTECT_OPTIONS` (15 entries).

````python
"""
Sheet writing helpers shared by every tab of the workbook.

:class:`SheetWriter` wraps an XlsxWriter worksheet and adds the house style:
title band, navigation strip, section headers, KPI tiles, label/value settings
rows and - most importantly - :meth:`SheetWriter.grid`, which turns a
:class:`~meal_planner_spreadsheet.layout.SheetSpec` into a finished register
with headers, drop downs, conditional formats, sample rows, frozen panes and
locked formula cells.
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .config import Config, Palette
from .formats import CALC, FormatBook, INPUT, PLAIN
from .layout import (NAV_ROW, SUBTITLE_ROW, TITLE_ROW, Registry,
                     SheetSpec, col_letter, quote_sheet)

P = Palette

#: Sheets linked from the navigation strip at the top of every tab.
NAV_LINKS: List[Tuple[str, str]] = [
    ("Start Here", "Start Here"),
    ("Dashboard", "Dashboard"),
    ("Recipes", "Recipes"),
    ("Week", "Weekly Planner"),
    ("Month", "Monthly Planner"),
    ("Grocery", "Grocery List"),
    ("Pantry", "Pantry"),
    ("Budget", "Budget"),
    ("Help", "Help"),
]

#: Worksheet protection: formulas are locked, everything else stays usable.
PROTECT_OPTIONS: Dict[str, bool] = {
    "objects": False,
    "scenarios": False,
    "format_cells": True,
    "format_columns": True,
    "format_rows": True,
    "insert_columns": False,
    "insert_rows": True,
    "insert_hyperlinks": True,
    "delete_columns": False,
    "delete_rows": True,
    "select_locked_cells": True,
    "sort": False,          # stops rows being re-sorted out of sync with IDs
    "autofilter": True,     # filtering stays available
    "pivot_tables": False,
    "select_unlocked_cells": True,
}


class SheetWriter:
    """Convenience wrapper around a worksheet."""

    def __init__(self, book, fmts: FormatBook, cfg: Config, reg: Registry,
                 name: str, tab_color: Optional[str] = None,
                 landscape: bool = True) -> None:
        self.book = book
        self.f = fmts
        self.cfg = cfg
        self.reg = reg
        self.name = name
        self.ws = book.add_worksheet(name)
        self._protected = False
        if tab_color:
            self.ws.set_tab_color(tab_color)
        if landscape:
            self.ws.set_landscape()
        self.ws.set_paper(9)  # A4
        self.ws.set_margins(0.4, 0.4, 0.6, 0.5)
        self.ws.set_header(
            f"&L&\"Calibri,Bold\"&12{''}"
            f"&C&\"Calibri,Bold\"&11{name}"
            f"&R&\"Calibri,Italic\"&9{cfg.product_name}")
        self.ws.set_footer(
            f"&L&\"Calibri,Italic\"&8{cfg.author} - v{cfg.version}"
            "&C&\"Calibri,Italic\"&8Page &P of &N"
            "&R&\"Calibri,Italic\"&8&D")

    # ------------------------------------------------------------- plumbing
    def write(self, row: int, col: int, value: Any, fmt=None):
        """Type aware write (0 based row/col, like XlsxWriter)."""
        if value is None:
            if fmt is not None:
                self.ws.write_blank(row, col, None, fmt)
            return
        if isinstance(value, str) and value.startswith("="):
            self.ws.write_formula(row, col, value, fmt)
        elif isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
            self.ws.write_datetime(row, col, value, fmt)
        elif isinstance(value, bool):
            self.ws.write_string(row, col, "Yes" if value else "No", fmt)
        elif isinstance(value, (int, float)):
            self.ws.write_number(row, col, value, fmt)
        else:
            self.ws.write_string(row, col, str(value), fmt)

    def merge(self, row: int, first_col: int, last_col: int, value: Any,
              fmt=None, row2: Optional[int] = None):
        row2 = row if row2 is None else row2
        if first_col == last_col and row == row2:
            self.write(row, first_col, value, fmt)
        else:
            text = "" if value is None else value
            if isinstance(text, (_dt.datetime, _dt.date, _dt.time)):
                self.ws.merge_range(row, first_col, row2, last_col, "", fmt)
                self.ws.write_datetime(row, first_col, text, fmt)
            else:
                self.ws.merge_range(row, first_col, row2, last_col, text, fmt)

    def set_col(self, first: int, last: int, width: float, fmt=None,
                hidden: bool = False):
        options = {"hidden": 1} if hidden else {}
        self.ws.set_column(first, last, width, fmt, options)

    def row_height(self, row: int, height: float):
        self.ws.set_row(row, height)

    # ----------------------------------------------------------- page chrome
    def banner(self, title: str, subtitle: str = "", last_col: int = 9,
               nav: bool = True, badge: str = "",
               tail_text: str = "") -> None:
        """Title row + strapline + navigation strip (Excel rows 1-3).

        `tail_text` is written in the empty space to the right of the
        navigation links - used for the register help line on sheets whose
        header row leaves no spare row for it.
        """
        r = TITLE_ROW - 1
        self.row_height(r, 34)
        split = max(last_col - 3, 1)
        self.merge(r, 0, split, title, self.f.chrome("title"))
        self.merge(r, split + 1, last_col,
                   badge or f"{self.cfg.author}  |  v{self.cfg.version}",
                   self.f.chrome("title_right"))

        r = SUBTITLE_ROW - 1
        self.row_height(r, 18)
        self.merge(r, 0, last_col, subtitle, self.f.chrome("subtitle"))

        if nav:
            self.nav_strip(NAV_ROW - 1, last_col, tail_text=tail_text)

    def nav_strip(self, row: int, last_col: int,
                  links: Optional[Sequence[Tuple[str, str]]] = None,
                  tail_text: str = "") -> int:
        """Row of internal links. Returns the first column left over."""
        links = list(links or NAV_LINKS)
        self.row_height(row, 16)
        col = 0
        fmt = self.f.chrome("nav")
        for label, sheet in links:
            if col + 1 > last_col:
                break
            self.ws.merge_range(row, col, row, min(col + 1, last_col), "", fmt)
            self.ws.write_url(row, col, f"internal:{quote_sheet(sheet)}!A1",
                              fmt, f"> {label}")
            col += 2
        if col <= last_col:
            room = last_col - col
            if tail_text and room >= 4:
                self.merge(row, col, last_col, tail_text,
                           self.f.chrome("note", bg_color=P.note_bg))
            else:
                self.merge(row, col, last_col, "",
                           self.f.chrome("nav_plain"))
        return col

    def section(self, row: int, text: str, last_col: int, first_col: int = 0):
        self.row_height(row, 20)
        self.merge(row, first_col, last_col, text, self.f.chrome("section"))
        return row + 1

    def subsection(self, row: int, text: str, last_col: int,
                   first_col: int = 0):
        self.row_height(row, 17)
        self.merge(row, first_col, last_col, text,
                   self.f.chrome("subsection"))
        return row + 1

    def note(self, row: int, text: str, last_col: int, first_col: int = 0,
             height: float = 28, boxed: bool = True):
        self.row_height(row, height)
        self.merge(row, first_col, last_col, text,
                   self.f.chrome("note_box" if boxed else "note"))
        return row + 1

    def bullets(self, row: int, lines: Sequence[str], last_col: int,
                first_col: int = 0, height: float = 15):
        for line in lines:
            self.row_height(row, height)
            self.merge(row, first_col, last_col, line, self.f.chrome("label"))
            row += 1
        return row

    # -------------------------------------------------------- settings rows
    def field(self, row: int, label: str, value: Any = None, *,
              label_col: int = 1, value_col: int = 3,
              value_end: Optional[int] = None, style: str = "text",
              mode: str = INPUT, name: Optional[str] = None,
              validate: Optional[Dict[str, Any]] = None,
              help_text: str = "", help_col: Optional[int] = None,
              formula: Optional[str] = None):
        """A `Label | value | help` settings row, optionally named."""
        self.merge(row, label_col, value_col - 1, label,
                   self.f.chrome("label_bold"))
        fmt = self.f.cell(style, mode)
        end = value_end if value_end is not None else value_col
        if formula is not None:
            rendered = self.reg.render(formula)
            if end > value_col:
                self.ws.merge_range(row, value_col, row, end, "", fmt)
            self.ws.write_formula(row, value_col, rendered, fmt)
        else:
            if end > value_col:
                self.merge(row, value_col, end, None, fmt)
                self.write(row, value_col, value, fmt)
            else:
                self.write(row, value_col, value, fmt)
        if validate:
            self.ws.data_validation(row, value_col, row, value_col, validate)
        if help_text:
            hc = help_col if help_col is not None else end + 1
            self.merge(row, hc, hc + 3, help_text, self.f.chrome("note"))
        if name:
            self.book.define_name(
                name, f"={quote_sheet(self.name)}!"
                      f"${col_letter(value_col)}${row + 1}")
        return row + 1

    # ------------------------------------------------------------ KPI tiles
    def kpi(self, row: int, col: int, label: str, formula: str,
            style: str = "int", span: int = 2, note: str = ""):
        """Three row KPI tile: caption, big number, small note."""
        self.merge(row, col, col + span - 1, label, self.f.chrome("kpi_label"))
        self.row_height(row + 1, 26)
        fmt = self.f.chrome("kpi_value",
                            num_format=self._num_format(style))
        self.ws.merge_range(row + 1, col, row + 1, col + span - 1, "", fmt)
        self.ws.write_formula(row + 1, col, self.reg.render(formula), fmt)
        self.merge(row + 2, col, col + span - 1, note,
                   self.f.chrome("kpi_note"))
        return row + 3

    def num_format(self, style: str) -> str:
        """Number format string for a style key (public helper)."""
        return self._num_format(style)

    def _num_format(self, style: str) -> str:
        return {
            "money": self.cfg.money_format,
            "money0": self.cfg.money_format_0,
            "int": "#,##0",
            "number": "#,##0.00",
            "number1": "#,##0.0",
            "percent": "0.0%",
            "percent0": "0%",
            "date": self.cfg.date_format,
            "text": "General",
        }.get(style, "General")

    # ------------------------------------------------------- generic tables
    def table_head(self, row: int, col: int, headers: Sequence[str],
                   widths: Optional[Sequence[float]] = None,
                   align_left_first: bool = True):
        for i, head in enumerate(headers):
            fmt = self.f.chrome(
                "header_left" if (i == 0 and align_left_first) else "header")
            self.write(row, col + i, head, fmt)
            if widths:
                self.set_col(col + i, col + i, widths[i])
        self.row_height(row, 26)
        return row + 1

    def calc_row(self, row: int, col: int, values: Sequence[Any],
                 styles: Sequence[str], mode: str = CALC,
                 first_style: str = "text"):
        for i, value in enumerate(values):
            style = styles[i] if i < len(styles) else "text"
            fmt = self.f.cell(style, PLAIN if i == 0 and mode == PLAIN
                              else mode)
            if isinstance(value, str) and value.startswith("="):
                self.ws.write_formula(row, col + i, self.reg.render(value),
                                      fmt)
            else:
                self.write(row, col + i, value, fmt)
        return row + 1

    # ----------------------------------------------------------- the grid --
    def grid(self, spec: SheetSpec, extra_help: str = "") -> None:
        """Render a whole register sheet from its specification."""
        ws = self.ws
        last_col = spec.last_col_index
        header_row0 = spec.header_row - 1
        first = spec.first_data_row
        last = spec.last_data_row

        help_text = extra_help or spec.help_text
        spare_row = spec.header_row - 1 > NAV_ROW   # room above the headers?
        self.banner(spec.title, spec.subtitle, last_col,
                    tail_text="" if spare_row else help_text)

        if help_text and spare_row:
            self.note(header_row0 - 1, help_text, last_col, height=26)

        # Column widths, default formats (so blank cells keep the right lock
        # state) and the header row itself.
        for idx, col in enumerate(spec.columns):
            mode = CALC if col.is_formula else INPUT
            self.set_col(idx, idx, col.width, self.f.cell(col.style, mode),
                         hidden=col.hidden)
            self.write(header_row0, idx, col.header,
                       self.f.chrome("header_left" if idx <= 1 else "header"))
        self.row_height(header_row0, 30)

        # Data rows: formulas everywhere, sample values in the input columns.
        samples = spec.samples if self.cfg.include_sample_data else []
        for offset in range(spec.rows):
            excel_row = first + offset
            row0 = excel_row - 1
            sample = samples[offset] if offset < len(samples) else {}
            for idx, col in enumerate(spec.columns):
                if col.is_formula:
                    fmt = self.f.cell(col.style, CALC)
                    ws.write_formula(
                        row0, idx,
                        self.reg.render(col.formula, excel_row, spec), fmt)
                elif col.key in sample and sample[col.key] is not None:
                    self.write(row0, idx, sample[col.key],
                               self.f.cell(col.style, INPUT))

        # Drop downs, help bubbles and conditional formats.
        for idx, col in enumerate(spec.columns):
            if col.validate:
                options = dict(col.validate)
                if "source" in options and isinstance(options["source"], str):
                    options["source"] = self.reg.render(options["source"])
                if col.note:
                    options.setdefault("input_title", col.header[:32])
                    options.setdefault("input_message", col.note[:255])
                ws.data_validation(first - 1, idx, last - 1, idx, options)
            if col.note:
                # Header comment: works in every Excel version and survives
                # the trip through Google Sheets.
                ws.write_comment(header_row0, idx, col.note, {
                    "author": self.cfg.author,
                    "width": 220,
                    "height": 100,
                    "x_scale": 1.1,
                })
            for rule in col.cf:
                options = dict(rule)
                fmt_key = options.pop("fmt", None)
                if fmt_key:
                    options["format"] = self.f.cf(fmt_key)
                for field_name in ("criteria", "value", "minimum", "maximum"):
                    val = options.get(field_name)
                    if isinstance(val, str) and ("[[" in val or "{" in val):
                        options[field_name] = self.reg.render(
                            val, first, spec)
                ws.conditional_format(first - 1, idx, last - 1, idx, options)

        ws.autofilter(header_row0, 0, last - 1, last_col)
        freeze_row = (spec.freeze_at if spec.freeze_at is not None
                      else spec.header_row)
        ws.freeze_panes(freeze_row, spec.freeze_cols)
        ws.print_area(0, 0, last - 1, last_col)
        ws.repeat_rows(0, header_row0 if spec.header_row <= 6 else 2)
        ws.fit_to_pages(1, 0)

    # ------------------------------------------------------------- finishing
    def finish(self, freeze_row: Optional[int] = None,
               freeze_col: int = 0, print_area: Optional[str] = None,
               fit_width: bool = True, gridlines: bool = True):
        if freeze_row is not None:
            self.ws.freeze_panes(freeze_row, freeze_col)
        if print_area:
            self.ws.print_area(print_area)
        if fit_width:
            self.ws.fit_to_pages(1, 0)
        if not gridlines:
            self.ws.hide_gridlines(2)
        self.protect()

    def protect(self):
        if self.cfg.protect_sheets and not self._protected:
            self.ws.protect(self.cfg.password, dict(PROTECT_OPTIONS))
            self._protected = True


# ---------------------------------------------------------------------------
# Small shared helpers
# ---------------------------------------------------------------------------

def yes_no(validate_note: str = "") -> Dict[str, Any]:
    opts: Dict[str, Any] = {"validate": "list", "source": ["Yes", "No"]}
    if validate_note:
        opts["input_message"] = validate_note
    return opts


def list_from(name: str) -> Dict[str, Any]:
    """Drop down fed by a defined name (see lists.py)."""
    return {"validate": "list", "source": f"={name}"}


def date_check() -> Dict[str, Any]:
    return {
        "validate": "date",
        "criteria": "between",
        "minimum": _dt.date(2000, 1, 1),
        "maximum": _dt.date(2100, 12, 31),
        "error_title": "Check the date",
        "error_message": "Enter a real date between 2000 and 2100.",
    }


def number_check(minimum: float = 0, maximum: float = 1000000,
                 decimal: bool = True) -> Dict[str, Any]:
    return {
        "validate": "decimal" if decimal else "integer",
        "criteria": "between",
        "minimum": minimum,
        "maximum": maximum,
        "error_title": "Check the value",
        "error_message": f"Enter a number between {minimum} and {maximum}.",
    }


def cf_text(text: str, fmt: str, criteria: str = "containing") -> Dict[str, Any]:
    return {"type": "text", "criteria": criteria, "value": text, "fmt": fmt}


def cf_formula(formula: str, fmt: str) -> Dict[str, Any]:
    return {"type": "formula", "criteria": formula, "fmt": fmt}


def cf_bar(color: str = "#2E7D8F") -> Dict[str, Any]:
    return {"type": "data_bar", "bar_color": color}
````

### `meal_planner_spreadsheet/names.py`

*50 lines · 1 function*

> Workbook level defined names that point at the register sheets.
>
> The drop down lists on the `Lists` tab define their own names. These are
> the ones that have to follow live data - recipe IDs, ingredient names,
> the people in the house - so they are built from the registry once every
> register exists. Each one is an ``OFFSET`` range that grows and shrinks
> with the number of filled rows, which is why a new recipe appears in the
> meal planner drop-down the moment you type its name.

| Defines | What it does |
|---------|--------------|
| `define_cross_sheet_names(book, reg)` | Create the live lookup names |

**Module constants:** `CROSS_SHEET_NAMES` (16 entries).

````python
"""Workbook level defined names that point at the register sheets.

The drop down lists on the `Lists` tab define their own names. These are
the ones that have to follow live data - recipe IDs, ingredient names,
the people in the house - so they are built from the registry once every
register exists. Each one is an ``OFFSET`` range that grows and shrinks
with the number of filled rows, which is why a new recipe appears in the
meal planner drop-down the moment you type its name.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # the two that drive the automation
    "RecipeIDs": ("Recipes", "id"),
    "RecipeNames": ("Recipes", "name"),
    "IngredientNames": ("Ingredients", "name"),
    "IngredientIDs": ("Ingredients", "id"),
    # supporting lists
    "MemberNames": ("Household", "name"),
    "StoreNames": ("Stores", "name"),
    "EventNames": ("Events", "name"),
    "PantryItems": ("Pantry", "ingredient"),
    "StapleNames": ("Staples", "ingredient"),
    "LeftoverItems": ("Leftovers", "item"),
    "GoalNames": ("Goals", "goal"),
    "PrepJobs": ("Meal Prep", "job"),
    # columns the dashboard and the reports scan
    "PlanDates": ("Meal Plan", "date"),
    "PlanKeys": ("Meal Plan", "key"),
    "PlanRecipes": ("Meal Plan", "recipe"),
    "SpendMonths": ("Grocery Spending", "month"),
}


def define_cross_sheet_names(book, reg: Registry) -> None:
    """Create the live lookup names. Call after every register is built."""
    for name, (sheet, key) in CROSS_SHEET_NAMES.items():
        if sheet not in reg:
            continue
        spec = reg[sheet]
        book.define_name(
            name,
            dynamic_name(sheet, spec.letter(key), spec.first_data_row,
                         spec.last_data_row))
````

### `meal_planner_spreadsheet/build.py`

*161 lines · 3 functions*

> Workbook builder: puts every tab together in the right order.
>
> Typical use::
>
>     from meal_planner_spreadsheet import build_workbook
>     build_workbook("Meal Planner & Grocery List.xlsx")

| Defines | What it does |
|---------|--------------|
| `build_register(book, fmts, cfg, reg, module)` | Create one register sheet from its :class:`SheetSpec` |
| `build_workbook(path, cfg)` | Write the whole product to *path* and return the path |
| `_set_properties(book, cfg)` |  |

**Module constants:** `REGISTER_MODULES` (16 items), `SHEET_ORDER` (36 items), `PAGE_BUILDERS` (20 entries), `REGISTER_BY_NAME` ({module.NAME: module for module in REGISTER_MODULES}).

````python
"""Workbook builder: puts every tab together in the right order.

Typical use::

    from meal_planner_spreadsheet import build_workbook
    build_workbook("Meal Planner & Grocery List.xlsx")
"""

from __future__ import annotations

import datetime as _dt
import inspect
import os
from typing import Callable, Dict, List, Optional

from xlsxwriter.workbook import Workbook

from .config import Config
from .formats import FormatBook
from .layout import Registry, SheetSpec
from .names import define_cross_sheet_names
from .sheets import (allergens, budget, converter, dashboard, events,
                     filters, fridge, goals, grocery, help_tab,
                     ingredients, leftovers, lists, meal_plan, members,
                     monthly, nutrition, pantry, prep, printables,
                     recipe_costs, recipe_lines, recipes, reports,
                     rotation, settings, spending, staples, start_here,
                     store_prices, stores, sustainability, waste,
                     weekly)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    members, recipes, recipe_lines, ingredients, meal_plan, grocery,
    pantry, leftovers, waste, spending, store_prices, stores, prep,
    staples, events, goals,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Start Here", "Dashboard", "Weekly Planner", "Monthly Planner", "Settings", "Lists",
    "Household", "Recipes", "Recipe Ingredients", "Ingredients",
    "Meal Plan", "Grocery List", "Pantry", "Staples", "Leftovers",
    "Food Waste", "Meal Prep", "Dietary Filters", "Nutrition",
    "Budget", "Recipe Costs", "Converter", "Rotation",
    "Food Safety", "Fridge & Freezer",
    "Grocery Spending", "Store Prices",
    "Stores", "Events", "Goals", "Sustainability", "Reports",
    "Print Week", "Print Shopping", "Print Recipe", "Help",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Start Here": start_here.build,
    "Dashboard": dashboard.build,
    "Weekly Planner": weekly.build,
    "Monthly Planner": monthly.build,
    "Nutrition": nutrition.build,
    "Budget": budget.build,
    "Dietary Filters": filters.build,
    "Recipe Costs": recipe_costs.build,
    "Converter": converter.build,
    "Rotation": rotation.build,
    "Food Safety": allergens.build,
    "Fridge & Freezer": fridge.build,
    "Settings": settings.build,
    "Lists": lists.build,
    "Sustainability": sustainability.build,
    "Reports": reports.build,
    "Print Week": printables.build_week,
    "Print Shopping": printables.build_shopping,
    "Print Recipe": printables.build_recipe,
    "Help": help_tab.build,
}

#: register modules by sheet name
REGISTER_BY_NAME = {module.NAME: module for module in REGISTER_MODULES}


# ---------------------------------------------------------------------------

def build_register(book, fmts: FormatBook, cfg: Config, reg: Registry,
                   module) -> SheetWriter:
    """Create one register sheet from its :class:`SheetSpec`."""
    spec: SheetSpec = reg[module.NAME]
    sheet = SheetWriter(book, fmts, cfg, reg, spec.name,
                        tab_color=spec.tab_color)
    sheet.grid(spec)
    decorate = getattr(module, "decorate", None)
    if decorate is not None:
        wants_registry = len(inspect.signature(decorate).parameters) >= 4
        if wants_registry:
            decorate(sheet, fmts, cfg, reg)
        else:
            decorate(sheet, fmts, cfg)
    sheet.protect()
    return sheet


def build_workbook(path: str = "Meal Planner & Grocery List.xlsx",
                   cfg: Optional[Config] = None) -> str:
    """Write the whole product to *path* and return the path."""
    cfg = cfg or Config()
    folder = os.path.dirname(os.path.abspath(path))
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)

    book = Workbook(path)
    _set_properties(book, cfg)

    fmts = FormatBook(book, cfg)
    reg = Registry()

    # 1. Register every spec first so cross sheet formulas resolve no
    #    matter which tab happens to be written first.
    for module in REGISTER_MODULES:
        reg.add(module.spec(cfg))

    # 2. Write the tabs in their final order.
    sheets: Dict[str, SheetWriter] = {}
    for name in SHEET_ORDER:
        if name in REGISTER_BY_NAME:
            sheets[name] = build_register(book, fmts, cfg, reg,
                                          REGISTER_BY_NAME[name])
        elif name in PAGE_BUILDERS:
            sheets[name] = PAGE_BUILDERS[name](book, fmts, cfg, reg)
        else:                                   # pragma: no cover
            raise KeyError(f"No builder for sheet '{name}'")

    # 3. Live lookup names (recipe IDs, ingredient names, who lives here).
    define_cross_sheet_names(book, reg)

    # 4. Final touches: open on the first page.
    first = sheets[SHEET_ORDER[0]].ws
    first.activate()
    first.set_first_sheet()

    book.close()
    return path


def _set_properties(book, cfg: Config) -> None:
    book.set_properties({
        "title": cfg.product_name,
        "subject": "Meal planning, grocery lists and food budgeting",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Household spreadsheet template",
        "keywords": ("meal planner, weekly meal planner, monthly meal "
                     "planner, recipe organizer, grocery list, shopping "
                     "list, pantry inventory, food budget, meal prep, "
                     "nutrition tracker, spreadsheet template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

### `meal_planner_spreadsheet/sample_data.py`

*855 lines · 1 class · 4 functions*

> Demo data: four weeks in the life of the Harper household.
>
> Two adults, a teenager and a six year old, a £180 weekly budget, a
> freezer with half a lasagne in it and a standing argument about
> mushrooms. Every tab opens with believable numbers so the dashboard,
> the grocery list and the reports mean something the moment you open the
> file - and `--blank` strips the lot out for your own data.
>
> Everything is dated relative to the build date, so the workbook never
> looks stale.

| Defines | What it does |
|---------|--------------|
| `d(base, offset)` |  |
| `week_monday(day)` |  |
| `class Sample` | fields: members, stores, ingredients, recipes, recipe_lines, meal_plan, pantry, leftovers |
| `get(cfg)` | Build (and cache) the demo data for this configuration |
| `_build(cfg)` |  |

**Module constants:** `INGREDIENTS` (47 items), `RECIPES` (18 items), `RECIPE_LINES` (18 entries).

````python
"""Demo data: four weeks in the life of the Harper household.

Two adults, a teenager and a six year old, a £180 weekly budget, a
freezer with half a lasagne in it and a standing argument about
mushrooms. Every tab opens with believable numbers so the dashboard,
the grocery list and the reports mean something the moment you open the
file - and `--blank` strips the lot out for your own data.

Everything is dated relative to the build date, so the workbook never
looks stale.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .config import Config


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def week_monday(day: _dt.date) -> _dt.date:
    return day - _dt.timedelta(days=day.weekday())


@dataclass
class Sample:
    members: List[Dict[str, Any]] = field(default_factory=list)
    stores: List[Dict[str, Any]] = field(default_factory=list)
    ingredients: List[Dict[str, Any]] = field(default_factory=list)
    recipes: List[Dict[str, Any]] = field(default_factory=list)
    recipe_lines: List[Dict[str, Any]] = field(default_factory=list)
    meal_plan: List[Dict[str, Any]] = field(default_factory=list)
    pantry: List[Dict[str, Any]] = field(default_factory=list)
    leftovers: List[Dict[str, Any]] = field(default_factory=list)
    waste: List[Dict[str, Any]] = field(default_factory=list)
    spending: List[Dict[str, Any]] = field(default_factory=list)
    store_prices: List[Dict[str, Any]] = field(default_factory=list)
    prep: List[Dict[str, Any]] = field(default_factory=list)
    staples: List[Dict[str, Any]] = field(default_factory=list)
    events: List[Dict[str, Any]] = field(default_factory=list)
    goals: List[Dict[str, Any]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# The price book
# ---------------------------------------------------------------------------
# name, category, aisle, store, unit, pack, price, kcal, pro, carb, fat,
# fibre, allergen, diet, shelf days, storage
INGREDIENTS = [
    ("Chicken breast", "Meat & seafood", "4 - Meat & fish",
     "Greenfield Market", "g", 600, 6.50, 165, 31, 0, 3.6, 0, "None",
     "No restriction", 3, "Fridge"),
    ("Chicken thighs", "Meat & seafood", "4 - Meat & fish",
     "Discount Larder", "g", 900, 4.80, 209, 26, 0, 11, 0, "None",
     "No restriction", 3, "Fridge"),
    ("Beef mince 5%", "Meat & seafood", "4 - Meat & fish",
     "Greenfield Market", "g", 500, 4.75, 137, 21, 0, 5, 0, "None",
     "No restriction", 3, "Fridge"),
    ("Salmon fillet", "Meat & seafood", "4 - Meat & fish",
     "Hill Farm Shop", "g", 240, 5.20, 208, 20, 0, 13, 0, "Fish",
     "Pescatarian", 2, "Fridge"),
    ("Eggs", "Dairy & eggs", "5 - Dairy & eggs", "Hill Farm Shop",
     "each", 12, 3.60, 72, 6.3, 0.4, 4.8, 0, "Eggs", "Vegetarian", 21,
     "Fridge"),
    ("Milk", "Dairy & eggs", "5 - Dairy & eggs", "Corner Express",
     "ml", 2272, 1.65, 42, 3.4, 4.8, 1, 0, "Milk", "Vegetarian", 7,
     "Fridge"),
    ("Cheddar", "Dairy & eggs", "5 - Dairy & eggs", "Greenfield Market",
     "g", 400, 3.40, 402, 25, 1.3, 33, 0, "Milk", "Vegetarian", 21,
     "Fridge"),
    ("Greek yoghurt", "Dairy & eggs", "5 - Dairy & eggs",
     "Greenfield Market", "g", 500, 2.10, 97, 9, 4, 5, 0, "Milk",
     "Vegetarian", 10, "Fridge"),
    ("Butter", "Dairy & eggs", "5 - Dairy & eggs", "Discount Larder",
     "g", 250, 2.25, 717, 0.9, 0.1, 81, 0, "Milk", "Vegetarian", 40,
     "Fridge"),
    ("Parmesan", "Dairy & eggs", "3 - Deli & cheese",
     "Greenfield Market", "g", 150, 3.20, 431, 38, 4.1, 29, 0, "Milk",
     "Vegetarian", 60, "Fridge"),
    ("Onions", "Produce", "1 - Produce", "Discount Larder", "g", 1000,
     0.95, 40, 1.1, 9.3, 0.1, 1.7, "None", "Vegan", 30, "Cupboard"),
    ("Garlic", "Produce", "1 - Produce", "Discount Larder", "each", 3,
     0.70, 4, 0.2, 1, 0, 0.1, "None", "Vegan", 60, "Cupboard"),
    ("Carrots", "Produce", "1 - Produce", "Hill Farm Shop", "g", 1000,
     0.80, 41, 0.9, 9.6, 0.2, 2.8, "None", "Vegan", 21, "Fridge"),
    ("Potatoes", "Produce", "1 - Produce", "Hill Farm Shop", "g", 2500,
     2.20, 77, 2, 17, 0.1, 2.2, "None", "Vegan", 21, "Cupboard"),
    ("Sweet potato", "Produce", "1 - Produce", "Greenfield Market",
     "g", 1000, 1.90, 86, 1.6, 20, 0.1, 3, "None", "Vegan", 14,
     "Cupboard"),
    ("Broccoli", "Produce", "1 - Produce", "Greenfield Market", "g",
     400, 1.10, 34, 2.8, 7, 0.4, 2.6, "None", "Vegan", 6, "Fridge"),
    ("Peppers", "Produce", "1 - Produce", "Greenfield Market", "each",
     3, 1.75, 31, 1, 6, 0.3, 2.1, "None", "Vegan", 8, "Fridge"),
    ("Courgette", "Produce", "1 - Produce", "Hill Farm Shop", "each",
     2, 1.10, 17, 1.2, 3.1, 0.3, 1, "None", "Vegan", 7, "Fridge"),
    ("Spinach", "Produce", "1 - Produce", "Greenfield Market", "g",
     250, 1.60, 23, 2.9, 3.6, 0.4, 2.2, "None", "Vegan", 4, "Fridge"),
    ("Mushrooms", "Produce", "1 - Produce", "Greenfield Market", "g",
     300, 1.20, 22, 3.1, 3.3, 0.3, 1, "None", "Vegan", 5, "Fridge"),
    ("Tomatoes", "Produce", "1 - Produce", "Hill Farm Shop", "g", 500,
     1.45, 18, 0.9, 3.9, 0.2, 1.2, "None", "Vegan", 7, "Fruit bowl"),
    ("Chopped tomatoes", "Pantry staples", "7 - Tins & jars",
     "Discount Larder", "g", 400, 0.55, 32, 1.6, 7, 0.3, 1.5, "None",
     "Vegan", 540, "Cupboard"),
    ("Lemon", "Produce", "1 - Produce", "Greenfield Market", "each", 4,
     1.00, 29, 1.1, 9.3, 0.3, 2.8, "None", "Vegan", 14, "Fruit bowl"),
    ("Bananas", "Produce", "1 - Produce", "Discount Larder", "each", 6,
     1.05, 89, 1.1, 23, 0.3, 2.6, "None", "Vegan", 6, "Fruit bowl"),
    ("Apples", "Produce", "1 - Produce", "Hill Farm Shop", "each", 6,
     1.80, 52, 0.3, 14, 0.2, 2.4, "None", "Vegan", 14, "Fruit bowl"),
    ("Frozen peas", "Frozen", "6 - Frozen", "Discount Larder", "g",
     900, 1.40, 81, 5.4, 14, 0.4, 5.1, "None", "Vegan", 365, "Freezer"),
    ("Frozen berries", "Frozen", "6 - Frozen", "Greenfield Market",
     "g", 500, 2.75, 58, 0.8, 12, 0.4, 3.5, "None", "Vegan", 365,
     "Freezer"),
    ("Pasta", "Pantry staples", "8 - Pasta, rice & grains",
     "Discount Larder", "g", 1000, 1.15, 371, 13, 75, 1.5, 3.2,
     "Cereals with gluten", "Vegan", 540, "Cupboard"),
    ("Rice", "Pantry staples", "8 - Pasta, rice & grains",
     "Discount Larder", "g", 2000, 2.60, 356, 7.5, 78, 0.9, 1.4,
     "None", "Vegan", 720, "Cupboard"),
    ("Couscous", "Pantry staples", "8 - Pasta, rice & grains",
     "Greenfield Market", "g", 500, 1.40, 376, 13, 77, 0.6, 5,
     "Cereals with gluten", "Vegan", 540, "Cupboard"),
    ("Red lentils", "Pantry staples", "7 - Tins & jars",
     "Discount Larder", "g", 500, 1.20, 352, 24, 60, 1.1, 11, "None",
     "Vegan", 720, "Cupboard"),
    ("Chickpeas", "Pantry staples", "7 - Tins & jars",
     "Discount Larder", "g", 400, 0.60, 139, 7.5, 21, 2.4, 6.4, "None",
     "Vegan", 720, "Cupboard"),
    ("Black beans", "Pantry staples", "7 - Tins & jars",
     "Discount Larder", "g", 400, 0.75, 132, 8.9, 24, 0.5, 8.7, "None",
     "Vegan", 720, "Cupboard"),
    ("Coconut milk", "Pantry staples", "7 - Tins & jars",
     "Online Grocer", "ml", 400, 1.30, 197, 2, 3, 20, 0, "None",
     "Vegan", 540, "Cupboard"),
    ("Olive oil", "Pantry staples", "10 - Spices & sauces",
     "Greenfield Market", "ml", 500, 4.50, 884, 0, 0, 100, 0, "None",
     "Vegan", 540, "Cupboard"),
    ("Soy sauce", "Pantry staples", "10 - Spices & sauces",
     "Online Grocer", "ml", 250, 1.85, 53, 8, 4.9, 0.1, 0.8, "Soya",
     "Vegan", 720, "Cupboard"),
    ("Curry paste", "Pantry staples", "10 - Spices & sauces",
     "Greenfield Market", "g", 180, 2.30, 180, 3, 12, 13, 3, "None",
     "Vegetarian", 365, "Cupboard"),
    ("Stock cubes", "Pantry staples", "10 - Spices & sauces",
     "Discount Larder", "each", 12, 1.10, 6, 0.3, 0.9, 0.1, 0, "None",
     "Vegan", 540, "Cupboard"),
    ("Paprika", "Spices & seasonings", "10 - Spices & sauces",
     "Greenfield Market", "g", 50, 1.25, 282, 14, 54, 13, 35, "None",
     "Vegan", 720, "Spice rack"),
    ("Cumin", "Spices & seasonings", "10 - Spices & sauces",
     "Greenfield Market", "g", 45, 1.20, 375, 18, 44, 22, 11, "None",
     "Vegan", 720, "Spice rack"),
    ("Bread", "Bakery", "2 - Bakery", "Corner Express", "g", 800, 1.35,
     265, 9, 49, 3.2, 2.7, "Cereals with gluten", "Vegan", 5,
     "Cupboard"),
    ("Tortilla wraps", "Bakery", "2 - Bakery", "Greenfield Market",
     "each", 8, 1.50, 218, 7, 36, 5, 2, "Cereals with gluten", "Vegan",
     14, "Cupboard"),
    ("Porridge oats", "Pantry staples", "8 - Pasta, rice & grains",
     "Discount Larder", "g", 1000, 1.30, 379, 13, 67, 6.5, 10, "None",
     "Vegan", 365, "Cupboard"),
    ("Peanut butter", "Pantry staples", "9 - Baking", "Online Grocer",
     "g", 340, 2.95, 588, 25, 20, 50, 6, "Peanuts", "Vegan", 365,
     "Cupboard"),
    ("Honey", "Pantry staples", "9 - Baking", "Hill Farm Shop", "g",
     340, 3.80, 304, 0.3, 82, 0, 0.2, "None", "Vegetarian", 720,
     "Cupboard"),
    ("Flour", "Pantry staples", "9 - Baking", "Discount Larder", "g",
     1500, 0.90, 364, 10, 76, 1, 2.7, "Cereals with gluten", "Vegan",
     365, "Cupboard"),
    ("Tofu", "Pantry staples", "5 - Dairy & eggs", "Online Grocer",
     "g", 396, 2.20, 76, 8, 1.9, 4.8, 0.3, "Soya", "Vegan", 14,
     "Fridge"),
]

# ---------------------------------------------------------------------------
# The recipe book
# ---------------------------------------------------------------------------
# name, category, slot, cuisine, serves, prep, cook, method, difficulty,
# diet, freezes, kid friendly, batch friendly, rating, tags
RECIPES = [
    ("Weeknight chicken traybake", "Main", "Dinner", "British", 4, 15,
     45, "Oven", "Easy", "No restriction", "No", "Yes", "Yes", 5,
     "one tray, no washing up"),
    ("Beef bolognese", "Main", "Dinner", "Italian", 6, 20, 60, "Hob",
     "Easy", "No restriction", "Yes", "Yes", "Yes", 5,
     "batch base, freezer friendly"),
    ("Red lentil dhal", "Main", "Dinner", "Indian", 4, 10, 30,
     "One pot", "Easy", "Vegan", "Yes", "No", "Yes", 4,
     "cheap, store cupboard"),
    ("Salmon, new potatoes & greens", "Main", "Dinner", "British", 2,
     10, 25, "Oven", "Easy", "Pescatarian", "No", "No", "No", 4,
     "payday treat"),
    ("Chickpea & spinach curry", "Main", "Dinner", "Indian", 4, 10, 25,
     "One pot", "Easy", "Vegan", "Yes", "No", "Yes", 4,
     "store cupboard"),
    ("Tomato & mascarpone pasta", "Main", "Dinner", "Italian", 4, 5,
     20, "Hob", "Easy", "Vegetarian", "No", "Yes", "No", 4,
     "twenty minutes flat"),
    ("Chicken fajitas", "Main", "Dinner", "Mexican", 4, 15, 20, "Hob",
     "Easy", "No restriction", "No", "Yes", "No", 5, "friday night"),
    ("Black bean chilli", "Main", "Dinner", "Mexican", 6, 15, 40,
     "Slow cooker", "Easy", "Vegan", "Yes", "Yes", "Yes", 4,
     "batch base"),
    ("Veg-packed shepherd's pie", "Main", "Dinner", "British", 6, 25,
     50, "Oven", "Medium", "No restriction", "Yes", "Yes", "Yes", 5,
     "hides the carrots"),
    ("Teriyaki salmon & rice", "Main", "Dinner", "Japanese", 2, 10, 20,
     "Hob", "Easy", "Pescatarian", "No", "No", "No", 4, ""),
    ("Overnight oats", "Breakfast", "Breakfast", "British", 2, 5, 0,
     "No cook", "Easy", "Vegetarian", "No", "Yes", "Yes", 4,
     "make it the night before"),
    ("Scrambled eggs on toast", "Breakfast", "Breakfast", "British", 2,
     5, 8, "Hob", "Easy", "Vegetarian", "No", "Yes", "No", 4, ""),
    ("Banana pancakes", "Breakfast", "Breakfast", "American", 4, 10,
     15, "Hob", "Easy", "Vegetarian", "Yes", "Yes", "No", 5,
     "saturday morning"),
    ("Couscous lunch jars", "Lunchbox", "Lunch", "Middle Eastern", 4,
     15, 10, "No cook", "Easy", "Vegan", "No", "No", "Yes", 4,
     "lunchbox, prep ahead"),
    ("Hummus & veg wraps", "Lunchbox", "Lunch", "Middle Eastern", 2, 10,
     0, "No cook", "Easy", "Vegan", "No", "Yes", "No", 3, "lunchbox"),
    ("Leftover soup", "Soup", "Lunch", "British", 4, 10, 25, "One pot",
     "Easy", "Vegan", "Yes", "No", "Yes", 4, "uses up leftovers"),
    ("Tofu stir fry", "Main", "Dinner", "Chinese", 4, 15, 15, "Hob",
     "Easy", "Vegan", "No", "No", "No", 3, ""),
    ("Apple crumble", "Dessert", "Dessert", "British", 6, 20, 35,
     "Oven", "Easy", "Vegetarian", "Yes", "Yes", "No", 5,
     "sunday pudding"),
]

#: recipe index -> [(ingredient, qty, unit, prep, optional)]
RECIPE_LINES: Dict[int, List[Any]] = {
    0: [("Chicken thighs", 800, "g", "skin on", "No"),
        ("Potatoes", 700, "g", "cut into chunks", "No"),
        ("Peppers", 2, "each", "sliced", "No"),
        ("Onions", 200, "g", "wedges", "No"),
        ("Olive oil", 2, "tbsp", "", "No"),
        ("Paprika", 2, "tsp", "", "No")],
    1: [("Beef mince 5%", 500, "g", "", "No"),
        ("Onions", 150, "g", "finely chopped", "No"),
        ("Carrots", 150, "g", "grated", "No"),
        ("Garlic", 2, "clove", "crushed", "No"),
        ("Chopped tomatoes", 800, "g", "", "No"),
        ("Pasta", 500, "g", "", "No"),
        ("Parmesan", 40, "g", "grated", "Yes")],
    2: [("Red lentils", 300, "g", "rinsed", "No"),
        ("Onions", 150, "g", "chopped", "No"),
        ("Garlic", 3, "clove", "crushed", "No"),
        ("Curry paste", 60, "g", "", "No"),
        ("Coconut milk", 400, "ml", "", "No"),
        ("Spinach", 100, "g", "", "Yes")],
    3: [("Salmon fillet", 240, "g", "", "No"),
        ("Potatoes", 500, "g", "halved", "No"),
        ("Broccoli", 200, "g", "", "No"),
        ("Lemon", 1, "each", "halved", "No"),
        ("Olive oil", 1, "tbsp", "", "No")],
    4: [("Chickpeas", 800, "g", "drained", "No"),
        ("Spinach", 200, "g", "", "No"),
        ("Onions", 150, "g", "chopped", "No"),
        ("Curry paste", 50, "g", "", "No"),
        ("Chopped tomatoes", 400, "g", "", "No"),
        ("Rice", 300, "g", "", "No")],
    5: [("Pasta", 400, "g", "", "No"),
        ("Chopped tomatoes", 400, "g", "", "No"),
        ("Garlic", 2, "clove", "sliced", "No"),
        ("Olive oil", 2, "tbsp", "", "No"),
        ("Parmesan", 50, "g", "grated", "No")],
    6: [("Chicken breast", 500, "g", "sliced", "No"),
        ("Peppers", 3, "each", "sliced", "No"),
        ("Onions", 150, "g", "sliced", "No"),
        ("Tortilla wraps", 8, "each", "", "No"),
        ("Cheddar", 100, "g", "grated", "Yes"),
        ("Paprika", 2, "tsp", "", "No")],
    7: [("Black beans", 800, "g", "drained", "No"),
        ("Chopped tomatoes", 800, "g", "", "No"),
        ("Onions", 200, "g", "chopped", "No"),
        ("Peppers", 2, "each", "diced", "No"),
        ("Cumin", 2, "tsp", "", "No"),
        ("Rice", 400, "g", "", "No")],
    8: [("Beef mince 5%", 500, "g", "", "No"),
        ("Potatoes", 1000, "g", "mashed", "No"),
        ("Carrots", 300, "g", "diced", "No"),
        ("Frozen peas", 200, "g", "", "No"),
        ("Onions", 150, "g", "chopped", "No"),
        ("Stock cubes", 1, "each", "", "No"),
        ("Butter", 40, "g", "for the mash", "No")],
    9: [("Salmon fillet", 240, "g", "", "No"),
        ("Rice", 150, "g", "", "No"),
        ("Soy sauce", 3, "tbsp", "", "No"),
        ("Honey", 1, "tbsp", "", "No"),
        ("Broccoli", 150, "g", "", "No")],
    10: [("Porridge oats", 100, "g", "", "No"),
         ("Milk", 300, "ml", "", "No"),
         ("Greek yoghurt", 100, "g", "", "No"),
         ("Frozen berries", 100, "g", "", "No"),
         ("Honey", 1, "tbsp", "", "Yes")],
    11: [("Eggs", 4, "each", "beaten", "No"),
         ("Bread", 120, "g", "toasted", "No"),
         ("Butter", 20, "g", "", "No"),
         ("Milk", 30, "ml", "", "Yes")],
    12: [("Flour", 200, "g", "", "No"),
         ("Bananas", 2, "each", "mashed", "No"),
         ("Eggs", 2, "each", "", "No"),
         ("Milk", 250, "ml", "", "No"),
         ("Butter", 20, "g", "for the pan", "No")],
    13: [("Couscous", 200, "g", "", "No"),
         ("Chickpeas", 400, "g", "drained", "No"),
         ("Tomatoes", 200, "g", "diced", "No"),
         ("Courgette", 1, "each", "roasted", "No"),
         ("Lemon", 1, "each", "juiced", "No"),
         ("Olive oil", 2, "tbsp", "", "No")],
    14: [("Tortilla wraps", 2, "each", "", "No"),
         ("Chickpeas", 200, "g", "mashed", "No"),
         ("Peppers", 1, "each", "sliced", "No"),
         ("Spinach", 50, "g", "", "No")],
    15: [("Carrots", 300, "g", "chopped", "No"),
         ("Potatoes", 300, "g", "chopped", "No"),
         ("Onions", 150, "g", "chopped", "No"),
         ("Stock cubes", 2, "each", "", "No"),
         ("Olive oil", 1, "tbsp", "", "No")],
    16: [("Tofu", 396, "g", "cubed", "No"),
         ("Broccoli", 200, "g", "florets", "No"),
         ("Peppers", 2, "each", "sliced", "No"),
         ("Soy sauce", 4, "tbsp", "", "No"),
         ("Rice", 300, "g", "", "No"),
         ("Garlic", 2, "clove", "sliced", "No")],
    17: [("Apples", 6, "each", "peeled and sliced", "No"),
         ("Flour", 200, "g", "", "No"),
         ("Butter", 100, "g", "cold, cubed", "No"),
         ("Honey", 60, "g", "", "Yes")],
}

_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg)
    return _CACHE[key]


def _build(cfg: Config) -> Sample:
    rnd = random.Random(20261011)
    today = cfg.today
    monday = week_monday(today)
    start = d(monday, -14)          # two weeks of history, two weeks ahead
    s = Sample()

    # ---------------------------------------------------------- household --
    s.members = [
        {"name": "Maya Harper", "role": "Cook", "age_group": "Adult",
         "portion": 1, "diet": "No restriction", "allergies": "None",
         "intolerances": "", "dislikes": "Olives",
         "loves": "Anything with chilli", "breakfast": "Yes",
         "lunch": "No", "dinner": "Yes", "packed": "No",
         "cal_target": 2000, "cooks": "Yes",
         "cook_days": "Most nights", "shops": "Yes",
         "jobs": "Meal plan, big shop", "birthday": d(start, 120),
         "status": "Active",
         "notes": "Does the planning on a Friday with a coffee."},
        {"name": "Tom Harper", "role": "Shopper", "age_group": "Adult",
         "portion": 1.2, "diet": "No restriction", "allergies": "None",
         "intolerances": "", "dislikes": "Mushrooms",
         "loves": "Curry night", "breakfast": "No", "lunch": "No",
         "dinner": "Yes", "packed": "Yes", "cal_target": 2400,
         "cooks": "Yes", "cook_days": "Tuesday and Thursday",
         "shops": "Yes", "jobs": "Top-up shop, washing up",
         "birthday": d(start, 35), "status": "Active",
         "notes": "Takes a packed lunch four days out of five."},
        {"name": "Rory Harper", "role": "Helper", "age_group": "Teenager",
         "portion": 1.4, "diet": "Vegetarian", "allergies": "None",
         "intolerances": "", "dislikes": "Anything 'mixed up'",
         "loves": "Pasta, pancakes", "breakfast": "Yes", "lunch": "Yes",
         "dinner": "Yes", "packed": "Yes", "cal_target": 2600,
         "cooks": "Yes", "cook_days": "Sunday pudding",
         "shops": "No", "jobs": "Lunchboxes", "birthday": d(start, 200),
         "status": "Active",
         "notes": "Went vegetarian in March - needs a protein swap on "
                  "meat nights."},
        {"name": "Esme Harper", "role": "Child", "age_group": "Child",
         "portion": 0.6, "diet": "No restriction",
         "allergies": "Peanuts", "intolerances": "",
         "dislikes": "Mushrooms, anything spicy", "loves": "Pancakes",
         "breakfast": "Yes", "lunch": "Yes", "dinner": "Yes",
         "packed": "Yes", "cal_target": 1400, "cooks": "No",
         "cook_days": "", "shops": "No", "jobs": "Laying the table",
         "birthday": d(start, 70), "status": "Active",
         "notes": "Peanut allergy - no peanut butter in the house, "
                  "check every label."},
    ]

    # ------------------------------------------------------------ stores --
    s.stores = [
        {"name": "Greenfield Market", "type": "Supermarket",
         "location": "High Street", "travel_min": 12,
         "best_for": "The weekly shop, decent veg", "delivery": "Yes",
         "delivery_fee": 3.50, "min_order": 40,
         "slot": "Saturday 9-10am", "loyalty": "Greenfield points",
         "open_hours": "7am - 10pm", "price_level": "Mid-range",
         "rating": 4, "status": "Active",
         "notes": "Default store for most of the price book."},
        {"name": "Discount Larder", "type": "Discounter",
         "location": "Retail park", "travel_min": 18,
         "best_for": "Store cupboard, tins, frozen", "delivery": "No",
         "delivery_fee": None, "min_order": None, "slot": "",
         "loyalty": "", "open_hours": "8am - 8pm",
         "price_level": "Budget", "rating": 4, "status": "Active",
         "notes": "Worth a monthly stock-up run, not a weekly trip."},
        {"name": "Hill Farm Shop", "type": "Farm shop",
         "location": "Hill Lane", "travel_min": 20,
         "best_for": "Veg box, eggs, meat for special occasions",
         "delivery": "Yes", "delivery_fee": 2.00, "min_order": 25,
         "slot": "Thursday afternoon", "loyalty": "",
         "open_hours": "9am - 5pm", "price_level": "Premium",
         "rating": 5, "status": "Active",
         "notes": "Dearer per 100g but the veg lasts twice as long."},
        {"name": "Corner Express", "type": "Convenience",
         "location": "End of the road", "travel_min": 3,
         "best_for": "Milk and bread emergencies", "delivery": "No",
         "delivery_fee": None, "min_order": None, "slot": "",
         "loyalty": "", "open_hours": "6am - 11pm",
         "price_level": "Premium", "rating": 3, "status": "Active",
         "notes": "Convenient and expensive - the top-up habit to "
                  "break."},
        {"name": "Online Grocer", "type": "Online delivery",
         "location": "Delivered", "travel_min": 0,
         "best_for": "Heavy and bulky, world food aisle",
         "delivery": "Yes", "delivery_fee": 4.95, "min_order": 50,
         "slot": "Wednesday 7-8pm", "loyalty": "Free delivery pass",
         "open_hours": "Any time", "price_level": "Mid-range",
         "rating": 4, "status": "Active", "notes": ""},
        {"name": "Curry House", "type": "Takeaway",
         "location": "High Street", "travel_min": 10,
         "best_for": "Friday night when nobody wants to cook",
         "delivery": "Yes", "delivery_fee": 2.50, "min_order": 20,
         "slot": "", "loyalty": "", "open_hours": "5pm - 11pm",
         "price_level": "Premium", "rating": 4, "status": "Active",
         "notes": "Logged here so the takeaway spend lands in the "
                  "same budget as the shopping."},
        {"name": "The Lantern", "type": "Restaurant",
         "location": "Market Square", "travel_min": 15,
         "best_for": "Birthdays and anniversaries", "delivery": "No",
         "delivery_fee": None, "min_order": None, "slot": "",
         "loyalty": "", "open_hours": "12pm - 10pm",
         "price_level": "Premium", "rating": 5, "status": "Active",
         "notes": "Eating out counts as food spend - that is rather "
                  "the point of tracking it."},
    ]

    # ------------------------------------------------------- ingredients --
    for row in INGREDIENTS:
        (name, cat, aisle, store, unit, pack, price, kcal, pro, carb,
         fat, fibre, allergen, diet, shelf, storage) = row
        s.ingredients.append({
            "name": name, "category": cat, "aisle": aisle,
            "store": store, "base_unit": unit, "pack_size": pack,
            "pack_price": price, "calories": kcal, "protein": pro,
            "carbs": carb, "fat": fat, "fibre": fibre,
            "sugar": round(carb * 0.25, 1), "sodium": int(kcal * 0.8),
            "allergen": allergen, "diet": diet, "season": "All year",
            "shelf_days": shelf, "storage": storage,
            "last_bought": d(monday, -rnd.randint(1, 10)),
            "brand": "" if rnd.random() < 0.6 else "Own brand",
            "substitute": "", "notes": "",
        })

    # ----------------------------------------------------------- recipes --
    for idx, row in enumerate(RECIPES):
        (name, cat, slot, cuisine, serves, prep, cook, method, diff,
         diet, freezes, kid, batch, rating, tags) = row
        lines = RECIPE_LINES.get(idx, [])
        s.recipes.append({
            "name": name, "category": cat, "meal_slot": slot,
            "cuisine": cuisine, "servings": serves, "prep_min": prep,
            "cook_min": cook, "method": method, "difficulty": diff,
            "diet": diet, "freezes": freezes, "kid_friendly": kid,
            "prep_friendly": batch, "rating": rating,
            "family_score": max(1, min(5, rating - (idx % 2))),
            "favourite": "Yes" if rating >= 5 else "No",
            "spice": "Mild" if cuisine in ("Indian", "Mexican") else "None",
            "season": "All year", "tags": tags,
            "one_pot": "Yes" if method in ("One pot", "Slow cooker")
                       else "No",
            "no_cook": "Yes" if method == "No cook" else "No",
            "description": f"{name} - {len(lines)} ingredients, serves "
                           f"{serves}.",
            "equipment": {"Oven": "Roasting tin", "Slow cooker":
                          "Slow cooker", "Hob": "Big pan"}.get(method, ""),
            "source": "Family favourite" if rating >= 5 else "",
            "link": "", "photo": "",
            "method_notes": "1. Prep everything first. 2. Cook as usual. "
                            "3. Taste before you serve it.",
            "added": d(start, -rnd.randint(30, 400)),
            "notes": "",
        })
        for ing, qty, unit, prep_note, optional in lines:
            s.recipe_lines.append({
                "recipe": f"R-{idx + 1:04d}", "ingredient": ing,
                "qty": qty, "unit": unit, "prep": prep_note,
                "optional": optional, "section": "",
            })

    # --------------------------------------------------------- meal plan --
    dinners = [0, 1, 2, 5, 6, 3, 8, 7, 4, 9, 16, 1, 0, 2]
    breakfasts = [11, 10, 11, 12, 10, 11, 12]
    lunches = [13, 14, 15, 14, 13, 15, 14]
    statuses = {True: "Eaten", False: "Planned"}
    for day_offset in range(28):
        day = d(start, day_offset)
        past = day < today
        # breakfast on most days
        if day_offset % 7 not in (2, 4):
            r = breakfasts[day_offset % len(breakfasts)]
            s.meal_plan.append({
                "date": day, "slot": "Breakfast",
                "recipe": f"R-{r + 1:04d}", "servings": 3,
                "shop": "Yes" if not past else "No",
                "status": statuses[past], "cook": "Maya Harper",
                "prep_ahead": "No",
                "notes": "",
            })
        # lunch - packed lunches in the week, leftovers at the weekend
        if day.weekday() < 5:
            r = lunches[day_offset % len(lunches)]
            s.meal_plan.append({
                "date": day, "slot": "Lunch", "recipe": f"R-{r + 1:04d}",
                "servings": 2, "shop": "Yes" if not past else "No",
                "status": statuses[past], "cook": "Rory Harper",
                "prep_ahead": "Yes" if day.weekday() == 0 else "No",
                "notes": "Lunchboxes" if day.weekday() < 5 else "",
            })
        else:
            s.meal_plan.append({
                "date": day, "slot": "Lunch", "recipe": None,
                "freetext": "Leftovers from the fridge", "servings": 4,
                "shop": "No", "status": statuses[past],
                "cook": "", "prep_ahead": "No", "notes": "",
            })
        # dinner every day
        if day_offset == 12:
            s.meal_plan.append({
                "date": day, "slot": "Dinner", "recipe": None,
                "freetext": "Takeaway - curry house", "servings": 4,
                "shop": "No", "status": "Takeaway", "cook": "",
                "actual_cost": 38.50, "prep_ahead": "No",
                "notes": "Friday treat. Over budget but worth it."})
        else:
            r = dinners[day_offset % len(dinners)]
            s.meal_plan.append({
                "date": day, "slot": "Dinner", "recipe": f"R-{r + 1:04d}",
                "servings": 4, "shop": "Yes" if not past else "No",
                "status": statuses[past],
                "cook": "Tom Harper" if day.weekday() in (1, 3)
                        else "Maya Harper",
                "leftovers": 2 if r in (1, 7, 8) else None,
                "prep_ahead": "Yes" if r in (1, 7, 8) else "No",
                "event": "Sunday lunch with Nan"
                         if day.weekday() == 6 and day_offset > 20 else "",
                "notes": "",
            })
        # the occasional pudding
        if day.weekday() == 6:
            s.meal_plan.append({
                "date": day, "slot": "Dessert", "recipe": "R-0018",
                "servings": 6, "shop": "Yes" if not past else "No",
                "status": statuses[past], "cook": "Rory Harper",
                "prep_ahead": "No", "notes": ""})

    # ------------------------------------------------------------ pantry --
    pantry_rows = [
        ("Rice", "Cupboard", 1200, "g", "No", 18, 300),
        ("Pasta", "Cupboard", 700, "g", "Yes", 25, 400),
        ("Chopped tomatoes", "Cupboard", 1600, "g", "No", 300, 800),
        ("Chickpeas", "Cupboard", 1200, "g", "No", 400, 400),
        ("Red lentils", "Cupboard", 400, "g", "Yes", 200, 250),
        ("Olive oil", "Cupboard", 300, "ml", "Yes", 200, 250),
        ("Stock cubes", "Cupboard", 7, "each", "Yes", 200, 6),
        ("Paprika", "Spice rack", 30, "g", "Yes", 300, 20),
        ("Cumin", "Spice rack", 12, "g", "Yes", 300, 20),
        ("Flour", "Cupboard", 900, "g", "Yes", 120, 500),
        ("Porridge oats", "Cupboard", 550, "g", "Yes", 90, 400),
        ("Honey", "Cupboard", 210, "g", "Yes", 400, 100),
        ("Soy sauce", "Cupboard", 180, "ml", "Yes", 300, 100),
        ("Curry paste", "Fridge", 90, "g", "Yes", 30, 60),
        ("Milk", "Fridge", 1100, "ml", "Yes", 4, 1000),
        ("Butter", "Fridge", 160, "g", "Yes", 20, 125),
        ("Cheddar", "Fridge", 220, "g", "Yes", 12, 150),
        ("Eggs", "Fridge", 7, "each", "No", 14, 6),
        ("Greek yoghurt", "Fridge", 300, "g", "Yes", 5, 250),
        ("Carrots", "Fridge", 600, "g", "No", 9, 300),
        ("Potatoes", "Cupboard", 1700, "g", "No", 12, 1000),
        ("Onions", "Cupboard", 700, "g", "No", 20, 500),
        ("Spinach", "Fridge", 120, "g", "Yes", 2, 0),
        ("Mushrooms", "Fridge", 150, "g", "Yes", 1, 0),
        ("Frozen peas", "Freezer", 600, "g", "Yes", 200, 400),
        ("Frozen berries", "Freezer", 350, "g", "Yes", 250, 200),
        ("Beef mince 5%", "Freezer", 500, "g", "No", 120, 0),
        ("Chicken thighs", "Freezer", 900, "g", "No", 150, 0),
        ("Bread", "Cupboard", 400, "g", "Yes", 2, 0),
        ("Tortilla wraps", "Cupboard", 4, "each", "Yes", 9, 0),
    ]
    for name, loc, qty, unit, opened, days, minimum in pantry_rows:
        s.pantry.append({
            "ingredient": name, "location": loc, "qty": qty,
            "unit": unit, "opened": opened,
            "bought": d(today, -rnd.randint(2, 20)),
            "expiry": d(today, days) if days <= 30 else None,
            "min_qty": minimum or None,
            "staple": "Yes" if minimum else "No",
            "notes": "Needs using" if days <= 2 else "",
        })

    # --------------------------------------------------------- leftovers --
    leftover_rows = [
        ("Half the bolognese", "R-0002", -2, 3, "Fridge", "Planned"),
        ("Shepherd's pie - two portions", "R-0009", -1, 2, "Fridge",
         "In the fridge"),
        ("Chilli for the freezer", "R-0008", -6, 4, "Freezer", "Frozen"),
        ("Dhal", "R-0003", -3, 2, "Fridge", "Eaten"),
        ("Roast potatoes", None, -4, 1, "Fridge", "Binned"),
        ("Apple crumble", "R-0018", -2, 3, "Fridge", "Eaten"),
        ("Traybake chicken", "R-0001", -1, 2, "Fridge", "Planned"),
        ("Soup", "R-0016", -5, 2, "Freezer", "Frozen"),
    ]
    for item, recipe, offset, portions, where, status in leftover_rows:
        s.leftovers.append({
            "item": item, "recipe": recipe, "cooked": d(today, offset),
            "portions": portions, "location": where,
            "container": "Glass box" if where == "Fridge"
                         else "Freezer bag",
            "planned_for": d(today, 1) if status == "Planned" else None,
            "eaten_by": "Tom Harper" if status == "Planned" else "",
            "status": status,
            "notes": "Label it or it will be a mystery in a month"
                     if where == "Freezer" else "",
        })

    # -------------------------------------------------------- food waste --
    waste_rows = [
        ("Spinach", 120, "g", "Went mouldy", "Yes", "Fridge",
         "Food waste caddy", "Buy the small bag, or freeze half"),
        ("Bread", 300, "g", "Expired", "Yes", "Cupboard", "Compost",
         "Freeze half the loaf on the day we buy it"),
        ("Milk", 500, "ml", "Expired", "Yes", "Fridge",
         "General bin", "Check the date before the top-up shop"),
        ("Mushrooms", 150, "g", "Forgot about it", "Yes", "Fridge",
         "Compost", "Keep them at eye level, not in the drawer"),
        ("Potatoes", 400, "g", "Went mouldy", "Yes", "Cupboard",
         "Compost", "Smaller bag from the farm shop"),
        ("Greek yoghurt", 200, "g", "Expired", "Yes", "Fridge",
         "General bin", "One pot a week is enough"),
        ("Carrots", 200, "g", "Plate scrapings", "No", "Fridge",
         "Food waste caddy", "Esme's portion is too big"),
        ("Tomatoes", 150, "g", "Bought too much", "Yes", "Fruit bowl",
         "Compost", "Only buy them when a recipe needs them"),
    ]
    for i, (item, qty, unit, reason, avoidable, where, disposal,
            prevent) in enumerate(waste_rows):
        s.waste.append({
            "date": d(today, -(i * 3 + 2)), "item": item, "qty": qty,
            "unit": unit, "reason": reason, "avoidable": avoidable,
            "where": where, "disposal": disposal, "prevent": prevent,
        })

    # ---------------------------------------------------------- spending --
    spend_rows = [
        (-21, "Greenfield Market", "Weekly shop", "Weekly shop", 142.80,
         "Yes", "Weekly shop"),
        (-19, "Corner Express", "Top-up shop", "Milk and bread", 6.40,
         "No", "Top-up shop"),
        (-17, "Hill Farm Shop", "Top-up shop", "Veg box", 18.50, "Yes",
         "Weekly shop"),
        (-14, "Greenfield Market", "Weekly shop", "Weekly shop", 151.25,
         "Yes", "Weekly shop"),
        (-12, "Online Grocer", "Bulk shop", "Store cupboard stock-up",
         62.40, "Yes", "Weekly shop"),
        (-11, "Corner Express", "Top-up shop", "Milk", 2.10, "No",
         "Top-up shop"),
        (-9, "Curry House", "Takeaway", "Friday takeaway", 38.50, "No",
         "Takeaway"),
        (-7, "Greenfield Market", "Weekly shop", "Weekly shop", 138.60,
         "Yes", "Weekly shop"),
        (-6, "Discount Larder", "Bulk shop", "Tins and frozen", 31.20,
         "Yes", "Weekly shop"),
        (-5, "Corner Express", "Top-up shop", "Bread", 1.35, "No",
         "Top-up shop"),
        (-4, "The Lantern", "Eating out", "Birthday lunch", 64.00,
         "Yes", "Special occasion"),
        (-2, "Greenfield Market", "Weekly shop", "Weekly shop", 147.95,
         "Yes", "Weekly shop"),
        (-1, "Hill Farm Shop", "Top-up shop", "Eggs and veg", 12.75,
         "Yes", "Weekly shop"),
    ]
    for offset, store, kind, item, amount, planned, line in spend_rows:
        s.spending.append({
            "date": d(today, offset), "store": store, "type": kind,
            "item": item, "amount": amount, "payer": "Maya Harper"
            if amount > 20 else "Tom Harper",
            "method": "Card", "planned": planned, "budget_line": line,
            "waste": "No", "notes": "",
        })

    # ------------------------------------------------------ store prices --
    compare = ["Chicken breast", "Beef mince 5%", "Pasta", "Rice",
               "Milk", "Cheddar", "Olive oil", "Eggs", "Chopped tomatoes",
               "Porridge oats", "Butter", "Bread"]
    lookup = {row[0]: row for row in INGREDIENTS}
    for name in compare:
        base = lookup[name]
        pack, price, unit = base[5], base[6], base[4]
        for store, factor in (("Greenfield Market", 1.0),
                              ("Discount Larder", 0.82),
                              ("Corner Express", 1.28)):
            s.store_prices.append({
                "ingredient": name, "store": store,
                "brand": "Own brand" if factor < 1 else "Branded",
                "pack_size": pack, "price": round(price * factor, 2),
                "offer": "No",
                "checked": d(today, -rnd.randint(3, 60)),
                "notes": "",
            })

    # --------------------------------------------------------- meal prep --
    prep_rows = [
        ("Big pot of bolognese", "R-0002", 2, "Freezer", 6,
         "Maya Harper", "Done"),
        ("Chilli for the freezer", "R-0008", 2, "Freezer", 6,
         "Tom Harper", "Done"),
        ("Couscous lunch jars", "R-0014", 1, "Fridge", 4,
         "Rory Harper", "Done"),
        ("Chop all the veg for the week", None, 1, "Fridge", 0,
         "Maya Harper", "Done"),
        ("Overnight oats x4", "R-0011", 2, "Fridge", 4, "Rory Harper",
         "To do"),
        ("Shepherd's pie - double batch", "R-0009", 2, "Freezer", 4,
         "Maya Harper", "To do"),
        ("Pancake batter for Saturday", "R-0013", 1, "Fridge", 1,
         "Rory Harper", "To do"),
    ]
    for i, (job, recipe, batches, where, containers, who,
            status) in enumerate(prep_rows):
        day = d(monday, -7 + (i * 3))
        s.prep.append({
            "job": job, "date": day, "recipe": recipe,
            "batches": batches, "storage": where,
            "containers": containers or None,
            "covers_from": day, "covers_to": d(day, 6),
            "cook": who, "status": status,
            "minutes": 45 if batches > 1 else 25,
            "notes": "",
        })

    # ---------------------------------------------------------- staples --
    staple_rows = [
        ("Rice", 1000, 250), ("Pasta", 1000, 400),
        ("Chopped tomatoes", 1600, 400), ("Olive oil", 500, 120),
        ("Stock cubes", 10, 3), ("Flour", 1000, 150),
        ("Porridge oats", 1000, 350), ("Milk", 2000, 1800),
        ("Butter", 250, 80), ("Eggs", 12, 8),
        ("Chickpeas", 1200, 300), ("Red lentils", 500, 150),
        ("Frozen peas", 900, 200), ("Soy sauce", 250, 40),
        ("Honey", 340, 40),
    ]
    for name, par, usage in staple_rows:
        s.staples.append({
            "ingredient": name, "par": par, "usage": usage,
            "bulk": "Yes" if par >= 1000 else "No", "notes": "",
        })

    # ----------------------------------------------------------- events --
    s.events = [
        {"name": "Sunday lunch with Nan", "type": "Sunday lunch",
         "date": d(monday, 6), "time": "13:00", "guests": 6,
         "adults": 5, "children": 1, "dietary": "Nan does not eat pork",
         "budget": 35.00, "menu": "Roast chicken, all the trimmings, "
         "apple crumble", "prep_start": d(monday, 5),
         "shopping_done": "No", "host": "Maya Harper",
         "status": "Planned",
         "notes": "Rory is doing the pudding."},
        {"name": "Esme's birthday tea", "type": "Birthday",
         "date": d(today, 23), "time": "16:00", "guests": 12,
         "adults": 4, "children": 8,
         "dietary": "Two peanut allergies, one dairy free",
         "budget": 60.00, "menu": "Pancakes, fruit, chocolate cake",
         "prep_start": d(today, 21), "shopping_done": "No",
         "host": "Maya Harper", "status": "Planned",
         "notes": "Nothing with peanuts in the house that day."},
        {"name": "Book club", "type": "Dinner party",
         "date": d(today, 9), "time": "19:30", "guests": 7,
         "adults": 7, "children": 0, "dietary": "One vegan",
         "budget": 45.00, "menu": "Dhal, flatbreads, salad",
         "prep_start": d(today, 9), "shopping_done": "No",
         "host": "Maya Harper", "status": "Menu set", "notes": ""},
        {"name": "Tom's work barbecue", "type": "Barbecue",
         "date": d(today, -10), "time": "12:00", "guests": 15,
         "adults": 12, "children": 3, "dietary": "",
         "budget": 70.00, "actual_cost": 82.40,
         "menu": "Burgers, salads, the usual", "prep_start": d(today, -11),
         "shopping_done": "Yes", "host": "Tom Harper", "status": "Done",
         "notes": "Bought far too much bread again."},
    ]

    # ------------------------------------------------------------ goals --
    s.goals = [
        {"goal": "Cook at home six nights a week", "area": "Habit",
         "measure": "Home cooked meals a week", "target": 18,
         "direction": "Hit or beat", "why": "Takeaways were creeping "
         "back up to twice a week", "start": d(today, -60),
         "review": d(today, 7), "owner": "Maya Harper",
         "status": "In progress", "notes": ""},
        {"goal": "Keep the grocery shop under £780 a month",
         "area": "Budget", "measure": "Grocery spend this month",
         "target": 780, "direction": "Stay under",
         "why": "It drifted to £900 last winter", "start": d(today, -90),
         "review": d(today, 14), "owner": "Maya Harper",
         "status": "On track", "notes": ""},
        {"goal": "Two new recipes a month", "area": "Variety",
         "measure": "New recipes this month", "target": 2,
         "direction": "Hit or beat",
         "why": "We were eating the same six dinners on a loop",
         "start": d(today, -120), "review": d(today, 21),
         "owner": "Rory Harper", "status": "In progress", "notes": ""},
        {"goal": "Waste less than £15 of food a month", "area": "Waste",
         "measure": "Food waste this month", "target": 15,
         "direction": "Stay under",
         "why": "The salad drawer was a graveyard",
         "start": d(today, -45), "review": d(today, -2),
         "owner": "Tom Harper", "status": "Behind",
         "notes": "Mostly bagged salad and bread."},
        {"goal": "Four takeaways a month, no more", "area": "Budget",
         "measure": "Takeaways a month", "target": 4,
         "direction": "Stay under", "why": "£40 a go adds up fast",
         "start": d(today, -30), "review": d(today, 10),
         "owner": "Tom Harper", "status": "On track", "notes": ""},
        {"goal": "Five packed lunches a week", "area": "Health",
         "measure": "Packed lunches a week", "target": 5,
         "direction": "Hit or beat",
         "why": "Meal deals were £7 a day", "start": d(today, -20),
         "review": d(today, 12), "owner": "Rory Harper",
         "status": "In progress", "notes": ""},
    ]

    return s
````

---

## 3. The food registers

One module per data tab. Each returns a `SheetSpec` - a list of columns with their formulas - and the writer turns that into a formatted, validated, protected sheet. These five are the chain that makes the shopping list write itself.

### `meal_planner_spreadsheet/sheets/__init__.py`

*6 lines*

> One module per tab.
>
> Register sheets expose ``NAME`` and ``spec(cfg)`` (plus an optional
> ``decorate``); page sheets expose ``NAME`` and ``build(book, fmts, cfg,
> reg)``. `build.py` wires them together.

````python
"""One module per tab.

Register sheets expose ``NAME`` and ``spec(cfg)`` (plus an optional
``decorate``); page sheets expose ``NAME`` and ``build(book, fmts, cfg,
reg)``. `build.py` wires them together.
"""
````

### `meal_planner_spreadsheet/sheets/common.py`

*206 lines · 20 functions*

> Small building blocks shared by the register sheets.

| Defines | What it does |
|---------|--------------|
| `id_col(trigger, prefix, digits, width, header, note)` | Auto numbered ID column (blank until the trigger column is filled) |
| `month_col(date_key, key, header, width)` | `yyyy-mm` text column - the key every monthly roll-up groups on |
| `weekday_col(date_key, key, header, width)` |  |
| `week_col(date_key, key, header, width)` | ISO-ish week number, honouring the week start day on Settings |
| `alert_cf()` | Traffic light colouring for the 'Alert' / 'Check' columns |
| `status_cf()` | Colouring for the meal, shopping and prep status columns |
| `rate_cf(target_name)` | Green above the target rate, amber below it |
| `bar_cf(color)` |  |
| `scale_cf()` |  |
| `money_cf()` |  |
| `dup_cf()` |  |
| `flag_yes_no_cf()` |  |
| `yesno(note)` |  |
| `score_validation(maximum)` |  |
| `blank(trigger, body)` | `=IF(trigger="","",body)` with the standard row placeholders |
| `lookup_recipe(target, key)` | Pull a column of the Recipes tab for the recipe ID in `key` |
| `lookup_ingredient(target, key)` | Pull a column of the Ingredients price book for the name in `key` |
| `to_base(qty_key, unit_key)` | Convert a quantity in any unit into the base unit for its type |
| `base_unit_of(unit_key)` | `g`, `ml` or `each` for the unit chosen in `unit_key` |
| `min_number(value_range, cond, cap)` | Smallest positive number in `value_range` where `cond` holds |

**Module constants:** `ROW_OFFSET` (HEADER_ROW).

````python
"""Small building blocks shared by the register sheets."""

from __future__ import annotations

from typing import Any, Dict, List

from ..layout import HEADER_ROW, Col
from ..writer import cf_text, yes_no

# Everything below assumes the standard geometry: headers on row 4, data from
# row 5 - ROW()-4 therefore gives a 1 based record number.
ROW_OFFSET = HEADER_ROW


def id_col(trigger: str, prefix: str, digits: int = 4, width: float = 11,
           header: str = "ID", note: str = "") -> Col:
    """Auto numbered ID column (blank until the trigger column is filled).

    `prefix` is a formula fragment - either a defined name such as
    ``RecipePrefix`` or a quoted literal such as ``'"MEAL-"'``.
    """
    zeros = "0" * digits
    return Col(
        "id", header, width, "id",
        formula=(f'=IF({{{trigger}}}{{r}}="","",'
                 f'{prefix}&TEXT(ROW()-{ROW_OFFSET},"{zeros}"))'),
        note=note or "Generated automatically from the row number. "
                     "Filter the list rather than re-sorting it so the IDs "
                     "keep matching your other tabs.")


def month_col(date_key: str, key: str = "month", header: str = "Month",
              width: float = 10) -> Col:
    """`yyyy-mm` text column - the key every monthly roll-up groups on."""
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"yyyy-mm"))'),
               note="Automatic. Every monthly report groups on this.")


def weekday_col(date_key: str, key: str = "weekday", header: str = "Day",
                width: float = 9) -> Col:
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"ddd"))'))


def week_col(date_key: str, key: str = "week", header: str = "Week",
             width: float = 7) -> Col:
    """ISO-ish week number, honouring the week start day on Settings."""
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'WEEKNUM({{{date_key}}}{{r}},'
                        f'IF(WeekStart="Sunday",1,2)))'),
               note="Week number of the year, using the week start day set "
                    "on the Settings tab.")


def alert_cf() -> List[Dict[str, Any]]:
    """Traffic light colouring for the 'Alert' / 'Check' columns."""
    return [
        cf_text("OK", "good", "begins with"),
        cf_text("Overdue", "bad"),
        cf_text("Late", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Conflict", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Duplicate", "bad"),
        cf_text("Action", "warn"),
        cf_text("Due", "warn"),
        cf_text("Check", "warn"),
        cf_text("Add", "warn"),
        cf_text("Today", "info"),
        cf_text("Upcoming", "info"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    """Colouring for the meal, shopping and prep status columns."""
    return [
        cf_text("Prepared", "good"),
        cf_text("Cooked", "good"),
        cf_text("Eaten", "good"),
        cf_text("Bought", "good"),
        cf_text("Purchased", "good"),
        cf_text("Done", "good"),
        cf_text("Complete", "good"),
        cf_text("In stock", "good"),
        cf_text("Yes", "good"),
        cf_text("Planned", "info"),
        cf_text("Scheduled", "info"),
        cf_text("Shopping", "info"),
        cf_text("On the list", "info"),
        cf_text("Prepping", "info"),
        cf_text("Frozen", "info"),
        cf_text("To do", "warn"),
        cf_text("Not started", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Use soon", "warn"),
        cf_text("Defrosting", "warn"),
        cf_text("Substituted", "warn"),
        cf_text("Out of stock", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Skipped", "bad"),
        cf_text("Binned", "bad"),
        cf_text("Wasted", "bad"),
        cf_text("Cancelled", "bad"),
        cf_text("Archived", "muted"),
        cf_text("Not needed", "muted"),
    ]


def rate_cf(target_name: str = "CalorieTarget") -> List[Dict[str, Any]]:
    """Green above the target rate, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def bar_cf(color: str = "#4E8A63") -> List[Dict[str, Any]]:
    return [{"type": "data_bar", "bar_color": color}]


def scale_cf() -> List[Dict[str, Any]]:
    return [{"type": "3_color_scale", "min_color": "#FBD9DD",
             "mid_color": "#FFF1CF", "max_color": "#D9F2E4"}]


def money_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0, "fmt": "bold_good"}]


def dup_cf() -> List[Dict[str, Any]]:
    return [cf_text("Duplicate", "bad"), cf_text("OK", "good")]


def flag_yes_no_cf() -> List[Dict[str, Any]]:
    return [cf_text("No", "warn"), cf_text("Yes", "good")]


def yesno(note: str = "") -> Dict[str, Any]:
    return yes_no(note)


def score_validation(maximum: int = 10) -> Dict[str, Any]:
    return {
        "validate": "integer",
        "criteria": "between",
        "minimum": 1,
        "maximum": maximum,
        "error_title": f"Score 1-{maximum}",
        "error_message": f"Enter a whole number from 1 (low) to {maximum} "
                         f"(high).",
    }


def blank(trigger: str, body: str) -> str:
    """`=IF(trigger="","",body)` with the standard row placeholders."""
    return f'=IF({{{trigger}}}{{r}}="","",{body})'


def lookup_recipe(target: str, key: str = "recipe") -> str:
    """Pull a column of the Recipes tab for the recipe ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Recipes.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Recipes.id]],0)),"Not found"))')


def lookup_ingredient(target: str, key: str = "ingredient") -> str:
    """Pull a column of the Ingredients price book for the name in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Ingredients.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Ingredients.name]],0)),"Add to Ingredients"))')


#: grams / millilitres / each - the base units everything is converted to
def to_base(qty_key: str, unit_key: str) -> str:
    """Convert a quantity in any unit into the base unit for its type.

    The factor comes from the unit table on the Lists tab, so adding a
    unit there is all it takes to support it everywhere.
    """
    return (f'=IF({{{qty_key}}}{{r}}="","",IFERROR({{{qty_key}}}{{r}}*'
            f'INDEX(UnitFactors,MATCH({{{unit_key}}}{{r}},UnitNames,0)),""))')


def base_unit_of(unit_key: str) -> str:
    """`g`, `ml` or `each` for the unit chosen in `unit_key`."""
    return (f'=IF({{{unit_key}}}{{r}}="","",IFERROR('
            f'INDEX(UnitBases,MATCH({{{unit_key}}}{{r}},UnitNames,0)),"?"))')


def min_number(value_range: str, cond: str, cap: int = 999999) -> str:
    """Smallest positive number in `value_range` where `cond` holds.

    Excel 2010 has no MINIFS, and a plain SUMPRODUCT(MIN(...)) trips over
    the empty strings that formula columns leave behind, so the values are
    coerced with IFERROR(...*1,0) and the non matching rows are swapped for
    a big number.  Returns `cap` when nothing matches - test for that.
    """
    safe = f"IFERROR({value_range}*1,0)"
    hit = f"(({cond})*({safe}>0))"
    return f"SUMPRODUCT(MIN({hit}*{safe}+(1-{hit})*{cap}))"
````

### `meal_planner_spreadsheet/sheets/lists.py`

*286 lines · 1 function*

> `Lists` - every drop down in the workbook, plus the unit table.
>
> Each list becomes a workbook level defined name, so a column can point at
> it with ``validate=list_from("MealSlotList")``. Edit a list here and every
> drop down that uses it follows.
>
> The unit table is the engine behind the automatic grocery list: each unit
> carries the factor that converts it to a base unit (grams, millilitres or
> "each"), which is what lets the workbook add 2 tbsp of oil to 100 ml of
> oil and come out with a sensible number.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Lists'), `LAST_COL` (9), `LISTS` (49 items), `UNITS` (30 items).

````python
"""`Lists` - every drop down in the workbook, plus the unit table.

Each list becomes a workbook level defined name, so a column can point at
it with ``validate=list_from("MealSlotList")``. Edit a list here and every
drop down that uses it follows.

The unit table is the engine behind the automatic grocery list: each unit
carries the factor that converts it to a base unit (grams, millilitres or
"each"), which is what lets the workbook add 2 tbsp of oil to 100 ml of
oil and come out with a sensible number.
"""

from __future__ import annotations

from typing import List, Tuple

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Lists"
LAST_COL = 9

#: name -> (header, [values]) - written one list per column block
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("YesNoNAList", ["Yes", "No", "N/A"]),
    ("MealSlotList", ["Breakfast", "Morning snack", "Lunch",
                      "Afternoon snack", "Dinner", "Evening snack",
                      "Dessert"]),
    ("MealStatusList", ["Planned", "Prepared", "Eaten", "Skipped",
                        "Ate out", "Takeaway", "Leftovers", "Cancelled"]),
    ("RecipeCategoryList", ["Main", "Side", "Soup", "Salad", "Breakfast",
                            "Snack", "Dessert", "Baking", "Sauce & dip",
                            "Drink", "Batch base", "Lunchbox"]),
    ("CuisineList", ["American", "British", "Chinese", "French", "Greek",
                     "Indian", "Italian", "Japanese", "Korean", "Mexican",
                     "Middle Eastern", "Spanish", "Thai", "Vietnamese",
                     "Caribbean", "African", "Fusion", "Other"]),
    ("CookingMethodList", ["Oven", "Hob", "One pot", "Slow cooker",
                           "Pressure cooker", "Air fryer", "Grill",
                           "Barbecue", "Microwave", "No cook",
                           "Blender", "Bread machine"]),
    ("DifficultyList", ["Easy", "Medium", "Hard"]),
    ("SpiceLevelList", ["None", "Mild", "Medium", "Hot", "Very hot"]),
    ("PortionSizeList", ["Small", "Regular", "Large", "Extra large"]),
    ("DietList", ["Vegetarian", "Vegan", "Pescatarian", "Gluten-free",
                  "Dairy-free", "Egg-free", "Nut-free", "Soy-free",
                  "Low sodium", "Low carb", "High protein", "High fibre",
                  "Halal", "Kosher", "No restriction"]),
    ("AllergenList", ["Celery", "Cereals with gluten", "Crustaceans",
                      "Eggs", "Fish", "Lupin", "Milk", "Molluscs",
                      "Mustard", "Nuts", "Peanuts", "Sesame", "Soya",
                      "Sulphites", "None"]),
    ("GroceryCategoryList", ["Produce", "Meat & seafood", "Dairy & eggs",
                             "Bakery", "Pantry staples", "Frozen",
                             "Beverages", "Spices & seasonings",
                             "Snacks", "Baby & kids", "Household",
                             "Health & personal care", "Pet", "Other"]),
    ("AisleList", ["1 - Produce", "2 - Bakery", "3 - Deli & cheese",
                   "4 - Meat & fish", "5 - Dairy & eggs", "6 - Frozen",
                   "7 - Tins & jars", "8 - Pasta, rice & grains",
                   "9 - Baking", "10 - Spices & sauces", "11 - Snacks",
                   "12 - Drinks", "13 - Household", "14 - Health",
                   "Online only", "Market stall"]),
    ("StorageList", ["Pantry", "Cupboard", "Fridge", "Freezer",
                     "Fruit bowl", "Spice rack", "Garage shelf",
                     "Cellar", "Other"]),
    ("ContainerList", ["Glass box", "Plastic box", "Freezer bag",
                       "Jar", "Foil tray", "Vacuum bag", "Original pack",
                       "Wrapped plate", "Bowl with lid"]),
    ("ShoppingStatusList", ["To buy", "On the list", "Bought",
                            "Substituted", "Out of stock", "Not needed",
                            "Ordered"]),
    ("PriorityList", ["Must have", "Nice to have", "Only if cheap",
                      "Stock up"]),
    ("TaskStatusList", ["To do", "In progress", "Done", "Blocked",
                        "Not needed"]),
    ("PrepTaskList", ["Wash & chop", "Marinate", "Batch cook", "Bake",
                      "Portion & label", "Freeze", "Defrost", "Soak",
                      "Make sauce", "Pack lunchboxes", "Shop", "Tidy"]),
    ("MemberRoleList", ["Cook", "Shopper", "Helper", "Dishes",
                        "Lunchbox duty", "Guest", "Child"]),
    ("EventTypeList", ["Birthday", "Christmas", "Thanksgiving", "Easter",
                       "Dinner party", "Barbecue", "Picnic", "Potluck",
                       "Sunday lunch", "Anniversary", "School event",
                       "Other"]),
    ("LeftoverStatusList", ["In the fridge", "In the freezer", "Planned",
                            "Eaten", "Frozen for later", "Binned",
                            "Given away"]),
    ("WasteReasonList", ["Expired", "Went mouldy", "Cooked too much",
                         "Did not like it", "Forgot about it",
                         "Freezer burn", "Left out too long",
                         "Bought too much", "Plate scrapings", "Other"]),
    ("WasteActionList", ["Binned", "Composted", "Fed to pets",
                         "Given away", "Donated"]),
    ("StoreTypeList", ["Supermarket", "Discounter", "Local grocer",
                       "Butcher", "Fishmonger", "Bakery", "Farm shop",
                       "Market", "Online delivery", "Wholesale",
                       "Convenience", "Takeaway", "Restaurant"]),
    ("SpendCategoryList", ["Groceries", "Fresh top-up", "Bulk shop",
                           "Takeaway", "Restaurant", "Coffee & drinks",
                           "Work lunch", "School meals", "Snacks",
                           "Subscription box", "Delivery fee",
                           "Household"]),
    ("GoalCategoryList", ["Budget", "Health", "Waste", "Habit",
                          "Variety", "Skill", "Time"]),
    ("GoalStatusList", ["Not started", "In progress", "On track",
                        "Behind", "Complete", "Paused"]),
    ("FrequencyList", ["Daily", "Weekly", "Fortnightly", "Monthly",
                       "Occasionally", "One-off"]),
    ("SeasonList", ["All year", "Spring", "Summer", "Autumn", "Winter"]),
    ("RatingList", ["5 - Make it again", "4 - Good", "3 - Fine",
                    "2 - Needs work", "1 - Never again"]),
    ("TextureList", ["Any", "Soft", "Crunchy", "Creamy", "Crispy",
                     "Smooth", "No lumps", "Not slimy"]),
    ("MonthList", ["January", "February", "March", "April", "May",
                   "June", "July", "August", "September", "October",
                   "November", "December"]),
    ("WeekdayList", ["Monday", "Tuesday", "Wednesday", "Thursday",
                     "Friday", "Saturday", "Sunday"]),
    ("WeekStartList", ["Monday", "Sunday"]),
    ("SystemList", ["Metric", "Imperial"]),
    ("RoundingList", ["Whole packs", "Exact quantity"]),
    ("ActiveList", ["Active", "Occasional", "Away", "Inactive"]),
    ("AgeGroupList", ["Adult", "Teenager", "Child", "Toddler", "Baby",
                      "Senior"]),
    ("BudgetLineList", ["Weekly shop", "Top-up shop", "Eating out",
                        "Takeaway", "Coffee & treats", "Lunches out",
                        "Special occasion", "Household", "Pet food"]),
    ("SpendTypeList", ["Weekly shop", "Top-up shop", "Bulk shop",
                       "Takeaway", "Eating out", "Work lunch",
                       "Coffee & treats", "Delivery fee",
                       "Special occasion", "Household"]),
    ("PaymentList", ["Card", "Cash", "Online", "Loyalty points",
                     "Voucher", "Shared account"]),
    ("PriceLevelList", ["Budget", "Mid-range", "Premium", "Varies"]),
    ("PrepStatusList", ["To do", "Prepping", "Done", "Skipped"]),
    ("EventStatusList", ["Idea", "Planned", "Menu set", "Shopping done",
                         "Done", "Cancelled"]),
    ("DirectionList", ["Hit or beat", "Stay under"]),
    ("DisposalList", ["General bin", "Food waste caddy", "Compost",
                      "Fed to pets", "Given away"]),
    ("MeasureList", ["Home cooked meals a week", "Takeaways a month",
                     "Grocery spend this month", "New recipes this month",
                     "Food waste this month", "Packed lunches a week",
                     "Leftovers eaten", "Meals prepped ahead",
                     "Your own count"]),
]

#: unit -> (base unit, how many base units one of it is, what it measures)
#: The factors are the ordinary cooking conversions; the "each" units are
#: counted rather than measured, so they convert to themselves.
UNITS: List[Tuple[str, str, float, str]] = [
    ("g", "g", 1, "Weight"),
    ("kg", "g", 1000, "Weight"),
    ("mg", "g", 0.001, "Weight"),
    ("oz", "g", 28.3495, "Weight"),
    ("lb", "g", 453.592, "Weight"),
    ("ml", "ml", 1, "Volume"),
    ("l", "ml", 1000, "Volume"),
    ("tsp", "ml", 5, "Volume"),
    ("tbsp", "ml", 15, "Volume"),
    ("cup", "ml", 240, "Volume"),
    ("fl oz", "ml", 29.5735, "Volume"),
    ("pint (US)", "ml", 473.176, "Volume"),
    ("pint (UK)", "ml", 568.261, "Volume"),
    ("quart", "ml", 946.353, "Volume"),
    ("gallon (US)", "ml", 3785.41, "Volume"),
    ("each", "each", 1, "Count"),
    ("clove", "each", 1, "Count"),
    ("slice", "each", 1, "Count"),
    ("pinch", "each", 1, "Count"),
    ("handful", "each", 1, "Count"),
    ("bunch", "each", 1, "Count"),
    ("sprig", "each", 1, "Count"),
    ("can", "each", 1, "Count"),
    ("jar", "each", 1, "Count"),
    ("packet", "each", 1, "Count"),
    ("bag", "each", 1, "Count"),
    ("bottle", "each", 1, "Count"),
    ("loaf", "each", 1, "Count"),
    ("sheet", "each", 1, "Count"),
    ("to taste", "each", 0, "Count"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.grey_text)

    s.banner("Drop Down Lists",
             "Every list the workbook offers you. Edit, add or delete an "
             "entry and the drop down that uses it changes everywhere.",
             LAST_COL)

    row = 4
    row = s.note(
        row,
        "Add new entries directly under the last one in a column - the "
        "lists are sized with room to grow, and blanks are ignored. Do "
        "not delete a whole column: a drop down pointing at a missing "
        "list stops working.",
        LAST_COL, height=30)
    row += 1

    # ------------------------------------------------------------ the lists
    row = s.section(row, "The drop downs", LAST_COL)
    per_row = 5
    spare = 8                       # blank rows kept inside every name
    top = row
    col = 0
    tallest = 0
    for index, (name, values) in enumerate(LISTS):
        if index and index % per_row == 0:
            top += tallest + 2
            col = 0
            tallest = 0
        s.write(top, col, name.replace("List", ""), fmts.chrome("header"))
        s.set_col(col, col, 22)
        for offset, value in enumerate(values):
            s.write(top + 1 + offset, col, value, fmts.input("text"))
        for offset in range(len(values), len(values) + spare):
            s.write(top + 1 + offset, col, None, fmts.input("text"))
        first = top + 2             # 1 based first value row
        last = top + 1 + len(values) + spare
        book.define_name(
            name,
            f"={quote_sheet(NAME)}!${chr(65 + col)}${first}:"
            f"${chr(65 + col)}${last}")
        tallest = max(tallest, len(values) + spare)
        col += 2
    row = top + tallest + 3

    # -------------------------------------------------------- the unit table
    row = s.section(row, "Measurement units and how they convert", LAST_COL)
    row = s.note(
        row,
        "This little table is what makes the grocery list add up. Every "
        "quantity you type on a recipe is multiplied by the factor below "
        "to get grams, millilitres or a simple count, so two tablespoons "
        "of oil and 100 ml of oil end up on one line of the shopping "
        "list. Add a unit here and it works everywhere.",
        LAST_COL, height=42)
    row += 1

    s.table_head(row, 0, ["Unit", "Measures", "Base unit",
                          "1 unit = this many base units"],
                 [14, 12, 12, 30])
    first_unit = row + 1
    for index, (unit, base, factor, kind) in enumerate(UNITS):
        r = first_unit + index
        s.write(r, 0, unit, fmts.input("text"))
        s.write(r, 1, kind, fmts.input("text"))
        s.write(r, 2, base, fmts.input("text"))
        s.write(r, 3, factor, fmts.input("number1"))
    spare_units = 10
    for index in range(len(UNITS), len(UNITS) + spare_units):
        r = first_unit + index
        for c, style in enumerate(["text", "text", "text", "number1"]):
            s.write(r, c, None, fmts.input(style))
    last_unit = first_unit + len(UNITS) + spare_units

    sheet_ref = quote_sheet(NAME)
    book.define_name("UnitNames",
                     f"={sheet_ref}!$A${first_unit + 1}:$A${last_unit}")
    book.define_name("UnitKinds",
                     f"={sheet_ref}!$B${first_unit + 1}:$B${last_unit}")
    book.define_name("UnitBases",
                     f"={sheet_ref}!$C${first_unit + 1}:$C${last_unit}")
    book.define_name("UnitFactors",
                     f"={sheet_ref}!$D${first_unit + 1}:$D${last_unit}")
    book.define_name("UnitList",
                     f"={sheet_ref}!$A${first_unit + 1}:$A${last_unit}")
    row = last_unit + 2

    row = s.note(
        row,
        "A note on cups and grams: converting a volume into a weight "
        "depends on what is in the cup (a cup of flour and a cup of "
        "honey weigh very different amounts). The workbook keeps weights "
        "and volumes apart rather than guessing - if you buy an "
        "ingredient by weight, write its recipe quantities by weight too.",
        LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/settings.py`

*239 lines · 1 function*

> The `Settings` tab - household setup, planning window, budget and targets.
>
> Everything the rest of the workbook needs to know about your household is
> entered once here and published as a defined name (HouseName, PlanStart,
> WeeklyBudget, CalorieTarget, ExpiryWarningDays ...) so the formulas stay
> readable.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `SHEET` ('Settings'), `LAST_COL` (8).

````python
"""
The `Settings` tab - household setup, planning window, budget and targets.

Everything the rest of the workbook needs to know about your household is
entered once here and published as a defined name (HouseName, PlanStart,
WeeklyBudget, CalorieTarget, ExpiryWarningDays ...) so the formulas stay
readable.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, date_check, list_from, number_check

SHEET = "Settings"
LAST_COL = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, SHEET, tab_color=Palette.accent)
    h = cfg.household

    s.banner("Household Setup & Settings",
             "Fill in the cream cells once. Grey cells are locked "
             "formulas. Everything else in the workbook reads these "
             "values.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 30, fmts.chrome("label"))
    s.set_col(2, 2, 16, fmts.chrome("label"))
    s.set_col(3, 3, 26)
    s.set_col(4, 4, 20)
    s.set_col(5, 8, 17, fmts.chrome("note"))

    r = 4  # 0 based row index -> Excel row 5

    # -------------------------------------------------------- 1. household
    r = s.section(r, "1. Your household", LAST_COL)
    r = s.field(r, "Household name", h.name, name="HouseName", value_end=4,
                help_text="Shown on every printable page and report.")
    r = s.field(r, "Main cook / planner", h.cook, name="HouseCook",
                value_end=4)
    r = s.field(r, "Email", h.email, name="HouseEmail", value_end=4)
    r = s.field(r, "Town / city", h.town, name="HouseTown", value_end=4)
    r = s.field(r, "People to cook for", cfg.members, style="int",
                name="HouseMembers",
                validate=number_check(1, 30, decimal=False),
                help_text="Used as the default serving count and to work "
                          "out the cost per person.")
    r = s.field(r, "Usual supermarket", h.store, name="MainStore",
                value_end=4,
                help_text="The store the grocery list defaults to.")
    r = s.field(r, "Notes about the household", h.notes, name="HouseNotes",
                value_end=4)

    # --------------------------------------------------------- 2. planning
    r += 1
    r = s.section(r, "2. Planning window", LAST_COL)
    r = s.field(r, "Planning year", cfg.year, style="int", name="PlanYear",
                validate=number_check(2000, 2100, decimal=False),
                help_text="The year the monthly planner and the reports "
                          "open on.")
    r = s.field(r, "Meal planning starts on", cfg.start_date, style="date",
                name="PlanStart", validate=date_check(),
                help_text="The first day of week 1 on the weekly planner. "
                          "Usually a Monday.")
    r = s.field(r, "Week starts on", cfg.week_start, name="WeekStart",
                validate=list_from("WeekStartList"),
                help_text="Used by the planner grids and the week numbers.")
    r = s.field(r, "Shopping day", cfg.shopping_day, name="ShoppingDay",
                validate=list_from("WeekdayList"),
                help_text="The grocery list counts down to this day.")
    r = s.field(r, "Meal prep day", cfg.prep_day, name="PrepDay",
                validate=list_from("WeekdayList"))
    r = s.field(r, "Measurement system", cfg.measurement_system,
                name="MeasureSystem", validate=list_from("SystemList"),
                help_text="Metric shows grams and millilitres, Imperial "
                          "shows ounces and cups. Both always work - this "
                          "only changes what the converter leads with.")
    r = s.field(r, "Default servings per meal", cfg.default_servings,
                style="int", name="DefaultServings",
                validate=number_check(1, 50, decimal=False))
    r = s.field(r, "Shopping list covers from", "=PlanStart", style="date",
                name="ShopFrom", validate=date_check(),
                help_text="The grocery list adds up every meal planned "
                          "between these two dates. Type over the dates "
                          "whenever you want a different window - the "
                          "whole list rebuilds itself.")
    r = s.field(r, "Shopping list covers to", "=ShopFrom+6", style="date",
                name="ShopTo", validate=date_check(),
                help_text="Six days after the start gives you one week. "
                          "Make it +13 for a fortnight's shop.")
    r = s.field(r, "Rounding on the grocery list", cfg.rounding,
                name="RoundingRule", value_end=4,
                validate=list_from("RoundingList"),
                help_text="'Whole packs' turns 1.3 tins into 2 tins - "
                          "what the shop actually sells you. 'Exact "
                          "quantity' leaves the raw number.")

    # ----------------------------------------------------------- 3. budget
    r += 1
    r = s.section(r, "3. Food budget", LAST_COL)
    r = s.field(r, "Weekly food budget", cfg.weekly_food_budget,
                style="money", name="WeeklyBudget",
                validate=number_check(0, 100000),
                help_text="Groceries only - eating out has its own line.")
    r = s.field(r, "Monthly food budget", cfg.monthly_food_budget,
                style="money", name="MonthlyBudget",
                validate=number_check(0, 1000000))
    r = s.field(r, "Yearly food budget", cfg.annual_food_budget,
                style="money", name="AnnualBudget",
                validate=number_check(0, 10000000))
    r = s.field(r, "Eating out / takeaway budget (month)",
                cfg.eating_out_budget, style="money", name="EatingOutBudget",
                validate=number_check(0, 100000))
    r = s.field(r, "Food waste cost you will accept (month)",
                cfg.waste_cost_target, style="money", name="WasteCostTarget",
                validate=number_check(0, 100000),
                help_text="The waste dashboard turns red above this.")

    # -------------------------------------------------------- 4. nutrition
    r += 1
    r = s.section(r, "4. Daily nutrition targets (per person)", LAST_COL)
    r = s.note(r, "Optional. Leave them as they are if you only want the "
                  "planning and the shopping - nothing else depends on "
                  "them. These are general placeholders, not medical "
                  "advice; set your own from your own guidance.",
               LAST_COL, first_col=1, height=30)
    r = s.field(r, "Calories a day", cfg.calorie_target, style="int",
                name="CalorieTarget", validate=number_check(0, 20000,
                                                            decimal=False))
    r = s.field(r, "Protein (g)", cfg.protein_target, style="int",
                name="ProteinTarget",
                validate=number_check(0, 1000, decimal=False))
    r = s.field(r, "Carbohydrates (g)", cfg.carb_target, style="int",
                name="CarbTarget",
                validate=number_check(0, 2000, decimal=False))
    r = s.field(r, "Fat (g)", cfg.fat_target, style="int", name="FatTarget",
                validate=number_check(0, 1000, decimal=False))
    r = s.field(r, "Fibre (g)", cfg.fibre_target, style="int",
                name="FibreTarget",
                validate=number_check(0, 500, decimal=False))
    r = s.field(r, "Sugar limit (g)", cfg.sugar_limit, style="int",
                name="SugarLimit",
                validate=number_check(0, 1000, decimal=False))
    r = s.field(r, "Sodium limit (mg)", cfg.sodium_limit, style="int",
                name="SodiumLimit",
                validate=number_check(0, 20000, decimal=False))
    r = s.field(r, "Water target (glasses a day)", cfg.water_target,
                style="int", name="WaterTarget",
                validate=number_check(0, 40, decimal=False))

    # --------------------------------------------------------- 5. reminders
    r += 1
    r = s.section(r, "5. Reminder windows", LAST_COL)
    r = s.field(r, "Warn me this many days before food expires",
                cfg.expiry_warning_days, style="int",
                name="ExpiryWarningDays",
                validate=number_check(0, 120, decimal=False),
                help_text="Drives the 'use it soon' flags on the fridge, "
                          "freezer and pantry tabs.")
    r = s.field(r, "Days a cooked leftover keeps in the fridge",
                cfg.leftover_safe_days, style="int", name="LeftoverDays",
                validate=number_check(0, 30, decimal=False),
                help_text="A guide for the leftovers tab, not a food "
                          "safety ruling - go by how it looks and smells, "
                          "and by local guidance.")
    r = s.field(r, "Days before a freezer item is 'old'",
                cfg.freezer_review_days, style="int", name="FreezerDays",
                validate=number_check(0, 1000, decimal=False))
    r = s.field(r, "Days of stock that counts as low", cfg.low_stock_days,
                style="int", name="LowStockDays",
                validate=number_check(0, 365, decimal=False))
    r = s.field(r, "Flag prep tasks this many days early",
                cfg.prep_reminder_days, style="int", name="PrepLeadDays",
                validate=number_check(0, 30, decimal=False))
    r = s.field(r, "Days before a meal can repeat", cfg.rotation_gap_days,
                style="int", name="RotationGapDays",
                validate=number_check(0, 365, decimal=False),
                help_text="The rotation planner flags anything cooked "
                          "again sooner than this.")

    # ----------------------------------------------------------- 6. habits
    r += 1
    r = s.section(r, "6. Habit targets", LAST_COL)
    r = s.field(r, "Home cooked meals a week", cfg.home_cooked_target,
                style="int", name="HomeCookedTarget",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "New recipes a month", cfg.new_recipe_target,
                style="int", name="NewRecipeTarget",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "Takeaways a month (limit)", cfg.takeaway_limit,
                style="int", name="TakeawayLimit",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "Packed lunches a week", cfg.packed_lunch_target,
                style="int", name="PackedLunchTarget",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "Food waste target (share of spend)",
                cfg.waste_reduction_target, style="percent",
                name="WasteTarget", validate=number_check(0, 1))

    # --------------------------------------------------------- 7. prefixes
    r += 1
    r = s.section(r, "7. ID prefixes", LAST_COL)
    r = s.note(r, "The letters in front of the automatic numbers. Change "
                  "them before you start entering data, not after.",
               LAST_COL, first_col=1, height=24)
    for label, value, name in [
            ("Recipe IDs", "R-", "RecipePrefix"),
            ("Ingredient IDs", "ING-", "IngredientPrefix"),
            ("Planned meal IDs", "M-", "MealPrefix"),
            ("Pantry item IDs", "P-", "PantryPrefix"),
            ("Shopping line IDs", "G-", "GroceryPrefix"),
            ("Leftover IDs", "L-", "LeftoverPrefix"),
            ("Event IDs", "E-", "EventPrefix"),
            ("Purchase IDs", "S-", "SpendPrefix")]:
        r = s.field(r, label, value, name=name)

    # ----------------------------------------------------------- 8. product
    r += 1
    r = s.section(r, "8. About this workbook", LAST_COL)
    r = s.field(r, "Product", cfg.product_name, mode="calc", value_end=5)
    r = s.field(r, "Version", cfg.version, mode="calc")
    r = s.field(r, "Created by", cfg.author, mode="calc")
    r = s.field(r, "Sheet protection password", cfg.password, mode="calc",
                help_text="Review > Unprotect Sheet if you need to change "
                          "a formula.")
    r = s.field(r, "Currency", f"{cfg.currency_symbol} "
                               f"({cfg.currency_code})", mode="calc")
    r += 1
    s.note(r,
           "Changed your mind about something here? Everything updates "
           "straight away - the planner, the grocery list, the budget and "
           "the dashboard all read these cells rather than keeping their "
           "own copy.", LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/members.py`

*132 lines · 2 functions*

> `Household` - who you are cooking for.
>
> One row per person: what they will not eat, what they must not eat, how
> hungry they are and which meals they are actually in for. The planner
> uses the portion factor to work out how much food a meal really needs,
> and the allergy column drives the warnings on the recipes.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Household').

````python
"""`Household` - who you are cooking for.

One row per person: what they will not eat, what they must not eat, how
hungry they are and which meals they are actually in for. The planner
uses the portion factor to work out how much food a meal really needs,
and the allergy column drives the warnings on the recipes.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, status_cf

NAME = "Household"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"HH-"', digits=2, width=8),
        Col("name", "Name", 18, "text",
            note="Everyone who eats here, including the ones who only "
                 "turn up at weekends."),
        Col("role", "Role", 14, "text", validate=list_from("MemberRoleList")),
        Col("age_group", "Age group", 13, "text",
            validate=list_from("AgeGroupList"),
            note="Only used to suggest a portion size."),
        Col("portion", "Portion factor", 11, "number1",
            validate=number_check(0.25, 3),
            note="1 is an adult portion. A small child might be 0.5, a "
                 "teenager 1.5. The planner multiplies by this."),
        Col("diet", "Diet", 15, "text", validate=list_from("DietList")),
        Col("allergies", "Allergies", 20, "text",
            validate=list_from("AllergenList"),
            note="The serious ones. Recipes containing them are "
                 "flagged - but always read the label yourself."),
        Col("intolerances", "Avoids", 20, "text",
            note="Things they would rather not eat: 'mushrooms', "
                 "'anything spicy'."),
        Col("dislikes", "Will not touch", 22, "text"),
        Col("loves", "Favourites", 22, "text",
            note="Worth filling in - it makes the 'what shall we have?' "
                 "question much easier."),
        Col("breakfast", "In for breakfast?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("lunch", "In for lunch?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("dinner", "In for dinner?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("packed", "Takes a packed lunch?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("cal_target", "Daily calories", 11, "number",
            validate=number_check(0, 6000),
            note="Leave blank to use the household target on the "
                 "Settings tab."),
        Col("eff_cals", "Calories used", 11, "number",
            formula='=IF({name}{r}="","",IF({cal_target}{r}="",'
                    'CalorieTarget*IF({portion}{r}="",1,{portion}{r}),'
                    '{cal_target}{r}))', hidden=True),
        Col("cooks", "Cooks?", 9, "flag", validate=list_from("YesNoList")),
        Col("cook_days", "Cooking nights", 16, "text",
            note="'Tuesday and Thursday'. The weekly planner shows it "
                 "so nobody can claim they forgot."),
        Col("shops", "Does the shopping?", 10, "flag",
            validate=list_from("YesNoList")),
        Col("jobs", "Kitchen jobs", 20, "text",
            note="Washing up, bins, packing lunches."),
        Col("meals_planned", "Meals planned for them", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Meal Plan.cook]],{name}{r}))',
            note="How many meals this person is down to cook."),
        Col("birthday", "Birthday", 11, "date", validate=date_check(),
            note="So the special events tab can remind you."),
        Col("status", "Status", 12, "status",
            validate=list_from("ActiveList"), cf=status_cf()),
        Col("alert", "Check", 24, "status",
            formula='=IF({name}{r}="","",IF({portion}{r}="",'
                    '"Add - no portion factor",IF(AND({allergies}{r}<>"",'
                    '{allergies}{r}<>"None"),"Check - allergy on file",'
                    '"OK")))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Household & Preferences",
        subtitle="Who eats here, what they can and cannot eat, and how "
                 "much of it - the numbers behind every portion in the "
                 "workbook.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).members,
        help_text="Fill this in once. The portion factors add up to the "
                  "household size the planner uses when you do not "
                  "type a number of servings.",
        extra_notes=[
            "Portion factor is the honest way to plan for a family "
            "with a toddler and a teenager: two adults, a 0.5 and a "
            "1.5 is 4 portions, not 4 identical plates.",
            "Allergy flags are an aid, not a medical check. Recipes "
            "are flagged when an ingredient carries the same allergen "
            "tag - always read the actual packaging.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "The household in numbers", last)
    row = s.table_head(row, 0, ["People", "Total portions",
                                "Daily calories for the house",
                                "Allergies on file", "Packed lunches",
                                "Cooks in the house"],
                       [10, 14, 24, 16, 15, 17])
    row = s.calc_row(row, 0, [
        '=COUNTIFS([[Household.name]],"<>",[[Household.status]],'
        '"Active")',
        '=SUMIFS([[Household.portion]],[[Household.status]],"Active")',
        '=SUMIFS([[Household.eff_cals]],[[Household.status]],"Active")',
        '=COUNTIFS([[Household.allergies]],"<>",'
        '[[Household.allergies]],"<>None")',
        '=COUNTIF([[Household.packed]],"Yes")',
        '=COUNTIF([[Household.cooks]],"Yes")',
    ], ["int", "number1", "number", "int", "int", "int"])
````

### `meal_planner_spreadsheet/sheets/recipes.py`

*294 lines · 1 function*

> `Recipes` - the recipe database.
>
> One row per recipe: what it is, how long it takes, who it suits and how
> it scored at the table. The cost and the nutrition are **not** typed in -
> they are added up from the Recipe Ingredients tab, so changing the price
> of butter re-prices every recipe that uses it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Recipes').

````python
"""`Recipes` - the recipe database.

One row per recipe: what it is, how long it takes, who it suits and how
it scored at the table. The cost and the nutrition are **not** typed in -
they are added up from the Recipe Ingredients tab, so changing the price
of butter re-prices every recipe that uses it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, score_validation

NAME = "Recipes"

#: the Recipe Ingredients columns the roll-ups sum
_LINES = "[[Recipe Ingredients.recipe]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "RecipePrefix", digits=4, width=10,
               note="The ID the meal planner, the shopping list and the "
                    "cost tabs all refer to. Generated for you."),
        Col("name", "Recipe", 30, "text",
            note="One row per recipe. The ingredients go on the Recipe "
                 "Ingredients tab, one line each."),
        Col("category", "Category", 14, "text",
            validate=list_from("RecipeCategoryList")),
        Col("meal_slot", "Usual meal", 14, "text",
            validate=list_from("MealSlotList"),
            note="Which meal of the day it normally turns up at."),
        Col("cuisine", "Cuisine", 14, "text",
            validate=list_from("CuisineList")),
        Col("description", "What it is", 34, "wrap"),
        Col("servings", "Serves", 8, "int",
            validate=number_check(1, 100, decimal=False),
            note="How many people the quantities on the Recipe "
                 "Ingredients tab feed. Everything scales from this."),
        Col("prep_min", "Prep (min)", 10, "int",
            validate=number_check(0, 600, decimal=False)),
        Col("cook_min", "Cook (min)", 10, "int",
            validate=number_check(0, 1440, decimal=False)),
        Col("total_min", "Total (min)", 11, "int",
            formula='=IF({name}{r}="","",IF(AND({prep_min}{r}="",'
                    '{cook_min}{r}=""),"",'
                    'IF({prep_min}{r}="",0,{prep_min}{r})+'
                    'IF({cook_min}{r}="",0,{cook_min}{r})))',
            cf=[{"type": "cell", "criteria": "<=", "value": 30,
                 "fmt": "good"},
                {"type": "cell", "criteria": ">", "value": 90,
                 "fmt": "warn"}]),
        Col("quick", "Quick meal?", 11, "flag",
            formula='=IF({total_min}{r}="","",'
                    'IF({total_min}{r}<=30,"Yes","No"))',
            note="30 minutes or less, start to finish."),
        Col("method", "Method", 14, "text",
            validate=list_from("CookingMethodList")),
        Col("difficulty", "Difficulty", 11, "text",
            validate=list_from("DifficultyList")),
        Col("equipment", "Equipment", 20, "text",
            note="Anything you need out of the cupboard: blender, "
                 "slow cooker, baking tin."),
        Col("one_pot", "One pot?", 9, "flag", validate=list_from(
            "YesNoList")),
        Col("no_cook", "No cook?", 9, "flag", validate=list_from(
            "YesNoList")),
        Col("prep_friendly", "Batch / prep friendly?", 13, "flag",
            validate=list_from("YesNoList"),
            note="Does it survive being made in advance and reheated?"),
        Col("freezes", "Freezes well?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("kid_friendly", "Kid friendly?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("diet", "Suits", 15, "text", validate=list_from("DietList"),
            note="The dietary filters use this plus the allergens "
                 "worked out from the ingredients."),
        Col("spice", "Spice level", 11, "text",
            validate=list_from("SpiceLevelList")),
        Col("season", "Best season", 12, "text",
            validate=list_from("SeasonList")),
        Col("tags", "Tags", 20, "text",
            note="Free text: 'payday treat', 'uses up leftovers', "
                 "'lunchbox'. Filter on it."),
        Col("source", "Where it came from", 20, "text"),
        Col("link", "Link", 24, "link",
            note="Paste the URL, or the page number of the book."),
        Col("photo", "Photo link", 18, "link",
            note="A spreadsheet cannot keep the photo itself tidy - "
                 "paste a cloud link instead."),
        Col("method_notes", "How to make it", 46, "wrap",
            note="Short numbered steps. 1. Chop... 2. Fry... Keep it to "
                 "what you would scribble on a card."),
        # ------------------------------------------------ rolled up totals
        Col("line_count", "Ingredients", 11, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _LINES + ',{id}{r}))',
            note="Counted from the Recipe Ingredients tab."),
        Col("cost", "Cost to make", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS('
                    '[[Recipe Ingredients.line_cost]],'
                    + _LINES + ',{id}{r}))',
            note="Added up from the ingredient prices. Change a price "
                 "once and every recipe re-prices itself."),
        Col("cost_serving", "Cost per serving", 13, "money",
            formula='=IF(OR({cost}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",{cost}{r}/{servings}{r}))',
            cf=[{"type": "3_color_scale", "min_color": "#D9F2E4",
                 "mid_color": "#FFF1CF", "max_color": "#FBD9DD"}]),
        Col("calories", "Calories a serving", 13, "number",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_cals]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("protein", "Protein (g)", 11, "number1",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_protein]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("carbs", "Carbs (g)", 10, "number1",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_carbs]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("fat", "Fat (g)", 9, "number1",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_fat]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("fibre", "Fibre (g)", 10, "number1",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_fibre]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("allergens", "Allergens in it", 22, "text",
            formula='=IF({id}{r}="","",IF(COUNTIFS(' + _LINES + ',{id}{r},'
                    '[[Recipe Ingredients.allergen]],"<>",'
                    '[[Recipe Ingredients.allergen]],"<>None")=0,'
                    '"None listed",COUNTIFS(' + _LINES + ',{id}{r},'
                    '[[Recipe Ingredients.allergen]],"<>",'
                    '[[Recipe Ingredients.allergen]],"<>None")&'
                    '" ingredient(s) flagged"))',
            note="Counted from the allergen column of each ingredient. "
                 "Always read the actual labels as well.",
            cf=[{"type": "text", "criteria": "containing",
                 "value": "flagged", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "None listed", "fmt": "good"}]),
        Col("cost_rank", "Cost rank", 9, "int",
            formula='=IF(OR({cost_serving}{r}="",NOT(ISNUMBER('
                    '{cost_serving}{r}))),"",COUNTIFS({cost_serving_first}:'
                    '{cost_serving_last},"<"&{cost_serving}{r})+'
                    'COUNTIFS({cost_serving_first}:{cost_serving}{r},'
                    '{cost_serving}{r}))',
            hidden=True,
            note="1 = cheapest a serving. Used by the Recipe Costs "
                 "page."),
        Col("match", "Matches the filter", 9, "int",
            formula='=IF({name}{r}="",0,IF(AND('
                    'OR(FilterDiet="Any",{diet}{r}=FilterDiet),'
                    'OR(FilterCategory="Any",{category}{r}='
                    'FilterCategory),'
                    'OR(FilterMethod="Any",{method}{r}=FilterMethod),'
                    'OR(FilterTime=0,AND({total_min}{r}<>"",'
                    '{total_min}{r}<=FilterTime)),'
                    'OR(FilterCost=0,AND({cost_serving}{r}<>"",'
                    'ISNUMBER({cost_serving}{r}),'
                    '{cost_serving}{r}<=FilterCost)),'
                    'OR(FilterAllergen="Any",COUNTIFS('
                    '[[Recipe Ingredients.recipe]],{id}{r},'
                    '[[Recipe Ingredients.allergen]],FilterAllergen)=0)'
                    '),1,0))',
            hidden=True,
            note="1 when the recipe passes every control on the "
                 "Dietary Filters page."),
        Col("match_rank", "Filter rank", 9, "int",
            formula='=IF({match}{r}=0,"",SUM({match_first}:{match}{r}))',
            hidden=True),
        Col("times_planned_rank", "Popularity rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({times_planned_first}:{times_planned_last},">"&{times_planned}{r})+'
                    'COUNTIFS({times_planned_first}:{times_planned}{r},{times_planned}{r}))',
            hidden=True, note="1 = planned most often."),
        Col("calories_rank", "Calorie rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({calories_first}:{calories_last},">"&{calories}{r})+'
                    'COUNTIFS({calories_first}:{calories}{r},{calories}{r}))',
            hidden=True, note="1 = most calories a serving."),
        Col("protein_rank", "Protein rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({protein_first}:{protein_last},">"&{protein}{r})+'
                    'COUNTIFS({protein_first}:{protein}{r},{protein}{r}))',
            hidden=True, note="1 = most protein a serving."),
        Col("fibre_rank", "Fibre rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({fibre_first}:{fibre_last},">"&{fibre}{r})+'
                    'COUNTIFS({fibre_first}:{fibre}{r},{fibre}{r}))',
            hidden=True, note="1 = most fibre a serving."),
        Col("rotation_rank", "Rotation rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({sort_key_first}:'
                    '{sort_key_last},">"&{sort_key}{r})+'
                    'COUNTIFS({sort_key_first}:{sort_key}{r},'
                    '{sort_key}{r}))',
            hidden=True,
            note="1 = the recipe you have left alone the longest."),
        Col("sort_key", "Sort key", 9, "number1",
            formula='=IF({name}{r}="","",IF({last_cooked}{r}="",99999,'
                    '{days_since}{r}))', hidden=True),
        Col("plan_scale", "Batches in the shopping window", 12, "number1",
            formula='=IF({id}{r}="",0,SUMIFS([[Meal Plan.scale]],'
                    '[[Meal Plan.recipe]],{id}{r},[[Meal Plan.date]],'
                    '">="&ShopFrom,[[Meal Plan.date]],"<="&ShopTo,'
                    '[[Meal Plan.shop]],"Yes"))',
            hidden=True,
            note="How many recipes-worth the grocery list has to buy "
                 "for. Driven by the dates on the Grocery List tab."),
        Col("plan_cost", "Cost in the window", 12, "money",
            formula='=IF(OR({id}{r}="",{cost}{r}=""),"",'
                    '{cost}{r}*{plan_scale}{r})',
            hidden=True),
        Col("missing_price", "Missing a price", 12, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _LINES + ',{id}{r},'
                    '[[Recipe Ingredients.line_cost]],""))',
            hidden=True),
        # ------------------------------------------------------- the family
        Col("rating", "Your rating", 11, "rating",
            validate=score_validation(5),
            note="1 to 5. The rotation planner leans on this."),
        Col("family_score", "Family score", 11, "rating",
            validate=score_validation(5),
            note="What everyone else thought."),
        Col("favourite", "Favourite?", 10, "flag",
            validate=list_from("YesNoList")),
        Col("times_planned", "Times planned", 12, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Meal Plan.recipe]],'
                    '{id}{r}))'),
        Col("times_cooked", "Times cooked", 12, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Meal Plan.recipe]],'
                    '{id}{r},[[Meal Plan.status]],"Prepared")+'
                    'COUNTIFS([[Meal Plan.recipe]],{id}{r},'
                    '[[Meal Plan.status]],"Eaten"))'),
        Col("last_cooked", "Last cooked", 12, "date",
            formula='=IF({id}{r}="","",IF(COUNTIFS([[Meal Plan.recipe]],'
                    '{id}{r},[[Meal Plan.date]],"<="&TODAY())=0,"",'
                    'SUMPRODUCT(MAX(([[Meal Plan.recipe]]={id}{r})*'
                    '([[Meal Plan.date]]<=TODAY())*'
                    '[[Meal Plan.date]]))))',
            note="Automatic, from the meal plan."),
        Col("days_since", "Days since", 10, "int",
            formula='=IF({last_cooked}{r}="","",TODAY()-'
                    '{last_cooked}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=RotationGapDays",
                 "fmt": "warn"}]),
        Col("notes", "Tweaks & notes", 28, "wrap",
            note="'Double the garlic', 'kids pick out the peppers', "
                 "'needs 10 minutes longer in our oven'."),
        Col("added", "Added on", 11, "date", validate=date_check()),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "name",
                'IF({servings}{r}="","Add - how many does it serve?",'
                'IF({line_count}{r}=0,"Add - no ingredients listed yet",'
                'IF({missing_price}{r}>0,"Check - '
                '"&{missing_price}{r}&" ingredient(s) have no price",'
                'IF(AND({days_since}{r}<>"",'
                '{days_since}{r}<RotationGapDays),'
                '"Cooked recently","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Recipe Database",
        subtitle="Every recipe you cook, with its cost and nutrition "
                 "worked out from the ingredients rather than typed in.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).recipes,
        help_text="Add the recipe here first, then list its ingredients "
                  "on the Recipe Ingredients tab using the ID in column "
                  "A. That is what makes the grocery list automatic.",
        extra_notes=[
            "Serves matters more than anything else on this tab: the "
            "shopping list multiplies the ingredients by (people you "
            "are feeding / serves), so get it right and everything "
            "else follows.",
            "Cost, calories and allergens are grey because they are "
            "added up from the ingredient lines - there is nothing to "
            "type.",
            "Filter the Suits, Quick meal and Freezes columns to answer "
            "'what can I make in 20 minutes that the vegetarian will "
            "eat?'",
        ],
    )
````

### `meal_planner_spreadsheet/sheets/recipe_lines.py`

*195 lines · 2 functions*

> `Recipe Ingredients` - one line per ingredient in a recipe.
>
> This is the engine room. Every line says *recipe R-0007 needs 300 g of
> chicken breast*; the workbook converts that into the base unit, prices
> it from the Ingredients tab, works out the nutrition, multiplies it by
> how many times the recipe appears in your meal plan and hands the answer
> to the grocery list. If the automatic shopping list ever looks wrong,
> this is the tab to check.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Totals strip: how healthy is the engine room? |

**Module constants:** `NAME` ('Recipe Ingredients').

````python
"""`Recipe Ingredients` - one line per ingredient in a recipe.

This is the engine room. Every line says *recipe R-0007 needs 300 g of
chicken breast*; the workbook converts that into the base unit, prices
it from the Ingredients tab, works out the nutrition, multiplies it by
how many times the recipe appears in your meal plan and hands the answer
to the grocery list. If the automatic shopping list ever looks wrong,
this is the tab to check.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, base_unit_of, id_col, lookup_ingredient,
                     lookup_recipe, to_base)

NAME = "Recipe Ingredients"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("ingredient", '"RL-"', digits=4, width=9),
        Col("recipe", "Recipe ID", 11, "text",
            validate=list_from("RecipeIDs"),
            note="Pick the recipe this line belongs to. The drop-down "
                 "lists every ID on the Recipes tab."),
        Col("recipe_name", "Recipe", 26, "text",
            formula=lookup_recipe("name")),
        Col("ingredient", "Ingredient", 24, "text",
            validate=list_from("IngredientNames"),
            note="Must match the Ingredients tab exactly - choose from "
                 "the drop-down and it always will."),
        Col("qty", "Quantity", 10, "number1",
            validate=number_check(0, 100000),
            note="For the number of servings on the Recipes tab, not "
                 "per person."),
        Col("unit", "Unit", 11, "text", validate=list_from("UnitList"),
            note="Any unit from the table on the Lists tab - kg, g, ml, "
                 "tbsp, cup, each. It is converted for you."),
        Col("prep", "Prepared how", 18, "text",
            note="'finely chopped', 'drained', 'at room temperature'."),
        Col("section", "Part of the recipe", 16, "text",
            note="Optional: 'sauce', 'topping', 'marinade' - keeps long "
                 "recipes readable."),
        Col("optional", "Optional?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Optional ingredients are still costed but are marked "
                 "on the shopping list so you can skip them."),
        # ------------------------------------------------- the conversion
        Col("base_qty", "In base units", 12, "number1",
            formula=to_base("qty", "unit"),
            note="Automatic: the quantity converted to g, ml or each "
                 "using the unit table on the Lists tab."),
        Col("base_unit", "Base unit", 10, "text",
            formula=base_unit_of("unit")),
        Col("price_unit", "Priced in", 10, "text",
            formula=lookup_ingredient("base_unit")),
        Col("unit_cost", "Cost per unit", 12, "money",
            formula=lookup_ingredient("unit_cost"), hidden=True),
        Col("line_cost", "Line cost", 11, "money",
            formula='=IF(OR({base_qty}{r}="",{unit_cost}{r}="",'
                    'NOT(ISNUMBER({unit_cost}{r}))),"",'
                    'IF({base_unit}{r}<>{price_unit}{r},"",'
                    '{base_qty}{r}*{unit_cost}{r}))',
            note="Quantity x the price per unit from the Ingredients "
                 "tab. Blank if the units do not match - fix the unit "
                 "and it fills itself in."),
        Col("cost_serving", "Cost a serving", 12, "money",
            formula='=IF({line_cost}{r}="","",IFERROR({line_cost}{r}/'
                    'INDEX([[Recipes.servings]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0)),""))', hidden=True),
        # -------------------------------------------------- the nutrition
        Col("per100", "Per 100", 9, "number",
            formula='=IF({base_qty}{r}="","",'
                    'IF({base_unit}{r}="each",{base_qty}{r},'
                    '{base_qty}{r}/100))',
            hidden=True,
            note="How many 'per 100 g' blocks this line is worth."),
        Col("line_cals", "Calories", 10, "number",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.calories]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")'),
        Col("line_protein", "Protein (g)", 10, "number1",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.protein]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")',
            hidden=True),
        Col("line_carbs", "Carbs (g)", 10, "number1",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.carbs]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")',
            hidden=True),
        Col("line_fat", "Fat (g)", 9, "number1",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.fat]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")',
            hidden=True),
        Col("line_fibre", "Fibre (g)", 10, "number1",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.fibre]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")',
            hidden=True),
        Col("allergen", "Allergen", 13, "text",
            formula=lookup_ingredient("allergen"),
            cf=[{"type": "text", "criteria": "not containing",
                 "value": "None", "fmt": "warn"}]),
        Col("category", "Shopping category", 15, "text",
            formula=lookup_ingredient("category"), hidden=True),
        # -------------------------------------- what the meal plan needs
        Col("plan_factor", "Planned batches", 11, "number1",
            formula='=IFERROR(INDEX([[Recipes.plan_scale]],MATCH('
                    '{recipe}{r},[[Recipes.id]],0)),0)',
            hidden=True,
            note="How many recipes-worth of this dish the shopping "
                 "window asks for."),
        Col("need_base", "Needed for the plan", 13, "number1",
            formula='=IF(OR({base_qty}{r}="",{plan_factor}{r}=0),"",'
                    '{base_qty}{r}*{plan_factor}{r})',
            note="Quantity x planned batches. This is the number the "
                 "grocery list adds up."),
        Col("scale_key", "Scale key", 10, "text",
            formula='=IF({ingredient}{r}="","",{recipe}{r}&"#"&'
                    'COUNTIFS({recipe_first}:{recipe}{r},{recipe}{r}))',
            hidden=True,
            note="Recipe ID plus the line number within that recipe - "
                 "how the recipe scaler walks the ingredient list."),
        Col("key", "Key", 10, "text",
            formula='=IF({ingredient}{r}="","",{recipe}{r}&"|"&'
                    '{ingredient}{r})', hidden=True),
        Col("alert", "Check", 26, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({recipe}{r}="","Add - which recipe?",'
                    'IF(COUNTIFS([[Ingredients.name]],{ingredient}{r})=0,'
                    '"Missing - not in the Ingredients tab",'
                    'IF({qty}{r}="","Add - no quantity",'
                    'IF({unit}{r}="","Add - no unit",'
                    'IF({base_unit}{r}<>{price_unit}{r},'
                    '"Check - recipe in "&{base_unit}{r}&'
                    '", priced in "&{price_unit}{r},'
                    'IF(COUNTIFS({key_first}:{key_last},{key}{r})>1,'
                    '"Duplicate - same ingredient twice","OK")))))))',
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Recipe Ingredients",
        subtitle="One line per ingredient per recipe. This is what "
                 "turns a meal plan into a shopping list.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=4,
        samples=sample(cfg).recipe_lines,
        help_text="Sort or filter by Recipe ID to work on one recipe at "
                  "a time. Everything to the right of the unit column "
                  "fills itself in.",
        extra_notes=[
            "Quantities are for the recipe as written - the serves "
            "figure on the Recipes tab - and the planner scales them "
            "up or down for the number of people you are feeding.",
            "Units are converted through the table on the Lists tab. "
            "Weight and volume are deliberately kept apart: 100 ml of "
            "oil is not 100 g, so price oil in whichever one you cook "
            "with and the line cost appears.",
            "A line with a unit mismatch is flagged in the Check "
            "column and left out of the cost rather than quietly "
            "guessing at it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals strip: how healthy is the engine room?"""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Ingredient line check", last)
    sheet.table_head(row, 0, ["Lines entered", "Recipes covered",
                              "Lines with a cost", "Unit mismatches",
                              "Not in the price book",
                              "Total value of all recipes"],
                     [15, 16, 16, 15, 20, 24])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Recipe Ingredients.ingredient]],"<>")',
        '=SUMPRODUCT((COUNTIF([[Recipe Ingredients.recipe]],'
        '[[Recipes.id]])>0)*([[Recipes.id]]<>""))',
        '=COUNTIFS([[Recipe Ingredients.line_cost]],">0")',
        '=COUNTIF([[Recipe Ingredients.alert]],"Check - *")',
        '=COUNTIF([[Recipe Ingredients.alert]],"Missing*")',
        '=SUM([[Recipe Ingredients.line_cost]])',
    ], ["int", "int", "int", "int", "int", "money"])
````

### `meal_planner_spreadsheet/sheets/ingredients.py`

*230 lines · 2 functions*

> `Ingredients` - the price book every cost and shopping line reads.
>
> One row per ingredient you buy. It holds the category and aisle (so the
> grocery list can group itself), the pack you normally buy, what that pack
> costs and the nutrition per 100 g or 100 ml. Everything else in the
> workbook - recipe costs, the shopping list, the pantry value, the
> nutrition tracker - looks its numbers up here, so a price only ever gets
> typed in one place.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Summary strip under the price book |

**Module constants:** `NAME` ('Ingredients').

````python
"""`Ingredients` - the price book every cost and shopping line reads.

One row per ingredient you buy. It holds the category and aisle (so the
grocery list can group itself), the pack you normally buy, what that pack
costs and the nutrition per 100 g or 100 ml. Everything else in the
workbook - recipe costs, the shopping list, the pantry value, the
nutrition tracker - looks its numbers up here, so a price only ever gets
typed in one place.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, min_number

NAME = "Ingredients"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "IngredientPrefix", digits=4, width=10),
        Col("name", "Ingredient", 24, "text",
            note="Use one spelling and stick to it - 'chicken breast', "
                 "not 'chicken breasts' as well. The grocery list adds "
                 "up by this name."),
        Col("category", "Category", 16, "text",
            validate=list_from("GroceryCategoryList"),
            note="Groups the shopping list."),
        Col("aisle", "Aisle", 17, "text", validate=list_from("AisleList"),
            note="Sorts the printed list into the order you walk round "
                 "the shop."),
        Col("store", "Usual store", 18, "text",
            validate=list_from("StoreNames")),
        Col("brand", "Preferred brand", 16, "text"),
        Col("base_unit", "Measured in", 12, "text",
            validate=list_from("UnitList"),
            note="The unit you cook with: g, ml or each. Recipe "
                 "quantities in any other unit are converted to this."),
        Col("pack_size", "Pack size", 11, "number1",
            validate=number_check(0, 1000000),
            note="How much is in the pack you normally buy, in the unit "
                 "above. A 400 g tin = 400."),
        Col("pack_price", "Pack price", 11, "money",
            validate=number_check(0, 100000)),
        Col("unit_cost", "Cost per unit", 12, "money",
            formula='=IF(OR({pack_size}{r}="",{pack_price}{r}=""),"",'
                    'IF({pack_size}{r}=0,"",{pack_price}{r}/'
                    '{pack_size}{r}))',
            note="Automatic. This is what every recipe cost is built "
                 "from."),
        Col("per_100", "Cost per 100", 11, "money",
            formula='=IF({unit_cost}{r}="","",{unit_cost}{r}*100)',
            note="Cost per 100 g or 100 ml - the number to compare "
                 "between brands and pack sizes."),
        Col("avg_paid", "Average price paid", 13, "money",
            formula='=IF({name}{r}="","",IFERROR(AVERAGEIFS('
                    '[[Grocery Spending.unit_price]],'
                    '[[Grocery Spending.item]],{name}{r},'
                    '[[Grocery Spending.unit_price]],">0"),""))',
            note="What you have actually been paying, from the purchase "
                 "log."),
        Col("best_price", "Cheapest store price", 14, "money",
            formula='=IF({name}{r}="","",IF(COUNTIFS('
                    '[[Store Prices.ingredient]],{name}{r})=0,"",'
                    'IF(' + min_number('[[Store Prices.unit_cost]]',
                                       '[[Store Prices.ingredient]]='
                                       '{name}{r}') + '>=999999,"",'
                    + min_number('[[Store Prices.unit_cost]]',
                                 '[[Store Prices.ingredient]]='
                                 '{name}{r}') + ')))',
            note="The lowest price per unit any store on the Store "
                 "Prices tab charges for it."),
        Col("best_store", "Cheapest at", 16, "text",
            formula='=IF({best_price}{r}="","",IFERROR(INDEX('
                    '[[Store Prices.store]],MATCH(1,INDEX(('
                    '[[Store Prices.ingredient]]={name}{r})*'
                    '([[Store Prices.unit_cost]]={best_price}{r}),0),0)),'
                    '""))'),
        Col("saving", "Saving per 100", 12, "money",
            formula='=IF(OR({best_price}{r}="",{unit_cost}{r}=""),"",'
                    'MAX(0,({unit_cost}{r}-{best_price}{r})*100))',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}],
            note="What you would save per 100 g / 100 ml by buying it "
                 "where it is cheapest."),
        Col("calories", "Calories per 100 / item", 12, "number",
            validate=number_check(0, 2000),
            note="Per 100 g or 100 ml from the label - or per item for "
                 "anything measured in 'each', like an egg."),
        Col("protein", "Protein per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("carbs", "Carbs per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("fat", "Fat per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("fibre", "Fibre per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("sugar", "Sugar per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("sodium", "Sodium per 100 (mg)", 12, "number",
            validate=number_check(0, 100000)),
        Col("allergen", "Allergen", 14, "text",
            validate=list_from("AllergenList"),
            note="Flags every recipe that uses this ingredient."),
        Col("diet", "Suits", 15, "text", validate=list_from("DietList"),
            note="The strictest diet this ingredient fits."),
        Col("season", "Best season", 12, "text",
            validate=list_from("SeasonList")),
        Col("shelf_days", "Keeps for (days)", 12, "int",
            validate=number_check(0, 3650),
            note="Typical shelf life once you get it home. Used to "
                 "suggest an expiry date when you add it to the pantry."),
        Col("storage", "Store it in", 13, "text",
            validate=list_from("StorageList")),
        Col("substitute", "Swap it for", 20, "text",
            note="What you use when the shop has run out."),
        Col("last_bought", "Last bought", 12, "date", validate=date_check(),
            note="Updated by hand, or left to the purchase log."),
        Col("in_pantry", "In the pantry now", 13, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Pantry.base_qty]],'
                    '[[Pantry.ingredient]],{name}{r}))',
            note="Automatic, in the unit above. Comes from the Pantry "
                 "tab."),
        Col("pantry_value", "Value in stock", 12, "money",
            formula='=IF(OR({in_pantry}{r}="",{unit_cost}{r}=""),"",'
                    '{in_pantry}{r}*{unit_cost}{r})'),
        Col("used_in", "Used in recipes", 12, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Recipe Ingredients.ingredient]],{name}{r}))',
            note="How many of your recipes call for it."),
        Col("need_now", "Needed for the plan", 13, "number1",
            formula='=IF({name}{r}="","",SUMIFS('
                    '[[Recipe Ingredients.need_base]],'
                    '[[Recipe Ingredients.ingredient]],{name}{r}))',
            note="Everything your planned meals call for in the "
                 "shopping window, in base units."),
        Col("spend_key", "Spend key", 9, "number",
            formula='=IF({name}{r}="",0,IFERROR({per_100}{r}*1,0)*'
                    'IFERROR({used_in}{r}*1,0))',
            hidden=True,
            note="Cost per 100 times the number of recipes that use "
                 "it - a rough 'where the money goes' score."),
        Col("spend_rank", "Spend rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({spend_key_first}:'
                    '{spend_key_last},">"&{spend_key}{r})+'
                    'COUNTIFS({spend_key_first}:{spend_key}{r},'
                    '{spend_key}{r}))',
            hidden=True, note="1 = the ingredient you spend most on."),
        Col("waste_count", "Times binned", 9, "int",
            formula='=IF({name}{r}="",0,COUNTIFS([[Food Waste.item]],'
                    '{name}{r}))',
            hidden=True),
        Col("waste_rank", "Waste rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({waste_count_first}:'
                    '{waste_count_last},">"&{waste_count}{r})+'
                    'COUNTIFS({waste_count_first}:{waste_count}{r},'
                    '{waste_count}{r}))',
            hidden=True, note="1 = binned most often."),
        Col("need_rank", "Rank", 7, "int",
            formula='=IF({name}{r}="","",IF({need_now}{r}>0,'
                    'COUNTIF({need_now}${first}:{need_now}{r},">0"),""))',
            hidden=True,
            note="Position on the automatic grocery list."),
        Col("to_buy", "On the list now", 12, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Grocery List.to_buy]],'
                    '[[Grocery List.ingredient]],{name}{r}))'),
        Col("alert", "Check", 24, "status",
            formula=blank(
                "name",
                'IF({pack_price}{r}="","Add - no price yet",'
                'IF({pack_size}{r}="","Add - no pack size",'
                'IF(COUNTIFS({name_first}:{name_last},{name}{r})>1,'
                '"Duplicate - same name twice",'
                'IF({used_in}{r}=0,"Check - no recipe uses it","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Ingredients & Prices",
        subtitle="Your price book. Every recipe cost, every shopping "
                 "total and every nutrition figure in the workbook is "
                 "built from this one list.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).ingredients,
        help_text="Fill in the pack size and pack price for the twenty "
                  "or thirty things you buy every week and the rest of "
                  "the workbook starts doing the maths for you.",
        extra_notes=[
            "Pack size is in the unit you cook with: a 1 kg bag of "
            "flour measured in g is 1000, a 6 pack of eggs measured in "
            "'each' is 6.",
            "Cost per 100 is the honest comparison between a big bag "
            "and a small one - the supermarket shelf label does the "
            "same sum.",
            "Nutrition is per 100 g or 100 ml so it scales with any "
            "recipe quantity. Copy it off the label; leave it blank if "
            "you are not tracking.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the price book."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1        # 1 based Excel row
    last = spec_.last_col_index
    row = sheet.section(row, "Price book summary", last)
    sheet.table_head(row, 0, ["Ingredients listed", "With a price",
                              "Missing a price", "Value in the pantry",
                              "Average cost per 100",
                              "Could be cheaper elsewhere"],
                     [18, 14, 15, 18, 19, 24])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Ingredients.name]],"<>")',
        '=COUNTIFS([[Ingredients.name]],"<>",'
        '[[Ingredients.pack_price]],"<>")',
        '=COUNTIFS([[Ingredients.name]],"<>",'
        '[[Ingredients.pack_price]],"")',
        '=SUM([[Ingredients.pantry_value]])',
        '=IFERROR(AVERAGE([[Ingredients.per_100]]),"")',
        '=COUNTIFS([[Ingredients.saving]],">0")',
    ], ["int", "int", "int", "money", "money", "int"])
````

### `meal_planner_spreadsheet/sheets/meal_plan.py`

*177 lines · 2 functions*

> `Meal Plan` - every planned meal, one row at a time.
>
> The Weekly Planner and Monthly Planner tabs are pretty views of this
> list. The grocery list, the budget, the nutrition tracker and the
> rotation report all read it. Add a row here - date, meal, recipe, how
> many you are feeding - and the rest of the workbook reacts.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Totals strip under the plan |

**Module constants:** `NAME` ('Meal Plan').

````python
"""`Meal Plan` - every planned meal, one row at a time.

The Weekly Planner and Monthly Planner tabs are pretty views of this
list. The grocery list, the budget, the nutrition tracker and the
rotation report all read it. Add a row here - date, meal, recipe, how
many you are feeding - and the rest of the workbook reacts.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, lookup_recipe, id_col, month_col, status_cf,
                     week_col, weekday_col)

NAME = "Meal Plan"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "MealPrefix", digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="One row per meal. Use the Weekly Planner tab to see "
                 "it laid out as a week."),
        weekday_col("date"),
        week_col("date"),
        month_col("date"),
        Col("slot", "Meal", 12, "text", validate=list_from("MealSlotList"),
            note="Breakfast, Lunch, Dinner, Snack... set your own on "
                 "the Settings tab."),
        Col("recipe", "Recipe ID", 11, "text",
            validate=list_from("RecipeIDs"),
            note="Leave blank for a meal you are not cooking from a "
                 "recipe - eating out, leftovers, cereal."),
        Col("recipe_name", "What you are eating", 28, "text",
            formula='=IF({recipe}{r}<>"",IFERROR(INDEX('
                    '[[Recipes.name]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0)),"Not found"),{freetext}{r})',
            note="Filled in from the recipe, or from whatever you typed "
                 "in the next column."),
        Col("freetext", "Or type it in", 22, "text",
            note="'Leftover chilli', 'Fish and chips', 'Mum cooking'. "
                 "No recipe needed."),
        Col("servings", "People to feed", 11, "int",
            validate=number_check(0, 100, decimal=False),
            note="Defaults to your household size if you leave it "
                 "blank - set that on the Settings tab."),
        Col("eff_servings", "Servings used", 11, "number1",
            formula='=IF({date}{r}="","",IF({servings}{r}="",'
                    'DefaultServings,{servings}{r}))', hidden=True),
        Col("scale", "Recipe batches", 11, "number1",
            formula='=IF({recipe}{r}="",0,IFERROR({eff_servings}{r}/'
                    'INDEX([[Recipes.servings]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0)),0))',
            note="People to feed divided by what the recipe serves. "
                 "1.5 means one and a half times the quantities."),
        Col("shop", "Shop for it?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Yes puts its ingredients on the grocery list. Set it "
                 "to No for a meal you have already shopped for, or "
                 "one you are eating out."),
        Col("status", "Status", 13, "status",
            validate=list_from("MealStatusList"), cf=status_cf()),
        Col("cook", "Who is cooking", 15, "text",
            validate=list_from("MemberNames")),
        Col("eating", "Who is eating", 18, "text",
            note="Leave blank for everybody, or name the ones who are "
                 "in."),
        Col("leftovers", "Planned leftovers", 11, "number1",
            validate=number_check(0, 100),
            note="Extra servings you mean to make. They show up on the "
                 "Leftovers tab."),
        Col("prep_ahead", "Prep ahead?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Flags it on the Meal Prep tab."),
        Col("event", "Occasion", 16, "text",
            validate=list_from("EventNames"),
            note="Links the meal to anything on the Special Events tab."),
        Col("cost", "Planned cost", 12, "money",
            formula='=IF({recipe}{r}="","",IFERROR(INDEX('
                    '[[Recipes.cost_serving]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0))*{eff_servings}{r},""))',
            note="Cost per serving from the recipe x the number of "
                 "people. What it should cost, not what you spent."),
        Col("actual_cost", "Actual spend", 12, "money",
            validate=number_check(0, 100000),
            note="Only for meals out or takeaways - the weekly shop "
                 "goes on the Grocery Spending tab."),
        Col("calories", "Calories each", 11, "number",
            formula=lookup_recipe("calories")),
        Col("protein", "Protein each (g)", 11, "number1",
            formula=lookup_recipe("protein"), hidden=True),
        Col("carbs", "Carbs each (g)", 11, "number1",
            formula=lookup_recipe("carbs"), hidden=True),
        Col("fat", "Fat each (g)", 10, "number1",
            formula=lookup_recipe("fat"), hidden=True),
        Col("fibre", "Fibre each (g)", 10, "number1",
            formula=lookup_recipe("fibre"), hidden=True),
        Col("cals_total", "Calories on the table", 12, "number",
            formula='=IF(OR({calories}{r}="",NOT(ISNUMBER('
                    '{calories}{r}))),"",{calories}{r}*'
                    '{eff_servings}{r})', hidden=True),
        Col("minutes", "Minutes in the kitchen", 11, "int",
            formula=lookup_recipe("total_min"), hidden=True),
        Col("category", "Category", 13, "text",
            formula=lookup_recipe("category"), hidden=True),
        Col("home_cooked", "Cooked at home?", 11, "flag",
            formula='=IF({date}{r}="","",IF({recipe}{r}<>"","Yes",'
                    'IF(OR({freetext}{r}="",ISNUMBER(SEARCH("takeaway",'
                    '{freetext}{r})),ISNUMBER(SEARCH("out",'
                    '{freetext}{r}))),"No","Yes")))', hidden=True),
        Col("key", "Key", 12, "text",
            formula='=IF({date}{r}="","",TEXT({date}{r},"yyyy-mm-dd")&'
                    '"|"&{slot}{r})', hidden=True,
            note="Date and meal together - how the weekly and monthly "
                 "planners find this row."),
        Col("alert", "Check", 24, "status",
            formula='=IF({date}{r}="","",'
                    'IF(AND({recipe}{r}="",{freetext}{r}=""),'
                    '"Add - nothing planned yet",'
                    'IF(COUNTIFS({key_first}:{key_last},{key}{r})>1,'
                    '"Duplicate - two meals in one slot",'
                    'IF(AND({recipe}{r}<>"",{recipe_name}{r}='
                    '"Not found"),"Missing - recipe ID not on the '
                    'Recipes tab",IF(AND({date}{r}<TODAY(),'
                    'OR({status}{r}="Planned",{status}{r}="")),'
                    '"Check - did you cook it?","OK")))))',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Meal Plan",
        subtitle="Every planned meal in one list. The weekly and "
                 "monthly planners, the grocery list and the budget all "
                 "read from here.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).meal_plan,
        help_text="You can plan straight onto this tab, or use the "
                  "Weekly Planner and let it write the rows for you.",
        extra_notes=[
            "'Shop for it?' is the switch between a plan and a "
            "shopping list. Set it to No for anything you have already "
            "bought, and the ingredients drop off the list.",
            "'Recipe batches' is how the scaling works: feeding 6 from "
            "a recipe that serves 4 gives 1.5, and every ingredient is "
            "multiplied by 1.5 on the grocery list.",
            "A meal with no recipe is still worth a row - the budget "
            "and the eating-out counter both use it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals strip under the plan."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Plan summary", last)
    sheet.table_head(row, 0, ["Meals planned", "From a recipe",
                              "Cooked at home", "Planned cost",
                              "Eaten out / takeaway", "Still to cook"],
                     [14, 14, 14, 16, 19, 14])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Meal Plan.date]],">0")',
        '=COUNTIF([[Meal Plan.recipe]],"<>")',
        '=COUNTIF([[Meal Plan.home_cooked]],"Yes")',
        '=SUM([[Meal Plan.cost]])',
        '=COUNTIF([[Meal Plan.home_cooked]],"No")',
        '=COUNTIFS([[Meal Plan.status]],"Planned")',
    ], ["int", "int", "int", "money", "int", "int"])
````

### `meal_planner_spreadsheet/sheets/grocery.py`

*285 lines · 2 functions*

> `Grocery List` - the shopping list that writes itself.
>
> Nothing in the first ten columns is typed. The list looks at every meal
> planned between the two dates on the Settings tab, multiplies each
> recipe by the number of people you are feeding, adds up the ingredients,
> takes off whatever is already in the pantry and tells you what is left
> to buy, in whole packs, with a running cost.
>
> You tick things off as you shop, and anything the recipes do not cover -
> bin bags, birthday candles - goes in the extras block underneath.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Totals, the budget check, the store split and the extras block |

**Module constants:** `NAME` ('Grocery List'), `EXTRA_ROWS` (24).

````python
"""`Grocery List` - the shopping list that writes itself.

Nothing in the first ten columns is typed. The list looks at every meal
planned between the two dates on the Settings tab, multiplies each
recipe by the number of people you are feeding, adds up the ingredients,
takes off whatever is already in the pantry and tells you what is left
to buy, in whole packs, with a running cost.

You tick things off as you shop, and anything the recipes do not cover -
bin bags, birthday candles - goes in the extras block underneath.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..writer import list_from, number_check
from .common import ROW_OFFSET, lookup_ingredient, status_cf

NAME = "Grocery List"

EXTRA_ROWS = 24


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("rank", "#", 5, "int",
            formula=f"=ROW()-{ROW_OFFSET}", hidden=True),
        Col("ingredient", "Item", 24, "text",
            formula='=IFERROR(INDEX([[Ingredients.name]],MATCH('
                    '{rank}{r},[[Ingredients.need_rank]],0)),"")',
            note="Automatic. Every ingredient your planned meals need, "
                 "in the order they appear in the price book."),
        Col("category", "Category", 15, "text",
            formula=lookup_ingredient("category")),
        Col("aisle", "Aisle", 16, "text",
            formula=lookup_ingredient("aisle"),
            note="Sort by this column to walk the shop once."),
        Col("store", "Store", 16, "text",
            formula=lookup_ingredient("store")),
        Col("needed", "Needed", 10, "number1",
            formula='=IF({ingredient}{r}="","",IFERROR(INDEX('
                    '[[Ingredients.need_now]],MATCH({ingredient}{r},'
                    '[[Ingredients.name]],0)),""))',
            note="Everything the planned meals call for, in base "
                 "units."),
        Col("unit", "Unit", 7, "text",
            formula=lookup_ingredient("base_unit")),
        Col("in_pantry", "Already have", 11, "number1",
            formula='=IF({ingredient}{r}="","",IFERROR(INDEX('
                    '[[Ingredients.in_pantry]],MATCH({ingredient}{r},'
                    '[[Ingredients.name]],0)),0))',
            note="Straight from the Pantry tab - this is the "
                 "deduction that stops you buying a fourth jar of "
                 "cumin."),
        Col("short", "Still short", 10, "number1",
            formula='=IF({needed}{r}="","",MAX(0,{needed}{r}-'
                    'IF({in_pantry}{r}="",0,{in_pantry}{r})))',
            cf=[{"type": "cell", "criteria": "=", "value": 0,
                 "fmt": "muted"}]),
        Col("pack_size", "Pack size", 10, "number1",
            formula=lookup_ingredient("pack_size"), hidden=True),
        Col("pack_price", "Pack price", 10, "money",
            formula=lookup_ingredient("pack_price"), hidden=True),
        Col("unit_cost", "Cost per unit", 10, "money",
            formula=lookup_ingredient("unit_cost"), hidden=True),
        Col("packs", "Packs to buy", 11, "number1",
            formula='=IF(OR({short}{r}="",{short}{r}=0),"",'
                    'IF(OR({pack_size}{r}="",NOT(ISNUMBER('
                    '{pack_size}{r})),{pack_size}{r}=0),"",'
                    'ROUNDUP({short}{r}/{pack_size}{r},0)))',
            note="How many of the pack you normally buy. 1.2 tins "
                 "becomes 2 tins, because shops do not sell 0.2 of a "
                 "tin."),
        Col("buy_qty", "Buy this much", 12, "number1",
            formula='=IF({short}{r}="","",IF(RoundingRule='
                    '"Exact quantity",{short}{r},IF({packs}{r}="",'
                    '{short}{r},{packs}{r}*{pack_size}{r})))',
            note="Follows the rounding rule on the Settings tab."),
        Col("est_cost", "Estimated cost", 12, "money",
            formula='=IF({buy_qty}{r}="","",IF(RoundingRule='
                    '"Exact quantity",IF({unit_cost}{r}="","",'
                    '{buy_qty}{r}*{unit_cost}{r}),'
                    'IF(OR({packs}{r}="",{pack_price}{r}=""),"",'
                    '{packs}{r}*{pack_price}{r})))',
            cf=[{"type": "data_bar", "bar_color": "#4E8A63"}]),
        Col("running", "Running total", 12, "money",
            formula='=IF({ingredient}{r}="","",SUM({est_cost}${first}:'
                    '{est_cost}{r}))',
            note="Watch this against your weekly budget as you plan."),
        Col("to_buy", "To buy", 10, "number1",
            formula='=IF({got}{r}="Yes",0,IF({buy_qty}{r}="",0,'
                    '{buy_qty}{r}))', hidden=True),
        Col("buy_rank", "Buy rank", 8, "int",
            formula='=IF(OR({ingredient}{r}="",{short}{r}=0,'
                    '{short}{r}=""),"",COUNTIFS({short_first}:'
                    '{short}{r},">0"))',
            hidden=True,
            note="Position on the printable shopping list."),
        Col("recipes", "For how many meals", 11, "int",
            formula='=IF({ingredient}{r}="","",COUNTIFS('
                    '[[Recipe Ingredients.ingredient]],{ingredient}{r},'
                    '[[Recipe Ingredients.need_base]],">0"))',
            note="How many of the week's recipes want it."),
        Col("optional", "Optional only?", 11, "flag",
            formula='=IF({ingredient}{r}="","",IF(COUNTIFS('
                    '[[Recipe Ingredients.ingredient]],{ingredient}{r},'
                    '[[Recipe Ingredients.need_base]],">0",'
                    '[[Recipe Ingredients.optional]],"No")=0,"Yes",'
                    '"No"))',
            note="Yes means every recipe that wants it treats it as "
                 "optional - skip it if money is tight.",
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "muted"}]),
        Col("allergen", "Allergen", 12, "text",
            formula=lookup_ingredient("allergen"), hidden=True),
        # ----------------------------------------------- what you do
        Col("got", "Got it?", 9, "flag", validate=list_from("YesNoList"),
            note="Tick as you shop. Ticked rows drop out of the "
                 "remaining total below."),
        Col("paid", "Actually paid", 11, "money",
            validate=number_check(0, 100000),
            note="Optional - fill it in at the till and the workbook "
                 "shows you the gap between estimate and reality."),
        Col("swap", "Swapped for", 16, "text",
            note="What you bought instead when the shelf was empty."),
        Col("note", "Note", 20, "wrap"),
        Col("status", "Status", 12, "status",
            formula='=IF({ingredient}{r}="","",IF({got}{r}="Yes",'
                    '"Bought",IF({short}{r}=0,"In stock",'
                    'IF({est_cost}{r}="","Check - no price",'
                    '"On the list"))))',
            cf=status_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Automatic Grocery List",
        subtitle="Built from your meal plan, minus what is already in "
                 "the pantry, grouped by aisle and priced as you go.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=[],
        help_text="Nothing to type on the left hand side - change the "
                  "shopping window on the Settings tab and the whole "
                  "list rebuilds. Tick 'Got it?' as you shop.",
        extra_notes=[
            "The list covers every meal planned between 'Shopping list "
            "covers from' and 'covers to' on the Settings tab, where "
            "'Shop for it?' on the Meal Plan says Yes.",
            "Quantities are in base units - g, ml or each. 'Buy this "
            "much' rounds them up to whole packs unless you change "
            "the rounding rule on the Settings tab.",
            "Ticking 'Got it?' does not delete the row - it keeps the "
            "record of the shop and takes the item out of the "
            "remaining total.",
            "An ingredient only appears here if it is on the "
            "Ingredients tab. Anything flagged 'Missing' on the Recipe "
            "Ingredients tab is quietly left out, so clear those first.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals, the budget check, the store split and the extras block."""
    spec_ = reg[NAME]
    s = sheet
    ws = s.ws
    last = spec_.last_col_index
    row = spec_.last_data_row + 1          # 0 based, one row under the grid

    # ------------------------------------------------------------ totals --
    row = s.section(row, "This shop", last)
    row = s.table_head(row, 0, ["Items on the list", "Already in stock",
                                "Still to buy", "Estimated cost",
                                "Picked up so far", "Left to pick up",
                                "Actually paid"],
                       [16, 15, 13, 15, 16, 15, 14])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Grocery List.ingredient]],"<>")',
        '=COUNTIFS([[Grocery List.short]],0,'
        '[[Grocery List.ingredient]],"<>")',
        '=COUNTIFS([[Grocery List.short]],">0")',
        '=SUM([[Grocery List.est_cost]])',
        '=COUNTIF([[Grocery List.got]],"Yes")',
        '=SUMIFS([[Grocery List.est_cost]],[[Grocery List.got]],"<>Yes")',
        '=SUM([[Grocery List.paid]])',
    ], ["int", "int", "int", "money", "int", "money", "money"])
    row += 1

    # ------------------------------------------------------ budget check --
    row = s.note(row, "The estimate is measured against the weekly "
                      "grocery budget on the Settings tab. Change the "
                      "shopping window there too - this list follows it.",
                 last)
    row = s.table_head(row, 0, ["Weekly budget", "This shop", "Difference",
                                "Shopping window",
                                "Days to the shopping day"],
                       [15, 14, 14, 28, 20])
    diff_row = row
    row = s.calc_row(row, 0, [
        "=WeeklyBudget",
        '=SUM([[Grocery List.est_cost]])',
        '=WeeklyBudget-SUM([[Grocery List.est_cost]])',
        '=TEXT(ShopFrom,"d mmm")&" to "&TEXT(ShopTo,"d mmm yyyy")',
        '=MAX(0,MOD(IFERROR(MATCH(ShoppingDay,WeekdayList,0),6)-'
        'WEEKDAY(TODAY(),2),7))',
    ], ["money", "money", "money", "text", "int"])
    ws.conditional_format(diff_row, 2, diff_row, 2,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(diff_row, 2, diff_row, 2,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    row += 1

    # --------------------------------------------------- store by store --
    row = s.section(row, "Store by store", last)
    row = s.table_head(row, 0, ["Store", "Items", "Estimated cost",
                                "Still to pick up", "Cheaper elsewhere"],
                       [22, 10, 15, 16, 18])
    for i in range(6):
        cell = f"A{row + 1}"
        s.write(row, 0, f'=IFERROR(IF(INDEX(StoreNames,{i + 1})=0,"",'
                        f'INDEX(StoreNames,{i + 1})),"")',
                fmts.calc("text"))
        s.calc_row(row, 1, [
            f'=IF({cell}="","",COUNTIFS([[Grocery List.store]],{cell}))',
            f'=IF({cell}="","",SUMIFS([[Grocery List.est_cost]],'
            f'[[Grocery List.store]],{cell}))',
            f'=IF({cell}="","",SUMIFS([[Grocery List.est_cost]],'
            f'[[Grocery List.store]],{cell},[[Grocery List.got]],'
            f'"<>Yes"))',
            f'=IF({cell}="","",SUMIFS([[Ingredients.saving]],'
            f'[[Ingredients.store]],{cell}))',
        ], ["int", "money", "money", "money"])
        row += 1
    row += 1

    # ----------------------------------------------------- extras block --
    row = s.section(row, "Extras - things no recipe asked for", last)
    row = s.note(row, "Bin bags, washing up liquid, a birthday cake. "
                      "These rows are yours to type in and they are "
                      "added to the grand total at the bottom.", last)
    row = s.table_head(row, 0, ["Item", "Category", "How many", "Unit",
                                "Estimated cost", "Store", "Got it?",
                                "Note"],
                       [24, 15, 10, 8, 14, 16, 9, 26])
    first_extra = row
    styles = ["text", "text", "number1", "text", "money", "text", "flag",
              "wrap"]
    for i in range(EXTRA_ROWS):
        r = first_extra + i
        for col_i, style in enumerate(styles):
            s.write(r, col_i, None, fmts.input(style))
        ws.data_validation(r, 1, r, 1, list_from("GroceryCategoryList"))
        ws.data_validation(r, 3, r, 3, list_from("UnitList"))
        ws.data_validation(r, 5, r, 5, list_from("StoreNames"))
        ws.data_validation(r, 6, r, 6, list_from("YesNoList"))
    last_extra = first_extra + EXTRA_ROWS - 1
    row = last_extra + 1
    s.merge(row, 0, 3, "Extras total", fmts.chrome("total_label"))
    extras_sum = f"=SUM(E{first_extra + 1}:E{last_extra + 1})"
    ws.write_formula(row, 4, extras_sum,
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    s.merge(row, 5, min(last, 9),
            "Typed in by you - no recipe needed.",
            fmts.chrome("total_value", align="left"))
    row += 2

    # ------------------------------------------------------ grand total --
    s.merge(row, 0, 3, "Everything in the trolley",
            fmts.chrome("total_label"))
    ws.write_formula(
        row, 4,
        s.reg.render('=SUM([[Grocery List.est_cost]])') +
        f"+SUM(E{first_extra + 1}:E{last_extra + 1})",
        fmts.chrome("total_value", num_format=cfg.money_format))
    s.merge(row, 5, min(last, 9),
            "Recipes plus extras. Compare it with the weekly "
            "budget above.",
            fmts.chrome("total_value", align="left"))
````

---

## 4. The kitchen registers

What is in the house and what happens to it: the pantry, the staples, leftovers, waste, batch cooking, the shops and what you spend in them, events and goals.

### `meal_planner_spreadsheet/sheets/pantry.py`

*188 lines · 2 functions*

> `Pantry` - what is actually in the house.
>
> Every jar, bag and tin you have, with how much is left and when it goes
> off. The grocery list subtracts this from what your meals need, so you
> stop buying a fourth bag of rice; the expiry columns tell you what to
> cook first.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Stock value and what needs attention |

**Module constants:** `NAME` ('Pantry').

````python
"""`Pantry` - what is actually in the house.

Every jar, bag and tin you have, with how much is left and when it goes
off. The grocery list subtracts this from what your meals need, so you
stop buying a fourth bag of rice; the expiry columns tell you what to
cook first.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, base_unit_of, id_col, lookup_ingredient,
                     status_cf, to_base)

NAME = "Pantry"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("ingredient", "PantryPrefix", digits=4, width=10),
        Col("ingredient", "Ingredient", 24, "text",
            validate=list_from("IngredientNames"),
            note="Choose from the Ingredients tab so the grocery list "
                 "can match it. Add anything new there first."),
        Col("category", "Category", 15, "text",
            formula=lookup_ingredient("category")),
        Col("location", "Where it lives", 14, "text",
            validate=list_from("StorageList"),
            note="Cupboard, Fridge, Freezer, Spice rack... The fridge "
                 "and freezer tab filters on this."),
        Col("qty", "How much is left", 11, "number1",
            validate=number_check(0, 100000),
            note="Update it when you notice, not every meal. Even a "
                 "rough number beats no number."),
        Col("unit", "Unit", 10, "text", validate=list_from("UnitList")),
        Col("base_qty", "In base units", 11, "number1",
            formula=to_base("qty", "unit"),
            note="What the grocery list deducts."),
        Col("base_unit", "Base unit", 9, "text",
            formula=base_unit_of("unit"), hidden=True),
        Col("opened", "Opened?", 9, "flag", validate=list_from("YesNoList")),
        Col("bought", "Bought on", 11, "date", validate=date_check()),
        Col("expiry", "Use by", 11, "date", validate=date_check(),
            note="Leave it blank and the workbook estimates one from "
                 "the shelf life on the Ingredients tab."),
        Col("shelf_days", "Keeps for", 9, "int",
            formula=lookup_ingredient("shelf_days"), hidden=True),
        Col("est_expiry", "Best before (worked out)", 12, "date",
            formula='=IF({expiry}{r}<>"",{expiry}{r},'
                    'IF(OR({bought}{r}="",{shelf_days}{r}="",'
                    'NOT(ISNUMBER({shelf_days}{r}))),"",'
                    '{bought}{r}+{shelf_days}{r}))',
            note="Your date if you typed one, otherwise bought date + "
                 "shelf life."),
        Col("days_left", "Days left", 9, "int",
            formula='=IF({est_expiry}{r}="","",{est_expiry}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": "=ExpiryWarningDays", "fmt": "warn"},
                {"type": "cell", "criteria": ">",
                 "value": "=ExpiryWarningDays", "fmt": "good"}]),
        Col("unit_cost", "Cost per unit", 10, "money",
            formula=lookup_ingredient("unit_cost"), hidden=True),
        Col("value", "Value", 11, "money",
            formula='=IF(OR({base_qty}{r}="",{unit_cost}{r}="",'
                    'NOT(ISNUMBER({unit_cost}{r}))),"",'
                    '{base_qty}{r}*{unit_cost}{r})',
            note="What is sitting in your cupboards, at price book "
                 "prices."),
        Col("min_qty", "Keep at least", 10, "number1",
            validate=number_check(0, 100000),
            note="Your re-order level. Below this it is flagged as low "
                 "stock - the staples list uses it."),
        Col("needed", "Plan needs", 11, "number1",
            formula='=IF({ingredient}{r}="","",IFERROR(INDEX('
                    '[[Ingredients.need_now]],MATCH({ingredient}{r},'
                    '[[Ingredients.name]],0)),0))', hidden=True),
        Col("after_plan", "Left after the plan", 12, "number1",
            formula='=IF({base_qty}{r}="","",{base_qty}{r}-'
                    'IF({needed}{r}="",0,{needed}{r}))',
            note="What should be left once you have cooked this "
                 "week's meals.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "warn"}]),
        Col("staple", "Staple?", 9, "flag", validate=list_from("YesNoList"),
            note="Things you never want to run out of."),
        Col("fridge_rank", "Fridge rank", 9, "int",
            formula='=IF(OR({ingredient}{r}="",{location}{r}<>'
                    '"Fridge",NOT(ISNUMBER({days_left}{r}))),"",'
                    'COUNTIFS({location_first}:{location_last},"Fridge",'
                    '{days_left_first}:{days_left_last},"<"&'
                    '{days_left}{r})+COUNTIFS({location_first}:'
                    '{location}{r},"Fridge",{days_left_first}:'
                    '{days_left}{r},{days_left}{r}))',
            hidden=True,
            note="1 = the fridge item with the least time left."),
        Col("freezer_rank", "Freezer rank", 9, "int",
            formula='=IF(OR({ingredient}{r}="",{location}{r}<>'
                    '"Freezer",NOT(ISNUMBER({days_left}{r}))),"",'
                    'COUNTIFS({location_first}:{location_last},"Freezer",'
                    '{days_left_first}:{days_left_last},"<"&'
                    '{days_left}{r})+COUNTIFS({location_first}:'
                    '{location}{r},"Freezer",{days_left_first}:'
                    '{days_left}{r},{days_left}{r}))',
            hidden=True,
            note="1 = the freezer item with the least time left."),
        Col("status", "Status", 14, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({base_qty}{r}=0,"Out of stock",'
                    'IF(AND({days_left}{r}<>"",{days_left}{r}<0),'
                    '"Expired",IF(AND({days_left}{r}<>"",'
                    '{days_left}{r}<=ExpiryWarningDays),"Use soon",'
                    'IF(AND({min_qty}{r}<>"",{base_qty}{r}<'
                    '{min_qty}{r}),"Low stock","In stock")))))',
            cf=status_cf()),
        Col("alert", "What to do", 28, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({status}{r}="Expired","Expired - bin it and log '
                    'it on the Food Waste tab",'
                    'IF({status}{r}="Use soon","Use soon - plan a meal '
                    'with it",IF({status}{r}="Out of stock",'
                    'IF({staple}{r}="Yes","Missing - staple has run '
                    'out","Add - restock when convenient"),'
                    'IF({status}{r}="Low stock","Check - running low",'
                    'IF({after_plan}{r}<0,"Check - the plan needs more '
                    'than you have","OK"))))))',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Pantry Inventory",
        subtitle="What you already own, what it is worth and what is "
                 "about to go off - the tab that shrinks your shopping "
                 "list.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).pantry,
        help_text="One row per thing you have in stock. You do not "
                  "have to be perfect - even counting the ten things "
                  "you buy most often saves real money.",
        extra_notes=[
            "The grocery list takes 'In base units' off whatever your "
            "meals need. An empty pantry simply means nothing is "
            "deducted.",
            "Leave 'Use by' blank and the workbook works one out from "
            "the shelf life on the Ingredients tab - a sensible guess, "
            "not a guarantee. Trust the label and your nose first.",
            "'Keep at least' turns this into a re-order list: anything "
            "below it is flagged as low stock on the dashboard.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Stock value and what needs attention."""
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Pantry at a glance", last)
    row = s.table_head(row, 0, ["Items in stock", "Value in stock",
                                "Expired", "Use within "
                                f"{cfg.expiry_warning_days} days",
                                "Low stock", "Out of stock",
                                "Staples below par"],
                       [15, 15, 11, 18, 12, 13, 18])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Pantry.ingredient]],"<>")',
        '=SUM([[Pantry.value]])',
        '=COUNTIF([[Pantry.status]],"Expired")',
        '=COUNTIF([[Pantry.status]],"Use soon")',
        '=COUNTIF([[Pantry.status]],"Low stock")',
        '=COUNTIF([[Pantry.status]],"Out of stock")',
        '=COUNTIFS([[Pantry.staple]],"Yes",[[Pantry.status]],'
        '"<>In stock")',
    ], ["int", "money", "int", "int", "int", "int", "int"])
    row += 1
    s.note(row, "Tip: filter the Status column on 'Use soon' before you "
                "plan next week. Cooking what you already have is the "
                "single biggest saving in this workbook.", last)
````

### `meal_planner_spreadsheet/sheets/staples.py`

*123 lines · 2 functions*

> `Staples` - the things you never want to run out of.
>
> Salt, olive oil, pasta, nappies for the dog's dinner. Set a par level
> once and the workbook tells you when stock has dropped below it, how
> long it will last at your usual rate, and roughly when to buy more.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Staples').

````python
"""`Staples` - the things you never want to run out of.

Salt, olive oil, pasta, nappies for the dog's dinner. Set a par level
once and the workbook tells you when stock has dropped below it, how
long it will last at your usual rate, and roughly when to buy more.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, id_col, lookup_ingredient, status_cf

NAME = "Staples"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("ingredient", '"STP-"', digits=3, width=9),
        Col("ingredient", "Staple", 24, "text",
            validate=list_from("IngredientNames")),
        Col("category", "Category", 15, "text",
            formula=lookup_ingredient("category")),
        Col("unit", "Unit", 8, "text",
            formula=lookup_ingredient("base_unit")),
        Col("par", "Always keep", 11, "number1",
            validate=number_check(0, 100000),
            note="Your par level: the amount you want on the shelf at "
                 "all times, in base units."),
        Col("in_stock", "In stock", 11, "number1",
            formula='=IF({ingredient}{r}="","",IFERROR(INDEX('
                    '[[Ingredients.in_pantry]],MATCH({ingredient}{r},'
                    '[[Ingredients.name]],0)),0))'),
        Col("gap", "Short by", 10, "number1",
            formula='=IF(OR({par}{r}="",{in_stock}{r}=""),"",'
                    'MAX(0,{par}{r}-{in_stock}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("usage", "Used a week", 10, "number1",
            validate=number_check(0, 100000),
            note="Roughly how much the house gets through in a week. "
                 "A guess is fine - it only drives the countdown."),
        Col("weeks_left", "Weeks left", 10, "number1",
            formula='=IF(OR({usage}{r}="",{usage}{r}=0,'
                    '{in_stock}{r}=""),"",{in_stock}{r}/{usage}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "between", "minimum": 1,
                 "maximum": 2, "fmt": "warn"}]),
        Col("buy_by", "Buy before", 11, "date",
            formula='=IF({weeks_left}{r}="","",TODAY()+'
                    'ROUND({weeks_left}{r}*7,0))'),
        Col("pack_size", "Pack size", 10, "number1",
            formula=lookup_ingredient("pack_size"), hidden=True),
        Col("pack_price", "Pack price", 10, "money",
            formula=lookup_ingredient("pack_price"), hidden=True),
        Col("packs", "Packs to buy", 10, "number1",
            formula='=IF(OR({gap}{r}="",{gap}{r}=0,{pack_size}{r}="",'
                    'NOT(ISNUMBER({pack_size}{r})),{pack_size}{r}=0),'
                    '"",ROUNDUP({gap}{r}/{pack_size}{r},0))'),
        Col("cost", "Cost to restock", 11, "money",
            formula='=IF(OR({packs}{r}="",{pack_price}{r}="",'
                    'NOT(ISNUMBER({pack_price}{r}))),"",'
                    '{packs}{r}*{pack_price}{r})'),
        Col("store", "Buy it at", 16, "text",
            formula=lookup_ingredient("store")),
        Col("bulk", "Worth buying big?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Only if you have the space and it does not go off - "
                 "bulk buying something you bin is not a saving."),
        Col("status", "Status", 13, "status",
            formula='=IF({ingredient}{r}="","",IF({in_stock}{r}=0,'
                    '"Out of stock",IF(AND({par}{r}<>"",'
                    '{in_stock}{r}<{par}{r}),"Low stock","In stock")))',
            cf=status_cf()),
        Col("alert", "Check", 24, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({status}{r}="Out of stock","Missing - add it '
                    'to the list",IF(AND({weeks_left}{r}<>"",'
                    '{weeks_left}{r}<1),"Due - less than a week left",'
                    'IF({status}{r}="Low stock","Check - below your '
                    'par level","OK"))))',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Pantry Staples",
        subtitle="Par levels for the things that should always be in "
                 "the cupboard, with a countdown to the day they run "
                 "out.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_soft,
        freeze_cols=2,
        samples=sample(cfg).staples,
        help_text="Twenty staples is plenty. Set the par level to what "
                  "makes you comfortable, not what fits in the "
                  "cupboard.",
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Restock list", last)
    row = s.table_head(row, 0, ["Staples tracked", "Below par",
                                "Out of stock", "Running out this week",
                                "Cost to restock everything"],
                       [15, 12, 13, 20, 26])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Staples.ingredient]],"<>")',
        '=COUNTIF([[Staples.status]],"Low stock")',
        '=COUNTIF([[Staples.status]],"Out of stock")',
        '=COUNTIFS([[Staples.weeks_left]],"<1",'
        '[[Staples.weeks_left]],">=0")',
        '=SUM([[Staples.cost]])',
    ], ["int", "int", "int", "int", "money"])
````

### `meal_planner_spreadsheet/sheets/leftovers.py`

*144 lines · 2 functions*

> `Leftovers` - the half a lasagne in the fridge.
>
> Log what is left after a meal, where you put it and how long it is good
> for. The dashboard nags you about it, the planner lets you drop it into
> tomorrow's lunch slot, and anything that does get thrown away rolls
> through to the Food Waste tab with a price on it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Leftovers').

````python
"""`Leftovers` - the half a lasagne in the fridge.

Log what is left after a meal, where you put it and how long it is good
for. The dashboard nags you about it, the planner lets you drop it into
tomorrow's lunch slot, and anything that does get thrown away rolls
through to the Food Waste tab with a price on it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, lookup_recipe, status_cf

NAME = "Leftovers"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("item", "LeftoverPrefix", digits=4, width=10),
        Col("item", "What it is", 26, "text",
            note="'Half the chilli', 'roast chicken - two portions'."),
        Col("recipe", "From which recipe", 11, "text",
            validate=list_from("RecipeIDs"),
            note="Optional. Links it back so the cost per portion is "
                 "worked out for you."),
        Col("recipe_name", "Recipe", 22, "text",
            formula=lookup_recipe("name"), hidden=True),
        Col("cooked", "Cooked on", 11, "date", validate=date_check()),
        Col("portions", "Portions left", 10, "number1",
            validate=number_check(0, 100)),
        Col("location", "Stored in", 12, "text",
            validate=list_from("StorageList"),
            note="Fridge or freezer - it changes how long you have."),
        Col("container", "Container", 14, "text",
            note="'Big glass box', 'blue lid' - so you can find it."),
        Col("window", "Days it keeps", 10, "int",
            formula='=IF({location}{r}="","",IF({location}{r}='
                    '"Freezer",FreezerDays,LeftoverDays))',
            hidden=True),
        Col("use_by", "Eat it by", 11, "date",
            formula='=IF(OR({cooked}{r}="",{window}{r}=""),"",'
                    '{cooked}{r}+{window}{r})',
            note="Cooked date plus the fridge or freezer window from "
                 "the Settings tab."),
        Col("days_left", "Days left", 9, "int",
            formula='=IF({use_by}{r}="","",{use_by}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 1, "fmt": "warn"}]),
        Col("cost_portion", "Cost a portion", 11, "money",
            formula=lookup_recipe("cost_serving"), hidden=True),
        Col("value", "Worth", 11, "money",
            formula='=IF(OR({portions}{r}="",{cost_portion}{r}="",'
                    'NOT(ISNUMBER({cost_portion}{r}))),"",'
                    '{portions}{r}*{cost_portion}{r})',
            note="What you would pay to make it again - the number "
                 "that makes leftovers feel worth eating."),
        Col("planned_for", "Planned for", 11, "date",
            validate=date_check(),
            note="When you intend to eat it. Add the meal to the plan "
                 "as well and it shows up in the week."),
        Col("eaten_by", "Who is having it", 14, "text",
            validate=list_from("MemberNames")),
        Col("queue_rank", "Queue rank", 9, "int",
            formula='=IF(OR({item}{r}="",{status}{r}="Eaten",'
                    '{status}{r}="Binned",NOT(ISNUMBER('
                    '{days_left}{r}))),"",COUNTIFS({status_first}:'
                    '{status_last},"<>Eaten",{status_first}:'
                    '{status_last},"<>Binned",{days_left_first}:'
                    '{days_left_last},"<"&{days_left}{r})+'
                    'COUNTIFS({status_first}:{status}{r},"<>Eaten",'
                    '{status_first}:{status}{r},"<>Binned",'
                    '{days_left_first}:{days_left}{r},{days_left}{r}))',
            hidden=True,
            note="1 = eat this one next. Drives the Fridge & Freezer "
                 "queue and the printable week."),
        Col("status", "Status", 13, "status",
            validate=list_from("LeftoverStatusList"), cf=status_cf()),
        Col("outcome", "What happened", 26, "status",
            formula='=IF({item}{r}="","",'
                    'IF({status}{r}="Eaten","OK - eaten",'
                    'IF({status}{r}="Binned","Wasted - log the cost",'
                    'IF({status}{r}="Frozen","OK - in the freezer",'
                    'IF(AND({days_left}{r}<>"",{days_left}{r}<0),'
                    '"Expired - check before eating",'
                    'IF(AND({days_left}{r}<>"",{days_left}{r}<=1),'
                    '"Today - eat it now","OK"))))))',
            cf=alert_cf()),
        Col("notes", "Notes", 22, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Leftovers Tracker",
        subtitle="What is in the fridge from last night, what it is "
                 "worth and how long you have to eat it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_soft,
        freeze_cols=2,
        samples=sample(cfg).leftovers,
        help_text="Log it while you are putting it in the fridge - it "
                  "takes ten seconds and turns into a free lunch.",
        extra_notes=[
            "The fridge and freezer windows come from the Settings "
            "tab (fridge "
            f"{cfg.leftover_safe_days} days, freezer "
            f"{cfg.freezer_review_days} "
            "days by default). They are guidance, not a guarantee - "
            "look, smell and use your judgement.",
            "Mark a portion 'Eaten' and it counts towards the money "
            "you saved; mark it 'Binned' and it shows up as waste.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Leftovers scoreboard", last)
    row = s.table_head(row, 0, ["Portions in the house", "Worth",
                                "Eat today or tomorrow", "Past their date",
                                "Eaten so far", "Binned so far",
                                "Money saved by eating them"],
                       [18, 12, 19, 16, 13, 13, 24])
    row = s.calc_row(row, 0, [
        '=SUMIFS([[Leftovers.portions]],[[Leftovers.status]],'
        '"<>Eaten",[[Leftovers.status]],"<>Binned")',
        '=SUMIFS([[Leftovers.value]],[[Leftovers.status]],"<>Eaten",'
        '[[Leftovers.status]],"<>Binned")',
        '=COUNTIFS([[Leftovers.days_left]],">=0",'
        '[[Leftovers.days_left]],"<=1")',
        '=COUNTIFS([[Leftovers.days_left]],"<0",'
        '[[Leftovers.status]],"<>Eaten")',
        '=COUNTIF([[Leftovers.status]],"Eaten")',
        '=COUNTIF([[Leftovers.status]],"Binned")',
        '=SUMIFS([[Leftovers.value]],[[Leftovers.status]],"Eaten")',
    ], ["number1", "money", "int", "int", "int", "int", "money"])
````

### `meal_planner_spreadsheet/sheets/waste.py`

*131 lines · 2 functions*

> `Food Waste` - the uncomfortable tab.
>
> Every bag of salad that turned to soup in the drawer, with a price on
> it. It is the fastest way to find the £300 a year most households throw
> out: log it for a month, read the reasons column, and change one thing.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Food Waste').

````python
"""`Food Waste` - the uncomfortable tab.

Every bag of salad that turned to soup in the drawer, with a price on
it. It is the fastest way to find the £300 a year most households throw
out: log it for a month, read the reasons column, and change one thing.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, base_unit_of, id_col, lookup_ingredient,
                     month_col, to_base, week_col)

NAME = "Food Waste"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"W-"', digits=4, width=8),
        Col("date", "Date", 11, "date", validate=date_check()),
        week_col("date"),
        month_col("date"),
        Col("item", "What went in the bin", 24, "text",
            validate=list_from("IngredientNames"),
            note="Pick an ingredient where you can - it prices the "
                 "waste for you."),
        Col("other", "Or describe it", 20, "text",
            note="'Half a takeaway', 'the kids' leftovers'."),
        Col("qty", "How much", 10, "number1",
            validate=number_check(0, 100000)),
        Col("unit", "Unit", 9, "text", validate=list_from("UnitList")),
        Col("base_qty", "In base units", 11, "number1",
            formula=to_base("qty", "unit"), hidden=True),
        Col("base_unit", "Base unit", 9, "text",
            formula=base_unit_of("unit"), hidden=True),
        Col("unit_cost", "Cost per unit", 10, "money",
            formula=lookup_ingredient("unit_cost", key="item"),
            hidden=True),
        Col("value", "What it cost", 11, "money",
            formula='=IF(OR({base_qty}{r}="",{unit_cost}{r}="",'
                    'NOT(ISNUMBER({unit_cost}{r}))),{manual_cost}{r},'
                    '{base_qty}{r}*{unit_cost}{r})',
            note="Worked out from the price book, or your own figure "
                 "in the next column."),
        Col("manual_cost", "Your estimate", 11, "money",
            validate=number_check(0, 10000),
            note="Use this for anything not in the price book."),
        Col("reason", "Why", 18, "text", validate=list_from("WasteReasonList"),
            note="Be honest. The pattern in this column is the whole "
                 "point of the tab."),
        Col("avoidable", "Avoidable?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Peelings and bones are not. A forgotten bag of "
                 "spinach is."),
        Col("where", "Where it was", 13, "text",
            validate=list_from("StorageList")),
        Col("disposal", "How it was thrown", 14, "text",
            validate=list_from("DisposalList"),
            note="Compost and food waste collections are better than "
                 "landfill, but not buying it is better still."),
        Col("prevent", "What would have stopped it", 30, "wrap",
            note="'Freeze half on the day', 'buy the small bag', "
                 "'actually check the fridge first'."),
        Col("month_total", "Wasted this month", 11, "money",
            formula='=IF({date}{r}="","",SUMIFS({value_first}:'
                    '{value_last},{month_first}:{month_last},'
                    '{month}{r}))', hidden=True),
        Col("alert", "Check", 22, "status",
            formula='=IF({date}{r}="","",IF({avoidable}{r}="Yes",'
                    'IF({value}{r}>5,"Check - avoidable and expensive",'
                    '"Check - avoidable"),"OK"))',
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Food Waste Log",
        subtitle="What got thrown away, what it cost and what would "
                 "have stopped it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent_soft,
        freeze_cols=2,
        samples=sample(cfg).waste,
        help_text="Log it for four weeks. Almost nobody keeps this up "
                  "forever, and almost everybody changes what they buy "
                  "after the first month.",
        extra_notes=[
            "The value is worked out from the price book, so a "
            "quarter of a bag of rocket costs what you paid for it, "
            "not a guess.",
            "The 'what would have stopped it' column is the one worth "
            "re-reading before you plan next week.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "What waste is costing you", last)
    row = s.table_head(row, 0, ["Items binned", "Total cost",
                                "This month", "Avoidable share",
                                "A year at this rate", "Target a month",
                                "Against target"],
                       [13, 13, 13, 15, 18, 15, 14])
    diff_row = row
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Food Waste.date]],">0")',
        '=SUM([[Food Waste.value]])',
        '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
        'TEXT(TODAY(),"yyyy-mm"))',
        '=IFERROR(SUMIFS([[Food Waste.value]],'
        '[[Food Waste.avoidable]],"Yes")/SUM([[Food Waste.value]]),"")',
        '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
        'TEXT(TODAY(),"yyyy-mm"))*12',
        "=WasteCostTarget",
        '=WasteCostTarget-SUMIFS([[Food Waste.value]],'
        '[[Food Waste.month]],TEXT(TODAY(),"yyyy-mm"))',
    ], ["int", "money", "money", "percent0", "money", "money", "money"])
    s.ws.conditional_format(diff_row, 6, diff_row, 6,
                            {"type": "cell", "criteria": "<", "value": 0,
                             "format": fmts.cf("bad")})
    s.ws.conditional_format(diff_row, 6, diff_row, 6,
                            {"type": "cell", "criteria": ">=", "value": 0,
                             "format": fmts.cf("good")})
````

### `meal_planner_spreadsheet/sheets/prep.py`

*151 lines · 2 functions*

> `Meal Prep` - the Sunday afternoon plan.
>
> Batch cooking only pays off if you know what to make, how many portions
> it gives you and where they are going. One row per prep job: what you
> are cooking, how long it takes, how many portions come out and which
> days of the week they cover.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Meal Prep').

````python
"""`Meal Prep` - the Sunday afternoon plan.

Batch cooking only pays off if you know what to make, how many portions
it gives you and where they are going. One row per prep job: what you
are cooking, how long it takes, how many portions come out and which
days of the week they cover.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, lookup_recipe, status_cf, weekday_col

NAME = "Meal Prep"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("job", '"PR-"', digits=3, width=8),
        Col("job", "Prep job", 26, "text",
            note="'Chop all the veg', 'cook a big pot of bolognese', "
                 "'portion the chicken for the freezer'."),
        Col("date", "Prep day", 11, "date", validate=date_check(),
            note="Your usual prep day is set on the Settings tab."),
        weekday_col("date"),
        Col("recipe", "Recipe", 11, "text", validate=list_from("RecipeIDs")),
        Col("recipe_name", "What you are cooking", 24, "text",
            formula=lookup_recipe("name")),
        Col("batches", "Batches", 9, "number1",
            validate=number_check(0, 50),
            note="How many times the recipe. Two batches of a recipe "
                 "that serves four is eight portions."),
        Col("servings_each", "Serves (each)", 10, "int",
            formula=lookup_recipe("servings"), hidden=True),
        Col("portions", "Portions made", 11, "number1",
            formula='=IF(OR({batches}{r}="",{servings_each}{r}="",'
                    'NOT(ISNUMBER({servings_each}{r}))),"",'
                    '{batches}{r}*{servings_each}{r})'),
        Col("minutes_each", "Minutes (recipe)", 10, "int",
            formula=lookup_recipe("total_min"), hidden=True),
        Col("minutes", "Time needed", 11, "int",
            validate=number_check(0, 1440, decimal=False),
            note="Your estimate. Blank uses the recipe's own prep plus "
                 "cook time."),
        Col("eff_minutes", "Minutes used", 10, "int",
            formula='=IF({job}{r}="","",IF({minutes}{r}<>"",'
                    '{minutes}{r},IF(ISNUMBER({minutes_each}{r}),'
                    '{minutes_each}{r},0)))', hidden=True),
        Col("per_portion", "Minutes a portion", 11, "number1",
            formula='=IF(OR({portions}{r}="",{portions}{r}=0,'
                    '{eff_minutes}{r}=0),"",{eff_minutes}{r}/'
                    '{portions}{r})',
            note="The number that proves batch cooking works - watch "
                 "it fall as the batches get bigger.",
            cf=[{"type": "3_color_scale", "min_color": "#D9F2E4",
                 "mid_color": "#FFF1CF", "max_color": "#FBD9DD"}]),
        Col("storage", "Stored in", 12, "text",
            validate=list_from("StorageList")),
        Col("containers", "Containers", 11, "int",
            validate=number_check(0, 100, decimal=False)),
        Col("covers_from", "Covers meals from", 11, "date",
            validate=date_check()),
        Col("covers_to", "Covers meals to", 11, "date",
            validate=date_check()),
        Col("meals_covered", "Meals it covers", 11, "int",
            formula='=IF(OR({covers_from}{r}="",{covers_to}{r}=""),"",'
                    'COUNTIFS([[Meal Plan.date]],">="&{covers_from}{r},'
                    '[[Meal Plan.date]],"<="&{covers_to}{r},'
                    '[[Meal Plan.recipe]],{recipe}{r}))',
            note="Planned meals in that window that use this recipe."),
        Col("cook", "Who is doing it", 14, "text",
            validate=list_from("MemberNames")),
        Col("cost", "Cost of the batch", 11, "money",
            formula='=IF(OR({batches}{r}="",{recipe}{r}=""),"",'
                    'IFERROR(INDEX([[Recipes.cost]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0))*{batches}{r},""))'),
        Col("cost_portion", "Cost a portion", 11, "money",
            formula='=IF(OR({cost}{r}="",{portions}{r}="",'
                    '{portions}{r}=0),"",{cost}{r}/{portions}{r})'),
        Col("todo_rank", "To-do rank", 9, "int",
            formula='=IF(OR({job}{r}="",{status}{r}="Done",'
                    'NOT(ISNUMBER({date}{r}))),"",COUNTIFS('
                    '{status_first}:{status_last},"<>Done",'
                    '{date_first}:{date_last},"<"&{date}{r})+'
                    'COUNTIFS({status_first}:{status}{r},"<>Done",'
                    '{date_first}:{date}{r},{date}{r}))',
            hidden=True,
            note="1 = the next unfinished prep job by date."),
        Col("status", "Status", 12, "status",
            validate=list_from("PrepStatusList"), cf=status_cf()),
        Col("alert", "Check", 24, "status",
            formula='=IF({job}{r}="","",'
                    'IF({status}{r}="Done","OK",'
                    'IF(AND({date}{r}<>"",{date}{r}<TODAY()),'
                    '"Overdue - still not made",'
                    'IF({date}{r}=TODAY(),"Today - prep day",'
                    'IF(AND({date}{r}<>"",{date}{r}-TODAY()<='
                    'PrepLeadDays),"Due - coming up","OK")))))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Meal Prep & Batch Cooking",
        subtitle="What you are making ahead, how many portions it "
                 "gives you and what each one costs in money and "
                 "minutes.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).prep,
        help_text="Plan two or three jobs for your prep day, not ten. "
                  "The minutes-a-portion column tells you which ones "
                  "are worth repeating.",
        extra_notes=[
            "Portions made feeds the leftovers tab: cook eight, eat "
            "four, and the other four are lunches with a value "
            "attached.",
            "Cost a portion is the number to compare with a shop "
            "bought lunch. It is usually the most convincing column "
            "in the workbook.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Batch cooking pays off", last)
    row = s.table_head(row, 0, ["Prep jobs", "Portions made",
                                "Hours in the kitchen",
                                "Average minutes a portion",
                                "Cost of the batches",
                                "Average cost a portion", "Still to make"],
                       [12, 14, 18, 22, 18, 20, 14])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Meal Prep.job]],"<>")',
        '=SUM([[Meal Prep.portions]])',
        '=SUM([[Meal Prep.eff_minutes]])/60',
        '=IFERROR(AVERAGE([[Meal Prep.per_portion]]),"")',
        '=SUM([[Meal Prep.cost]])',
        '=IFERROR(SUM([[Meal Prep.cost]])/SUM([[Meal Prep.portions]]),"")',
        '=COUNTIFS([[Meal Prep.status]],"<>Done",[[Meal Prep.job]],"<>")',
    ], ["int", "number1", "number1", "number1", "money", "money", "int"])
````

### `meal_planner_spreadsheet/sheets/spending.py`

*139 lines · 2 functions*

> `Grocery Spending` - what you actually spent.
>
> The planned cost on the meal plan is a forecast. This is the receipt.
> Log the shop (or each line of it if you are keeping a close eye on
> prices) and the budget tabs compare the two, month by month.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Grocery Spending').

````python
"""`Grocery Spending` - what you actually spent.

The planned cost on the meal plan is a forecast. This is the receipt.
Log the shop (or each line of it if you are keeping a close eye on
prices) and the budget tabs compare the two, month by month.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, id_col, month_col, status_cf, week_col,
                     weekday_col)

NAME = "Grocery Spending"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "SpendPrefix", digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        weekday_col("date"),
        week_col("date"),
        month_col("date"),
        Col("store", "Store", 18, "text", validate=list_from("StoreNames")),
        Col("type", "Type of spend", 15, "text",
            validate=list_from("SpendTypeList"),
            note="Weekly shop, top-up, takeaway, eating out, treat. "
                 "The budget splits on this."),
        Col("item", "Item or description", 24, "text",
            validate=list_from("IngredientNames"),
            note="Pick an ingredient to track its price over time, or "
                 "just type 'weekly shop'."),
        Col("qty", "Quantity", 9, "number1",
            validate=number_check(0, 100000)),
        Col("unit", "Unit", 8, "text", validate=list_from("UnitList")),
        Col("amount", "Amount", 11, "money",
            validate=number_check(0, 1000000),
            note="What left your account."),
        Col("unit_price", "Price per unit", 11, "money",
            formula='=IF(OR({amount}{r}="",{qty}{r}="",{qty}{r}=0),"",'
                    '{amount}{r}/{qty}{r})',
            hidden=True),
        Col("payer", "Who paid", 14, "text",
            validate=list_from("MemberNames")),
        Col("method", "Paid with", 13, "text",
            validate=list_from("PaymentList")),
        Col("planned", "Planned?", 9, "flag",
            validate=list_from("YesNoList"),
            note="No means it was an impulse buy - the dashboard "
                 "counts them."),
        Col("budget_line", "Budget line", 15, "text",
            validate=list_from("BudgetLineList")),
        Col("waste", "Thrown away?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Set to Yes if this one ended up in the bin - it "
                 "feeds the waste report."),
        Col("receipt", "Receipt link", 18, "link"),
        Col("month_budget", "That month's budget", 11, "money",
            formula="=MonthlyBudget", hidden=True),
        Col("running_month", "Spent this month", 12, "money",
            formula='=IF({date}{r}="","",SUMIFS({amount_first}:'
                    '{amount_last},{month_first}:{month_last},'
                    '{month}{r}))',
            note="Everything logged in the same month as this row."),
        Col("over", "Over budget?", 12, "status",
            formula='=IF({date}{r}="","",IF({running_month}{r}>'
                    'MonthlyBudget,"Check - over the monthly budget",'
                    '"OK"))',
            cf=alert_cf(), hidden=True),
        Col("status", "Status", 12, "status",
            formula='=IF({date}{r}="","",IF({amount}{r}="",'
                    '"Add - no amount",IF({waste}{r}="Yes","Wasted",'
                    'IF({planned}{r}="No","Check - unplanned",'
                    '"Purchased"))))',
            cf=status_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Grocery Spending",
        subtitle="Every shop, top-up and takeaway, so the budget tabs "
                 "can compare what you planned with what you paid.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).spending,
        help_text="One line per shop is plenty to track a budget. Line "
                  "by line entry is only worth it for the handful of "
                  "items whose price you want to watch.",
        extra_notes=[
            "Choosing an ingredient from the drop-down lets the price "
            "book show what you have actually been paying - useful "
            "when the shelf price and the real price drift apart.",
            "Takeaways and meals out belong here too, under their own "
            "spend type, otherwise the monthly food total lies to you.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Spending summary", last)
    row = s.table_head(row, 0, ["Shops logged", "Total spent",
                                "This month", "Monthly budget",
                                "Left this month", "Eating out",
                                "Unplanned buys", "Average shop"],
                       [13, 13, 13, 14, 15, 13, 14, 14])
    diff_row = row
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Grocery Spending.date]],">0")',
        '=SUM([[Grocery Spending.amount]])',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
        "=MonthlyBudget",
        '=MonthlyBudget-SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.type]],"Eating out")+'
        'SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.type]],"Takeaway")',
        '=COUNTIF([[Grocery Spending.planned]],"No")',
        '=IFERROR(AVERAGEIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.type]],"Weekly shop"),"")',
    ], ["int", "money", "money", "money", "money", "money", "int",
        "money"])
    s.ws.conditional_format(diff_row, 4, diff_row, 4,
                            {"type": "cell", "criteria": "<", "value": 0,
                             "format": fmts.cf("bad")})
    s.ws.conditional_format(diff_row, 4, diff_row, 4,
                            {"type": "cell", "criteria": ">=", "value": 0,
                             "format": fmts.cf("good")})
````

### `meal_planner_spreadsheet/sheets/store_prices.py`

*131 lines · 2 functions*

> `Store Prices` - the same thing, priced in three shops.
>
> The honest version of 'shop around'. Log what a pack costs where, and
> the workbook works out the price per unit, the cheapest shop, and
> whether the drive to the big supermarket is actually worth it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Store Prices').

````python
"""`Store Prices` - the same thing, priced in three shops.

The honest version of 'shop around'. Log what a pack costs where, and
the workbook works out the price per unit, the cheapest shop, and
whether the drive to the big supermarket is actually worth it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, id_col, lookup_ingredient,
                     min_number)

NAME = "Store Prices"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("ingredient", '"SP-"', digits=4, width=8),
        Col("ingredient", "Ingredient", 24, "text",
            validate=list_from("IngredientNames")),
        Col("store", "Store", 18, "text", validate=list_from("StoreNames")),
        Col("brand", "Brand", 16, "text",
            note="Own brand against the big name is usually the "
                 "biggest single saving in a shop."),
        Col("pack_size", "Pack size", 10, "number1",
            validate=number_check(0, 1000000),
            note="In the base unit for that ingredient - g, ml or "
                 "each."),
        Col("unit", "Unit", 8, "text",
            formula=lookup_ingredient("base_unit")),
        Col("price", "Pack price", 11, "money",
            validate=number_check(0, 100000)),
        Col("offer", "On offer?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Offer prices are worth logging but do not plan the "
                 "month around them."),
        Col("unit_cost", "Cost per unit", 11, "money",
            formula='=IF(OR({price}{r}="",{pack_size}{r}="",'
                    '{pack_size}{r}=0),"",{price}{r}/{pack_size}{r})',
            hidden=True),
        Col("per_100", "Cost per 100", 11, "money",
            formula='=IF({unit_cost}{r}="","",{unit_cost}{r}*100)',
            note="The comparable number. This is the column to sort."),
        Col("book_cost", "Price book cost", 11, "money",
            formula=lookup_ingredient("unit_cost"), hidden=True),
        Col("vs_book", "Against the price book", 11, "money",
            formula='=IF(OR({unit_cost}{r}="",{book_cost}{r}="",'
                    'NOT(ISNUMBER({book_cost}{r}))),"",'
                    '({unit_cost}{r}-{book_cost}{r})*100)',
            note="Negative means this shop is cheaper per 100 than the "
                 "price you have in the Ingredients tab.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("best", "Cheapest so far?", 11, "flag",
            formula='=IF(OR({ingredient}{r}="",{per_100}{r}=""),"",'
                    'IF({unit_cost}{r}<=' + min_number(
                        '[[Store Prices.unit_cost]]',
                        '[[Store Prices.ingredient]]={ingredient}{r}')
                    + '+0.0000001,"Yes","No"))',
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"}],
            note="Yes means no other shop you have logged beats this "
                 "price per unit."),
        Col("checked", "Price checked", 11, "date", validate=date_check()),
        Col("age_days", "How old", 9, "int",
            formula='=IF({checked}{r}="","",TODAY()-{checked}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 90,
                 "fmt": "warn"}]),
        Col("alert", "Check", 24, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({price}{r}="","Add - no price",'
                    'IF({pack_size}{r}="","Add - no pack size",'
                    'IF(AND({age_days}{r}<>"",{age_days}{r}>180),'
                    '"Check - over six months old","OK"))))',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Store Price Comparison",
        subtitle="What the same ingredient costs in each shop, reduced "
                 "to one comparable number.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent_soft,
        freeze_cols=3,
        samples=sample(cfg).store_prices,
        help_text="You do not need to price everything. The twenty "
                  "things you buy every week are where the money is.",
        extra_notes=[
            "Cost per 100 is the only fair comparison - it is what "
            "the shelf label shows, and it is how the Ingredients tab "
            "picks the cheapest shop for each item.",
            "Prices go stale. Anything you checked more than six "
            "months ago is flagged so you know not to trust it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "What shopping around is worth", last)
    row = s.table_head(row, 0, ["Prices logged", "Ingredients covered",
                                "Stores compared", "Checked this year",
                                "Total saving if you always bought "
                                "cheapest"],
                       [14, 18, 16, 16, 34])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Store Prices.ingredient]],"<>")',
        '=SUMPRODUCT((COUNTIF([[Store Prices.ingredient]],'
        '[[Ingredients.name]])>0)*([[Ingredients.name]]<>""))',
        '=SUMPRODUCT((COUNTIF([[Store Prices.store]],'
        '[[Stores.name]])>0)*([[Stores.name]]<>""))',
        '=COUNTIFS([[Store Prices.checked]],">="&DATE(YEAR(TODAY()),1,1))',
        '=SUM([[Ingredients.saving]])',
    ], ["int", "int", "int", "int", "money"])
    row += 1
    s.note(row, "The saving on the right is the sum of every 'cheaper "
                "elsewhere' gap per 100 units. It is a ceiling, not a "
                "promise - petrol, time and the fact that nobody "
                "shops in four places all take a bite out of it.", last)
````

### `meal_planner_spreadsheet/sheets/stores.py`

*95 lines · 1 function*

> `Stores` - where you shop.
>
> Opening hours, delivery slots, loyalty cards and - the useful bit - how
> much of your money each shop takes and how its prices compare.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Stores').

````python
"""`Stores` - where you shop.

Opening hours, delivery slots, loyalty cards and - the useful bit - how
much of your money each shop takes and how its prices compare.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, id_col, status_cf

NAME = "Stores"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"ST-"', digits=2, width=8),
        Col("name", "Store", 20, "text",
            note="Supermarkets, the market stall, the butcher, the "
                 "online shop - anywhere food money goes."),
        Col("type", "Type", 15, "text", validate=list_from("StoreTypeList")),
        Col("location", "Where", 20, "text"),
        Col("travel_min", "Minutes away", 10, "int",
            validate=number_check(0, 300, decimal=False)),
        Col("best_for", "Best for", 24, "text",
            note="'Veg and bread', 'cheap store cupboard', 'the only "
                 "place that does the oat milk'."),
        Col("delivery", "Delivers?", 10, "flag",
            validate=list_from("YesNoList")),
        Col("delivery_fee", "Delivery fee", 11, "money",
            validate=number_check(0, 1000)),
        Col("min_order", "Minimum order", 11, "money",
            validate=number_check(0, 100000)),
        Col("slot", "Usual slot", 16, "text",
            note="'Thursday 6-7pm'. Handy next to the shopping day."),
        Col("loyalty", "Loyalty card", 16, "text"),
        Col("open_hours", "Opening hours", 18, "text"),
        Col("price_level", "Price level", 12, "text",
            validate=list_from("PriceLevelList")),
        Col("rating", "Your rating", 10, "rating",
            validate=number_check(1, 5, decimal=False)),
        Col("items_priced", "Items priced here", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Store Prices.store]],{name}{r}))',
            note="How many lines of the price book you have checked "
                 "at this shop."),
        Col("cheapest_on", "Cheapest on", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Ingredients.best_store]],{name}{r}))',
            note="How many ingredients this shop wins on price."),
        Col("spent", "Spent here", 12, "money",
            formula='=IF({name}{r}="","",SUMIFS('
                    '[[Grocery Spending.amount]],'
                    '[[Grocery Spending.store]],{name}{r}))'),
        Col("visits", "Shops logged", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Grocery Spending.store]],{name}{r}))'),
        Col("avg_shop", "Average shop", 11, "money",
            formula='=IF(OR({visits}{r}="",{visits}{r}=0),"",'
                    '{spent}{r}/{visits}{r})'),
        Col("on_list", "On this week's list", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Grocery List.store]],{name}{r}))'),
        Col("spend_rank", "Spend rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({spent_first}:'
                    '{spent_last},">"&{spent}{r})+COUNTIFS('
                    '{spent_first}:{spent}{r},{spent}{r}))',
            hidden=True, note="1 = where most of the money goes."),
        Col("status", "Status", 12, "status",
            validate=list_from("ActiveList"), cf=status_cf()),
        Col("alert", "Check", 22, "status",
            formula='=IF({name}{r}="","",IF({items_priced}{r}=0,'
                    '"Add - no prices logged here yet",'
                    'IF(AND({delivery}{r}="Yes",{min_order}{r}=""),'
                    '"Add - minimum order?","OK")))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Grocery Stores",
        subtitle="Where you shop, what each one is good for and what "
                 "each one costs you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_soft,
        freeze_cols=2,
        samples=sample(cfg).stores,
        help_text="Add a store here and it appears in every store "
                  "drop-down in the workbook.",
    )
````

### `meal_planner_spreadsheet/sheets/events.py`

*113 lines · 1 function*

> `Events` - birthdays, Christmas, the in-laws on Sunday.
>
> Occasions need a head count, a menu, a budget and a cooking timetable.
> This tab holds the occasion; the meals themselves go on the meal plan
> with the event picked in the Occasion column, so the shopping list
> picks them up like any other meal.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Events').

````python
"""`Events` - birthdays, Christmas, the in-laws on Sunday.

Occasions need a head count, a menu, a budget and a cooking timetable.
This tab holds the occasion; the meals themselves go on the meal plan
with the event picked in the Occasion column, so the shopping list
picks them up like any other meal.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, status_cf

NAME = "Events"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "EventPrefix", digits=3, width=9),
        Col("name", "Occasion", 24, "text",
            note="'Dad's birthday dinner', 'Christmas Day', 'book "
                 "club'. The name appears in the meal plan drop-down."),
        Col("type", "Type", 14, "text", validate=list_from("EventTypeList")),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("days_to_go", "Days to go", 10, "int",
            formula='=IF({date}{r}="","",{date}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "muted"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 7, "fmt": "warn"}]),
        Col("time", "Time", 10, "time"),
        Col("guests", "Guests", 9, "int",
            validate=number_check(0, 500, decimal=False),
            note="Including the household - it is the number you cook "
                 "for."),
        Col("adults", "Adults", 8, "int",
            validate=number_check(0, 500, decimal=False)),
        Col("children", "Children", 9, "int",
            validate=number_check(0, 500, decimal=False)),
        Col("portions", "Portions to cook", 11, "number1",
            formula='=IF({guests}{r}="","",IF(AND({adults}{r}="",'
                    '{children}{r}=""),{guests}{r},'
                    'IF({adults}{r}="",0,{adults}{r})+'
                    'IF({children}{r}="",0,{children}{r})*0.5))',
            note="Children count as half a portion - change the "
                 "formula if your children eat like mine."),
        Col("dietary", "Dietary needs", 24, "text",
            note="'One vegan, one coeliac, no nuts at all'. Check it "
                 "against the Allergens tab before you shop."),
        Col("budget", "Budget", 11, "money",
            validate=number_check(0, 100000)),
        Col("planned_cost", "Planned food cost", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Meal Plan.cost]],'
                    '[[Meal Plan.event]],{name}{r}))',
            note="From every meal on the plan tagged with this "
                 "occasion."),
        Col("actual_cost", "Actual cost", 11, "money",
            validate=number_check(0, 100000)),
        Col("vs_budget", "Against budget", 12, "money",
            formula='=IF({budget}{r}="","",{budget}{r}-'
                    'IF({actual_cost}{r}<>"",{actual_cost}{r},'
                    'IF({planned_cost}{r}="",0,{planned_cost}{r})))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"}]),
        Col("meals", "Meals planned", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Meal Plan.event]],{name}{r}))'),
        Col("menu", "Menu", 30, "wrap",
            note="The running order: nibbles, starter, main, pudding."),
        Col("prep_start", "Start cooking", 11, "date",
            validate=date_check(),
            note="Work backwards from the meal - most of the stress "
                 "is in the timetable, not the food."),
        Col("shopping_done", "Shopping done?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("host", "Who is hosting", 14, "text",
            validate=list_from("MemberNames")),
        Col("status", "Status", 12, "status",
            validate=list_from("EventStatusList"), cf=status_cf()),
        Col("alert", "Check", 26, "status",
            formula='=IF({name}{r}="","",'
                    'IF({status}{r}="Done","OK",'
                    'IF(AND({days_to_go}{r}<>"",{days_to_go}{r}<0),'
                    '"Check - date has passed",'
                    'IF(AND({days_to_go}{r}<=7,{meals}{r}=0),'
                    '"Action - nothing on the menu yet",'
                    'IF(AND({days_to_go}{r}<=3,{shopping_done}{r}<>'
                    '"Yes"),"Due - shopping not done","OK")))))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap",
            note="What worked, what to never attempt again."),
    ]

    return SheetSpec(
        name=NAME,
        title="Special Events",
        subtitle="Occasions with a head count, a budget and a "
                 "timetable - and a menu that flows into the normal "
                 "shopping list.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).events,
        help_text="Add the occasion here, then add its meals to the "
                  "Meal Plan with this occasion chosen - that is what "
                  "gets the ingredients onto the grocery list.",
    )
````

### `meal_planner_spreadsheet/sheets/goals.py`

*153 lines · 3 functions*

> `Goals` - what you are trying to change, and whether it is working.
>
> Cook at home five nights a week. Spend less than £120. Try two new
> recipes a month. Waste under a fiver. Each row is a goal with a target,
> a live measurement pulled from the rest of the workbook and a simple
> on-track flag.

| Defines | What it does |
|---------|--------------|
| `_measure_formula()` | Nested IF that turns the chosen measure into a live number |
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Goals'), `MEASURES` (8 items).

````python
"""`Goals` - what you are trying to change, and whether it is working.

Cook at home five nights a week. Spend less than £120. Try two new
recipes a month. Waste under a fiver. Each row is a goal with a target,
a live measurement pulled from the rest of the workbook and a simple
on-track flag.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, status_cf

NAME = "Goals"

#: the measures the 'Measured by' drop-down can read automatically
MEASURES = [
    ("Home cooked meals a week",
     'COUNTIFS([[Meal Plan.date]],">="&TODAY()-7,[[Meal Plan.date]],'
     '"<="&TODAY(),[[Meal Plan.home_cooked]],"Yes")'),
    ("Takeaways a month",
     'COUNTIFS([[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"),'
     '[[Grocery Spending.type]],"Takeaway")'),
    ("Grocery spend this month",
     'SUMIFS([[Grocery Spending.amount]],[[Grocery Spending.month]],'
     'TEXT(TODAY(),"yyyy-mm"))'),
    ("New recipes this month",
     'COUNTIFS([[Recipes.added]],">="&DATE(YEAR(TODAY()),'
     'MONTH(TODAY()),1))'),
    ("Food waste this month",
     'SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
     'TEXT(TODAY(),"yyyy-mm"))'),
    ("Packed lunches a week",
     'COUNTIFS([[Meal Plan.date]],">="&TODAY()-7,[[Meal Plan.slot]],'
     '"Lunch",[[Meal Plan.home_cooked]],"Yes")'),
    ("Leftovers eaten",
     'COUNTIFS([[Leftovers.status]],"Eaten")'),
    ("Meals prepped ahead",
     'COUNTIFS([[Meal Prep.status]],"Done")'),
]


def _measure_formula() -> str:
    """Nested IF that turns the chosen measure into a live number."""
    body = '""'
    for label, expr in reversed(MEASURES):
        body = f'IF({{measure}}{{r}}="{label}",{expr},{body})'
    return f'=IF({{measure}}{{r}}="","",IFERROR({body},""))'


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("goal", '"GL-"', digits=3, width=8),
        Col("goal", "Goal", 30, "text",
            note="Write it as something you would actually say: 'stop "
                 "ordering pizza on Fridays'."),
        Col("area", "Area", 14, "text", validate=list_from("GoalCategoryList")),
        Col("measure", "Measured by", 24, "text",
            validate=list_from("MeasureList"),
            note="Pick one of the measures the workbook can count for "
                 "itself, or choose 'Your own count' and fill in the "
                 "actual column by hand."),
        Col("target", "Target", 10, "number1",
            validate=number_check(0, 1000000)),
        Col("direction", "Aim to", 12, "text",
            validate=list_from("DirectionList"),
            note="'Hit or beat' for things you want more of, 'Stay "
                 "under' for spending and waste."),
        Col("auto", "Measured now", 11, "number1",
            formula=_measure_formula(), hidden=True),
        Col("manual", "Your own count", 11, "number1",
            validate=number_check(0, 1000000),
            note="Only needed for goals the workbook cannot count."),
        Col("actual", "Where you are", 11, "number1",
            formula='=IF({manual}{r}<>"",{manual}{r},{auto}{r})'),
        Col("progress", "Progress", 10, "percent0",
            formula='=IF(OR({target}{r}="",{target}{r}=0,'
                    '{actual}{r}=""),"",IF({direction}{r}='
                    '"Stay under",IF({actual}{r}=0,1,'
                    'MIN(2,{target}{r}/{actual}{r})),'
                    'MIN(2,{actual}{r}/{target}{r})))',
            cf=[{"type": "data_bar", "bar_color": "#4E8A63"}]),
        Col("on_track", "On track?", 11, "flag",
            formula='=IF({progress}{r}="","",IF({progress}{r}>=1,'
                    '"Yes","No"))',
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"}]),
        Col("why", "Why it matters", 26, "wrap",
            note="The reason you will remember in week three."),
        Col("start", "Started", 11, "date", validate=date_check()),
        Col("review", "Review on", 11, "date", validate=date_check()),
        Col("owner", "Whose goal", 14, "text",
            validate=list_from("MemberNames")),
        Col("status", "Status", 12, "status",
            validate=list_from("GoalStatusList"), cf=status_cf()),
        Col("alert", "Check", 24, "status",
            formula='=IF({goal}{r}="","",'
                    'IF({status}{r}="Done","OK",'
                    'IF({target}{r}="","Add - no target set",'
                    'IF(AND({review}{r}<>"",{review}{r}<TODAY()),'
                    '"Due - time to review this",'
                    'IF({on_track}{r}="No","Check - behind target",'
                    '"OK")))))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Food Goals & Habits",
        subtitle="Three or four goals, measured by the numbers already "
                 "in this workbook.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).goals,
        help_text="Three goals you will keep beats ten you will not. "
                  "The 'Measured now' column updates itself every time "
                  "the workbook recalculates.",
        extra_notes=[
            "Progress is capped at 200% so one brilliant week does "
            "not make the bar meaningless.",
            "'Stay under' goals invert the sum: hitting £90 against a "
            "£120 target reads as ahead, not behind.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "How the habits are going", last)
    row = s.table_head(row, 0, ["Goals set", "On track", "Behind",
                                "Average progress", "Due a review",
                                "Finished"],
                       [12, 11, 11, 16, 14, 11])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Goals.goal]],"<>")',
        '=COUNTIF([[Goals.on_track]],"Yes")',
        '=COUNTIF([[Goals.on_track]],"No")',
        '=IFERROR(AVERAGE([[Goals.progress]]),"")',
        '=COUNTIFS([[Goals.review]],"<"&TODAY(),[[Goals.review]],">0",'
        '[[Goals.status]],"<>Done")',
        '=COUNTIF([[Goals.status]],"Done")',
    ], ["int", "int", "int", "percent0", "int", "int"])
````

---

## 5. The pages

The tabs that read the registers: the welcome page, the dashboard, the two planners, the filters, the money and nutrition pages, the converter and scaler, rotation, food safety, cold storage, sustainability, the reports, the three printables and the help tab.

### `meal_planner_spreadsheet/sheets/start_here.py`

*189 lines · 1 function*

> `Start Here` - the welcome page and the six step set-up.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Start Here'), `LAST_COL` (11), `STEPS` (6 items), `TABS` (22 items).

````python
"""`Start Here` - the welcome page and the six step set-up."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Start Here"
LAST_COL = 11

STEPS = [
    ("1", "Settings",
     "Your household, how many you feed, the weekly budget, your "
     "shopping day and your nutrition targets. Ten minutes, and it "
     "feeds every other tab."),
    ("2", "Household",
     "One row per person: portion size, diet, allergies, what they "
     "will not touch and which meals they are in for."),
    ("3", "Ingredients",
     "The price book. Pack size and pack price for the thirty things "
     "you buy most. This is what prices every recipe and every "
     "shopping list, so it is the tab worth the effort."),
    ("4", "Recipes + Recipe Ingredients",
     "Add a recipe, then list its ingredients one line each with a "
     "quantity and a unit. The cost, the calories and the allergens "
     "are worked out for you."),
    ("5", "Pantry",
     "What is already in the cupboard, the fridge and the freezer. "
     "This is what stops the grocery list buying things you own."),
    ("6", "Meal Plan",
     "Plan the week: date, meal, recipe, how many you are feeding. "
     "The Weekly Planner shows it as a week and the Grocery List "
     "builds itself."),
]

TABS = [
    ("Dashboard", "This week at a glance - cost, meals, alerts"),
    ("Weekly Planner", "Seven days on one page, ready to print"),
    ("Monthly Planner", "The month as a calendar grid"),
    ("Settings", "Household, budget, targets, shopping window"),
    ("Lists", "Every drop-down, plus the unit conversion table"),
    ("Household", "Who eats here, portions, diets and allergies"),
    ("Recipes", "THE RECIPE BOOK - cost and nutrition worked out"),
    ("Recipe Ingredients", "What goes in each recipe. The engine room"),
    ("Ingredients", "The price book: pack sizes, prices, nutrition"),
    ("Meal Plan", "THE MAIN TAB - one row per planned meal"),
    ("Grocery List", "Builds itself from the plan, minus the pantry"),
    ("Pantry", "What you own, what it is worth, what expires when"),
    ("Staples", "Par levels for the things you never want to run out of"),
    ("Leftovers", "What is in the fridge and how long you have"),
    ("Food Waste", "What got binned and what it cost"),
    ("Meal Prep", "Batch cooking: portions made, minutes per portion"),
    ("Grocery Spending", "Every shop and takeaway against the budget"),
    ("Store Prices", "The same item priced in three shops"),
    ("Stores", "Where you shop, hours, delivery, loyalty"),
    ("Events", "Birthdays and occasions with a head count and budget"),
    ("Goals", "Three or four food habits, measured automatically"),
    ("Help", "Every question answered, including the honest limits"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)

    s.banner(cfg.product_name,
             "Plan the week, price it before you shop, and stop buying "
             "things you already own.", LAST_COL)
    s.set_col(0, 0, 4)
    s.set_col(1, 1, 22)
    s.set_col(2, 2, 30)
    for c in range(3, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.note(
        row,
        "Everything in this workbook is linked. Type something once - a "
        "recipe, a price, a meal - and it turns up everywhere it is "
        "needed: the planner, the shopping list, the budget, the "
        "nutrition tracker and the alerts. Cream cells are yours to "
        "fill in; grey cells are formulas and are locked so they "
        "cannot be broken by accident.",
        LAST_COL, first_col=1, height=46)
    row += 1

    row = s.section(row, "Six steps to set it up", LAST_COL)
    s.table_head(row, 1, ["Step", "Where", "What to do"], [8, 26, 68])
    first = row + 1
    for i, (number, where, what) in enumerate(STEPS):
        r = first + i
        s.row_height(r, 32)
        s.write(r, 1, number, fmts.chrome("kpi_label"))
        s.write(r, 2, where, fmts.chrome("label_bold"))
        s.merge(r, 3, LAST_COL, what, fmts.chrome("note"))
    row = first + len(STEPS) + 1

    row = s.section(row, "How the automatic grocery list works", LAST_COL)
    row = s.bullets(row, [
        "1. The Meal Plan says you are cooking recipe R-0002 on Tuesday "
        "for 4 people.",
        "2. The recipe serves 6, so everything in it is multiplied by "
        "0.67 - that is the scaling, done for you.",
        "3. Each ingredient line is converted to a base unit: 2 tbsp "
        "of oil becomes 30 ml, 1 kg of mince becomes 1000 g.",
        "4. Every meal in the shopping window is added up by "
        "ingredient, so three recipes using onions make one line.",
        "5. Whatever the Pantry says you already own is taken off.",
        "6. What is left is rounded up to whole packs and priced from "
        "the Ingredients tab - that is your list, with a total, before "
        "you leave the house.",
    ], LAST_COL, first_col=1, height=26)
    row += 1

    row = s.section(row, "The weekly rhythm", LAST_COL)
    row = s.bullets(row, [
        "FRIDAY (10 minutes) - look at the Pantry 'Use soon' list and "
        "the Leftovers tab, then plan next week on the Meal Plan tab.",
        "SATURDAY MORNING - open the Grocery List, tick off anything "
        "you already have, print it or take the laptop, and shop.",
        "AFTER THE SHOP (2 minutes) - log the total on the Grocery "
        "Spending tab and update the Pantry for the big items.",
        "PREP DAY - work through the Meal Prep tab: two or three jobs, "
        "not ten.",
        "DURING THE WEEK - mark meals 'Eaten' as you go and log "
        "leftovers while you are putting them in the fridge.",
        "END OF THE MONTH - read the Dashboard and the Food Waste tab, "
        "and change one thing.",
    ], LAST_COL, first_col=1, height=26)
    row += 1

    row = s.section(row, "What is in the workbook", LAST_COL)
    s.table_head(row, 1, ["Tab", "What it is for"], [26, 70])
    first = row + 1
    for i, (tab, purpose) in enumerate(TABS):
        r = first + i
        s.write(r, 1, tab, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL, purpose, fmts.chrome("label"))
    row = first + len(TABS) + 1

    row = s.section(row, "Good to know", LAST_COL)
    row = s.bullets(row, [
        f"PROTECTION - every sheet is protected so the formulas cannot "
        f"be edited by accident. The password is '{cfg.password}'. "
        f"Review > Unprotect Sheet if you want to change something.",
        "ADDING ROWS - every register has spare rows ready to use. "
        "Right click a row inside the list and choose Insert to add "
        "more; the formulas copy down.",
        "SORTING - use the filter arrows in the header row rather than "
        "re-sorting the sheet, so the recipe IDs keep matching the "
        "other tabs.",
        "DROP DOWNS - every list lives on the Lists tab, including the "
        "unit table. Add a unit there and it works everywhere.",
        "UNITS - weights convert to grams, volumes to millilitres, and "
        "counted things stay as 'each'. Weight and volume are "
        "deliberately kept apart because 100 ml of oil is not 100 g.",
        "GOOGLE SHEETS - File > Import works, and everything except "
        "the sheet protection carries over.",
        "YOUR CURRENCY - set the symbol on the Settings tab, or "
        "rebuild the workbook with a different currency.",
    ], LAST_COL, first_col=1, height=28)
    row += 1

    row = s.section(row, "What a spreadsheet can and cannot do", LAST_COL)
    row = s.note(
        row,
        "Being straight with you. This workbook does the arithmetic on "
        "the information you give it: it cannot read the barcodes in "
        "your cupboard, it cannot check tonight's shelf prices, and it "
        "will not order the shopping for you. The automatic grocery "
        "list is only as good as your ingredient lines and units - a "
        "recipe that says 'a splash of oil' cannot be added up. "
        "Nutrition comes from the numbers you copy off the packet and "
        "is a guide, not medical advice. Allergen flags and use-by "
        "dates are a second pair of eyes, not a guarantee: always read "
        "the label and trust your nose. Reminders appear on the tabs - "
        "a spreadsheet cannot push a notification to your phone. What "
        "it does do is turn a vague plan into a priced shopping list "
        "in ten minutes a week, and show you exactly where the food "
        "money goes.",
        LAST_COL, first_col=1, height=92)
    row += 1

    s.merge(row, 1, LAST_COL,
            f"{cfg.product_name}  |  {cfg.author}  |  version "
            f"{cfg.version}", fmts.chrome("print_sub"))

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/dashboard.py`

*292 lines · 1 function*

> `Dashboard` - the one page that answers 'are we alright?'
>
> Eight numbers across the top, a list of things that need doing, this
> week's dinners, and twelve months of spending, waste and home cooking
> with a little column chart beside each row. Everything reads straight
> from the registers, so it is never out of date.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dashboard'), `LAST_COL` (11), `MONTHS` (12).

````python
"""`Dashboard` - the one page that answers 'are we alright?'

Eight numbers across the top, a list of things that need doing, this
week's dinners, and twelve months of spending, waste and home cooking
with a little column chart beside each row. Everything reads straight
from the registers, so it is never out of date.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Dashboard"
LAST_COL = 11
MONTHS = 12


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Dashboard",
             "This week, this month and the year so far - all from the "
             "numbers you have already entered.", LAST_COL)
    s.set_col(0, 0, 20)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Week of", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=WeekPick", fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("DashWeek", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "Month", fmts.chrome("label_right"))
    ws.write_formula(row, 3, '=TEXT(TODAY(),"mmmm yyyy")',
                     fmts.calc("text"))
    book.define_name("DashMonth", f"={quote_sheet(NAME)}!$D${row + 1}")
    s.merge(row, 4, 4, "Month key", fmts.chrome("label_right"))
    ws.write_formula(row, 5, '=TEXT(TODAY(),"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$F${row + 1}")
    s.merge(row, 6, 6, "Household", fmts.chrome("label_right"))
    ws.write_formula(row, 7, "=HouseMembers", fmts.calc("int"))
    s.merge(row, 8, 9, "Updated", fmts.chrome("label_right"))
    ws.write_formula(row, 10, "=TODAY()", fmts.calc("date"))
    row += 2

    # ---------------------------------------------------------- the KPIs --
    tiles_1 = [
        ("Meals planned this week",
         '=COUNTIFS([[Meal Plan.date]],">="&DashWeek,'
         '[[Meal Plan.date]],"<="&DashWeek+6)', "int",
         "every slot with something in it"),
        ("Planned cost this week",
         '=SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],">="&DashWeek,'
         '[[Meal Plan.date]],"<="&DashWeek+6)', "money",
         "at price book prices"),
        ("This week's shopping list",
         '=SUM([[Grocery List.est_cost]])', "money",
         "what the list adds up to"),
        ("Left in the weekly budget",
         '=WeeklyBudget-SUM([[Grocery List.est_cost]])', "money",
         "green is money left"),
    ]
    tiles_2 = [
        ("Spent this month",
         '=SUMIFS([[Grocery Spending.amount]],'
         '[[Grocery Spending.month]],DashKey)', "money",
         "every shop, takeaway and meal out"),
        ("Pantry value",
         '=SUM([[Pantry.value]])', "money",
         "what is sitting in the cupboards"),
        ("Food wasted this month",
         '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],DashKey)',
         "money", f"target under {cfg.currency_symbol}"
                  f"{cfg.waste_cost_target:,.0f}"),
        ("Cooked at home this week",
         '=COUNTIFS([[Meal Plan.date]],">="&DashWeek,'
         '[[Meal Plan.date]],"<="&DashWeek+6,'
         '[[Meal Plan.home_cooked]],"Yes")', "int",
         f"target {cfg.home_cooked_target} a week"),
    ]
    for tiles in (tiles_1, tiles_2):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, s.reg.render(formula), style=style,
                  span=3, note=note)
        row += 4

    # ----------------------------------------------- needs your attention
    row = s.section(row, "Needs your attention", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [26, 11, 46, 20])
    alerts = [
        ("Food about to go off",
         '=COUNTIF([[Pantry.status]],"Use soon")',
         f'="Use these within {cfg.expiry_warning_days} days or they '
         f'are money in the bin"', "Pantry"),
        ("Already expired",
         '=COUNTIF([[Pantry.status]],"Expired")',
         '="Past their date - check, then bin and log them"',
         "Pantry"),
        ("Leftovers waiting",
         '=SUMIFS([[Leftovers.portions]],[[Leftovers.status]],'
         '"<>Eaten",[[Leftovers.status]],"<>Binned")',
         '="Portions in the fridge or freezer worth "&'
         'TEXT(SUMIFS([[Leftovers.value]],[[Leftovers.status]],'
         '"<>Eaten",[[Leftovers.status]],"<>Binned"),"0.00")',
         "Leftovers"),
        ("Staples below par",
         '=COUNTIF([[Staples.status]],"Low stock")+'
         'COUNTIF([[Staples.status]],"Out of stock")',
         '="Things you said you never want to run out of"',
         "Staples"),
        ("Ingredients with no price",
         '=COUNTIFS([[Ingredients.name]],"<>",'
         '[[Ingredients.pack_price]],"")',
         '="Until these have a price, recipes using them are '
         'under-costed"', "Ingredients"),
        ("Recipe lines to fix",
         '=COUNTIF([[Recipe Ingredients.alert]],"Missing*")+'
         'COUNTIF([[Recipe Ingredients.alert]],"Check - *")',
         '="Unit mismatches or ingredients not in the price book - '
         'they are left out of the shopping list"',
         "Recipe Ingredients"),
        ("Days with nothing planned",
         "=" + "+".join(
             f'(COUNTIFS([[Meal Plan.date]],DashWeek+{i},'
             f'[[Meal Plan.slot]],"Dinner")=0)*1'
             for i in range(7)),
         '="Dinner slots still empty in the week you are looking at"',
         "Meal Plan"),
        ("Prep jobs outstanding",
         '=COUNTIFS([[Meal Prep.status]],"<>Done",'
         '[[Meal Prep.job]],"<>")',
         '="Batch cooking you planned but have not made yet"',
         "Meal Prep"),
        ("Goals behind target",
         '=COUNTIF([[Goals.on_track]],"No")',
         '="Worth five minutes with a cup of tea"', "Goals"),
    ]
    first_alert = row
    for label, count, message, where in alerts:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(count), fmts.calc("int"))
        s.merge(row, 2, LAST_COL - 2, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(message), fmts.calc("text"))
        s.merge(row, LAST_COL - 1, LAST_COL, where, fmts.chrome("label"))
        row += 1
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": "=", "value": 0,
                           "format": fmts.cf("good")})
    row += 1

    # --------------------------------------------------- this week's food
    row = s.section(row, "This week's dinners", LAST_COL)
    s.write(row, 0, "Day", fmts.chrome("header_left"))
    for i in range(7):
        ws.write_formula(row, i + 1, f"=DashWeek+{i}",
                         fmts.chrome("header", num_format="ddd d"))
    s.write(row, 8, "Week", fmts.chrome("header"))
    day_row = row
    row += 1
    s.write(row, 0, "Dinner", fmts.chrome("label_bold"))
    for i in range(7):
        cell = f"{chr(66 + i)}${day_row + 1}"
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IFERROR(INDEX([[Meal Plan.recipe_name]],'
                         f'MATCH(TEXT({cell},"yyyy-mm-dd")&"|Dinner",'
                         f'[[Meal Plan.key]],0)),"-")'),
            fmts.cell("wrap", "calc"))
    ws.write_formula(row, 8,
                     s.reg.render('=COUNTIFS([[Meal Plan.date]],'
                                  '">="&DashWeek,[[Meal Plan.date]],'
                                  '"<="&DashWeek+6,[[Meal Plan.slot]],'
                                  '"Dinner")'), fmts.calc("int"))
    s.row_height(row, 30)
    row += 1
    s.write(row, 0, "Cost", fmts.chrome("label_bold"))
    for i in range(7):
        cell = f"{chr(66 + i)}${day_row + 1}"
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IF({cell}="","",SUMIFS([[Meal Plan.cost]],'
                         f'[[Meal Plan.date]],{cell}))'),
            fmts.calc("money"))
    ws.write_formula(row, 8,
                     s.reg.render('=SUMIFS([[Meal Plan.cost]],'
                                  '[[Meal Plan.date]],">="&DashWeek,'
                                  '[[Meal Plan.date]],"<="&DashWeek+6)'),
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    # ----------------------------------------------------- the year table
    row = s.section(row, "The year so far", LAST_COL)
    row = s.table_head(row, 0, ["Month", "Grocery spend", "Budget",
                                "Difference", "Eating out", "Food waste",
                                "Meals cooked at home", "Trend"],
                       [12, 14, 12, 12, 12, 12, 20, 16])
    first_month = row
    for m in range(MONTHS):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        ws.write_formula(row, 0,
                         f'=TEXT(DATE(PlanYear,{m + 1},1),"mmm")',
                         fmts.calc("text"))
        s.calc_row(row, 1, [
            f'=SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{key})',
            "=MonthlyBudget",
            f'=MonthlyBudget-SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{key})',
            f'=SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{key},'
            f'[[Grocery Spending.type]],"Eating out")+'
            f'SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{key},'
            f'[[Grocery Spending.type]],"Takeaway")',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.month]],{key})',
            f'=COUNTIFS([[Meal Plan.month]],{key},'
            f'[[Meal Plan.home_cooked]],"Yes")',
        ], ["money", "money", "money", "money", "money", "int"])
        row += 1
    last_month = row - 1
    ws.conditional_format(first_month, 3, last_month, 3,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_month, 1, last_month, 1,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    ws.conditional_format(first_month, 5, last_month, 5,
                          {"type": "data_bar", "bar_color": "#D9703C"})

    # sparklines beside the table
    for col_idx, (src, colour) in enumerate(
            [(1, "#2F5E43"), (5, "#D9703C"), (6, "#4E8A63")]):
        letter = chr(65 + src)
        try:
            ws.add_sparkline(first_month + col_idx, 7, {
                "range": (f"{quote_sheet(NAME)}!{letter}"
                          f"{first_month + 1}:{letter}{last_month + 1}"),
                "type": "column",
                "series_color": colour,
            })
        except Exception:        # pragma: no cover - library differences
            pass
    row += 1

    s.merge(row, 0, LAST_COL,
            "Spend, waste and home cooking by month. The little charts "
            "on the right are, top to bottom: grocery spend, food "
            "waste, meals cooked at home.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------------------- the totals
    row = s.section(row, "The year in one line", LAST_COL)
    row = s.table_head(row, 0, ["Spent on food", "Food budget",
                                "Difference", "Wasted", "Waste as a "
                                "share", "Meals cooked at home",
                                "Recipes in the book", "Pantry value"],
                       [14, 13, 12, 11, 15, 20, 17, 13])
    s.calc_row(row, 0, [
        '=SUM([[Grocery Spending.amount]])',
        "=AnnualBudget",
        '=AnnualBudget-SUM([[Grocery Spending.amount]])',
        '=SUM([[Food Waste.value]])',
        '=IFERROR(SUM([[Food Waste.value]])/'
        'SUM([[Grocery Spending.amount]]),"")',
        '=COUNTIF([[Meal Plan.home_cooked]],"Yes")',
        '=COUNTIF([[Recipes.name]],"<>")',
        '=SUM([[Pantry.value]])',
    ], ["money", "money", "money", "money", "percent", "int", "int",
        "money"])
    ws.conditional_format(row, 4, row, 4,
                          {"type": "cell", "criteria": ">",
                           "value": cfg.waste_reduction_target,
                           "format": fmts.cf("warn")})
    row += 2

    s.note(row, "Everything on this page is live. If a number looks "
                "wrong, the fix is always on the tab named in the "
                "'Where to fix it' column - never here.", LAST_COL,
           height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/weekly.py`

*255 lines · 2 functions*

> `Weekly Planner` - the week on one page.
>
> Seven days across, your meal slots down, with the cost and the calories
> worked out underneath and a shopping summary at the bottom. Pick a week
> with the date at the top and the whole page re-reads the Meal Plan.
>
> Every cell here is a lookup, not a place to type: meals are added on the
> Meal Plan tab (one row each, with drop-downs) and this page follows
> along. That keeps one copy of the truth, which is what lets the grocery
> list be trustworthy.

| Defines | What it does |
|---------|--------------|
| `_slot_cell(day_cell, slot, column)` | Pull `column` from the meal planned in `slot` on `day_cell` |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Weekly Planner'), `DAYS` (7), `LAST_COL` (8), `DAY_WIDTH` (19.0).

````python
"""`Weekly Planner` - the week on one page.

Seven days across, your meal slots down, with the cost and the calories
worked out underneath and a shopping summary at the bottom. Pick a week
with the date at the top and the whole page re-reads the Meal Plan.

Every cell here is a lookup, not a place to type: meals are added on the
Meal Plan tab (one row each, with drop-downs) and this page follows
along. That keeps one copy of the truth, which is what lets the grocery
list be trustworthy.
"""

from __future__ import annotations

from ..config import Config, MAIN_SLOTS, MEAL_SLOTS, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Weekly Planner"
DAYS = 7
LAST_COL = 8                 # A = labels, B..H = the week, I = totals
DAY_WIDTH = 19.0

_KEY = "[[Meal Plan.key]]"


def _slot_cell(day_cell: str, slot: str, column: str) -> str:
    """Pull `column` from the meal planned in `slot` on `day_cell`."""
    key = f'TEXT({day_cell},"yyyy-mm-dd")&"|"&"{slot}"'
    return (f'=IF({day_cell}="","",IFERROR(INDEX('
            f'[[Meal Plan.{column}]],MATCH({key},{_KEY},0)),""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Weekly Meal Planner",
             "One week at a glance - what you are eating, what it costs "
             "and what still needs cooking.", LAST_COL)
    s.set_col(0, 0, 16)
    for c in range(1, DAYS + 1):
        s.set_col(c, c, DAY_WIDTH)
    s.set_col(DAYS + 1, DAYS + 1, 15)

    # ------------------------------------------------------ week picker --
    row = 4
    s.merge(row, 0, 0, "Week starting", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.start_date, fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("WeekPick", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "Week number", fmts.chrome("label_right"))
    ws.write_formula(row, 3,
                     '=WEEKNUM(WeekPick,IF(WeekStart="Sunday",1,2))',
                     fmts.calc("int"))
    s.merge(row, 4, 4, "People", fmts.chrome("label_right"))
    ws.write_formula(row, 5, "=DefaultServings", fmts.calc("int"))
    s.merge(row, 6, 6, "Weekly budget", fmts.chrome("label_right"))
    ws.write_formula(row, 7, "=WeeklyBudget", fmts.calc("money"))
    row += 1

    row = s.note(row,
                 "Change the date above to any Monday - every box on "
                 "the page follows it. Meals themselves are added on "
                 "the Meal Plan tab; this page is the view, not the "
                 "entry form.", LAST_COL, height=26)

    # ----------------------------------------------------- day headings --
    head = row
    s.write(head, 0, "Meal", fmts.chrome("header_left"))
    for i in range(DAYS):
        ws.write_formula(head, i + 1, f"=WeekPick+{i}",
                         fmts.chrome("header", num_format='ddd d mmm'))
    s.write(head, DAYS + 1, "Week total", fmts.chrome("header"))
    s.row_height(head, 26)
    day_cells = [f"{chr(66 + i)}${head + 1}" for i in range(DAYS)]
    row = head + 1

    # ----------------------------------------------------- the meal grid --
    slot_rows = {}
    for slot in MEAL_SLOTS:
        s.write(row, 0, slot, fmts.chrome("label_bold"))
        for i in range(DAYS):
            ws.write_formula(row, i + 1,
                             s.reg.render(_slot_cell(day_cells[i], slot,
                                                     "recipe_name")),
                             fmts.cell("wrap", "calc"))
        ws.write_formula(
            row, DAYS + 1,
            s.reg.render(f'=COUNTIFS([[Meal Plan.date]],">="&WeekPick,'
                         f'[[Meal Plan.date]],"<="&WeekPick+6,'
                         f'[[Meal Plan.slot]],"{slot}")'),
            fmts.calc("int"))
        s.row_height(row, 30)
        slot_rows[slot] = row
        row += 1

    # cost and calories strips -------------------------------------------
    s.write(row, 0, "Cost of the day", fmts.chrome("label_bold"))
    for i in range(DAYS):
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IF({day_cells[i]}="","",'
                         f'SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],'
                         f'{day_cells[i]}))'),
            fmts.calc("money"))
    ws.write_formula(row, DAYS + 1,
                     s.reg.render('=SUMIFS([[Meal Plan.cost]],'
                                  '[[Meal Plan.date]],">="&WeekPick,'
                                  '[[Meal Plan.date]],"<="&WeekPick+6)'),
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    cost_row = row
    row += 1

    s.write(row, 0, "Calories a head", fmts.chrome("label_bold"))
    for i in range(DAYS):
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IF({day_cells[i]}="","",IFERROR('
                         f'SUMIFS([[Meal Plan.cals_total]],'
                         f'[[Meal Plan.date]],{day_cells[i]})/'
                         f'DefaultServings,""))'),
            fmts.calc("number"))
    ws.write_formula(
        row, DAYS + 1,
        s.reg.render('=IFERROR(SUMIFS([[Meal Plan.cals_total]],'
                     '[[Meal Plan.date]],">="&WeekPick,'
                     '[[Meal Plan.date]],"<="&WeekPick+6)/'
                     '(DefaultServings*7),"")'),
        fmts.calc("number"))
    cals_row = row
    row += 1

    s.write(row, 0, "Who is cooking", fmts.chrome("label_bold"))
    for i in range(DAYS):
        ws.write_formula(
            row, i + 1,
            s.reg.render(_slot_cell(day_cells[i], MAIN_SLOTS[-1], "cook")),
            fmts.calc("text"))
    ws.write_formula(row, DAYS + 1,
                     s.reg.render('=COUNTIFS([[Meal Plan.date]],'
                                  '">="&WeekPick,[[Meal Plan.date]],'
                                  '"<="&WeekPick+6,[[Meal Plan.cook]],'
                                  '"<>")'),
                     fmts.calc("int"))
    row += 2

    # cost colouring: green under budget per day, amber over -------------
    daily = "WeeklyBudget/7"
    ws.conditional_format(cost_row, 1, cost_row, DAYS,
                          {"type": "formula",
                           "criteria": f"=AND(B{cost_row + 1}<>\"\","
                                       f"B{cost_row + 1}>{daily})",
                           "format": fmts.cf("warn")})
    ws.conditional_format(cost_row, 1, cost_row, DAYS,
                          {"type": "formula",
                           "criteria": f"=AND(B{cost_row + 1}<>\"\","
                                       f"B{cost_row + 1}<={daily})",
                           "format": fmts.cf("good")})
    ws.conditional_format(cals_row, 1, cals_row, DAYS,
                          {"type": "formula",
                           "criteria": f"=AND(B{cals_row + 1}<>\"\","
                                       f"B{cals_row + 1}>CalorieTarget)",
                           "format": fmts.cf("warn")})

    # ------------------------------------------------- the week in numbers
    row = s.section(row, "This week in numbers", LAST_COL)
    kpis = [
        ("Meals planned",
         '=COUNTIFS([[Meal Plan.date]],">="&WeekPick,'
         '[[Meal Plan.date]],"<="&WeekPick+6)', "int",
         "every slot with something in it"),
        ("Planned food cost",
         '=SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],">="&WeekPick,'
         '[[Meal Plan.date]],"<="&WeekPick+6)', "money",
         "recipes only - the shop may differ"),
        ("Against the budget",
         '=WeeklyBudget-SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],'
         '">="&WeekPick,[[Meal Plan.date]],"<="&WeekPick+6)', "money",
         "green is money left"),
        ("Cooked at home",
         '=COUNTIFS([[Meal Plan.date]],">="&WeekPick,'
         '[[Meal Plan.date]],"<="&WeekPick+6,'
         '[[Meal Plan.home_cooked]],"Yes")', "int",
         f"target {cfg.home_cooked_target} a week"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ---------------------------------------------------- shopping & prep
    row = s.section(row, "Before the week starts", LAST_COL)
    row = s.table_head(row, 0, ["Job", "Where it stands", "Detail"],
                       [26, 20, 44])
    jobs = [
        ("Shopping list built",
         '=IF(COUNTIF([[Grocery List.ingredient]],"<>")=0,'
         '"Nothing on it","'
         'Ready")',
         '="The list covers "&TEXT(ShopFrom,"d mmm")&" to "&'
         'TEXT(ShopTo,"d mmm")&" - "&'
         'COUNTIF([[Grocery List.ingredient]],"<>")&" items, about "&'
         'TEXT(SUM([[Grocery List.est_cost]]),"0.00")'),
        ("Shopping done",
         '=IF(COUNTIF([[Grocery List.got]],"Yes")=0,"Not started",'
         'IF(COUNTIFS([[Grocery List.short]],">0",'
         '[[Grocery List.got]],"<>Yes")=0,"Done","In progress"))',
         '=COUNTIF([[Grocery List.got]],"Yes")&" of "&'
         'COUNTIFS([[Grocery List.short]],">0")&" items ticked off"'),
        ("Prep jobs for this week",
         '=IF(COUNTIFS([[Meal Prep.date]],">="&WeekPick,'
         '[[Meal Prep.date]],"<="&WeekPick+6)=0,"None planned",'
         'IF(COUNTIFS([[Meal Prep.date]],">="&WeekPick,'
         '[[Meal Prep.date]],"<="&WeekPick+6,'
         '[[Meal Prep.status]],"<>Done")=0,"Done","To do"))',
         '=COUNTIFS([[Meal Prep.date]],">="&WeekPick,'
         '[[Meal Prep.date]],"<="&WeekPick+6)&" job(s) on the prep day"'),
        ("Leftovers to eat first",
         '=IF(COUNTIFS([[Leftovers.status]],"<>Eaten",'
         '[[Leftovers.status]],"<>Binned")=0,"None","Use soon")',
         '=SUMIFS([[Leftovers.portions]],[[Leftovers.status]],'
         '"<>Eaten",[[Leftovers.status]],"<>Binned")&'
         '" portion(s) in the fridge or freezer"'),
        ("Food about to go off",
         '=IF(COUNTIF([[Pantry.status]],"Use soon")=0,"All good",'
         '"Use soon")',
         '=COUNTIF([[Pantry.status]],"Use soon")&" pantry item(s) '
         'within "&ExpiryWarningDays&" days of their date"'),
    ]
    for label, state, detail in jobs:
        s.write(row, 0, label, fmts.chrome("label"))
        ws.write_formula(row, 1, s.reg.render(state), fmts.calc("status"))
        ws.write_formula(row, 2, s.reg.render(detail), fmts.calc("text"))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(detail), fmts.calc("text"))
        row += 1
    for kind, word in (("good", "Done"), ("good", "Ready"),
                       ("warn", "To do"), ("warn", "In progress"),
                       ("warn", "Use soon"), ("muted", "None"),
                       ("muted", "Nothing on it"), ("good", "All good"),
                       ("warn", "Not started")):
        ws.conditional_format(row - len(jobs), 1, row - 1, 1,
                              {"type": "text", "criteria": "containing",
                               "value": word, "format": fmts.cf(kind)})
    row += 1

    s.note(row, "Printing: this page is set to one page wide. Print it "
                "on a Sunday, stick it on the fridge and the week looks "
                "after itself.", LAST_COL, height=24)

    s.finish(freeze_row=head + 1, freeze_col=1,
             print_area=f"A1:I{row + 2}", gridlines=False)
    return s
````

### `meal_planner_spreadsheet/sheets/monthly.py`

*156 lines · 2 functions*

> `Monthly Planner` - the month as a calendar.
>
> Six weeks of boxes, each showing the dinner (and the lunch, if there is
> one) planned for that day with its cost underneath. Pick a month at the
> top; everything reads from the Meal Plan, so the grid is never out of
> step with the shopping list.

| Defines | What it does |
|---------|--------------|
| `_slot(day_cell, slot)` |  |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Monthly Planner'), `DAYS` (7), `WEEKS` (6), `LAST_COL` (DAYS - 1), `COL_WIDTH` (21.0).

````python
"""`Monthly Planner` - the month as a calendar.

Six weeks of boxes, each showing the dinner (and the lunch, if there is
one) planned for that day with its cost underneath. Pick a month at the
top; everything reads from the Meal Plan, so the grid is never out of
step with the shopping list.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from

NAME = "Monthly Planner"
DAYS = 7
WEEKS = 6
LAST_COL = DAYS - 1
COL_WIDTH = 21.0

_KEY = "[[Meal Plan.key]]"


def _slot(day_cell: str, slot: str) -> str:
    key = f'TEXT({day_cell},"yyyy-mm-dd")&"|"&"{slot}"'
    return (f'=IF({day_cell}="","",IFERROR(INDEX('
            f'[[Meal Plan.recipe_name]],MATCH({key},{_KEY},0)),""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Monthly Meal Planner",
             "The whole month on one page - dinners, lunches and what "
             "each day costs. Print it and pin it up.", LAST_COL)
    for c in range(DAYS):
        s.set_col(c, c, COL_WIDTH)

    # --------------------------------------------------- month controls --
    row = 4
    s.merge(row, 0, 0, "Month", fmts.chrome("label_bold"))
    s.merge(row, 1, 1, cfg.today.strftime("%B"), fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("MonthList"))
    book.define_name("PlanMonthName", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "First day", fmts.chrome("label_right"))
    ws.write_formula(
        row, 3,
        '=DATE(PlanYear,MAX(1,IFERROR(MATCH(PlanMonthName,MonthList,0),'
        'MONTH(TODAY()))),1)', fmts.calc("date"))
    book.define_name("PlanMonthStart", f"={quote_sheet(NAME)}!$D${row + 1}")
    s.merge(row, 4, 4, "Meals planned", fmts.chrome("label_right"))
    ws.write_formula(
        row, 5,
        s.reg.render('=COUNTIFS([[Meal Plan.date]],">="&PlanMonthStart,'
                     '[[Meal Plan.date]],"<="&EOMONTH(PlanMonthStart,0))'),
        fmts.calc("int"))
    s.merge(row, 6, 6, None, fmts.calc("money"))
    ws.write_formula(
        row, 6,
        s.reg.render('=SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],'
                     '">="&PlanMonthStart,[[Meal Plan.date]],'
                     '"<="&EOMONTH(PlanMonthStart,0))'),
        fmts.calc("money"))
    row += 1

    row = s.note(row,
                 "Each box shows the dinner in bold type, then the "
                 "lunch, then what the day costs. Empty boxes are days "
                 "with nothing planned - the gaps are the point.",
                 LAST_COL, height=26)

    # -------------------------------------------------- weekday headings --
    first_cell = ('=PlanMonthStart-MOD(WEEKDAY(PlanMonthStart,'
                  'IF(WeekStart="Sunday",1,2))-1,7)')
    head = row
    for i in range(DAYS):
        ws.write_formula(
            head, i,
            f'=TEXT({first_cell.lstrip("=")}+{i},"dddd")',
            fmts.chrome("header"))
    s.row_height(head, 24)
    row = head + 1

    # ------------------------------------------------------- the calendar
    for week in range(WEEKS):
        date_row = row
        meal_row = row + 1
        lunch_row = row + 2
        cost_row = row + 3
        for i in range(DAYS):
            offset = week * 7 + i
            ws.write_formula(
                date_row, i,
                f'={first_cell.lstrip("=")}+{offset}',
                fmts.chrome("day_number", num_format="d mmm"))
            cell = f"{chr(65 + i)}${date_row + 1}"
            ws.write_formula(meal_row, i,
                             s.reg.render(_slot(cell, "Dinner")),
                             fmts.chrome("day_item"))
            ws.write_formula(lunch_row, i,
                             s.reg.render(_slot(cell, "Lunch")),
                             fmts.cell("wrap", "calc"))
            ws.write_formula(
                cost_row, i,
                s.reg.render(f'=IF({cell}="","",IF(SUMIFS('
                             f'[[Meal Plan.cost]],[[Meal Plan.date]],'
                             f'{cell})=0,"",SUMIFS([[Meal Plan.cost]],'
                             f'[[Meal Plan.date]],{cell})))'),
                fmts.calc("money"))
            # grey out the days that belong to the month either side
            ws.conditional_format(
                date_row, i, cost_row, i,
                {"type": "formula",
                 "criteria": f'=MONTH({cell})<>MONTH(PlanMonthStart)',
                 "format": fmts.cf("muted")})
        s.row_height(date_row, 16)
        s.row_height(meal_row, 26)
        s.row_height(lunch_row, 22)
        s.row_height(cost_row, 16)
        row = cost_row + 1
    row += 1

    # ------------------------------------------------------ month summary
    row = s.section(row, "The month in numbers", LAST_COL)
    row = s.table_head(row, 0, ["Meals planned", "Days covered",
                                "Planned food cost", "Monthly budget",
                                "Difference", "Cooked at home",
                                "Eaten out"],
                       [14, 13, 17, 14, 13, 15, 13])
    span = ('[[Meal Plan.date]],">="&PlanMonthStart,'
            '[[Meal Plan.date]],"<="&EOMONTH(PlanMonthStart,0)')
    diff_row = row
    row = s.calc_row(row, 0, [
        f'=COUNTIFS({span})',
        f'=SUMPRODUCT((COUNTIFS({span},[[Meal Plan.date]],'
        f'[[Meal Plan.date]])>0)*1)/MAX(1,1)',
        f'=SUMIFS([[Meal Plan.cost]],{span})',
        "=MonthlyBudget",
        f'=MonthlyBudget-SUMIFS([[Meal Plan.cost]],{span})',
        f'=COUNTIFS({span},[[Meal Plan.home_cooked]],"Yes")',
        f'=COUNTIFS({span},[[Meal Plan.home_cooked]],"No")',
    ], ["int", "int", "money", "money", "money", "int", "int"])
    ws.conditional_format(diff_row, 4, diff_row, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row += 1

    s.note(row, "Tip: plan the month roughly - the same six or seven "
                "dinners on a rotation is fine - then fine tune one "
                "week at a time on the Weekly Planner.", LAST_COL,
           height=24)

    s.finish(freeze_row=head + 1, print_area=f"A1:G{row + 2}",
             gridlines=False)
    return s
````

### `meal_planner_spreadsheet/sheets/filters.py`

*148 lines · 1 function*

> `Dietary Filters` - 'what can I actually make tonight?'
>
> Six drop-downs at the top: diet, category, method, how long you have,
> what it can cost a serving and an allergen to avoid. Underneath, every
> recipe that passes, with its time, cost, calories and the reason it
> suits. It is the answer to the six o'clock question, and it is also how
> you cook for the vegetarian, the coeliac and the fussy one at once.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dietary Filters'), `LAST_COL` (9), `RESULTS` (25).

````python
"""`Dietary Filters` - 'what can I actually make tonight?'

Six drop-downs at the top: diet, category, method, how long you have,
what it can cost a serving and an allergen to avoid. Underneath, every
recipe that passes, with its time, cost, calories and the reason it
suits. It is the answer to the six o'clock question, and it is also how
you cook for the vegetarian, the coeliac and the fussy one at once.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

NAME = "Dietary Filters"
LAST_COL = 9
RESULTS = 25


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Dietary Filters",
             "Pick the rules, get the list. Every recipe that fits, "
             "with the time, the cost and the calories.", LAST_COL)
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 24)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 14)

    # ----------------------------------------------------- the controls --
    row = 4
    row = s.section(row, "Set the rules", LAST_COL)
    controls = [
        ("Diet", "FilterDiet", "Any", list_from("DietList"),
         "Vegetarian, vegan, gluten-free... 'Any' turns the filter off."),
        ("Category", "FilterCategory", "Any",
         list_from("RecipeCategoryList"),
         "Mains, soups, lunchboxes, puddings."),
        ("Method", "FilterMethod", "Any",
         list_from("CookingMethodList"),
         "Only the slow cooker? Only no-cook? Choose it here."),
        ("Ready within (minutes)", "FilterTime", 0,
         number_check(0, 600, decimal=False),
         "Total time, prep plus cook. 0 means no limit."),
        ("Cost a serving up to", "FilterCost", 0,
         number_check(0, 1000),
         "0 means no limit. Try your per-meal budget from the Budget "
         "tab."),
        ("Avoid this allergen", "FilterAllergen", "Any",
         list_from("AllergenList"),
         "Hides every recipe with that allergen anywhere in its "
         "ingredients."),
    ]
    for label, name, default, validate, help_text in controls:
        style = "int" if isinstance(default, int) else "text"
        row = s.field(row, label, default, style=style, name=name,
                      validate=validate, help_text=help_text)
    row += 1

    # --------------------------------------------------- what matched ---
    row = s.section(row, "What you can make", LAST_COL)
    s.merge(row, 0, 2, None, fmts.chrome("kpi_value"))
    ws.write_formula(row, 0,
                     s.reg.render('="Recipes that fit: "&'
                                  'SUM([[Recipes.match]])'),
                     fmts.chrome("kpi_value", align="left"))
    s.merge(row, 3, LAST_COL,
            "Blank rows mean nothing else matched - loosen a rule.",
            fmts.chrome("note"))
    row += 2

    row = s.table_head(row, 0, ["#", "Recipe", "Category", "Suits",
                                "Total time", "Cost a serving",
                                "Calories", "Rating", "Last cooked",
                                "Freezes?"],
                       [5, 26, 14, 15, 12, 14, 11, 10, 12, 10])
    first = row
    for n in range(RESULTS):
        s.write(row, 0, n + 1, fmts.chrome("label"))
        pick = (f'MATCH({n + 1},[[Recipes.match_rank]],0)')
        for col, key, style in [(1, "name", "text"),
                                (2, "category", "text"),
                                (3, "diet", "text"),
                                (4, "total_min", "int"),
                                (5, "cost_serving", "money"),
                                (6, "calories", "number"),
                                (7, "rating", "rating"),
                                (8, "last_cooked", "date"),
                                (9, "freezes", "flag")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Recipes.{key}]],'
                             f'{pick}),"")'),
                fmts.calc(style))
        row += 1
    last = row - 1
    ws.conditional_format(first, 4, last, 4,
                          {"type": "cell", "criteria": "<=", "value": 30,
                           "format": fmts.cf("good")})
    ws.conditional_format(first, 5, last, 5,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    row += 1

    # ---------------------------------------------- who can eat what ----
    row = s.section(row, "Who in the house this suits", LAST_COL)
    row = s.table_head(row, 0, ["Person", "Diet", "Allergies",
                                "Recipes that suit their diet",
                                "Recipes to avoid"],
                       [18, 16, 20, 26, 20])
    for i in range(6):
        cell = f"A{row + 1}"
        ws.write_formula(row, 0,
                         f'=IFERROR(IF(INDEX(MemberNames,{i + 1})=0,"",'
                         f'INDEX(MemberNames,{i + 1})),"")',
                         fmts.calc("text"))
        diet = (f'IFERROR(INDEX([[Household.diet]],MATCH({cell},'
                f'[[Household.name]],0)),"")')
        allergy = (f'IFERROR(INDEX([[Household.allergies]],MATCH({cell},'
                   f'[[Household.name]],0)),"")')
        s.calc_row(row, 1, [
            f'=IF({cell}="","",{diet})',
            f'=IF({cell}="","",{allergy})',
            f'=IF({cell}="","",IF({diet}="No restriction",'
            f'COUNTIF([[Recipes.name]],"<>"),COUNTIFS([[Recipes.diet]],'
            f'{diet})))',
            f'=IF({cell}="","",IF(OR({allergy}="",{allergy}="None"),0,'
            f'SUMPRODUCT((COUNTIFS([[Recipe Ingredients.recipe]],'
            f'[[Recipes.id]],[[Recipe Ingredients.allergen]],'
            f'{allergy})>0)*([[Recipes.name]]<>""))))',
        ], ["text", "text", "int", "int"])
        row += 1
    row += 1

    s.note(row,
           "Allergen filtering looks at the allergen tag on each "
           "ingredient in the price book, so it is only as complete as "
           "those tags. It is a sorting tool, not a safety check: for "
           "a real allergy, read the label on the actual packet every "
           "single time, and remember that 'may contain' warnings "
           "never appear in a spreadsheet.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/nutrition.py`

*208 lines · 1 function*

> `Nutrition` - calories and macros, without the guilt.
>
> Everything here is worked out from the recipes you planned and the
> nutrition you copied off the packets: calories, protein, carbs, fat and
> fibre a head, day by day, against the targets on the Settings tab. It
> is a guide to what you are eating, not medical advice, and it is only
> as accurate as the numbers in the price book.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Nutrition'), `LAST_COL` (9), `MACROS` (5 items).

````python
"""`Nutrition` - calories and macros, without the guilt.

Everything here is worked out from the recipes you planned and the
nutrition you copied off the packets: calories, protein, carbs, fat and
fibre a head, day by day, against the targets on the Settings tab. It
is a guide to what you are eating, not medical advice, and it is only
as accurate as the numbers in the price book.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Nutrition"
LAST_COL = 9

MACROS = [
    ("Calories", "cals_total", "CalorieTarget", "number", "kcal"),
    ("Protein", "protein", "ProteinTarget", "number1", "g"),
    ("Carbs", "carbs", "CarbTarget", "number1", "g"),
    ("Fat", "fat", "FatTarget", "number1", "g"),
    ("Fibre", "fibre", "FibreTarget", "number1", "g"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Calories & Macros",
             "What the plan adds up to, a head, a day - measured "
             "against your own targets.", LAST_COL)
    s.set_col(0, 0, 18)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ------------------------------------------------------ the targets --
    row = 4
    row = s.section(row, "Your daily targets", LAST_COL)
    row = s.table_head(row, 0, ["Calories", "Protein (g)", "Carbs (g)",
                                "Fat (g)", "Fibre (g)", "Sugar (g)",
                                "Sodium (mg)", "Water (glasses)"],
                       [12, 12, 12, 11, 12, 12, 13, 15])
    row = s.calc_row(row, 0, [
        "=CalorieTarget", "=ProteinTarget", "=CarbTarget", "=FatTarget",
        "=FibreTarget", "=SugarLimit", "=SodiumLimit", "=WaterTarget",
    ], ["number", "number1", "number1", "number1", "number1",
        "number1", "number", "int"])
    row = s.note(row, "Set these on the Settings tab. They are one "
                      "set of numbers for the household - the "
                      "Household tab holds a calorie figure per person "
                      "if you want to be more precise.", LAST_COL,
                 height=26)

    # --------------------------------------------------- the week ------
    row = s.section(row, "This week, a head, a day", LAST_COL)
    head = row
    s.write(head, 0, "", fmts.chrome("header_left"))
    for i in range(7):
        ws.write_formula(head, i + 1, f"=WeekPick+{i}",
                         fmts.chrome("header", num_format="ddd d"))
    s.write(head, 8, "Average", fmts.chrome("header"))
    s.row_height(head, 24)
    row = head + 1

    per_head = "MAX(1,DefaultServings)"
    for label, key, target, style, unit in MACROS:
        s.write(row, 0, f"{label} ({unit})", fmts.chrome("label_bold"))
        for i in range(7):
            cell = f"{chr(66 + i)}${head + 1}"
            if key == "cals_total":
                total = (f'SUMIFS([[Meal Plan.cals_total]],'
                         f'[[Meal Plan.date]],{cell})')
            else:
                total = (f'SUMPRODUCT(([[Meal Plan.date]]={cell})*'
                         f'IFERROR([[Meal Plan.{key}]]*1,0)*'
                         f'IFERROR([[Meal Plan.eff_servings]]*1,0))')
            ws.write_formula(
                row, i + 1,
                s.reg.render(f'=IF({cell}="","",IFERROR({total}/'
                             f'{per_head},""))'),
                fmts.calc(style))
        letter_first, letter_last = "B", "H"
        ws.write_formula(row, 8,
                         f'=IFERROR(AVERAGE({letter_first}{row + 1}:'
                         f'{letter_last}{row + 1}),"")',
                         fmts.chrome("total_value"))
        ws.conditional_format(row, 1, row, 7,
                              {"type": "formula",
                               "criteria": f'=AND(B{row + 1}<>"",'
                                           f'B{row + 1}>{target}*1.15)',
                               "format": fmts.cf("warn")})
        ws.conditional_format(row, 1, row, 7,
                              {"type": "formula",
                               "criteria": f'=AND(B{row + 1}<>"",'
                                           f'B{row + 1}>={target}*0.85,'
                                           f'B{row + 1}<={target}*1.15)',
                               "format": fmts.cf("good")})
        row += 1

    s.write(row, 0, "Meals logged", fmts.chrome("label_bold"))
    for i in range(7):
        cell = f"{chr(66 + i)}${head + 1}"
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IF({cell}="","",COUNTIFS([[Meal Plan.date]],'
                         f'{cell}))'), fmts.calc("int"))
    ws.write_formula(row, 8, f'=SUM(B{row + 1}:H{row + 1})',
                     fmts.chrome("total_value"))
    row += 2

    # -------------------------------------------------- against target --
    row = s.section(row, "How the week compares", LAST_COL)
    row = s.table_head(row, 0, ["", "Average a day", "Target",
                                "Difference", "% of target",
                                "Verdict"],
                       [18, 16, 13, 14, 14, 26])
    first_cmp = row
    for idx, (label, key, target, style, unit) in enumerate(MACROS):
        src = f"I{head + 2 + idx}"
        s.write(row, 0, f"{label} ({unit})", fmts.chrome("label_bold"))
        s.calc_row(row, 1, [
            f"={src}",
            f"={target}",
            f'=IFERROR({src}-{target},"")',
            f'=IFERROR({src}/{target},"")',
            f'=IF({src}="","No data yet",IF({src}<{target}*0.85,'
            f'"Under - add a bit more",IF({src}>{target}*1.15,'
            f'"Over - worth a look","About right")))',
        ], ["number1", "number1", "number1", "percent0", "text"])
        row += 1
    ws.conditional_format(first_cmp, 5, row - 1, 5,
                          {"type": "text", "criteria": "containing",
                           "value": "About right",
                           "format": fmts.cf("good")})
    ws.conditional_format(first_cmp, 5, row - 1, 5,
                          {"type": "text", "criteria": "containing",
                           "value": "Over",
                           "format": fmts.cf("warn")})
    ws.conditional_format(first_cmp, 5, row - 1, 5,
                          {"type": "text", "criteria": "containing",
                           "value": "Under",
                           "format": fmts.cf("info")})
    ws.conditional_format(first_cmp, 4, row - 1, 4,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    row += 1

    # ------------------------------------------------ by meal and person
    row = s.section(row, "Where the calories come from", LAST_COL)
    row = s.table_head(row, 0, ["Meal", "Calories a week",
                                "Share of the week", "Meals planned",
                                "Average a meal"],
                       [18, 17, 17, 15, 16])
    week = ('[[Meal Plan.date]],">="&WeekPick,[[Meal Plan.date]],'
            '"<="&WeekPick+6')
    for slot in ["Breakfast", "Lunch", "Dinner", "Dessert"]:
        s.write(row, 0, slot, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=SUMIFS([[Meal Plan.cals_total]],{week},'
            f'[[Meal Plan.slot]],"{slot}")',
            f'=IFERROR(SUMIFS([[Meal Plan.cals_total]],{week},'
            f'[[Meal Plan.slot]],"{slot}")/SUMIFS('
            f'[[Meal Plan.cals_total]],{week}),"")',
            f'=COUNTIFS({week},[[Meal Plan.slot]],"{slot}")',
            f'=IFERROR(SUMIFS([[Meal Plan.cals_total]],{week},'
            f'[[Meal Plan.slot]],"{slot}")/COUNTIFS({week},'
            f'[[Meal Plan.slot]],"{slot}"),"")',
        ], ["number", "percent0", "int", "number"])
        row += 1
    row += 1

    # ------------------------------------------------- top recipes -----
    row = s.section(row, "The recipes doing the heavy lifting", LAST_COL)
    row = s.table_head(row, 0, ["Highest calorie recipes", "A serving",
                                "Most protein", "A serving",
                                "Most fibre", "A serving"],
                       [26, 12, 22, 12, 22, 12])
    for n in range(cfg.top_rows):
        for col, (value_key, label_col) in enumerate(
                [("calories", 0), ("protein", 2), ("fibre", 4)]):
            pick = (f'MATCH({n + 1},[[Recipes.{value_key}_rank]],0)')
            big = f'INDEX([[Recipes.{value_key}]],{pick})'
            s.write(row, label_col, None, fmts.calc("text"))
            ws.write_formula(
                row, label_col,
                s.reg.render(f'=IFERROR(IF({big}=0,"",INDEX('
                             f'[[Recipes.name]],{pick})),"")'),
                fmts.calc("text"))
            ws.write_formula(
                row, label_col + 1,
                s.reg.render(f'=IFERROR(IF({big}=0,"",{big}),"")'),
                fmts.calc("number1"))
        row += 1
    row += 1

    s.note(row,
           "Honest limits: these numbers come from the 'per 100' "
           "figures you typed into the Ingredients tab, multiplied by "
           "the quantities in your recipes. Cooking losses, oil left "
           "in the pan, who actually ate the bigger portion - none of "
           "that is in here. Treat it as a direction of travel, not a "
           "food diary, and never as medical or dietary advice.",
           LAST_COL, height=50)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/budget.py`

*250 lines · 1 function*

> `Budget` - where the food money actually goes.
>
> Twelve months down the page: what you budgeted, what you spent, what
> you spent it on and what was left. Underneath, the split by type of
> spend and by store, the cost of eating out, and the weekly rhythm that
> tells you whether the budget is realistic or just optimistic.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Budget'), `LAST_COL` (10), `MONTHS` (12), `SPEND_TYPES` (10 items).

````python
"""`Budget` - where the food money actually goes.

Twelve months down the page: what you budgeted, what you spent, what
you spent it on and what was left. Underneath, the split by type of
spend and by store, the cost of eating out, and the weekly rhythm that
tells you whether the budget is realistic or just optimistic.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Budget"
LAST_COL = 10
MONTHS = 12

SPEND_TYPES = ["Weekly shop", "Top-up shop", "Bulk shop", "Takeaway",
               "Eating out", "Work lunch", "Coffee & treats",
               "Delivery fee", "Special occasion", "Household"]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Food Budget",
             "Planned against actual, month by month, with the split "
             "that shows where it really goes.", LAST_COL)
    s.set_col(0, 0, 16)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 14)

    # ------------------------------------------------------- the targets
    row = 4
    s.merge(row, 0, 0, "Weekly budget", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=WeeklyBudget", fmts.calc("money"))
    s.merge(row, 2, 2, "Monthly", fmts.chrome("label_right"))
    ws.write_formula(row, 3, "=MonthlyBudget", fmts.calc("money"))
    s.merge(row, 4, 4, "A year", fmts.chrome("label_right"))
    ws.write_formula(row, 5, "=AnnualBudget", fmts.calc("money"))
    s.merge(row, 6, 6, "Per person a week", fmts.chrome("label_right"))
    ws.write_formula(row, 7,
                     "=IFERROR(WeeklyBudget/MAX(1,HouseMembers),\"\")",
                     fmts.calc("money"))
    s.merge(row, 8, 8, "Per meal", fmts.chrome("label_right"))
    ws.write_formula(row, 9,
                     "=IFERROR(WeeklyBudget/MAX(1,HouseMembers*21),\"\")",
                     fmts.calc("money"))
    row += 1
    row = s.note(row, "Change these on the Settings tab. Per meal "
                      "assumes three meals a day for everyone - a "
                      "useful reality check when the weekly figure "
                      "feels mean.", LAST_COL, height=26)

    # ------------------------------------------------- month by month --
    row = s.section(row, "Month by month", LAST_COL)
    row = s.table_head(row, 0, ["Month", "Budget", "Groceries",
                                "Eating out & takeaway", "Other",
                                "Total spent", "Left", "% of budget",
                                "Food waste", "Spend a head"],
                       [12, 12, 13, 19, 11, 13, 12, 12, 12, 13])
    first = row
    for m in range(MONTHS):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        groceries = (f'SUMIFS([[Grocery Spending.amount]],'
                     f'[[Grocery Spending.month]],{key},'
                     f'[[Grocery Spending.type]],"*shop*")')
        out = (f'SUMIFS([[Grocery Spending.amount]],'
               f'[[Grocery Spending.month]],{key},'
               f'[[Grocery Spending.type]],"Takeaway")+'
               f'SUMIFS([[Grocery Spending.amount]],'
               f'[[Grocery Spending.month]],{key},'
               f'[[Grocery Spending.type]],"Eating out")')
        total = (f'SUMIFS([[Grocery Spending.amount]],'
                 f'[[Grocery Spending.month]],{key})')
        ws.write_formula(row, 0,
                         f'=TEXT(DATE(PlanYear,{m + 1},1),"mmmm")',
                         fmts.calc("text"))
        s.calc_row(row, 1, [
            "=MonthlyBudget",
            f"={groceries}",
            f"={out}",
            f"={total}-{groceries}-({out})",
            f"={total}",
            f"=MonthlyBudget-{total}",
            f'=IFERROR(({total})/MonthlyBudget,"")',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.month]],{key})',
            f'=IFERROR(({total})/MAX(1,HouseMembers),"")',
        ], ["money", "money", "money", "money", "money", "money",
            "percent0", "money", "money"])
        row += 1
    last = row - 1
    ws.conditional_format(first, 6, last, 6,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first, 6, last, 6,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(first, 7, last, 7,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    ws.conditional_format(first, 5, last, 5,
                          {"type": "data_bar", "bar_color": "#2F5E43"})

    # year total line
    s.merge(row, 0, 0, "Year", fmts.chrome("total_label"))
    totals = [
        "=AnnualBudget",
        f'=SUM(C{first + 1}:C{last + 1})',
        f'=SUM(D{first + 1}:D{last + 1})',
        f'=SUM(E{first + 1}:E{last + 1})',
        f'=SUM(F{first + 1}:F{last + 1})',
        f'=AnnualBudget-SUM(F{first + 1}:F{last + 1})',
        f'=IFERROR(SUM(F{first + 1}:F{last + 1})/AnnualBudget,"")',
        f'=SUM(I{first + 1}:I{last + 1})',
        f'=IFERROR(SUM(F{first + 1}:F{last + 1})/MAX(1,HouseMembers),"")',
    ]
    for i, formula in enumerate(totals):
        num = "0%" if i == 6 else cfg.money_format
        ws.write_formula(row, i + 1, formula,
                         fmts.chrome("total_value", num_format=num))
    row += 2

    # ------------------------------------------------- split by type ---
    row = s.section(row, "What the money goes on", LAST_COL)
    row = s.table_head(row, 0, ["Type of spend", "This month",
                                "Year to date", "Share of the year",
                                "Times logged", "Average"],
                       [20, 14, 14, 16, 13, 13])
    first_type = row
    for kind in SPEND_TYPES:
        s.write(row, 0, kind, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}",'
            f'[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
            f'=SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}")',
            f'=IFERROR(SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}")/'
            f'SUM([[Grocery Spending.amount]]),"")',
            f'=COUNTIFS([[Grocery Spending.type]],"{kind}")',
            f'=IFERROR(AVERAGEIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}"),"")',
        ], ["money", "money", "percent0", "int", "money"])
        row += 1
    ws.conditional_format(first_type, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#D9703C"})
    row += 1

    # ------------------------------------------------- split by store --
    row = s.section(row, "Store by store", LAST_COL)
    row = s.table_head(row, 0, ["Store", "Shops", "Spent",
                                "Average shop", "Share", "Price level"],
                       [20, 10, 13, 14, 12, 14])
    for i in range(8):
        cell = f"A{row + 1}"
        ws.write_formula(row, 0,
                         f'=IFERROR(IF(INDEX(StoreNames,{i + 1})=0,"",'
                         f'INDEX(StoreNames,{i + 1})),"")',
                         fmts.calc("text"))
        s.calc_row(row, 1, [
            f'=IF({cell}="","",COUNTIFS([[Grocery Spending.store]],'
            f'{cell}))',
            f'=IF({cell}="","",SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.store]],{cell}))',
            f'=IF({cell}="","",IFERROR(AVERAGEIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.store]],'
            f'{cell}),""))',
            f'=IF({cell}="","",IFERROR(SUMIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.store]],'
            f'{cell})/SUM([[Grocery Spending.amount]]),""))',
            f'=IF({cell}="","",IFERROR(INDEX([[Stores.price_level]],'
            f'MATCH({cell},[[Stores.name]],0)),""))',
        ], ["int", "money", "money", "percent0", "text"])
        row += 1
    row += 1

    # ---------------------------------------------- planned vs actual --
    row = s.section(row, "Planned against actual", LAST_COL)
    row = s.note(row, "The planner prices your meals from the "
                      "Ingredients tab. The till tells you what really "
                      "happened. The gap between them is the honest "
                      "measure of how good your price book is - and of "
                      "how much gets bought that was never planned.",
                 LAST_COL, height=32)
    row = s.table_head(row, 0, ["", "Planned (recipes)",
                                "Actual (receipts)", "Difference",
                                "Unplanned buys", "Impulse spend"],
                       [16, 18, 18, 14, 14, 14])
    gap_row = row
    s.write(row, 0, "This month", fmts.chrome("label_bold"))
    s.calc_row(row, 1, [
        '=SUMIFS([[Meal Plan.cost]],[[Meal Plan.month]],'
        'TEXT(TODAY(),"yyyy-mm"))',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))-'
        'SUMIFS([[Meal Plan.cost]],[[Meal Plan.month]],'
        'TEXT(TODAY(),"yyyy-mm"))',
        '=COUNTIFS([[Grocery Spending.planned]],"No",'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.planned]],"No",'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
    ], ["money", "money", "money", "int", "money"])
    row += 1
    s.write(row, 0, "The year", fmts.chrome("label_bold"))
    s.calc_row(row, 1, [
        '=SUM([[Meal Plan.cost]])',
        '=SUM([[Grocery Spending.amount]])',
        '=SUM([[Grocery Spending.amount]])-SUM([[Meal Plan.cost]])',
        '=COUNTIF([[Grocery Spending.planned]],"No")',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.planned]],"No")',
    ], ["money", "money", "money", "int", "money"])
    ws.conditional_format(gap_row, 3, row, 3,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 2

    # ------------------------------------------------- savings tracker --
    row = s.section(row, "What the workbook has saved you", LAST_COL)
    row = s.table_head(row, 0, ["Leftovers eaten",
                                "Cheaper shops available",
                                "Waste avoided against target",
                                "Batch cooking against takeaway"],
                       [18, 22, 26, 28])
    s.calc_row(row, 0, [
        '=SUMIFS([[Leftovers.value]],[[Leftovers.status]],"Eaten")',
        '=SUM([[Ingredients.saving]])',
        '=MAX(0,WasteCostTarget-SUMIFS([[Food Waste.value]],'
        '[[Food Waste.month]],TEXT(TODAY(),"yyyy-mm")))',
        '=IFERROR(SUM([[Meal Prep.portions]])*'
        '(EatingOutBudget/MAX(1,TakeawayLimit*2))-'
        'SUM([[Meal Prep.cost]]),"")',
    ], ["money", "money", "money", "money"])
    row += 2
    s.note(row, "These are estimates, not accounting. 'Leftovers "
                "eaten' values the portions you did not have to cook "
                "again; 'cheaper shops available' is the ceiling if "
                "you always bought where it is cheapest; the batch "
                "cooking figure compares your prepped portions with "
                "what a takeaway portion costs you.", LAST_COL,
           height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/recipe_costs.py`

*181 lines · 1 function*

> `Recipe Costs` - what every meal actually costs to put on the table.
>
> Recipes ranked cheapest first, with the cost a serving, the cost per
> 100 calories and what the same meal would cost as a takeaway. It is the
> tab that settles arguments: the curry you think is expensive is 90p a
> head, and the 'cheap' pasta bake is not.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Recipe Costs'), `LAST_COL` (9), `ROWS` (25).

````python
"""`Recipe Costs` - what every meal actually costs to put on the table.

Recipes ranked cheapest first, with the cost a serving, the cost per
100 calories and what the same meal would cost as a takeaway. It is the
tab that settles arguments: the curry you think is expensive is 90p a
head, and the 'cheap' pasta bake is not.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, number_check

NAME = "Recipe Costs"
LAST_COL = 9
ROWS = 25


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent_soft)
    ws = s.ws

    s.banner("Recipe Costing",
             "Every recipe priced from your own shopping, cheapest "
             "first.", LAST_COL)
    s.set_col(0, 0, 6)
    s.set_col(1, 1, 28)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 14)

    # --------------------------------------------------------- the KPIs
    row = 4
    kpis = [
        ("Cheapest meal a serving",
         '=IFERROR(MIN(IFERROR([[Recipes.cost_serving]]*1,999999)),"")',
         "money", "the one to cook when money is tight"),
        ("Average meal a serving",
         '=IFERROR(AVERAGE([[Recipes.cost_serving]]),"")', "money",
         "across every recipe with a price"),
        ("Dearest meal a serving",
         '=IFERROR(MAX(IFERROR([[Recipes.cost_serving]]*1,0)),"")',
         "money", "worth saving for a Friday"),
        ("Recipes still unpriced",
         '=COUNTIFS([[Recipes.name]],"<>",[[Recipes.cost]],0)+'
         'COUNTIFS([[Recipes.name]],"<>",[[Recipes.cost]],"")',
         "int", "add the ingredient prices and they appear"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2 + 1, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ----------------------------------------------- the cost benchmark
    row = s.section(row, "What you are comparing against", LAST_COL)
    row = s.field(row, "A takeaway for the household costs",
                  round(cfg.eating_out_budget / 4, 2), style="money",
                  name="TakeawayCost",
                  validate=number_check(0, 1000),
                  help_text="Your own figure for one takeaway meal for "
                            "everybody. Used in the saving column "
                            "below.")
    row = s.field(row, "A bought lunch costs", 6.50, style="money",
                  name="BoughtLunch", validate=number_check(0, 200),
                  help_text="Meal deal, café sandwich, canteen - "
                            "whatever you would otherwise pay.")
    row += 1

    # ------------------------------------------------- the ranked table
    row = s.section(row, "Every recipe, cheapest first", LAST_COL)
    row = s.table_head(row, 0, ["#", "Recipe", "Serves", "Cost to make",
                                "Cost a serving", "Per 100 calories",
                                "Minutes", "Saving against a takeaway",
                                "Rating"],
                       [5, 28, 9, 14, 14, 16, 11, 24, 10])
    first = row
    for n in range(ROWS):
        pick = f'MATCH({n + 1},[[Recipes.cost_rank]],0)'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "name", "text"),
                                (2, "servings", "int"),
                                (3, "cost", "money"),
                                (4, "cost_serving", "money"),
                                (6, "total_min", "int"),
                                (8, "rating", "rating")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Recipes.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        r = row + 1
        cals = f'IFERROR(INDEX([[Recipes.calories]],{pick}),0)'
        ws.write_formula(
            row, 5,
            s.reg.render(f'=IF(E{r}="","",IFERROR(E{r}/({cals}/100),""))'),
            fmts.calc("money"))
        ws.write_formula(
            row, 7,
            f'=IF(D{r}="","",TakeawayCost-D{r})', fmts.calc("money"))
        row += 1
    last = row - 1
    ws.conditional_format(first, 4, last, 4,
                          {"type": "3_color_scale",
                           "min_color": "#D9F2E4", "mid_color": "#FFF1CF",
                           "max_color": "#FBD9DD"})
    ws.conditional_format(first, 7, last, 7,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(first, 3, last, 3,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    row += 1

    # ------------------------------------------- where the money goes --
    row = s.section(row, "The expensive ingredients", LAST_COL)
    row = s.note(row, "Sorted by what you spend on them across every "
                      "recipe - the five lines at the top are where a "
                      "swap or a cheaper shop pays off most.",
                 LAST_COL, height=24)
    row = s.table_head(row, 0, ["", "Ingredient", "Cost per 100",
                                "Used in recipes", "Total across "
                                "recipes", "Cheapest store",
                                "Saving available"],
                       [5, 26, 14, 15, 20, 18, 17])
    for n in range(cfg.top_rows * 2):
        pick = f'MATCH({n + 1},[[Ingredients.spend_rank]],0)'
        big = f'INDEX([[Ingredients.spend_key]],{pick})'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "name", "text"),
                                (2, "per_100", "money"),
                                (3, "used_in", "int"),
                                (5, "best_store", "text"),
                                (6, "saving", "money")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(IF({big}=0,"",INDEX('
                             f'[[Ingredients.{key}]],{pick})),"")'),
                fmts.calc(style))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IFERROR(IF({big}=0,"",{big}),"")'),
            fmts.calc("money"))
        row += 1
    row += 1

    # ------------------------------------------------ cost by category
    row = s.section(row, "Cost by kind of meal", LAST_COL)
    row = s.table_head(row, 0, ["Category", "Recipes", "Average a "
                                "serving", "Cheapest", "Dearest",
                                "Average minutes"],
                       [18, 11, 18, 13, 13, 16])
    for kind in ["Main", "Side", "Soup", "Salad", "Breakfast", "Snack",
                 "Dessert", "Lunchbox", "Batch base"]:
        s.write(row, 0, kind, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Recipes.category]],"{kind}")',
            f'=IFERROR(AVERAGEIFS([[Recipes.cost_serving]],'
            f'[[Recipes.category]],"{kind}"),"")',
            f'=IFERROR(IF(COUNTIFS([[Recipes.category]],"{kind}")=0,"",'
            f'SUMPRODUCT(MIN((([[Recipes.category]]="{kind}")*'
            f'IFERROR([[Recipes.cost_serving]]*1,0)>0)*'
            f'IFERROR([[Recipes.cost_serving]]*1,0)+'
            f'(1-(([[Recipes.category]]="{kind}")*'
            f'IFERROR([[Recipes.cost_serving]]*1,0)>0))*9999))),"")',
            f'=IFERROR(SUMPRODUCT(MAX(([[Recipes.category]]="{kind}")*'
            f'IFERROR([[Recipes.cost_serving]]*1,0))),"")',
            f'=IFERROR(AVERAGEIFS([[Recipes.total_min]],'
            f'[[Recipes.category]],"{kind}"),"")',
        ], ["int", "money", "money", "money", "int"])
        row += 1
    row += 1

    s.note(row,
           "A recipe is only as well priced as its ingredient lines. "
           "Anything flagged on the Recipe Ingredients tab - a missing "
           "price, a unit that does not match the price book - is left "
           "out of these figures rather than guessed at, so a "
           "suspiciously cheap recipe usually means a missing line "
           "rather than a bargain.",
           LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/converter.py`

*205 lines · 1 function*

> `Converter` - unit conversion and recipe scaling, on one page.
>
> Two tools. The top half converts any quantity between the units in the
> table on the Lists tab - grams to ounces, millilitres to cups, a
> tablespoon to a teaspoon. The bottom half takes one of your recipes and
> rewrites every ingredient line for a different number of people, with
> the new cost.
>
> Weight and volume are deliberately not interchangeable: 100 ml of oil
> is not 100 g of oil, and a workbook that pretends otherwise quietly
> ruins your baking.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Converter'), `LAST_COL` (9), `QUICK_ROWS` (10), `SCALE_ROWS` (18).

````python
"""`Converter` - unit conversion and recipe scaling, on one page.

Two tools. The top half converts any quantity between the units in the
table on the Lists tab - grams to ounces, millilitres to cups, a
tablespoon to a teaspoon. The bottom half takes one of your recipes and
rewrites every ingredient line for a different number of people, with
the new cost.

Weight and volume are deliberately not interchangeable: 100 ml of oil
is not 100 g of oil, and a workbook that pretends otherwise quietly
ruins your baking.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Converter"
LAST_COL = 9
QUICK_ROWS = 10
SCALE_ROWS = 18


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Converter & Recipe Scaler",
             "Change the units, or change how many people you are "
             "feeding - the quantities follow.", LAST_COL)
    s.set_col(0, 0, 22)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 14)

    # ------------------------------------------------- the quick convert
    row = 4
    row = s.section(row, "Convert a quantity", LAST_COL)
    row = s.note(row, "Type an amount and a unit on the left, choose "
                      "the unit you want on the right. Anything that "
                      "cannot be converted - grams into millilitres, "
                      "say - says so instead of guessing.",
                 LAST_COL, height=26)
    row = s.table_head(row, 0, ["Amount", "From", "To", "Answer",
                                "In base units", "Measures"],
                       [12, 14, 14, 16, 16, 14])
    first_q = row
    for i in range(QUICK_ROWS):
        s.write(row, 0, 100 if i == 0 else None, fmts.input("number1"))
        ws.data_validation(row, 1, row, 1, list_from("UnitList"))
        s.write(row, 1, "g" if i == 0 else None, fmts.input("text"))
        ws.data_validation(row, 2, row, 2, list_from("UnitList"))
        s.write(row, 2, "oz" if i == 0 else None, fmts.input("text"))
        r = row + 1
        base = (f'A{r}*INDEX(UnitFactors,MATCH(B{r},UnitNames,0))')
        to_factor = f'INDEX(UnitFactors,MATCH(C{r},UnitNames,0))'
        same = (f'INDEX(UnitBases,MATCH(B{r},UnitNames,0))='
                f'INDEX(UnitBases,MATCH(C{r},UnitNames,0))')
        ws.write_formula(
            row, 3,
            f'=IF(OR(A{r}="",B{r}="",C{r}=""),"",IFERROR(IF({same},'
            f'{base}/{to_factor},"Cannot convert "&'
            f'INDEX(UnitKinds,MATCH(B{r},UnitNames,0))&" to "&'
            f'INDEX(UnitKinds,MATCH(C{r},UnitNames,0))),"Check the '
            f'units"))', fmts.calc("number1"))
        ws.write_formula(
            row, 4,
            f'=IF(OR(A{r}="",B{r}=""),"",IFERROR({base}&" "&'
            f'INDEX(UnitBases,MATCH(B{r},UnitNames,0)),""))',
            fmts.calc("text"))
        ws.write_formula(
            row, 5,
            f'=IF(B{r}="","",IFERROR(INDEX(UnitKinds,MATCH(B{r},'
            f'UnitNames,0)),""))', fmts.calc("text"))
        row += 1
    ws.conditional_format(first_q, 3, row - 1, 3,
                          {"type": "text", "criteria": "containing",
                           "value": "Cannot", "format": fmts.cf("warn")})
    row += 1

    # ------------------------------------------------- the common ones --
    row = s.section(row, "The ones you look up every time", LAST_COL)
    row = s.table_head(row, 0, ["Metric", "=", "Imperial / cups",
                                "Metric", "=", "Imperial / cups"],
                       [18, 4, 22, 18, 4, 22])
    pairs = [
        ("1 teaspoon", "5 ml", "100 g flour", "about 3/4 cup"),
        ("1 tablespoon", "15 ml (3 tsp)", "100 g sugar", "1/2 cup"),
        ("1 cup", "240 ml", "100 g butter", "7 tablespoons"),
        ("250 ml", "about 1 cup", "100 g rice", "1/2 cup uncooked"),
        ("500 ml", "about 2 cups", "1 egg (medium)", "about 50 g"),
        ("1 litre", "about 4 cups", "1 clove garlic", "1 teaspoon minced"),
        ("100 g", "3.5 oz", "Oven 180C fan", "350F / gas 4"),
        ("500 g", "1 lb 2 oz", "Oven 200C fan", "400F / gas 6"),
        ("1 kg", "2 lb 3 oz", "Oven 220C fan", "425F / gas 7"),
    ]
    for left, left_v, right, right_v in pairs:
        s.write(row, 0, left, fmts.chrome("label"))
        s.write(row, 1, "=", fmts.chrome("label"))
        s.write(row, 2, left_v, fmts.chrome("label_bold"))
        s.write(row, 3, right, fmts.chrome("label"))
        s.write(row, 4, "=", fmts.chrome("label"))
        s.write(row, 5, right_v, fmts.chrome("label_bold"))
        row += 1
    row = s.note(row, "Cup measures are the US cup (240 ml). Flour and "
                      "sugar weights are approximate because a cup of "
                      "flour depends on how hard you pack it - weigh "
                      "it if the baking matters.", LAST_COL, height=26)

    # ----------------------------------------------------- the scaler ---
    row = s.section(row, "Scale a recipe", LAST_COL)
    row = s.field(row, "Recipe",
                  "R-0001" if cfg.include_sample_data else "",
                  name="ScaleRecipe",
                  validate=list_from("RecipeIDs"),
                  help_text="Pick the recipe ID. The Recipes tab lists "
                            "them.")
    row = s.field(row, "People to feed", cfg.default_servings,
                  style="int", name="ScalePeople",
                  validate=number_check(1, 200, decimal=False),
                  help_text="Every quantity below is multiplied to "
                            "suit.")
    s.merge(row, 1, 2, "Recipe name", fmts.chrome("label_right"))
    ws.write_formula(row, 3,
                     s.reg.render('=IFERROR(INDEX([[Recipes.name]],'
                                  'MATCH(ScaleRecipe,[[Recipes.id]],0)),'
                                  '"Pick a recipe above")'),
                     fmts.calc("text"))
    s.merge(row, 4, 4, "Normally serves", fmts.chrome("label_right"))
    ws.write_formula(row, 5,
                     s.reg.render('=IFERROR(INDEX([[Recipes.servings]],'
                                  'MATCH(ScaleRecipe,[[Recipes.id]],0)),'
                                  '"")'), fmts.calc("int"))
    s.merge(row, 6, 6, "Scale factor", fmts.chrome("label_right"))
    ws.write_formula(row, 7,
                     s.reg.render('=IFERROR(ScalePeople/INDEX('
                                  '[[Recipes.servings]],MATCH('
                                  'ScaleRecipe,[[Recipes.id]],0)),"")'),
                     fmts.calc("number1"))
    book.define_name("ScaleFactor", f"={quote_sheet(NAME)}!$H${row + 1}")
    row += 2

    row = s.table_head(row, 0, ["Ingredient", "Original", "Unit",
                                "You need", "In base units", "Prepared",
                                "Cost", "Optional?"],
                       [26, 12, 10, 14, 16, 18, 12, 10])
    first_s = row
    for n in range(SCALE_ROWS):
        pick = (f'MATCH(ScaleRecipe&"#"&{n + 1},'
                f'[[Recipe Ingredients.scale_key]],0)')
        for col, key, style in [(0, "ingredient", "text"),
                                (1, "qty", "number1"),
                                (2, "unit", "text"),
                                (5, "prep", "text"),
                                (7, "optional", "flag")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX('
                             f'[[Recipe Ingredients.{key}]],{pick}),"")'),
                fmts.calc(style))
        r = row + 1
        ws.write_formula(row, 3,
                         f'=IF(B{r}="","",IFERROR(ROUND(B{r}*'
                         f'ScaleFactor,2),""))', fmts.calc("number1"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IF(A{r}="","",IFERROR(ROUND(INDEX('
                         f'[[Recipe Ingredients.base_qty]],{pick})*'
                         f'ScaleFactor,0)&" "&INDEX('
                         f'[[Recipe Ingredients.base_unit]],{pick}),""))'),
            fmts.calc("text"))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=IF(A{r}="","",IFERROR(INDEX('
                         f'[[Recipe Ingredients.line_cost]],{pick})*'
                         f'ScaleFactor,""))'), fmts.calc("money"))
        row += 1
    s.merge(row, 0, 2, "Scaled totals", fmts.chrome("total_label"))
    ws.write_formula(row, 6, f'=SUM(G{first_s + 1}:G{row})',
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    s.merge(row, 7, LAST_COL, None, fmts.chrome("total_value",
                                                align="left"))
    ws.write_formula(
        row, 7,
        s.reg.render('=IFERROR("That is "&TEXT(ScalePeople,"0")&'
                     '" servings at "&TEXT(SUM(G'
                     + str(first_s + 1) + ':G' + str(row) + ')/'
                     'MAX(1,ScalePeople),"0.00")&" each","")'),
        fmts.chrome("total_value", align="left"))
    row += 2

    s.note(row,
           "Scaling multiplies every quantity by the same number, "
           "which is right for ingredients and wrong for cooking "
           "times. A double batch of stew takes a little longer to "
           "come to the boil, not twice as long; a doubled cake needs "
           "a bigger tin, not two of the same. Spices and salt "
           "rarely want doubling either - taste as you go.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/rotation.py`

*201 lines · 1 function*

> `Rotation` - stop eating the same six dinners.
>
> Which recipes you lean on, which you have not cooked for months, how
> much variety the last four weeks actually had, and a shortlist of
> things to bring back. Nobody needs a hundred recipes - they need the
> twenty they like, in an order that does not feel like a loop.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Rotation'), `LAST_COL` (8), `ROWS` (20).

````python
"""`Rotation` - stop eating the same six dinners.

Which recipes you lean on, which you have not cooked for months, how
much variety the last four weeks actually had, and a shortlist of
things to bring back. Nobody needs a hundred recipes - they need the
twenty they like, in an order that does not feel like a loop.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Rotation"
LAST_COL = 8
ROWS = 20


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Meal Rotation & Variety",
             "What you cook, how often, and what has been forgotten at "
             "the back of the recipe book.", LAST_COL)
    s.set_col(0, 0, 6)
    s.set_col(1, 1, 30)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 15)

    # ------------------------------------------------------------ KPIs
    row = 4
    kpis = [
        ("Recipes in the book",
         '=COUNTIF([[Recipes.name]],"<>")', "int",
         "everything you could cook"),
        ("Cooked in the last 30 days",
         '=SUMPRODUCT((COUNTIFS([[Meal Plan.recipe]],[[Recipes.id]],'
         '[[Meal Plan.date]],">="&TODAY()-30,[[Meal Plan.date]],'
         '"<="&TODAY())>0)*([[Recipes.name]]<>""))', "int",
         "the ones in active service"),
        ("Never cooked",
         '=COUNTIFS([[Recipes.name]],"<>",[[Recipes.times_planned]],0)',
         "int", "worth a go, or worth deleting"),
        ("Variety score",
         '=IFERROR(SUMPRODUCT((COUNTIFS([[Meal Plan.recipe]],'
         '[[Recipes.id]],[[Meal Plan.date]],">="&TODAY()-30,'
         '[[Meal Plan.date]],"<="&TODAY())>0)*'
         '([[Recipes.name]]<>""))/MAX(1,COUNTIFS([[Meal Plan.date]],'
         '">="&TODAY()-30,[[Meal Plan.date]],"<="&TODAY(),'
         '[[Meal Plan.recipe]],"<>")),"")', "percent0",
         "different recipes / meals cooked"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4
    row = s.note(row, "A variety score of 100% means every meal in the "
                      "last month was a different recipe; 25% means "
                      "you cooked the same four things. Neither is "
                      "right or wrong - but if it is low and you are "
                      "bored, the list underneath is where to look.",
                 LAST_COL, height=30)

    # ------------------------------------------- longest untouched ----
    row = s.section(row, "Not cooked for the longest", LAST_COL)
    row = s.table_head(row, 0, ["#", "Recipe", "Category", "Last cooked",
                                "Days ago", "Rating", "Cost a serving",
                                "Minutes"],
                       [5, 30, 15, 14, 12, 10, 15, 11])
    first = row
    for n in range(ROWS):
        pick = f'MATCH({n + 1},[[Recipes.rotation_rank]],0)'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "name", "text"),
                                (2, "category", "text"),
                                (3, "last_cooked", "date"),
                                (5, "rating", "rating"),
                                (6, "cost_serving", "money"),
                                (7, "total_min", "int")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Recipes.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IFERROR(IF(INDEX([[Recipes.last_cooked]],'
                         f'{pick})="","Never",TODAY()-INDEX('
                         f'[[Recipes.last_cooked]],{pick})),"")'),
            fmts.calc("int"))
        row += 1
    last = row - 1
    ws.conditional_format(first, 4, last, 4,
                          {"type": "text", "criteria": "containing",
                           "value": "Never", "format": fmts.cf("info")})
    ws.conditional_format(first, 5, last, 5,
                          {"type": "cell", "criteria": ">=", "value": 4,
                           "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------- most cooked ----
    row = s.section(row, "Your greatest hits", LAST_COL)
    row = s.table_head(row, 0, ["", "Recipe", "Times planned",
                                "Times cooked", "Rating",
                                "Cost a serving", "Last cooked"],
                       [5, 30, 15, 14, 10, 15, 14])
    for n in range(cfg.top_rows * 2):
        pick = f'MATCH({n + 1},[[Recipes.times_planned_rank]],0)'
        big = f'INDEX([[Recipes.times_planned]],{pick})'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "name", "text"),
                                (3, "times_cooked", "int"),
                                (4, "rating", "rating"),
                                (5, "cost_serving", "money"),
                                (6, "last_cooked", "date")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(IF({big}=0,"",INDEX('
                             f'[[Recipes.{key}]],{pick})),"")'),
                fmts.calc(style))
        ws.write_formula(row, 2,
                         s.reg.render(f'=IFERROR(IF({big}=0,"",{big}),"")'),
                         fmts.calc("int"))
        row += 1
    row += 1

    # ------------------------------------------------ variety by type -
    row = s.section(row, "Variety by cuisine and method", LAST_COL)
    row = s.table_head(row, 0, ["Cuisine", "Recipes", "Meals planned",
                                "Share of meals", "", "Method",
                                "Recipes", "Meals planned"],
                       [16, 10, 14, 14, 3, 16, 10, 14])
    cuisines = ["British", "Italian", "Indian", "Mexican", "Chinese",
                "Japanese", "Middle Eastern", "Thai", "American",
                "Other"]
    methods = ["Oven", "Hob", "One pot", "Slow cooker", "Air fryer",
               "No cook", "Grill", "Pressure cooker", "Microwave",
               "Barbecue"]
    for cuisine, method in zip(cuisines, methods):
        s.write(row, 0, cuisine, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Recipes.cuisine]],"{cuisine}")',
            f'=SUMPRODUCT(([[Recipes.cuisine]]="{cuisine}")*'
            f'IFERROR([[Recipes.times_planned]]*1,0))',
            f'=IFERROR(SUMPRODUCT(([[Recipes.cuisine]]="{cuisine}")*'
            f'IFERROR([[Recipes.times_planned]]*1,0))/'
            f'MAX(1,SUM(IFERROR([[Recipes.times_planned]]*1,0))),"")',
        ], ["int", "int", "percent0"])
        s.write(row, 5, method, fmts.chrome("label"))
        s.calc_row(row, 6, [
            f'=COUNTIFS([[Recipes.method]],"{method}")',
            f'=SUMPRODUCT(([[Recipes.method]]="{method}")*'
            f'IFERROR([[Recipes.times_planned]]*1,0))',
        ], ["int", "int"])
        row += 1
    row += 1

    # ------------------------------------------------ repeat warning --
    row = s.section(row, "Cooked twice inside the rotation gap", LAST_COL)
    row = s.note(row, f"Your rotation gap is set to "
                      f"{cfg.rotation_gap_days} days on the Settings "
                      f"tab. These recipes have come round faster than "
                      f"that - fine if they are favourites, worth "
                      f"swapping if they are just habit.",
                 LAST_COL, height=28)
    row = s.table_head(row, 0, ["Recipe", "Days since", "Times planned",
                                "Rating", "Suggested swap"],
                       [30, 13, 14, 10, 30])
    for n in range(cfg.top_rows):
        s.write(row, 0, None, fmts.calc("text"))
        rev = f'COUNTIF([[Recipes.name]],"<>")-{n}'
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX([[Recipes.name]],MATCH({rev},'
                         f'[[Recipes.rotation_rank]],0)),"")'),
            fmts.calc("text"))
        for col, key, style in [(1, "days_since", "int"),
                                (2, "times_planned", "int"),
                                (3, "rating", "rating")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Recipes.{key}]],'
                             f'MATCH({rev},[[Recipes.rotation_rank]],0)),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IFERROR(INDEX([[Recipes.name]],MATCH({n + 1},'
                         f'[[Recipes.rotation_rank]],0)),"")'),
            fmts.calc("text"))
        row += 1
    row += 1

    s.note(row, "The 'suggested swap' column is deliberately simple: "
                "it offers the recipe you have left alone the longest. "
                "A spreadsheet cannot know that nobody fancies fish on "
                "a Monday.", LAST_COL, height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/allergens.py`

*223 lines · 1 function*

> `Food Safety` - allergens, storage and the things worth checking.
>
> Three parts: a grid showing which allergens each person reacts to and
> how many recipes contain them, a storage and shelf-life reference, and
> the cooking temperatures worth keeping somewhere you can see them.
>
> Everything here is a prompt, not a guarantee. The workbook only knows
> the allergen tags you typed into the price book, and a tag is not a
> label.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Food Safety'), `LAST_COL` (9), `ALLERGENS` (14 items), `STORAGE` (5 items), `TEMPS` (8 items).

````python
"""`Food Safety` - allergens, storage and the things worth checking.

Three parts: a grid showing which allergens each person reacts to and
how many recipes contain them, a storage and shelf-life reference, and
the cooking temperatures worth keeping somewhere you can see them.

Everything here is a prompt, not a guarantee. The workbook only knows
the allergen tags you typed into the price book, and a tag is not a
label.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Food Safety"
LAST_COL = 9

ALLERGENS = ["Celery", "Cereals with gluten", "Crustaceans", "Eggs",
             "Fish", "Lupin", "Milk", "Molluscs", "Mustard", "Nuts",
             "Peanuts", "Sesame", "Soya", "Sulphites"]

STORAGE = [
    ("Fridge", "0-5C", "Milk, cooked leftovers, opened jars, raw meat "
                       "on the bottom shelf"),
    ("Freezer", "-18C", "Batch cooking, bread, herbs, anything near "
                        "its date"),
    ("Cupboard", "Cool and dark", "Tins, pasta, rice, flour, oil away "
                                  "from the hob"),
    ("Fruit bowl", "Room temperature", "Bananas, tomatoes, citrus, "
                                       "avocados"),
    ("Not the fridge", "Room temperature", "Potatoes, onions, garlic, "
                                           "bread, honey"),
]

TEMPS = [
    ("Poultry", "75C / 165F", "No pink, juices run clear"),
    ("Minced meat & burgers", "70C / 160F", "All the way through"),
    ("Pork", "70C / 160F", "Rest it for five minutes"),
    ("Whole cuts of beef & lamb", "55-70C", "Rare to well done - your "
                                            "call"),
    ("Fish", "63C / 145F", "Flakes easily"),
    ("Reheating anything", "75C / 165F", "Piping hot right through, "
                                         "once only"),
    ("Fridge", "0-5C", "Check it with a thermometer once a year"),
    ("Freezer", "-18C or colder", "Food stays safe indefinitely; "
                                  "quality drops"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Food Safety & Allergens",
             "Who reacts to what, where things live and how hot it "
             "needs to be.", LAST_COL)
    s.set_col(0, 0, 24)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 15)

    # ---------------------------------------------------- the headline --
    row = 4
    kpis = [
        ("People with an allergy",
         '=COUNTIFS([[Household.allergies]],"<>",'
         '[[Household.allergies]],"<>None")', "int",
         "from the Household tab"),
        ("Ingredients carrying an allergen",
         '=COUNTIFS([[Ingredients.allergen]],"<>",'
         '[[Ingredients.allergen]],"<>None")', "int",
         "tagged in the price book"),
        ("Recipes with a flagged ingredient",
         '=COUNTIF([[Recipes.allergens]],"*flagged*")', "int",
         "check these against the real labels"),
        ("Pantry items past their date",
         '=COUNTIF([[Pantry.status]],"Expired")', "int",
         "deal with these first"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ----------------------------------------------- the allergen grid --
    row = s.section(row, "Allergen grid", LAST_COL)
    row = s.note(row, "Each row is one of the fourteen allergens that "
                      "must be declared on food labels in the UK and "
                      "EU. 'In the house' counts the ingredients in "
                      "your price book carrying that tag.",
                 LAST_COL, height=26)
    head = row
    s.write(head, 0, "Allergen", fmts.chrome("header_left"))
    s.write(head, 1, "Ingredients", fmts.chrome("header"))
    s.write(head, 2, "Recipes", fmts.chrome("header"))
    s.write(head, 3, "In the pantry", fmts.chrome("header"))
    for i in range(5):
        ws.write_formula(head, 4 + i,
                         f'=IFERROR(IF(INDEX(MemberNames,{i + 1})=0,"",'
                         f'INDEX(MemberNames,{i + 1})),"")',
                         fmts.chrome("header"))
    s.row_height(head, 26)
    row = head + 1
    first_grid = row
    for allergen in ALLERGENS:
        s.write(row, 0, allergen, fmts.chrome("label_bold"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Ingredients.allergen]],"{allergen}")',
            f'=SUMPRODUCT((COUNTIFS([[Recipe Ingredients.recipe]],'
            f'[[Recipes.id]],[[Recipe Ingredients.allergen]],'
            f'"{allergen}")>0)*([[Recipes.name]]<>""))',
            f'=SUMPRODUCT((COUNTIFS([[Ingredients.name]],'
            f'[[Pantry.ingredient]])>0)*0)+COUNTIFS('
            f'[[Pantry.ingredient]],"<>")*0+SUMPRODUCT('
            f'([[Pantry.ingredient]]<>"")*(COUNTIFS('
            f'[[Ingredients.name]],[[Pantry.ingredient]],'
            f'[[Ingredients.allergen]],"{allergen}")>0))',
        ], ["int", "int", "int"])
        for i in range(5):
            person = f"{chr(69 + i)}${head + 1}"
            ws.write_formula(
                row, 4 + i,
                s.reg.render(f'=IF({person}="","",IF(IFERROR(INDEX('
                             f'[[Household.allergies]],MATCH({person},'
                             f'[[Household.name]],0)),"")="{allergen}",'
                             f'"AVOID","ok"))'),
                fmts.calc("flag"))
        row += 1
    ws.conditional_format(first_grid, 4, row - 1, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "AVOID", "format": fmts.cf("bad")})
    ws.conditional_format(first_grid, 4, row - 1, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "ok", "format": fmts.cf("muted")})
    ws.conditional_format(first_grid, 1, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#D9703C"})
    row += 1

    # --------------------------------------------- recipes to watch ----
    row = s.section(row, "Recipes to check before you serve them",
                    LAST_COL)
    row = s.table_head(row, 0, ["Recipe", "Allergen note", "Serves",
                                "Category", "Last cooked"],
                       [30, 26, 10, 15, 14])
    for n in range(cfg.top_rows * 2):
        pick = f'MATCH({n + 1},[[Recipes.cost_rank]],0)'
        for col, key, style in [(0, "name", "text"),
                                (1, "allergens", "text"),
                                (2, "servings", "int"),
                                (3, "category", "text"),
                                (4, "last_cooked", "date")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(IF(INDEX([[Recipes.allergens]],'
                             f'{pick})="None listed","",INDEX('
                             f'[[Recipes.{key}]],{pick})),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    # --------------------------------------------------- storage rules -
    row = s.section(row, "Where things live", LAST_COL)
    row = s.table_head(row, 0, ["Place", "Temperature", "What goes there",
                                "In your pantry now"],
                       [18, 18, 44, 18])
    for place, temp, what in STORAGE:
        s.write(row, 0, place, fmts.chrome("label_bold"))
        s.write(row, 1, temp, fmts.chrome("label"))
        s.merge(row, 2, LAST_COL - 1, what, fmts.chrome("label"))
        ws.write_formula(
            row, LAST_COL,
            s.reg.render(f'=COUNTIFS([[Pantry.location]],"{place}")'),
            fmts.calc("int"))
        row += 1
    row += 1

    # ------------------------------------------------------ temperatures
    row = s.section(row, "Cooked through means", LAST_COL)
    row = s.table_head(row, 0, ["What", "Temperature", "How you know"],
                       [28, 18, 50])
    for what, temp, how in TEMPS:
        s.write(row, 0, what, fmts.chrome("label_bold"))
        s.write(row, 1, temp, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, how, fmts.chrome("label"))
        row += 1
    row += 1

    # ------------------------------------------------- leftovers rules -
    row = s.section(row, "Leftovers, reheating and the freezer", LAST_COL)
    row = s.bullets(row, [
        "COOL IT FAST - get leftovers into the fridge within two "
        "hours, in a shallow container so they cool quickly.",
        f"THE FRIDGE WINDOW - this workbook assumes "
        f"{cfg.leftover_safe_days} days for cooked leftovers; change "
        f"it on the Settings tab if your household rule differs.",
        f"THE FREEZER WINDOW - {cfg.freezer_review_days} days is the "
        f"quality reminder, not a safety limit. Frozen food stays "
        f"safe; it just stops tasting of much.",
        "REHEAT ONCE - reheat a portion once, until it is piping hot "
        "all the way through. Not twice.",
        "DEFROST IN THE FRIDGE - overnight, not on the worktop. Plan "
        "it on the Meal Prep tab the day before.",
        "RICE - cool it and refrigerate it quickly, eat it within a "
        "day, reheat it once, properly.",
        "WHEN IN DOUBT - the use-by date on the packet beats anything "
        "this spreadsheet works out. Throw it out and log it on the "
        "Food Waste tab.",
    ], LAST_COL, height=26)
    row += 1

    s.note(row,
           "Plain English: this tab organises what you told it. It "
           "cannot see 'may contain' warnings, recipe changes by the "
           "manufacturer, or cross-contamination in your own kitchen. "
           "For a serious allergy, the label on the packet in your "
           "hand is the only source that counts - every time, even "
           "for a product you have bought for years.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/fridge.py`

*162 lines · 2 functions*

> `Fridge & Freezer` - what is in the cold bits, and how long you have.
>
> The Pantry tab holds everything; this page pulls out the fridge, the
> freezer and the leftovers and sorts them by urgency, because that is
> the list you actually want at five o'clock on a Tuesday.

| Defines | What it does |
|---------|--------------|
| `_pick(location, n)` | Nth item in `location`, soonest to expire first |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Fridge & Freezer'), `LAST_COL` (8), `ROWS` (16).

````python
"""`Fridge & Freezer` - what is in the cold bits, and how long you have.

The Pantry tab holds everything; this page pulls out the fridge, the
freezer and the leftovers and sorts them by urgency, because that is
the list you actually want at five o'clock on a Tuesday.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Fridge & Freezer"
LAST_COL = 8
ROWS = 16


def _pick(location: str, n: int) -> str:
    """Nth item in `location`, soonest to expire first."""
    return f'MATCH({n},[[Pantry.{location.lower()}_rank]],0)'


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Fridge & Freezer",
             "The cold storage list, sorted by what needs eating "
             "first.", LAST_COL)
    s.set_col(0, 0, 6)
    s.set_col(1, 1, 26)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 15)

    row = 4
    kpis = [
        ("In the fridge",
         '=COUNTIFS([[Pantry.location]],"Fridge")', "int",
         "items logged"),
        ("In the freezer",
         '=COUNTIFS([[Pantry.location]],"Freezer")', "int",
         "items logged"),
        ("Needs eating this week",
         '=COUNTIF([[Pantry.status]],"Use soon")', "int",
         f"within {cfg.expiry_warning_days} days of the date"),
        ("Value in cold storage",
         '=SUMIFS([[Pantry.value]],[[Pantry.location]],"Fridge")+'
         'SUMIFS([[Pantry.value]],[[Pantry.location]],"Freezer")',
         "money", "at price book prices"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    for location, blurb in [
            ("Fridge", "Soonest to expire at the top. Plan a meal "
                       "round the first three and the rest looks "
                       "after itself."),
            ("Freezer", "The freezer is where food goes to be "
                        "forgotten. Label everything, and work "
                        "through the top of this list once a month.")]:
        row = s.section(row, location, LAST_COL)
        row = s.note(row, blurb, LAST_COL, height=24)
        row = s.table_head(row, 0, ["#", "Item", "How much", "Unit",
                                    "Use by", "Days left", "Value",
                                    "Status"],
                           [5, 26, 12, 10, 13, 12, 12, 14])
        first = row
        for n in range(ROWS):
            pick = _pick(location, n + 1)
            s.write(row, 0, n + 1, fmts.chrome("label"))
            for col, key, style in [(1, "ingredient", "text"),
                                    (2, "qty", "number1"),
                                    (3, "unit", "text"),
                                    (4, "est_expiry", "date"),
                                    (5, "days_left", "int"),
                                    (6, "value", "money"),
                                    (7, "status", "status")]:
                ws.write_formula(
                    row, col,
                    s.reg.render(f'=IFERROR(INDEX([[Pantry.{key}]],'
                                 f'{pick}),"")'), fmts.calc(style))
            row += 1
        last = row - 1
        ws.conditional_format(first, 5, last, 5,
                              {"type": "cell", "criteria": "<",
                               "value": 0, "format": fmts.cf("bad")})
        ws.conditional_format(first, 5, last, 5,
                              {"type": "cell", "criteria": "between",
                               "minimum": 0,
                               "maximum": cfg.expiry_warning_days,
                               "format": fmts.cf("warn")})
        ws.conditional_format(first, 7, last, 7,
                              {"type": "text", "criteria": "containing",
                               "value": "Expired",
                               "format": fmts.cf("bad")})
        row += 1

    # ------------------------------------------------------- leftovers
    row = s.section(row, "Leftovers waiting to be eaten", LAST_COL)
    row = s.table_head(row, 0, ["#", "What it is", "Portions", "Where",
                                "Cooked", "Eat by", "Days left",
                                "Worth"],
                       [5, 28, 10, 12, 13, 13, 12, 12])
    first = row
    for n in range(8):
        pick = f'MATCH({n + 1},[[Leftovers.queue_rank]],0)'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "item", "text"),
                                (2, "portions", "number1"),
                                (3, "location", "text"),
                                (4, "cooked", "date"),
                                (5, "use_by", "date"),
                                (6, "days_left", "int"),
                                (7, "value", "money")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Leftovers.{key}]],'
                             f'{pick}),"")'), fmts.calc(style))
        row += 1
    ws.conditional_format(first, 6, row - 1, 6,
                          {"type": "cell", "criteria": "<=", "value": 1,
                           "format": fmts.cf("warn")})
    row += 1

    # ------------------------------------------------- defrost planner
    row = s.section(row, "Taking something out of the freezer", LAST_COL)
    row = s.table_head(row, 0, ["Size", "In the fridge",
                                "In cold water", "Never"],
                       [22, 22, 24, 30])
    defrost = [
        ("Small portion (1-2 servings)", "Overnight",
         "1-2 hours, change the water", "On the worktop all day"),
        ("Family batch (4-6 servings)", "24 hours",
         "Not really practical", "In the sun or a warm kitchen"),
        ("Whole chicken", "24-36 hours",
         "2-3 hours, change the water", "Refreezing once thawed"),
        ("Bread, cake, pastry", "A couple of hours out of the freezer",
         "-", "Worrying about it - it is fine"),
        ("Soup, stew, sauce", "Overnight",
         "Straight into the pan on a low heat",
         "Reheating twice"),
    ]
    for size, fridge, water, never in defrost:
        s.write(row, 0, size, fmts.chrome("label_bold"))
        s.write(row, 1, fridge, fmts.chrome("label"))
        s.merge(row, 2, 3, water, fmts.chrome("label"))
        s.merge(row, 4, LAST_COL, never, fmts.chrome("note"))
        row += 1
    row += 1

    s.note(row, "Days left comes from the use-by date you typed in, or "
                "from the shelf life on the Ingredients tab if you "
                "left it blank. Both are estimates - the packet and "
                "your own senses outrank this page every time.",
           LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/sustainability.py`

*197 lines · 1 function*

> `Sustainability` - the waste report, and what to do about it.
>
> Where the binned food comes from, what it costs a year, which reasons
> keep repeating, and a handful of changes that actually work. Reducing
> food waste is the one 'green' thing that also puts money straight back
> in your pocket, which is why it has a tab of its own.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Sustainability'), `LAST_COL` (8), `REASONS` (10 items), `HABITS` (7 items).

````python
"""`Sustainability` - the waste report, and what to do about it.

Where the binned food comes from, what it costs a year, which reasons
keep repeating, and a handful of changes that actually work. Reducing
food waste is the one 'green' thing that also puts money straight back
in your pocket, which is why it has a tab of its own.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Sustainability"
LAST_COL = 8

REASONS = ["Expired", "Went mouldy", "Cooked too much",
           "Did not like it", "Forgot about it", "Freezer burn",
           "Left out too long", "Bought too much", "Plate scrapings",
           "Other"]

HABITS = [
    ("Shop your own cupboards first",
     "The Pantry deduction on the grocery list does this for you - "
     "the saving only lands if you keep the pantry roughly honest."),
    ("Freeze half on the day you buy it",
     "Bread, milk, cooked rice, herbs, cheese. The freezer window on "
     "the Settings tab tracks it."),
    ("Plan one 'use it up' meal a week",
     "Soup, fried rice, frittata, traybake. Add it to the Meal Plan "
     "with no recipe and the ingredients come out of the fridge "
     "rather than the shop."),
    ("Buy the smaller bag of salad",
     "The big bag is cheaper per 100 g only if you eat all of it. "
     "Check the Food Waste log before you decide."),
    ("Use the Leftovers tab as a menu",
     "Portions with a value attached get eaten; mystery boxes get "
     "binned."),
    ("Keep a 'eat me first' shelf",
     "One shelf, everything close to its date. Costs nothing, works "
     "better than any app."),
    ("Cook from the freezer once a week",
     "The Fridge & Freezer page sorts it by how long it has been in "
     "there."),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Food Waste & Sustainability",
             "What goes in the bin, what it costs, and the handful of "
             "habits that change it.", LAST_COL)
    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 15)

    row = 4
    kpis = [
        ("Wasted this month",
         '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
         'TEXT(TODAY(),"yyyy-mm"))', "money",
         f"target under {cfg.currency_symbol}"
         f"{cfg.waste_cost_target:,.0f}"),
        ("A year at this rate",
         '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
         'TEXT(TODAY(),"yyyy-mm"))*12', "money",
         "the number worth staring at"),
        ("Avoidable share",
         '=IFERROR(SUMIFS([[Food Waste.value]],'
         '[[Food Waste.avoidable]],"Yes")/SUM([[Food Waste.value]]),"")',
         "percent0", "peelings and bones are not avoidable"),
        ("Waste as a share of the shop",
         '=IFERROR(SUM([[Food Waste.value]])/'
         'SUM([[Grocery Spending.amount]]),"")', "percent0",
         f"target under {cfg.waste_reduction_target:.0%}"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ------------------------------------------------------ by reason --
    row = s.section(row, "Why food gets thrown away", LAST_COL)
    row = s.table_head(row, 0, ["Reason", "Times", "Cost",
                                "Share of the waste", "Avoidable?",
                                "Most common item"],
                       [24, 10, 13, 18, 12, 24])
    first = row
    for reason in REASONS:
        s.write(row, 0, reason, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Food Waste.reason]],"{reason}")',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.reason]],'
            f'"{reason}")',
            f'=IFERROR(SUMIFS([[Food Waste.value]],'
            f'[[Food Waste.reason]],"{reason}")/'
            f'SUM([[Food Waste.value]]),"")',
            f'=IF(COUNTIFS([[Food Waste.reason]],"{reason}")=0,"",'
            f'IF(COUNTIFS([[Food Waste.reason]],"{reason}",'
            f'[[Food Waste.avoidable]],"Yes")>0,"Yes","No"))',
            f'=IFERROR(INDEX([[Food Waste.item]],MATCH("{reason}",'
            f'[[Food Waste.reason]],0)),"")',
        ], ["int", "money", "percent0", "flag", "text"])
        row += 1
    ws.conditional_format(first, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#D9703C"})
    row += 1

    # ------------------------------------------------- worst offenders
    row = s.section(row, "The things you bin most", LAST_COL)
    row = s.table_head(row, 0, ["", "Item", "Times binned",
                                "Cost so far", "Still buying it?",
                                "What would stop it"],
                       [5, 24, 14, 14, 16, 34])
    for n in range(cfg.top_rows):
        pick = f'MATCH({n + 1},[[Ingredients.waste_rank]],0)'
        big = f'INDEX([[Ingredients.waste_count]],{pick})'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IFERROR(IF({big}=0,"",INDEX('
                         f'[[Ingredients.name]],{pick})),"")'),
            fmts.calc("text"))
        ws.write_formula(
            row, 2, s.reg.render(f'=IFERROR(IF({big}=0,"",{big}),"")'),
            fmts.calc("int"))
        item = f"B{row + 1}"
        s.calc_row(row, 3, [
            f'=IF({item}="","",SUMIFS([[Food Waste.value]],'
            f'[[Food Waste.item]],{item}))',
            f'=IF({item}="","",IF(COUNTIFS([[Grocery List.ingredient]],'
            f'{item})>0,"On this week\'s list","Not this week"))',
            f'=IF({item}="","",IFERROR(INDEX([[Food Waste.prevent]],'
            f'MATCH({item},[[Food Waste.item]],0)),""))',
        ], ["money", "text", "text"])
        row += 1
    row += 1

    # ---------------------------------------------------- disposal ----
    row = s.section(row, "Where it ends up", LAST_COL)
    row = s.table_head(row, 0, ["Disposal route", "Times", "Cost",
                                "Share"], [24, 12, 14, 14])
    for route in ["General bin", "Food waste caddy", "Compost",
                  "Fed to pets", "Given away"]:
        s.write(row, 0, route, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Food Waste.disposal]],"{route}")',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.disposal]],'
            f'"{route}")',
            f'=IFERROR(COUNTIFS([[Food Waste.disposal]],"{route}")/'
            f'COUNTIF([[Food Waste.date]],">0"),"")',
        ], ["int", "money", "percent0"])
        row += 1
    row += 1

    # -------------------------------------------------- what to change
    row = s.section(row, "Seven things that actually work", LAST_COL)
    for habit, detail in HABITS:
        s.write(row, 0, habit, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, detail, fmts.chrome("note"))
        s.row_height(row, 28)
        row += 1
    row += 1

    # ---------------------------------------------------- seasonality -
    row = s.section(row, "Eating with the seasons", LAST_COL)
    row = s.table_head(row, 0, ["Season", "Ingredients tagged",
                                "Recipes tagged", "On the list now"],
                       [16, 20, 18, 18])
    for season in ["All year", "Spring", "Summer", "Autumn", "Winter"]:
        s.write(row, 0, season, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Ingredients.season]],"{season}")',
            f'=COUNTIFS([[Recipes.season]],"{season}")',
            f'=SUMPRODUCT(([[Grocery List.ingredient]]<>"")*'
            f'(COUNTIFS([[Ingredients.name]],'
            f'[[Grocery List.ingredient]],[[Ingredients.season]],'
            f'"{season}")>0))',
        ], ["int", "int", "int"])
        row += 1
    row += 1

    s.note(row,
           "The honest version: a spreadsheet cannot measure your "
           "carbon footprint, and anyone who says their meal planner "
           "can is guessing. What this page does measure is money in "
           "the bin, which is the part you control and the part that "
           "changes behaviour. Log a month properly and the pattern "
           "is usually obvious - and usually three or four items.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/reports.py`

*294 lines · 1 function*

> `Reports` - the month in one page, and the year in one table.
>
> Pick a month at the top and every table underneath follows it: what
> you spent, what you cooked, what you wasted, how the week-by-week
> pattern moved. Underneath that, a twelve row year table you can copy
> straight into a message to whoever else pays for the food.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Reports'), `LAST_COL` (9).

````python
"""`Reports` - the month in one page, and the year in one table.

Pick a month at the top and every table underneath follows it: what
you spent, what you cooked, what you wasted, how the week-by-week
pattern moved. Underneath that, a twelve row year table you can copy
straight into a message to whoever else pays for the food.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from
#: the Budget tab owns the list; the two pages must agree
from .budget import SPEND_TYPES

NAME = "Reports"
LAST_COL = 9


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Reporting & Analytics",
             "Pick a month; every table on this page follows it.",
             LAST_COL)
    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 14)

    # ------------------------------------------------------ the picker --
    row = 4
    s.write(row, 0, "Month", fmts.chrome("label_bold"))
    s.write(row, 1, "", fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("MonthList"))
    book.define_name("ReportMonth", f"={quote_sheet(NAME)}!$B${row + 1}")
    ws.write_formula(row, 2,
                     '=IFERROR(DATE(VALUE(LEFT(ReportMonth,4)),'
                     'VALUE(RIGHT(ReportMonth,2)),1),'
                     'DATE(YEAR(TODAY()),MONTH(TODAY()),1))',
                     fmts.calc("date"))
    book.define_name("ReportStart", f"={quote_sheet(NAME)}!$C${row + 1}")
    s.merge(row, 3, LAST_COL,
            "Leave it blank and the page shows the month you are in. "
            "The list comes from the Lists tab.", fmts.chrome("note"))
    ws.write_formula(row, 11, s.reg.render('=MAX(1,(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,1,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,2,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,3,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,4,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,5,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,6,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,7,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,8,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,9,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,10,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,11,1),"yyyy-mm"))>0)*1+(COUNTIFS([[Grocery Spending.month]],TEXT(DATE(PlanYear,12,1),"yyyy-mm"))>0)*1)'),
                     fmts.calc("int"))
    book.define_name("ReportMonths", f"={quote_sheet(NAME)}!$L${row + 1}")
    ws.set_column(11, 11, None, None, {"hidden": 1})
    row += 2

    M = 'TEXT(ReportStart,"yyyy-mm")'

    kpis = [
        ("Total food spend",
         f'=SUMIFS([[Grocery Spending.amount]],'
         f'[[Grocery Spending.month]],{M})', "money", "every receipt"),
        ("Cost per head per day",
         f'=IFERROR(SUMIFS([[Grocery Spending.amount]],'
         f'[[Grocery Spending.month]],{M})/HouseMembers/'
         f'DAY(EOMONTH(ReportStart,0)),"")', "money",
         "the comparable number"),
        ("Meals cooked at home",
         f'=SUMIFS([[Meal Plan.home_cooked]],[[Meal Plan.month]],{M})',
         "int", "from the meal plan"),
        ("Food binned",
         f'=SUMIFS([[Food Waste.value]],[[Food Waste.month]],{M})',
         "money", "the avoidable part hurts"),
        ("Against budget",
         f'=IFERROR(SUMIFS([[Grocery Spending.amount]],'
         f'[[Grocery Spending.month]],{M})/MonthlyBudget,"")',
         "percent0", "under 100% is the win"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ------------------------------------------------- spend breakdown --
    row = s.section(row, "Where the money went", LAST_COL)
    row = s.table_head(row, 0, ["Type of spend", "This month", "Share",
                                "Last month", "Change",
                                "Monthly average"],
                       [26, 14, 12, 14, 14, 16])
    first = row
    for kind in SPEND_TYPES:
        s.write(row, 0, kind, fmts.chrome("label"))
        this = (f'SUMIFS([[Grocery Spending.amount]],'
                f'[[Grocery Spending.month]],{M},'
                f'[[Grocery Spending.type]],"{kind}")')
        prev = (f'SUMIFS([[Grocery Spending.amount]],'
                f'[[Grocery Spending.month]],'
                f'TEXT(EDATE(ReportStart,-1),"yyyy-mm"),'
                f'[[Grocery Spending.type]],"{kind}")')
        s.calc_row(row, 1, [
            f'={this}',
            f'=IFERROR({this}/SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{M}),"")',
            f'={prev}',
            f'=IF({prev}=0,"",{this}/{prev}-1)',
            f'=IFERROR(SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.type]],"{kind}")/ReportMonths,"")',
        ], ["money", "percent0", "money", "percent", "money"])
        row += 1
    s.write(row, 0, "Total", fmts.chrome("total_label"))
    for col in (1, 3):
        letter = chr(ord("A") + col)
        ws.write_formula(row, col,
                         f"=SUM({letter}{first + 1}:{letter}{row})",
                         fmts.chrome("total_value",
                                     num_format=cfg.money_format))
    for col in (2, 4, 5):
        s.write(row, col, None, fmts.chrome("total_value"))
    ws.conditional_format(first, 4, row - 1, 4,
                          {"type": "cell", "criteria": ">",
                           "value": 0.1, "format": fmts.cf("bad")})
    ws.conditional_format(first, 4, row - 1, 4,
                          {"type": "cell", "criteria": "<",
                           "value": -0.1, "format": fmts.cf("good")})
    row += 2

    # ------------------------------------------------ cooking the month --
    row = s.section(row, "What the month looked like in the kitchen",
                    LAST_COL)
    row = s.table_head(row, 0, ["Measure", "This month", "Target",
                                "Verdict", "Notes"],
                       [30, 14, 14, 16, 40])
    lines = [
        ("Meals planned",
         f'=COUNTIFS([[Meal Plan.month]],{M})', '=""', "int",
         "every slot with something in it"),
        ("Dinners cooked at home",
         f'=COUNTIFS([[Meal Plan.month]],{M},[[Meal Plan.slot]],'
         f'"Dinner",[[Meal Plan.home_cooked]],1)',
         '=HomeCookedTarget*DAY(EOMONTH(ReportStart,0))/7', "int",
         "the headline habit"),
        ("Takeaways and meals out",
         f'=COUNTIFS([[Meal Plan.month]],{M},[[Meal Plan.home_cooked]],0)',
         '=TakeawayLimit*DAY(EOMONTH(ReportStart,0))/7', "int",
         "no judgement, just the number"),
        ("Different recipes used",
         f'=SUMPRODUCT((COUNTIFS([[Meal Plan.month]],{M},'
         f'[[Meal Plan.recipe]],[[Recipes.id]])>0)*1)',
         '=""', "int", "variety, not volume"),
        ("New recipes added",
         '=COUNTIFS([[Recipes.added]],">="&ReportStart,'
         '[[Recipes.added]],"<="&EOMONTH(ReportStart,0))',
         '=NewRecipeTarget', "int", "one a month is plenty"),
        ("Packed lunches",
         f'=COUNTIFS([[Meal Plan.month]],{M},[[Meal Plan.slot]],'
         f'"Lunch",[[Meal Plan.home_cooked]],1)',
         '=PackedLunchTarget*DAY(EOMONTH(ReportStart,0))/7', "int",
         "where the easy money is"),
        ("Batch cook sessions",
         '=COUNTIFS([[Meal Prep.date]],">="&ReportStart,'
         '[[Meal Prep.date]],"<="&EOMONTH(ReportStart,0))',
         '=""', "int", "from the Meal Prep tab"),
        ("Portions frozen",
         '=SUMIFS([[Meal Prep.portions]],[[Meal Prep.date]],">="&'
         'ReportStart,[[Meal Prep.date]],"<="&EOMONTH(ReportStart,0))',
         '=""', "number1", "future you says thanks"),
    ]
    for label, formula, target, style, note in lines:
        s.write(row, 0, label, fmts.chrome("label"))
        s.calc_row(row, 1, [
            formula, target,
            f'=IF(C{row + 1}="","",IF(B{row + 1}>=C{row + 1},'
            f'"On target","Short"))',
        ], [style, style, "status"])
        s.write(row, 4, note, fmts.chrome("note"))
        row += 1
    row += 2

    # ------------------------------------------------------ week by week --
    row = s.section(row, "Week by week inside the month", LAST_COL)
    row = s.table_head(row, 0, ["Week beginning", "Meals planned",
                                "Planned cost", "Spent", "Wasted",
                                "Cost per head"],
                       [18, 14, 14, 14, 14, 16])
    for n in range(5):
        start = f"A{row + 1}"
        ws.write_formula(
            row, 0,
            f'=IF(ReportStart+{n * 7}>EOMONTH(ReportStart,0),"",'
            f'ReportStart+{n * 7})', fmts.calc("date"))
        s.calc_row(row, 1, [
            f'=IF({start}="","",COUNTIFS([[Meal Plan.date]],">="&{start},'
            f'[[Meal Plan.date]],"<="&{start}+6))',
            f'=IF({start}="","",SUMIFS([[Meal Plan.cost]],'
            f'[[Meal Plan.date]],">="&{start},[[Meal Plan.date]],'
            f'"<="&{start}+6))',
            f'=IF({start}="","",SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.date]],">="&{start},'
            f'[[Grocery Spending.date]],"<="&{start}+6))',
            f'=IF({start}="","",SUMIFS([[Food Waste.value]],'
            f'[[Food Waste.date]],">="&{start},[[Food Waste.date]],'
            f'"<="&{start}+6))',
            f'=IF({start}="","",IFERROR(SUMIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.date]],'
            f'">="&{start},[[Grocery Spending.date]],"<="&{start}+6)/'
            f'HouseMembers,""))',
        ], ["int", "money", "money", "money", "money"])
        row += 1
    row += 2

    # ------------------------------------------------------- year table --
    row = s.section(row, "The year in one table", LAST_COL)
    row = s.table_head(row, 0, ["Month", "Spend", "Budget", "Over/under",
                                "Meals at home", "Waste",
                                "Cost per head per day", "Takeaways",
                                "Shops"],
                       [14, 13, 13, 13, 14, 12, 18, 12, 10])
    y_first = row
    for m in range(12):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        ws.write_formula(row, 0, f"=DATE(PlanYear,{m + 1},1)",
                         fmts.calc("text", num_format="mmmm"))
        spend = (f'SUMIFS([[Grocery Spending.amount]],'
                 f'[[Grocery Spending.month]],{key})')
        s.calc_row(row, 1, [
            f'={spend}',
            '=MonthlyBudget',
            f'=MonthlyBudget-{spend}',
            f'=SUMIFS([[Meal Plan.home_cooked]],[[Meal Plan.month]],'
            f'{key})',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.month]],{key})',
            f'=IFERROR({spend}/HouseMembers/DAY(DATE(PlanYear,'
            f'{m + 2},0)),"")',
            f'=COUNTIFS([[Meal Plan.month]],{key},'
            f'[[Meal Plan.home_cooked]],0)',
            f'=COUNTIFS([[Grocery Spending.month]],{key})',
        ], ["money", "money", "money", "int", "money", "money", "int",
            "int"])
        row += 1
    s.write(row, 0, "Year", fmts.chrome("total_label"))
    for col in (1, 2, 3, 4, 5, 7, 8):
        letter = chr(ord("A") + col)
        style = "int" if col in (4, 7, 8) else "money"
        ws.write_formula(
            row, col, f"=SUM({letter}{y_first + 1}:{letter}{row})",
            fmts.chrome("total_value",
                        num_format=s.num_format(style)))
    ws.write_formula(
        row, 6, f"=IFERROR(AVERAGE(G{y_first + 1}:G{row}),0)",
        fmts.chrome("total_value", num_format=cfg.money_format))
    ws.conditional_format(y_first, 1, row - 1, 1,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    ws.conditional_format(y_first, 3, row - 1, 3,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row += 2

    # ---------------------------------------------------- store summary --
    row = s.section(row, "Shops, ranked by what you spend there",
                    LAST_COL)
    row = s.table_head(row, 0, ["Store", "Visits", "Total spent",
                                "Average basket", "Share of spend",
                                "Price level"],
                       [24, 12, 14, 16, 14, 14])
    for n in range(8):
        pick = f'MATCH({n + 1},[[Stores.spend_rank]],0)'
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX([[Stores.name]],{pick}),"")'),
            fmts.calc("text"))
        store = f"A{row + 1}"
        s.calc_row(row, 1, [
            f'=IF({store}="","",COUNTIFS([[Grocery Spending.store]],'
            f'{store}))',
            f'=IF({store}="","",SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.store]],{store}))',
            f'=IF({store}="","",IFERROR(SUMIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.store]],'
            f'{store})/COUNTIFS([[Grocery Spending.store]],{store}),""))',
            f'=IF({store}="","",IFERROR(SUMIFS('
            f'[[Grocery Spending.amount]],[[Grocery Spending.store]],'
            f'{store})/SUM([[Grocery Spending.amount]]),""))',
            f'=IF({store}="","",IFERROR(INDEX([[Stores.price_level]],'
            f'MATCH({store},[[Stores.name]],0)),""))',
        ], ["int", "money", "money", "percent0", "text"])
        row += 1
    row += 1

    s.note(row,
           "Copy any of these tables straight out - they are plain "
           "cells, not a pivot table, so paste works properly. If a "
           "month looks empty, check the Grocery Spending tab has "
           "dates in it: every table on this page keys off the month "
           "column, which is worked out from the date you type.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/printables.py`

*359 lines · 3 functions*

> The three printable pages.
>
> * `Print Week` - the week's meals on one sheet of A4 for the fridge
>   door, with the prep jobs and the leftovers to eat.
> * `Print Shopping` - the grocery list grouped by aisle with tick boxes
>   and a running total, sized for a pocket.
> * `Print Recipe` - one recipe as a card: ingredients scaled to the
>   number of people you choose, method, cost and nutrition.
>
> All three read from the same registers as everything else, so they are
> never a stale copy.

| Defines | What it does |
|---------|--------------|
| `build_week(book, fmts, cfg, reg)` |  |
| `build_shopping(book, fmts, cfg, reg)` |  |
| `build_recipe(book, fmts, cfg, reg)` |  |

**Module constants:** `WEEK_NAME` ('Print Week'), `SHOP_NAME` ('Print Shopping'), `RECIPE_NAME` ('Print Recipe').

````python
"""The three printable pages.

* `Print Week` - the week's meals on one sheet of A4 for the fridge
  door, with the prep jobs and the leftovers to eat.
* `Print Shopping` - the grocery list grouped by aisle with tick boxes
  and a running total, sized for a pocket.
* `Print Recipe` - one recipe as a card: ingredients scaled to the
  number of people you choose, method, cost and nutrition.

All three read from the same registers as everything else, so they are
never a stale copy.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from, number_check

WEEK_NAME = "Print Week"
SHOP_NAME = "Print Shopping"
RECIPE_NAME = "Print Recipe"


# ---------------------------------------------------------------------------
# Print Week
# ---------------------------------------------------------------------------

def build_week(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = WEEK_NAME, 5
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("This Week's Meals", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 14)
    s.set_col(1, 1, 34)
    s.set_col(2, 2, 30)
    s.set_col(3, 3, 12)
    s.set_col(4, 5, 14)

    row = 4
    s.merge(row, 0, 1, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0,
                     '=HouseName&" - week of "&TEXT(WeekPick,'
                     '"d mmmm yyyy")', fmts.chrome("print_title"))
    s.merge(row, 2, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 2,
                     s.reg.render('="Planned cost "&TEXT(SUMIFS('
                                  '[[Meal Plan.cost]],[[Meal Plan.date]],'
                                  '">="&WeekPick,[[Meal Plan.date]],'
                                  '"<="&WeekPick+6),"0.00")&'
                                  '"   |   Shopping day: "&ShoppingDay&'
                                  '"   |   Prep day: "&PrepDay'),
                     fmts.chrome("print_sub"))
    row += 2

    row = s.table_head(row, 0, ["Day", "Dinner", "Lunch", "Cost",
                                "Who is cooking", "Done"],
                       [14, 34, 30, 12, 16, 10])
    first = row
    for i in range(7):
        day = f"A{row + 1}"
        ws.write_formula(row, 0, f"=WeekPick+{i}",
                         fmts.chrome("label_bold", num_format="ddd d mmm"))
        for col, slot in ((1, "Dinner"), (2, "Lunch")):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Meal Plan.recipe_name]],'
                             f'MATCH(TEXT({day},"yyyy-mm-dd")&"|{slot}",'
                             f'[[Meal Plan.key]],0)),"")'),
                fmts.cell("wrap", "plain"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF(SUMIFS([[Meal Plan.cost]],'
                         f'[[Meal Plan.date]],{day})=0,"",SUMIFS('
                         f'[[Meal Plan.cost]],[[Meal Plan.date]],'
                         f'{day}))'),
            fmts.cell("money", "plain"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IFERROR(INDEX([[Meal Plan.cook]],MATCH('
                         f'TEXT({day},"yyyy-mm-dd")&"|Dinner",'
                         f'[[Meal Plan.key]],0)),"")'),
            fmts.cell("text", "plain"))
        s.write(row, 5, None, fmts.chrome("form_box"))
        s.row_height(row, 30)
        row += 1
    s.merge(row, 0, 2, "Week total", fmts.chrome("total_label"))
    ws.write_formula(row, 3,
                     f"=SUM(D{first + 1}:D{row})",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    s.merge(row, 4, LAST_COL, None, fmts.chrome("total_value"))
    row += 2

    row = s.section(row, "Before the week starts", LAST_COL)
    row = s.table_head(row, 0, ["Prep job", "When", "Portions",
                                "Who", "Done"], [34, 14, 12, 16, 10])
    for n in range(5):
        pick = f'MATCH({n + 1},[[Meal Prep.todo_rank]],0)'
        for col, key, style in [(0, "job", "text"), (1, "date", "date"),
                                (2, "portions", "number1"),
                                (3, "cook", "text")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Meal Prep.{key}]],'
                             f'{pick}),"")'), fmts.cell(style, "plain"))
        s.write(row, 4, None, fmts.chrome("form_box"))
        row += 1
    row += 1

    row = s.section(row, "Eat these first", LAST_COL)
    row = s.table_head(row, 0, ["Leftover or item", "Where", "Eat by",
                                "Portions", "Done"],
                       [34, 14, 14, 12, 10])
    for n in range(5):
        pick = f'MATCH({n + 1},[[Leftovers.queue_rank]],0)'
        for col, key, style in [(0, "item", "text"),
                                (1, "location", "text"),
                                (2, "use_by", "date"),
                                (3, "portions", "number1")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Leftovers.{key}]],'
                             f'{pick}),"")'), fmts.cell(style, "plain"))
        s.write(row, 4, None, fmts.chrome("form_box"))
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0,
                     f'="{cfg.product_name}  |  "&HouseName&"  |  "&'
                     f'TEXT(TODAY(),"d mmm yyyy")',
                     fmts.chrome("print_sub"))

    s.finish(freeze_row=None, print_area=f"A1:F{row + 1}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Shopping
# ---------------------------------------------------------------------------

def build_shopping(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = SHOP_NAME, 5
    ROWS = 45
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Shopping List", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 6)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 18)
    s.set_col(3, 3, 14)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 10)

    row = 4
    s.merge(row, 0, 2, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0,
                     '=HouseName&" - "&TEXT(ShopFrom,"d mmm")&" to "&'
                     'TEXT(ShopTo,"d mmm")', fmts.chrome("print_title"))
    s.merge(row, 3, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 3,
                     s.reg.render('="Estimated "&TEXT(SUM('
                                  '[[Grocery List.est_cost]]),"0.00")&'
                                  '"  |  budget "&TEXT(WeeklyBudget,'
                                  '"0.00")'),
                     fmts.chrome("print_sub"))
    row += 2

    row = s.table_head(row, 0, ["Got", "Item", "Aisle", "How much",
                                "Cost", "Store"],
                       [6, 30, 18, 14, 12, 14])
    first = row
    for n in range(ROWS):
        pick = f'MATCH({n + 1},[[Grocery List.buy_rank]],0)'
        s.write(row, 0, None, fmts.chrome("form_box"))
        for col, key, style in [(1, "ingredient", "text"),
                                (2, "aisle", "text"),
                                (4, "est_cost", "money"),
                                (5, "store", "text")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Grocery List.{key}]],'
                             f'{pick}),"")'), fmts.cell(style, "plain"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IFERROR(IF(INDEX([[Grocery List.buy_qty]],'
                         f'{pick})="","",ROUND(INDEX('
                         f'[[Grocery List.buy_qty]],{pick}),0)&" "&'
                         f'INDEX([[Grocery List.unit]],{pick})),"")'),
            fmts.cell("text", "plain"))
        s.row_height(row, 18)
        row += 1
    s.merge(row, 0, 3, "Total", fmts.chrome("total_label"))
    ws.write_formula(row, 4, f"=SUM(E{first + 1}:E{row})",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    s.merge(row, 5, LAST_COL, None, fmts.chrome("total_value"))
    row += 2

    row = s.section(row, "Extras", LAST_COL)
    for n in range(6):
        s.write(row, 0, None, fmts.chrome("form_box"))
        s.merge(row, 1, 3, None, fmts.chrome("form_line"))
        s.merge(row, 4, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 20)
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL,
            "Printed from the Grocery List tab - change the shopping "
            "window on Settings and print again.",
            fmts.chrome("print_sub"))

    s.finish(freeze_row=None, print_area=f"A1:F{row + 1}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Recipe
# ---------------------------------------------------------------------------

def build_recipe(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = RECIPE_NAME, 5
    LINES = 16
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.accent_soft)
    ws = s.ws

    s.banner("Recipe Card", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 26)
    s.set_col(1, 1, 14)
    s.set_col(2, 2, 12)
    s.set_col(3, 3, 22)
    s.set_col(4, 5, 16)

    # the two controls
    row = 4
    s.merge(row, 0, 0, "Recipe", fmts.chrome("label_bold"))
    s.write(row, 1, "R-0001" if cfg.include_sample_data else "",
            fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("RecipeIDs"))
    book.define_name("CardRecipe", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "For how many", fmts.chrome("label_right"))
    s.write(row, 3, cfg.default_servings, fmts.input("int"))
    ws.data_validation(row, 3, row, 3, number_check(1, 200,
                                                    decimal=False))
    book.define_name("CardPeople", f"={quote_sheet(NAME)}!$D${row + 1}")
    s.merge(row, 4, LAST_COL, "Pick a recipe ID, set the number of "
                              "people, print.", fmts.chrome("note"))
    row += 1

    base = 'MATCH(CardRecipe,' + s.reg.render('[[Recipes.id]]') + ',0)'
    ws.write_formula(row, 7,
                     s.reg.render(f'=IFERROR(CardPeople/INDEX('
                                  f'[[Recipes.servings]],{base}),1)'),
                     fmts.calc("number1"))
    book.define_name("CardFactor", f"={quote_sheet(NAME)}!$H${row + 1}")
    ws.set_column(7, 7, None, None, {"hidden": 1})
    row += 1

    # the title block
    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0,
                     s.reg.render(f'=IFERROR(INDEX([[Recipes.name]],'
                                  f'{base}),"Choose a recipe above")'),
                     fmts.chrome("print_title"))
    s.row_height(row, 30)
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(
        row, 0,
        s.reg.render(f'=IFERROR(INDEX([[Recipes.category]],{base})&'
                     f'"  |  "&INDEX([[Recipes.cuisine]],{base})&'
                     f'"  |  serves "&TEXT(CardPeople,"0")&'
                     f'"  |  "&TEXT(INDEX([[Recipes.total_min]],{base}),'
                     f'"0")&" minutes  |  "&TEXT(INDEX('
                     f'[[Recipes.cost_serving]],{base})*1,"0.00")&'
                     f'" a serving","")'), fmts.chrome("print_sub"))
    row += 2

    # ingredients
    row = s.table_head(row, 0, ["Ingredient", "You need", "Unit",
                                "Prepared", "Optional?"],
                       [26, 14, 12, 22, 12])
    for n in range(LINES):
        pick = (f'MATCH(CardRecipe&"#"&{n + 1},'
                f'[[Recipe Ingredients.scale_key]],0)')
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX('
                         f'[[Recipe Ingredients.ingredient]],{pick}),"")'),
            fmts.cell("text", "plain"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IFERROR(ROUND(INDEX('
                         f'[[Recipe Ingredients.qty]],{pick})*'
                         f'CardFactor,2),"")'),
            fmts.cell("number1", "plain"))
        for col, key in ((2, "unit"), (3, "prep"), (4, "optional")):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX('
                             f'[[Recipe Ingredients.{key}]],{pick}),"")'),
                fmts.cell("text", "plain"))
        s.row_height(row, 18)
        row += 1
    row += 1

    # method
    row = s.section(row, "Method", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0,
                     s.reg.render(f'=IFERROR(INDEX('
                                  f'[[Recipes.method_notes]],{base}),"")'),
                     fmts.cell("wrap", "plain"))
    s.row_height(row, 90)
    row += 2

    # the numbers
    row = s.section(row, "Per serving", LAST_COL)
    row = s.table_head(row, 0, ["Cost", "Calories", "Protein (g)",
                                "Carbs (g)", "Fat (g)", "Fibre (g)"])
    s.calc_row(row, 0, [
        s.reg.render(f'=IFERROR(INDEX([[Recipes.cost_serving]],{base}),'
                     f'"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.calories]],{base}),"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.protein]],{base}),"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.carbs]],{base}),"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.fat]],{base}),"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.fibre]],{base}),"")'),
    ], ["money", "number", "number1", "number1", "number1", "number1"],
        mode="plain")
    row += 2

    s.merge(row, 0, LAST_COL, None, fmts.chrome("note_box"))
    ws.write_formula(
        row, 0,
        s.reg.render(f'=IFERROR("Allergens: "&INDEX('
                     f'[[Recipes.allergens]],{base})&"   |   Notes: "&'
                     f'INDEX([[Recipes.notes]],{base}),"")'),
        fmts.chrome("note_box"))
    s.row_height(row, 32)
    row += 2

    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0,
                     f'=HouseName&"  |  {cfg.product_name}  |  "&'
                     f'TEXT(TODAY(),"d mmm yyyy")',
                     fmts.chrome("print_sub"))

    s.finish(freeze_row=None, print_area=f"A1:F{row + 1}",
             gridlines=False, fit_width=True)
    return s
````

### `meal_planner_spreadsheet/sheets/help_tab.py`

*282 lines · 1 function*

> `Help` - how the thing works, in the order you will need it.
>
> Setup in ten minutes, then a tab-by-tab reference, the automation
> chain written out in plain English, what to do when a number looks
> wrong, the honest limits, and the keyboard shortcuts worth knowing.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Help'), `LAST_COL` (7), `SETUP` (7 items), `TABS` (35 items), `CHAIN` (7 items), `TROUBLE` (7 items), `LIMITS` (5 items), `SHORTCUTS` (7 items).

````python
"""`Help` - how the thing works, in the order you will need it.

Setup in ten minutes, then a tab-by-tab reference, the automation
chain written out in plain English, what to do when a number looks
wrong, the honest limits, and the keyboard shortcuts worth knowing.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Help"
LAST_COL = 7

SETUP = [
    ("1. Settings",
     "Household name, who cooks, the week start, your shopping window, "
     "budgets and targets. Ten minutes here and everything else has "
     "something to work with."),
    ("2. Household",
     "One row per person. The portion factor is the important one - a "
     "teenager eating 1.4 portions and a toddler eating 0.6 is why "
     "'serves 4' never quite works."),
    ("3. Lists",
     "Only if you want different categories, cuisines, stores or "
     "units. Everything else reads its dropdowns from here, so edit "
     "the list rather than typing a one-off value."),
    ("4. Ingredients",
     "Name, pack size, pack price, base unit. This is the tab that "
     "makes costing work; without prices the grocery list still works "
     "but the money columns stay blank."),
    ("5. Recipes",
     "Name, servings, timings, category. Then the lines go on Recipe "
     "Ingredients - one row per ingredient per recipe."),
    ("6. Meal Plan",
     "Pick a date, a slot and a recipe. The Weekly Planner and Monthly "
     "Planner are views of this tab, so you can plan from whichever "
     "one suits you."),
    ("7. Grocery List",
     "Nothing to do. Set the shopping window on Settings and the list "
     "builds itself from the plan, minus what the Pantry says you "
     "already have."),
]

TABS = [
    ("Start Here", "The ten minute setup and where everything lives."),
    ("Dashboard", "The week at a glance: cost, calories, what needs "
                  "attention, what is being eaten."),
    ("Weekly Planner", "Seven days by seven meal slots. Pick a week "
                       "with the date at the top."),
    ("Monthly Planner", "The same plan as a calendar month, with "
                        "dinners in the day cells."),
    ("Settings", "Every number the workbook uses to make a decision."),
    ("Lists", "The dropdown lists, and the unit conversion table."),
    ("Household", "Who eats, what they eat, what they cannot eat."),
    ("Recipes", "The recipe book. Cost, calories and macros come back "
                "from the ingredient lines."),
    ("Recipe Ingredients", "One row per ingredient per recipe. The "
                           "engine room."),
    ("Ingredients", "Prices, pack sizes, nutrition per 100, allergens, "
                    "storage."),
    ("Meal Plan", "Every planned meal as a row - the single source of "
                  "truth behind both planners."),
    ("Grocery List", "Built from the plan minus the pantry. Sorted by "
                     "aisle for the actual shop."),
    ("Pantry", "What is in the cupboards, with value and expiry."),
    ("Staples", "The things you always want in: tick when they run "
                "low."),
    ("Leftovers", "Portions in the fridge, with a use-by and a value."),
    ("Food Waste", "What got binned, why, and what it cost."),
    ("Meal Prep", "Batch cooking sessions and what they cover."),
    ("Dietary Filters", "Six dropdowns; the recipes that survive all "
                        "six."),
    ("Nutrition", "Daily calories and macros against the targets."),
    ("Budget", "Twelve months of planned against actual, by type and "
               "by store."),
    ("Recipe Costs", "Cheapest meals first, and where the money goes."),
    ("Converter", "Unit conversion and the recipe scaler."),
    ("Rotation", "Variety: what you cook too often and what you have "
                 "forgotten."),
    ("Food Safety", "Allergens by household member, storage times and "
                    "cooking temperatures."),
    ("Fridge & Freezer", "Cold storage sorted by what needs eating "
                         "first."),
    ("Grocery Spending", "Every receipt. This is what the budget "
                         "compares against."),
    ("Store Prices", "The same item at different shops, and the "
                     "cheapest one."),
    ("Stores", "Where you shop, travel time, delivery, price level."),
    ("Events", "Dinners, parties and holidays with their own budget."),
    ("Goals", "A handful of food habits with a measure attached."),
    ("Sustainability", "Why food gets binned and what stops it."),
    ("Reports", "A month on one page, and the year in one table."),
    ("Print Week", "The week's meals on one sheet for the fridge "
                   "door."),
    ("Print Shopping", "The shopping list with tick boxes."),
    ("Print Recipe", "One recipe scaled to any number of people."),
]

CHAIN = [
    "You set a shopping window on Settings (Shop from / Shop to).",
    "Meal Plan rows inside that window are counted. Each row knows its "
    "recipe and how many people are eating, so it knows its scale.",
    "Recipes totals up how many batches of each recipe the window "
    "needs.",
    "Recipe Ingredients multiplies every line by those batches and "
    "converts it to a base unit - grams, millilitres or each.",
    "Ingredients adds up the total needed for the window, takes off "
    "what the Pantry says you already have, and ranks what is short.",
    "Grocery List reads that ranking, rounds up to whole packs, prices "
    "it from the cheapest store, and groups it by aisle.",
    "Print Shopping prints that same list with tick boxes.",
]

TROUBLE = [
    ("The grocery list is empty",
     "Check Shop from and Shop to on Settings cover days that actually "
     "have meals planned, and that those meals point at a recipe "
     "rather than free text."),
    ("An ingredient is on the list but I have loads",
     "The Pantry quantity is what gets deducted. If the cupboard is "
     "full but the Pantry says zero, the list will buy it again."),
    ("Costs are blank or far too low",
     "An ingredient with no pack price costs nothing. The Recipes tab "
     "has a 'missing a price' flag that tells you which."),
    ("A conversion looks wrong",
     "Weight and volume are never converted into each other - a "
     "spreadsheet cannot know the density of your flour. Enter the "
     "line in the unit you will measure it in."),
    ("#N/A or #VALUE! in a cell",
     "Nearly always a name typed slightly differently from the list - "
     "'Tomatoes' and 'tomatoes ' are two ingredients. Use the "
     "dropdowns."),
    ("A sheet will not let me type",
     "Formula cells are locked on purpose. The password is on the "
     "Start Here tab; Review then Unprotect Sheet if you really want "
     "to change a calculation."),
    ("The numbers did not update",
     "Press F9. If that does not do it, Formulas then Calculation "
     "Options then Automatic."),
]

LIMITS = [
    "Automatic shopping lists need structured ingredients and "
    "consistent units. 'A splash of oil' cannot be added up; '15 ml "
    "oil' can.",
    "Nutrition is only as good as the per-100 figures you type in, and "
    "this is not medical or dietary advice. If something matters "
    "clinically, check it with a professional.",
    "Allergen and storage columns are memory aids, not guarantees. "
    "Always read the actual label - recipes change, so do products.",
    "Nothing here can send you a push notification, email you, or "
    "order the shopping. It is a spreadsheet; it works when you open "
    "it.",
    "Prices go stale. The store price tab is a snapshot of what you "
    "last saw, not a live feed.",
]

SHORTCUTS = [
    ("Ctrl + ;", "Today's date into the cell"),
    ("Alt + Down", "Open the dropdown in a validated cell"),
    ("F9", "Recalculate everything"),
    ("Ctrl + Page Down", "Next tab"),
    ("Ctrl + Shift + Down", "Select to the bottom of the column"),
    ("Ctrl + D", "Copy the cell above - the quickest way to extend a "
                 "formula column into new rows"),
    ("Ctrl + F", "Find - faster than scrolling the Recipes tab"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)

    s.banner("Help & How It Works",
             "Setup, the tab list, the automation chain, and what to "
             "do when a number looks wrong.", LAST_COL)
    s.set_col(0, 0, 24)
    s.set_col(1, LAST_COL, 16)

    row = 4
    row = s.note(row,
                 "Short version: fill in Settings, Household, "
                 "Ingredients and Recipes, then plan some meals. "
                 "Everything else is worked out for you. Every white "
                 "cell is yours to type in; every shaded cell is a "
                 "formula and is locked.",
                 LAST_COL, height=44)
    row += 1

    row = s.section(row, "Setting up, in order", LAST_COL)
    for label, detail in SETUP:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, detail, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    row = s.section(row, "What each tab is for", LAST_COL)
    row = s.table_head(row, 0, ["Tab", "What it does"], [24, 16])
    for tab, what in TABS:
        s.write(row, 0, tab, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, what, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    row = s.section(row, "How the shopping list builds itself",
                    LAST_COL)
    for i, step in enumerate(CHAIN, start=1):
        s.write(row, 0, f"Step {i}", fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, step, fmts.chrome("note"))
        s.row_height(row, 30)
        row += 1
    row = s.note(row,
                 "Every step is an ordinary formula you can click on "
                 "and read. There are no macros in this file, which is "
                 "also why it opens without a security warning.",
                 LAST_COL, height=30)
    row += 1

    row = s.section(row, "When something looks wrong", LAST_COL)
    row = s.table_head(row, 0, ["Symptom", "What it usually is"],
                       [24, 16])
    for symptom, fix in TROUBLE:
        s.write(row, 0, symptom, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, fix, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    row = s.section(row, "What this workbook cannot do", LAST_COL)
    for limit in LIMITS:
        s.merge(row, 0, LAST_COL, limit, fmts.chrome("note"))
        s.row_height(row, 30)
        row += 1
    row += 1

    row = s.section(row, "Shortcuts worth knowing", LAST_COL)
    row = s.table_head(row, 0, ["Keys", "What it does"], [24, 16])
    for keys, what in SHORTCUTS:
        s.write(row, 0, keys, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, what, fmts.chrome("note"))
        s.row_height(row, 22)
        row += 1
    row += 1

    row = s.section(row, "Adding more rows", LAST_COL)
    row = s.bullets(row, [
        "Every register has spare rows below the sample data - just "
        "keep typing.",
        "If you run out, select the last row, copy it, and insert "
        "copied cells. That carries the formulas and the dropdowns "
        "with it.",
        "Formula columns are locked; use the password on Start Here "
        "if you need to unlock a sheet first.",
        "Deleting rows is safe. Deleting whole columns is not - other "
        "tabs point at them by name.",
    ], LAST_COL, height=20)
    row += 1

    row = s.section(row, "Printing", LAST_COL)
    row = s.bullets(row, [
        "Print Week, Print Shopping and Print Recipe are already set "
        "up to fit a page - just press print.",
        "Everything else is set to fit the page width; use landscape "
        "for the wide registers.",
        "The printables read live data, so there is never a stale "
        "copy to tidy up.",
    ], LAST_COL, height=20)
    row += 1

    s.note(row,
           f"{cfg.product_name} - built by {cfg.author}. No macros, no "
           f"add-ins, no internet connection required. Works in Excel "
           f"2010 and later, and in Google Sheets and LibreOffice with "
           f"minor cosmetic differences.",
           LAST_COL, height=36)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

---

## 6. Quality checks (optional)

Not needed to build the workbook - these are the checks that keep it honest. The first two need `openpyxl`, the third needs `formulas`.

### `tools/validate_meal_workbook.py`

*414 lines · 6 functions*

> Static checks on a generated Meal Planner workbook.
>
> Run it with any Python that has openpyxl installed::
>
>     python tools/validate_meal_workbook.py output/planner.xlsx
>
> It never opens Excel - it parses the XML and reports anything that would
> break for a buyer: unresolved template tokens, formulas pointing at the
> wrong sheet, defined names that do not exist, missing sheet protection,
> formula cells that are not locked, and input cells that are.

| Defines | What it does |
|---------|--------------|
| `_colnum(ref)` |  |
| `check_raw_xml(path)` | Overlapping merges and stray hyperlinks - Excel calls these corrupt |
| `paren_depth(formula)` | Open brackets minus closing ones, ignoring anything inside text |
| `_specs()` | Every register specification, keyed by sheet name |
| `check_chain(wb, problems, notes)` | The grocery list only builds itself while these links hold |
| `check(path)` |  |

**Module constants:** `FUNCTION_RE` (re.compile(...)), `NAME_RE` (re.compile(...)), `BARE_COL_RE` (re.compile(...)), `KNOWN_FUNCTIONS` ({'IF', 'IFERROR', 'SUM', 'SUMIF', 'SUMIFS', 'COUNT', 'COU...), `MODERN_ONLY` ({'XLOOKUP', 'IFS', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCA...), `SEEDED_INPUTS` (3 entries), `CHAIN_CHECKS` (17 items).

````python
"""Static checks on a generated Meal Planner workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_meal_workbook.py output/planner.xlsx

It never opens Excel - it parses the XML and reports anything that would
break for a buyer: unresolved template tokens, formulas pointing at the
wrong sheet, defined names that do not exist, missing sheet protection,
formula cells that are not locked, and input cells that are.
"""

from __future__ import annotations

import os
import re
import sys
import zipfile
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:                                   # pragma: no cover
    print("openpyxl is required: pip install openpyxl")
    raise SystemExit(2)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from meal_planner_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402

FUNCTION_RE = re.compile(r"\b([A-Z][A-Z0-9\.]{1,})\s*\(")
NAME_RE = re.compile(r"(?<![A-Za-z0-9_.!$\"])([A-Za-z_][A-Za-z0-9_.]{2,})"
                     r"(?![A-Za-z0-9_.(!])")
#: a column letter with no row number - almost always a missing {r}
BARE_COL_RE = re.compile(r"(?<![A-Za-z0-9_.!$\"])([A-Z]{1,3})"
                         r"(?![A-Za-z0-9_.(!$:])")

KNOWN_FUNCTIONS = {
    "IF", "IFERROR", "SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTA", "COUNTIF",
    "COUNTIFS", "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "MAX", "MIN", "ROUND",
    "ROUNDUP", "ROUNDDOWN", "INDEX", "MATCH", "OFFSET", "TEXT", "VALUE",
    "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "PROPER",
    "CONCATENATE", "TODAY", "NOW", "DATE", "DAY", "MONTH", "YEAR", "WEEKDAY",
    "WEEKNUM", "EOMONTH", "EDATE", "DATEDIF", "DAYS", "NETWORKDAYS", "HOUR",
    "MINUTE", "TIME", "SUMPRODUCT", "LARGE", "SMALL", "RANK", "ABS", "AND",
    "OR", "NOT", "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISNA", "NA",
    "CHOOSE", "VLOOKUP", "HLOOKUP", "LOOKUP", "SUBSTITUTE", "REPT",
    "MEDIAN", "STDEV", "FLOOR", "CEILING", "INT", "MOD", "SIGN", "RAND",
    "SEARCH", "FIND", "EXACT", "N", "T", "TRUE", "FALSE", "ROW", "ROWS",
    "COLUMN", "COLUMNS", "YEARFRAC", "INDIRECT", "ADDRESS", "CELL",
    "SUBTOTAL", "TRANSPOSE", "PRODUCT", "POWER", "SQRT", "COUNTBLANK",
    "ISERR", "ISLOGICAL", "ISREF", "TYPE", "NUMBERVALUE", "DATEVALUE",
    "TIMEVALUE", "WORKDAY", "EOMONTH", "FREQUENCY", "PERCENTILE",
}

#: functions that do not exist in Excel 2010 / Google Sheets
MODERN_ONLY = {
    "XLOOKUP", "IFS", "MAXIFS", "MINIFS", "TEXTJOIN", "CONCAT", "SWITCH",
    "LET", "LAMBDA", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE",
    "TEXTSPLIT", "AGGREGATE", "IFNA",
}


def _colnum(ref: str) -> tuple:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    letters, row = match.group(1), int(match.group(2))
    num = 0
    for char in letters:
        num = num * 26 + (ord(char) - 64)
    return num, row


def check_raw_xml(path: str) -> list:
    """Overlapping merges and stray hyperlinks - Excel calls these corrupt."""
    found = []
    with zipfile.ZipFile(path) as zf:
        book = zf.read("xl/workbook.xml").decode()
        sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="rId(\d+)"',
                            book)
        sheet_names = {name for name, _ in sheets}
        for name, rid in sheets:
            try:
                xml = zf.read(f"xl/worksheets/sheet{rid}.xml").decode()
            except KeyError:
                continue
            merges = re.findall(r'<mergeCell ref="([^"]+)"/>', xml)
            boxes = []
            for ref in merges:
                start, _, end = ref.partition(":")
                end = end or start
                c1, r1 = _colnum(start)
                c2, r2 = _colnum(end)
                boxes.append((c1, r1, c2, r2, ref))
            for i, a in enumerate(boxes):
                for b in boxes[i + 1:]:
                    if (a[0] <= b[2] and b[0] <= a[2]
                            and a[1] <= b[3] and b[1] <= a[3]):
                        found.append(f"{name}: merges overlap "
                                     f"{a[4]} / {b[4]}")
            for location in re.findall(
                    r'<hyperlink ref="[^"]+" location="([^"]+)"', xml):
                target = location.split("!")[0].strip("'").replace("''", "'")
                if target not in sheet_names:
                    found.append(f"{name}: hyperlink points at missing "
                                 f"sheet '{target}'")
            for link in re.findall(r'<hyperlink ref="([^"]+)"', xml):
                lc, lr = _colnum(link)
                for c1, r1, c2, r2, ref in boxes:
                    if (c1 <= lc <= c2 and r1 <= lr <= r2
                            and (lc, lr) != (c1, r1)):
                        found.append(f"{name}: hyperlink {link} is inside "
                                     f"merged range {ref}")
    return found


def paren_depth(formula: str) -> int:
    """Open brackets minus closing ones, ignoring anything inside text.

    Excel escapes a quote inside a string by doubling it, so `""` is not
    the end of the string. Deeply nested IF() chains built by string
    concatenation are the one bug this whole file exists to catch: Excel
    and openpyxl both accept a short formula silently, and the buyer
    sees the damage.
    """
    depth = 0
    in_text = False
    i = 0
    while i < len(formula):
        char = formula[i]
        if in_text:
            if char == '"':
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        i += 1
    return depth


#: input cells that ship with a formula in them as a helpful default -
#: the buyer is meant to be able to type a fixed value straight over it
SEEDED_INPUTS = {
    "Dashboard!B5": "the week picker starts on the planner's week",
    "Settings!D22": "shopping window starts on the planning start date",
    "Settings!D23": "shopping window is a week long to begin with",
}

#: the automation chain, link by link: (sheet, column key, what the
#: formula in that column has to still be pointing at)
CHAIN_CHECKS = [
    ("Recipes", "plan_scale", ("sheet", "Meal Plan")),
    ("Recipe Ingredients", "plan_factor", ("sheet", "Recipes")),
    ("Recipe Ingredients", "need_base",
     ("column", "Recipe Ingredients", "plan_factor")),
    ("Ingredients", "need_now", ("sheet", "Recipe Ingredients")),
    ("Ingredients", "need_rank", ("column", "Ingredients", "need_now")),
    ("Grocery List", "ingredient",
     ("column", "Ingredients", "need_rank")),
    ("Grocery List", "needed", ("sheet", "Ingredients")),
    ("Ingredients", "in_pantry", ("sheet", "Pantry")),
    ("Grocery List", "in_pantry", ("sheet", "Ingredients")),
    ("Grocery List", "short", ("column", "Grocery List", "needed")),
    ("Grocery List", "packs", ("column", "Grocery List", "short")),
    ("Grocery List", "est_cost", ("column", "Grocery List", "packs")),
    ("Meal Plan", "scale", ("sheet", "Recipes")),
    ("Meal Plan", "key", ("column", "Meal Plan", "slot")),
    ("Pantry", "days_left", ("column", "Pantry", "est_expiry")),
    ("Leftovers", "queue_rank", ("column", "Leftovers", "days_left")),
    ("Recipes", "cost", ("sheet", "Recipe Ingredients")),
]


def _specs():
    """Every register specification, keyed by sheet name."""
    from meal_planner_spreadsheet.build import REGISTER_MODULES
    from meal_planner_spreadsheet.config import Config
    cfg = Config()
    return {module.NAME: module.spec(cfg) for module in REGISTER_MODULES}


def check_chain(wb, problems: list, notes: list) -> None:
    """The grocery list only builds itself while these links hold.

    Every one of them is a formula in row 5 of a register, so the check
    is cheap: find the column by its key in the specification, read the
    formula, and make sure it still mentions the thing it depends on.
    """
    specs = _specs()
    checked = 0
    for sheet_name, key, expect in CHAIN_CHECKS:
        spec = specs.get(sheet_name)
        if spec is None or sheet_name not in wb.sheetnames:
            problems.append(f"missing register {sheet_name}")
            continue
        try:
            column = spec.index(key) + 1
        except Exception:
            problems.append(f"{sheet_name}: column '{key}' has gone")
            continue
        formula = wb[sheet_name].cell(row=5, column=column).value
        if not isinstance(formula, str) or not formula.startswith("="):
            problems.append(f"{sheet_name}.{key} is no longer a formula")
            continue
        if expect[0] == "sheet":
            found = expect[1] in formula
            label = f"the {expect[1]} tab"
        else:
            letter = specs[expect[1]].letter(expect[2])
            label = f"{expect[1]}.{expect[2]} (column {letter})"
            if expect[1] == sheet_name:
                # a link inside the same register is a plain A1 reference
                found = bool(re.search(r"(?<![A-Z$])\$?" + letter +
                                       r"\$?5(?![0-9])", formula))
            else:
                found = f"${letter}$" in formula
        if not found:
            problems.append(
                f"{sheet_name}.{key} no longer points at {label} - "
                f"the automation chain is broken")
        checked += 1
    notes.append(f"automation chain: {checked} links verified")


def check(path: str) -> int:
    problems: list[str] = list(check_raw_xml(path))
    notes: list[str] = []
    if problems:
        print(f"--- {path}")
        for problem in problems[:40]:
            print(f"    - {problem}")
        return 1
    wb = openpyxl.load_workbook(path, data_only=False)

    # ---------------------------------------------------------- structure
    names = wb.sheetnames
    if names != EXPECTED_SHEETS:
        problems.append(f"sheet order/={names}")
    hidden = [ws.title for ws in wb.worksheets if ws.sheet_state != "visible"]
    if hidden:
        problems.append(f"no tab should be hidden, got {hidden}")

    check_chain(wb, problems, notes)

    # --------------------------------------------------------- properties
    props = wb.properties
    if props.creator != "Novality Store":
        problems.append(f"creator is {props.creator!r}")
    if props.lastModifiedBy != "Novality Store":
        problems.append(f"lastModifiedBy is {props.lastModifiedBy!r}")

    # --------------------------------------------------------- protection
    for ws in wb.worksheets:
        prot = ws.protection
        if not prot.sheet:
            problems.append(f"{ws.title}: sheet protection off")
        elif not prot.password:
            problems.append(f"{ws.title}: protected without a password")

    # -------------------------------------------------------- definednames
    defined = set(wb.defined_names)
    for ws in wb.worksheets:
        defined |= set(ws.defined_names)
    notes.append(f"{len(defined)} defined names")

    # ------------------------------------------------------------ formulas
    used_names: Counter = Counter()
    unknown_functions: Counter = Counter()
    bare_refs: dict = {}
    bad_functions: Counter = Counter()
    formula_cells = 0
    locked_formula_problem = 0
    unlocked_input_cells = 0
    per_sheet = defaultdict(int)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    formula_cells += 1
                    per_sheet[ws.title] += 1
                    if "[[" in value or "{" in value or "}" in value:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unresolved "
                            f"template token -> {value[:90]}")
                    if "#REF" in value:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: #REF -> "
                            f"{value[:90]}")
                    where = f"{ws.title}!{cell.coordinate}"
                    if (cell.protection and not cell.protection.locked
                            and where not in SEEDED_INPUTS):
                        locked_formula_problem += 1
                        if locked_formula_problem < 25:
                            problems.append(
                                f"{where}: formula cell is NOT locked")
                    stripped = re.sub(r'"[^"]*"', '""', value)
                    stripped = re.sub(r"'[^']*'!", "SHEET!", stripped)
                    if value.count('"') % 2:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: odd number of "
                            f"quotes -> {value[:80]}")
                    depth = paren_depth(value)
                    if depth:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unbalanced "
                            f"brackets ({depth:+d}) -> {value[:80]}")
                    if ",)" in stripped or "(," in stripped or \
                            ",," in stripped:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: empty argument "
                            f"-> {value[:80]}")
                    for func in FUNCTION_RE.findall(stripped):
                        if func in MODERN_ONLY:
                            bad_functions[func] += 1
                        elif func not in KNOWN_FUNCTIONS:
                            unknown_functions[func] += 1
                    for token in NAME_RE.findall(stripped):
                        if token in KNOWN_FUNCTIONS:
                            continue
                        used_names[token] += 1
                    for token in BARE_COL_RE.findall(stripped):
                        if token in KNOWN_FUNCTIONS:
                            continue
                        bare_refs[f"{ws.title}!{cell.coordinate} {token}"] = \
                            value[:90]
                elif value is not None and cell.protection and \
                        not cell.protection.locked:
                    unlocked_input_cells += 1

    # sheet names used inside formulas must exist
    sheet_refs = Counter()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    body = re.sub(r'"[^"]*"', '""', value)
                    for ref in re.findall(r"'([^']+)'!", body):
                        sheet_refs[ref] += 1
                    for ref in re.findall(r"(?<![A-Za-z0-9_'])"
                                          r"([A-Za-z][A-Za-z0-9_]*)!", body):
                        sheet_refs[ref] += 1
    for ref, count in sheet_refs.items():
        if ref not in names:
            problems.append(f"formula points at missing sheet "
                            f"'{ref}' ({count}x)")
    notes.append("cross sheet references: " + ", ".join(
        f"{k} {v}" for k, v in sorted(sheet_refs.items(),
                                      key=lambda kv: -kv[1])[:6]))

    for where, formula in list(bare_refs.items())[:15]:
        if where.rsplit(" ", 1)[-1] not in defined:
            problems.append(f"{where}: column letter with no row "
                            f"number -> {formula}")
    for func, count in unknown_functions.items():
        problems.append(f"unknown function {func}() used {count}x")
    for func, count in bad_functions.items():
        problems.append(f"{func}() needs a modern Excel - used {count}x")

    missing = {}
    for token, count in used_names.items():
        if token in defined:
            continue
        if token in KNOWN_FUNCTIONS or token.upper() in KNOWN_FUNCTIONS:
            continue
        if re.match(r"^[A-Z]{1,3}\d+$", token):
            continue
        missing[token] = count
    for token, count in sorted(missing.items(), key=lambda kv: -kv[1])[:25]:
        problems.append(f"undefined name used {count}x: {token}")

    # ----------------------------------------------------- data validation
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            formula = (dv.formula1 or "").lstrip("=")
            if dv.type == "list" and formula and \
                    re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", formula):
                if formula not in defined:
                    problems.append(
                        f"{ws.title}: drop down uses missing name "
                        f"{formula}")

    # ------------------------------------------------------------- report
    with zipfile.ZipFile(path) as zf:
        size = sum(i.file_size for i in zf.infolist())
    notes.append(f"{formula_cells:,} formula cells, "
                 f"{unlocked_input_cells:,} unlocked value cells")
    notes.append(f"uncompressed xml {size / 1_000_000:.1f} MB")

    print(f"--- {path}")
    for note in notes:
        print(f"  . {note}")
    print("  . formulas per sheet: " + ", ".join(
        f"{k} {v}" for k, v in sorted(per_sheet.items(),
                                      key=lambda kv: -kv[1])[:10]))
    if problems:
        print(f"  ! {len(problems)} problem(s):")
        for problem in problems[:60]:
            print(f"    - {problem}")
        return 1
    print("  OK - no problems found")
    return 0


if __name__ == "__main__":
    target = (sys.argv[1] if len(sys.argv) > 1
              else "output/Meal Planner & Grocery List.xlsx")
    raise SystemExit(check(target))
````

### `tools/check_meal_sample.py`

*244 lines · 5 functions*

> Cross-check the example records in a generated meal planner workbook.
>
> The demo data has to join up, otherwise the dashboard opens full of
> zeros and every lookup column fills with "Not found". This reads the
> literal values out of the finished file and makes sure every recipe
> line points at a real recipe and a real ingredient, every planned meal
> points at a real recipe, every pantry item and leftover is a real
> ingredient or recipe, and that there is enough of it to make the
> dashboard look like something.
>
>     python tools/check_meal_sample.py output/test.xlsx

| Defines | What it does |
|---------|--------------|
| `data_rows()` | Last data row of every register, straight from its specification |
| `headers(ws, header_row)` |  |
| `column_values(ws, header, header_row)` | Every non-empty literal value under the given header |
| `check(path)` |  |
| `main(argv)` |  |

**Module constants:** `HEADER_ROW` (4), `LAST_ROW` (0 entries).

````python
"""Cross-check the example records in a generated meal planner workbook.

The demo data has to join up, otherwise the dashboard opens full of
zeros and every lookup column fills with "Not found". This reads the
literal values out of the finished file and makes sure every recipe
line points at a real recipe and a real ingredient, every planned meal
points at a real recipe, every pantry item and leftover is a real
ingredient or recipe, and that there is enough of it to make the
dashboard look like something.

    python tools/check_meal_sample.py output/test.xlsx
"""

from __future__ import annotations

import os
import sys
from collections import Counter

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

HEADER_ROW = 4


def data_rows() -> dict:
    """Last data row of every register, straight from its specification.

    Without this the checks read the summary strip underneath the grid
    and report labels like 'Value in stock' as a missing ingredient.
    """
    from meal_planner_spreadsheet.build import REGISTER_MODULES
    from meal_planner_spreadsheet.config import Config
    cfg = Config()
    return {module.NAME: module.spec(cfg).last_data_row + 1
            for module in REGISTER_MODULES}


LAST_ROW: dict = {}


def headers(ws, header_row: int = HEADER_ROW) -> dict:
    out = {}
    for cell in ws[header_row]:
        if isinstance(cell.value, str) and cell.value.strip():
            out.setdefault(cell.value.strip(), cell.column)
    return out


def column_values(ws, header: str, header_row: int = HEADER_ROW) -> list:
    """Every non-empty literal value under the given header."""
    found = headers(ws, header_row)
    if header not in found:
        raise KeyError(f"{ws.title}: no column called {header!r} "
                       f"(have {sorted(found)[:10]} ...)")
    col = found[header]
    out = []
    last = min(LAST_ROW.get(ws.title, ws.max_row), ws.max_row)
    for row in range(header_row + 1, last + 1):
        value = ws.cell(row=row, column=col).value
        if value is None:
            continue
        if isinstance(value, str):
            value = value.strip()
            if not value or value.startswith("="):
                continue
        out.append(value)
    return out


def check(path: str) -> int:
    LAST_ROW.update(data_rows())
    wb = openpyxl.load_workbook(path, data_only=False)
    problems: list = []
    notes: list = []

    def sheet(name):
        if name not in wb.sheetnames:
            problems.append(f"missing sheet {name}")
            return None
        return wb[name]

    recipes = sheet("Recipes")
    lines = sheet("Recipe Ingredients")
    plan = sheet("Meal Plan")
    ingredients = sheet("Ingredients")
    if None in (recipes, lines, plan, ingredients):
        for problem in problems:
            print(f"    - {problem}")
        return 1

    recipe_names = column_values(recipes, "Recipe")
    ingredient_names = column_values(ingredients, "Ingredient")
    notes.append(f"{len(recipe_names)} example recipes, "
                 f"{len(ingredient_names)} ingredients")
    for label, values in (("recipe", recipe_names),
                          ("ingredient", ingredient_names)):
        duplicates = [v for v, n in Counter(values).items() if n > 1]
        if duplicates:
            problems.append(f"duplicate {label} names: {duplicates[:5]}")

    def id_link(from_sheet: str, from_col: str, to_sheet: str,
                count_col: str) -> None:
        """IDs are generated by formulas, so check the number is a row."""
        src, dst = sheet(from_sheet), sheet(to_sheet)
        if src is None or dst is None:
            return
        try:
            values = [str(v) for v in column_values(src, from_col)]
            limit = len(column_values(dst, count_col))
        except KeyError as error:
            problems.append(str(error))
            return
        bad = []
        for value in values:
            digits = "".join(c for c in value if c.isdigit())
            if not digits or not 1 <= int(digits) <= limit:
                bad.append(value)
        if bad:
            problems.append(
                f"{from_sheet}.{from_col}: {len(bad)} reference(s) point "
                f"past the {limit} rows on {to_sheet} -> "
                f"{sorted(set(bad))[:5]}")
        else:
            notes.append(f"{from_sheet}.{from_col} -> {to_sheet} "
                         f"({len(values)} rows, max {limit}) ok")

    def link(from_sheet: str, from_col: str, to_sheet: str,
             to_col: str) -> None:
        src, dst = sheet(from_sheet), sheet(to_sheet)
        if src is None or dst is None:
            return
        try:
            values = column_values(src, from_col)
            allowed = set(column_values(dst, to_col))
        except KeyError as error:
            problems.append(str(error))
            return
        missing = sorted({str(v) for v in values}
                         - {str(a) for a in allowed})
        if missing:
            problems.append(
                f"{from_sheet}.{from_col}: {len(missing)} value(s) not in "
                f"{to_sheet}.{to_col} -> {missing[:5]}")
        else:
            notes.append(f"{from_sheet}.{from_col} -> {to_sheet}.{to_col} "
                         f"({len(values)} rows) ok")

    # the spine: recipe lines, planned meals, prep jobs and leftovers
    id_link("Recipe Ingredients", "Recipe ID", "Recipes", "Recipe")
    id_link("Meal Plan", "Recipe ID", "Recipes", "Recipe")
    id_link("Meal Prep", "Recipe", "Recipes", "Recipe")
    id_link("Leftovers", "From which recipe", "Recipes", "Recipe")

    link("Recipe Ingredients", "Ingredient", "Ingredients", "Ingredient")
    link("Pantry", "Ingredient", "Ingredients", "Ingredient")
    link("Store Prices", "Ingredient", "Ingredients", "Ingredient")
    link("Store Prices", "Store", "Stores", "Store")
    link("Grocery Spending", "Store", "Stores", "Store")
    link("Food Waste", "What went in the bin", "Ingredients",
         "Ingredient")
    link("Staples", "Staple", "Ingredients", "Ingredient")
    link("Meal Plan", "Who is cooking", "Household", "Name")
    link("Meal Prep", "Who is doing it", "Household", "Name")

    # every recipe needs lines, or the costing and the list do nothing
    line_recipes = set(str(v) for v in
                       column_values(lines, "Recipe ID"))
    if len(line_recipes) < len(recipe_names):
        problems.append(
            f"only {len(line_recipes)} of {len(recipe_names)} recipes "
            f"have ingredient lines - the rest cost nothing")
    else:
        notes.append(f"{len(column_values(lines, 'Ingredient'))} recipe "
                     f"lines across {len(line_recipes)} recipes")

    # prices: an ingredient with no pack price silently costs nothing
    priced = [v for v in column_values(ingredients, "Pack price")
              if isinstance(v, (int, float)) and v > 0]
    if len(priced) < len(ingredient_names) * 0.9:
        problems.append(f"only {len(priced)} of {len(ingredient_names)} "
                        f"ingredients have a price")
    else:
        notes.append(f"{len(priced)} ingredients priced")

    # the planner itself
    dates = [v for v in column_values(plan, "Date") if hasattr(v, "year")]
    if len(dates) < 20:
        problems.append(f"only {len(dates)} planned meals - the week "
                        f"views will look empty")
    else:
        span = (max(dates) - min(dates)).days + 1
        notes.append(f"{len(dates)} planned meals over {span} days "
                     f"({min(dates):%d %b} to {max(dates):%d %b})")
    slots = Counter(column_values(plan, "Meal"))
    if slots.get("Dinner", 0) < 14:
        problems.append(f"only {slots.get('Dinner', 0)} dinners planned")
    else:
        notes.append("meals by slot: " + ", ".join(
            f"{k} {v}" for k, v in slots.most_common(4)))

    # money, waste and the rest of the dashboard tiles
    for name, header, least, label in (
            ("Grocery Spending", "Amount", 10, "receipts"),
            ("Food Waste", "How much", 5, "waste entries"),
            ("Pantry", "How much is left", 15, "pantry items"),
            ("Leftovers", "Portions left", 4, "leftover portions"),
            ("Store Prices", "Pack price", 20, "store prices"),
            ("Household", "Portion factor", 2, "household members")):
        ws = sheet(name)
        if ws is None:
            continue
        try:
            values = [v for v in column_values(ws, header)
                      if isinstance(v, (int, float))]
        except KeyError as error:
            problems.append(str(error))
            continue
        if len(values) < least:
            problems.append(f"only {len(values)} {label} - expected at "
                            f"least {least}")
        else:
            notes.append(f"{len(values)} {label}")

    print(f"--- {path}")
    for note in notes:
        print(f"  . {note}")
    if problems:
        print(f"  ! {len(problems)} problem(s):")
        for problem in problems[:30]:
            print(f"    - {problem}")
        return 1
    print("  OK - the example data joins up")
    return 0


def main(argv: list) -> int:
    targets = argv[1:] or ["output/Meal Planner & Grocery List.xlsx"]
    return max(check(path) for path in targets)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
````

### `tools/recalc_check.py`

*216 lines · 4 functions*

> Really calculate every formula in a generated workbook and report any
> cell that ends up as an Excel error (#REF!, #DIV/0!, #VALUE! ...).
>
> `validate_workbook.py` reads the file; this one *runs* it.  It loads the
> workbook into the `formulas` calculation engine, evaluates the whole
> dependency graph and then looks at the answer in every formula cell, so a
> formula that parses perfectly but divides by zero on an empty row is still
> caught.
>
>     pip install formulas openpyxl
>     python tools/recalc_check.py "output/Pet Sitting Business Spreadsheet.xlsx"
>
> Notes
> -----
> * The engine cannot resolve `OFFSET()` based defined names, so a working
>   copy of the file is made first with those names flattened to plain
>   ranges.  The copy is a test artefact - the real workbook keeps its
>   dynamic names and is never touched.
> * Only cells that actually hold a formula are judged, so help text that
>   merely mentions "#N/A" is not mistaken for a broken formula.
> * The engine is hungry: a full-size 400-row build is around 50,000
>   formulas and wants several GB of memory.  Check a smaller build
>   instead - the formulas are identical, there are just fewer rows of
>   them:
>
>       python build_pet_sitting_workbook.py --rows 120 -o /tmp/small.xlsx
>       python tools/recalc_check.py /tmp/small.xlsx
>
>   Pass --force to run it on a full-size file anyway.

| Defines | What it does |
|---------|--------------|
| `flatten_dynamic_names(src, dst)` | Save a copy of *src* at *dst* with OFFSET names turned into ranges |
| `formula_cells(path)` | `{(SHEET NAME upper, coordinate)}` for every cell holding a formula |
| `cell_text(value)` |  |
| `main(argv)` |  |

**Module constants:** `BIG` (25000), `ERROR_TEXT` (re.compile(...)), `DYNAMIC_NAME` (re.compile(...)).

````python
"""Really calculate every formula in a generated workbook and report any
cell that ends up as an Excel error (#REF!, #DIV/0!, #VALUE! ...).

`validate_workbook.py` reads the file; this one *runs* it.  It loads the
workbook into the `formulas` calculation engine, evaluates the whole
dependency graph and then looks at the answer in every formula cell, so a
formula that parses perfectly but divides by zero on an empty row is still
caught.

    pip install formulas openpyxl
    python tools/recalc_check.py "output/Pet Sitting Business Spreadsheet.xlsx"

Notes
-----
* The engine cannot resolve `OFFSET()` based defined names, so a working
  copy of the file is made first with those names flattened to plain
  ranges.  The copy is a test artefact - the real workbook keeps its
  dynamic names and is never touched.
* Only cells that actually hold a formula are judged, so help text that
  merely mentions "#N/A" is not mistaken for a broken formula.
* The engine is hungry: a full-size 400-row build is around 50,000
  formulas and wants several GB of memory.  Check a smaller build
  instead - the formulas are identical, there are just fewer rows of
  them:

      python build_pet_sitting_workbook.py --rows 120 -o /tmp/small.xlsx
      python tools/recalc_check.py /tmp/small.xlsx

  Pass --force to run it on a full-size file anyway.
"""

from __future__ import annotations

import argparse
import collections
import os
import re
import sys
import tempfile
import time

#: Formula count above which the engine needs more memory than a laptop
#: is likely to enjoy handing over.
BIG = 25000

ERROR_TEXT = re.compile(r"#(REF!|VALUE!|NAME\?|DIV/0!|NUM!|NULL!|N/A)")

# OFFSET(Sheet!$A$5,0,0,MAX(1,COUNTIF(Sheet!$A$5:$A$254,"?*")),1)
DYNAMIC_NAME = re.compile(
    r"^OFFSET\((?P<sheet>'[^']+'|[A-Za-z0-9_.]+)!"
    r"\$(?P<col>[A-Z]+)\$(?P<row>\d+),0,0,"
    r"MAX\(1,COUNTIF\([^,]+,\"\?\*\"\)\),1\)$")


def flatten_dynamic_names(src: str, dst: str) -> int:
    """Save a copy of *src* at *dst* with OFFSET names turned into ranges."""
    import warnings
    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        book = openpyxl.load_workbook(src)
    changed = 0
    for _name, defined in list(book.defined_names.items()):
        match = DYNAMIC_NAME.match(defined.value or "")
        if not match:
            continue
        last = re.search(r":\$[A-Z]+\$(\d+)", defined.value)
        if not last:
            continue
        col = match.group("col")
        defined.value = (f"{match.group('sheet')}!"
                         f"${col}${match.group('row')}:${col}${last.group(1)}")
        changed += 1
    book.save(dst)
    return changed


def formula_cells(path: str) -> set:
    """`{(SHEET NAME upper, coordinate)}` for every cell holding a formula."""
    import warnings
    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        book = openpyxl.load_workbook(path)
    cells = set()
    for sheet in book.worksheets:
        upper = sheet.title.upper()
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    cells.add((upper, cell.coordinate))
    return cells


def cell_text(value) -> str:
    try:
        return str(value.value[0, 0])
    except Exception:
        try:
            return str(value.value)
        except Exception:
            return ""


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook")
    parser.add_argument("--list", type=int, default=12, metavar="N",
                        help="how many example cells to print per error type")
    parser.add_argument("--keep", action="store_true",
                        help="keep the flattened working copy")
    parser.add_argument("--show", default="", metavar="CELLS",
                        help="also print the calculated value of these "
                             "cells, e.g. \"Dashboard!A10,Reports!B12\"")
    parser.add_argument("--force", action="store_true",
                        help="check a big workbook even though the engine "
                             "will need several GB of memory")
    args = parser.parse_args(argv)

    path = args.workbook
    if not os.path.isfile(path):
        print(f"no such file: {path}")
        return 2
    try:
        import formulas
    except ImportError:
        print("this check needs the calculation engine: pip install formulas")
        return 2

    print(path, flush=True)
    started = time.time()
    folder = tempfile.mkdtemp(prefix="recalc_")
    working = os.path.join(folder, os.path.basename(path))
    count = flatten_dynamic_names(path, working)
    print(f"  . flattened {count} dynamic names into a working copy",
          flush=True)

    wanted = formula_cells(path)
    print(f"  . {len(wanted):,} formula cells to check", flush=True)
    if len(wanted) > BIG and not args.force:
        print("  ! that is a lot - the engine would need several GB of "
              "memory.\n"
              "    Check a smaller build instead (the formulas are the "
              "same, there\n"
              "    are just fewer rows of them):\n"
              "      python build_pet_sitting_workbook.py --rows 120 "
              "-o /tmp/small.xlsx\n"
              "      python tools/recalc_check.py /tmp/small.xlsx\n"
              "    or pass --force to go ahead anyway.")
        return 2

    model = formulas.ExcelModel().loads(working).finish()
    print(f"  . loaded in {time.time() - started:.0f}s", flush=True)
    clock = time.time()
    solution = model.calculate()
    print(f"  . calculated {len(solution):,} nodes in "
          f"{time.time() - clock:.0f}s", flush=True)

    book = os.path.basename(working).upper()
    counts = collections.Counter()
    examples = collections.defaultdict(list)
    seen = 0
    for key, value in solution.items():
        if "!" not in key or "]" not in key:
            continue
        sheet, _, coord = key.partition("!")
        sheet = sheet.split("]")[-1].rstrip("'").upper()
        if book and sheet.endswith(".XLSX"):
            continue
        coord = coord.replace("$", "")
        if (sheet, coord) not in wanted:
            continue
        seen += 1
        found = ERROR_TEXT.search(cell_text(value))
        if not found:
            continue
        token = f"{sheet} {found.group(0)}"
        counts[token] += 1
        if len(examples[token]) < args.list:
            examples[token].append(coord)

    if args.show:
        wanted_show = {}
        for item in args.show.split(","):
            item = item.strip()
            if not item:
                continue
            sheet, _, coord = item.rpartition("!")
            wanted_show[(sheet.strip("'").upper(),
                         coord.replace("$", "").upper())] = item
        print("  . values asked for:")
        for key, value in solution.items():
            if "!" not in key or "]" not in key:
                continue
            sheet, _, coord = key.partition("!")
            sheet = sheet.split("]")[-1].rstrip("'").upper()
            coord = coord.replace("$", "")
            label = wanted_show.get((sheet, coord))
            if label:
                print(f"    {label:28s} = {cell_text(value)[:70]}")

    print(f"  . {seen:,} of them evaluated by the engine", flush=True)
    if not counts:
        print("  OK - every formula worked out to a real answer")
        return 0
    print(f"  ! {sum(counts.values())} cell(s) ended up as an error:")
    for token, number in counts.most_common():
        print(f"    - {token} x{number}: {', '.join(examples[token])}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
````

### `tools/export_source_md.py`

*922 lines · 10 functions*

> Write every source file of a generator into one markdown document.
>
>     python tools/export_source_md.py                   # -> SOURCE_CODE.md
>     python tools/export_source_md.py --product social  # -> SOURCE_CODE_SOCIAL.md
>     python tools/export_source_md.py -o /tmp/src.md
>
> Handy when the code has to travel somewhere that is not a git checkout:
> one file, every module in reading order, each in a fenced block with its
> real path above it, plus a short description taken from the module's own
> docstring so the two can never disagree.

| Defines | What it does |
|---------|--------------|
| `outline(path)` | `[(kind, name, note)]` for the classes, functions and constants |
| `describe_value(node)` | A short, human description of a module level constant |
| `first_sentence(text)` |  |
| `blurb(path)` | First sentence of the module docstring, as a one line description |
| `full_docstring(path)` |  |
| `anchor(heading)` |  |
| `tab_table(product)` | A row per tab: order, name, kind, module, columns, rows |
| `data_dictionary(product)` | One table per register: every column, its type and where it comes from |
| `file_table(sections)` |  |
| `main(argv)` |  |

**Module constants:** `ROOT` (os.path.dirname(...)), `SHARED_TOOLS` (2 items), `PET_SECTIONS` (5 items), `PET` (9 entries), `SOCIAL_SECTIONS` (7 items), `SOCIAL` (9 entries), `MEAL_SECTIONS` (6 items), `MEAL` (11 entries), `PRODUCTS` (3 entries), `ARCHITECTURE` ('Everything here is plain Python writing a `.xlsx` file t...), `DEFAULT_CLOSING` ('**Hide the `Calc` engine**, activate the first tab and c...), `CALC_ENGINE` ('### The hidden `Calc` tab\n\nDashboards do not aggregate...), `RANK_ENGINE` ('### Ranking columns instead of a hidden engine\n\nThere ...).

````python
"""Write every source file of a generator into one markdown document.

    python tools/export_source_md.py                   # -> SOURCE_CODE.md
    python tools/export_source_md.py --product social  # -> SOURCE_CODE_SOCIAL.md
    python tools/export_source_md.py -o /tmp/src.md

Handy when the code has to travel somewhere that is not a git checkout:
one file, every module in reading order, each in a fenced block with its
real path above it, plus a short description taken from the module's own
docstring so the two can never disagree.
"""

from __future__ import annotations

import argparse
import ast
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SHARED_TOOLS = [
    "tools/export_source_md.py",
    "tools/make_download_zip.py",
]

# ---------------------------------------------------------------------------
# Product A - pet sitting
# ---------------------------------------------------------------------------

PET_SECTIONS = [
    ("1. Run it",
     "Save the files below keeping these exact paths, then run the one "
     "liner. Nothing else is needed - no project file, no build step.\n\n"
     "```bash\n"
     "pip install XlsxWriter\n"
     "python build_pet_sitting_workbook.py\n"
     "```\n\n"
     "The workbook lands in `output/Pet Sitting Business Spreadsheet.xlsx`.",
     ["build_pet_sitting_workbook.py"]),
    ("2. The package",
     "Configuration, the formula template engine, cell formats, the sheet "
     "writer and the demo data. These are the parts every tab leans on.",
     ["pet_sitting_spreadsheet/__init__.py",
      "pet_sitting_spreadsheet/__main__.py",
      "pet_sitting_spreadsheet/cli.py",
      "pet_sitting_spreadsheet/config.py",
      "pet_sitting_spreadsheet/layout.py",
      "pet_sitting_spreadsheet/formats.py",
      "pet_sitting_spreadsheet/writer.py",
      "pet_sitting_spreadsheet/names.py",
      "pet_sitting_spreadsheet/build.py",
      "pet_sitting_spreadsheet/sample_data.py"]),
    ("3. The registers",
     "One module per data tab. Each returns a `SheetSpec` - a list of "
     "columns with their formulas - and the writer turns that into a "
     "formatted, validated, protected sheet.",
     ["pet_sitting_spreadsheet/sheets/__init__.py",
      "pet_sitting_spreadsheet/sheets/common.py",
      "pet_sitting_spreadsheet/sheets/lists.py",
      "pet_sitting_spreadsheet/sheets/settings.py",
      "pet_sitting_spreadsheet/sheets/services.py",
      "pet_sitting_spreadsheet/sheets/clients.py",
      "pet_sitting_spreadsheet/sheets/pets.py",
      "pet_sitting_spreadsheet/sheets/bookings.py",
      "pet_sitting_spreadsheet/sheets/visits.py",
      "pet_sitting_spreadsheet/sheets/health.py",
      "pet_sitting_spreadsheet/sheets/incidents.py",
      "pet_sitting_spreadsheet/sheets/messages.py",
      "pet_sitting_spreadsheet/sheets/invoices.py",
      "pet_sitting_spreadsheet/sheets/payments.py",
      "pet_sitting_spreadsheet/sheets/expenses.py",
      "pet_sitting_spreadsheet/sheets/mileage.py",
      "pet_sitting_spreadsheet/sheets/staff.py",
      "pet_sitting_spreadsheet/sheets/inventory.py",
      "pet_sitting_spreadsheet/sheets/reminders.py"]),
    ("4. The pages",
     "The tabs that read the registers: the hidden calculation engine, the "
     "dashboard, the reports, the planner, the welcome tab, the printable "
     "pages and the help tab.",
     ["pet_sitting_spreadsheet/sheets/calc.py",
      "pet_sitting_spreadsheet/sheets/dashboard.py",
      "pet_sitting_spreadsheet/sheets/reports.py",
      "pet_sitting_spreadsheet/sheets/budget.py",
      "pet_sitting_spreadsheet/sheets/start_here.py",
      "pet_sitting_spreadsheet/sheets/printables.py",
      "pet_sitting_spreadsheet/sheets/help_tab.py"]),
    ("5. Quality checks (optional)",
     "Not needed to build the workbook - these are the checks that keep it "
     "honest. The first two need `openpyxl`, the third needs `formulas`.",
     ["tools/validate_workbook.py",
      "tools/check_sample_data.py",
      "tools/recalc_check.py"] + SHARED_TOOLS),
]

PET = {
    "package": "pet_sitting_spreadsheet",
    "builder": "build_pet_sitting_workbook.py",
    "doc": "README.md",
    "tabs": 27,
    "title": "Pet Sitting & Dog Walking Business Spreadsheet",
    "subtitle": ("Complete source code of the generator that writes the "
                 "`.xlsx` product."),
    "output": "SOURCE_CODE.md",
    "intro": [
        "Writes a 27 tab `.xlsx` workbook with about 50,000 formulas: client "
        "and pet records, bookings, visit logs, pricing, invoices, payments, "
        "expenses, mileage, a dashboard, reports and automatic reminders. "
        "Document author is **Novality Store**; every formula cell is locked "
        "and every sheet protected with the password **`premium`**, while "
        "the cells you are meant to type in stay open.",
        "",
        "Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the "
        "much older 0.3.0, so it does not rely on anything recent. The "
        "workbook itself is Excel 2010 safe — no dynamic arrays, no "
        "`XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens "
        "in Google Sheets, LibreOffice and Numbers.",
        "",
        "```bash",
        "python build_pet_sitting_workbook.py --help",
        "",
        "  -o, --output PATH     where to write the file",
        "  --blank               no example data, just the empty workbook",
        "  --no-protect          leave the sheets unlocked",
        "  --password TEXT       protection password (default: premium)",
        "  --author TEXT         document author  (default: Novality Store)",
        "  --business TEXT       business name on the cover and invoices",
        "  --currency SYMBOL     currency symbol  (default: $)",
        "  --unit {km,miles}     mileage unit",
        "  --rows N              rows of capacity in the big registers",
        "  --today YYYY-MM-DD    pretend it is this date (for the demo data)",
        "```",
    ],
    "sections": PET_SECTIONS,
}

# ---------------------------------------------------------------------------
# Product B - social media content planner
# ---------------------------------------------------------------------------

_PKG = "social_media_spreadsheet"

SOCIAL_SECTIONS = [
    ("1. Run it",
     "Save the files below keeping these exact paths, then run the one "
     "liner. Nothing else is needed - no project file, no build step.\n\n"
     "```bash\n"
     "pip install XlsxWriter\n"
     "python build_social_media_workbook.py\n"
     "```\n\n"
     "The workbook lands in "
     "`output/Social Media Content Planner.xlsx`.",
     ["build_social_media_workbook.py"]),
    ("2. The package",
     "Configuration, the formula template engine, cell formats, the sheet "
     "writer, the cross sheet names and the demo data. These are the parts "
     "every tab leans on.",
     [f"{_PKG}/__init__.py",
      f"{_PKG}/__main__.py",
      f"{_PKG}/cli.py",
      f"{_PKG}/config.py",
      f"{_PKG}/layout.py",
      f"{_PKG}/formats.py",
      f"{_PKG}/writer.py",
      f"{_PKG}/names.py",
      f"{_PKG}/build.py",
      f"{_PKG}/sample_data.py"]),
    ("3. Setup & planning registers",
     "One module per data tab. Each returns a `SheetSpec` - a list of "
     "columns with their formulas - and the writer turns that into a "
     "formatted, validated, protected sheet. These are the tabs you fill "
     "in before you start posting.",
     [f"{_PKG}/sheets/__init__.py",
      f"{_PKG}/sheets/common.py",
      f"{_PKG}/sheets/lists.py",
      f"{_PKG}/sheets/settings.py",
      f"{_PKG}/sheets/accounts.py",
      f"{_PKG}/sheets/personas.py",
      f"{_PKG}/sheets/pillars.py",
      f"{_PKG}/sheets/calendar_sheet.py",
      f"{_PKG}/sheets/ideas.py",
      f"{_PKG}/sheets/captions.py",
      f"{_PKG}/sheets/hashtags.py",
      f"{_PKG}/sheets/platforms.py"]),
    ("4. Campaign & production registers",
     "The tabs that run the work: campaigns and launches, the seasonal "
     "plan, collaborations, repurposing, research and the team.",
     [f"{_PKG}/sheets/campaigns.py",
      f"{_PKG}/sheets/launches.py",
      f"{_PKG}/sheets/seasonal.py",
      f"{_PKG}/sheets/influencers.py",
      f"{_PKG}/sheets/ugc.py",
      f"{_PKG}/sheets/repurposing.py",
      f"{_PKG}/sheets/competitors.py",
      f"{_PKG}/sheets/trends.py",
      f"{_PKG}/sheets/assets.py",
      f"{_PKG}/sheets/team.py",
      f"{_PKG}/sheets/tasks.py",
      f"{_PKG}/sheets/compliance.py"]),
    ("5. Measurement registers",
     "Where the numbers go in: monthly account analytics, per post "
     "performance, tests, paid promotion, goals and the budget.",
     [f"{_PKG}/sheets/analytics.py",
      f"{_PKG}/sheets/performance.py",
      f"{_PKG}/sheets/ab_tests.py",
      f"{_PKG}/sheets/ads.py",
      f"{_PKG}/sheets/goals.py",
      f"{_PKG}/sheets/budget.py"]),
    ("6. The pages",
     "The tabs that read the registers: the hidden calculation engine, the "
     "two dashboards, the calendar grid, the production pipeline, the "
     "brand voice page, the reports, the welcome tab, the printables and "
     "the help tab.",
     [f"{_PKG}/sheets/calc.py",
      f"{_PKG}/sheets/dashboard.py",
      f"{_PKG}/sheets/annual.py",
      f"{_PKG}/sheets/month_view.py",
      f"{_PKG}/sheets/workflow.py",
      f"{_PKG}/sheets/brand_voice.py",
      f"{_PKG}/sheets/reports.py",
      f"{_PKG}/sheets/start_here.py",
      f"{_PKG}/sheets/printables.py",
      f"{_PKG}/sheets/help_tab.py"]),
    ("7. Quality checks (optional)",
     "Not needed to build the workbook - these are the checks that keep it "
     "honest. The first two need `openpyxl`, the third needs `formulas`.",
     ["tools/validate_social_workbook.py",
      "tools/check_social_sample.py",
      "tools/recalc_check.py"] + SHARED_TOOLS),
]

SOCIAL = {
    "package": "social_media_spreadsheet",
    "builder": "build_social_media_workbook.py",
    "doc": "README_SOCIAL.md",
    "tabs": 43,
    "title": "Social Media Content Calendar & Analytics Planner",
    "subtitle": ("Complete source code of the generator that writes the "
                 "`.xlsx` product."),
    "output": "SOURCE_CODE_SOCIAL.md",
    "intro": [
        "Writes a 43 tab `.xlsx` workbook with over 110,000 formulas: a "
        "12 month content calendar, an idea bank, caption and hashtag "
        "banks, five platform planners, campaigns, launches, influencer "
        "and UGC tracking, repurposing, competitor and trend research, "
        "monthly analytics, per post performance, two dashboards, printable "
        "briefs and a reports page. Document author is **Novality Store**; "
        "every formula cell is locked and every sheet protected with the "
        "password **`premium`**, while the cells you are meant to type in "
        "stay open.",
        "",
        "Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the "
        "much older 0.3.0, so it does not rely on anything recent. The "
        "workbook itself is Excel 2010 safe — no dynamic arrays, no "
        "`XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens "
        "in Google Sheets, LibreOffice and Numbers.",
        "",
        "```bash",
        "python build_social_media_workbook.py --help",
        "",
        "  -o, --output PATH     where to write the file",
        "  --blank               no example data, just the empty workbook",
        "  --no-protect          leave the sheets unlocked",
        "  --password TEXT       protection password (default: premium)",
        "  --author TEXT         document author  (default: Novality Store)",
        "  --brand TEXT          brand name on the cover and the dashboards",
        "  --currency SYMBOL     currency symbol  (default: $)",
        "  --year YYYY           the year the calendar is built for",
        "  --rows N              rows of capacity in the big registers",
        "  --today YYYY-MM-DD    pretend it is this date (for the demo data)",
        "```",
    ],
    "sections": SOCIAL_SECTIONS,
}

# ---------------------------------------------------------------------------
# Product C - automated meal planner
# ---------------------------------------------------------------------------

_MEAL = "meal_planner_spreadsheet"

MEAL_SECTIONS = [
    ("1. Run it",
     "Save the files below keeping these exact paths, then run the one "
     "liner. Nothing else is needed - no project file, no build step.\n\n"
     "```bash\n"
     "pip install XlsxWriter\n"
     "python build_meal_planner_workbook.py\n"
     "```\n\n"
     "The workbook lands in "
     "`output/Meal Planner & Grocery List.xlsx`.",
     ["build_meal_planner_workbook.py"]),
    ("2. The package",
     "Configuration, the formula template engine, cell formats, the sheet "
     "writer, the cross sheet names and the demo household. These are the "
     "parts every tab leans on.",
     [f"{_MEAL}/__init__.py",
      f"{_MEAL}/__main__.py",
      f"{_MEAL}/cli.py",
      f"{_MEAL}/config.py",
      f"{_MEAL}/layout.py",
      f"{_MEAL}/formats.py",
      f"{_MEAL}/writer.py",
      f"{_MEAL}/names.py",
      f"{_MEAL}/build.py",
      f"{_MEAL}/sample_data.py"]),
    ("3. The food registers",
     "One module per data tab. Each returns a `SheetSpec` - a list of "
     "columns with their formulas - and the writer turns that into a "
     "formatted, validated, protected sheet. These five are the chain "
     "that makes the shopping list write itself.",
     [f"{_MEAL}/sheets/__init__.py",
      f"{_MEAL}/sheets/common.py",
      f"{_MEAL}/sheets/lists.py",
      f"{_MEAL}/sheets/settings.py",
      f"{_MEAL}/sheets/members.py",
      f"{_MEAL}/sheets/recipes.py",
      f"{_MEAL}/sheets/recipe_lines.py",
      f"{_MEAL}/sheets/ingredients.py",
      f"{_MEAL}/sheets/meal_plan.py",
      f"{_MEAL}/sheets/grocery.py"]),
    ("4. The kitchen registers",
     "What is in the house and what happens to it: the pantry, the "
     "staples, leftovers, waste, batch cooking, the shops and what you "
     "spend in them, events and goals.",
     [f"{_MEAL}/sheets/pantry.py",
      f"{_MEAL}/sheets/staples.py",
      f"{_MEAL}/sheets/leftovers.py",
      f"{_MEAL}/sheets/waste.py",
      f"{_MEAL}/sheets/prep.py",
      f"{_MEAL}/sheets/spending.py",
      f"{_MEAL}/sheets/store_prices.py",
      f"{_MEAL}/sheets/stores.py",
      f"{_MEAL}/sheets/events.py",
      f"{_MEAL}/sheets/goals.py"]),
    ("5. The pages",
     "The tabs that read the registers: the welcome page, the dashboard, "
     "the two planners, the filters, the money and nutrition pages, the "
     "converter and scaler, rotation, food safety, cold storage, "
     "sustainability, the reports, the three printables and the help "
     "tab.",
     [f"{_MEAL}/sheets/start_here.py",
      f"{_MEAL}/sheets/dashboard.py",
      f"{_MEAL}/sheets/weekly.py",
      f"{_MEAL}/sheets/monthly.py",
      f"{_MEAL}/sheets/filters.py",
      f"{_MEAL}/sheets/nutrition.py",
      f"{_MEAL}/sheets/budget.py",
      f"{_MEAL}/sheets/recipe_costs.py",
      f"{_MEAL}/sheets/converter.py",
      f"{_MEAL}/sheets/rotation.py",
      f"{_MEAL}/sheets/allergens.py",
      f"{_MEAL}/sheets/fridge.py",
      f"{_MEAL}/sheets/sustainability.py",
      f"{_MEAL}/sheets/reports.py",
      f"{_MEAL}/sheets/printables.py",
      f"{_MEAL}/sheets/help_tab.py"]),
    ("6. Quality checks (optional)",
     "Not needed to build the workbook - these are the checks that keep "
     "it honest. The first two need `openpyxl`, the third needs "
     "`formulas`.",
     ["tools/validate_meal_workbook.py",
      "tools/check_meal_sample.py",
      "tools/recalc_check.py"] + SHARED_TOOLS),
]

MEAL = {
    "package": "meal_planner_spreadsheet",
    "builder": "build_meal_planner_workbook.py",
    "doc": "README_MEAL.md",
    "tabs": 36,
    "title": "Automated Meal Planner & Grocery List Spreadsheet",
    "subtitle": ("Complete source code of the generator that writes the "
                 "`.xlsx` product."),
    "output": "SOURCE_CODE_MEAL.md",
    "closing": ("**Activate the first tab** and close the file - there "
                "is no hidden sheet to hide in this one."),
    "engine_heading": "Ranking columns instead of a hidden engine",
    "intro": [
        "Writes a 36 tab `.xlsx` workbook with over 72,000 formulas: a "
        "weekly and a monthly meal planner, a recipe book with costed "
        "ingredient lines, a grocery list that builds itself from the "
        "plan and deducts the pantry, fridge and freezer tracking, "
        "leftovers and waste, a food budget, recipe costing and store "
        "price comparison, unit conversion and recipe scaling, dietary "
        "filters, calories and macros, batch cooking, meal rotation, "
        "food safety, a dashboard, reports and three printable pages. "
        "Document author is **Novality Store**; every formula cell is "
        "locked and every sheet protected with the password "
        "**`premium`**, while the cells you are meant to type in stay "
        "open.",
        "",
        "Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on "
        "the much older 0.3.0, so it does not rely on anything recent. "
        "The workbook itself is Excel 2010 safe \u2014 no dynamic arrays, "
        "no `XLOOKUP`, no `IFS`, nothing that needs array entry \u2014 and "
        "it opens in Google Sheets, LibreOffice and Numbers.",
        "",
        "```bash",
        "python build_meal_planner_workbook.py --help",
        "",
        "  -o, --output PATH     where to write the file",
        "  --blank               no example data, just the empty workbook",
        "  --no-protect          leave the sheets unlocked",
        "  --password TEXT       protection password (default: premium)",
        "  --author TEXT         document author  (default: Novality Store)",
        "  --household TEXT      household name on the cover and printables",
        "  --cook TEXT           the main cook / planner",
        "  --currency SYMBOL     currency symbol  (default: $)",
        "  --year YYYY           the year the planner is built for",
        "  --week-start DAY      Monday or Sunday",
        "  --servings N          default number of people a meal feeds",
        "  --budget AMOUNT       weekly food budget",
        "  --units {Metric,Imperial}   which unit set the drop downs lead with",
        "  --rows N              rows of capacity in the big registers",
        "  --today YYYY-MM-DD    pretend it is this date (for the demo data)",
        "```",
    ],
    "sections": MEAL_SECTIONS,
}

PRODUCTS = {"pet": PET, "social": SOCIAL, "meal": MEAL}


# ---------------------------------------------------------------------------
# Chapter 0 - how the generator is put together (same engine for both
# products, so the prose is written once and filled in per product)
# ---------------------------------------------------------------------------

ARCHITECTURE = """\
Everything here is plain Python writing a `.xlsx` file through XlsxWriter.
There is no template workbook, no copy of an existing file and no manual
step: the spreadsheet is described in code, and the code is the only
source of truth for what ends up in a cell.

Two kinds of tab exist.

* A **register** is a data tab — a header row and hundreds of rows under
  it. Its module exposes `NAME`, a `spec(cfg)` that returns a
  `SheetSpec` (a list of `Col` objects) and, optionally, a
  `decorate(sheet, fmts, cfg)` that adds totals or notes underneath the
  grid. Nothing in a register module writes cells directly.
* A **page** is a hand-built tab — the dashboards, the reports, the
  printables, the welcome and help pages. Its module exposes
  `build(book, fmts, cfg, reg)` and writes with `SheetWriter` helpers:
  `banner`, `section`, `kpi`, `table_head`, `calc_row`, `note`,
  `bullets`, `field`, `merge`, `finish`.

### The build pipeline

`{builder}` parses the command line into a `Config`, then
`{package}.build.build_workbook(path, cfg)` does six things in order:

1. **Open the workbook** and stamp the document properties — title,
   author, company, keywords and a comment naming the protection
   password. The author is the same on every generated file.
2. **Build the `FormatBook`.** Every visual style the product uses is
   created once, up front: a style (text, date, money, percent, status,
   link, …) crossed with a mode (`INPUT` cream and unlocked, `CALC` grey
   and locked, `PLAIN`), plus the page chrome (titles, section bars, KPI
   tiles, notes) and the conditional formatting palette.
3. **Register every `SheetSpec` first**, before a single tab is written.
   This is what lets any sheet refer to any other sheet's columns by
   name regardless of which tab happens to be written first.
4. **Write the tabs in `SHEET_ORDER`.** Registers go through
   `build_register()` — `SheetWriter.grid(spec)` writes widths, formats,
   headers, formulas, sample rows, drop downs, header comments,
   conditional formats, the autofilter, frozen panes and the print
   setup — then the module's `decorate()` runs and the sheet is
   protected. Pages call their own `build()`.
5. **Define the live lookup names** (`ContentIDs`, `TeamNames`, …) from
   the registry. Each is an `OFFSET` range that grows and shrinks with
   the number of filled rows, so a drop down never shows blanks.
6. {closing}

### Describing a column

A column is declared once, for one row, with placeholders:

```python
Col("er", "Engagement rate", 11, "percent",
    formula='=IFERROR({{engagements}}{{r}}/{{reach}}{{r}},"")',
    cf=[{{"type": "cell", "criteria": ">=", "value": "=TargetEngagement",
         "fmt": "good"}}])
```

`Col` carries the key, the header text, the width, the style, the
formula template, the data validation, the header comment, the
conditional formats, whether the column is hidden and which outline
group it belongs to. `Registry.render()` expands the template per row:

| Token | Becomes |
|-------|---------|
| `{{r}}` / `{{r_prev}}` | this row number / the row above |
| `{{key}}` | the column letter of another column on the same sheet |
| `{{first}}` / `{{last}}` | the first and last data row of the sheet |
| `[[Sheet.key]]` | an absolute range on another sheet, e.g. `'Content Calendar'!$G$5:$G$604` |
| `[[Sheet.key#12]]` | one cell of that column |
| `[[Sheet.key#]]` | the same row as the formula being written |

Helpers cover the patterns Excel 2010 makes awkward — `max_if`,
`min_if`, `blank_if`, `safe`, `lookup`, `dynamic_name`, `quote_sheet`.

### The rules every formula obeys

* **Excel 2010 safe.** `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `SMALL`,
  `LARGE`, `RANK`, `EDATE`, `EOMONTH`, `SUMPRODUCT(MAX(...))`. No
  dynamic arrays, no `XLOOKUP`, no `IFS`, no `AGGREGATE`, nothing that
  needs Ctrl+Shift+Enter. The file opens in Excel 2010 and later, Google
  Sheets, LibreOffice and Numbers.
* **No charts.** The vendored XlsxWriter in this repository cannot write
  them, so trends are drawn with `add_sparkline`, data bars and
  three-colour scales — which also survive a Google Sheets import.
* **Nth-row lookups without array formulas.** Hidden composite key
  columns (`date + ROW()/100000`) plus `SMALL` and `INDEX/MATCH` answer
  "the third thing on 14 March" and "the next eight due" with ordinary
  formulas.
* **One number, one home.** Nothing is typed into two tabs. Every other
  appearance of a figure is a formula pointing back at the one place it
  was entered.
* **Locked where it matters.** Formula cells are locked, input cells are
  not, and every sheet is protected with the same password — so the
  maths cannot be typed over by accident but the cream cells stay open.

{engine}
"""

DEFAULT_CLOSING = ("**Hide the `Calc` engine**, activate the first tab "
                   "and close the file.")

CALC_ENGINE = """\
### The hidden `Calc` tab

Dashboards do not aggregate the registers directly. A hidden sheet
pre-computes one row per month, per platform/service, per category and
per owner, each with its own ranking key, and the visible pages read
those columns with plain `INDEX`/`MATCH`. It keeps the heavy `SUMIFS`
work in one place, keeps recalculation predictable, and makes every
dashboard number traceable to a single cell."""

RANK_ENGINE = """\
### Ranking columns instead of a hidden engine

There is no hidden `Calc` tab in this product. Anything a page needs to
rank - the next thing to buy, the recipe you have left alone longest,
the leftovers to eat first, the fridge item closest to its date - is a
**hidden rank column on the register itself**:

```python
Col("rotation_rank", "Rotation rank", 9, "int",
    formula='=IF({{name}}{{r}}="","",COUNTIFS({{sort_key_first}}:'
            '{{sort_key_last}},">"&{{sort_key}}{{r}})+'
            'COUNTIFS({{sort_key_first}}:{{sort_key}}{{r}},'
            '{{sort_key}}{{r}}))',
    hidden=True)
```

`COUNTIFS(whole column, ">" & this row) + COUNTIFS(up to this row, this
row)` is a dense rank that breaks ties by row order, so the ranks come
out 1, 2, 3 ... with no gaps and no duplicates. The page then asks for
the nth item with an ordinary `MATCH(n, rank_range, 0)` - a plain range
lookup, which is why there is **no array formula and nothing needing
Ctrl+Shift+Enter anywhere in the workbook**. The tempting alternative,
`LARGE(IF(...))` over a computed array, reads better in Python and
quietly returns the wrong answer in Excel 2010.

The ranks in use: `Ingredients.need_rank` (what to buy next),
`Grocery List.buy_rank` (position on the printed list),
`Ingredients.spend_rank` and `waste_rank`, `Recipes.cost_rank`,
`match_rank`, `rotation_rank`, `times_planned_rank`, `calories_rank`,
`protein_rank` and `fibre_rank`, `Pantry.fridge_rank` and
`freezer_rank`, `Leftovers.queue_rank`, `Meal Prep.todo_rank` and
`Stores.spend_rank`."""


def outline(path: str) -> list:
    """`[(kind, name, note)]` for the classes, functions and constants."""
    with open(os.path.join(ROOT, path), encoding="utf-8") as handle:
        source = handle.read()
    try:
        tree = ast.parse(source)
    except SyntaxError:                                  # pragma: no cover
        return []
    items = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            args = [a.arg for a in node.args.args]
            if node.args.vararg:
                args.append("*" + node.args.vararg.arg)
            if node.args.kwonlyargs:
                args.append("*")
                args += [a.arg for a in node.args.kwonlyargs]
            if node.args.kwarg:
                args.append("**" + node.args.kwarg.arg)
            note = first_sentence(ast.get_docstring(node) or "")
            items.append(("function", f"{node.name}({', '.join(args)})",
                          note))
        elif isinstance(node, ast.ClassDef):
            bases = []
            for base in node.bases:
                try:
                    bases.append(ast.unparse(base))
                except Exception:                        # pragma: no cover
                    pass
            name = node.name + (f"({', '.join(bases)})" if bases else "")
            note = first_sentence(ast.get_docstring(node) or "")
            fields = [n.target.id for n in node.body
                      if isinstance(n, ast.AnnAssign)
                      and isinstance(n.target, ast.Name)]
            if fields and not note:
                note = "fields: " + ", ".join(fields[:8])
            items.append(("class", name, note))
        elif isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = ([node.target] if isinstance(node, ast.AnnAssign)
                       else node.targets)
            for target in targets:
                if not isinstance(target, ast.Name):
                    continue
                if not re.match(r"^[A-Z][A-Z0-9_]*$", target.id):
                    continue
                items.append(("constant", target.id,
                              describe_value(node.value)))
    return items


def describe_value(node) -> str:
    """A short, human description of a module level constant."""
    if node is None:
        return ""
    if isinstance(node, ast.Constant):
        text = repr(node.value)
        return text if len(text) < 60 else f"{text[:57]}..."
    if isinstance(node, (ast.List, ast.Tuple)):
        return f"{len(node.elts)} items"
    if isinstance(node, ast.Dict):
        return f"{len(node.keys)} entries"
    if isinstance(node, ast.Call):
        try:
            name = ast.unparse(node.func)
        except Exception:                                # pragma: no cover
            name = "call"
        return f"{name}(...)"
    try:
        text = ast.unparse(node)
    except Exception:                                    # pragma: no cover
        return ""
    return text if len(text) < 60 else f"{text[:57]}..."


def first_sentence(text: str) -> str:
    text = " ".join(text.split())
    if not text:
        return ""
    cut = re.search(r"\.(\s|$)", text)
    if cut:
        text = text[:cut.start()]
    return text.strip().rstrip(".")


def blurb(path: str) -> str:
    """First sentence of the module docstring, as a one line description."""
    with open(os.path.join(ROOT, path), encoding="utf-8") as handle:
        source = handle.read()
    try:
        text = ast.get_docstring(ast.parse(source)) or ""
    except SyntaxError:                                  # pragma: no cover
        text = ""
    return first_sentence(text)


def full_docstring(path: str) -> str:
    with open(os.path.join(ROOT, path), encoding="utf-8") as handle:
        source = handle.read()
    try:
        return (ast.get_docstring(ast.parse(source)) or "").strip()
    except SyntaxError:                                  # pragma: no cover
        return ""


def anchor(heading: str) -> str:
    slug = heading.lower()
    slug = re.sub(r"[^\w\s-]", "", slug)
    return re.sub(r"\s+", "-", slug.strip())


def tab_table(product: dict) -> list:
    """A row per tab: order, name, kind, module, columns, rows.

    Imports the product package, so it degrades to nothing if the
    generator is not importable from where the tool is run.
    """
    if ROOT not in sys.path:
        sys.path.insert(0, ROOT)
    try:
        package = __import__(f"{product['package']}.build",
                             fromlist=["build"])
        config = __import__(f"{product['package']}.config",
                            fromlist=["config"])
    except Exception:                                    # pragma: no cover
        return []
    cfg = config.Config()
    lines = ["| # | Tab | Kind | Written by | Columns | Rows |",
             "|---|-----|------|------------|--------:|-----:|"]
    for number, name in enumerate(package.SHEET_ORDER, start=1):
        module = package.REGISTER_BY_NAME.get(name)
        if module is not None:
            spec = module.spec(cfg)
            # a register is usually a module, but the platform planners
            # are instances of one shared class
            where = (getattr(module, "__name__", None)
                     or type(module).__module__)
            path = where.replace(".", "/") + ".py"
            lines.append(f"| {number} | **{name}** | register | "
                         f"`{path}` | {len(spec.columns)} | {spec.rows} |")
        else:
            builder = package.PAGE_BUILDERS.get(name)
            path = (getattr(builder, "__module__", "")
                    .replace(".", "/") + ".py")
            lines.append(f"| {number} | **{name}** | page | `{path}` "
                         f"| - | - |")
    return lines


def data_dictionary(product: dict) -> list:
    """One table per register: every column, its type and where it comes
    from. Generated from the specs themselves, so it cannot drift."""
    if ROOT not in sys.path:
        sys.path.insert(0, ROOT)
    try:
        package = __import__(f"{product['package']}.build",
                             fromlist=["build"])
        config = __import__(f"{product['package']}.config",
                            fromlist=["config"])
    except Exception:                                    # pragma: no cover
        return []
    cfg = config.Config()
    lines = []
    for name in package.SHEET_ORDER:
        module = package.REGISTER_BY_NAME.get(name)
        if module is None:
            continue
        spec = module.spec(cfg)
        typed = sum(1 for c in spec.columns if not c.is_formula)
        lines += [f"#### {spec.name}", ""]
        if spec.subtitle:
            lines += [f"{spec.subtitle}", ""]
        lines += [f"{len(spec.columns)} columns, {spec.rows} rows of "
                  f"capacity, {typed} of them typed in by hand.", "",
                  "| Col | Header | Key | Format | Filled by |",
                  "|-----|--------|-----|--------|-----------|"]
        for col in spec.columns:
            source = "formula" if col.is_formula else "you type it"
            if col.validate:
                picked = str(col.validate.get("source", "")).lstrip("=")
                if picked and re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", picked):
                    source = f"drop down `{picked}`"
                elif not col.is_formula:
                    source = "you type it (checked)"
            if col.hidden:
                source += ", hidden"
            lines.append(f"| {spec.letter(col.key)} | {col.header} | "
                         f"`{col.key}` | {col.style} | {source} |")
        lines.append("")
    return lines


def file_table(sections: list) -> list:
    lines = ["| File | Lines | What it is |", "|------|------:|------------|"]
    for _heading, _intro, paths in sections:
        for path in paths:
            with open(os.path.join(ROOT, path), encoding="utf-8") as handle:
                count = len(handle.read().splitlines())
            note = blurb(path) or ""
            if len(note) > 110:
                note = note[:107].rsplit(" ", 1)[0] + "..."
            lines.append(f"| [`{path}`](#{anchor(path)}) | {count:,} | "
                         f"{note} |")
    return lines


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--product", choices=sorted(PRODUCTS),
                        default="pet",
                        help="which generator to export (default: pet)")
    parser.add_argument("-o", "--output", default=None)
    parser.add_argument("--plain", action="store_true",
                        help="just the code, without the architecture "
                             "chapter and the per file outlines")
    args = parser.parse_args(argv)

    product = PRODUCTS[args.product]
    sections = product["sections"]
    detailed = not args.plain

    lines = [f"# {product['title']} — source code", "",
             product["subtitle"], ""]
    lines += product["intro"]
    lines += ["", "---", "", "## Contents", ""]

    if detailed:
        lines += ["* [0. How it is put together](#0-how-it-is-put-together)",
                  "  * [The build pipeline](#the-build-pipeline)",
                  "  * [Describing a column](#describing-a-column)",
                  "  * [The rules every formula obeys]"
                  "(#the-rules-every-formula-obeys)",
                  f"  * [{product.get('engine_heading', 'The hidden `Calc` tab')}]"
                  f"(#{anchor(product.get('engine_heading', 'The hidden Calc tab'))})",
                  "  * [Every tab at a glance](#every-tab-at-a-glance)",
                  "  * [Every file at a glance](#every-file-at-a-glance)"]
    for heading, _intro, paths in sections:
        lines.append(f"* [{heading}](#{anchor(heading)})")
        for path in paths:
            lines.append(f"  * [`{path}`](#{anchor(path)})")
    if detailed:
        lines.append("* [Appendix A. Data dictionary]"
                     "(#appendix-a-data-dictionary)")
    lines += ["", "---", ""]

    if detailed:
        lines += ["## 0. How it is put together", ""]
        lines += [ARCHITECTURE.format(
            builder=product["builder"], package=product["package"],
            engine=(RANK_ENGINE if product.get("engine_heading")
                    else CALC_ENGINE),
            closing=product.get("closing", DEFAULT_CLOSING)), ""]
        tabs = tab_table(product)
        if tabs:
            lines += ["### Every tab at a glance", "",
                      f"{product['tabs']} tabs in the order they appear in "
                      f"the finished workbook. \"Rows\" is the capacity of "
                      f"a register at the default `--rows` setting.", ""]
            lines += tabs + [""]
        lines += ["### Every file at a glance", ""]
        lines += file_table(sections) + ["", "---", ""]

    total = 0
    for heading, intro, paths in sections:
        lines += [f"## {heading}", "", intro, ""]
        for path in paths:
            full = os.path.join(ROOT, path)
            with open(full, encoding="utf-8") as handle:
                code = handle.read().rstrip("\n")
            count = len(code.splitlines())
            total += count
            lines.append(f"### `{path}`")
            lines.append("")
            if detailed:
                items = outline(path)
                classes = sum(1 for kind, _n, _d in items
                              if kind == "class")
                functions = sum(1 for kind, _n, _d in items
                                if kind == "function")
                stats = [f"{count:,} lines"]
                if classes:
                    stats.append(f"{classes} class"
                                 f"{'es' if classes > 1 else ''}")
                if functions:
                    stats.append(f"{functions} function"
                                 f"{'s' if functions > 1 else ''}")
                lines += [f"*{' · '.join(stats)}*", ""]
                doc = full_docstring(path)
                if doc:
                    lines += ["> " + line if line else ">"
                              for line in doc.splitlines()]
                    lines.append("")
                named = [item for item in items if item[0] != "constant"]
                constants = [item for item in items if item[0] == "constant"]
                if named:
                    lines += ["| Defines | What it does |",
                              "|---------|--------------|"]
                    for kind, name, note in named:
                        tag = "class " if kind == "class" else ""
                        lines.append(f"| `{tag}{name}` | {note} |")
                    lines.append("")
                if constants:
                    text = ", ".join(f"`{name}` ({note})" if note
                                     else f"`{name}`"
                                     for _k, name, note in constants[:14])
                    more = (f" and {len(constants) - 14} more"
                            if len(constants) > 14 else "")
                    lines += [f"**Module constants:** {text}{more}.", ""]
            else:
                note = blurb(path)
                if note:
                    lines.append(f"{note}. *({count} lines)*")
                else:
                    lines.append(f"*({count} lines)*")
                lines.append("")
            # four backticks, so a fence inside a docstring cannot
            # close the block early
            lines += ["````python", code, "````", ""]
        lines.append("---")
        lines.append("")

    if detailed:
        dictionary = data_dictionary(product)
        if dictionary:
            lines += ["## Appendix A. Data dictionary", "",
                      "Every column of every register tab, straight from "
                      "the specs that build them: the Excel column letter, "
                      "the header a buyer sees, the key formulas use, the "
                      "number format and whether the cell is typed in, "
                      "chosen from a drop down or calculated. Pages "
                      "(dashboards, reports, printables) are not listed - "
                      "they have no fixed grid.", ""]
            lines += dictionary + ["---", ""]

    lines += [f"*{total:,} lines of Python in "
              f"{sum(len(p) for _h, _i, p in sections)} files.*", ""]

    out = args.output or product["output"]
    with open(out, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines))
    size = os.path.getsize(out)
    print(f"{out}  ({total:,} lines of source, {size / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
````

### `tools/make_download_zip.py`

*232 lines · 2 functions*

> Pack a generator into a single zip people can download and run.
>
>     python tools/make_download_zip.py
>     python tools/make_download_zip.py --product social
>     python tools/make_download_zip.py --product meal
>
> Writes `download/<product>-source.zip` containing the generator, the
> quality tools, the documentation and a `requirements.txt`, all under one
> folder so it unpacks tidily.
>
> The vendored copy of XlsxWriter in this repository is deliberately left
> out: the zip is meant to be unpacked anywhere and run against a normal
> `pip install XlsxWriter`.  Timestamps are fixed so rebuilding the zip
> without changing any source produces exactly the same bytes.

| Defines | What it does |
|---------|--------------|
| `gather(product)` | `[(archive name, text)]` for everything that goes in the zip |
| `main(argv)` |  |

**Module constants:** `ROOT` (os.path.dirname(...)), `STAMP` (6 items), `SHARED_TOOLS` (3 items), `REQUIREMENTS` ('# The generator needs one package.\nXlsxWriter>=3.0\n\n#...), `PET_HOW_TO_RUN` ('Pet Sitting & Dog Walking Business Spreadsheet\n========...), `SOCIAL_HOW_TO_RUN` ('Social Media Content Calendar & Analytics Planner\n=====...), `MEAL_HOW_TO_RUN` ('Automated Meal Planner & Grocery List Spreadsheet\n=====...), `PRODUCTS` (3 entries).

````python
"""Pack a generator into a single zip people can download and run.

    python tools/make_download_zip.py
    python tools/make_download_zip.py --product social
    python tools/make_download_zip.py --product meal

Writes `download/<product>-source.zip` containing the generator, the
quality tools, the documentation and a `requirements.txt`, all under one
folder so it unpacks tidily.

The vendored copy of XlsxWriter in this repository is deliberately left
out: the zip is meant to be unpacked anywhere and run against a normal
`pip install XlsxWriter`.  Timestamps are fixed so rebuilding the zip
without changing any source produces exactly the same bytes.
"""

from __future__ import annotations

import argparse
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#: Same date for every entry, so the zip only changes when the code does.
STAMP = (2026, 1, 1, 0, 0, 0)

SHARED_TOOLS = ["tools/recalc_check.py", "tools/export_source_md.py",
                "tools/make_download_zip.py"]

REQUIREMENTS = """\
# The generator needs one package.
XlsxWriter>=3.0

# Only for the optional checks in tools/
# openpyxl>=3.1     the validator and the sample data check
# formulas>=1.3     recalc_check.py
"""

PET_HOW_TO_RUN = """\
Pet Sitting & Dog Walking Business Spreadsheet
==============================================

    pip install XlsxWriter
    python build_pet_sitting_workbook.py

The workbook appears in:

    output/Pet Sitting Business Spreadsheet.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --business "My Name"    your business name on the cover and invoices
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --unit miles            mileage in miles instead of km
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 1000             more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

README.md explains how it is built, what is on each of the 27 tabs and
how to change things.  SOURCE_CODE.md is the whole program in one
readable document.
"""

SOCIAL_HOW_TO_RUN = """\
Social Media Content Calendar & Analytics Planner
=================================================

    pip install XlsxWriter
    python build_social_media_workbook.py

The workbook appears in:

    output/Social Media Content Planner.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --brand "My Brand"      your brand name on the cover and dashboards
    --year 2027             the year the 12 month calendar is built for
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 600              more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

Start on the "Start Here" tab inside the workbook.  README_SOCIAL.md
explains how it is built and what is on each of the 43 tabs;
SOURCE_CODE_SOCIAL.md is the whole program in one readable document.
"""

MEAL_HOW_TO_RUN = """\
Automated Meal Planner & Grocery List Spreadsheet
==================================================

    pip install XlsxWriter
    python build_meal_planner_workbook.py

The workbook appears in:

    output/Meal Planner & Grocery List.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --household "Our House" your household name on the cover and prints
    --cook "Your Name"      the main cook / planner
    --servings 6            how many people a meal is normally cooked for
    --budget 200            the weekly food budget
    --units Imperial        lead with cups, ounces and pounds
    --year 2027             the year the planner is built for
    --week-start Sunday     weeks run Sunday to Saturday
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 200              more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

Fill in Settings and Household first, then Ingredients and Recipes, then
plan some meals - the Grocery List builds itself from the plan and takes
off whatever the Pantry says you already have.  README_MEAL.md explains
how it is built and what is on each of the 36 tabs; SOURCE_CODE_MEAL.md
is the whole program in one readable document.
"""


PRODUCTS = {
    "pet": {
        "folder": "pet-sitting-spreadsheet",
        "output": os.path.join("download",
                               "pet-sitting-spreadsheet-source.zip"),
        "files": ["build_pet_sitting_workbook.py", "README.md",
                  "SOURCE_CODE.md"],
        "trees": ["pet_sitting_spreadsheet"],
        "tools": ["tools/validate_workbook.py", "tools/check_sample_data.py"]
                 + SHARED_TOOLS,
        "how_to_run": PET_HOW_TO_RUN,
    },
    "social": {
        "folder": "social-media-content-planner",
        "output": os.path.join(
            "download", "social-media-content-planner-source.zip"),
        "files": ["build_social_media_workbook.py", "README_SOCIAL.md",
                  "SOURCE_CODE_SOCIAL.md"],
        "trees": ["social_media_spreadsheet"],
        "tools": ["tools/validate_social_workbook.py",
                  "tools/check_social_sample.py"] + SHARED_TOOLS,
        "how_to_run": SOCIAL_HOW_TO_RUN,
    },
    "meal": {
        "folder": "meal-planner-grocery-list",
        "output": os.path.join(
            "download", "meal-planner-grocery-list-source.zip"),
        "files": ["build_meal_planner_workbook.py", "README_MEAL.md",
                  "SOURCE_CODE_MEAL.md"],
        "trees": ["meal_planner_spreadsheet"],
        "tools": ["tools/validate_meal_workbook.py",
                  "tools/check_meal_sample.py"] + SHARED_TOOLS,
        "how_to_run": MEAL_HOW_TO_RUN,
    },
}


def gather(product: dict) -> list:
    """`[(archive name, text)]` for everything that goes in the zip."""
    items = []
    for name in product["files"]:
        with open(os.path.join(ROOT, name), encoding="utf-8") as handle:
            items.append((name, handle.read()))
    for tree in product["trees"]:
        for folder, dirs, names in os.walk(os.path.join(ROOT, tree)):
            dirs[:] = sorted(d for d in dirs if d != "__pycache__")
            for name in sorted(names):
                if not name.endswith(".py"):
                    continue
                full = os.path.join(folder, name)
                rel = os.path.relpath(full, ROOT).replace(os.sep, "/")
                with open(full, encoding="utf-8") as handle:
                    items.append((rel, handle.read()))
    for name in product["tools"]:
        with open(os.path.join(ROOT, name), encoding="utf-8") as handle:
            items.append((name, handle.read()))
    items.append(("requirements.txt", REQUIREMENTS))
    items.append(("HOW TO RUN.txt", product["how_to_run"]))
    return items


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--product", choices=sorted(PRODUCTS),
                        default="pet",
                        help="which generator to pack (default: pet)")
    parser.add_argument("-o", "--output", default=None)
    args = parser.parse_args(argv)

    product = PRODUCTS[args.product]
    out = args.output or product["output"]
    folder = os.path.dirname(os.path.abspath(out))
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)

    items = gather(product)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, text in items:
            info = zipfile.ZipInfo(f"{product['folder']}/{name}",
                                   date_time=STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            zf.writestr(info, text)

    size = os.path.getsize(out)
    lines = sum(len(text.splitlines()) for name, text in items
                if name.endswith(".py"))
    print(f"{out}  ({len(items)} files, {lines:,} lines of Python, "
          f"{size / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
````

---

## Appendix A. Data dictionary

Every column of every register tab, straight from the specs that build them: the Excel column letter, the header a buyer sees, the key formulas use, the number format and whether the cell is typed in, chosen from a drop down or calculated. Pages (dashboards, reports, printables) are not listed - they have no fixed grid.

#### Household

Who eats here, what they can and cannot eat, and how much of it - the numbers behind every portion in the workbook.

25 columns, 12 rows of capacity, 21 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Name | `name` | text | you type it |
| C | Role | `role` | text | drop down `MemberRoleList` |
| D | Age group | `age_group` | text | drop down `AgeGroupList` |
| E | Portion factor | `portion` | number1 | you type it (checked) |
| F | Diet | `diet` | text | drop down `DietList` |
| G | Allergies | `allergies` | text | drop down `AllergenList` |
| H | Avoids | `intolerances` | text | you type it |
| I | Will not touch | `dislikes` | text | you type it |
| J | Favourites | `loves` | text | you type it |
| K | In for breakfast? | `breakfast` | flag | drop down `YesNoList` |
| L | In for lunch? | `lunch` | flag | drop down `YesNoList` |
| M | In for dinner? | `dinner` | flag | drop down `YesNoList` |
| N | Takes a packed lunch? | `packed` | flag | drop down `YesNoList` |
| O | Daily calories | `cal_target` | number | you type it (checked) |
| P | Calories used | `eff_cals` | number | formula, hidden |
| Q | Cooks? | `cooks` | flag | drop down `YesNoList` |
| R | Cooking nights | `cook_days` | text | you type it |
| S | Does the shopping? | `shops` | flag | drop down `YesNoList` |
| T | Kitchen jobs | `jobs` | text | you type it |
| U | Meals planned for them | `meals_planned` | int | formula |
| V | Birthday | `birthday` | date | you type it (checked) |
| W | Status | `status` | status | drop down `ActiveList` |
| X | Check | `alert` | status | formula |
| Y | Notes | `notes` | wrap | you type it |

#### Recipes

Every recipe you cook, with its cost and nutrition worked out from the ingredients rather than typed in.

58 columns, 200 rows of capacity, 29 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Recipe | `name` | text | you type it |
| C | Category | `category` | text | drop down `RecipeCategoryList` |
| D | Usual meal | `meal_slot` | text | drop down `MealSlotList` |
| E | Cuisine | `cuisine` | text | drop down `CuisineList` |
| F | What it is | `description` | wrap | you type it |
| G | Serves | `servings` | int | you type it (checked) |
| H | Prep (min) | `prep_min` | int | you type it (checked) |
| I | Cook (min) | `cook_min` | int | you type it (checked) |
| J | Total (min) | `total_min` | int | formula |
| K | Quick meal? | `quick` | flag | formula |
| L | Method | `method` | text | drop down `CookingMethodList` |
| M | Difficulty | `difficulty` | text | drop down `DifficultyList` |
| N | Equipment | `equipment` | text | you type it |
| O | One pot? | `one_pot` | flag | drop down `YesNoList` |
| P | No cook? | `no_cook` | flag | drop down `YesNoList` |
| Q | Batch / prep friendly? | `prep_friendly` | flag | drop down `YesNoList` |
| R | Freezes well? | `freezes` | flag | drop down `YesNoList` |
| S | Kid friendly? | `kid_friendly` | flag | drop down `YesNoList` |
| T | Suits | `diet` | text | drop down `DietList` |
| U | Spice level | `spice` | text | drop down `SpiceLevelList` |
| V | Best season | `season` | text | drop down `SeasonList` |
| W | Tags | `tags` | text | you type it |
| X | Where it came from | `source` | text | you type it |
| Y | Link | `link` | link | you type it |
| Z | Photo link | `photo` | link | you type it |
| AA | How to make it | `method_notes` | wrap | you type it |
| AB | Ingredients | `line_count` | int | formula |
| AC | Cost to make | `cost` | money | formula |
| AD | Cost per serving | `cost_serving` | money | formula |
| AE | Calories a serving | `calories` | number | formula |
| AF | Protein (g) | `protein` | number1 | formula |
| AG | Carbs (g) | `carbs` | number1 | formula |
| AH | Fat (g) | `fat` | number1 | formula |
| AI | Fibre (g) | `fibre` | number1 | formula |
| AJ | Allergens in it | `allergens` | text | formula |
| AK | Cost rank | `cost_rank` | int | formula, hidden |
| AL | Matches the filter | `match` | int | formula, hidden |
| AM | Filter rank | `match_rank` | int | formula, hidden |
| AN | Popularity rank | `times_planned_rank` | int | formula, hidden |
| AO | Calorie rank | `calories_rank` | int | formula, hidden |
| AP | Protein rank | `protein_rank` | int | formula, hidden |
| AQ | Fibre rank | `fibre_rank` | int | formula, hidden |
| AR | Rotation rank | `rotation_rank` | int | formula, hidden |
| AS | Sort key | `sort_key` | number1 | formula, hidden |
| AT | Batches in the shopping window | `plan_scale` | number1 | formula, hidden |
| AU | Cost in the window | `plan_cost` | money | formula, hidden |
| AV | Missing a price | `missing_price` | int | formula, hidden |
| AW | Your rating | `rating` | rating | you type it (checked) |
| AX | Family score | `family_score` | rating | you type it (checked) |
| AY | Favourite? | `favourite` | flag | drop down `YesNoList` |
| AZ | Times planned | `times_planned` | int | formula |
| BA | Times cooked | `times_cooked` | int | formula |
| BB | Last cooked | `last_cooked` | date | formula |
| BC | Days since | `days_since` | int | formula |
| BD | Tweaks & notes | `notes` | wrap | you type it |
| BE | Added on | `added` | date | you type it (checked) |
| BF | Check | `alert` | status | formula |

#### Recipe Ingredients

One line per ingredient per recipe. This is what turns a meal plan into a shopping list.

28 columns, 1200 rows of capacity, 7 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Recipe ID | `recipe` | text | drop down `RecipeIDs` |
| C | Recipe | `recipe_name` | text | formula |
| D | Ingredient | `ingredient` | text | drop down `IngredientNames` |
| E | Quantity | `qty` | number1 | you type it (checked) |
| F | Unit | `unit` | text | drop down `UnitList` |
| G | Prepared how | `prep` | text | you type it |
| H | Part of the recipe | `section` | text | you type it |
| I | Optional? | `optional` | flag | drop down `YesNoList` |
| J | In base units | `base_qty` | number1 | formula |
| K | Base unit | `base_unit` | text | formula |
| L | Priced in | `price_unit` | text | formula |
| M | Cost per unit | `unit_cost` | money | formula, hidden |
| N | Line cost | `line_cost` | money | formula |
| O | Cost a serving | `cost_serving` | money | formula, hidden |
| P | Per 100 | `per100` | number | formula, hidden |
| Q | Calories | `line_cals` | number | formula |
| R | Protein (g) | `line_protein` | number1 | formula, hidden |
| S | Carbs (g) | `line_carbs` | number1 | formula, hidden |
| T | Fat (g) | `line_fat` | number1 | formula, hidden |
| U | Fibre (g) | `line_fibre` | number1 | formula, hidden |
| V | Allergen | `allergen` | text | formula |
| W | Shopping category | `category` | text | formula, hidden |
| X | Planned batches | `plan_factor` | number1 | formula, hidden |
| Y | Needed for the plan | `need_base` | number1 | formula |
| Z | Scale key | `scale_key` | text | formula, hidden |
| AA | Key | `key` | text | formula, hidden |
| AB | Check | `alert` | status | formula |

#### Ingredients

Your price book. Every recipe cost, every shopping total and every nutrition figure in the workbook is built from this one list.

41 columns, 300 rows of capacity, 23 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Ingredient | `name` | text | you type it |
| C | Category | `category` | text | drop down `GroceryCategoryList` |
| D | Aisle | `aisle` | text | drop down `AisleList` |
| E | Usual store | `store` | text | drop down `StoreNames` |
| F | Preferred brand | `brand` | text | you type it |
| G | Measured in | `base_unit` | text | drop down `UnitList` |
| H | Pack size | `pack_size` | number1 | you type it (checked) |
| I | Pack price | `pack_price` | money | you type it (checked) |
| J | Cost per unit | `unit_cost` | money | formula |
| K | Cost per 100 | `per_100` | money | formula |
| L | Average price paid | `avg_paid` | money | formula |
| M | Cheapest store price | `best_price` | money | formula |
| N | Cheapest at | `best_store` | text | formula |
| O | Saving per 100 | `saving` | money | formula |
| P | Calories per 100 / item | `calories` | number | you type it (checked) |
| Q | Protein per 100 (g) | `protein` | number1 | you type it (checked) |
| R | Carbs per 100 (g) | `carbs` | number1 | you type it (checked) |
| S | Fat per 100 (g) | `fat` | number1 | you type it (checked) |
| T | Fibre per 100 (g) | `fibre` | number1 | you type it (checked) |
| U | Sugar per 100 (g) | `sugar` | number1 | you type it (checked) |
| V | Sodium per 100 (mg) | `sodium` | number | you type it (checked) |
| W | Allergen | `allergen` | text | drop down `AllergenList` |
| X | Suits | `diet` | text | drop down `DietList` |
| Y | Best season | `season` | text | drop down `SeasonList` |
| Z | Keeps for (days) | `shelf_days` | int | you type it (checked) |
| AA | Store it in | `storage` | text | drop down `StorageList` |
| AB | Swap it for | `substitute` | text | you type it |
| AC | Last bought | `last_bought` | date | you type it (checked) |
| AD | In the pantry now | `in_pantry` | number1 | formula |
| AE | Value in stock | `pantry_value` | money | formula |
| AF | Used in recipes | `used_in` | int | formula |
| AG | Needed for the plan | `need_now` | number1 | formula |
| AH | Spend key | `spend_key` | number | formula, hidden |
| AI | Spend rank | `spend_rank` | int | formula, hidden |
| AJ | Times binned | `waste_count` | int | formula, hidden |
| AK | Waste rank | `waste_rank` | int | formula, hidden |
| AL | Rank | `need_rank` | int | formula, hidden |
| AM | On the list now | `to_buy` | number1 | formula |
| AN | Check | `alert` | status | formula |
| AO | Notes | `notes` | wrap | you type it |

#### Meal Plan

Every planned meal in one list. The weekly and monthly planners, the grocery list and the budget all read from here.

33 columns, 500 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Day | `weekday` | text | formula |
| D | Week | `week` | int | formula |
| E | Month | `month` | text | formula |
| F | Meal | `slot` | text | drop down `MealSlotList` |
| G | Recipe ID | `recipe` | text | drop down `RecipeIDs` |
| H | What you are eating | `recipe_name` | text | formula |
| I | Or type it in | `freetext` | text | you type it |
| J | People to feed | `servings` | int | you type it (checked) |
| K | Servings used | `eff_servings` | number1 | formula, hidden |
| L | Recipe batches | `scale` | number1 | formula |
| M | Shop for it? | `shop` | flag | drop down `YesNoList` |
| N | Status | `status` | status | drop down `MealStatusList` |
| O | Who is cooking | `cook` | text | drop down `MemberNames` |
| P | Who is eating | `eating` | text | you type it |
| Q | Planned leftovers | `leftovers` | number1 | you type it (checked) |
| R | Prep ahead? | `prep_ahead` | flag | drop down `YesNoList` |
| S | Occasion | `event` | text | drop down `EventNames` |
| T | Planned cost | `cost` | money | formula |
| U | Actual spend | `actual_cost` | money | you type it (checked) |
| V | Calories each | `calories` | number | formula |
| W | Protein each (g) | `protein` | number1 | formula, hidden |
| X | Carbs each (g) | `carbs` | number1 | formula, hidden |
| Y | Fat each (g) | `fat` | number1 | formula, hidden |
| Z | Fibre each (g) | `fibre` | number1 | formula, hidden |
| AA | Calories on the table | `cals_total` | number | formula, hidden |
| AB | Minutes in the kitchen | `minutes` | int | formula, hidden |
| AC | Category | `category` | text | formula, hidden |
| AD | Cooked at home? | `home_cooked` | flag | formula, hidden |
| AE | Key | `key` | text | formula, hidden |
| AF | Check | `alert` | status | formula |
| AG | Notes | `notes` | wrap | you type it |

#### Grocery List

Built from your meal plan, minus what is already in the pantry, grouped by aisle and priced as you go.

26 columns, 300 rows of capacity, 4 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | # | `rank` | int | formula, hidden |
| B | Item | `ingredient` | text | formula |
| C | Category | `category` | text | formula |
| D | Aisle | `aisle` | text | formula |
| E | Store | `store` | text | formula |
| F | Needed | `needed` | number1 | formula |
| G | Unit | `unit` | text | formula |
| H | Already have | `in_pantry` | number1 | formula |
| I | Still short | `short` | number1 | formula |
| J | Pack size | `pack_size` | number1 | formula, hidden |
| K | Pack price | `pack_price` | money | formula, hidden |
| L | Cost per unit | `unit_cost` | money | formula, hidden |
| M | Packs to buy | `packs` | number1 | formula |
| N | Buy this much | `buy_qty` | number1 | formula |
| O | Estimated cost | `est_cost` | money | formula |
| P | Running total | `running` | money | formula |
| Q | To buy | `to_buy` | number1 | formula, hidden |
| R | Buy rank | `buy_rank` | int | formula, hidden |
| S | For how many meals | `recipes` | int | formula |
| T | Optional only? | `optional` | flag | formula |
| U | Allergen | `allergen` | text | formula, hidden |
| V | Got it? | `got` | flag | drop down `YesNoList` |
| W | Actually paid | `paid` | money | you type it (checked) |
| X | Swapped for | `swap` | text | you type it |
| Y | Note | `note` | wrap | you type it |
| Z | Status | `status` | status | formula |

#### Pantry

What you already own, what it is worth and what is about to go off - the tab that shrinks your shopping list.

25 columns, 300 rows of capacity, 10 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Ingredient | `ingredient` | text | drop down `IngredientNames` |
| C | Category | `category` | text | formula |
| D | Where it lives | `location` | text | drop down `StorageList` |
| E | How much is left | `qty` | number1 | you type it (checked) |
| F | Unit | `unit` | text | drop down `UnitList` |
| G | In base units | `base_qty` | number1 | formula |
| H | Base unit | `base_unit` | text | formula, hidden |
| I | Opened? | `opened` | flag | drop down `YesNoList` |
| J | Bought on | `bought` | date | you type it (checked) |
| K | Use by | `expiry` | date | you type it (checked) |
| L | Keeps for | `shelf_days` | int | formula, hidden |
| M | Best before (worked out) | `est_expiry` | date | formula |
| N | Days left | `days_left` | int | formula |
| O | Cost per unit | `unit_cost` | money | formula, hidden |
| P | Value | `value` | money | formula |
| Q | Keep at least | `min_qty` | number1 | you type it (checked) |
| R | Plan needs | `needed` | number1 | formula, hidden |
| S | Left after the plan | `after_plan` | number1 | formula |
| T | Staple? | `staple` | flag | drop down `YesNoList` |
| U | Fridge rank | `fridge_rank` | int | formula, hidden |
| V | Freezer rank | `freezer_rank` | int | formula, hidden |
| W | Status | `status` | status | formula |
| X | What to do | `alert` | status | formula |
| Y | Notes | `notes` | wrap | you type it |

#### Staples

Par levels for the things that should always be in the cupboard, with a countdown to the day they run out.

19 columns, 100 rows of capacity, 5 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Staple | `ingredient` | text | drop down `IngredientNames` |
| C | Category | `category` | text | formula |
| D | Unit | `unit` | text | formula |
| E | Always keep | `par` | number1 | you type it (checked) |
| F | In stock | `in_stock` | number1 | formula |
| G | Short by | `gap` | number1 | formula |
| H | Used a week | `usage` | number1 | you type it (checked) |
| I | Weeks left | `weeks_left` | number1 | formula |
| J | Buy before | `buy_by` | date | formula |
| K | Pack size | `pack_size` | number1 | formula, hidden |
| L | Pack price | `pack_price` | money | formula, hidden |
| M | Packs to buy | `packs` | number1 | formula |
| N | Cost to restock | `cost` | money | formula |
| O | Buy it at | `store` | text | formula |
| P | Worth buying big? | `bulk` | flag | drop down `YesNoList` |
| Q | Status | `status` | status | formula |
| R | Check | `alert` | status | formula |
| S | Notes | `notes` | wrap | you type it |

#### Leftovers

What is in the fridge from last night, what it is worth and how long you have to eat it.

19 columns, 120 rows of capacity, 10 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | What it is | `item` | text | you type it |
| C | From which recipe | `recipe` | text | drop down `RecipeIDs` |
| D | Recipe | `recipe_name` | text | formula, hidden |
| E | Cooked on | `cooked` | date | you type it (checked) |
| F | Portions left | `portions` | number1 | you type it (checked) |
| G | Stored in | `location` | text | drop down `StorageList` |
| H | Container | `container` | text | you type it |
| I | Days it keeps | `window` | int | formula, hidden |
| J | Eat it by | `use_by` | date | formula |
| K | Days left | `days_left` | int | formula |
| L | Cost a portion | `cost_portion` | money | formula, hidden |
| M | Worth | `value` | money | formula |
| N | Planned for | `planned_for` | date | you type it (checked) |
| O | Who is having it | `eaten_by` | text | drop down `MemberNames` |
| P | Queue rank | `queue_rank` | int | formula, hidden |
| Q | Status | `status` | status | drop down `LeftoverStatusList` |
| R | What happened | `outcome` | status | formula |
| S | Notes | `notes` | wrap | you type it |

#### Food Waste

What got thrown away, what it cost and what would have stopped it.

20 columns, 150 rows of capacity, 11 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Week | `week` | int | formula |
| D | Month | `month` | text | formula |
| E | What went in the bin | `item` | text | drop down `IngredientNames` |
| F | Or describe it | `other` | text | you type it |
| G | How much | `qty` | number1 | you type it (checked) |
| H | Unit | `unit` | text | drop down `UnitList` |
| I | In base units | `base_qty` | number1 | formula, hidden |
| J | Base unit | `base_unit` | text | formula, hidden |
| K | Cost per unit | `unit_cost` | money | formula, hidden |
| L | What it cost | `value` | money | formula |
| M | Your estimate | `manual_cost` | money | you type it (checked) |
| N | Why | `reason` | text | drop down `WasteReasonList` |
| O | Avoidable? | `avoidable` | flag | drop down `YesNoList` |
| P | Where it was | `where` | text | drop down `StorageList` |
| Q | How it was thrown | `disposal` | text | drop down `DisposalList` |
| R | What would have stopped it | `prevent` | wrap | you type it |
| S | Wasted this month | `month_total` | money | formula, hidden |
| T | Check | `alert` | status | formula |

#### Meal Prep

What you are making ahead, how many portions it gives you and what each one costs in money and minutes.

25 columns, 120 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Prep job | `job` | text | you type it |
| C | Prep day | `date` | date | you type it (checked) |
| D | Day | `weekday` | text | formula |
| E | Recipe | `recipe` | text | drop down `RecipeIDs` |
| F | What you are cooking | `recipe_name` | text | formula |
| G | Batches | `batches` | number1 | you type it (checked) |
| H | Serves (each) | `servings_each` | int | formula, hidden |
| I | Portions made | `portions` | number1 | formula |
| J | Minutes (recipe) | `minutes_each` | int | formula, hidden |
| K | Time needed | `minutes` | int | you type it (checked) |
| L | Minutes used | `eff_minutes` | int | formula, hidden |
| M | Minutes a portion | `per_portion` | number1 | formula |
| N | Stored in | `storage` | text | drop down `StorageList` |
| O | Containers | `containers` | int | you type it (checked) |
| P | Covers meals from | `covers_from` | date | you type it (checked) |
| Q | Covers meals to | `covers_to` | date | you type it (checked) |
| R | Meals it covers | `meals_covered` | int | formula |
| S | Who is doing it | `cook` | text | drop down `MemberNames` |
| T | Cost of the batch | `cost` | money | formula |
| U | Cost a portion | `cost_portion` | money | formula |
| V | To-do rank | `todo_rank` | int | formula, hidden |
| W | Status | `status` | status | drop down `PrepStatusList` |
| X | Check | `alert` | status | formula |
| Y | Notes | `notes` | wrap | you type it |

#### Grocery Spending

Every shop, top-up and takeaway, so the budget tabs can compare what you planned with what you paid.

23 columns, 400 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Day | `weekday` | text | formula |
| D | Week | `week` | int | formula |
| E | Month | `month` | text | formula |
| F | Store | `store` | text | drop down `StoreNames` |
| G | Type of spend | `type` | text | drop down `SpendTypeList` |
| H | Item or description | `item` | text | drop down `IngredientNames` |
| I | Quantity | `qty` | number1 | you type it (checked) |
| J | Unit | `unit` | text | drop down `UnitList` |
| K | Amount | `amount` | money | you type it (checked) |
| L | Price per unit | `unit_price` | money | formula, hidden |
| M | Who paid | `payer` | text | drop down `MemberNames` |
| N | Paid with | `method` | text | drop down `PaymentList` |
| O | Planned? | `planned` | flag | drop down `YesNoList` |
| P | Budget line | `budget_line` | text | drop down `BudgetLineList` |
| Q | Thrown away? | `waste` | flag | drop down `YesNoList` |
| R | Receipt link | `receipt` | link | you type it |
| S | That month's budget | `month_budget` | money | formula, hidden |
| T | Spent this month | `running_month` | money | formula |
| U | Over budget? | `over` | status | formula, hidden |
| V | Status | `status` | status | formula |
| W | Notes | `notes` | wrap | you type it |

#### Store Prices

What the same ingredient costs in each shop, reduced to one comparable number.

17 columns, 300 rows of capacity, 8 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Ingredient | `ingredient` | text | drop down `IngredientNames` |
| C | Store | `store` | text | drop down `StoreNames` |
| D | Brand | `brand` | text | you type it |
| E | Pack size | `pack_size` | number1 | you type it (checked) |
| F | Unit | `unit` | text | formula |
| G | Pack price | `price` | money | you type it (checked) |
| H | On offer? | `offer` | flag | drop down `YesNoList` |
| I | Cost per unit | `unit_cost` | money | formula, hidden |
| J | Cost per 100 | `per_100` | money | formula |
| K | Price book cost | `book_cost` | money | formula, hidden |
| L | Against the price book | `vs_book` | money | formula |
| M | Cheapest so far? | `best` | flag | formula |
| N | Price checked | `checked` | date | you type it (checked) |
| O | How old | `age_days` | int | formula |
| P | Check | `alert` | status | formula |
| Q | Notes | `notes` | wrap | you type it |

#### Stores

Where you shop, what each one is good for and what each one costs you.

24 columns, 40 rows of capacity, 15 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Store | `name` | text | you type it |
| C | Type | `type` | text | drop down `StoreTypeList` |
| D | Where | `location` | text | you type it |
| E | Minutes away | `travel_min` | int | you type it (checked) |
| F | Best for | `best_for` | text | you type it |
| G | Delivers? | `delivery` | flag | drop down `YesNoList` |
| H | Delivery fee | `delivery_fee` | money | you type it (checked) |
| I | Minimum order | `min_order` | money | you type it (checked) |
| J | Usual slot | `slot` | text | you type it |
| K | Loyalty card | `loyalty` | text | you type it |
| L | Opening hours | `open_hours` | text | you type it |
| M | Price level | `price_level` | text | drop down `PriceLevelList` |
| N | Your rating | `rating` | rating | you type it (checked) |
| O | Items priced here | `items_priced` | int | formula |
| P | Cheapest on | `cheapest_on` | int | formula |
| Q | Spent here | `spent` | money | formula |
| R | Shops logged | `visits` | int | formula |
| S | Average shop | `avg_shop` | money | formula |
| T | On this week's list | `on_list` | int | formula |
| U | Spend rank | `spend_rank` | int | formula, hidden |
| V | Status | `status` | status | drop down `ActiveList` |
| W | Check | `alert` | status | formula |
| X | Notes | `notes` | wrap | you type it |

#### Events

Occasions with a head count, a budget and a timetable - and a menu that flows into the normal shopping list.

23 columns, 40 rows of capacity, 16 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Occasion | `name` | text | you type it |
| C | Type | `type` | text | drop down `EventTypeList` |
| D | Date | `date` | date | you type it (checked) |
| E | Days to go | `days_to_go` | int | formula |
| F | Time | `time` | time | you type it |
| G | Guests | `guests` | int | you type it (checked) |
| H | Adults | `adults` | int | you type it (checked) |
| I | Children | `children` | int | you type it (checked) |
| J | Portions to cook | `portions` | number1 | formula |
| K | Dietary needs | `dietary` | text | you type it |
| L | Budget | `budget` | money | you type it (checked) |
| M | Planned food cost | `planned_cost` | money | formula |
| N | Actual cost | `actual_cost` | money | you type it (checked) |
| O | Against budget | `vs_budget` | money | formula |
| P | Meals planned | `meals` | int | formula |
| Q | Menu | `menu` | wrap | you type it |
| R | Start cooking | `prep_start` | date | you type it (checked) |
| S | Shopping done? | `shopping_done` | flag | drop down `YesNoList` |
| T | Who is hosting | `host` | text | drop down `MemberNames` |
| U | Status | `status` | status | drop down `EventStatusList` |
| V | Check | `alert` | status | formula |
| W | Notes | `notes` | wrap | you type it |

#### Goals

Three or four goals, measured by the numbers already in this workbook.

18 columns, 40 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Goal | `goal` | text | you type it |
| C | Area | `area` | text | drop down `GoalCategoryList` |
| D | Measured by | `measure` | text | drop down `MeasureList` |
| E | Target | `target` | number1 | you type it (checked) |
| F | Aim to | `direction` | text | drop down `DirectionList` |
| G | Measured now | `auto` | number1 | formula, hidden |
| H | Your own count | `manual` | number1 | you type it (checked) |
| I | Where you are | `actual` | number1 | formula |
| J | Progress | `progress` | percent0 | formula |
| K | On track? | `on_track` | flag | formula |
| L | Why it matters | `why` | wrap | you type it |
| M | Started | `start` | date | you type it (checked) |
| N | Review on | `review` | date | you type it (checked) |
| O | Whose goal | `owner` | text | drop down `MemberNames` |
| P | Status | `status` | status | drop down `GoalStatusList` |
| Q | Check | `alert` | status | formula |
| R | Notes | `notes` | wrap | you type it |

---

*11,435 lines of Python in 52 files.*
