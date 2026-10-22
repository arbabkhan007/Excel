# Social Media Content Calendar & Analytics Planner — generator

> **This repository holds six spreadsheet products** built on the same
> engine, with the same conventions, the same author (Novality Store) and
> the same `premium` sheet protection:
>
> * **Pet Sitting & Dog Walking Business Spreadsheet** —
>   [README.md](README.md), `build_pet_sitting_workbook.py`, 27 tabs.
> * **Social Media Content Calendar & Analytics Planner** — this
>   README, `build_social_media_workbook.py`, 43 tabs.
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
>   [README_SALON.md](README_SALON.md),
>   `build_salon_workbook.py`, 20 tabs.

Python code that builds a finished, ready-to-sell Excel workbook for
anyone who plans social media content: a 12-month calendar, the banks and
trackers around it, and the dashboards that read it back.

> **Listing title:** Social Media Content Planner Spreadsheet, 12-Month
> Content Calendar, Instagram, TikTok, Pinterest, YouTube and Facebook
> Analytics Tracker
> **Author / creator of every generated file:** Novality Store
> **Sheet protection password:** `premium` (every formula cell is locked)

```bash
python build_social_media_workbook.py
# -> output/Social Media Content Planner.xlsx   (43 tabs, ~1.5 MB)
```

The repository also contains the XlsxWriter library itself (`xlsxwriter/`),
so the command above runs with **no installation and no dependencies** from
the repository root. Outside this repo, `pip install xlsxwriter` first — the
generator is tested against the bundled 0.3.0 and against XlsxWriter 3.2.

The companion products in this repository are the
[Pet Sitting & Dog Walking Business Spreadsheet](README.md), the
[Automated Meal Planner & Grocery List Spreadsheet](README_MEAL.md)
and the
[Cleaning Business CRM & Pricing Calculator](README_CLEANING.md);
all four share the same engine, the same conventions and the same
quality tools.

---

## 1. What gets built

43 tabs, 110,000+ formulas, all joined up and all driven by drop downs.

| # | Tab | What it does |
|---|-----|--------------|
| 1 | **Start Here** | Welcome page, the six-step set-up, the weekly rhythm, a map of every tab, the colour key and an honest note on what a spreadsheet can and cannot do |
| 2 | **Dashboard** | Pick a month (or "All year"): 24 KPI tiles, a 12-month table with sparklines, platform and pillar breakdowns, top 8 posts, the alert list, what is coming up and what is overdue |
| 3 | **Annual Dashboard** | The same year read four ways — month by month, quarter by quarter, platform by platform and pillar by pillar — plus goals, budget and the year's best posts |
| 4 | **Settings** | Brand details, plan year, week start, targets, budget, prefixes and the reminder windows every other tab reads |
| 5 | **Lists** | Every drop down in the workbook, 42 named lists, editable in place |
| 6 | **Brand Voice** | Tone words, what you never say, the hashtag and emoji rules, colours, fonts, the bio and a caption checklist |
| 7 | **Accounts** | Every profile: handle, type, audience, followers, a six-point optimisation score, this year's reach and engagement, and a "gone quiet" warning |
| 8 | **Audience Personas** | Who you are talking to, what they want, what stops them, and how much of the plan is actually aimed at them |
| 9 | **Content Pillars** | Your themes, the share of posts each should get, and how far the real plan drifts from it |
| 10 | **Content Calendar** | **The main tab.** One row per post for the whole year: date, slot, platform, title, pillar, format, campaign, objective, funnel stage, persona, caption, hashtag group, asset, CTA, owner, status, deadlines, link and results |
| 11 | **Month View** | The same calendar as a printable six-week grid, three posts per day, today highlighted |
| 12 | **Workflow** | The production pipeline — idea, brief, draft, review, approved, scheduled, published — with counts, what is late and what is due this week |
| 13 | **Idea Bank** | Every idea with a score out of 10, its source, the pillar it serves and whether it has been scheduled yet |
| 14 | **Caption Bank** | Hooks, bodies, CTAs, the full caption, a character count against each platform's limit, and how often each one has been used |
| 15 | **Hashtag Bank** | Hashtags grouped into reusable sets, with type, size band, search intent, competition and days rested |
| 16-20 | **Instagram, TikTok, Pinterest, YouTube, Facebook** | One planner per platform with the fields only that platform needs — carousel slides, sounds, board names, thumbnails, link stickers — and the shared metrics pulled in from Performance, never retyped |
| 21 | **Campaigns** | Themed pushes with a budget, dates, posts, reach, leads, revenue, cost per lead and return on spend |
| 22 | **Launch Planner** | A launch broken into teaser, launch day and follow-up, with the content each phase needs |
| 23 | **Seasonal Planner** | Holidays and awareness days with a "start planning on" date that counts down |
| 24 | **Influencers** | Outreach status, fee, deliverables, code, and what the collaboration actually returned |
| 25 | **UGC Tracker** | Customer content, permission, usage rights and when those rights expire |
| 26 | **Repurposing** | One good post, used again: where it goes next, in what format, by when |
| 27 | **Competitors** | Who you are up against, what they post, how often, and where the gap is |
| 28 | **Trends** | Sounds, formats and topics worth jumping on, with a window that closes |
| 29 | **Analytics** | **Type here once a month.** One row per account per month: followers, reach, impressions, engagements, clicks, leads, sales, revenue, spend — the rest is calculated |
| 30 | **Performance** | Post by post results keyed on the content ID: engagement, save and share rates, revenue per 1,000 reached, rank and a verdict |
| 31 | **AB Tests** | Change one thing, measure it, keep the winner — with an uplift and a "how much to trust it" read |
| 32 | **Ads** | Paid promotion: budget, spend, reach, clicks, sales, cost per click, cost per sale and return |
| 33 | **Assets** | The creative library with links, rights, expiry dates and how often each asset has been used |
| 34 | **Goals** | The numbers you are trying to move, with progress bars and days left |
| 35 | **Budget** | What social media costs month by month, planned against actual, with the variance |
| 36 | **Team** | Who does what, their rate, their posts, their open and overdue tasks |
| 37 | **Tasks** | The jobs behind the posts, with owners, due dates and a countdown |
| 38 | **Compliance** | Disclosure, rights, accessibility and quality — a 24-point checklist with a score |
| 39 | **Reports** | Pick two dates: headline numbers, platform and pillar tables, top posts, campaigns and a written summary line, ready to print or PDF |
| 40 | **Print Brief** | Pick a content ID and the whole brief prints on one page for whoever makes it |
| 41 | **Print Week** | Seven days of posts on one landscape page for the wall |
| 42 | **Help** | How the tabs fit together, 15 questions answered, every metric in plain English, a troubleshooting table and the honest limitations |
| 43 | **Calc** | Hidden calculation engine — months, platforms, pillars, formats, weekdays, slots, owners, campaigns and accounts, pre-aggregated so the dashboards stay fast |

### Product rules baked in

* **Nothing is typed twice.** A post is planned once on the Content
  Calendar, its results are recorded once on Performance, and the account
  totals are recorded once a month on Analytics. Every other tab reads
  those three.
* **Every formula cell is locked**, every sheet is protected with
  `premium`, and the cream coloured cells are the only ones left open.
* **Excel 2010 safe.** `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `SMALL`,
  `LARGE`, `RANK`, `EOMONTH`, `SUMPRODUCT(MAX(...))` — no dynamic arrays,
  no `XLOOKUP`, no `IFS`, nothing that needs Ctrl+Shift+Enter. It opens in
  Google Sheets, LibreOffice and Numbers.
* **No charts, by design** — the vendored XlsxWriter cannot write them.
  Trends are shown with sparklines, data bars and colour scales, which
  survive a Google Sheets import better anyway.
* **Nth-row lookups without array formulas.** Hidden composite key columns
  (`date + ROW()/100000`) plus `SMALL` and `INDEX/MATCH` give "the third
  post on 14 March" and "the next eight posts due" from ordinary formulas.
* **Every date filter is a real date range** (`>=` start, `<=` end), so
  the dashboard, the month view and the reports all agree.

---

## 1b. Download it

Everything is committed to the repository, so these links work straight
from GitHub — right click, *Save link as*.

| What | Size | Link |
|------|-----:|------|
| **The finished workbook, with the example brand filled in** | 1,460 KB | [social-media-content-planner.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/social-media-content-planner.xlsx) |
| **The finished workbook, empty and ready to sell** | 1,419 KB | [social-media-content-planner-blank.xlsx](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/social-media-content-planner-blank.xlsx) |
| **The generator source, zipped** | 295 KB | [social-media-content-planner-source.zip](https://github.com/arbabkhan007/Excel/raw/refs/heads/arena/89c8bec5-excel/download/social-media-content-planner-source.zip) |

The two `.xlsx` files are the product itself: 43 tabs, every formula
cell locked, every sheet protected with `premium`, author Novality
Store. Open the blank one, set the plan year on **Settings**, and start
on the **Start Here** tab.

The zip is for anyone who wants to rebuild or change it — unzip anywhere
and run:

```bash
pip install -r requirements.txt
python build_social_media_workbook.py
```

It holds the generator, the quality tools, this README,
`SOURCE_CODE_SOCIAL.md`, a `requirements.txt` and a plain
`HOW TO RUN.txt`. The vendored copy of XlsxWriter in this repository is
left out on purpose — the zip runs against a normal
`pip install XlsxWriter` (tested on 3.2.0). Rebuild everything after a
code change with:

```bash
python build_social_media_workbook.py -o download/social-media-content-planner.xlsx
python build_social_media_workbook.py --blank -o download/social-media-content-planner-blank.xlsx
python tools/make_download_zip.py --product social
```

---

## 2. Command line

```bash
python build_social_media_workbook.py [options]
python -m social_media_spreadsheet   [options]      # same thing
```

| Option | Default | What it does |
|--------|---------|--------------|
| `-o, --output PATH` | `output/Social Media Content Planner.xlsx` | where to write the file |
| `--blank` | off | no example data, just the empty workbook |
| `--no-protect` | off | leave the sheets unlocked (for development) |
| `--password TEXT` | `premium` | sheet protection password |
| `--author TEXT` | `Novality Store` | document author |
| `--brand TEXT` | `Wildflower Studio` | brand name on the cover and dashboards |
| `--currency SYMBOL` | `$` | currency symbol, e.g. `£`, `€`, `A$` |
| `--currency-code XXX` | `USD` | three letter code used in headings |
| `--year YYYY` | this year | the year the 12-month calendar is built for |
| `--week-start DAY` | `Monday` | `Monday` or `Sunday` — drives the month grid |
| `--rows N` | 600 | rows of capacity in the Content Calendar; every other register scales with it |
| `--today YYYY-MM-DD` | today | pretend it is this date (keeps the demo data reproducible) |
| `--quiet` | off | print only the output path |

```bash
# a blank British workbook for 2027, ready to sell
python build_social_media_workbook.py --blank --year 2027 \
    --currency "£" --currency-code GBP --week-start Monday \
    -o "output/Social Media Planner 2027.xlsx"
```

---

## 3. How the code is organised

```
build_social_media_workbook.py      one line wrapper
social_media_spreadsheet/
    config.py       every default: palette, row counts, targets, brand
    layout.py       SheetSpec / Col and the formula template engine
    formats.py      the whole format book (styles x input/calc/plain)
    writer.py       SheetWriter - banners, nav, sections, KPI tiles, grids
    names.py        the cross sheet defined names
    sample_data.py  one coherent demo brand, generated from a fixed seed
    build.py        tab order, page builders, the workbook assembly
    sheets/
        settings.py lists.py common.py        the setup tabs
        accounts.py personas.py pillars.py    the planning registers
        calendar_sheet.py ideas.py captions.py hashtags.py platforms.py
        campaigns.py launches.py seasonal.py influencers.py ugc.py
        repurposing.py competitors.py trends.py assets.py team.py tasks.py
        compliance.py
        analytics.py performance.py ab_tests.py ads.py goals.py budget.py
        calc.py       the hidden engine
        dashboard.py annual.py month_view.py workflow.py brand_voice.py
        reports.py start_here.py printables.py help_tab.py
tools/              the quality checks
```

A **register** is a data tab. Its module exposes `NAME`, a
`register(cfg, reg)` that returns a `SheetSpec` (a list of `Col`s with
their formulas, widths, drop downs, comments and conditional formats) and
an optional `decorate(sheet, fmts, cfg)` for anything extra underneath the
grid. A **page** is a hand-built tab — the dashboards, the reports, the
printables — written directly with `SheetWriter`.

### Formula templates

Columns are written once, for one row, with placeholders:

```python
Col("er", "Engagement rate", 11, "percent",
    formula='=IFERROR({engagements}{r}/{reach}{r},"")')
```

* `{r}` — the current row, `{r_prev}` — the row above
* `{key}` — the column letter of another column on the same sheet
* `{first}` / `{last}` — the first and last data row
* `[[Sheet.key]]` — an absolute range on another sheet,
  `[[Sheet.key#12]]` one cell of it
* helpers for the patterns Excel 2010 makes awkward: `max_if`, `min_if`,
  `blank_if`, `safe`, `lookup`, `dynamic_name`

The registry renders them per row, so a column's logic lives in exactly
one place.

### Adding something

* **A column** — add a `Col` to that sheet's `register()`. Widths,
  formats, validation, comments and conditional formatting come with it.
* **A drop down** — add the list to `sheets/lists.py`; it becomes a named
  range automatically and any column can point at it with
  `validate=list_from("MyList")`.
* **A tab** — write the module, add it to `sheets/__init__.py`, then to
  `REGISTER_MODULES` (or `PAGE_BUILDERS`) and `SHEET_ORDER` in `build.py`.
* **A KPI tile** — one `sheet.kpi(row, col, label, formula, style)` call.

---

## 4. Quality checks

```bash
python build_social_media_workbook.py -o output/test.xlsx
python tools/validate_social_workbook.py output/test.xlsx   # needs openpyxl
python tools/check_social_sample.py      output/test.xlsx

python build_social_media_workbook.py --rows 60 -o /tmp/small.xlsx
python tools/recalc_check.py /tmp/small.xlsx                # needs formulas
```

`validate_social_workbook.py` reports: tab order and the hidden Calc
sheet, document author, sheet protection and passwords, overlapping
merges or hyperlinks trapped inside merges (Excel calls these corrupt),
unresolved template tokens, `#REF!`, unbalanced brackets or quotes, empty
arguments, unknown or too-modern functions, formulas pointing at a missing
sheet, defined names and drop down sources that do not exist, column
letters left without a row number, and formula cells that are not locked.

`check_social_sample.py` re-reads the finished file and confirms the
example data joins up — that every content ID on Performance, the five
platform planners and the repurposing tab points at a post that exists,
that every pillar, caption, hashtag group, campaign, persona, owner and
account name matches its source tab, and that there are enough published
posts and monthly analytics rows for the dashboards to show something.

`recalc_check.py` goes further and actually *runs* the workbook: it loads
it into the `formulas` calculation engine, evaluates the whole dependency
graph, then reads the answer out of every formula cell — so a formula that
parses perfectly but divides by zero on an empty row is still caught.
Check a `--rows 60` build rather than a full-size one: the formulas are
identical, there are just fewer rows of them, and 110,000 of them need
far more memory than a laptop will give you (the tool says so and stops
unless you pass `--force`). Add `--show "Dashboard!A10,Reports!D11"` to
print the calculated value of particular cells.

Current status: all three clean and `pyflakes` clean. 14,349 formulas in
the `--rows 60` build were calculated and every one lands on a real
answer.

---

## 5. What a spreadsheet cannot do

Worth saying plainly, in the listing as well as in the Help tab:

* it **cannot fetch analytics automatically** — no spreadsheet can; that
  needs a platform API. You copy about fifteen numbers per account once a
  month, which takes around ten minutes;
* it **cannot publish or schedule posts** — pair it with a scheduler and
  keep this as the plan and the record;
* it cannot see comments, messages or mentions;
* revenue attribution is only as good as the numbers you give it — use UTM
  links and your shop's own reports;
* the Compliance tab is a prompt list, not legal advice.

What it does give you is a year of content planned in one place, every
caption and hashtag reusable, and a history of what worked that no
scheduling app hands back to you.

---

## 6. What is deliberately *not* in version 1

The brief listed 890 features across 41 categories; v1 ships the ones a
working creator or small brand uses every week. These are the natural
premium / v2 additions:

* API or Zapier style integrations that pull platform analytics in;
* an automation tab that writes scheduling CSV exports for Later, Buffer
  or Metricool;
* multi-brand or agency mode with a client switcher on every tab;
* statistical significance maths on the A/B tests;
* influencer contract templates and payment schedules;
* a Google Sheets edition with Apps Script automation.

---

© Novality Store. Generated workbooks are licensed for personal use and for
a single business.
