# Social Media Content Calendar & Analytics Planner — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 43 tab `.xlsx` workbook with over 110,000 formulas: a 12 month content calendar, an idea bank, caption and hashtag banks, five platform planners, campaigns, launches, influencer and UGC tracking, repurposing, competitor and trend research, monthly analytics, per post performance, two dashboards, printable briefs and a reports page. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_social_media_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --brand TEXT          brand name on the cover and the dashboards
  --currency SYMBOL     currency symbol  (default: $)
  --year YYYY           the year the calendar is built for
  --rows N              rows of capacity in the big registers
  --today YYYY-MM-DD    pretend it is this date (for the demo data)
```

---

## Contents

* [1. Run it](#1-run-it)
  * [`build_social_media_workbook.py`](#build_social_media_workbookpy)
* [2. The package](#2-the-package)
  * [`social_media_spreadsheet/__init__.py`](#social_media_spreadsheet__init__py)
  * [`social_media_spreadsheet/__main__.py`](#social_media_spreadsheet__main__py)
  * [`social_media_spreadsheet/cli.py`](#social_media_spreadsheetclipy)
  * [`social_media_spreadsheet/config.py`](#social_media_spreadsheetconfigpy)
  * [`social_media_spreadsheet/layout.py`](#social_media_spreadsheetlayoutpy)
  * [`social_media_spreadsheet/formats.py`](#social_media_spreadsheetformatspy)
  * [`social_media_spreadsheet/writer.py`](#social_media_spreadsheetwriterpy)
  * [`social_media_spreadsheet/names.py`](#social_media_spreadsheetnamespy)
  * [`social_media_spreadsheet/build.py`](#social_media_spreadsheetbuildpy)
  * [`social_media_spreadsheet/sample_data.py`](#social_media_spreadsheetsample_datapy)
* [3. Setup & planning registers](#3-setup-planning-registers)
  * [`social_media_spreadsheet/sheets/__init__.py`](#social_media_spreadsheetsheets__init__py)
  * [`social_media_spreadsheet/sheets/common.py`](#social_media_spreadsheetsheetscommonpy)
  * [`social_media_spreadsheet/sheets/lists.py`](#social_media_spreadsheetsheetslistspy)
  * [`social_media_spreadsheet/sheets/settings.py`](#social_media_spreadsheetsheetssettingspy)
  * [`social_media_spreadsheet/sheets/accounts.py`](#social_media_spreadsheetsheetsaccountspy)
  * [`social_media_spreadsheet/sheets/personas.py`](#social_media_spreadsheetsheetspersonaspy)
  * [`social_media_spreadsheet/sheets/pillars.py`](#social_media_spreadsheetsheetspillarspy)
  * [`social_media_spreadsheet/sheets/calendar_sheet.py`](#social_media_spreadsheetsheetscalendar_sheetpy)
  * [`social_media_spreadsheet/sheets/ideas.py`](#social_media_spreadsheetsheetsideaspy)
  * [`social_media_spreadsheet/sheets/captions.py`](#social_media_spreadsheetsheetscaptionspy)
  * [`social_media_spreadsheet/sheets/hashtags.py`](#social_media_spreadsheetsheetshashtagspy)
  * [`social_media_spreadsheet/sheets/platforms.py`](#social_media_spreadsheetsheetsplatformspy)
* [4. Campaign & production registers](#4-campaign-production-registers)
  * [`social_media_spreadsheet/sheets/campaigns.py`](#social_media_spreadsheetsheetscampaignspy)
  * [`social_media_spreadsheet/sheets/launches.py`](#social_media_spreadsheetsheetslaunchespy)
  * [`social_media_spreadsheet/sheets/seasonal.py`](#social_media_spreadsheetsheetsseasonalpy)
  * [`social_media_spreadsheet/sheets/influencers.py`](#social_media_spreadsheetsheetsinfluencerspy)
  * [`social_media_spreadsheet/sheets/ugc.py`](#social_media_spreadsheetsheetsugcpy)
  * [`social_media_spreadsheet/sheets/repurposing.py`](#social_media_spreadsheetsheetsrepurposingpy)
  * [`social_media_spreadsheet/sheets/competitors.py`](#social_media_spreadsheetsheetscompetitorspy)
  * [`social_media_spreadsheet/sheets/trends.py`](#social_media_spreadsheetsheetstrendspy)
  * [`social_media_spreadsheet/sheets/assets.py`](#social_media_spreadsheetsheetsassetspy)
  * [`social_media_spreadsheet/sheets/team.py`](#social_media_spreadsheetsheetsteampy)
  * [`social_media_spreadsheet/sheets/tasks.py`](#social_media_spreadsheetsheetstaskspy)
  * [`social_media_spreadsheet/sheets/compliance.py`](#social_media_spreadsheetsheetscompliancepy)
* [5. Measurement registers](#5-measurement-registers)
  * [`social_media_spreadsheet/sheets/analytics.py`](#social_media_spreadsheetsheetsanalyticspy)
  * [`social_media_spreadsheet/sheets/performance.py`](#social_media_spreadsheetsheetsperformancepy)
  * [`social_media_spreadsheet/sheets/ab_tests.py`](#social_media_spreadsheetsheetsab_testspy)
  * [`social_media_spreadsheet/sheets/ads.py`](#social_media_spreadsheetsheetsadspy)
  * [`social_media_spreadsheet/sheets/goals.py`](#social_media_spreadsheetsheetsgoalspy)
  * [`social_media_spreadsheet/sheets/budget.py`](#social_media_spreadsheetsheetsbudgetpy)
* [6. The pages](#6-the-pages)
  * [`social_media_spreadsheet/sheets/calc.py`](#social_media_spreadsheetsheetscalcpy)
  * [`social_media_spreadsheet/sheets/dashboard.py`](#social_media_spreadsheetsheetsdashboardpy)
  * [`social_media_spreadsheet/sheets/annual.py`](#social_media_spreadsheetsheetsannualpy)
  * [`social_media_spreadsheet/sheets/month_view.py`](#social_media_spreadsheetsheetsmonth_viewpy)
  * [`social_media_spreadsheet/sheets/workflow.py`](#social_media_spreadsheetsheetsworkflowpy)
  * [`social_media_spreadsheet/sheets/brand_voice.py`](#social_media_spreadsheetsheetsbrand_voicepy)
  * [`social_media_spreadsheet/sheets/reports.py`](#social_media_spreadsheetsheetsreportspy)
  * [`social_media_spreadsheet/sheets/start_here.py`](#social_media_spreadsheetsheetsstart_herepy)
  * [`social_media_spreadsheet/sheets/printables.py`](#social_media_spreadsheetsheetsprintablespy)
  * [`social_media_spreadsheet/sheets/help_tab.py`](#social_media_spreadsheetsheetshelp_tabpy)
* [7. Quality checks (optional)](#7-quality-checks-optional)
  * [`tools/validate_social_workbook.py`](#toolsvalidate_social_workbookpy)
  * [`tools/check_social_sample.py`](#toolscheck_social_samplepy)
  * [`tools/recalc_check.py`](#toolsrecalc_checkpy)
  * [`tools/export_source_md.py`](#toolsexport_source_mdpy)
  * [`tools/make_download_zip.py`](#toolsmake_download_zippy)

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_social_media_workbook.py
```

The workbook lands in `output/Social Media Content Planner.xlsx`.

### `build_social_media_workbook.py`

Build the Social Media Content Calendar & Analytics Planner. *(32 lines)*

````python
#!/usr/bin/env python3
"""
Build the Social Media Content Calendar & Analytics Planner.

Double click it, or run::

    python build_social_media_workbook.py

The finished .xlsx lands in the ``output`` folder next to this script.
Author of the generated workbook: Novality Store.

Options (all optional)::

    python build_social_media_workbook.py --blank        # no example data
    python build_social_media_workbook.py --currency "£" --currency-code GBP
    python build_social_media_workbook.py --year 2027
    python build_social_media_workbook.py -o "~/Desktop/Planner.xlsx"

See ``python build_social_media_workbook.py --help`` for the full list.
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from social_media_spreadsheet.cli import main   # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer, the cross sheet names and the demo data. These are the parts every tab leans on.

### `social_media_spreadsheet/__init__.py`

Social Media Content Calendar & Analytics Planner ================================================= Generates a ready-to-sell Excel workbook for social media managers, creators and small brands: a 12 month content calendar, idea / caption / hashtag banks, a production workflow, campaigns, launches, influencer and UGC trackers, repurposing, competitor and trend research, per-platform planners for Instagram, TikTok, Pinterest, YouTube and Facebook, monthly analytics, post level performance scoring, A/B tests, ads, assets, goals, budget and ROI, team tasks, dashboards, reports and printable briefs. *(32 lines)*

````python
"""
Social Media Content Calendar & Analytics Planner
=================================================

Generates a ready-to-sell Excel workbook for social media managers, creators
and small brands: a 12 month content calendar, idea / caption / hashtag
banks, a production workflow, campaigns, launches, influencer and UGC
trackers, repurposing, competitor and trend research, per-platform planners
for Instagram, TikTok, Pinterest, YouTube and Facebook, monthly analytics,
post level performance scoring, A/B tests, ads, assets, goals, budget and
ROI, team tasks, dashboards, reports and printable briefs.

    python -m social_media_spreadsheet --output "My Planner.xlsx"

or from Python::

    from social_media_spreadsheet import build_workbook, Config
    build_workbook("planner.xlsx", Config(include_sample_data=False))

Every formula cell is locked and every sheet is protected with the password
in :class:`~social_media_spreadsheet.config.Config` (default ``premium``), so
the maths cannot be edited by accident while the input cells stay open.
"""

from .build import build_workbook
from .config import AUTHOR, FORMULA_PASSWORD, Config

__all__ = ["build_workbook", "Config", "AUTHOR", "FORMULA_PASSWORD",
           "__version__"]

__version__ = "1.0.0"
__author__ = AUTHOR
````

### `social_media_spreadsheet/__main__.py`

``python -m social_media_spreadsheet``. *(6 lines)*

````python
"""``python -m social_media_spreadsheet``."""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
````

### `social_media_spreadsheet/cli.py`

Command line front end. *(117 lines)*

````python
"""Command line front end.

    python -m social_media_spreadsheet
    python -m social_media_spreadsheet -o "output/Planner.xlsx" --blank
    python -m social_media_spreadsheet --currency "£" --year 2027 --rows 300
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
    "output", "Social Media Content Planner.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="social_media_spreadsheet",
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
    parser.add_argument("--brand", default=None,
                        help="pre-fill the brand name on Business Setup")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="planning year the calendar opens on")
    parser.add_argument("--week-start", choices=["Monday", "Sunday"],
                        default=None, help="first day of the week")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (default row counts are "
                             "multiplied to fit this many calendar rows)")
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
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.brand:
        cfg.brand.name = args.brand
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
    print(f"  sheets        : {len(SHEET_ORDER)} "
          f"(1 hidden calculation engine)")
    print(f"  planning year : {cfg.year}")
    print(f"  example data  : {'no' if not cfg.include_sample_data else 'yes'}")
    print(f"  protection    : "
          f"{'off' if not cfg.protect_sheets else 'on, password ' + cfg.password}")
    print(f"  written to    : {os.path.abspath(path)}  ({size:,.0f} KB)")
    return 0


if __name__ == "__main__":        # pragma: no cover
    raise SystemExit(main())
````

### `social_media_spreadsheet/config.py`

Configuration for the Social Media Content Calendar & Analytics Planner. *(221 lines)*

````python
"""
Configuration for the Social Media Content Calendar & Analytics Planner.

Everything a shop owner may want to re-brand lives here, so the workbook can
be regenerated for a different brand, currency or country without touching
any of the sheet code.

Author of the generated workbooks: Novality Store.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import date


# ---------------------------------------------------------------------------
# Product / branding constants
# ---------------------------------------------------------------------------

AUTHOR = "Novality Store"
PRODUCT_NAME = "Social Media Content Calendar & Analytics Planner"
PRODUCT_SUBTITLE = (
    "12-Month Content Calendar, Caption & Hashtag Banks, Campaigns and "
    "Instagram, TikTok, Pinterest, YouTube & Facebook Analytics"
)
PRODUCT_VERSION = "1.0.0"

#: Title suggested for the shop listing.
PRODUCT_LISTING_TITLE = (
    "Social Media Content Planner Spreadsheet, 12-Month Content Calendar, "
    "Instagram, TikTok, Pinterest, YouTube and Facebook Analytics Tracker"
)

#: Password that protects every sheet so formula cells cannot be edited by
#: accident.  Input cells stay unlocked, formula cells are locked.
FORMULA_PASSWORD = "premium"

#: The platforms the workbook reports on out of the box.  Change these on the
#: Lists tab and every drop down follows; the five analytics columns on the
#: dashboards follow :data:`TRACKED_PLATFORMS`.
TRACKED_PLATFORMS = ["Instagram", "TikTok", "Pinterest", "YouTube",
                     "Facebook"]


# ---------------------------------------------------------------------------
# Colour palette (indigo / coral scheme used across every tab)
# ---------------------------------------------------------------------------

class Palette:
    ink = "#241B38"          # darkest indigo - titles
    primary = "#3B2A6B"      # headers
    primary_mid = "#6350A8"  # section bands
    primary_soft = "#E7E3F6"  # light band / alternating rows
    accent = "#F2665E"       # coral accent
    accent_soft = "#FDEBE9"
    input_bg = "#FFFCF7"     # editable cell background (warm white)
    input_border = "#E8DCC8"
    calc_bg = "#F1EFF8"      # locked formula cell background (soft lilac)
    calc_border = "#D8D2EA"
    white = "#FFFFFF"
    grey_text = "#5F5A70"
    line = "#C7C1DA"
    good_bg = "#D9F2E4"
    good_text = "#0E6B45"
    warn_bg = "#FFF1CF"
    warn_text = "#8A5A00"
    bad_bg = "#FBD9DD"
    bad_text = "#9C1C2E"
    info_bg = "#DEE8FA"
    info_text = "#2A4A8B"
    note_bg = "#F8F7FC"


# ---------------------------------------------------------------------------
# Row capacity of every register (how many pre-formatted / pre-formulated rows)
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Settings": 0,
    "Lists": 40,
    "Accounts": 30,
    "Audience Personas": 12,
    "Content Pillars": 20,
    "Content Calendar": 600,
    "Idea Bank": 300,
    "Caption Bank": 200,
    "Hashtag Bank": 300,
    "Instagram": 200,
    "TikTok": 200,
    "Pinterest": 200,
    "YouTube": 120,
    "Facebook": 200,
    "Campaigns": 60,
    "Launch Planner": 40,
    "Seasonal Planner": 60,
    "Influencers": 60,
    "UGC Tracker": 100,
    "Repurposing": 150,
    "Competitors": 40,
    "Trends": 80,
    "Analytics": 150,
    "Performance": 400,
    "AB Tests": 40,
    "Ads": 100,
    "Assets": 250,
    "Goals": 40,
    "Budget": 60,
    "Team": 30,
    "Tasks": 60,
}


@dataclass
class BrandDefaults:
    """Sample values written into the Business Setup (Settings) sheet."""

    name: str = "Wildflower Studio"
    owner: str = "Jordan Avery"
    email: str = "hello@wildflowerstudio.com"
    phone: str = "+1 (555) 017 4410"
    website: str = "www.wildflowerstudio.com"
    shop: str = "shop.wildflowerstudio.com"
    industry: str = "Handmade home & lifestyle"
    niche: str = "Sustainable ceramics and home styling"
    audience: str = ("Women 28-45 who love slow living, handmade homeware "
                     "and styling their first home")
    goal: str = "Grow to 25,000 followers and 1,200 shop visits a month"
    logo_link: str = "brand-assets/wildflower-logo.png"
    palette: str = "#3B2A6B indigo, #F2665E coral, #F4EFE6 oat"
    fonts: str = "Headings: Canela. Body: Inter."
    timezone: str = "GMT-05:00 Eastern Time"
    hashtag: str = "#wildflowerstudio"


@dataclass
class Config:
    """All generator options."""

    output_path: str = "output/Social Media Content Planner.xlsx"

    # Branding / document properties
    author: str = AUTHOR
    company: str = AUTHOR
    product_name: str = PRODUCT_NAME
    product_subtitle: str = PRODUCT_SUBTITLE
    version: str = PRODUCT_VERSION
    release_date: str = "2026-10-11"

    # Protection
    password: str = FORMULA_PASSWORD
    protect_sheets: bool = True

    # Formats
    currency_symbol: str = "$"
    currency_code: str = "USD"
    date_format: str = "dd-mmm-yyyy"
    time_format: str = "h:mm AM/PM"

    # Brand / planning defaults pre-filled on the Settings tab
    brand: BrandDefaults = field(default_factory=BrandDefaults)
    plan_year: int = 0               # 0 = the year of `today`
    week_start: str = "Monday"
    posts_per_week_target: int = 3
    monthly_post_target: int = 13
    engagement_target: float = 0.045
    follower_goal: int = 25000
    reach_goal: int = 120000
    click_goal: int = 1800
    lead_goal: int = 240
    conversion_goal: int = 90
    revenue_goal: float = 36000.0
    monthly_budget: float = 900.0
    annual_budget: float = 10800.0
    roas_target: float = 4.0
    cost_per_lead_target: float = 6.0
    production_days: int = 5          # draft deadline before publishing
    approval_days: int = 2            # approval deadline before publishing
    overdue_grace_days: int = 0
    idea_review_days: int = 14
    trend_window_days: int = 14
    asset_review_days: int = 180
    inactive_account_days: int = 30
    top_rows: int = 5                 # size of the "top content" tables

    # Content switches
    include_sample_data: bool = True
    rows: dict = field(default_factory=lambda: dict(DEFAULT_ROWS))
    today: date = field(default_factory=date.today)

    # ---------------------------------------------------------------- helpers
    def row_count(self, sheet: str) -> int:
        return int(self.rows.get(sheet, 100))

    @property
    def year(self) -> int:
        return self.plan_year or self.today.year

    def scale_rows(self, calendar_rows: int) -> None:
        """Scale every register in place so the calendar holds N rows."""
        base = DEFAULT_ROWS["Content Calendar"]
        factor = max(0.02, float(calendar_rows) / float(base))
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
        return f'"{s}"#,##0.00;[Red]-"{s}"#,##0.00'

    @property
    def money_format_0(self) -> str:
        s = self.currency_symbol
        return f'"{s}"#,##0;[Red]-"{s}"#,##0'
````

### `social_media_spreadsheet/layout.py`

Layout engine. *(244 lines)*

````python
"""
Layout engine.

The workbook is described declaratively: every register sheet (Content Calendar,
Caption Bank, Analytics, ...) is a :class:`SheetSpec` made of :class:`Col`
objects.  Formulas
are written as templates that never contain hard coded column letters, so
columns can be reordered or inserted without breaking a single formula:

    "=IF({title}{r}=\\"\\",\\"\\",SUMIFS([[Performance.reach]],[[Performance.content]],{id}{r}))"

    {r}            -> the current Excel row number (1 based)
    {key}          -> column letter of the column `key` on the *same* sheet
    [[Sheet.key]]  -> absolute range of that column's data rows, e.g.
                      Performance!$K$5:$K$304
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

### `social_media_spreadsheet/formats.py`

Cell format factory. *(200 lines)*

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

### `social_media_spreadsheet/writer.py`

Sheet writing helpers shared by every tab of the workbook. *(444 lines)*

````python
"""
Sheet writing helpers shared by every tab of the workbook.

:class:`SheetWriter` wraps an XlsxWriter worksheet and adds the house style:
title band, navigation strip, section headers, KPI tiles, label/value settings
rows and - most importantly - :meth:`SheetWriter.grid`, which turns a
:class:`~social_media_spreadsheet.layout.SheetSpec` into a finished register
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
    ("Calendar", "Content Calendar"),
    ("Ideas", "Idea Bank"),
    ("Captions", "Caption Bank"),
    ("Hashtags", "Hashtag Bank"),
    ("Analytics", "Analytics"),
    ("Reports", "Reports"),
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

### `social_media_spreadsheet/names.py`

Workbook level defined names that point at the register sheets. *(46 lines)*

````python
"""Workbook level defined names that point at the register sheets.

The drop down lists on the `Lists` tab define their own names; these are the
ones that have to follow live data (content IDs, caption IDs, campaign names,
...), so they are built from the registry once every register has been
created.  Each one is an ``OFFSET`` range that grows and shrinks with the
number of filled rows.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    "AccountNames": ("Accounts", "account"),
    "PersonaNames": ("Audience Personas", "name"),
    "PillarNames": ("Content Pillars", "name"),
    "ContentIDs": ("Content Calendar", "id"),
    "ContentTitles": ("Content Calendar", "title"),
    "CaptionIDs": ("Caption Bank", "id"),
    "CaptionTitles": ("Caption Bank", "title"),
    "IdeaIDs": ("Idea Bank", "id"),
    "CampaignNames": ("Campaigns", "name"),
    "LaunchNames": ("Launch Planner", "product"),
    "InfluencerNames": ("Influencers", "name"),
    "CompetitorNames": ("Competitors", "name"),
    "TrendNames": ("Trends", "name"),
    "AssetIDs": ("Assets", "id"),
    "GoalNames": ("Goals", "name"),
    "TeamNames": ("Team", "name"),
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

### `social_media_spreadsheet/build.py`

Workbook builder: puts every tab together in the right order. *(160 lines)*

````python
"""Workbook builder: puts every tab together in the right order.

Typical use::

    from social_media_spreadsheet import build_workbook
    build_workbook("Social Media Content Planner.xlsx")
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
from .sheets import (ab_tests, accounts, ads, analytics, annual, assets,
                     brand_voice, budget, calc, calendar_sheet, campaigns,
                     captions, competitors, compliance, dashboard, goals,
                     hashtags, help_tab, ideas, influencers, launches,
                     lists, month_view, performance, personas, pillars,
                     platforms, printables, reports, repurposing, seasonal,
                     settings, start_here, tasks, team, trends, ugc,
                     workflow)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    accounts, personas, pillars, calendar_sheet, ideas, captions, hashtags,
    platforms.INSTAGRAM, platforms.TIKTOK, platforms.PINTEREST,
    platforms.YOUTUBE, platforms.FACEBOOK, campaigns, launches, seasonal,
    influencers, ugc, repurposing, competitors, trends, analytics,
    performance, ab_tests, ads, assets, goals, budget, team, tasks,
    compliance, calc,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Start Here", "Dashboard", "Annual Dashboard", "Settings", "Lists",
    "Brand Voice", "Accounts", "Audience Personas", "Content Pillars",
    "Content Calendar", "Month View", "Workflow",
    "Idea Bank", "Caption Bank", "Hashtag Bank",
    "Instagram", "TikTok", "Pinterest", "YouTube", "Facebook",
    "Campaigns", "Launch Planner", "Seasonal Planner",
    "Influencers", "UGC Tracker", "Repurposing", "Competitors", "Trends",
    "Analytics", "Performance", "AB Tests", "Ads", "Assets",
    "Goals", "Budget", "Team", "Tasks", "Compliance",
    "Reports", "Print Brief", "Print Week", "Help",
    "Calc",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Start Here": start_here.build,
    "Dashboard": dashboard.build,
    "Annual Dashboard": annual.build,
    "Month View": month_view.build,
    "Workflow": workflow.build,
    "Settings": settings.build,
    "Lists": lists.build,
    "Brand Voice": brand_voice.build,
    "Reports": reports.build,
    "Print Brief": printables.build_brief,
    "Print Week": printables.build_week,
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


def build_workbook(path: str = "Social Media Content Planner.xlsx",
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

    # 3. Live lookup names (content IDs, campaign names, team members).
    define_cross_sheet_names(book, reg)

    # 4. Final touches: hide the engine, open on the first page.
    if calc.NAME in sheets:
        sheets[calc.NAME].ws.hide()
    first = sheets[SHEET_ORDER[0]].ws
    first.activate()
    first.set_first_sheet()

    book.close()
    return path


def _set_properties(book, cfg: Config) -> None:
    book.set_properties({
        "title": cfg.product_name,
        "subject": "Social media content planning and analytics",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("social media, content calendar, content planner, "
                     "instagram, tiktok, pinterest, youtube, facebook, "
                     "analytics, engagement, hashtags, spreadsheet "
                     "template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked and "
                     f"every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

### `social_media_spreadsheet/sample_data.py`

Demo data. *(1178 lines)*

````python
"""
Demo data.

A small handmade-ceramics brand with five social accounts, ten weeks of
content, captions, hashtags, campaigns, influencers, analytics and post
level performance.  Everything is written into the input (cream) cells so
the dashboards, reports and alerts show meaningful numbers the moment the
workbook is opened, and everything is dated relative to the build date.

Switch it all off with ``--blank``.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .config import Config


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def month_key(day: _dt.date) -> str:
    return day.strftime("%Y-%m")


@dataclass
class Sample:
    accounts: List[Dict[str, Any]] = field(default_factory=list)
    personas: List[Dict[str, Any]] = field(default_factory=list)
    pillars: List[Dict[str, Any]] = field(default_factory=list)
    calendar: List[Dict[str, Any]] = field(default_factory=list)
    ideas: List[Dict[str, Any]] = field(default_factory=list)
    captions: List[Dict[str, Any]] = field(default_factory=list)
    hashtags: List[Dict[str, Any]] = field(default_factory=list)
    instagram: List[Dict[str, Any]] = field(default_factory=list)
    tiktok: List[Dict[str, Any]] = field(default_factory=list)
    pinterest: List[Dict[str, Any]] = field(default_factory=list)
    youtube: List[Dict[str, Any]] = field(default_factory=list)
    facebook: List[Dict[str, Any]] = field(default_factory=list)
    campaigns: List[Dict[str, Any]] = field(default_factory=list)
    launches: List[Dict[str, Any]] = field(default_factory=list)
    seasonal: List[Dict[str, Any]] = field(default_factory=list)
    influencers: List[Dict[str, Any]] = field(default_factory=list)
    ugc: List[Dict[str, Any]] = field(default_factory=list)
    repurposing: List[Dict[str, Any]] = field(default_factory=list)
    competitors: List[Dict[str, Any]] = field(default_factory=list)
    trends: List[Dict[str, Any]] = field(default_factory=list)
    analytics: List[Dict[str, Any]] = field(default_factory=list)
    performance: List[Dict[str, Any]] = field(default_factory=list)
    tests: List[Dict[str, Any]] = field(default_factory=list)
    ads: List[Dict[str, Any]] = field(default_factory=list)
    assets: List[Dict[str, Any]] = field(default_factory=list)
    goals: List[Dict[str, Any]] = field(default_factory=list)
    budget: List[Dict[str, Any]] = field(default_factory=list)
    team: List[Dict[str, Any]] = field(default_factory=list)
    tasks: List[Dict[str, Any]] = field(default_factory=list)


_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (once) the demo dataset for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg)
    return _CACHE[key]


# ---------------------------------------------------------------------------
# The content plan.  (days from today, platform, format, pillar, title)
# Negative days are already published, positive days are still to come.
# ---------------------------------------------------------------------------

PLAN = [
    (-56, "Instagram", "Carousel", "Educate",
     "5 ways to style a single stem vase"),
    (-55, "Pinterest", "Standard Pin", "Educate",
     "Autumn table styling checklist"),
    (-54, "TikTok", "Reel", "Behind the scenes",
     "Throwing a mug from start to finish"),
    (-52, "Instagram", "Reel", "Inspire",
     "Morning light in the studio"),
    (-50, "Facebook", "Feed post", "Community",
     "Meet the makers behind your mugs"),
    (-49, "Instagram", "Story", "Promote",
     "Last 12 speckled bowls of the batch"),
    (-48, "YouTube", "Long video", "Educate",
     "How a glaze is mixed - full studio tour"),
    (-47, "Pinterest", "Standard Pin", "Inspire",
     "Neutral shelf styling in 6 steps"),
    (-45, "TikTok", "Short video", "Behind the scenes",
     "Kiln opening - did they survive?"),
    (-44, "Instagram", "Feed post", "Community",
     "Your kitchen shelfies - UGC round up"),
    (-42, "Instagram", "Carousel", "Educate",
     "Stoneware vs porcelain - which to buy"),
    (-41, "Facebook", "Photo", "Promote",
     "Autumn collection now in the shop"),
    (-40, "Pinterest", "Video Pin", "Educate",
     "Three ways to fold a linen napkin"),
    (-38, "TikTok", "Reel", "Entertain",
     "Pottery fails you asked to see"),
    (-37, "Instagram", "Reel", "Promote",
     "Unboxing the autumn table set"),
    (-35, "YouTube", "Short video", "Educate",
     "Why handmade mugs keep tea hotter"),
    (-34, "Instagram", "Story", "Community",
     "This or that - glaze colour vote"),
    (-33, "Pinterest", "Standard Pin", "Promote",
     "The autumn table set - shop the look"),
    (-31, "Instagram", "Carousel", "Inspire",
     "A slow Sunday table, styled three ways"),
    (-30, "Facebook", "Feed post", "Educate",
     "Caring for handmade ceramics"),
    (-28, "TikTok", "Reel", "Behind the scenes",
     "Packing orders at 6am"),
    (-27, "Instagram", "Feed post", "Personal",
     "Why I left my desk job for clay"),
    (-25, "Pinterest", "Standard Pin", "Educate",
     "Small kitchen storage ideas"),
    (-24, "Instagram", "Reel", "Educate",
     "The 3 mugs every home needs"),
    (-23, "YouTube", "Long video", "Behind the scenes",
     "A week in a one woman ceramics studio"),
    (-21, "Instagram", "Carousel", "Community",
     "Answering your 9 most asked questions"),
    (-20, "TikTok", "Short video", "Inspire",
     "Satisfying trimming sounds"),
    (-19, "Facebook", "Link post", "Promote",
     "Restock alert - speckled dinnerware"),
    (-17, "Instagram", "Reel", "Promote",
     "Restock in 60 seconds"),
    (-16, "Pinterest", "Video Pin", "Inspire",
     "Styling a console table for autumn"),
    (-14, "Instagram", "Story", "Promote",
     "Studio sale - 20% off seconds"),
    (-13, "TikTok", "Reel", "Educate",
     "What 'seconds' means in pottery"),
    (-12, "Instagram", "Carousel", "Educate",
     "How to photograph your table setting"),
    (-10, "YouTube", "Short video", "Entertain",
     "Clay or chocolate? guess the sound"),
    (-9, "Instagram", "Feed post", "Community",
     "Customer story - Hannah's wedding set"),
    (-7, "Instagram", "Reel", "Inspire",
     "Autumn table reveal"),
    (-6, "Pinterest", "Standard Pin", "Educate",
     "Gift guide - presents for home lovers"),
    (-5, "TikTok", "Reel", "Promote",
     "Gift sets are live"),
    (-4, "Facebook", "Feed post", "Community",
     "Local market this weekend"),
    (-3, "Instagram", "Carousel", "Promote",
     "The gift guide - 8 ideas under 60"),
    (-2, "Instagram", "Story", "Community",
     "Ask me anything about ordering"),
    (-1, "TikTok", "Short video", "Behind the scenes",
     "Wrapping a gift order"),
    (0, "Instagram", "Reel", "Educate",
     "Three table settings, one dinner set"),
    (1, "Pinterest", "Standard Pin", "Promote",
     "Holiday gift guide pin"),
    (2, "Instagram", "Carousel", "Inspire",
     "Colour palette: oat, clay and sage"),
    (3, "YouTube", "Long video", "Educate",
     "Beginner's guide to buying handmade"),
    (4, "TikTok", "Reel", "Entertain",
     "Guess the glaze - part two"),
    (5, "Instagram", "Feed post", "Promote",
     "Holiday shipping dates"),
    (6, "Facebook", "Feed post", "Promote",
     "Holiday gift guide is live"),
    (8, "Instagram", "Reel", "Behind the scenes",
     "Glazing the holiday collection"),
    (9, "Pinterest", "Video Pin", "Educate",
     "How to set a holiday table"),
    (11, "Instagram", "Carousel", "Community",
     "Your holiday tables - call for photos"),
    (13, "TikTok", "Reel", "Promote",
     "Last posting day countdown"),
    (15, "Instagram", "Story", "Promote",
     "Final orders before the break"),
]

#: campaign for a plan row, by index range
CAMPAIGN_FOR = [
    (0, 17, "Autumn Table Styling"),
    (18, 29, "Studio Restock"),
    (30, 54, "Holiday Gift Guide"),
]

OBJECTIVES = {
    "Educate": "Brand authority",
    "Inspire": "Awareness",
    "Behind the scenes": "Community",
    "Community": "Engagement",
    "Promote": "Sales",
    "Entertain": "Awareness",
    "Personal": "Community",
}

CTAS = {
    "Educate": "Save this",
    "Inspire": "Follow for more",
    "Behind the scenes": "Comment below",
    "Community": "Tag a friend",
    "Promote": "Shop now",
    "Entertain": "Share with a friend",
    "Personal": "Reply to this",
}

FUNNEL = {
    "Educate": "Consideration",
    "Inspire": "Awareness",
    "Behind the scenes": "Awareness",
    "Community": "Retention",
    "Promote": "Conversion",
    "Entertain": "Awareness",
    "Personal": "Awareness",
}

SLOTS = {
    "Instagram": "Morning",
    "TikTok": "Evening",
    "Pinterest": "Lunchtime",
    "YouTube": "Afternoon",
    "Facebook": "Morning",
}

OWNERS = ["Jordan Avery", "Priya Nair", "Sam Okonkwo", "Elena Dias"]


# ---------------------------------------------------------------------------

def _build(cfg: Config) -> Sample:
    rnd = random.Random(20261011)
    today = cfg.today
    s = Sample()
    brand = cfg.brand.name

    # ------------------------------------------------------------- accounts
    account_plan = [
        ("Instagram", "Wildflower Studio", "@wildflower.studio",
         "instagram.com/wildflower.studio", "Business account", 8420,
         "Main shop window - product, styling and community", 5),
        ("TikTok", "Wildflower Studio", "@wildflowerstudio",
         "tiktok.com/@wildflowerstudio", "Creator account", 5260,
         "Reach new people with process and personality", 4),
        ("Pinterest", "Wildflower Studio", "wildflowerstudio",
         "pinterest.com/wildflowerstudio", "Business account", 3180,
         "Evergreen search traffic to the shop", 3),
        ("YouTube", "Wildflower Studio", "@wildflowerstudio",
         "youtube.com/@wildflowerstudio", "Channel", 1240,
         "Long form studio films and buying guides", 1),
        ("Facebook", "Wildflower Studio", "wildflowerstudiohome",
         "facebook.com/wildflowerstudiohome", "Page", 2640,
         "Local customers, markets and events", 2),
    ]
    for i, (platform, name, handle, url, kind, followers, purpose,
            per_week) in enumerate(account_plan):
        s.accounts.append(dict(
            platform=platform, account=name, handle=handle, url=url,
            type=kind, owner=OWNERS[i % len(OWNERS)], purpose=purpose,
            audience="Home lovers 28-45" if i != 3 else "Makers and buyers",
            followers=followers, launched=d(today, -720 + i * 90),
            target_per_week=per_week,
            bio_ok="Yes", photo_ok="Yes", cover_ok="Yes" if i != 1 else "No",
            link_ok="Yes", contact_ok="Yes" if i != 3 else "No",
            branding_ok="Yes", access="Owner + manager, 2FA on",
            notes="Link in bio points at the shop" if i == 0 else ""))

    # ------------------------------------------------------------- personas
    persona_plan = [
        ("First-home Freya", "28-34", "City flats, UK & US",
         "Slow living, thrifting, cooking for friends",
         "Her flat feels unfinished and everything matches too much",
         "A home that feels personal without spending a fortune",
         "Wants pieces with a story", "Worries handmade means fragile",
         "Instagram", "Carousel", "Awareness"),
        ("Gifting Grace", "35-45", "Suburban US & Canada",
         "Hosting, gift giving, seasonal decorating",
         "Finding a present that is not generic, close to the deadline",
         "To be the person who gives the thoughtful gift",
         "Buys for others, values packaging",
         "Will it arrive in time?", "Pinterest", "Standard Pin",
         "Conversion"),
        ("Interiors Isla", "30-42", "Design-led cities",
         "Interiors, ceramics, neutral palettes, slow design",
         "Mass produced homeware feels soulless",
         "A considered home she is proud to photograph",
         "Follows makers, buys a piece at a time",
         "Is the colour true to the photo?", "Instagram", "Reel",
         "Consideration"),
    ]
    for (name, age, place, interests, pain, goal, motivation, objection,
         platform, fmt, stage) in persona_plan:
        s.personas.append(dict(
            name=name, age=age, location=place, interests=interests,
            pain=pain, goal=goal, motivation=motivation,
            objection=objection, platform=platform, format=fmt,
            stage=stage,
            engagement="Replies to stories, saves carousels",
            faq="How long does delivery take?",
            language="'shelfie', 'tablescape', 'slow made'",
            awareness_idea="Styling inspiration reels",
            consideration_idea="Material and care explainers",
            conversion_idea="Shop the look carousels",
            retention_idea="Customer features and care tips",
            notes="Highest spend per order" if name == "Gifting Grace"
                  else ""))

    # ------------------------------------------------------------- pillars
    pillar_plan = [
        ("Educate", 0.25, "Teach one useful thing about living with "
                          "handmade pieces", "Brand authority",
         "Carousel", "How to, explainers, care guides"),
        ("Inspire", 0.20, "Show the feeling of a slow, styled home",
         "Awareness", "Reel", "Styling, light, seasonal tables"),
        ("Behind the scenes", 0.18, "Let people watch the making",
         "Community", "Reel", "Studio days, kiln openings, packing"),
        ("Community", 0.12, "Talk with people, not at them",
         "Engagement", "Story", "Questions, votes, customer features"),
        ("Promote", 0.15, "Ask for the sale, clearly and kindly",
         "Sales", "Carousel", "Launches, restocks, gift guides"),
        ("Entertain", 0.05, "Make someone smile and share it",
         "Awareness", "Reel", "Studio bloopers, trends, satisfying clips"),
        ("Personal", 0.05, "The person behind the brand", "Community",
         "Feed post", "Story of the maker, lessons, opinions"),
    ]
    for name, share, purpose, objective, fmt, examples in pillar_plan:
        s.pillars.append(dict(
            name=name, target=share, purpose=purpose, objective=objective,
            format=fmt, examples=examples,
            persona="First-home Freya" if name == "Educate"
                    else "Interiors Isla",
            notes=""))

    # ------------------------------------------------------------- captions
    caption_plan = [
        ("Single stem styling", "Educate",
         "One stem. Three minutes. A table that looks considered.",
         "Save this for your next dinner", "Save this"),
        ("Kiln opening", "Behind the scenes",
         "Opening the kiln never stops being terrifying.",
         "Tell me which glaze wins", "Comment below"),
        ("Autumn collection live", "Promote",
         "The autumn table is in the shop, and there are 40 pieces.",
         "Shop the collection", "Shop now"),
        ("Stoneware vs porcelain", "Educate",
         "Stoneware chips less. Porcelain is lighter. Here is how to "
         "choose.", "Save for your next order", "Save this"),
        ("Customer shelfie", "Community",
         "Your shelves are better styled than mine, honestly.",
         "Tag us in yours", "Tag a friend"),
        ("Why I left my desk job", "Personal",
         "Three years ago I was answering emails. Now I answer to a kiln.",
         "Tell me your leap story", "Reply to this"),
        ("Restock alert", "Promote",
         "The speckled dinnerware is back, and it goes quickly.",
         "Link in bio", "Link in bio"),
        ("Care guide", "Educate",
         "Handmade does not mean delicate. It means made by hands.",
         "Save the care guide", "Save this"),
        ("Gift guide", "Promote",
         "Eight presents for the person whose home always smells of "
         "coffee.", "Shop the guide", "Shop now"),
        ("Slow Sunday", "Inspire",
         "Nothing on the calendar, everything on the table.",
         "Follow for slow living", "Follow for more"),
        ("Studio week", "Behind the scenes",
         "Six days, 180 mugs, one very tired kiln.",
         "Watch the full film", "Watch to the end"),
        ("Holiday shipping", "Promote",
         "Last posting dates are confirmed - here they are.",
         "Order before Friday", "Shop now"),
    ]
    for i, (title, pillar, body, cta_line, cta) in enumerate(caption_plan):
        platform = ["Instagram", "TikTok", "Pinterest", "Facebook"][i % 4]
        hook = body.split(".")[0] + "."
        full = f"{hook}\n\n{body}\n\n{cta_line} {cfg.brand.hashtag}"
        s.captions.append(dict(
            title=title, platform=platform, pillar=pillar, hook=hook,
            body=body, cta=cta_line, full=full,
            short=hook, long=full + "\n\nMade slowly in the studio.",
            persona="First-home Freya" if i % 2 else "Interiors Isla",
            tone="Warm" if i % 3 else "Playful",
            emojis="One or two, never a wall",
            group=f"Group {1 + i % 4}", cta_type=cta,
            offer="20% off seconds" if "sale" in title.lower() else "",
            approval="Approved" if i < 9 else "Awaiting review",
            created=d(today, -70 + i * 5), edited=d(today, -14 + i),
            reusable="Yes" if i % 3 else "No",
            notes="Best performing hook so far" if i == 0 else ""))

    # ------------------------------------------------------------ hashtags
    tag_plan = [
        ("#handmadeceramics", "Broad reach", "Educate", 860000),
        ("#slowliving", "Broad reach", "Inspire", 1200000),
        ("#tablescape", "Niche", "Inspire", 240000),
        ("#potterylife", "Community", "Behind the scenes", 410000),
        ("#wildflowerstudio", "Branded", "Community", 1800),
        ("#handmademug", "Niche", "Promote", 96000),
        ("#stoneware", "Niche", "Educate", 150000),
        ("#kitchenshelfie", "Community", "Community", 72000),
        ("#giftsforhomelovers", "Seasonal", "Promote", 18000),
        ("#autumntable", "Seasonal", "Inspire", 34000),
        ("#madeinthestudio", "Branded", "Behind the scenes", 2400),
        ("#ceramicartist", "Broad reach", "Behind the scenes", 540000),
        ("#smallbatchpottery", "Niche", "Promote", 21000),
        ("#neutralhome", "Broad reach", "Inspire", 680000),
        ("#shopsmall", "Community", "Promote", 2100000),
        ("#glazetest", "Niche", "Behind the scenes", 42000),
        ("#dinnerpartyideas", "Niche", "Educate", 88000),
        ("#holidaygiftguide", "Campaign", "Promote", 120000),
        ("#handmadehome", "Broad reach", "Inspire", 390000),
        ("#ceramicsofinstagram", "Community", "Community", 310000),
        ("#localmaker", "Location", "Community", 64000),
        ("#speckledpottery", "Product", "Promote", 9800),
        ("#mugshot", "Trending", "Entertain", 55000),
        ("#studiotour", "Niche", "Behind the scenes", 47000),
    ]
    keywords = {
        "Educate": ("how to style a table", "Informational"),
        "Inspire": ("autumn tablescape ideas", "Inspirational"),
        "Promote": ("handmade mug gift", "Transactional"),
        "Community": ("handmade home community", "Navigational"),
        "Behind the scenes": ("pottery studio tour", "Informational"),
        "Entertain": ("satisfying pottery video", "Informational"),
    }
    for i, (tag, kind, pillar, volume) in enumerate(tag_plan):
        keyword, intent = keywords.get(pillar, ("handmade ceramics",
                                                "Informational"))
        s.hashtags.append(dict(
            tag=tag, type=kind, pillar=pillar,
            topic=pillar.lower(),
            platform=["Instagram", "TikTok", "Pinterest", "Facebook"][i % 4],
            group=f"Group {1 + i % 4}",
            keyword=keyword, related=f"{keyword}, {pillar.lower()} ideas",
            intent=intent,
            difficulty=["Low", "Medium", "High"][i % 3],
            volume=volume, uses=rnd.randint(2, 28),
            last_used=d(today, -rnd.randint(1, 40)),
            banned="No",
            notes="Branded - always include" if kind == "Branded" else ""))

    # ------------------------------------------------------------- calendar
    def campaign_for(index: int) -> str:
        for start, end, name in CAMPAIGN_FOR:
            if start <= index <= end:
                return name
        return ""

    caption_titles = [c["title"] for c in s.captions]
    asset_ids = [f"AS-{i + 1:04d}" for i in range(14)]

    for i, (offset, platform, fmt, pillar, title) in enumerate(PLAN):
        if offset < 0:
            status = "Published"
        elif offset <= 2:
            status = "Scheduled"
        elif offset <= 6:
            status = "Approved"
        elif offset <= 9:
            status = "Draft"
        elif offset <= 12:
            status = "In production"
        else:
            status = "Brief"
        day = d(today, offset)
        s.calendar.append(dict(
            date=day, slot=SLOTS.get(platform, "Morning"),
            platform=platform, title=title,
            description=f"{pillar} post for {platform.lower()}: {title}.",
            pillar=pillar, format=fmt, campaign=campaign_for(i),
            objective=OBJECTIVES.get(pillar, "Awareness"),
            funnel=FUNNEL.get(pillar, "Awareness"),
            persona=["First-home Freya", "Gifting Grace",
                     "Interiors Isla"][i % 3],
            caption=caption_titles[i % len(caption_titles)],
            hashtag_group=f"Group {1 + i % 4}",
            asset=asset_ids[i % len(asset_ids)],
            cta=CTAS.get(pillar, "Follow for more"),
            owner=OWNERS[i % len(OWNERS)],
            status=status,
            url=(f"{platform.lower()}.com/{brand.lower().replace(' ', '')}"
                 f"/p/{abs(offset):02d}{i:02d}") if offset < 0 else "",
            published="Yes" if offset < 0 else "No",
            notes="Top performer" if i in (3, 24, 35) else ""))

    # --------------------------------------------------------- performance
    # Only published posts are measured.  Engagement is loosely driven by the
    # platform so the comparison tables have something to show.
    base_reach = {"Instagram": 4200, "TikTok": 9800, "Pinterest": 2600,
                  "YouTube": 1500, "Facebook": 1400}
    published = [(i, row) for i, row in enumerate(s.calendar)
                 if row["status"] == "Published"]
    for i, row in published:
        platform = row["platform"]
        spread = rnd.uniform(0.55, 1.85)
        reach = int(base_reach[platform] * spread)
        impressions = int(reach * rnd.uniform(1.15, 1.6))
        rate = rnd.uniform(0.028, 0.085)
        if row["pillar"] in ("Educate", "Community"):
            rate *= 1.25
        if row["pillar"] == "Promote":
            rate *= 0.75
        engagements = max(6, int(reach * rate))
        likes = int(engagements * 0.72)
        comments = max(1, int(engagements * 0.09))
        shares = max(1, int(engagements * 0.07))
        saves = engagements - likes - comments - shares
        clicks = int(reach * rnd.uniform(0.004, 0.022))
        conversions = int(clicks * rnd.uniform(0.02, 0.09))
        s.performance.append(dict(
            content=f"C-{i + 1:04d}", reach=reach, impressions=impressions,
            likes=likes, comments=comments, shares=shares,
            saves=max(0, saves), clicks=clicks, conversions=conversions,
            revenue=round(conversions * rnd.uniform(34, 72), 2),
            watch=(round(rnd.uniform(4.5, 31.0), 1)
                   if platform in ("TikTok", "YouTube", "Instagram")
                   else 0),
            target_rate=0.045,
            notes="Hook tested in A/B 2" if i == 35 else ""))

    # ------------------------------------------------- platform planners --
    def content_of(platform: str, limit: int) -> List[tuple]:
        out = []
        for i, row in enumerate(s.calendar):
            if row["platform"] == platform:
                out.append((f"C-{i + 1:04d}", row))
            if len(out) >= limit:
                break
        return out

    for cid, row in content_of("Instagram", 9):
        reel = row["format"] in ("Reel", "Story")
        s.instagram.append(dict(
            content=cid, format=row["format"], date=row["date"],
            hook="Open on the finished table, then rewind" if reel
                 else "Strong first slide question",
            slides=0 if reel else rnd.choice([5, 6, 7, 8, 10]),
            duration=rnd.choice([12, 18, 24, 30]) if reel else 0,
            cover="covers/ig-" + cid.lower() + ".jpg",
            collab="No", story_sequence="Hook > proof > poll > link"
            if row["format"] == "Story" else "",
            profile_visits=rnd.randint(40, 320),
            followers_gained=rnd.randint(4, 85),
            notes=""))

    for cid, row in content_of("TikTok", 7):
        s.tiktok.append(dict(
            content=cid, date=row["date"],
            hook="First line on screen in under 2 seconds",
            sound="Trending - 'slow morning' lo-fi",
            script="Hook > process > payoff > CTA",
            shots="Wheel close up, hands, finished piece",
            duration=rnd.choice([15, 21, 28, 34, 45]),
            text_on_screen="3 mugs every home needs",
            cover="covers/tt-" + cid.lower() + ".jpg",
            series="Studio diaries" if rnd.random() > 0.5 else "",
            watch_time=round(rnd.uniform(5.5, 22.0), 1),
            completion=round(rnd.uniform(0.28, 0.74), 3),
            profile_views=rnd.randint(60, 540),
            followers_gained=rnd.randint(8, 160),
            notes=""))

    for cid, row in content_of("Pinterest", 7):
        s.pinterest.append(dict(
            content=cid, date=row["date"],
            board="Autumn table styling",
            board_url="pinterest.com/wildflowerstudio/autumn-table",
            pin_title=row["title"],
            pin_description=f"{row['title']} - ideas from a small "
                            f"ceramics studio.",
            destination="shop.wildflowerstudio.com/autumn",
            image="pins/" + cid.lower() + ".png",
            dimensions="1000 x 1500",
            keyword="autumn tablescape ideas",
            intent="Inspirational",
            seasonal="Yes" if row["pillar"] == "Promote" else "No",
            outbound_clicks=rnd.randint(12, 180),
            audience_growth=rnd.randint(5, 60),
            refresh_due=d(today, rnd.randint(30, 150)),
            notes=""))

    for cid, row in content_of("YouTube", 5):
        long_form = row["format"] == "Long video"
        s.youtube.append(dict(
            content=cid, date=row["date"],
            format="Long video" if long_form else "Short",
            title_option=row["title"],
            description="Full studio film with chapters and links.",
            script="Cold open > intro > 4 chapters > CTA",
            thumbnail="thumbs/" + cid.lower() + ".png",
            thumbnail_idea="Hands + kiln glow + 3 word title",
            hook="The kiln is at 1240 degrees and I am nervous",
            chapters="00:00 intro, 02:10 glaze, 06:40 firing",
            keywords="pottery studio tour, handmade ceramics",
            tags="pottery, ceramics, studio vlog",
            playlist="Studio films",
            end_screen="Subscribe + next film",
            watch_time=round(rnd.uniform(120, 940), 1),
            avg_duration=round(rnd.uniform(2.4, 7.8), 1),
            retention=round(rnd.uniform(0.32, 0.61), 3),
            ctr=round(rnd.uniform(0.035, 0.092), 3),
            subs_gained=rnd.randint(4, 70), subs_lost=rnd.randint(0, 6),
            traffic="Browse + suggested",
            notes=""))

    for cid, row in content_of("Facebook", 6):
        s.facebook.append(dict(
            content=cid, date=row["date"], format=row["format"],
            audience="Local + existing customers",
            targeting="Home decor interest, 40km radius",
            link="shop.wildflowerstudio.com",
            reactions=rnd.randint(18, 160), link_clicks=rnd.randint(5, 90),
            followers_gained=rnd.randint(1, 30),
            group_post="Yes" if rnd.random() > 0.7 else "No",
            event="Autumn market" if row["pillar"] == "Community" else "",
            notes=""))

    # ------------------------------------------------------------ campaigns
    campaign_plan = [
        ("Autumn Table Styling", "Awareness", -60, -32, 320.0, 286.40,
         "Autumn collection", "Running"),
        ("Studio Restock", "Sales", -31, -12, 180.0, 164.20,
         "Speckled dinnerware", "Complete"),
        ("Holiday Gift Guide", "Sales", -8, 30, 650.0, 212.75,
         "Gift sets", "Running"),
        ("Evergreen Education", "Brand authority", -120, 180, 0.0, 0.0,
         "All products", "Running"),
    ]
    for (name, objective, start, end, budget, spend, product,
         status) in campaign_plan:
        s.campaigns.append(dict(
            name=name, objective=objective, description=(
                f"{objective} campaign for {product.lower()}."),
            start=d(today, start), end=d(today, end),
            audience="First-home Freya" if objective != "Sales"
                     else "Gifting Grace",
            message="Slow made pieces for a table you actually use",
            product=product,
            platforms="Instagram, TikTok, Pinterest",
            owner=OWNERS[0], budget=budget, spend=spend,
            offer="Free shipping over 75" if objective == "Sales" else "",
            landing="shop.wildflowerstudio.com/" +
                    name.lower().replace(" ", "-"),
            cta="Shop now" if objective == "Sales" else "Follow for more",
            utm="utm_source=social&utm_campaign=" +
                name.lower().replace(" ", "_"),
            leads=rnd.randint(18, 90), conversions=rnd.randint(6, 48),
            revenue=round(rnd.uniform(400, 2600), 2),
            status=status, notes=""))

    # ------------------------------------------------------------- launches
    launch_plan = [
        ("Autumn Table Set", -40, "Complete"),
        ("Holiday Gift Boxes", 12, "In production"),
    ]
    for product, offset, status in launch_plan:
        launch = d(today, offset)
        s.launches.append(dict(
            product=product, launch=launch,
            campaign="Autumn Table Styling" if offset < 0
                     else "Holiday Gift Guide",
            pre_start=d(today, offset - 14), post_end=d(today, offset + 14),
            teaser="Three teaser stories + one reel",
            announcement="Carousel + email on launch morning",
            demo="Reel showing the set on a real table",
            benefits="Dishwasher safe, stacks, made in small batches",
            testimonial="Hannah's wedding set story",
            offer="Early access for the email list",
            code="EARLY10", code_uses=rnd.randint(8, 60),
            reminder=d(today, offset - 1),
            email="Teaser, launch, last call",
            influencer="Mara Lindqvist",
            budget=420.0, engagement=rnd.randint(400, 2400),
            leads=rnd.randint(20, 110), sales=rnd.randint(12, 90),
            revenue=round(rnd.uniform(900, 4200), 2),
            status=status,
            review="Sold out in 6 days - make 30% more next time"
                   if offset < 0 else ""))

    # ------------------------------------------------------------- seasonal
    year = cfg.year
    seasonal_plan = [
        (_dt.date(year, 1, 1), "New Year", "Public holiday", "Fresh start"),
        (_dt.date(year, 2, 14), "Valentine's Day", "Cultural moment",
         "Gifts for two"),
        (_dt.date(year, 3, 20), "First day of spring", "Cultural moment",
         "Spring table"),
        (_dt.date(year, 5, 10), "Mother's Day", "Cultural moment",
         "Gift guide"),
        (_dt.date(year, 6, 21), "Father's Day", "Cultural moment",
         "Mugs for dads"),
        (_dt.date(year, 8, 20), "Back to school", "Seasonal sale",
         "Lunch and storage"),
        (_dt.date(year, 9, 22), "First day of autumn", "Cultural moment",
         "Autumn table launch"),
        (_dt.date(year, 10, 31), "Halloween", "Public holiday",
         "Spooky glaze fun"),
        (_dt.date(year, 11, 27), "Black Friday", "Seasonal sale",
         "Best offer of the year"),
        (_dt.date(year, 11, 30), "Cyber Monday", "Seasonal sale",
         "Online only bundle"),
        (_dt.date(year, 12, 12), "Last posting day", "Industry event",
         "Shipping deadline reminders"),
        (_dt.date(year, 12, 25), "Christmas Day", "Public holiday",
         "Thank you post"),
    ]
    for day, name, kind, angle in seasonal_plan:
        s.seasonal.append(dict(
            date=day, name=name, type=kind, angle=angle,
            campaign="Holiday Gift Guide" if day.month in (11, 12) else "",
            pillar="Promote" if kind == "Seasonal sale" else "Inspire",
            platform="Instagram, Pinterest",
            lead_days=14 if kind == "Seasonal sale" else 7,
            hashtags="#holidaygiftguide" if day.month in (11, 12)
                     else "#slowliving",
            budget=250.0 if kind == "Seasonal sale" else 0.0,
            status="Planned", notes=""))

    # ---------------------------------------------------------- influencers
    influencer_plan = [
        ("Mara Lindqvist", "@maramakes", 24800, 0.052, "Gifted product",
         "Agreed", 0.0, -20),
        ("Tom Whitfield", "@slowhomeuk", 61200, 0.031, "Paid post",
         "Complete", 450.0, -38),
        ("Aisha Rahman", "@thehostingedit", 15400, 0.067, "Affiliate",
         "Negotiating", 0.0, -6),
        ("Nina Castro", "@ceramicdaily", 92000, 0.024, "Account takeover",
         "Contacted", 0.0, -3),
    ]
    for (name, handle, followers, rate, kind, status, fee,
         offset) in influencer_plan:
        s.influencers.append(dict(
            name=name, handle=handle,
            url=f"instagram.com/{handle.strip('@')}",
            contact=f"{handle.strip('@')}@email.com",
            niche="Interiors and slow living", followers=followers,
            engagement=rate, demographic="78% women, 25-44, UK & US",
            type=kind, campaign="Holiday Gift Guide",
            outreach=d(today, offset), status=status,
            response="Replied in 2 days" if status != "Contacted" else "",
            rate=fee, deliverables="1 reel + 3 stories + 1 grid post",
            deadline=d(today, offset + 30),
            approval="Approved" if status == "Complete"
                     else "Awaiting review",
            code=f"{name.split()[0].upper()}10",
            cost=fee + (60.0 if kind == "Gifted product" else 0.0),
            leads=rnd.randint(0, 70), sales=rnd.randint(0, 40),
            revenue=round(rnd.uniform(0, 2200), 2), notes=""))

    # ------------------------------------------------------------------ ugc
    ugc_plan = [
        ("Hannah Price", "Tagged post", "Instagram", "Wedding dinner set",
         5, "Granted"),
        ("Diego Marin", "Story mention", "Instagram", "Speckled mug",
         5, "Granted"),
        ("Keira O'Neill", "Product review", "Other", "Breakfast set",
         4, "Granted"),
        ("Yusuf Demir", "Hashtag", "TikTok", "Plant pots", 5, "Requested"),
        ("Lucy Braun", "Competition entry", "Instagram", "Serving bowl",
         4, "Granted"),
        ("Omar Haddad", "Direct message", "Instagram", "Gift box", 5,
         "Not requested"),
    ]
    for i, (person, source, platform, product, score,
            permission) in enumerate(ugc_plan):
        s.ugc.append(dict(
            person=person, source=source, platform=platform,
            url=f"{platform.lower()}.com/p/ugc{i + 1:03d}",
            type="Photo" if i % 2 else "Video clip", product=product,
            quote="It is the first mug I reach for every single morning.",
            score=score, permission=permission,
            rights="Social use, 12 months" if permission == "Granted"
                   else "",
            expires=d(today, 330) if permission == "Granted" else None,
            attribution=f"Credit @{person.split()[0].lower()}",
            approval="Approved" if permission == "Granted"
                     else "Awaiting review",
            repost=d(today, 3 + i * 2),
            destination="Instagram Story + Facebook",
            engagement=rnd.randint(120, 900), notes=""))

    # ---------------------------------------------------------- repurposing
    repurpose_sources = [2, 6, 12, 22, 24, 30, 35, 37]
    destinations = [("Pinterest", "Standard Pin"), ("Instagram", "Carousel"),
                    ("YouTube", "Short video"), ("Facebook", "Feed post"),
                    ("TikTok", "Reel"), ("Instagram", "Story"),
                    ("Pinterest", "Video Pin"), ("Instagram", "Reel")]
    for i, index in enumerate(repurpose_sources):
        row = s.calendar[index]
        platform, fmt = destinations[i % len(destinations)]
        done = i < 5
        s.repurposing.append(dict(
            source=f"C-{index + 1:04d}", destination_platform=platform,
            new_format=fmt,
            idea=f"Cut the best 20 seconds for {platform}",
            angle="Lead with the result, not the process",
            caption=f"{row['title']} - the short version",
            cta="Save this", asset=f"AS-{(i % 14) + 1:04d}",
            owner=OWNERS[(i + 1) % len(OWNERS)],
            deadline=d(today, -14 + i * 4),
            status="Published" if done else "Planned",
            new_date=d(today, -10 + i * 4) if done else d(today, 7 + i),
            new_url=(f"{platform.lower()}.com/p/re{i:03d}") if done else "",
            notes=""))

    # ----------------------------------------------------------- competitors
    competitor_plan = [
        ("Clay & Co", "Instagram", 48200, 5, "Educate + promote"),
        ("The Pottery Edit", "TikTok", 112000, 7, "Behind the scenes"),
        ("Oat Ceramics", "Pinterest", 26400, 4, "Styling inspiration"),
        ("Studio Lune", "Instagram", 18900, 3, "Personal storytelling"),
    ]
    for i, (name, platform, size, per_week, pillars) in enumerate(
            competitor_plan):
        s.competitors.append(dict(
            name=name, url=f"{platform.lower()}.com/{name.lower()}",
            platform=platform, niche="Handmade ceramics", audience=size,
            frequency=f"{per_week} posts a week", pillars=pillars,
            best_post="Kiln opening reel, 240k views",
            engagement=round(rnd.uniform(0.018, 0.061), 3),
            formats="Reels and carousels",
            campaign="Running a 12 days of gifting series"
                     if i == 0 else "",
            questions="Is it dishwasher safe?",
            trending="Glaze reveal videos",
            news="Pinterest pushing video pins harder",
            gap="Nobody explains how to choose a size",
            opportunity="Own the 'how to choose' search",
            reviewed=d(today, -7 - i * 3),
            source=f"{platform.lower()}.com/{name.lower()}",
            notes=""))

    # ---------------------------------------------------------------- trends
    trend_plan = [
        ("Slow morning lo-fi sound", "Audio", "TikTok", 9, -4),
        ("Guess the sound", "Format", "TikTok", 7, -9),
        ("Shelfie tour", "Format", "Instagram", 8, -16),
        ("12 days of gifting", "Topic", "Instagram", 9, -2),
        ("Before and after styling", "Format", "Pinterest", 6, -22),
    ]
    for name, kind, platform, score, offset in trend_plan:
        s.trends.append(dict(
            name=name, platform=platform, found=d(today, offset),
            url=f"{platform.lower()}.com/trend/{name.split()[0].lower()}",
            type=kind, audio=name if kind == "Audio" else "",
            topic=name, audience="First-home Freya", score=score,
            deadline=d(today, offset + 14),
            idea=f"Use {name.lower()} for a styling reveal",
            owner=OWNERS[1], status="Published" if offset < -10
                                    else "In production",
            posted=d(today, offset + 5) if offset < -10 else None,
            reach=rnd.randint(2000, 24000) if offset < -10 else 0,
            engagement=rnd.randint(120, 1900) if offset < -10 else 0,
            outcome="Best reach of the month" if offset < -10 else "",
            notes=""))

    # ------------------------------------------------------------- analytics
    #: three months of account level numbers per platform
    months = []
    cursor = today.replace(day=1)
    for back in (2, 1, 0):
        month = cursor
        for _ in range(back):
            month = (month - _dt.timedelta(days=1)).replace(day=1)
        months.append(month)
    growth = {"Instagram": 180, "TikTok": 260, "Pinterest": 90,
              "YouTube": 45, "Facebook": 30}
    for account in s.accounts:
        platform = account["platform"]
        followers = account["followers"] - growth[platform] * 3
        for month in months:
            gained = growth[platform] + rnd.randint(-30, 70)
            lost = rnd.randint(2, 28)
            start = followers
            followers = start + gained - lost
            posts = len([row for row in s.calendar
                         if row["platform"] == platform
                         and month_key(row["date"]) == month_key(month)
                         and row["status"] == "Published"])
            reach = int((posts or 1) * {
                "Instagram": 5200, "TikTok": 11800, "Pinterest": 3400,
                "YouTube": 1900, "Facebook": 1600}[platform]
                * rnd.uniform(0.8, 1.3))
            impressions = int(reach * rnd.uniform(1.2, 1.55))
            engagements = int(reach * rnd.uniform(0.035, 0.075))
            clicks = int(reach * rnd.uniform(0.006, 0.02))
            leads = int(clicks * rnd.uniform(0.08, 0.2))
            conversions = int(leads * rnd.uniform(0.2, 0.45))
            s.analytics.append(dict(
                month=month, platform=platform,
                account=account["account"],
                followers_start=start, followers_end=followers,
                lost=lost, posts=posts, reach=reach,
                impressions=impressions, engagements=engagements,
                likes=int(engagements * 0.74),
                comments=int(engagements * 0.1),
                shares=int(engagements * 0.08),
                saves=int(engagements * 0.08),
                profile_visits=int(reach * 0.035),
                clicks=clicks, leads=leads, conversions=conversions,
                revenue=round(conversions * rnd.uniform(38, 74), 2),
                spend=round({"Instagram": 120, "TikTok": 60,
                             "Pinterest": 40, "YouTube": 0,
                             "Facebook": 50}[platform]
                            * rnd.uniform(0.7, 1.4), 2),
                notes=""))

    # ------------------------------------------------------------ A/B tests
    test_plan = [
        ("Hook: question vs statement", "Hook", "Instagram",
         "Question hook beats a statement hook on saves",
         "Which mug do you reach for?", "This is the mug you reach for",
         -30, -23, "Version A", 0.061, 0.043),
        ("Carousel length", "Caption length", "Instagram",
         "Five slides hold attention better than ten",
         "5 slides", "10 slides", -22, -15, "Version B", 0.038, 0.052),
        ("Posting time", "Posting time", "TikTok",
         "Evening posts out-reach morning posts",
         "08:00 post", "19:00 post", -14, -7, "Version B", 0.031, 0.049),
        ("Thumbnail style", "Thumbnail", "YouTube",
         "Hands in frame lift click through rate",
         "Product only", "Hands holding piece", -6, 8, "", 0.041, 0.0),
    ]
    for (name, variable, platform, hypothesis, a, b, start, end, winner,
         a_rate, b_rate) in test_plan:
        s.tests.append(dict(
            name=name, variable=variable, platform=platform,
            hypothesis=hypothesis, version_a=a, version_b=b,
            start=d(today, start), end=d(today, end),
            metric="Engagement rate", secondary="Saves",
            audience="Same account, similar pillar",
            a_result=a_rate, b_result=b_rate,
            a_clicks=rnd.randint(20, 180), b_clicks=rnd.randint(20, 180),
            a_conversions=rnd.randint(1, 20),
            b_conversions=rnd.randint(1, 20),
            winner=winner,
            confidence="Small sample - repeat next month" if winner
                       else "Still running",
            next_test="Try the question hook on TikTok" if winner else "",
            status="Complete" if winner else "Running"))

    # ------------------------------------------------------------------ ads
    ad_plan = [
        ("Autumn collection - traffic", "Instagram", "Traffic", -58, -40,
         180.0, 174.20),
        ("Restock retargeting", "Facebook", "Retargeting", -30, -16,
         120.0, 118.65),
        ("Gift guide - sales", "Instagram", "Sales", -8, 14, 320.0, 96.40),
        ("Pin promotion", "Pinterest", "Traffic", -12, 18, 90.0, 38.20),
        ("Studio film - views", "YouTube", "Video views", -24, -10, 70.0,
         68.90),
        ("Local market - awareness", "Facebook", "Awareness", -6, 2, 45.0,
         41.10),
    ]
    for (name, platform, objective, start, end, budget, spend) in ad_plan:
        impressions = int(spend * rnd.uniform(220, 480))
        reach = int(impressions * rnd.uniform(0.55, 0.8))
        clicks = int(impressions * rnd.uniform(0.008, 0.021))
        conversions = int(clicks * rnd.uniform(0.02, 0.08))
        s.ads.append(dict(
            name=name, platform=platform, objective=objective,
            campaign="Holiday Gift Guide" if start > -20
                     else "Autumn Table Styling",
            audience="Interests: handmade, interiors, slow living",
            segment="Lookalike 2% of buyers" if "retarget" in name.lower()
                    else "Cold interest audience",
            creative=f"AS-{rnd.randint(1, 14):04d}",
            copy="Slow made pieces for a table you actually use",
            headline="Handmade ceramics, made in small batches",
            landing="shop.wildflowerstudio.com",
            start=d(today, start), end=d(today, end),
            budget=budget, spend=spend, impressions=impressions,
            reach=reach, clicks=clicks, conversions=conversions,
            revenue=round(conversions * rnd.uniform(42, 88), 2),
            notes=""))

    # ---------------------------------------------------------------- assets
    asset_plan = [
        ("Autumn table hero", "Photo", "1080 x 1350"),
        ("Studio hands close up", "Photo", "1080 x 1350"),
        ("Kiln opening clip", "Video clip", "1080 x 1920"),
        ("Glaze pour slow motion", "Video clip", "1080 x 1920"),
        ("Carousel template - teal", "Canva template", "1080 x 1350"),
        ("Carousel template - oat", "Canva template", "1080 x 1350"),
        ("Reel cover set", "Reel cover", "1080 x 1920"),
        ("Pin template - vertical", "Canva template", "1000 x 1500"),
        ("Studio film thumbnail", "Thumbnail", "1280 x 720"),
        ("Product shot - speckled mug", "Product shot", "2000 x 2000"),
        ("Product shot - dinner set", "Product shot", "2000 x 2000"),
        ("Gift box flat lay", "Photo", "1080 x 1080"),
        ("Hannah wedding set UGC", "UGC clip", "1080 x 1920"),
        ("Brand logo pack", "Logo", "Vector"),
    ]
    for i, (title, kind, size) in enumerate(asset_plan):
        s.assets.append(dict(
            title=title, type=kind,
            link=f"drive/wildflower/{title.lower().replace(' ', '-')}",
            storage="Google Drive / Brand assets",
            template="canva.com/wildflower-templates"
                     if "template" in kind.lower() else "",
            creator=OWNERS[i % len(OWNERS)],
            created=d(today, -200 + i * 11),
            updated=d(today, -40 + i * 2),
            dimensions=size,
            platform="Instagram, Pinterest" if "1350" in size
                     else "All platforms",
            approval="Approved" if i < 12 else "Awaiting review",
            rights="Owned" if kind != "UGC clip" else "UGC with permission",
            expires=d(today, 300) if kind == "UGC clip" else None,
            product="Autumn collection" if i < 5 else "Gift sets",
            campaign="Autumn Table Styling" if i < 6
                     else "Holiday Gift Guide",
            reuse=rnd.randint(1, 9),
            notes="Highest saving carousel template" if i == 4 else ""))

    # ----------------------------------------------------------------- goals
    goal_plan = [
        ("Instagram followers", "Followers", "Instagram", 7400, 9500,
         300),
        ("TikTok followers", "Followers", "TikTok", 4300, 7000, 300),
        ("Average engagement rate", "Engagement", "All platforms", 0.038,
         0.055, 180),
        ("Monthly website clicks", "Website traffic", "All platforms",
         900, 1800, 120),
        ("Email sign ups", "Email list", "All platforms", 210, 600, 240),
        ("Monthly leads", "Leads", "All platforms", 90, 240, 180),
        ("Posts published per month", "Content output", "All platforms",
         34, 52, 90),
        ("Revenue from social", "Revenue", "All platforms", 9800, 36000,
         300),
    ]
    for i, (name, category, platform, baseline, target,
            days) in enumerate(goal_plan):
        s.goals.append(dict(
            name=name, category=category, platform=platform,
            description=f"Move {name.lower()} from {baseline} to {target}.",
            baseline=baseline, target=target,
            target_date=d(today, days),
            current=round(baseline + (target - baseline)
                          * rnd.uniform(0.25, 0.78), 3),
            milestone_month=round(baseline + (target - baseline) / 12, 2),
            milestone_quarter=round(baseline + (target - baseline) / 4, 2),
            owner=OWNERS[i % len(OWNERS)],
            campaign="Holiday Gift Guide" if category == "Revenue" else "",
            plan="Content Calendar", budget=round(target * 0.02, 2),
            status="On track" if i % 3 else "At risk",
            notes=""))

    # ---------------------------------------------------------------- budget
    budget_plan = [
        ("Paid advertising", 350.0, 348.55, "Meta + Pinterest ads"),
        ("Influencer fees", 150.0, 450.0, "One paid post this quarter"),
        ("Content production", 120.0, 96.40, "Props, flowers, linens"),
        ("Design & editing", 90.0, 88.00, "Video editor, 4 hours"),
        ("Software & subscriptions", 74.0, 74.00, "Scheduler + Canva"),
        ("Contractor fees", 60.0, 0.0, "Spare capacity"),
        ("Props & samples", 40.0, 61.20, "Autumn styling props"),
        ("Training", 25.0, 0.0, "Short course fund"),
        ("Stock media", 15.0, 12.99, "Music licence"),
        ("Printing", 20.0, 18.40, "Market flyers"),
    ]
    for category, planned, actual, note in budget_plan:
        s.budget.append(dict(
            category=category, month=today.replace(day=1),
            platform="All platforms" if category != "Paid advertising"
                     else "Instagram",
            campaign="Holiday Gift Guide" if actual > 100 else "",
            planned=planned, actual=actual, notes=note))

    # ------------------------------------------------------------------ team
    team_plan = [
        ("Jordan Avery", "Owner", "jordan@wildflowerstudio.com", 12),
        ("Priya Nair", "Social media manager", "priya@wildflower.com", 20),
        ("Sam Okonkwo", "Video editor", "sam@freelance.com", 8),
        ("Elena Dias", "Copywriter", "elena@freelance.com", 6),
    ]
    for name, role, email, hours in team_plan:
        s.team.append(dict(
            name=name, role=role, contact=email,
            platform="All platforms", hours=hours,
            rate=0.0 if role == "Owner" else round(28 + hours, 2),
            reviews="Yes" if role in ("Owner", "Social media manager")
                    else "No",
            approves="Yes" if role == "Owner" else "No",
            notes=""))

    task_plan = [
        ("Write December captions", "Elena Dias", -2, "In progress",
         "High"),
        ("Edit holiday reel", "Sam Okonkwo", 1, "Not started", "High"),
        ("Shoot gift box flat lay", "Jordan Avery", 2, "Not started",
         "Medium"),
        ("Schedule week 50", "Priya Nair", 3, "Not started", "High"),
        ("Refresh top 5 pins", "Priya Nair", 6, "Not started", "Medium"),
        ("Reply to UGC permissions", "Priya Nair", -1, "In progress",
         "High"),
        ("Brief influencer collab", "Jordan Avery", 5, "Not started",
         "Medium"),
        ("Update hashtag groups", "Priya Nair", 9, "Not started", "Low"),
        ("Monthly analytics entry", "Priya Nair", 12, "Not started",
         "High"),
        ("Plan January pillars", "Jordan Avery", 16, "Not started", "Low"),
    ]
    for title, owner, offset, status, priority in task_plan:
        s.tasks.append(dict(
            task=title, owner=owner, due=d(today, offset), status=status,
            priority=priority,
            campaign="Holiday Gift Guide" if offset > 0 else "",
            content="", notes=""))

    # ------------------------------------------------------------- ideas --
    idea_plan = [
        ("The 5 mug sizes explained", "Educate", "Instagram", "Carousel", 9),
        ("A day in the studio in 60 seconds", "Behind the scenes", "TikTok",
         "Reel", 8),
        ("How to style open shelving", "Educate", "Pinterest",
         "Standard Pin", 9),
        ("Clay sourcing - where it comes from", "Educate", "YouTube",
         "Long video", 6),
        ("Customer kitchen tour", "Community", "Instagram", "Reel", 7),
        ("Glaze colour vote", "Community", "Instagram", "Story", 8),
        ("Packing an order ASMR", "Behind the scenes", "TikTok",
         "Short video", 7),
        ("What 100 hours of throwing looks like", "Personal", "Instagram",
         "Carousel", 6),
        ("Three dinner party tables", "Inspire", "Pinterest", "Video Pin",
         8),
        ("Myth: handmade is fragile", "Educate", "Instagram", "Reel", 9),
        ("Studio playlist", "Personal", "Facebook", "Feed post", 4),
        ("Seconds sale explainer", "Promote", "Instagram", "Story", 7),
        ("Gift wrapping tutorial", "Educate", "TikTok", "Reel", 8),
        ("Behind a wholesale order", "Behind the scenes", "YouTube",
         "Short video", 5),
        ("Ask the maker anything", "Community", "Instagram", "Story", 7),
        ("Why we price what we price", "Personal", "Instagram", "Carousel",
         8),
        ("Styling a tiny flat", "Inspire", "Pinterest", "Standard Pin", 7),
        ("New glaze test results", "Behind the scenes", "Instagram",
         "Feed post", 6),
    ]
    for i, (title, pillar, platform, fmt, score) in enumerate(idea_plan):
        used = i < 4
        s.ideas.append(dict(
            title=title, description=f"{title} - told in a {fmt.lower()}.",
            persona=["First-home Freya", "Gifting Grace",
                     "Interiors Isla"][i % 3],
            problem="People do not know how to choose or care for pieces",
            objective=OBJECTIVES.get(pillar, "Awareness"),
            pillar=pillar, platform=platform, format=fmt,
            hook=f"What nobody tells you about {title.lower()}",
            message="Handmade is for using, not for display",
            cta=CTAS.get(pillar, "Follow for more"),
            source=["Customer question", "Competitor post", "Trend",
                    "Own idea"][i % 4],
            url="", competitor="Clay & Co" if i % 5 == 0 else "",
            trend="Shelfie tour" if i % 7 == 0 else "",
            evergreen="Yes" if pillar in ("Educate", "Inspire") else "No",
            seasonal="Yes" if "gift" in title.lower() else "No",
            score=score,
            status="Used" if used else ("Shortlisted" if score >= 8
                                        else "New"),
            created=d(today, -60 + i * 3),
            reviewed=d(today, -30 + i * 2) if i % 2 else None,
            planned=d(today, 7 + i * 2) if not used else None,
            content=f"C-{i + 1:04d}" if used else "",
            notes=""))

    return s
````

---

## 3. Setup & planning registers

One module per data tab. Each returns a `SheetSpec` - a list of columns with their formulas - and the writer turns that into a formatted, validated, protected sheet. These are the tabs you fill in before you start posting.

### `social_media_spreadsheet/sheets/__init__.py`

One module per tab of the workbook. *(21 lines)*

````python
"""One module per tab of the workbook."""

from . import (ab_tests, accounts, ads, analytics, annual, assets,
               brand_voice, budget, calc, calendar_sheet, campaigns,
               captions, common, competitors, compliance, dashboard,
               goals, hashtags, help_tab, ideas, influencers,
               launches, lists, month_view, performance, personas,
               pillars, platforms, printables, reports, repurposing,
               seasonal, settings, start_here, tasks, team, trends,
               ugc, workflow)

__all__ = [
    "ab_tests", "accounts", "ads", "analytics", "annual", "assets",
    "brand_voice", "budget", "calc", "calendar_sheet", "campaigns",
    "captions", "common", "competitors", "compliance", "dashboard",
    "goals", "hashtags", "help_tab", "ideas", "influencers",
    "launches", "lists", "month_view", "performance", "personas",
    "pillars", "platforms", "printables", "reports", "repurposing",
    "seasonal", "settings", "start_here", "tasks", "team", "trends",
    "ugc", "workflow",
]
````

### `social_media_spreadsheet/sheets/common.py`

Small building blocks shared by the register sheets. *(167 lines)*

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
    ``ContentPrefix`` or a quoted literal such as ``'"IDEA-"'``.
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
    """Colouring for the content workflow and every other status column."""
    return [
        cf_text("Published", "good"),
        cf_text("Approved", "good"),
        cf_text("Complete", "good"),
        cf_text("Active", "good"),
        cf_text("Agreed", "good"),
        cf_text("Yes", "good"),
        cf_text("Won", "good"),
        cf_text("Scheduled", "info"),
        cf_text("In production", "info"),
        cf_text("Running", "info"),
        cf_text("Sent", "info"),
        cf_text("Testing", "info"),
        cf_text("Idea", "warn"),
        cf_text("Brief", "warn"),
        cf_text("Draft", "warn"),
        cf_text("Awaiting review", "warn"),
        cf_text("Pending", "warn"),
        cf_text("On hold", "warn"),
        cf_text("Revision needed", "bad"),
        cf_text("Rejected", "bad"),
        cf_text("Cancelled", "bad"),
        cf_text("Missed", "bad"),
        cf_text("Overdue", "bad"),
        cf_text("Declined", "bad"),
        cf_text("Archived", "muted"),
        cf_text("Paused", "muted"),
        cf_text("Not started", "muted"),
    ]


def rate_cf(target_name: str = "TargetEngagement") -> List[Dict[str, Any]]:
    """Green above the target rate, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def bar_cf(color: str = "#6350A8") -> List[Dict[str, Any]]:
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


def lookup_content(target: str, key: str = "content") -> str:
    """Pull a column of the Content Calendar for the content ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Content Calendar.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Content Calendar.id]],0)),"Not found"))')
````

### `social_media_spreadsheet/sheets/lists.py`

The `Lists` tab: every drop down list in the workbook lives here. *(194 lines)*

````python
"""
The `Lists` tab: every drop down list in the workbook lives here.

Each column becomes a *dynamic* defined name (OFFSET + COUNTIF) so the user
can add or remove options and every drop down in the workbook follows along.
"""

from __future__ import annotations

from collections import OrderedDict
from typing import List, Tuple

from ..config import Config, Palette
from ..layout import (DATA_ROW, HEADER_ROW, Registry, col_letter,
                      dynamic_name, quote_sheet)
from ..writer import SheetWriter

SHEET = "Lists"
ROWS = 40

#: name -> (column header, options)
LISTS: "OrderedDict[str, Tuple[str, List[str]]]" = OrderedDict([
    ("YesNoList", ("Yes / No", ["Yes", "No"])),
    ("YesNoNAList", ("Yes / No / N/A", ["Yes", "No", "N/A"])),
    ("PlatformList", ("Platform", [
        "Instagram", "TikTok", "Pinterest", "YouTube", "Facebook",
        "LinkedIn", "X (Twitter)", "Threads", "Blog", "Email newsletter",
        "Other"])),
    ("AccountTypeList", ("Account type", [
        "Business account", "Creator account", "Personal profile",
        "Page", "Group", "Channel", "Shop"])),
    ("FormatList", ("Content format", [
        "Feed post", "Carousel", "Reel", "Story", "Short video",
        "Long video", "Live", "Photo", "Graphic", "Infographic",
        "Text post", "Poll", "Standard Pin", "Video Pin", "Idea Pin",
        "Blog post", "Newsletter", "Podcast", "UGC repost"])),
    ("ContentStatusList", ("Content status", [
        "Idea", "Brief", "In production", "Draft", "Awaiting review",
        "Revision needed", "Approved", "Scheduled", "Published",
        "Archived", "Cancelled"])),
    ("ObjectiveList", ("Objective", [
        "Awareness", "Engagement", "Traffic", "Leads", "Sales",
        "Community", "Retention", "Brand authority", "Launch support"])),
    ("FunnelStageList", ("Funnel stage", [
        "Awareness", "Consideration", "Conversion", "Retention"])),
    ("CTAList", ("Call to action", [
        "Follow for more", "Save this", "Share with a friend",
        "Comment below", "Link in bio", "Shop now", "Sign up",
        "Download the guide", "Book a call", "Watch to the end",
        "Reply to this", "Tag a friend", "None"])),
    ("HookTypeList", ("Hook type", [
        "Question", "Bold statement", "Statistic", "Story", "Problem",
        "Myth bust", "How-to", "Listicle", "Before and after", "POV",
        "Tutorial", "Behind the scenes"])),
    ("ToneList", ("Tone of voice", [
        "Friendly", "Playful", "Inspiring", "Educational", "Professional",
        "Warm", "Bold", "Calm", "Witty", "Honest"])),
    ("PriorityList", ("Priority", ["High", "Medium", "Low"])),
    ("IdeaStatusList", ("Idea status", [
        "New", "Shortlisted", "Scheduled", "Used", "On hold",
        "Archived"])),
    ("IdeaSourceList", ("Idea source", [
        "Own idea", "Customer question", "Comment", "Direct message",
        "Product review", "Competitor post", "Trend", "Keyword research",
        "Analytics", "Team suggestion", "Newsletter", "Repurposed"])),
    ("SeasonalTypeList", ("Seasonal event type", [
        "Public holiday", "Awareness day", "Industry event",
        "Seasonal sale", "Product launch", "Local event",
        "Cultural moment"])),
    ("CampaignStatusList", ("Campaign status", [
        "Planned", "Running", "Paused", "Complete", "Cancelled"])),
    ("HashtagTypeList", ("Hashtag type", [
        "Broad reach", "Niche", "Community", "Branded", "Campaign",
        "Product", "Location", "Seasonal", "Trending"])),
    ("SearchIntentList", ("Search intent", [
        "Informational", "Inspirational", "Commercial", "Navigational",
        "Transactional"])),
    ("DifficultyList", ("Keyword difficulty", ["Low", "Medium", "High"])),
    ("AssetTypeList", ("Asset type", [
        "Photo", "Video clip", "Graphic", "Canva template", "Reel cover",
        "Thumbnail", "Carousel set", "Audio", "Font", "Logo", "Mockup",
        "UGC clip", "Product shot"])),
    ("UsageRightsList", ("Usage rights", [
        "Owned", "Licensed", "Royalty free", "UGC with permission",
        "Influencer content", "Stock library", "Expired"])),
    ("ApprovalStatusList", ("Approval status", [
        "Not started", "Awaiting review", "Approved", "Revision needed",
        "Rejected"])),
    ("CollabTypeList", ("Collaboration type", [
        "Gifted product", "Paid post", "Affiliate", "Brand ambassador",
        "Giveaway", "Account takeover", "Guest content", "Event"])),
    ("OutreachStatusList", ("Outreach status", [
        "Not contacted", "Contacted", "Replied", "Negotiating", "Agreed",
        "Declined", "Complete"])),
    ("UGCPermissionList", ("UGC permission", [
        "Not requested", "Requested", "Granted", "Refused", "Expired"])),
    ("UGCSourceList", ("UGC source", [
        "Tagged post", "Story mention", "Product review", "Email",
        "Direct message", "Hashtag", "Competition entry",
        "Influencer content"])),
    ("RoleList", ("Team role", [
        "Owner", "Social media manager", "Content creator", "Copywriter",
        "Designer", "Video editor", "Photographer", "Virtual assistant",
        "Contractor", "Agency"])),
    ("TaskStatusList", ("Task status", [
        "Not started", "In progress", "Blocked", "Complete"])),
    ("GoalCategoryList", ("Goal category", [
        "Followers", "Engagement", "Reach", "Website traffic", "Leads",
        "Sales", "Revenue", "Content output", "Email list",
        "Community"])),
    ("GoalStatusList", ("Goal status", [
        "Not started", "On track", "At risk", "Behind", "Achieved",
        "Missed"])),
    ("CostCategoryList", ("Cost category", [
        "Paid advertising", "Influencer fees", "Content production",
        "Design & editing", "Software & subscriptions", "Contractor fees",
        "Props & samples", "Training", "Stock media", "Printing",
        "Other"])),
    ("AdObjectiveList", ("Ad objective", [
        "Awareness", "Traffic", "Engagement", "Video views",
        "Lead generation", "Sales", "Retargeting", "App installs"])),
    ("TestVariableList", ("Test variable", [
        "Hook", "Caption length", "Thumbnail", "Posting time", "Format",
        "Call to action", "Hashtag set", "Cover image", "Audio", "Offer",
        "First comment"])),
    ("TestStatusList", ("Test status", [
        "Planned", "Running", "Complete", "Inconclusive"])),
    ("TrendTypeList", ("Trend type", [
        "Audio", "Format", "Topic", "Challenge", "Meme", "Hashtag",
        "Platform feature"])),
    ("RepurposeStatusList", ("Repurpose status", [
        "Planned", "In progress", "Published", "Not suitable"])),
    ("FrequencyList", ("Frequency", [
        "Daily", "3 x week", "Weekly", "Fortnightly", "Monthly",
        "Quarterly", "Ad hoc"])),
    ("TimeSlotList", ("Posting slot", [
        "Early morning", "Morning", "Lunchtime", "Afternoon", "Evening",
        "Night"])),
    ("ComplianceStatusList", ("Compliance status", [
        "Not checked", "Checked", "Not required", "Issue found",
        "Fixed"])),
    ("MonthList", ("Month", [
        "All year", "January", "February", "March", "April", "May",
        "June", "July", "August", "September", "October", "November",
        "December"])),
    ("QuarterList", ("Quarter", [
        "All year", "Q1", "Q2", "Q3", "Q4"])),
    ("WeekStartList", ("Week starts on", ["Monday", "Sunday"])),
])


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    sheet = SheetWriter(book, fmts, cfg, reg, SHEET,
                        tab_color=Palette.grey_text)
    last_col = len(LISTS) - 1
    sheet.banner(
        "Customisable Drop Down Lists",
        "Edit, add or delete options below - every drop down in the "
        "workbook updates automatically. Keep each list in one unbroken "
        "block, starting in row 5.",
        last_col,
        tail_text="TIP: lists grow automatically (up to row 44). Deleting "
                  "an option does not change posts that already use it.")

    header_fmt = fmts.chrome("header")
    input_fmt = fmts.input("text")
    for idx, (name, (header, values)) in enumerate(LISTS.items()):
        sheet.set_col(idx, idx, max(16, min(26, len(header) + 8)), input_fmt)
        sheet.write(HEADER_ROW - 1, idx, header, header_fmt)
        for offset, value in enumerate(values):
            sheet.write(DATA_ROW - 1 + offset, idx, value, input_fmt)
        book.define_name(name, dynamic_name(SHEET, col_letter(idx),
                                            DATA_ROW, DATA_ROW + ROWS - 1))

    sheet.ws.set_row(HEADER_ROW - 1, 30)
    sheet.ws.freeze_panes(HEADER_ROW, 0)
    sheet.finish(fit_width=False)
    return sheet


def ref(name: str) -> str:
    """A1 reference to this list's column on the current row, e.g.

        Lists!$AH{r}

    Used by the hidden Calc sheet to mirror a list down a block of rows.
    """
    idx = list(LISTS).index(name)
    return f"{quote_sheet(SHEET)}!${col_letter(idx)}{{r}}"


def options(name: str) -> List[str]:
    """Static copy of a list (used for sample data sanity)."""
    return list(LISTS[name][1])
````

### `social_media_spreadsheet/sheets/settings.py`

The `Settings` tab - business setup, brand basics, targets and ID prefixes. *(195 lines)*

````python
"""
The `Settings` tab - business setup, brand basics, targets and ID prefixes.

Everything the rest of the workbook needs to know about the brand is entered
once here and published as a defined name (BrandName, PlanYear,
TargetEngagement, ProductionDays ...) so the formulas stay readable.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

SHEET = "Settings"
LAST_COL = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, SHEET, tab_color=Palette.accent)
    b = cfg.brand

    s.banner("Business Setup & Settings",
             "Fill in the cream cells once. Grey cells are locked formulas. "
             "Everything else in the workbook reads these values.",
             LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 26, fmts.chrome("label"))
    s.set_col(2, 2, 16, fmts.chrome("label"))
    s.set_col(3, 3, 30)
    s.set_col(4, 4, 20)
    s.set_col(5, 8, 17, fmts.chrome("note"))

    r = 4  # 0 based row index -> Excel row 5

    # ------------------------------------------------------------ 1. brand
    r = s.section(r, "1. Brand details", LAST_COL)
    r = s.field(r, "Business or brand name", b.name, name="BrandName",
                value_end=4,
                help_text="Shown on every printable page and report.")
    r = s.field(r, "Owner / main contact", b.owner, name="BrandOwner",
                value_end=4)
    r = s.field(r, "Email", b.email, name="BrandEmail", value_end=4)
    r = s.field(r, "Phone", b.phone, name="BrandPhone", value_end=4)
    r = s.field(r, "Website", b.website, name="BrandWeb", value_end=4)
    r = s.field(r, "Shop / booking link", b.shop, name="BrandShop",
                value_end=4,
                help_text="The link you send people to - used in the "
                          "campaign and launch planners.")
    r = s.field(r, "Industry", b.industry, name="BrandIndustry",
                value_end=4)
    r = s.field(r, "Niche / speciality", b.niche, name="BrandNiche",
                value_end=4)
    r = s.field(r, "Target audience", b.audience, name="BrandAudience",
                value_end=4,
                help_text="One sentence. The Audience Personas tab breaks "
                          "this down properly.")
    r = s.field(r, "Main marketing goal", b.goal, name="BrandGoal",
                value_end=4)
    r = s.field(r, "Branded hashtag", b.hashtag, name="BrandHashtag",
                value_end=4)

    # -------------------------------------------------------- 2. identity
    r += 1
    r = s.section(r, "2. Visual identity reference", LAST_COL)
    r = s.field(r, "Logo file / link", b.logo_link, name="BrandLogo",
                value_end=4,
                help_text="Paste a cloud link. Spreadsheets cannot preview "
                          "an external image, so keep the link handy.")
    r = s.field(r, "Brand colour palette", b.palette, name="BrandPalette",
                value_end=4)
    r = s.field(r, "Fonts / typography", b.fonts, name="BrandFonts",
                value_end=4)
    r = s.field(r, "Brand asset folder", "brand-assets/", name="BrandFolder",
                value_end=4)

    # -------------------------------------------------------- 3. calendar
    r += 1
    r = s.section(r, "3. Calendar & time settings", LAST_COL)
    r = s.field(r, "Planning year", cfg.year, style="int", name="PlanYear",
                validate=number_check(2000, 2100, decimal=False),
                help_text="The year the dashboards and the 12 month "
                          "calendar open on.")
    r = s.field(r, "Week starts on", cfg.week_start, name="WeekStart",
                validate=list_from("WeekStartList"),
                help_text="Used by the week numbers on the calendar.")
    r = s.field(r, "Time zone", b.timezone, name="BrandTimeZone",
                value_end=4,
                help_text="Posting times on the calendar are in this zone.")
    r = s.field(r, "Currency code", cfg.currency_code, name="CurrencyCode",
                help_text="Budget and revenue columns use the currency "
                          "symbol set when the workbook was built.")
    r = s.field(r, "Fiscal year starts in month", 1, style="int",
                name="FiscalStartMonth",
                validate=number_check(1, 12, decimal=False),
                help_text="1 = January. Used by the quarterly report.")

    # --------------------------------------------------------- 4. targets
    r += 1
    r = s.section(r, "4. Targets (the dashboards measure against these)",
                  LAST_COL)
    r = s.field(r, "Posts per week target", cfg.posts_per_week_target,
                style="int", name="PostsPerWeek",
                validate=number_check(0, 500, decimal=False))
    r = s.field(r, "Posts per month target", cfg.monthly_post_target,
                style="int", name="MonthlyPostTarget",
                validate=number_check(0, 2000, decimal=False),
                help_text="Actual against planned is on the dashboard.")
    r = s.field(r, "Engagement rate target", cfg.engagement_target,
                style="percent", name="TargetEngagement",
                validate=number_check(0, 1),
                help_text="Posts at or above this rate turn green.")
    r = s.field(r, "Follower goal (end of year)", cfg.follower_goal,
                style="int", name="FollowerGoal",
                validate=number_check(0, 100000000, decimal=False))
    r = s.field(r, "Reach goal (per month)", cfg.reach_goal, style="int",
                name="ReachGoal",
                validate=number_check(0, 100000000, decimal=False))
    r = s.field(r, "Website click goal (per month)", cfg.click_goal,
                style="int", name="ClickGoal",
                validate=number_check(0, 10000000, decimal=False))
    r = s.field(r, "Lead goal (per month)", cfg.lead_goal, style="int",
                name="LeadGoal",
                validate=number_check(0, 1000000, decimal=False))
    r = s.field(r, "Conversion goal (per month)", cfg.conversion_goal,
                style="int", name="ConversionGoal",
                validate=number_check(0, 1000000, decimal=False))
    r = s.field(r, "Revenue goal (per year)", cfg.revenue_goal,
                style="money", name="RevenueGoal",
                validate=number_check(0, 100000000))

    # ---------------------------------------------------------- 5. budget
    r += 1
    r = s.section(r, "5. Budget & return", LAST_COL)
    r = s.field(r, "Marketing budget per month", cfg.monthly_budget,
                style="money", name="MonthlyBudget",
                validate=number_check(0, 10000000))
    r = s.field(r, "Marketing budget per year", cfg.annual_budget,
                style="money", name="AnnualBudget",
                validate=number_check(0, 100000000))
    r = s.field(r, "Return on ad spend target", cfg.roas_target,
                style="number1", name="RoasTarget",
                validate=number_check(0, 1000),
                help_text="4 means every 1 spent should bring back 4.")
    r = s.field(r, "Cost per lead target", cfg.cost_per_lead_target,
                style="money", name="CplTarget",
                validate=number_check(0, 100000))

    # -------------------------------------------------------- 6. workflow
    r += 1
    r = s.section(r, "6. Workflow & reminder windows (days)", LAST_COL)
    r = s.field(r, "Draft ready before publishing", cfg.production_days,
                style="int", name="ProductionDays",
                validate=number_check(0, 90, decimal=False),
                help_text="The calendar works out your draft deadline from "
                          "the publishing date.")
    r = s.field(r, "Approval needed before publishing", cfg.approval_days,
                style="int", name="ApprovalDays",
                validate=number_check(0, 90, decimal=False))
    r = s.field(r, "Idea review window", cfg.idea_review_days, style="int",
                name="IdeaReviewDays",
                validate=number_check(0, 365, decimal=False),
                help_text="Ideas older than this still sitting at 'New' are "
                          "flagged.")
    r = s.field(r, "Trend acts-fast window", cfg.trend_window_days,
                style="int", name="TrendWindowDays",
                validate=number_check(0, 365, decimal=False))
    r = s.field(r, "Asset refresh window", cfg.asset_review_days,
                style="int", name="AssetReviewDays",
                validate=number_check(0, 3650, decimal=False))
    r = s.field(r, "Account counts as quiet after", cfg.inactive_account_days,
                style="int", name="QuietAccountDays",
                validate=number_check(0, 365, decimal=False))

    # ------------------------------------------------------- 7. ID prefix
    r += 1
    r = s.section(r, "7. Reference prefixes", LAST_COL)
    r = s.field(r, "Content ID prefix", "C-", name="ContentPrefix",
                help_text="C-0001, C-0002 ... on the Content Calendar.")
    r = s.field(r, "Idea ID prefix", "IDEA-", name="IdeaPrefix")
    r = s.field(r, "Caption ID prefix", "CAP-", name="CaptionPrefix")
    r = s.field(r, "Hashtag ID prefix", "TAG-", name="HashtagPrefix")
    r = s.field(r, "Campaign ID prefix", "CMP-", name="CampaignPrefix")
    r = s.field(r, "Asset ID prefix", "AS-", name="AssetPrefix")

    r += 1
    r = s.note(r, "Tip: everything on this tab is a named cell, so you can "
                  "use the names in your own formulas too - type "
                  "=TargetEngagement or =BrandName in any empty cell to see. "
                  "The drop down options live on the Lists tab.",
               LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `social_media_spreadsheet/sheets/accounts.py`

`Accounts` - one row per social media account, with a profile checklist. *(147 lines)*

````python
"""`Accounts` - one row per social media account, with a profile checklist."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col,
                     rate_cf, status_cf)

NAME = "Accounts"

_YEAR_START = 'DATE(PlanYear,1,1)'
_YEAR_END = 'DATE(PlanYear,12,31)'

#: the six profile checks that make up the optimisation score
CHECKS = [
    ("bio_ok", "Bio / description checked"),
    ("photo_ok", "Profile picture on brand"),
    ("cover_ok", "Cover image current"),
    ("link_ok", "Website link works"),
    ("contact_ok", "Contact details correct"),
    ("branding_ok", "Platform branding done"),
]

_LAST_POST = ('SUMPRODUCT(MAX(([[Content Calendar.platform]]={platform}{r})'
              '*([[Content Calendar.status]]="Published")'
              '*[[Content Calendar.date]]))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("account", '"ACC-"', digits=3),
        Col("platform", "Platform", 14, "text",
            validate=list_from("PlatformList"),
            note="Pick from the Lists tab. Everything else in the workbook "
                 "groups on this."),
        Col("account", "Account / profile name", 22, "text"),
        Col("handle", "Username / handle", 20, "text"),
        Col("url", "Profile URL", 30, "link"),
        Col("type", "Account type", 16, "text",
            validate=list_from("AccountTypeList")),
        Col("owner", "Account owner", 16, "text",
            validate=list_from("TeamNames")),
        Col("purpose", "What this account is for", 34, "wrap"),
        Col("audience", "Primary audience", 22, "text"),
        Col("followers", "Followers today", 13, "int",
            validate=number_check(0, 1000000000, decimal=False),
            note="Type today's number. The monthly history lives on the "
                 "Analytics tab.",
            cf=bar_cf()),
        Col("launched", "Account started", 13, "date",
            validate=date_check()),
        Col("target_per_week", "Posts per week target", 11, "int",
            validate=number_check(0, 100, decimal=False)),
    ]

    for key, header in CHECKS:
        columns.append(Col(key, header, 12, "flag",
                           validate=list_from("YesNoList"),
                           cf=flag_yes_no_cf()))

    checks = "+".join(f'(({{{key}}}{{r}}="Yes")*1)' for key, _h in CHECKS)
    columns += [
        Col("profile_score", "Profile score", 11, "percent",
            formula=blank("account", f"({checks})/{len(CHECKS)}"),
            cf=[{"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "warn"}],
            note="Share of the six profile checks marked Yes."),
        Col("access", "Access & responsibility notes", 30, "wrap"),
        Col("posts_year", "Posts this year", 11, "int",
            formula=blank("platform",
                          'COUNTIFS([[Content Calendar.platform]],'
                          '{platform}{r},[[Content Calendar.date]],">="&'
                          f'{_YEAR_START},[[Content Calendar.date]],"<="&'
                          f'{_YEAR_END})')),
        Col("published_year", "Published", 10, "int",
            formula=blank("platform",
                          'COUNTIFS([[Content Calendar.platform]],'
                          '{platform}{r},[[Content Calendar.status]],'
                          '"Published",[[Content Calendar.date]],">="&'
                          f'{_YEAR_START},[[Content Calendar.date]],"<="&'
                          f'{_YEAR_END})')),
        Col("reach_year", "Reach this year", 13, "int",
            formula=blank("account",
                          'SUMIFS([[Analytics.reach]],'
                          '[[Analytics.account]],{account}{r},'
                          '[[Analytics.platform]],{platform}{r})'),
            cf=bar_cf()),
        Col("engagements_year", "Engagements", 12, "int",
            formula=blank("account",
                          'SUMIFS([[Analytics.engagements]],'
                          '[[Analytics.account]],{account}{r},'
                          '[[Analytics.platform]],{platform}{r})')),
        Col("er_year", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach_year}{r}="",{reach_year}{r}=0),"",'
                    '{engagements_year}{r}/{reach_year}{r})',
            cf=rate_cf()),
        Col("growth_year", "Follower growth", 12, "int",
            formula=blank("account",
                          'SUMIFS([[Analytics.followers_end]],'
                          '[[Analytics.account]],{account}{r})'
                          '-SUMIFS([[Analytics.followers_start]],'
                          '[[Analytics.account]],{account}{r})')),
        Col("last_post", "Last published", 13, "date",
            formula=('=IF({platform}{r}="","",'
                     f'IF({_LAST_POST}=0,"",{_LAST_POST}))')),
        Col("days_quiet", "Days since", 9, "int",
            formula='=IF({last_post}{r}="","",TODAY()-{last_post}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": "=QuietAccountDays",
                 "fmt": "warn"}]),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "account",
                'IF({profile_score}{r}<1,"Action - finish the profile '
                'checklist",IF(AND(ISNUMBER({days_quiet}{r}),'
                '{days_quiet}{r}>QuietAccountDays),"Check - account has '
                'gone quiet",IF(AND({er_year}{r}<>"",'
                '{er_year}{r}<TargetEngagement),"Check - engagement below '
                'target","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap"),
    ]

    # keep the status colouring available for the Yes/No checks
    for col in columns:
        if col.style == "flag" and not col.cf:
            col.cf = status_cf()

    return SheetSpec(
        name=NAME,
        title="Social Media Accounts",
        subtitle="One row per account. The grey columns count your posts, "
                 "pull the reach and engagement from the Analytics tab and "
                 "tell you when a profile needs attention.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).accounts,
        help_text="Add every account you post to - the dashboards compare "
                  "them side by side. Mark the six profile checks Yes once "
                  "you have optimised each profile.",
    )
````

### `social_media_spreadsheet/sheets/personas.py`

`Audience Personas` - who you are talking to, and what to make for them. *(85 lines)*

````python
"""`Audience Personas` - who you are talking to, and what to make for them."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from
from .common import alert_cf, bar_cf, blank, id_col

NAME = "Audience Personas"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"P-"', digits=2, width=8),
        Col("name", "Persona name", 20, "text",
            note="Give them a name you will actually remember, e.g. "
                 "'First-home Freya'."),
        Col("age", "Age range", 11, "text"),
        Col("location", "Location / market", 20, "text"),
        Col("interests", "Interests & hobbies", 30, "wrap"),
        Col("pain", "Main challenge / pain point", 32, "wrap"),
        Col("goal", "What they want", 30, "wrap"),
        Col("motivation", "Buying motivation", 26, "wrap"),
        Col("objection", "Common objection", 26, "wrap"),
        Col("platform", "Favourite platform", 15, "text",
            validate=list_from("PlatformList")),
        Col("format", "Favourite format", 15, "text",
            validate=list_from("FormatList")),
        Col("stage", "Journey stage", 14, "text",
            validate=list_from("FunnelStageList")),
        Col("engagement", "How they engage", 24, "wrap"),
        Col("faq", "Question they always ask", 28, "wrap"),
        Col("language", "Words they use", 24, "wrap"),
        Col("awareness_idea", "Awareness content idea", 28, "wrap"),
        Col("consideration_idea", "Consideration idea", 28, "wrap"),
        Col("conversion_idea", "Conversion idea", 28, "wrap"),
        Col("retention_idea", "Retention idea", 28, "wrap"),
        Col("posts", "Posts planned", 11, "int",
            formula=blank("name",
                          'COUNTIF([[Content Calendar.persona]],'
                          '{name}{r})'),
            cf=bar_cf()),
        Col("published", "Published", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.persona]],{name}{r},'
                          '[[Content Calendar.status]],"Published")')),
        Col("ideas", "Ideas waiting", 11, "int",
            formula=blank("name",
                          'COUNTIFS([[Idea Bank.persona]],{name}{r},'
                          '[[Idea Bank.status]],"<>Used")')),
        Col("captions", "Captions written", 12, "int",
            formula=blank("name",
                          'COUNTIF([[Caption Bank.persona]],{name}{r})')),
        Col("share", "Share of the plan", 12, "percent",
            formula='=IF(OR({name}{r}="",COUNTA([[Content Calendar.persona]])'
                    '=0),"",{posts}{r}/MAX(1,COUNTA('
                    '[[Content Calendar.persona]])))'),
        Col("alert", "Action needed", 24, "status",
            formula=blank(
                "name",
                'IF({posts}{r}=0,"Action - nothing planned for this '
                'persona",IF({captions}{r}=0,"Add - no captions written '
                'yet","OK"))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Audience Research & Customer Personas",
        subtitle="Three or four personas is plenty. Every post on the "
                 "calendar can be pointed at one of them, and the grey "
                 "columns show whether you are actually making content "
                 "for each.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).personas,
        help_text="Write these in your customer's words, not marketing "
                  "words. The four 'idea' columns become entries in the "
                  "Idea Bank.",
    )
````

### `social_media_spreadsheet/sheets/pillars.py`

`Content Pillars` - the themes you post about, and the mix you planned. *(126 lines)*

````python
"""`Content Pillars` - the themes you post about, and the mix you planned."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, bar_cf, blank, id_col, rate_cf

NAME = "Content Pillars"

_YEAR = ('[[Content Calendar.date]],">="&DATE(PlanYear,1,1),'
         '[[Content Calendar.date]],"<="&DATE(PlanYear,12,31)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"PIL-"', digits=2, width=9),
        Col("name", "Pillar name", 20, "text",
            note="Five or six is the sweet spot. These feed the pillar "
                 "drop down on the Content Calendar."),
        Col("target", "Target share of posts", 13, "percent",
            validate=number_check(0, 1),
            note="Aim for the shares to add up to 100%. The total is "
                 "checked at the bottom of the grey columns.",
            cf=bar_cf()),
        Col("purpose", "What this pillar is for", 34, "wrap"),
        Col("objective", "Main objective", 16, "text",
            validate=list_from("ObjectiveList")),
        Col("format", "Usual format", 15, "text",
            validate=list_from("FormatList")),
        Col("persona", "Mainly for", 18, "text",
            validate=list_from("PersonaNames")),
        Col("examples", "Example post ideas", 36, "wrap"),
        Col("planned", "Posts planned", 11, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.pillar]],{name}{r},'
                          f'{_YEAR})'),
            cf=bar_cf()),
        Col("published", "Published", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.pillar]],{name}{r},'
                          '[[Content Calendar.status]],"Published",'
                          f'{_YEAR})')),
        Col("actual", "Actual share", 11, "percent",
            formula='=IF(OR({name}{r}="",SUM([[Content Pillars.planned]])=0)'
                    ',"",{planned}{r}/SUM([[Content Pillars.planned]]))'),
        Col("variance", "Off target by", 11, "percent",
            formula='=IF(OR({actual}{r}="",{target}{r}=""),"",'
                    '{actual}{r}-{target}{r})',
            cf=[{"type": "cell", "criteria": "between", "minimum": -0.05,
                 "maximum": 0.05, "fmt": "good"},
                {"type": "cell", "criteria": "not between",
                 "minimum": -0.1, "maximum": 0.1, "fmt": "bad"}]),
        Col("reach", "Reach", 12, "int",
            formula=blank("name",
                          'SUMIFS([[Performance.reach]],'
                          '[[Performance.pillar]],{name}{r})'),
            cf=bar_cf()),
        Col("engagements", "Engagements", 12, "int",
            formula=blank("name",
                          'SUMIFS([[Performance.engagements]],'
                          '[[Performance.pillar]],{name}{r})')),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("clicks", "Clicks", 10, "int",
            formula=blank("name",
                          'SUMIFS([[Performance.clicks]],'
                          '[[Performance.pillar]],{name}{r})')),
        Col("best", "Best post", 30, "text",
            formula='=IF(OR({name}{r}="",{published}{r}=0),"",'
                    'IFERROR(INDEX([[Performance.title]],MATCH('
                    'SUMPRODUCT(MAX(([[Performance.pillar]]={name}{r})'
                    '*[[Performance.rank_key]])),'
                    '[[Performance.rank_key]],0)),""))',
            note="The post with the highest engagement rate for this "
                 "pillar."),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "name",
                'IF({planned}{r}=0,"Action - no posts planned",'
                'IF({variance}{r}>0.1,"Check - over-used pillar",'
                'IF({variance}{r}<-0.1,"Check - under-used pillar","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Content Pillars & Strategy",
        subtitle="Set the mix you want, then watch the grey columns tell "
                 "you the mix you actually published.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).pillars,
        help_text="Target shares should add up to 100%. 'Off target by' "
                  "turns red when a pillar drifts more than 10 points "
                  "away from its target.",
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals strip under the grid: do the target shares add up?"""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1          # 1 based Excel row
    last = spec_.last_col_index
    s = sheet
    s.merge(row, 0, 1, "Totals", fmts.chrome("total_label"))
    total_fmt = fmts.chrome("total_value", num_format="0%")
    s.ws.write_formula(
        row, spec_.index("target"),
        f"=SUM({spec_.letter('target')}{spec_.first_data_row}:"
        f"{spec_.letter('target')}{spec_.last_data_row})", total_fmt)
    s.ws.write_formula(
        row, spec_.index("planned"),
        f"=SUM({spec_.letter('planned')}{spec_.first_data_row}:"
        f"{spec_.letter('planned')}{spec_.last_data_row})",
        fmts.chrome("total_value", num_format="#,##0"))
    s.merge(row, spec_.index("purpose"), min(last, spec_.index("examples")),
            "The target column should add up to 100%.",
            fmts.chrome("total_value", align="left"))
````

### `social_media_spreadsheet/sheets/calendar_sheet.py`

`Content Calendar` - the master register. *(186 lines)*

````python
"""`Content Calendar` - the master register. Every post lives on this tab.

Everything else in the workbook either feeds this sheet (captions, hashtag
groups, assets, campaigns) or reads it (the dashboards, the workflow view,
the reports and the printable briefs).
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, dup_cf, id_col, month_col, rate_cf,
                     status_cf, week_col, weekday_col)

NAME = "Content Calendar"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("title", "ContentPrefix",
               note="C-0001, C-0002 ... Quote this ID on the Performance, "
                    "Repurposing and platform tabs."),
        Col("date", "Publish date", 13, "date", validate=date_check(),
            note="The day it goes live. Draft and approval deadlines are "
                 "worked out from this."),
        Col("slot", "Time slot", 12, "text",
            validate=list_from("TimeSlotList")),
        month_col("date"),
        week_col("date"),
        weekday_col("date"),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("title", "Content title", 34, "text",
            note="Short and recognisable - this is what shows up in the "
                 "reports and the top content tables."),
        Col("description", "What it is about", 38, "wrap"),
        Col("pillar", "Content pillar", 16, "text",
            validate=list_from("PillarNames")),
        Col("format", "Format", 14, "text",
            validate=list_from("FormatList")),
        Col("campaign", "Campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("objective", "Objective", 15, "text",
            validate=list_from("ObjectiveList")),
        Col("funnel", "Funnel stage", 14, "text",
            validate=list_from("FunnelStageList")),
        Col("persona", "Written for", 18, "text",
            validate=list_from("PersonaNames")),
        Col("caption", "Caption (from the bank)", 24, "text",
            validate=list_from("CaptionTitles")),
        Col("caption_check", "Caption ready", 12, "status",
            formula='=IF({title}{r}="","",IF({caption}{r}="",'
                    '"Add - no caption",IFERROR(INDEX('
                    '[[Caption Bank.approval]],MATCH({caption}{r},'
                    '[[Caption Bank.title]],0)),"Check - not in bank")))',
            cf=alert_cf() + status_cf()),
        Col("hashtag_group", "Hashtag group", 14, "text",
            note="Type a group name you have used on the Hashtag Bank "
                 "tab, e.g. 'Group 1'."),
        Col("hashtag_count", "Tags in group", 11, "int",
            formula='=IF({hashtag_group}{r}="","",COUNTIF('
                    '[[Hashtag Bank.group]],{hashtag_group}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 3,
                 "fmt": "warn"}]),
        Col("asset", "Creative asset", 13, "text",
            validate=list_from("AssetIDs")),
        Col("asset_check", "Asset approved", 12, "status",
            formula='=IF({title}{r}="","",IF({asset}{r}="",'
                    '"Add - no asset",IFERROR(INDEX([[Assets.approval]],'
                    'MATCH({asset}{r},[[Assets.id]],0)),'
                    '"Check - not in library")))',
            cf=alert_cf() + status_cf()),
        Col("cta", "Call to action", 18, "text",
            validate=list_from("CTAList")),
        Col("owner", "Owner", 16, "text",
            validate=list_from("TeamNames")),
        Col("status", "Status", 15, "status",
            validate=list_from("ContentStatusList"), cf=status_cf(),
            note="Move it along: Idea > Brief > In production > Draft > "
                 "Awaiting review > Approved > Scheduled > Published."),
        Col("draft_due", "Draft due", 12, "date",
            formula=blank("date", "{date}{r}-ProductionDays")),
        Col("approve_due", "Approval due", 12, "date",
            formula=blank("date", "{date}{r}-ApprovalDays")),
        Col("days_to_go", "Days to go", 10, "int",
            formula=blank("date", "{date}{r}-TODAY()"),
            cf=[{"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 2, "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "muted"}]),
        Col("url", "Published link", 28, "link"),
        Col("published", "Live?", 10, "status",
            formula='=IF({title}{r}="","",IF({status}{r}="Published",'
                    'IF({url}{r}="","Add - link missing","Yes"),"No"))',
            cf=alert_cf() + [{"type": "text", "criteria": "containing",
                              "value": "Yes", "fmt": "good"}]),
        Col("reach", "Reach", 11, "int",
            formula='=IF({id}{r}="","",IF(COUNTIF([[Performance.content]],'
                    '{id}{r})=0,"",SUMIFS([[Performance.reach]],'
                    '[[Performance.content]],{id}{r})))',
            note="Pulled from the Performance tab - enter the numbers "
                 "there once and they appear here."),
        Col("engagements", "Engagements", 12, "int",
            formula='=IF({reach}{r}="","",SUMIFS('
                    '[[Performance.engagements]],'
                    '[[Performance.content]],{id}{r}))'),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("clicks", "Clicks", 9, "int",
            formula='=IF({reach}{r}="","",SUMIFS([[Performance.clicks]],'
                    '[[Performance.content]],{id}{r}))'),
        Col("repurposed", "Repurposed", 10, "int",
            formula='=IF({id}{r}="","",COUNTIF([[Repurposing.source]],'
                    '{id}{r}))'),
        Col("day_key", "Day key", 11, "number",
            formula='=IF(OR({title}{r}="",{date}{r}=""),"",'
                    '{date}{r}+ROW()/100000)',
            hidden=True,
            note="Hidden helper: lets the Month View grid list the posts "
                 "that fall on each day without sorting the sheet."),
        Col("queue_key", "Queue key", 11, "number",
            formula='=IF(OR({title}{r}="",{date}{r}=""),"",'
                    'IF(OR({status}{r}="Published",'
                    '{status}{r}="Cancelled",{status}{r}="Archived"),"",'
                    'IF({date}{r}<TODAY(),"",{date}{r}+ROW()/100000)))',
            hidden=True,
            note="Hidden helper: lets the dashboard list the next posts "
                 "due without sorting the sheet."),
        Col("late_key", "Late key", 11, "number",
            formula='=IF(OR({title}{r}="",{date}{r}=""),"",'
                    'IF(OR({status}{r}="Published",'
                    '{status}{r}="Cancelled",{status}{r}="Archived"),"",'
                    'IF({date}{r}>=TODAY(),"",{date}{r}+ROW()/100000)))',
            hidden=True,
            note="Hidden helper: feeds the overdue list on the "
                 "Dashboard and the Workflow tab."),
        Col("dup_check", "Duplicate check", 13, "status",
            formula=blank("title",
                          'IF(COUNTIFS([[Content Calendar.title]],'
                          '{title}{r},[[Content Calendar.platform]],'
                          '{platform}{r},[[Content Calendar.date]],'
                          '{date}{r})>1,"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 30, "status",
            formula=blank(
                "title",
                'IF(AND({date}{r}<>"",{date}{r}<TODAY(),'
                '{status}{r}<>"Published",{status}{r}<>"Cancelled",'
                '{status}{r}<>"Archived"),"Overdue - date has passed",'
                'IF(AND({status}{r}="Published",{url}{r}=""),'
                '"Add - publish link missing",'
                'IF(AND({date}{r}=TODAY(),{status}{r}<>"Published"),'
                '"Today - due to go live",'
                'IF(AND({draft_due}{r}<>"",{draft_due}{r}<TODAY(),'
                'OR({status}{r}="Idea",{status}{r}="Brief",'
                '{status}{r}="In production")),"Late - draft not ready",'
                'IF(AND({days_to_go}{r}>=0,{days_to_go}{r}<=3,'
                'OR({status}{r}="Draft",{status}{r}="Awaiting review")),'
                '"Due - needs approval",'
                'IF(AND({title}{r}<>"",{caption}{r}=""),'
                '"Add - caption missing","OK"))))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="12-Month Content Calendar",
        subtitle="One row per post, on any platform, for the whole year. "
                 "Filter by month, platform, pillar or status - the grey "
                 "columns keep the deadlines, links and results up to "
                 "date for you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).calendar,
        help_text="Type the publish date and a title - the ID, month, "
                  "week, deadlines and alerts fill themselves in. Use the "
                  "filter arrows rather than re-sorting so the content IDs "
                  "stay matched to the other tabs.",
    )
````

### `social_media_spreadsheet/sheets/ideas.py`

`Idea Bank` - the parking place for every content idea. *(101 lines)*

````python
"""`Idea Bank` - the parking place for every content idea."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, dup_cf, id_col, lookup_content,
                     score_validation, scale_cf, status_cf, yesno)

NAME = "Idea Bank"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("title", "IdeaPrefix", digits=3),
        Col("title", "Idea", 34, "text",
            note="Capture it quickly - you can tidy it up later."),
        Col("description", "What it would look like", 38, "wrap"),
        Col("persona", "Written for", 18, "text",
            validate=list_from("PersonaNames")),
        Col("problem", "Problem it solves", 30, "wrap"),
        Col("objective", "Objective", 15, "text",
            validate=list_from("ObjectiveList")),
        Col("pillar", "Content pillar", 16, "text",
            validate=list_from("PillarNames")),
        Col("platform", "Best platform", 14, "text",
            validate=list_from("PlatformList")),
        Col("format", "Format", 14, "text",
            validate=list_from("FormatList")),
        Col("hook", "Hook / opening line", 32, "wrap"),
        Col("message", "Key message", 28, "wrap"),
        Col("cta", "Call to action", 18, "text",
            validate=list_from("CTAList")),
        Col("source", "Where it came from", 18, "text",
            validate=list_from("IdeaSourceList")),
        Col("url", "Reference link", 24, "link"),
        Col("competitor", "Seen at (competitor)", 18, "text"),
        Col("trend", "Related trend", 18, "text"),
        Col("evergreen", "Evergreen?", 11, "flag", validate=yesno(),
            cf=status_cf()),
        Col("seasonal", "Seasonal?", 10, "flag", validate=yesno(),
            cf=status_cf()),
        Col("score", "Priority score (1-10)", 12, "rating",
            validate=score_validation(),
            note="How strong is it? 10 = make this next.",
            cf=scale_cf()),
        Col("status", "Status", 15, "status",
            validate=list_from("IdeaStatusList"), cf=status_cf()),
        Col("created", "Captured on", 12, "date", validate=date_check()),
        Col("reviewed", "Last reviewed", 12, "date", validate=date_check()),
        Col("planned", "Planned for", 12, "date", validate=date_check()),
        Col("content", "Became content ID", 13, "text",
            note="Paste the C-0000 ID from the Content Calendar once the "
                 "idea has been scheduled."),
        Col("content_title", "On the calendar as", 28, "text",
            formula=lookup_content("title")),
        Col("content_date", "Publish date", 12, "date",
            formula=lookup_content("date")),
        Col("age", "Days in the bank", 11, "int",
            formula=blank("created", "TODAY()-{created}{r}")),
        Col("review_due", "Review by", 12, "date",
            formula='=IF({title}{r}="","",IF({status}{r}="Used","",'
                    'IF({reviewed}{r}="",{created}{r},{reviewed}{r})'
                    '+IdeaReviewDays))'),
        Col("dup_check", "Duplicate check", 13, "status",
            formula=blank("title",
                          'IF(COUNTIF([[Idea Bank.title]],{title}{r})>1,'
                          '"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "title",
                'IF({status}{r}="Used","Done - on the calendar",'
                'IF(AND({content}{r}<>"",{content_title}{r}=""),'
                '"Check - content ID not found",'
                'IF(AND({review_due}{r}<>"",{review_due}{r}<TODAY()),'
                '"Review - idea is going stale",'
                'IF(AND({score}{r}>=8,{status}{r}<>"Scheduled",'
                '{content}{r}=""),"Plan - strong idea, not scheduled",'
                '"OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Content Idea Bank",
        subtitle="Never start from a blank page again. Score each idea, "
                 "then move the best ones onto the calendar and paste the "
                 "content ID back here.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).ideas,
        help_text="Filter on status 'New' and sort by score when you are "
                  "planning a month. Ideas that have not been looked at "
                  "for a while are flagged for review.",
    )
````

### `social_media_spreadsheet/sheets/captions.py`

`Caption Bank` - reusable hooks, captions and calls to action. *(112 lines)*

````python
"""`Caption Bank` - reusable hooks, captions and calls to action."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, dup_cf, id_col, rate_cf, status_cf,
                     yesno)

NAME = "Caption Bank"

#: character limits used by the length check (generous, platform published)
LIMIT = ('IF({platform}{r}="Instagram",2200,IF({platform}{r}="TikTok",2200,'
         'IF({platform}{r}="Pinterest",500,IF({platform}{r}="YouTube",5000,'
         'IF({platform}{r}="Facebook",63206,IF({platform}{r}="X (Twitter)",'
         '280,IF({platform}{r}="LinkedIn",3000,IF({platform}{r}="Threads",'
         '500,2200))))))))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("title", "CaptionPrefix", digits=3),
        Col("title", "Caption name", 24, "text",
            note="This name appears in the caption drop down on the "
                 "Content Calendar, so keep it unique."),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("pillar", "Content pillar", 16, "text",
            validate=list_from("PillarNames")),
        Col("persona", "Written for", 18, "text",
            validate=list_from("PersonaNames")),
        Col("tone", "Tone of voice", 13, "text",
            validate=list_from("ToneList")),
        Col("hook", "Hook (first line)", 34, "wrap",
            note="The only line most people read. Make it earn the tap."),
        Col("body", "Body", 44, "wrap"),
        Col("cta", "Closing call to action", 26, "wrap"),
        Col("full", "Full caption", 48, "wrap",
            note="Paste the finished caption here, ready to copy into the "
                 "app."),
        Col("short", "Short version", 34, "wrap"),
        Col("long", "Long version", 44, "wrap"),
        Col("cta_type", "CTA type", 18, "text",
            validate=list_from("CTAList")),
        Col("emojis", "Emoji guidance", 20, "text"),
        Col("offer", "Offer mentioned", 18, "text"),
        Col("group", "Hashtag group", 13, "text",
            note="Which hashtag group pairs with this caption."),
        Col("reusable", "Reusable?", 10, "flag", validate=yesno(),
            cf=status_cf()),
        Col("approval", "Approval", 15, "status",
            validate=list_from("ApprovalStatusList"), cf=status_cf()),
        Col("created", "Written on", 12, "date", validate=date_check()),
        Col("edited", "Last edited", 12, "date", validate=date_check()),
        Col("chars", "Characters", 10, "int",
            formula=blank("full", "LEN({full}{r})")),
        Col("limit", "Platform limit", 11, "int",
            formula=blank("platform", LIMIT)),
        Col("fits", "Length check", 12, "status",
            formula='=IF(OR({full}{r}="",{limit}{r}=""),"",'
                    'IF({chars}{r}>{limit}{r},"Too long - trim it",'
                    'IF({chars}{r}<40,"Check - very short","OK")))',
            cf=alert_cf() + [{"type": "text", "criteria": "containing",
                              "value": "Too long", "fmt": "bad"}]),
        Col("uses", "Times used", 10, "int",
            formula=blank("title",
                          'COUNTIF([[Content Calendar.caption]],'
                          '{title}{r})')),
        Col("last_used", "Last used", 12, "date",
            formula='=IF({uses}{r}=0,"",IF({title}{r}="","",'
                    'SUMPRODUCT(MAX(([[Content Calendar.caption]]='
                    '{title}{r})*[[Content Calendar.date]]))))'),
        Col("avg_er", "Average engagement", 13, "percent",
            formula='=IF({uses}{r}=0,"",IFERROR(SUMIFS('
                    '[[Performance.engagements]],'
                    '[[Performance.caption]],{title}{r})/SUMIFS('
                    '[[Performance.reach]],[[Performance.caption]],'
                    '{title}{r}),""))',
            cf=rate_cf()),
        Col("dup_check", "Duplicate check", 13, "status",
            formula=blank("title",
                          'IF(COUNTIF([[Caption Bank.title]],{title}{r})>1,'
                          '"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "title",
                'IF({approval}{r}="Rejected","Rewrite - not approved",'
                'IF({full}{r}="","Add - caption not written yet",'
                'IF({approval}{r}<>"Approved","Check - waiting for '
                'approval",IF({uses}{r}=0,"Use - never used yet","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Caption & Copy Bank",
        subtitle="Write it once, use it again. The length check compares "
                 "your caption against the platform limit, and the grey "
                 "columns show how often each caption has been used and "
                 "how it performed.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).captions,
        help_text="Keep the caption name short - it is what you pick from "
                  "the drop down on the Content Calendar.",
    )
````

### `social_media_spreadsheet/sheets/hashtags.py`

`Hashtag Bank` - hashtag sets, keywords and search intent. *(95 lines)*

````python
"""`Hashtag Bank` - hashtag sets, keywords and search intent."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, dup_cf, id_col, status_cf,
                     yesno)

NAME = "Hashtag Bank"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("tag", "HashtagPrefix", digits=3),
        Col("tag", "Hashtag", 22, "text",
            note="Include the # so you can copy a whole group straight "
                 "into the app."),
        Col("type", "Type", 14, "text",
            validate=list_from("HashtagTypeList"),
            note="Mix broad, niche, community and branded tags in every "
                 "group."),
        Col("group", "Group", 11, "text",
            note="Groups let you reuse a tested set. Type the same group "
                 "name on the Content Calendar."),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("pillar", "Content pillar", 16, "text",
            validate=list_from("PillarNames")),
        Col("topic", "Topic", 16, "text"),
        Col("keyword", "Keyword / search term", 26, "text",
            note="The words people actually type into the search box - "
                 "these matter on Pinterest, YouTube and TikTok."),
        Col("related", "Related keywords", 30, "wrap"),
        Col("intent", "Search intent", 15, "text",
            validate=list_from("SearchIntentList")),
        Col("difficulty", "Competition", 12, "text",
            validate=list_from("DifficultyList"), cf=status_cf()),
        Col("volume", "Posts / search volume", 14, "int",
            validate=number_check(0, 10000000000, decimal=False),
            cf=bar_cf()),
        Col("size", "Size band", 13, "text",
            formula='=IF({volume}{r}="","",IF({volume}{r}>=1000000,'
                    '"Huge - very competitive",IF({volume}{r}>=250000,'
                    '"Large",IF({volume}{r}>=25000,"Medium - sweet spot",'
                    '"Small - easy to rank"))))'),
        Col("uses", "Times used (manual)", 12, "int",
            validate=number_check(0, 100000, decimal=False)),
        Col("posts", "Posts using this group", 13, "int",
            formula=blank("group",
                          'COUNTIF([[Content Calendar.hashtag_group]],'
                          '{group}{r})')),
        Col("last_used", "Last used", 12, "date", validate=date_check()),
        Col("rest_days", "Days rested", 10, "int",
            formula=blank("last_used", "TODAY()-{last_used}{r}")),
        Col("banned", "Banned / flagged?", 12, "flag", validate=yesno(),
            note="Some tags get hidden by the platform. Check before a "
                 "campaign.",
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "good"}]),
        Col("dup_check", "Duplicate check", 13, "status",
            formula=blank("tag",
                          'IF(COUNTIF([[Hashtag Bank.tag]],{tag}{r})>1,'
                          '"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "tag",
                'IF({banned}{r}="Yes","Remove - flagged by the platform",'
                'IF(AND({volume}{r}<>"",{volume}{r}>=1000000),'
                '"Check - huge tag, pair with niche ones",'
                'IF({posts}{r}=0,"Add - group not on the calendar yet",'
                '"OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Hashtag & Keyword Bank",
        subtitle="Build tested groups of 20-30 tags, keep the keywords "
                 "next to them for search, and rotate the groups so your "
                 "posts do not look identical.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).hashtags,
        help_text="Filter on Group to copy a whole set at once. The size "
                  "band tells you whether a tag is realistic to rank for.",
    )
````

### `social_media_spreadsheet/sheets/platforms.py`

The five platform planners, built from one shared factory. *(350 lines)*

````python
"""The five platform planners, built from one shared factory.

Every planner works the same way: pick a content ID from the Content
Calendar, fill in the fields that only matter on that platform, and let the
grey columns pull the title, date, status and the shared results (reach,
engagement, clicks) from the Performance tab.  Nothing is ever typed twice.
"""

from __future__ import annotations

from typing import Any, Dict, List

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, bar_cf, id_col, lookup_content, rate_cf, yesno

CALENDAR = "Content Calendar"


# --------------------------------------------------------------------- base

def _head(platform: str, prefix: str) -> List[Col]:
    """Content ID + the details pulled back from the Content Calendar."""
    return [
        id_col("content", f'"{prefix}"', digits=3, width=9),
        Col("content", "Content ID", 12, "text",
            validate=list_from("ContentIDs"),
            note=f"Pick the C-0000 ID of the {platform} post from the "
                 f"Content Calendar. Everything grey fills itself in."),
        Col("title", "Content title", 30, "text",
            formula=lookup_content("title")),
        Col("date", "Publish date", 12, "date",
            formula=lookup_content("date")),
        Col("pillar", "Pillar", 15, "text",
            formula=lookup_content("pillar")),
        Col("status", "Status", 14, "status",
            formula=lookup_content("status")),
    ]


def _shared_results() -> List[Col]:
    """Results every platform shares - pulled from the Performance tab."""
    pull = ('=IF({content}{r}="","",IF(COUNTIF([[Performance.content]],'
            '{content}{r})=0,"",SUMIFS([[Performance.%s]],'
            '[[Performance.content]],{content}{r})))')
    return [
        Col("reach", "Reach", 11, "int", formula=pull % "reach",
            note="From the Performance tab - enter each post's numbers "
                 "once, there.",
            cf=bar_cf()),
        Col("impressions", "Impressions", 11, "int",
            formula=pull % "impressions"),
        Col("engagements", "Engagements", 11, "int",
            formula=pull % "engagements"),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("saves", "Saves", 9, "int", formula=pull % "saves"),
        Col("shares", "Shares", 9, "int", formula=pull % "shares"),
        Col("comments", "Comments", 10, "int", formula=pull % "comments"),
        Col("clicks", "Clicks", 9, "int", formula=pull % "clicks"),
    ]


def _tail(alert_body: str) -> List[Col]:
    return [
        Col("alert", "Action needed", 28, "status",
            formula=f'=IF({{content}}{{r}}="","",{alert_body})',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]


class Planner:
    """A tiny module-like object so `build.py` can treat it as a sheet."""

    def __init__(self, name: str, platform: str, prefix: str, title: str,
                 subtitle: str, help_text: str, middle, alert: str,
                 samples: str, tab_color: str, rows_key: str) -> None:
        self.NAME = name
        self.platform = platform
        self.prefix = prefix
        self.title = title
        self.subtitle = subtitle
        self.help_text = help_text
        self.middle = middle
        self.alert = alert
        self.samples = samples
        self.tab_color = tab_color
        self.rows_key = rows_key

    def spec(self, cfg: Config) -> SheetSpec:
        columns = (_head(self.platform, self.prefix) + self.middle()
                   + _shared_results() + _tail(self.alert))
        return SheetSpec(
            name=self.NAME,
            title=self.title,
            subtitle=self.subtitle,
            columns=columns,
            rows=cfg.row_count(self.rows_key),
            tab_color=self.tab_color,
            freeze_cols=3,
            samples=getattr(sample(cfg), self.samples),
            help_text=self.help_text,
        )


def _secs(lo: int, hi: int) -> Dict[str, Any]:
    return number_check(lo, hi, decimal=False)


# ---------------------------------------------------------------- instagram

def _instagram_cols() -> List[Col]:
    return [
        Col("format", "Instagram format", 14, "text",
            validate=list_from("FormatList")),
        Col("hook", "Hook / first slide", 32, "wrap"),
        Col("slides", "Carousel slides", 11, "int",
            validate=_secs(0, 20),
            note="Carousels of 6-10 slides tend to hold attention best."),
        Col("duration", "Reel length (sec)", 12, "int",
            validate=_secs(0, 900)),
        Col("cover", "Cover image / file", 22, "text"),
        Col("story_sequence", "Story sequence", 28, "wrap",
            note="Plan the frames: hook, proof, poll, link."),
        Col("collab", "Collab post?", 11, "flag", validate=yesno()),
        Col("profile_visits", "Profile visits", 12, "int",
            validate=_secs(0, 10000000)),
        Col("followers_gained", "Followers gained", 13, "int",
            validate=_secs(-100000, 10000000)),
    ]


INSTAGRAM = Planner(
    name="Instagram", platform="Instagram", prefix="IG-",
    title="Instagram Planner",
    subtitle="Grid, reels and stories in one place. Fill in the Instagram "
             "specific fields - the title, date and results come from the "
             "calendar and the Performance tab.",
    help_text="Reels under 30 seconds and carousels of 6-10 slides are the "
              "usual sweet spots. Plan your story frames in the sequence "
              "column.",
    middle=_instagram_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF(AND({format}{r}="Reel",{duration}{r}=0),'
          '"Add - reel length missing",IF({cover}{r}="",'
          '"Add - no cover image",IF(AND({reach}{r}<>"",'
          '{er}{r}<TargetEngagement),"Review - below your target rate",'
          '"OK"))))',
    samples="instagram", tab_color=Palette.primary_mid,
    rows_key="Instagram")


# ------------------------------------------------------------------- tiktok

def _tiktok_cols() -> List[Col]:
    return [
        Col("hook", "First 2 seconds", 30, "wrap",
            note="TikTok lives or dies on the first two seconds."),
        Col("sound", "Sound / audio", 24, "text"),
        Col("script", "Script outline", 32, "wrap"),
        Col("shots", "Shot list", 28, "wrap"),
        Col("duration", "Length (sec)", 11, "int", validate=_secs(0, 1800)),
        Col("text_on_screen", "Text on screen", 24, "wrap"),
        Col("cover", "Cover frame", 20, "text"),
        Col("series", "Series name", 18, "text"),
        Col("watch_time", "Avg watch time (sec)", 13, "number1",
            validate=number_check(0, 100000)),
        Col("completion", "Completion rate", 12, "percent",
            validate=number_check(0, 1),
            cf=[{"type": "cell", "criteria": ">=", "value": 0.5,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.25,
                 "fmt": "warn"}]),
        Col("profile_views", "Profile views", 11, "int",
            validate=_secs(0, 10000000)),
        Col("followers_gained", "Followers gained", 13, "int",
            validate=_secs(-100000, 10000000)),
    ]


TIKTOK = Planner(
    name="TikTok", platform="TikTok", prefix="TT-",
    title="TikTok Planner",
    subtitle="Hooks, sounds, scripts and shot lists, plus the watch time "
             "and completion rate that decide how far a video travels.",
    help_text="Write the hook as the words on screen in the first two "
              "seconds. Completion rate matters more than length.",
    middle=_tiktok_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF({hook}{r}="","Add - no hook written",'
          'IF({sound}{r}="","Add - no sound chosen",'
          'IF(AND({completion}{r}<>"",{completion}{r}<0.25),'
          '"Review - most people scroll away early","OK"))))',
    samples="tiktok", tab_color=Palette.primary_mid, rows_key="TikTok")


# ---------------------------------------------------------------- pinterest

def _pinterest_cols() -> List[Col]:
    return [
        Col("board", "Board", 22, "text"),
        Col("board_url", "Board link", 26, "link"),
        Col("pin_title", "Pin title", 30, "text",
            note="Pinterest is a search engine - put the keyword in the "
                 "title."),
        Col("pin_description", "Pin description", 38, "wrap"),
        Col("destination", "Destination link", 28, "link"),
        Col("image", "Image file", 20, "text"),
        Col("dimensions", "Size (2:3 is best)", 13, "text"),
        Col("keyword", "Target keyword", 24, "text"),
        Col("intent", "Search intent", 14, "text",
            validate=list_from("SearchIntentList")),
        Col("seasonal", "Seasonal pin?", 11, "flag", validate=yesno(),
            note="Seasonal pins want publishing 30-45 days early."),
        Col("outbound_clicks", "Outbound clicks", 12, "int",
            validate=_secs(0, 10000000)),
        Col("audience_growth", "Audience growth", 12, "int",
            validate=_secs(-100000, 10000000)),
        Col("refresh_due", "Refresh / repin on", 13, "date"),
    ]


PINTEREST = Planner(
    name="Pinterest", platform="Pinterest", prefix="PN-",
    title="Pinterest Planner",
    subtitle="Pins keep working for months, so plan the keyword, the board "
             "and the refresh date as carefully as the image.",
    help_text="Use 2:3 images (1000 x 1500), put your keyword in the pin "
              "title and description, and publish seasonal pins 30-45 days "
              "ahead.",
    middle=_pinterest_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF({keyword}{r}="","Add - no target keyword",'
          'IF({destination}{r}="","Add - no destination link",'
          'IF(AND({refresh_due}{r}<>"",{refresh_due}{r}<TODAY()),'
          '"Due - time to refresh this pin","OK"))))',
    samples="pinterest", tab_color=Palette.primary_mid,
    rows_key="Pinterest")


# ------------------------------------------------------------------ youtube

def _youtube_cols() -> List[Col]:
    return [
        Col("format", "Long or Short", 12, "text",
            validate=list_from("FormatList")),
        Col("title_option", "Video title", 32, "text",
            note="Under 60 characters so it is not cut off in search."),
        Col("title_chars", "Title length", 10, "int",
            formula='=IF({title_option}{r}="","",LEN({title_option}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 60,
                 "fmt": "warn"}]),
        Col("hook", "First 15 seconds", 30, "wrap"),
        Col("description", "Description", 36, "wrap"),
        Col("script", "Script / structure", 32, "wrap"),
        Col("chapters", "Chapters", 26, "wrap"),
        Col("thumbnail", "Thumbnail file", 20, "text"),
        Col("thumbnail_idea", "Thumbnail idea", 26, "wrap"),
        Col("keywords", "Target keywords", 26, "wrap"),
        Col("tags", "Tags", 24, "wrap"),
        Col("playlist", "Playlist", 18, "text"),
        Col("end_screen", "End screen / cards", 22, "text"),
        Col("watch_time", "Watch time (min)", 13, "number1",
            validate=number_check(0, 10000000)),
        Col("avg_duration", "Avg view (min)", 12, "number1",
            validate=number_check(0, 100000)),
        Col("retention", "Audience retention", 12, "percent",
            validate=number_check(0, 1),
            cf=[{"type": "cell", "criteria": ">=", "value": 0.5,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.3,
                 "fmt": "warn"}]),
        Col("ctr", "Thumbnail CTR", 11, "percent",
            validate=number_check(0, 1),
            cf=[{"type": "cell", "criteria": ">=", "value": 0.05,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.03,
                 "fmt": "warn"}]),
        Col("subs_gained", "Subscribers gained", 13, "int",
            validate=_secs(0, 10000000)),
        Col("subs_lost", "Subscribers lost", 12, "int",
            validate=_secs(0, 10000000)),
        Col("net_subs", "Net subscribers", 12, "int",
            formula='=IF({subs_gained}{r}="","",{subs_gained}{r}'
                    '-N({subs_lost}{r}))'),
        Col("traffic", "Main traffic source", 20, "text"),
    ]


YOUTUBE = Planner(
    name="YouTube", platform="YouTube", prefix="YT-",
    title="YouTube Planner",
    subtitle="Titles, thumbnails, chapters and the two numbers that drive "
             "the algorithm: click through rate and retention.",
    help_text="A 4-6% thumbnail CTR and 50%+ retention are healthy. Keep "
              "titles under 60 characters and always fill in the chapters.",
    middle=_youtube_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF({thumbnail}{r}="","Add - no thumbnail",'
          'IF(AND({ctr}{r}<>"",{ctr}{r}<0.03),'
          '"Review - thumbnail click rate is low",'
          'IF(AND({retention}{r}<>"",{retention}{r}<0.3),'
          '"Review - viewers leave early","OK"))))',
    samples="youtube", tab_color=Palette.primary_mid, rows_key="YouTube")


# ----------------------------------------------------------------- facebook

def _facebook_cols() -> List[Col]:
    return [
        Col("format", "Post format", 14, "text",
            validate=list_from("FormatList")),
        Col("audience", "Audience", 22, "text"),
        Col("targeting", "Targeting notes", 26, "wrap"),
        Col("link", "Link shared", 26, "link"),
        Col("group_post", "Posted to a group?", 12, "flag",
            validate=yesno()),
        Col("event", "Linked event", 20, "text"),
        Col("reactions", "Reactions", 11, "int",
            validate=_secs(0, 10000000)),
        Col("link_clicks", "Link clicks", 11, "int",
            validate=_secs(0, 10000000)),
        Col("followers_gained", "Page followers gained", 14, "int",
            validate=_secs(-100000, 10000000)),
    ]


FACEBOOK = Planner(
    name="Facebook", platform="Facebook", prefix="FB-",
    title="Facebook Planner",
    subtitle="Page posts, groups and events - the place where links still "
             "work and your existing customers still hang out.",
    help_text="Facebook rewards conversation. Reply to every comment in "
              "the first hour and keep link posts short.",
    middle=_facebook_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF(AND({link}{r}="",{format}{r}<>"Photo"),'
          '"Add - no link on a link post",IF(AND({reach}{r}<>"",'
          '{er}{r}<TargetEngagement),"Review - below your target rate",'
          '"OK")))',
    samples="facebook", tab_color=Palette.primary_mid, rows_key="Facebook")


#: every planner, in tab order
PLANNERS = [INSTAGRAM, TIKTOK, PINTEREST, YOUTUBE, FACEBOOK]
````

---

## 4. Campaign & production registers

The tabs that run the work: campaigns and launches, the seasonal plan, collaborations, repurposing, research and the team.

### `social_media_spreadsheet/sheets/campaigns.py`

`Campaigns` - themed pushes with a budget, a goal and a result. *(137 lines)*

````python
"""`Campaigns` - themed pushes with a budget, a goal and a result."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, rate_cf,
                     status_cf)

NAME = "Campaigns"

_POSTS = 'COUNTIF([[Content Calendar.campaign]],{name}{r})'
_PERF = ('IF(COUNTIF([[Performance.campaign]],{name}{r})=0,"",'
         'SUMIFS([[Performance.%s]],[[Performance.campaign]],{name}{r}))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "CampaignPrefix", digits=3),
        Col("name", "Campaign name", 24, "text",
            note="This name appears in the campaign drop down on the "
                 "Content Calendar, so keep it short and unique."),
        Col("objective", "Objective", 15, "text",
            validate=list_from("ObjectiveList")),
        Col("description", "What it is", 34, "wrap"),
        Col("product", "Product / service", 22, "text"),
        Col("audience", "Audience", 18, "text",
            validate=list_from("PersonaNames")),
        Col("message", "Key message", 30, "wrap"),
        Col("platforms", "Platforms", 22, "text"),
        Col("start", "Start date", 12, "date", validate=date_check()),
        Col("end", "End date", 12, "date", validate=date_check()),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("offer", "Offer / hook", 20, "text"),
        Col("landing", "Landing page", 26, "link"),
        Col("cta", "Call to action", 18, "text",
            validate=list_from("CTAList")),
        Col("utm", "Tracking link / UTM", 28, "text",
            note="Add a UTM tag so your website analytics can tell you "
                 "which campaign sent the visitor."),
        Col("budget", "Budget", 12, "money", validate=number_check(0, 1e9),
            cf=bar_cf()),
        Col("spend", "Spent so far", 12, "money",
            validate=number_check(0, 1e9)),
        Col("leads", "Leads", 9, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("conversions", "Sales / conversions", 12, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("status", "Status", 13, "status",
            validate=list_from("CampaignStatusList"), cf=status_cf()),
        Col("days", "Length (days)", 11, "int",
            formula='=IF(OR({start}{r}="",{end}{r}=""),"",'
                    '{end}{r}-{start}{r}+1)'),
        Col("posts", "Posts planned", 11, "int",
            formula=blank("name", _POSTS), cf=bar_cf()),
        Col("published", "Published", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.campaign]],'
                          '{name}{r},[[Content Calendar.status]],'
                          '"Published")')),
        Col("reach", "Reach", 12, "int",
            formula=blank("name", _PERF % "reach"), cf=bar_cf()),
        Col("engagements", "Engagements", 12, "int",
            formula=blank("name", _PERF % "engagements")),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("clicks", "Clicks", 9, "int",
            formula=blank("name", _PERF % "clicks")),
        Col("remaining", "Budget left", 12, "money",
            formula='=IF({budget}{r}="","",{budget}{r}-N({spend}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("budget_used", "Budget used", 11, "percent",
            formula='=IF(OR({budget}{r}="",{budget}{r}=0),"",'
                    'N({spend}{r})/{budget}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bad"},
                {"type": "data_bar", "bar_color": "#F2665E"}]),
        Col("cpl", "Cost per lead", 12, "money",
            formula='=IF(OR({leads}{r}="",{leads}{r}=0),"",'
                    'N({spend}{r})/{leads}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": "=CplTarget",
                 "fmt": "warn"}]),
        Col("cpa", "Cost per sale", 12, "money",
            formula='=IF(OR({conversions}{r}="",{conversions}{r}=0),"",'
                    'N({spend}{r})/{conversions}{r})'),
        Col("conv_rate", "Click to sale", 11, "percent",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({conversions}{r})/{clicks}{r})'),
        Col("roas", "Return on spend", 12, "number1",
            formula='=IF(OR({spend}{r}="",{spend}{r}=0),"",'
                    'N({revenue}{r})/{spend}{r})',
            note="Revenue divided by spend. 3.0 means three back for "
                 "every one spent.",
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": "=RoasTarget",
                 "fmt": "warn"}]),
        Col("profit", "Revenue less spend", 13, "money",
            formula='=IF(AND({revenue}{r}="",{spend}{r}=""),"",'
                    'N({revenue}{r})-N({spend}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(N({spend}{r})>N({budget}{r}),"Over budget",'
                'IF(AND({status}{r}="Running",{posts}{r}=0),'
                '"Action - running with no posts planned",'
                'IF(AND({end}{r}<>"",{end}{r}<TODAY(),'
                '{status}{r}="Running"),"Overdue - mark it complete",'
                'IF(AND({roas}{r}<>"",{roas}{r}<RoasTarget),'
                '"Review - return below target","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]
    return SheetSpec(
        name=NAME,
        title="Campaign Planner",
        subtitle="Give every push a name, a window and a budget, then let "
                 "the grey columns count the posts, pull the results and "
                 "work out the return.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).campaigns,
        help_text="Tag posts with the campaign name on the Content "
                  "Calendar - reach, engagement and clicks then roll up "
                  "here automatically.",
    )
````

### `social_media_spreadsheet/sheets/launches.py`

`Launch Planner` - the run-up, the launch day and the follow-through. *(120 lines)*

````python
"""`Launch Planner` - the run-up, the launch day and the follow-through."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, status_cf)

NAME = "Launch Planner"

_WINDOW = ('COUNTIFS([[Content Calendar.date]],">="&{pre_start}{r},'
           '[[Content Calendar.date]],"<="&{post_end}{r})')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("product", '"LA-"', digits=3, width=9),
        Col("product", "Product / offer", 24, "text"),
        Col("launch", "Launch date", 12, "date", validate=date_check(),
            note="The day it goes on sale. Everything else counts "
                 "backwards and forwards from here."),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("pre_start", "Teasing starts", 12, "date",
            validate=date_check()),
        Col("post_end", "Follow-up ends", 12, "date",
            validate=date_check()),
        Col("teaser", "Teaser content", 28, "wrap"),
        Col("announcement", "Announcement content", 28, "wrap"),
        Col("demo", "Demo / proof content", 28, "wrap"),
        Col("benefits", "Benefits to lead with", 28, "wrap"),
        Col("testimonial", "Testimonial / social proof", 26, "wrap"),
        Col("offer", "Launch offer", 22, "text"),
        Col("code", "Discount code", 14, "text"),
        Col("code_uses", "Code used", 10, "int",
            validate=number_check(0, 1000000, decimal=False)),
        Col("reminder", "Last-call reminder", 13, "date",
            validate=date_check()),
        Col("email", "Email sequence", 24, "wrap"),
        Col("influencer", "Influencer support", 20, "text",
            validate=list_from("InfluencerNames")),
        Col("budget", "Budget", 12, "money", validate=number_check(0, 1e9)),
        Col("engagement", "Engagements", 12, "int",
            validate=number_check(0, 1e8, decimal=False)),
        Col("leads", "Leads / waitlist", 12, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("sales", "Sales", 10, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("status", "Status", 13, "status",
            validate=list_from("CampaignStatusList"), cf=status_cf()),
        Col("days_to_go", "Days to launch", 11, "int",
            formula=blank("launch", "{launch}{r}-TODAY()")),
        Col("phase", "Phase today", 15, "text",
            formula='=IF({launch}{r}="","",IF(TODAY()<{pre_start}{r},'
                    '"Not started yet",IF(TODAY()<{launch}{r},'
                    '"Teasing",IF(TODAY()={launch}{r},"Launch day",'
                    'IF(TODAY()<={post_end}{r},"Follow-up",'
                    '"Finished")))))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Launch day", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Teasing", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Follow-up", "fmt": "info"}]),
        Col("runway", "Teasing days", 10, "int",
            formula='=IF(OR({launch}{r}="",{pre_start}{r}=""),"",'
                    '{launch}{r}-{pre_start}{r})',
            note="Two weeks of warming up beats one surprise post.",
            cf=[{"type": "cell", "criteria": "<", "value": 7,
                 "fmt": "warn"}]),
        Col("posts", "Posts in the window", 12, "int",
            formula='=IF(OR({pre_start}{r}="",{post_end}{r}=""),"",'
                    f'{_WINDOW})',
            cf=bar_cf()),
        Col("aov", "Average order", 12, "money",
            formula='=IF(OR({sales}{r}="",{sales}{r}=0),"",'
                    'N({revenue}{r})/{sales}{r})'),
        Col("conv", "Leads to sales", 11, "percent",
            formula='=IF(OR({leads}{r}="",{leads}{r}=0),"",'
                    'N({sales}{r})/{leads}{r})'),
        Col("roas", "Return on spend", 12, "number1",
            formula='=IF(OR({budget}{r}="",{budget}{r}=0),"",'
                    'N({revenue}{r})/{budget}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": "=RoasTarget",
                 "fmt": "warn"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "product",
                'IF(AND({runway}{r}<>"",{runway}{r}<7),'
                '"Action - less than a week of teasing",'
                'IF(AND({phase}{r}="Teasing",{posts}{r}=0),'
                '"Action - nothing on the calendar yet",'
                'IF(AND({launch}{r}<TODAY(),{status}{r}<>"Complete",'
                '{status}{r}<>"Cancelled"),"Overdue - close this launch '
                'off",IF({reminder}{r}=TODAY(),"Today - send the last '
                'call","OK"))))'),
            cf=alert_cf()),
        Col("review", "What to change next time", 30, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Product Launch Planner",
        subtitle="Plan the teasing, the launch day and the follow-up as "
                 "one piece of work, then write down what you would "
                 "change next time.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).launches,
        help_text="A fortnight of teasing, a clear launch day and a week "
                  "of follow-up sells far more than a single post.",
    )
````

### `social_media_spreadsheet/sheets/seasonal.py`

`Seasonal Planner` - holidays, awareness days and seasonal moments. *(74 lines)*

````python
"""`Seasonal Planner` - holidays, awareness days and seasonal moments."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, month_col, status_cf)

NAME = "Seasonal Planner"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"SEA-"', digits=3, width=10),
        Col("date", "Date", 12, "date", validate=date_check()),
        month_col("date"),
        Col("name", "Event / moment", 24, "text"),
        Col("type", "Type", 16, "text",
            validate=list_from("SeasonalTypeList")),
        Col("angle", "Your angle on it", 34, "wrap",
            note="Only join in where you have something real to say."),
        Col("pillar", "Content pillar", 16, "text",
            validate=list_from("PillarNames")),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("platform", "Platforms", 20, "text"),
        Col("lead_days", "Plan this many days ahead", 13, "int",
            validate=number_check(0, 365, decimal=False),
            note="Pinterest wants 30-45 days, Instagram about 14."),
        Col("hashtags", "Hashtags", 22, "text"),
        Col("budget", "Budget", 12, "money", validate=number_check(0, 1e9)),
        Col("status", "Status", 13, "status",
            validate=list_from("CampaignStatusList"), cf=status_cf()),
        Col("start_by", "Start work by", 12, "date",
            formula='=IF({date}{r}="","",{date}{r}-N({lead_days}{r}))'),
        Col("days_to_go", "Days to go", 10, "int",
            formula=blank("date", "{date}{r}-TODAY()"),
            cf=[{"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 14, "fmt": "warn"}]),
        Col("posts", "Posts planned", 11, "int",
            formula='=IF({date}{r}="","",COUNTIFS('
                    '[[Content Calendar.date]],">="&{start_by}{r},'
                    '[[Content Calendar.date]],"<="&{date}{r}))',
            cf=bar_cf()),
        Col("alert", "Action needed", 30, "status",
            formula=blank(
                "name",
                'IF(AND({date}{r}<TODAY(),{status}{r}<>"Complete",'
                '{status}{r}<>"Cancelled"),"Expired - the date has '
                'passed",IF(AND({start_by}{r}<=TODAY(),'
                '{date}{r}>=TODAY(),{posts}{r}=0),'
                '"Action - start planning this now",'
                'IF(AND({days_to_go}{r}>=0,{days_to_go}{r}<=14),'
                '"Upcoming - within a fortnight","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Seasonal & Holiday Content Planner",
        subtitle="The dates that come round every year. Set how far ahead "
                 "each one needs planning and the sheet tells you when to "
                 "start.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=3,
        samples=sample(cfg).seasonal,
        help_text="Fill this in once and it works for every year - just "
                  "update the dates each January.",
    )
````

### `social_media_spreadsheet/sheets/influencers.py`

`Influencers` - outreach, deliverables, fees and what came back. *(115 lines)*

````python
"""`Influencers` - outreach, deliverables, fees and what came back."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf,
                     status_cf)

NAME = "Influencers"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"INF-"', digits=3, width=10),
        Col("name", "Name", 20, "text"),
        Col("handle", "Handle", 16, "text"),
        Col("url", "Profile link", 26, "link"),
        Col("contact", "Email / contact", 24, "text"),
        Col("niche", "Niche", 22, "text"),
        Col("followers", "Followers", 12, "int",
            validate=number_check(0, 1e9, decimal=False), cf=bar_cf()),
        Col("engagement", "Their engagement rate", 13, "percent",
            validate=number_check(0, 1),
            note="Look at it before you look at follower count - 2-6% is "
                 "healthy.",
            cf=[{"type": "cell", "criteria": ">=", "value": 0.03,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.01,
                 "fmt": "warn"}]),
        Col("demographic", "Their audience", 26, "wrap"),
        Col("type", "Collaboration type", 18, "text",
            validate=list_from("CollabTypeList")),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("outreach", "Contacted on", 12, "date", validate=date_check()),
        Col("status", "Outreach status", 15, "status",
            validate=list_from("OutreachStatusList"), cf=status_cf()),
        Col("response", "Their response", 24, "wrap"),
        Col("rate", "Their fee", 12, "money",
            validate=number_check(0, 1e8)),
        Col("deliverables", "Agreed deliverables", 28, "wrap"),
        Col("deadline", "Due date", 12, "date", validate=date_check()),
        Col("approval", "Content approval", 15, "status",
            validate=list_from("ApprovalStatusList"), cf=status_cf()),
        Col("code", "Discount / affiliate code", 15, "text"),
        Col("cost", "Total cost", 12, "money",
            validate=number_check(0, 1e8),
            note="Fee plus product, postage and anything else you gave "
                 "away."),
        Col("leads", "Leads", 9, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("sales", "Sales", 9, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("cost_per_sale", "Cost per sale", 12, "money",
            formula='=IF(OR({sales}{r}="",{sales}{r}=0),"",'
                    'N({cost}{r})/{sales}{r})'),
        Col("roi", "Return on spend", 12, "number1",
            formula='=IF(OR({cost}{r}="",{cost}{r}=0),"",'
                    'N({revenue}{r})/{cost}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "bad"}]),
        Col("profit", "Revenue less cost", 13, "money",
            formula='=IF(AND({revenue}{r}="",{cost}{r}=""),"",'
                    'N({revenue}{r})-N({cost}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("value", "Cost per 1,000 followers", 13, "money",
            formula='=IF(OR({followers}{r}="",{followers}{r}=0),"",'
                    'N({cost}{r})/({followers}{r}/1000))'),
        Col("work_again", "Worth repeating?", 13, "status",
            formula='=IF({name}{r}="","",IF({roi}{r}="","Not measured '
                    'yet",IF({roi}{r}>=RoasTarget,"Yes - do it again",'
                    'IF({roi}{r}>=1,"Maybe - it washed its face",'
                    '"No - it cost more than it made"))))',
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Maybe", "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "No -",
                 "fmt": "bad"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(AND({status}{r}="Contacted",{outreach}{r}<>"",'
                'TODAY()-{outreach}{r}>7),"Action - follow up, no reply '
                'in a week",IF(AND({deadline}{r}<>"",'
                '{deadline}{r}<TODAY(),{status}{r}<>"Complete"),'
                '"Overdue - content not delivered",'
                'IF(AND({status}{r}="Agreed",{approval}{r}<>"Approved"),'
                '"Check - content still to approve","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Influencer & Collaboration Tracker",
        subtitle="From first message to final invoice - and, most "
                 "importantly, whether the collaboration paid for itself.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).influencers,
        help_text="Judge a creator on engagement rate and audience fit, "
                  "not follower count. Record the total cost including "
                  "gifted product.",
    )
````

### `social_media_spreadsheet/sheets/ugc.py`

`UGC Tracker` - customer content, permission and where you reposted it. *(100 lines)*

````python
"""`UGC Tracker` - customer content, permission and where you reposted it."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, scale_cf, score_validation,
                     status_cf)

NAME = "UGC Tracker"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("person", '"UGC-"', digits=3, width=10),
        Col("person", "Customer / creator", 22, "text"),
        Col("source", "Where you found it", 18, "text",
            validate=list_from("UGCSourceList")),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("url", "Link to the post", 28, "link"),
        Col("type", "Type", 14, "text", validate=list_from("AssetTypeList")),
        Col("product", "Product shown", 22, "text"),
        Col("quote", "Quote / review", 36, "wrap"),
        Col("score", "Quality score (1-10)", 12, "rating",
            validate=score_validation(), cf=scale_cf()),
        Col("permission", "Permission", 14, "status",
            validate=list_from("UGCPermissionList"),
            note="Always ask before you repost. A screenshot of the yes "
                 "is worth keeping.",
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Not requested", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Refused", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Expired", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Requested", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Granted", "fmt": "good"}]),
        Col("rights", "Usage rights agreed", 22, "wrap"),
        Col("expires", "Rights expire", 12, "date", validate=date_check()),
        Col("attribution", "Credit line", 20, "text"),
        Col("approval", "Approved to use", 15, "status",
            validate=list_from("ApprovalStatusList"), cf=status_cf()),
        Col("repost", "Repost planned for", 13, "date",
            validate=date_check()),
        Col("destination", "Where you will use it", 24, "text"),
        Col("engagement", "Engagements it got", 13, "int",
            validate=number_check(0, 1e8, decimal=False)),
        Col("days_left", "Rights left (days)", 12, "int",
            formula=blank("expires", "{expires}{r}-TODAY()"),
            cf=[{"type": "cell", "criteria": "<", "value": 30,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("usable", "Safe to post?", 14, "status",
            formula=blank(
                "person",
                'IF({permission}{r}<>"Granted","No - no permission yet",'
                'IF(AND({expires}{r}<>"",{expires}{r}<TODAY()),'
                '"No - rights expired",IF({approval}{r}<>"Approved",'
                '"Check - not approved yet","Yes")))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Check", "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "No -",
                 "fmt": "bad"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "person",
                'IF({permission}{r}="Requested","Action - chase the '
                'permission",IF(AND({usable}{r}="Yes",{repost}{r}=""),'
                '"Plan - schedule this repost",'
                'IF(AND({repost}{r}<>"",{repost}{r}<TODAY(),'
                '{approval}{r}="Approved"),"Overdue - repost date has '
                'passed",IF(AND({days_left}{r}<>"",{days_left}{r}<30,'
                '{days_left}{r}>=0),"Check - rights expire soon",'
                '"OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="User Generated Content Tracker",
        subtitle="Your customers make your best adverts. Track what they "
                 "post, whether you have permission to use it, and when "
                 "those rights run out.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).ugc,
        help_text="Never repost without asking. Record the permission, "
                  "the credit line and the date the rights expire.",
    )
````

### `social_media_spreadsheet/sheets/repurposing.py`

`Repurposing` - one good post, used five more times. *(89 lines)*

````python
"""`Repurposing` - one good post, used five more times."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, id_col, lookup_content, rate_cf,
                     status_cf)

NAME = "Repurposing"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("source", '"RP-"', digits=3, width=9),
        Col("source", "Original content ID", 13, "text",
            validate=list_from("ContentIDs"),
            note="The post you are reusing. Pick its C-0000 ID."),
        Col("source_title", "Original post", 30, "text",
            formula=lookup_content("title", "source")),
        Col("source_platform", "Was on", 13, "text",
            formula=lookup_content("platform", "source")),
        Col("source_date", "Original date", 12, "date",
            formula=lookup_content("date", "source")),
        Col("source_er", "How it performed", 13, "percent",
            formula='=IF({source}{r}="","",IFERROR(SUMIFS('
                    '[[Performance.engagements]],[[Performance.content]],'
                    '{source}{r})/SUMIFS([[Performance.reach]],'
                    '[[Performance.content]],{source}{r}),""))',
            note="Repurpose what already worked - this is the engagement "
                 "rate of the original.",
            cf=rate_cf()),
        Col("destination_platform", "Reuse on", 14, "text",
            validate=list_from("PlatformList")),
        Col("new_format", "New format", 14, "text",
            validate=list_from("FormatList")),
        Col("idea", "What you will change", 32, "wrap"),
        Col("angle", "New angle", 28, "wrap"),
        Col("caption", "New caption / title", 30, "wrap"),
        Col("cta", "Call to action", 18, "text",
            validate=list_from("CTAList")),
        Col("asset", "Asset to use", 13, "text",
            validate=list_from("AssetIDs")),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("deadline", "Do it by", 12, "date", validate=date_check()),
        Col("status", "Status", 14, "status",
            validate=list_from("RepurposeStatusList"), cf=status_cf()),
        Col("new_date", "Reposted on", 12, "date", validate=date_check()),
        Col("new_url", "New link", 26, "link"),
        Col("times_used", "Times reused", 11, "int",
            formula=blank("source",
                          'COUNTIF([[Repurposing.source]],{source}{r})')),
        Col("days_late", "Days late", 10, "int",
            formula='=IF(OR({deadline}{r}="",{status}{r}="Published",'
                    '{status}{r}="Not suitable"),"",'
                    'MAX(0,TODAY()-{deadline}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "source",
                'IF({source_title}{r}="Not found","Check - content ID not '
                'on the calendar",IF(AND({status}{r}="Published",'
                '{new_url}{r}=""),"Add - link missing",'
                'IF(AND({days_late}{r}<>"",{days_late}{r}>0),'
                '"Overdue - past the deadline",'
                'IF(AND({status}{r}="Planned",{deadline}{r}=""),'
                '"Add - no deadline set","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Content Repurposing Tracker",
        subtitle="Your best reel is a carousel, a pin, a short and a "
                 "newsletter. Plan the second life of every post that "
                 "worked.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=3,
        samples=sample(cfg).repurposing,
        help_text="Sort by 'How it performed' and repurpose from the top "
                  "down. Add the new post to the Content Calendar too, so "
                  "it gets its own ID and results.",
    )
````

### `social_media_spreadsheet/sheets/competitors.py`

`Competitors` - what the people in your niche are doing. *(82 lines)*

````python
"""`Competitors` - what the people in your niche are doing."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, bar_cf, blank, id_col, rate_cf

NAME = "Competitors"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"COM-"', digits=3, width=10),
        Col("name", "Competitor", 20, "text"),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("url", "Profile link", 26, "link"),
        Col("niche", "Their niche", 22, "text"),
        Col("audience", "Followers", 12, "int",
            validate=number_check(0, 1e9, decimal=False), cf=bar_cf()),
        Col("frequency", "How often they post", 18, "text"),
        Col("pillars", "Their content themes", 30, "wrap"),
        Col("formats", "Formats they use", 24, "wrap"),
        Col("best_post", "Their best post", 30, "wrap"),
        Col("engagement", "Their engagement rate", 13, "percent",
            validate=number_check(0, 1),
            note="Engagements divided by followers on a handful of recent "
                 "posts - near enough is good enough.",
            cf=rate_cf()),
        Col("campaign", "Campaign they are running", 26, "wrap"),
        Col("questions", "Questions their audience asks", 28, "wrap"),
        Col("trending", "What is working for them", 26, "wrap"),
        Col("news", "Platform / industry news", 26, "wrap"),
        Col("gap", "Gap you could fill", 28, "wrap",
            note="The most valuable column on the sheet."),
        Col("opportunity", "Your move", 28, "wrap"),
        Col("reviewed", "Last reviewed", 12, "date", validate=date_check()),
        Col("source", "Where you looked", 22, "text"),
        Col("vs_you", "Versus your rate", 13, "percent",
            formula='=IF({engagement}{r}="","",'
                    '{engagement}{r}-TargetEngagement)',
            note="Their engagement rate minus your target rate.",
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}]),
        Col("ideas", "Ideas taken from them", 11, "int",
            formula=blank("name",
                          'COUNTIF([[Idea Bank.competitor]],{name}{r})')),
        Col("days_since", "Days since review", 11, "int",
            formula=blank("reviewed", "TODAY()-{reviewed}{r}"),
            cf=[{"type": "cell", "criteria": ">", "value": 60,
                 "fmt": "warn"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(OR({reviewed}{r}="",{days_since}{r}>60),'
                '"Action - review this account again",'
                'IF({gap}{r}="","Add - no gap identified yet",'
                'IF({ideas}{r}=0,"Add - no ideas captured from them",'
                '"OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Competitor Research",
        subtitle="Watch four or five accounts properly rather than forty "
                 "badly. Look for the gap, then fill it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).competitors,
        help_text="Review every month or two. Note their best post and "
                  "the questions people ask in their comments - that is "
                  "free content research.",
    )
````

### `social_media_spreadsheet/sheets/trends.py`

`Trends` - sounds, formats and topics worth jumping on quickly. *(85 lines)*

````python
"""`Trends` - sounds, formats and topics worth jumping on quickly."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, scale_cf, score_validation,
                     status_cf)

NAME = "Trends"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"TR-"', digits=3, width=9),
        Col("name", "Trend / sound / format", 24, "text"),
        Col("type", "Type", 14, "text", validate=list_from("TrendTypeList")),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("found", "Spotted on", 12, "date", validate=date_check()),
        Col("url", "Example link", 26, "link"),
        Col("audio", "Audio / sound name", 22, "text"),
        Col("topic", "Topic", 20, "text"),
        Col("audience", "Who it suits", 18, "text",
            validate=list_from("PersonaNames")),
        Col("score", "Fit for your brand (1-10)", 13, "rating",
            validate=score_validation(),
            note="Only ride a trend if it fits. A forced trend reads as "
                 "a forced trend.",
            cf=scale_cf()),
        Col("idea", "How you would use it", 32, "wrap"),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("deadline", "Use it before", 12, "date", validate=date_check(),
            note="Trends go stale fast - a fortnight is usually the "
                 "whole window."),
        Col("status", "Status", 15, "status",
            validate=list_from("ContentStatusList"), cf=status_cf()),
        Col("posted", "Posted on", 12, "date", validate=date_check()),
        Col("reach", "Reach it got", 12, "int",
            validate=number_check(0, 1e9, decimal=False)),
        Col("engagement", "Engagements", 12, "int",
            validate=number_check(0, 1e8, decimal=False)),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({engagement}{r})/{reach}{r})'),
        Col("age", "Days old", 9, "int",
            formula=blank("found", "TODAY()-{found}{r}")),
        Col("window", "Days left", 10, "int",
            formula='=IF(OR({deadline}{r}="",{posted}{r}<>""),"",'
                    '{deadline}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 3,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("outcome", "What happened", 28, "wrap"),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(AND({posted}{r}="",{deadline}{r}<>"",'
                '{deadline}{r}<TODAY()),"Expired - the moment has passed",'
                'IF(AND({posted}{r}="",{score}{r}>=8),'
                '"Action - strong fit, make it now",'
                'IF(AND({posted}{r}="",{age}{r}>TrendWindowDays),'
                '"Check - this has been sitting here a while","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Trend Tracker",
        subtitle="Catch a trend in the first week or leave it alone. "
                 "Score the fit, set a use-by date, and record what "
                 "actually happened.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).trends,
        help_text="Check your platform's trending page once a week and "
                  "add anything that fits. Anything older than your trend "
                  "window gets flagged.",
    )
````

### `social_media_spreadsheet/sheets/assets.py`

`Assets` - the creative library: photos, clips, templates, thumbnails. *(92 lines)*

````python
"""`Assets` - the creative library: photos, clips, templates, thumbnails."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, dup_cf, id_col, status_cf)

NAME = "Assets"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("title", "AssetPrefix", digits=4,
               note="AS-0001, AS-0002 ... Quote this ID on the Content "
                    "Calendar, the Ads tab and the Repurposing tab."),
        Col("title", "Asset name", 26, "text"),
        Col("type", "Type", 15, "text", validate=list_from("AssetTypeList")),
        Col("link", "File link", 30, "link",
            note="Paste the Drive, Dropbox or Canva link. The spreadsheet "
                 "stores the link, not the file."),
        Col("storage", "Where it lives", 24, "text"),
        Col("template", "Template link", 24, "link"),
        Col("creator", "Made by", 16, "text",
            validate=list_from("TeamNames")),
        Col("created", "Created", 12, "date", validate=date_check()),
        Col("updated", "Last updated", 12, "date", validate=date_check()),
        Col("dimensions", "Size", 13, "text"),
        Col("platform", "Suits", 20, "text"),
        Col("product", "Product shown", 20, "text"),
        Col("campaign", "Campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("approval", "Approval", 15, "status",
            validate=list_from("ApprovalStatusList"), cf=status_cf()),
        Col("rights", "Usage rights", 20, "text",
            validate=list_from("UsageRightsList")),
        Col("expires", "Rights expire", 12, "date", validate=date_check()),
        Col("reuse", "Times reused (manual)", 12, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("used_posts", "Used on posts", 11, "int",
            formula=blank("id",
                          'COUNTIF([[Content Calendar.asset]],{id}{r})'),
            cf=bar_cf()),
        Col("used_ads", "Used in ads", 10, "int",
            formula=blank("id", 'COUNTIF([[Ads.creative]],{id}{r})')),
        Col("last_used", "Last used on", 12, "date",
            formula='=IF({used_posts}{r}=0,"",SUMPRODUCT(MAX('
                    '([[Content Calendar.asset]]={id}{r})'
                    '*[[Content Calendar.date]])))'),
        Col("age_days", "Days old", 9, "int",
            formula=blank("created", "TODAY()-{created}{r}")),
        Col("rights_left", "Rights left (days)", 12, "int",
            formula=blank("expires", "{expires}{r}-TODAY()"),
            cf=[{"type": "cell", "criteria": "<", "value": 30,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("dup_check", "Duplicate check", 12, "status",
            formula=blank("title",
                          'IF(COUNTIF([[Assets.title]],{title}{r})>1,'
                          '"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "title",
                'IF(AND({expires}{r}<>"",{expires}{r}<TODAY()),'
                '"Expired - stop using this file",'
                'IF({link}{r}="","Add - no file link saved",'
                'IF({approval}{r}<>"Approved","Check - not approved for '
                'use",IF({used_posts}{r}=0,"Use - never used yet",'
                '"OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Creative Asset Library",
        subtitle="Every photo, clip, template and thumbnail with its "
                 "link, its rights and a count of where it has been "
                 "used.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=3,
        samples=sample(cfg).assets,
        help_text="Keep the files in your cloud drive and the links "
                  "here. Licensed and UGC files get an expiry date so "
                  "you never use something you no longer have rights to.",
    )
````

### `social_media_spreadsheet/sheets/team.py`

`Team` - who does what, and how much of it. *(91 lines)*

````python
"""`Team` - who does what, and how much of it."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col, yesno)

NAME = "Team"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"TM-"', digits=2, width=8),
        Col("name", "Name", 20, "text",
            note="These names fill the 'Owner' drop downs across the "
                 "workbook."),
        Col("role", "Role", 20, "text", validate=list_from("RoleList")),
        Col("contact", "Email / contact", 26, "text"),
        Col("platform", "Looks after", 20, "text"),
        Col("hours", "Hours a week", 11, "number1",
            validate=number_check(0, 168), cf=bar_cf()),
        Col("rate", "Hourly rate", 12, "money",
            validate=number_check(0, 1e6)),
        Col("reviews", "Reviews content?", 12, "flag", validate=yesno(),
            cf=flag_yes_no_cf()),
        Col("approves", "Final approval?", 12, "flag", validate=yesno(),
            cf=flag_yes_no_cf()),
        Col("weekly_cost", "Cost a week", 12, "money",
            formula='=IF(OR({hours}{r}="",{rate}{r}=""),"",'
                    '{hours}{r}*{rate}{r})'),
        Col("monthly_cost", "Cost a month", 12, "money",
            formula='=IF({weekly_cost}{r}="","",'
                    '{weekly_cost}{r}*52/12)'),
        Col("posts", "Posts owned", 11, "int",
            formula=blank("name",
                          'COUNTIF([[Content Calendar.owner]],{name}{r})'),
            cf=bar_cf()),
        Col("published", "Published", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.owner]],{name}{r},'
                          '[[Content Calendar.status]],"Published")')),
        Col("late", "Late posts", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.owner]],{name}{r},'
                          '[[Content Calendar.alert]],"Overdue*")'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}]),
        Col("tasks_open", "Open tasks", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Tasks.owner]],{name}{r},'
                          '[[Tasks.status]],"<>Complete")')),
        Col("tasks_late", "Overdue tasks", 11, "int",
            formula=blank("name",
                          'COUNTIFS([[Tasks.owner]],{name}{r},'
                          '[[Tasks.status]],"<>Complete",'
                          '[[Tasks.due]],"<"&TODAY())'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("on_time", "On-time rate", 11, "percent",
            formula='=IF(OR({posts}{r}="",{posts}{r}=0),"",'
                    '1-{late}{r}/{posts}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": 0.9,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.75,
                 "fmt": "warn"}]),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "name",
                'IF({tasks_late}{r}>0,"Action - overdue tasks",'
                'IF({late}{r}>0,"Check - late content",'
                'IF({posts}{r}=0,"Add - nothing assigned","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Team & Responsibilities",
        subtitle="Even a team of one. Set who owns what, and the sheet "
                 "counts the posts and tasks sitting with each person.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).team,
        help_text="If it is just you, add one row with your name - every "
                  "owner drop down in the workbook reads this list.",
    )
````

### `social_media_spreadsheet/sheets/tasks.py`

`Tasks` - the jobs behind the posts. *(71 lines)*

````python
"""`Tasks` - the jobs behind the posts."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, id_col, lookup_content, status_cf,
                     week_col)

NAME = "Tasks"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("task", '"TSK-"', digits=3, width=9),
        Col("task", "Task", 34, "text"),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("due", "Due date", 12, "date", validate=date_check()),
        week_col("due"),
        Col("priority", "Priority", 11, "text",
            validate=list_from("PriorityList"),
            cf=[{"type": "text", "criteria": "containing", "value": "High",
                 "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Medium", "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "Low",
                 "fmt": "muted"}]),
        Col("status", "Status", 14, "status",
            validate=list_from("TaskStatusList"), cf=status_cf()),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("content", "Linked content", 13, "text",
            validate=list_from("ContentIDs")),
        Col("content_title", "Which post", 28, "text",
            formula=lookup_content("title")),
        Col("days_to_go", "Days to go", 10, "int",
            formula='=IF(OR({due}{r}="",{status}{r}="Complete"),"",'
                    '{due}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 2, "fmt": "warn"}]),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "task",
                'IF({status}{r}="Complete","OK",'
                'IF(AND({due}{r}<>"",{due}{r}<TODAY()),"Overdue",'
                'IF({due}{r}=TODAY(),"Today - due now",'
                'IF({status}{r}="Blocked","Action - unblock this",'
                'IF(AND({days_to_go}{r}<>"",{days_to_go}{r}<=2),'
                '"Due - this week","OK")))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Task Management",
        subtitle="Shoots, edits, approvals, replies and the monthly "
                 "analytics entry - everything that is not a post but "
                 "still has to happen.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).tasks,
        help_text="Filter on 'Overdue' in the action column each Monday "
                  "morning - that is your week sorted.",
    )
````

### `social_media_spreadsheet/sheets/compliance.py`

`Compliance` - the legal and quality checks, written out as a list. *(155 lines)*

````python
"""`Compliance` - the legal and quality checks, written out as a list.

The checks themselves are part of the product, so they are written onto
the sheet whether or not the demo data is switched on.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..formats import INPUT
from ..layout import Col, SheetSpec
from ..writer import date_check, list_from
from .common import alert_cf, blank, id_col, status_cf

NAME = "Compliance"

#: (area, check, why it matters, how often)
CHECKS = [
    ("Advertising", "Paid partnerships marked as an ad",
     "Undisclosed paid posts break advertising rules in most countries.",
     "Every post"),
    ("Advertising", "Gifted products disclosed",
     "Free product counts as payment - say so.", "Every post"),
    ("Advertising", "Affiliate links disclosed",
     "Commission links must be labelled.", "Every post"),
    ("Advertising", "Platform paid-partnership label switched on",
     "The caption hashtag is not always enough on its own.",
     "Every post"),
    ("Claims", "Product claims are true and provable",
     "'Handmade', 'organic' and 'sustainable' all have legal meanings.",
     "Every post"),
    ("Claims", "Before and after images are honest",
     "No filters or edits that change the result shown.", "Every post"),
    ("Claims", "Prices, discounts and stock are accurate",
     "Out-of-date offers cause refunds and complaints.", "Weekly"),
    ("Rights", "Music is licensed for business use",
     "The commercial sound library is not the same as the personal one.",
     "Every post"),
    ("Rights", "Images and video are yours or licensed",
     "Stock licences have limits - check before reposting.",
     "Every post"),
    ("Rights", "UGC has written permission",
     "A comment saying 'sure' is enough, a silence is not.",
     "Every repost"),
    ("Rights", "Credit given to creators and collaborators",
     "Tag the maker, photographer or customer.", "Every post"),
    ("Rights", "Licences and rights renewed before they expire",
     "The Assets tab flags expiry dates for you.", "Monthly"),
    ("Privacy", "People in photos have agreed to appear",
     "Especially children, customers and staff.", "Every post"),
    ("Privacy", "No customer details visible in screenshots",
     "Blur names, addresses, order numbers and emails.", "Every post"),
    ("Privacy", "Direct messages kept confidential",
     "Ask before you share a message publicly.", "Every post"),
    ("Privacy", "Email and data consent recorded",
     "Sign-ups from social still need a lawful basis.", "Monthly"),
    ("Competitions", "Giveaway rules written and linked",
     "Entry rules, closing date, prize and who can enter.",
     "Every giveaway"),
    ("Competitions", "Platform promotion rules followed",
     "Each platform has its own competition policy.",
     "Every giveaway"),
    ("Accessibility", "Alt text written for images",
     "Helps screen readers and improves reach.", "Every post"),
    ("Accessibility", "Captions or subtitles on video",
     "Most people watch with the sound off.", "Every post"),
    ("Accessibility", "Readable colour contrast on graphics",
     "Pale text on a pale photo loses people.", "Every post"),
    ("Quality", "Spelling and grammar checked",
     "Read it aloud once before it goes out.", "Every post"),
    ("Quality", "Links tested and tracked",
     "A broken link in bio costs you the whole week.", "Weekly"),
    ("Quality", "Scheduled posts reviewed before they go live",
     "Check nothing has changed since you wrote it.", "Weekly"),
    ("Quality", "Comments and messages answered",
     "Within a day is the standard people expect.", "Daily"),
    ("Brand", "Tone of voice matches the brand guide",
     "The Brand Voice tab is the reference.", "Every post"),
    ("Brand", "Logo, colours and fonts correct",
     "Consistency is what makes a small brand look established.",
     "Every post"),
    ("Brand", "Profile links and bios up to date",
     "Check the link in bio actually goes somewhere useful.",
     "Monthly"),
    ("Risk", "A plan exists for negative comments",
     "Decide in advance what you reply to and what you hide.",
     "Quarterly"),
    ("Risk", "Account recovery and two-factor set up",
     "Losing the account is the real worst case.", "Quarterly"),
]


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("check", '"CHK-"', digits=3, width=9),
        Col("area", "Area", 16, "text"),
        Col("check", "Check", 40, "wrap"),
        Col("why", "Why it matters", 42, "wrap"),
        Col("frequency", "How often", 14, "text",
            validate=list_from("FrequencyList")),
        Col("owner", "Who checks it", 16, "text",
            validate=list_from("TeamNames")),
        Col("status", "Status", 15, "status",
            validate=list_from("ComplianceStatusList"), cf=status_cf()),
        Col("last_checked", "Last checked", 12, "date",
            validate=date_check()),
        Col("evidence", "Evidence / link", 26, "link"),
        Col("action", "Action to take", 30, "wrap"),
        Col("days_since", "Days since", 10, "int",
            formula=blank("last_checked", "TODAY()-{last_checked}{r}")),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "check",
                'IF({status}{r}="Issue found","Action - fix this now",'
                'IF({status}{r}="Not required","OK",'
                'IF(OR({last_checked}{r}="",{status}{r}="Not checked"),'
                '"Check - never checked",IF(AND({frequency}{r}="Monthly",'
                '{days_since}{r}>31),"Due - monthly check",'
                'IF(AND({frequency}{r}="Weekly",{days_since}{r}>7),'
                '"Due - weekly check",IF(AND({frequency}{r}="Quarterly",'
                '{days_since}{r}>92),"Due - quarterly check","OK"))))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Compliance & Quality Checklist",
        subtitle="The boring list that keeps you out of trouble: "
                 "disclosure, music rights, permissions, accessibility "
                 "and the weekly quality checks.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.grey_text,
        freeze_cols=3,
        help_text="Work down the list, set a status and a date. Anything "
                  "due again is flagged in the last column. This is "
                  "guidance, not legal advice - check the rules where "
                  "you live.",
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Write the standard checklist (it is content, not demo data)."""
    spec_ = reg[NAME]
    text = fmts.cell("text", INPUT)
    wrap = fmts.cell("wrap", INPUT)
    for offset, (area, check, why, frequency) in enumerate(CHECKS):
        if offset >= spec_.rows:
            break
        row0 = spec_.first_data_row - 1 + offset
        sheet.write(row0, spec_.index("area"), area, text)
        sheet.write(row0, spec_.index("check"), check, wrap)
        sheet.write(row0, spec_.index("why"), why, wrap)
        sheet.write(row0, spec_.index("frequency"), frequency, text)
````

---

## 5. Measurement registers

Where the numbers go in: monthly account analytics, per post performance, tests, paid promotion, goals and the budget.

### `social_media_spreadsheet/sheets/analytics.py`

`Analytics` - the monthly numbers you copy from each platform. *(170 lines)*

````python
"""`Analytics` - the monthly numbers you copy from each platform.

One row per account per month.  This is the only place account level
numbers are typed; every dashboard, report and growth chart reads them
from here.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, dup_cf, id_col, money_cf,
                     rate_cf)

NAME = "Analytics"

_PREV = ('SUMIFS([[Analytics.%s]],[[Analytics.account]],{account}{r},'
         '[[Analytics.month_key]],TEXT(EDATE({month}{r},-1),"yyyy-mm"))')


def _int(hi: float = 1e9):
    return number_check(0, hi, decimal=False)


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("account", '"AN-"', digits=4, width=9),
        Col("month", "Month", 12, "date", validate=date_check(),
            note="Use the 1st of the month - 01/03/2026 means March 2026."),
        Col("month_key", "Month key", 10, "text",
            formula='=IF({month}{r}="","",TEXT({month}{r},"yyyy-mm"))',
            note="Automatic. Every monthly report groups on this."),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("account", "Account", 20, "text",
            validate=list_from("AccountNames")),
        Col("followers_start", "Followers at start", 13, "int",
            validate=_int()),
        Col("followers_end", "Followers at end", 13, "int",
            validate=_int(), cf=bar_cf()),
        Col("lost", "Unfollows", 10, "int", validate=_int(1e8)),
        Col("posts", "Posts published", 11, "int", validate=_int(10000),
            note="What the platform reports. The grey column next to it "
                 "counts what your calendar says."),
        Col("reach", "Reach", 12, "int", validate=_int(), cf=bar_cf()),
        Col("impressions", "Impressions", 12, "int", validate=_int()),
        Col("engagements", "Engagements", 12, "int", validate=_int(),
            note="Likes + comments + shares + saves, as the platform "
                 "reports it."),
        Col("likes", "Likes", 10, "int", validate=_int()),
        Col("comments", "Comments", 10, "int", validate=_int()),
        Col("shares", "Shares", 10, "int", validate=_int()),
        Col("saves", "Saves", 10, "int", validate=_int()),
        Col("profile_visits", "Profile visits", 12, "int", validate=_int()),
        Col("clicks", "Link clicks", 11, "int", validate=_int()),
        Col("leads", "Leads / sign-ups", 12, "int", validate=_int(1e8)),
        Col("conversions", "Sales", 10, "int", validate=_int(1e8)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("spend", "Spend", 12, "money", validate=number_check(0, 1e9)),
        # ---------------------------------------------------- calculated
        Col("gained", "New followers", 12, "int",
            formula='=IF({followers_end}{r}="","",{followers_end}{r}'
                    '-N({followers_start}{r})+N({lost}{r}))'),
        Col("net_growth", "Net growth", 11, "int",
            formula='=IF({followers_end}{r}="","",{followers_end}{r}'
                    '-N({followers_start}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("growth_rate", "Growth rate", 11, "percent",
            formula='=IF(OR({followers_start}{r}="",'
                    '{followers_start}{r}=0),"",{net_growth}{r}'
                    '/{followers_start}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": 0.02,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("planned", "On your calendar", 12, "int",
            formula='=IF(OR({month}{r}="",{platform}{r}=""),"",'
                    'COUNTIFS([[Content Calendar.platform]],{platform}{r},'
                    '[[Content Calendar.month]],{month_key}{r},'
                    '[[Content Calendar.status]],"Published"))'),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({engagements}{r})/{reach}{r})',
            note="Engagements divided by reach - the number to watch.",
            cf=rate_cf()),
        Col("reach_per_post", "Reach per post", 12, "int",
            formula='=IF(OR({posts}{r}="",{posts}{r}=0),"",'
                    'ROUND(N({reach}{r})/{posts}{r},0))'),
        Col("eng_per_post", "Engagements per post", 13, "int",
            formula='=IF(OR({posts}{r}="",{posts}{r}=0),"",'
                    'ROUND(N({engagements}{r})/{posts}{r},0))'),
        Col("reach_rate", "Reach vs followers", 12, "percent",
            formula='=IF(OR({followers_end}{r}="",{followers_end}{r}=0),'
                    '"",N({reach}{r})/{followers_end}{r})',
            note="Over 100% means you reached beyond your own audience."),
        Col("ctr", "Click rate", 11, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({clicks}{r})/{reach}{r})'),
        Col("lead_rate", "Clicks to leads", 11, "percent",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({leads}{r})/{clicks}{r})'),
        Col("conv_rate", "Leads to sales", 11, "percent",
            formula='=IF(OR({leads}{r}="",{leads}{r}=0),"",'
                    'N({conversions}{r})/{leads}{r})'),
        Col("cpl", "Cost per lead", 11, "money",
            formula='=IF(OR({leads}{r}="",{leads}{r}=0,{spend}{r}=""),"",'
                    '{spend}{r}/{leads}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": "=CplTarget",
                 "fmt": "warn"}]),
        Col("roas", "Return on spend", 12, "number1",
            formula='=IF(OR({spend}{r}="",{spend}{r}=0),"",'
                    'N({revenue}{r})/{spend}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": "=RoasTarget",
                 "fmt": "warn"}]),
        Col("prev_reach", "Reach last month", 12, "int",
            formula='=IF({month}{r}="","",' + (_PREV % "reach") + ")",
            hidden=True),
        Col("reach_change", "Reach vs last month", 13, "percent",
            formula='=IF(OR({prev_reach}{r}="",{prev_reach}{r}=0),"",'
                    'N({reach}{r})/{prev_reach}{r}-1)',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": -0.1,
                 "fmt": "bad"}]),
        Col("prev_followers", "Followers last month", 12, "int",
            formula='=IF({month}{r}="","",'
                    + (_PREV % "followers_end") + ")",
            hidden=True),
        Col("dup_check", "Duplicate check", 13, "status",
            formula='=IF({account}{r}="","",IF(COUNTIFS('
                    '[[Analytics.account]],{account}{r},'
                    '[[Analytics.month_key]],{month_key}{r})>1,'
                    '"Duplicate","OK"))',
            cf=dup_cf()),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "account",
                'IF({dup_check}{r}="Duplicate","Duplicate - this month is '
                'entered twice",IF(AND({reach}{r}<>"",{er}{r}<'
                'TargetEngagement),"Check - engagement below target",'
                'IF({net_growth}{r}<0,"Check - followers went down",'
                'IF(AND({posts}{r}<>"",{planned}{r}<>"",'
                '{posts}{r}<>{planned}{r}),"Check - post count does not '
                'match the calendar","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Monthly Analytics Data Entry",
        subtitle="Once a month, copy the headline numbers from each "
                 "platform into the cream cells. Everything else in the "
                 "workbook - dashboards, growth, reports - is calculated "
                 "from this one table.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=5,
        samples=sample(cfg).analytics,
        help_text="One row per account per month, always dated the 1st. "
                  "Set a reminder for the first working day of each "
                  "month - it takes about ten minutes.",
    )
````

### `social_media_spreadsheet/sheets/performance.py`

`Performance` - the numbers for one individual post. *(171 lines)*

````python
"""`Performance` - the numbers for one individual post.

Type the content ID and the metrics; the title, platform, pillar, campaign
and date are pulled from the Content Calendar, and the rates are worked
out for you.  This is the table the top-content lists, the pillar
comparison and the platform planners all read.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, bar_cf, blank, dup_cf, id_col, lookup_content,
                     money_cf, rate_cf, status_cf)

NAME = "Performance"


def _int(hi: float = 1e9):
    return number_check(0, hi, decimal=False)


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("content", '"PF-"', digits=4, width=9),
        Col("content", "Content ID", 12, "text",
            validate=list_from("ContentIDs"),
            note="Pick the post from the Content Calendar - everything "
                 "grey fills itself in."),
        Col("title", "Content title", 30, "text",
            formula=lookup_content("title")),
        Col("date", "Published", 12, "date",
            formula=lookup_content("date")),
        Col("month", "Month", 10, "text",
            formula=lookup_content("month")),
        Col("platform", "Platform", 13, "text",
            formula=lookup_content("platform")),
        Col("format", "Format", 13, "text",
            formula=lookup_content("format")),
        Col("pillar", "Pillar", 15, "text",
            formula=lookup_content("pillar")),
        Col("campaign", "Campaign", 18, "text",
            formula=lookup_content("campaign")),
        Col("caption", "Caption used", 20, "text",
            formula=lookup_content("caption")),
        Col("persona", "Written for", 16, "text",
            formula=lookup_content("persona")),
        Col("funnel", "Funnel stage", 13, "text",
            formula=lookup_content("funnel")),
        Col("slot", "Time slot", 12, "text",
            formula=lookup_content("slot")),
        Col("weekday", "Day", 9, "text",
            formula=lookup_content("weekday")),
        # ------------------------------------------------- typed numbers
        Col("reach", "Reach", 11, "int", validate=_int(), cf=bar_cf(),
            note="How many people saw it. If the platform only gives you "
                 "impressions, use that and be consistent."),
        Col("impressions", "Impressions", 11, "int", validate=_int()),
        Col("likes", "Likes", 10, "int", validate=_int(1e8)),
        Col("comments", "Comments", 10, "int", validate=_int(1e8)),
        Col("shares", "Shares", 10, "int", validate=_int(1e8)),
        Col("saves", "Saves", 10, "int", validate=_int(1e8)),
        Col("clicks", "Link clicks", 11, "int", validate=_int(1e8)),
        Col("watch", "Avg watch time (sec)", 13, "number1",
            validate=number_check(0, 100000)),
        Col("conversions", "Sales from it", 11, "int",
            validate=_int(1e7)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("target_rate", "Target rate", 11, "percent",
            validate=number_check(0, 1),
            note="Leave blank to use the workbook target from the "
                 "Settings tab."),
        # ---------------------------------------------------- calculated
        Col("engagements", "Engagements", 12, "int",
            formula='=IF({reach}{r}="","",N({likes}{r})+N({comments}{r})'
                    '+N({shares}{r})+N({saves}{r}))'),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("save_rate", "Save rate", 10, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({saves}{r})/{reach}{r})',
            note="Saves are the strongest signal on most platforms."),
        Col("share_rate", "Share rate", 10, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({shares}{r})/{reach}{r})'),
        Col("comment_rate", "Comment rate", 11, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({comments}{r})/{reach}{r})'),
        Col("ctr", "Click rate", 10, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({clicks}{r})/{reach}{r})'),
        Col("conv_rate", "Clicks to sales", 11, "percent",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({conversions}{r})/{clicks}{r})'),
        Col("rev_per_1k", "Revenue per 1,000 reached", 14, "money",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({revenue}{r})/({reach}{r}/1000))'),
        Col("target_used", "Target used", 10, "percent",
            formula='=IF({reach}{r}="","",IF({target_rate}{r}="",'
                    'TargetEngagement,{target_rate}{r}))',
            hidden=True),
        Col("vs_target", "Versus target", 11, "percent",
            formula='=IF(OR({er}{r}="",{target_used}{r}=""),"",'
                    '{er}{r}-{target_used}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "warn"}]),
        Col("rank_key", "Rank key", 10, "number1",
            formula='=IF({er}{r}="","",{er}{r}+ROW()/1000000)',
            hidden=True,
            note="Hidden helper: makes the engagement rate unique so the "
                 "top-post lookups never tie."),
        Col("rank", "Rank by engagement", 12, "int",
            formula='=IF({er}{r}="","",RANK({rank_key}{r},'
                    '[[Performance.rank_key]]))',
            cf=[{"type": "cell", "criteria": "<=", "value": 3,
                 "fmt": "good"}]),
        Col("verdict", "Verdict", 20, "status",
            formula='=IF({er}{r}="","",IF({er}{r}>={target_used}{r}*1.5,'
                    '"Top performer",IF({er}{r}>={target_used}{r},'
                    '"Above target",IF({er}{r}>={target_used}{r}*0.6,'
                    '"Below target","Weak - review it"))))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Top performer", "fmt": "bold_good"},
                {"type": "text", "criteria": "containing",
                 "value": "Above target", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Below target", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Weak", "fmt": "bad"}]),
        Col("dup_check", "Duplicate check", 12, "status",
            formula=blank("content",
                          'IF(COUNTIF([[Performance.content]],'
                          '{content}{r})>1,"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "content",
                'IF({title}{r}="Not found","Check - content ID not on the '
                'calendar",IF({reach}{r}="","Add - no numbers entered '
                'yet",IF({dup_check}{r}="Duplicate","Duplicate - this '
                'post is entered twice",IF({verdict}{r}="Top performer",'
                '"Repurpose this one","OK"))))'),
            cf=alert_cf() + status_cf() + [
                {"type": "text", "criteria": "containing",
                 "value": "Repurpose", "fmt": "info"}]),
        Col("notes", "What you learned", 28, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Content Performance Tracker",
        subtitle="Post by post results. Fill in the cream columns a few "
                 "days after publishing and the workbook ranks every "
                 "post, finds your best pillar and tells you what to "
                 "make more of.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).performance,
        help_text="Wait 48-72 hours before recording numbers. Be "
                  "consistent about whether you use reach or impressions "
                  "- the rates only compare if you are.",
    )
````

### `social_media_spreadsheet/sheets/ab_tests.py`

`A/B Tests` - change one thing, measure it, keep the winner. *(112 lines)*

````python
"""`A/B Tests` - change one thing, measure it, keep the winner."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, status_cf

NAME = "AB Tests"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"AB-"', digits=3, width=9),
        Col("name", "Test name", 24, "text"),
        Col("variable", "What you changed", 16, "text",
            validate=list_from("TestVariableList"),
            note="Change one thing only, or you will not know what "
                 "caused the difference."),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("hypothesis", "Hypothesis", 34, "wrap",
            note="'If I ... then ... because ...'"),
        Col("version_a", "Version A", 26, "wrap"),
        Col("version_b", "Version B", 26, "wrap"),
        Col("audience", "Kept the same", 24, "wrap"),
        Col("metric", "Main measure", 16, "text"),
        Col("secondary", "Second measure", 15, "text"),
        Col("start", "Started", 12, "date", validate=date_check()),
        Col("end", "Finished", 12, "date", validate=date_check()),
        Col("a_result", "A result", 11, "percent",
            validate=number_check(0, 10)),
        Col("b_result", "B result", 11, "percent",
            validate=number_check(0, 10)),
        Col("a_clicks", "A clicks", 10, "int",
            validate=number_check(0, 1e8, decimal=False)),
        Col("b_clicks", "B clicks", 10, "int",
            validate=number_check(0, 1e8, decimal=False)),
        Col("a_conversions", "A sales", 9, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("b_conversions", "B sales", 9, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("status", "Status", 13, "status",
            validate=list_from("TestStatusList"), cf=status_cf()),
        Col("difference", "Difference", 11, "percent",
            formula='=IF(OR({a_result}{r}="",{b_result}{r}=""),"",'
                    '{b_result}{r}-{a_result}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("uplift", "Uplift", 11, "percent",
            formula='=IF(OR({a_result}{r}="",{a_result}{r}=0,'
                    '{b_result}{r}=""),"",{b_result}{r}/{a_result}{r}-1)'),
        Col("winner", "Winner", 11, "text",
            formula='=IF(OR({a_result}{r}="",{b_result}{r}=""),"",'
                    'IF(ABS({uplift}{r})<0.1,"Too close",'
                    'IF({b_result}{r}>{a_result}{r},"B","A")))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Too close", "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "B",
                 "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "A",
                 "fmt": "good"}]),
        Col("sample_size", "Total clicks", 11, "int",
            formula='=IF(AND({a_clicks}{r}="",{b_clicks}{r}=""),"",'
                    'N({a_clicks}{r})+N({b_clicks}{r}))'),
        Col("trust", "How much to trust it", 22, "status",
            formula='=IF({winner}{r}="","",IF({sample_size}{r}<100,'
                    '"Low - too small a sample",IF(ABS({uplift}{r})<0.1,'
                    '"Low - the gap is noise",IF({sample_size}{r}<400,'
                    '"Fair - worth repeating","Good - act on it"))))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Good", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Fair", "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "Low",
                 "fmt": "bad"}]),
        Col("days", "Days run", 9, "int",
            formula='=IF(OR({start}{r}="",{end}{r}=""),"",'
                    '{end}{r}-{start}{r}+1)'),
        Col("confidence", "Your read on it", 26, "wrap"),
        Col("next_test", "Next test to run", 26, "wrap"),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "name",
                'IF(AND({status}{r}="Running",{end}{r}<>"",'
                '{end}{r}<TODAY()),"Due - test finished, record the '
                'result",IF(AND({status}{r}="Complete",'
                '{next_test}{r}=""),"Add - what will you test next?",'
                'IF({trust}{r}="Low - too small a sample",'
                '"Check - run it for longer","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="A/B Testing Tracker",
        subtitle="Two versions, one difference, one measure. Small "
                 "samples lie, so the sheet tells you how much to trust "
                 "each result.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).tests,
        help_text="Test one variable at a time - hook, cover, posting "
                  "time or call to action - and run it long enough to "
                  "collect a few hundred interactions.",
    )
````

### `social_media_spreadsheet/sheets/ads.py`

`Ads` - paid promotion, what it cost and what it brought back. *(124 lines)*

````python
"""`Ads` - paid promotion, what it cost and what it brought back."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, status_cf)

NAME = "Ads"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"AD-"', digits=3, width=9),
        Col("name", "Ad / ad set name", 24, "text"),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("objective", "Objective", 15, "text",
            validate=list_from("AdObjectiveList")),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("audience", "Audience", 28, "wrap"),
        Col("segment", "Targeting", 24, "wrap"),
        Col("creative", "Creative used", 13, "text",
            validate=list_from("AssetIDs")),
        Col("copy", "Ad copy", 32, "wrap"),
        Col("headline", "Headline", 26, "wrap"),
        Col("landing", "Landing page", 26, "link"),
        Col("start", "Start", 12, "date", validate=date_check()),
        Col("end", "End", 12, "date", validate=date_check()),
        Col("budget", "Budget", 12, "money", validate=number_check(0, 1e9),
            cf=bar_cf()),
        Col("spend", "Spent", 12, "money", validate=number_check(0, 1e9)),
        Col("impressions", "Impressions", 12, "int",
            validate=number_check(0, 1e10, decimal=False)),
        Col("reach", "Reach", 11, "int",
            validate=number_check(0, 1e10, decimal=False)),
        Col("clicks", "Clicks", 10, "int",
            validate=number_check(0, 1e9, decimal=False)),
        Col("conversions", "Sales", 10, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("status", "Status", 12, "status",
            formula='=IF({name}{r}="","",IF({start}{r}="","Planned",'
                    'IF(AND({start}{r}<=TODAY(),OR({end}{r}="",'
                    '{end}{r}>=TODAY())),"Running",IF({start}{r}>TODAY(),'
                    '"Planned","Complete"))))',
            cf=status_cf()),
        Col("days", "Days live", 9, "int",
            formula='=IF({start}{r}="","",IF({end}{r}="",'
                    'TODAY()-{start}{r}+1,MIN(TODAY(),{end}{r})'
                    '-{start}{r}+1))'),
        Col("daily_spend", "Spend per day", 11, "money",
            formula='=IF(OR({days}{r}="",{days}{r}<=0),"",'
                    'N({spend}{r})/{days}{r})'),
        Col("ctr", "Click rate", 10, "percent",
            formula='=IF(OR({impressions}{r}="",{impressions}{r}=0),"",'
                    'N({clicks}{r})/{impressions}{r})',
            note="1% and up is healthy on most platforms.",
            cf=[{"type": "cell", "criteria": ">=", "value": 0.01,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.005,
                 "fmt": "warn"}]),
        Col("cpc", "Cost per click", 11, "money",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({spend}{r})/{clicks}{r})'),
        Col("cpm", "Cost per 1,000 shown", 12, "money",
            formula='=IF(OR({impressions}{r}="",{impressions}{r}=0),"",'
                    'N({spend}{r})/({impressions}{r}/1000))'),
        Col("conv_rate", "Clicks to sales", 11, "percent",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({conversions}{r})/{clicks}{r})'),
        Col("cpa", "Cost per sale", 11, "money",
            formula='=IF(OR({conversions}{r}="",{conversions}{r}=0),"",'
                    'N({spend}{r})/{conversions}{r})'),
        Col("roas", "Return on spend", 12, "number1",
            formula='=IF(OR({spend}{r}="",{spend}{r}=0),"",'
                    'N({revenue}{r})/{spend}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "bad"}]),
        Col("profit", "Revenue less spend", 13, "money",
            formula='=IF(AND({revenue}{r}="",{spend}{r}=""),"",'
                    'N({revenue}{r})-N({spend}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("budget_used", "Budget used", 11, "percent",
            formula='=IF(OR({budget}{r}="",{budget}{r}=0),"",'
                    'N({spend}{r})/{budget}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bad"},
                {"type": "data_bar", "bar_color": "#F2665E"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(N({spend}{r})>N({budget}{r}),"Over budget - pause or '
                'top up",IF(AND({roas}{r}<>"",{roas}{r}<1),'
                '"Action - losing money, change it",'
                'IF(AND({ctr}{r}<>"",{ctr}{r}<0.005),'
                '"Check - the creative is not landing",'
                'IF(AND({status}{r}="Complete",{revenue}{r}=""),'
                '"Add - record the revenue","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Advertising Planner & Tracker",
        subtitle="Boosted posts and proper ad sets. Keep the spend, the "
                 "creative and the result in one row so you can see "
                 "which ads actually pay.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).ads,
        help_text="Boost what already worked organically - check the "
                  "Performance tab for your top posts first.",
    )
````

### `social_media_spreadsheet/sheets/goals.py`

`Goals` - the handful of numbers you are actually trying to move. *(116 lines)*

````python
"""`Goals` - the handful of numbers you are actually trying to move."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, status_cf

NAME = "Goals"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"GOAL-"', digits=2, width=10),
        Col("name", "Goal", 26, "text",
            note="Write it as a number you can check: 'Instagram "
                 "followers to 5,000'."),
        Col("category", "Category", 16, "text",
            validate=list_from("GoalCategoryList")),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("description", "What success looks like", 34, "wrap"),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("campaign", "Linked campaign", 18, "text",
            validate=list_from("CampaignNames")),
        Col("plan", "How you will get there", 26, "wrap"),
        Col("baseline", "Starting point", 12, "number",
            validate=number_check(-1e12, 1e12)),
        Col("current", "Where you are now", 13, "number",
            validate=number_check(-1e12, 1e12),
            note="Update this once a month when you fill in the "
                 "Analytics tab."),
        Col("target", "Target", 12, "number",
            validate=number_check(-1e12, 1e12)),
        Col("target_date", "By when", 12, "date", validate=date_check()),
        Col("milestone_month", "Monthly step", 12, "number",
            validate=number_check(-1e12, 1e12)),
        Col("milestone_quarter", "Quarterly step", 12, "number",
            validate=number_check(-1e12, 1e12)),
        Col("budget", "Budget for it", 12, "money",
            validate=number_check(0, 1e9)),
        Col("status", "Status", 13, "status",
            validate=list_from("GoalStatusList"), cf=status_cf()),
        Col("to_go", "Still to go", 12, "number",
            formula='=IF(OR({target}{r}="",{current}{r}=""),"",'
                    '{target}{r}-{current}{r})'),
        Col("progress", "Progress", 11, "percent",
            formula='=IF(OR({target}{r}="",{baseline}{r}="",'
                    '{target}{r}={baseline}{r}),"",MAX(0,'
                    '({current}{r}-{baseline}{r})'
                    '/({target}{r}-{baseline}{r})))',
            cf=[{"type": "data_bar", "bar_color": "#6350A8"},
                {"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"}]),
        Col("days_left", "Days left", 10, "int",
            formula=blank("target_date", "{target_date}{r}-TODAY()"),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "muted"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 30, "fmt": "warn"}]),
        Col("expected", "Where you should be", 13, "percent",
            formula='=IF(OR({target_date}{r}="",{start_ref}{r}="",'
                    '{target_date}{r}<={start_ref}{r}),"",'
                    'MIN(1,MAX(0,(TODAY()-{start_ref}{r})'
                    '/({target_date}{r}-{start_ref}{r}))))',
            note="How far through the time you are - compare it with "
                 "the progress column."),
        Col("start_ref", "Year start", 10, "date",
            formula='=IF({name}{r}="","",DATE(PlanYear,1,1))',
            hidden=True),
        Col("pace", "On track?", 20, "status",
            formula='=IF(OR({progress}{r}="",{expected}{r}=""),"",'
                    'IF({progress}{r}>=1,"Achieved",'
                    'IF({progress}{r}>={expected}{r},"On track",'
                    'IF({progress}{r}>={expected}{r}*0.8,"Slightly '
                    'behind","Behind - needs a push"))))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Achieved", "fmt": "bold_good"},
                {"type": "text", "criteria": "containing",
                 "value": "On track", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Slightly", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Behind", "fmt": "bad"}]),
        Col("needed_per_month", "Needed each month", 13, "number1",
            formula='=IF(OR({to_go}{r}="",{days_left}{r}="",'
                    '{days_left}{r}<=0),"",{to_go}{r}'
                    '/MAX(1,{days_left}{r}/30.4))'),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(AND({days_left}{r}<>"",{days_left}{r}<0,'
                '{progress}{r}<1),"Overdue - the date has passed",'
                'IF({pace}{r}="Behind - needs a push","Action - behind '
                'the pace you need",IF({current}{r}="","Add - no current '
                'number yet","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Goals & KPI Tracker",
        subtitle="Four or five goals, each with a number, a date and an "
                 "owner. The sheet compares how far you have come with "
                 "how far through the year you are.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).goals,
        help_text="Update the 'where you are now' column once a month. "
                  "Everything else works itself out.",
    )
````

### `social_media_spreadsheet/sheets/budget.py`

`Budget` - what social media costs you, month by month. *(90 lines)*

````python
"""`Budget` - what social media costs you, month by month."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, month_col)

NAME = "Budget"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("category", '"BU-"', digits=3, width=9),
        Col("month", "Month", 12, "date", validate=date_check(),
            note="Use the 1st of the month."),
        month_col("month", key="month_key", header="Month key"),
        Col("category", "Cost category", 20, "text",
            validate=list_from("CostCategoryList")),
        Col("platform", "Platform", 15, "text"),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("planned", "Planned", 12, "money",
            validate=number_check(0, 1e9), cf=bar_cf()),
        Col("actual", "Actual", 12, "money",
            validate=number_check(0, 1e9)),
        Col("variance", "Difference", 12, "money",
            formula='=IF(AND({planned}{r}="",{actual}{r}=""),"",'
                    'N({planned}{r})-N({actual}{r}))',
            note="Positive means you came in under budget.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"}]),
        Col("variance_pct", "Difference %", 11, "percent",
            formula='=IF(OR({planned}{r}="",{planned}{r}=0),"",'
                    '{variance}{r}/{planned}{r})'),
        Col("share", "Share of the month", 12, "percent",
            formula='=IF(OR({actual}{r}="",{month_key}{r}=""),"",'
                    'IFERROR({actual}{r}/SUMIFS([[Budget.actual]],'
                    '[[Budget.month_key]],{month_key}{r}),""))'),
        Col("revenue", "Revenue credited", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf(),
            note="Only if you can tie revenue to this line - leave it "
                 "blank otherwise."),
        Col("roi", "Return", 11, "number1",
            formula='=IF(OR({actual}{r}="",{actual}{r}=0,'
                    '{revenue}{r}=""),"",{revenue}{r}/{actual}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "bad"}]),
        Col("month_total", "Spent that month", 13, "money",
            formula=blank("month_key",
                          'SUMIFS([[Budget.actual]],'
                          '[[Budget.month_key]],{month_key}{r})')),
        Col("vs_budget", "Versus monthly budget", 13, "percent",
            formula='=IF(OR({month_total}{r}="",MonthlyBudget=0),"",'
                    '{month_total}{r}/MonthlyBudget)',
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bad"},
                {"type": "data_bar", "bar_color": "#F2665E"}]),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "category",
                'IF({vs_budget}{r}>1,"Over budget this month",'
                'IF(AND({actual}{r}<>"",{planned}{r}<>"",'
                '{actual}{r}>{planned}{r}*1.2),"Check - 20% over plan",'
                'IF(AND({roi}{r}<>"",{roi}{r}<1),"Review - costs more '
                'than it returns","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Marketing Budget & Return",
        subtitle="Ads, tools, props, editing and freelancers. One row "
                 "per cost per month - the dashboard totals it and "
                 "compares it with the revenue you recorded.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).budget,
        help_text="Set your monthly budget on the Settings tab. Anything "
                  "over it is flagged here and on the dashboard.",
    )
````

---

## 6. The pages

The tabs that read the registers: the hidden calculation engine, the two dashboards, the calendar grid, the production pipeline, the brand voice page, the reports, the welcome tab, the printables and the help tab.

### `social_media_spreadsheet/sheets/calc.py`

`Calc` - hidden helper sheet. *(336 lines)*

````python
"""
`Calc` - hidden helper sheet.

The 2007-2016 formula set has no SORT or FILTER, so the twelve month trend
series, the platform and pillar breakdowns and the "top ten" tables are
pre-computed here with plain SUMIFS / COUNTIFS and then pulled into the
dashboards and reports with LARGE + INDEX/MATCH.

Nothing here needs touching: the tab is hidden and every cell is a locked
formula.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from . import lists

NAME = "Calc"

# Dashboard period window (DashStart / DashEnd are defined on the Dashboard
# tab - one month, or the whole year).
_DASH = ('[[Content Calendar.date]],">="&DashStart,'
         '[[Content Calendar.date]],"<="&DashEnd')
_DASH_AN = ('[[Analytics.month]],">="&DashStart,'
            '[[Analytics.month]],"<="&DashEnd')
#: year window on the Content Calendar
_YEAR = ('[[Content Calendar.date]],">="&DATE(PlanYear,1,1),'
         '[[Content Calendar.date]],"<="&DATE(PlanYear,12,31)')


def _mirror(sheet: str, key: str) -> str:
    return f'=IF([[{sheet}.{key}#]]="","",[[{sheet}.{key}#]])'


def _from_list(name: str) -> str:
    ref = lists.ref(name)
    return f'=IF({ref}="","",{ref})'


def spec(cfg: Config) -> SheetSpec:
    rows = max(cfg.row_count("Performance"),
               cfg.row_count("Content Calendar"),
               cfg.row_count("Analytics"),
               cfg.row_count("Campaigns"), 60)

    columns = [
        # ------------------------------------------- 12 month series ------
        Col("m_start", "Month start", 12, "date",
            formula='=IF(ROW()-4>12,"",DATE(PlanYear,ROW()-4,1))'),
        Col("m_end", "Month end", 12, "date",
            formula='=IF({m_start}{r}="","",EOMONTH({m_start}{r},0))'),
        Col("m_key", "Month key", 10, "text",
            formula='=IF({m_start}{r}="","",TEXT({m_start}{r},"yyyy-mm"))'),
        Col("m_label", "Month", 9, "text",
            formula='=IF({m_start}{r}="","",TEXT({m_start}{r},"mmm"))'),
        Col("m_planned", "Posts planned", 11, "int",
            formula='=IF({m_key}{r}="","",COUNTIFS('
                    '[[Content Calendar.month]],{m_key}{r}))'),
        Col("m_published", "Published", 10, "int",
            formula='=IF({m_key}{r}="","",COUNTIFS('
                    '[[Content Calendar.month]],{m_key}{r},'
                    '[[Content Calendar.status]],"Published"))'),
        Col("m_reach", "Reach", 12, "int",
            formula='=IF({m_key}{r}="","",SUMIFS([[Analytics.reach]],'
                    '[[Analytics.month_key]],{m_key}{r}))'),
        Col("m_impressions", "Impressions", 12, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.impressions]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_engagements", "Engagements", 12, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.engagements]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_er", "Engagement rate", 12, "percent",
            formula='=IF(OR({m_reach}{r}="",{m_reach}{r}=0),"",'
                    '{m_engagements}{r}/{m_reach}{r})'),
        Col("m_followers", "Followers", 11, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.followers_end]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_growth", "Follower growth", 11, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.net_growth]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_clicks", "Clicks", 10, "int",
            formula='=IF({m_key}{r}="","",SUMIFS([[Analytics.clicks]],'
                    '[[Analytics.month_key]],{m_key}{r}))'),
        Col("m_leads", "Leads", 10, "int",
            formula='=IF({m_key}{r}="","",SUMIFS([[Analytics.leads]],'
                    '[[Analytics.month_key]],{m_key}{r}))'),
        Col("m_conversions", "Sales", 10, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.conversions]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_revenue", "Revenue", 13, "money",
            formula='=IF({m_key}{r}="","",SUMIFS([[Analytics.revenue]],'
                    '[[Analytics.month_key]],{m_key}{r}))'),
        Col("m_spend", "Spend", 12, "money",
            formula='=IF({m_key}{r}="","",SUMIFS([[Budget.actual]],'
                    '[[Budget.month_key]],{m_key}{r}))'),
        Col("m_profit", "Revenue less spend", 13, "money",
            formula='=IF({m_key}{r}="","",{m_revenue}{r}-{m_spend}{r})'),
        Col("m_target", "Target posts", 10, "int",
            formula='=IF({m_key}{r}="","",MonthlyPostTarget)'),
        # --------------------------------------------- platform block -----
        Col("p_name", "Platform", 14, "text",
            formula=_from_list("PlatformList")),
        Col("p_posts_m", "Posts (month)", 11, "int",
            formula='=IF({p_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.platform]],{p_name}{r},'
                    + _DASH + "))"),
        Col("p_published_m", "Published (month)", 11, "int",
            formula='=IF({p_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.platform]],{p_name}{r},'
                    '[[Content Calendar.status]],"Published",'
                    + _DASH + "))"),
        Col("p_reach_m", "Reach (month)", 12, "int",
            formula='=IF({p_name}{r}="","",SUMIFS([[Analytics.reach]],'
                    '[[Analytics.platform]],{p_name}{r},'
                    + _DASH_AN + "))"),
        Col("p_eng_m", "Engagements (month)", 12, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.engagements]],[[Analytics.platform]],'
                    '{p_name}{r},' + _DASH_AN + "))"),
        Col("p_er_m", "Engagement rate (month)", 12, "percent",
            formula='=IF(OR({p_reach_m}{r}="",{p_reach_m}{r}=0),"",'
                    '{p_eng_m}{r}/{p_reach_m}{r})'),
        Col("p_followers", "Followers (month)", 11, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.followers_end]],[[Analytics.platform]],'
                    '{p_name}{r},' + _DASH_AN + "))"),
        Col("p_growth_m", "Growth (month)", 11, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.net_growth]],[[Analytics.platform]],'
                    '{p_name}{r},' + _DASH_AN + "))"),
        Col("p_clicks_m", "Clicks (month)", 10, "int",
            formula='=IF({p_name}{r}="","",SUMIFS([[Analytics.clicks]],'
                    '[[Analytics.platform]],{p_name}{r},'
                    + _DASH_AN + "))"),
        Col("p_revenue_m", "Revenue (month)", 13, "money",
            formula='=IF({p_name}{r}="","",SUMIFS([[Analytics.revenue]],'
                    '[[Analytics.platform]],{p_name}{r},'
                    + _DASH_AN + "))"),
        Col("p_posts_y", "Posts (year)", 11, "int",
            formula='=IF({p_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.platform]],{p_name}{r},'
                    + _YEAR + "))"),
        Col("p_reach_y", "Reach (year)", 12, "int",
            formula='=IF({p_name}{r}="","",SUMIFS([[Analytics.reach]],'
                    '[[Analytics.platform]],{p_name}{r}))'),
        Col("p_eng_y", "Engagements (year)", 12, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.engagements]],[[Analytics.platform]],'
                    '{p_name}{r}))'),
        Col("p_er_y", "Engagement rate (year)", 12, "percent",
            formula='=IF(OR({p_reach_y}{r}="",{p_reach_y}{r}=0),"",'
                    '{p_eng_y}{r}/{p_reach_y}{r})'),
        Col("p_growth_y", "Growth (year)", 11, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.net_growth]],[[Analytics.platform]],'
                    '{p_name}{r}))'),
        Col("p_key", "Platform rank key", 12, "number",
            formula='=IF({p_name}{r}="","",IF({p_reach_y}{r}=0,"",'
                    '{p_reach_y}{r}+ROW()/1000000))'),
        # ------------------------------------------------ pillar block ----
        Col("pl_name", "Pillar", 16, "text",
            formula=_mirror("Content Pillars", "name")),
        Col("pl_posts", "Pillar posts", 11, "int",
            formula='=IF({pl_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.pillar]],{pl_name}{r},'
                    + _YEAR + "))"),
        Col("pl_reach", "Pillar reach", 12, "int",
            formula='=IF({pl_name}{r}="","",SUMIFS([[Performance.reach]],'
                    '[[Performance.pillar]],{pl_name}{r}))'),
        Col("pl_eng", "Pillar engagements", 12, "int",
            formula='=IF({pl_name}{r}="","",SUMIFS('
                    '[[Performance.engagements]],[[Performance.pillar]],'
                    '{pl_name}{r}))'),
        Col("pl_er", "Pillar engagement rate", 12, "percent",
            formula='=IF(OR({pl_reach}{r}="",{pl_reach}{r}=0),"",'
                    '{pl_eng}{r}/{pl_reach}{r})'),
        Col("pl_key", "Pillar rank key", 12, "number",
            formula='=IF({pl_er}{r}="","",{pl_er}{r}+ROW()/1000000)'),
        # ------------------------------------------------ format block ----
        Col("fm_name", "Format", 14, "text",
            formula=_from_list("FormatList")),
        Col("fm_posts", "Format posts", 11, "int",
            formula='=IF({fm_name}{r}="","",COUNTIF('
                    '[[Performance.format]],{fm_name}{r}))'),
        Col("fm_reach", "Format reach", 12, "int",
            formula='=IF({fm_name}{r}="","",SUMIFS([[Performance.reach]],'
                    '[[Performance.format]],{fm_name}{r}))'),
        Col("fm_eng", "Format engagements", 12, "int",
            formula='=IF({fm_name}{r}="","",SUMIFS('
                    '[[Performance.engagements]],[[Performance.format]],'
                    '{fm_name}{r}))'),
        Col("fm_er", "Format engagement rate", 12, "percent",
            formula='=IF(OR({fm_reach}{r}="",{fm_reach}{r}=0),"",'
                    '{fm_eng}{r}/{fm_reach}{r})'),
        Col("fm_key", "Format rank key", 12, "number",
            formula='=IF({fm_posts}{r}=0,"",IF({fm_er}{r}="","",'
                    '{fm_er}{r}+ROW()/1000000))'),
        # ------------------------------------------- day and slot block ---
        Col("dw_day", "Weekday", 9, "text",
            formula='=IF(ROW()-4>7,"",TEXT(DATE(2024,1,ROW()-4),"ddd"))'),
        Col("dw_posts", "Posts that day", 11, "int",
            formula='=IF({dw_day}{r}="","",COUNTIF('
                    '[[Performance.weekday]],{dw_day}{r}))'),
        Col("dw_reach", "Reach that day", 12, "int",
            formula='=IF({dw_day}{r}="","",SUMIFS([[Performance.reach]],'
                    '[[Performance.weekday]],{dw_day}{r}))'),
        Col("dw_eng", "Engagements that day", 12, "int",
            formula='=IF({dw_day}{r}="","",SUMIFS('
                    '[[Performance.engagements]],'
                    '[[Performance.weekday]],{dw_day}{r}))'),
        Col("dw_er", "Engagement rate that day", 12, "percent",
            formula='=IF(OR({dw_reach}{r}="",{dw_reach}{r}=0),"",'
                    '{dw_eng}{r}/{dw_reach}{r})'),
        Col("dw_key", "Weekday rank key", 12, "number",
            formula='=IF({dw_posts}{r}=0,"",IF({dw_er}{r}="","",'
                    '{dw_er}{r}+ROW()/1000000))'),
        Col("sl_slot", "Time slot", 13, "text",
            formula=_from_list("TimeSlotList")),
        Col("sl_posts", "Posts in slot", 11, "int",
            formula='=IF({sl_slot}{r}="","",COUNTIF('
                    '[[Performance.slot]],{sl_slot}{r}))'),
        Col("sl_reach", "Reach in slot", 12, "int",
            formula='=IF({sl_slot}{r}="","",SUMIFS([[Performance.reach]],'
                    '[[Performance.slot]],{sl_slot}{r}))'),
        Col("sl_eng", "Engagements in slot", 12, "int",
            formula='=IF({sl_slot}{r}="","",SUMIFS('
                    '[[Performance.engagements]],[[Performance.slot]],'
                    '{sl_slot}{r}))'),
        Col("sl_er", "Engagement rate in slot", 12, "percent",
            formula='=IF(OR({sl_reach}{r}="",{sl_reach}{r}=0),"",'
                    '{sl_eng}{r}/{sl_reach}{r})'),
        Col("sl_key", "Slot rank key", 12, "number",
            formula='=IF({sl_posts}{r}=0,"",IF({sl_er}{r}="","",'
                    '{sl_er}{r}+ROW()/1000000))'),
        # ------------------------------------------- performance mirror ---
        Col("pf_title", "Post", 28, "text",
            formula=_mirror("Performance", "title")),
        Col("pf_platform", "Post platform", 13, "text",
            formula=_mirror("Performance", "platform")),
        Col("pf_pillar", "Post pillar", 14, "text",
            formula=_mirror("Performance", "pillar")),
        Col("pf_date", "Post date", 12, "date",
            formula=_mirror("Performance", "date")),
        Col("pf_month", "Post month", 10, "text",
            formula=_mirror("Performance", "month")),
        Col("pf_reach", "Post reach", 11, "int",
            formula=_mirror("Performance", "reach")),
        Col("pf_er", "Post engagement rate", 12, "percent",
            formula=_mirror("Performance", "er")),
        Col("pf_key_m", "Top post key (period)", 13, "number",
            formula='=IF({pf_title}{r}="","",IF({pf_date}{r}="","",'
                    'IF(OR({pf_date}{r}<DashStart,{pf_date}{r}>DashEnd),'
                    '"",IF({pf_er}{r}="","",{pf_er}{r}+ROW()/1000000))))'),
        Col("pf_reach_key_m", "Top reach key (period)", 13, "number",
            formula='=IF({pf_reach}{r}="","",IF({pf_date}{r}="","",'
                    'IF(OR({pf_date}{r}<DashStart,{pf_date}{r}>DashEnd),'
                    '"",{pf_reach}{r}+ROW()/1000000)))'),
        Col("pf_key_y", "Top post key (year)", 13, "number",
            formula='=IF({pf_title}{r}="","",IF({pf_date}{r}="","",'
                    'IF(YEAR({pf_date}{r})<>PlanYear,"",'
                    'IF({pf_er}{r}="","",{pf_er}{r}+ROW()/1000000))))'),
        Col("pf_key_r", "Top post key (report range)", 13, "number",
            formula='=IF({pf_title}{r}="","",IF({pf_date}{r}="","",'
                    'IF(OR({pf_date}{r}<RepFrom,{pf_date}{r}>RepTo),"",'
                    'IF({pf_er}{r}="","",{pf_er}{r}+ROW()/1000000))))'),
        Col("pf_reach_key", "Top reach key (year)", 13, "number",
            formula='=IF({pf_reach}{r}="","",IF({pf_date}{r}="","",'
                    'IF(YEAR({pf_date}{r})<>PlanYear,"",'
                    '{pf_reach}{r}+ROW()/1000000)))'),
        # --------------------------------------------- workflow block -----
        Col("ws_status", "Workflow status", 15, "text",
            formula=_from_list("ContentStatusList")),
        Col("ws_count", "In this status", 11, "int",
            formula='=IF({ws_status}{r}="","",COUNTIF('
                    '[[Content Calendar.status]],{ws_status}{r}))'),
        Col("ws_month", "In status (month)", 11, "int",
            formula='=IF({ws_status}{r}="","",COUNTIFS('
                    '[[Content Calendar.status]],{ws_status}{r},'
                    + _DASH + "))"),
        Col("ow_name", "Owner", 16, "text", formula=_mirror("Team", "name")),
        Col("ow_posts", "Owner posts", 11, "int",
            formula='=IF({ow_name}{r}="","",COUNTIF('
                    '[[Content Calendar.owner]],{ow_name}{r}))'),
        Col("ow_open", "Owner in progress", 11, "int",
            formula='=IF({ow_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.owner]],{ow_name}{r},'
                    '[[Content Calendar.status]],"<>Published"))'),
        Col("ow_tasks", "Owner open tasks", 11, "int",
            formula='=IF({ow_name}{r}="","",COUNTIFS([[Tasks.owner]],'
                    '{ow_name}{r},[[Tasks.status]],"<>Complete"))'),
        # --------------------------------------------- campaign block -----
        Col("cm_name", "Campaign", 20, "text",
            formula=_mirror("Campaigns", "name")),
        Col("cm_revenue", "Campaign revenue", 13, "money",
            formula=_mirror("Campaigns", "revenue")),
        Col("cm_spend", "Campaign spend", 12, "money",
            formula=_mirror("Campaigns", "spend")),
        Col("cm_reach", "Campaign reach", 12, "int",
            formula=_mirror("Campaigns", "reach")),
        Col("cm_key", "Campaign rank key", 12, "number",
            formula='=IF({cm_name}{r}="","",IF(N({cm_revenue}{r})<=0,"",'
                    'N({cm_revenue}{r})+ROW()/1000000))'),
        # ------------------------------------------------ account block ---
        Col("ac_name", "Account", 20, "text",
            formula=_mirror("Accounts", "account")),
        Col("ac_platform", "Account platform", 13, "text",
            formula=_mirror("Accounts", "platform")),
        Col("ac_followers", "Account followers", 12, "int",
            formula=_mirror("Accounts", "followers")),
        Col("ac_growth", "Account growth (month)", 12, "int",
            formula='=IF({ac_name}{r}="","",SUMIFS('
                    '[[Analytics.net_growth]],[[Analytics.account]],'
                    '{ac_name}{r},' + _DASH_AN + "))"),
        Col("ac_key", "Account rank key", 12, "number",
            formula='=IF({ac_followers}{r}="","",IF({ac_followers}{r}=0,'
                    '"",{ac_followers}{r}+ROW()/1000000))'),
    ]

    return SheetSpec(
        name=NAME,
        title="Calculation engine (do not edit)",
        subtitle="Helper tables for the dashboards and reports. This tab "
                 "is hidden by default - unhide it only if you want to "
                 "check the maths.",
        columns=columns,
        rows=rows,
        tab_color="#9AA7AD",
        freeze_cols=1,
        help_text="Every cell on this tab is a locked formula.",
    )
````

### `social_media_spreadsheet/sheets/dashboard.py`

`Dashboard` - the one page view of the month. *(375 lines)*

````python
"""`Dashboard` - the one page view of the month."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..writer import SheetWriter, list_from

NAME = "Dashboard"
LAST_COL = 17            # column R - six KPI tiles of three columns each
FIRST_CALC_ROW = 5       # Calc sheet data starts on row 5
MONTH_ROWS = 12

_WINDOW = ('[[Content Calendar.date]],">="&DashStart,'
           '[[Content Calendar.date]],"<="&DashEnd')
_AN = ('[[Analytics.month]],">="&DashStart,'
       '[[Analytics.month]],"<="&DashEnd')


def _calc(key: str, row: int) -> str:
    return f"[[Calc.{key}#{row}]]"


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    ws = s.ws

    s.banner("Monthly Dashboard",
             "Pick a month, then read your whole social media presence at "
             "a glance. Every number is live - it updates as you fill in "
             "the calendar and the Analytics tab.",
             LAST_COL)

    for c in range(0, LAST_COL + 1):
        s.set_col(c, c, 12.5)

    # ------------------------------------------------------ period picker
    row = 4
    row = s.section(row, "Which month are you looking at?", LAST_COL)
    s.merge(row, 0, 1, "Month", fmts.chrome("label_bold"))
    s.merge(row, 2, 3, cfg.today.strftime("%B"), fmts.input("text"))
    ws.data_validation(row, 2, row, 2, list_from("MonthList"))
    book.define_name("DashMonthName", f"={quote_sheet(NAME)}!$C${row + 1}")

    s.merge(row, 4, 5, "Period start", fmts.chrome("label_bold"))
    s.merge(row, 6, 7, None, fmts.calc("date"))
    ws.write_formula(
        row, 6,
        '=IF(DashMonthName="All year",DATE(PlanYear,1,1),'
        'DATE(PlanYear,MAX(1,MATCH(DashMonthName,MonthList,0)-1),1))',
        fmts.calc("date"))
    book.define_name("DashStart", f"={quote_sheet(NAME)}!$G${row + 1}")

    s.merge(row, 8, 9, "Period end", fmts.chrome("label_bold"))
    s.merge(row, 10, 11, None, fmts.calc("date"))
    ws.write_formula(
        row, 10,
        '=IF(DashMonthName="All year",DATE(PlanYear,12,31),'
        'EOMONTH(DashStart,0))', fmts.calc("date"))
    book.define_name("DashEnd", f"={quote_sheet(NAME)}!$K${row + 1}")

    s.merge(row, 12, 13, "Month key", fmts.chrome("label_bold"))
    s.merge(row, 14, 15, None, fmts.calc("text"))
    ws.write_formula(row, 14, '=TEXT(DashStart,"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$O${row + 1}")

    s.merge(row, 16, 16, "Today", fmts.chrome("label_bold"))
    s.write(row, 17, "=TODAY()", fmts.calc("date"))
    row += 2

    # ------------------------------------------------------- output tiles
    row = s.section(row, "What you published", LAST_COL)
    s.kpi(row, 0, "Posts planned",
          f"=COUNTIFS({_WINDOW})", "int", 3,
          "Everything on the calendar in this period")
    s.kpi(row, 3, "Published",
          f'=COUNTIFS({_WINDOW},[[Content Calendar.status]],"Published")',
          "int", 3, "Actually went out")
    book.define_name("Dash_Published", f"={quote_sheet(NAME)}!$D${row + 2}")
    s.kpi(row, 6, "Still to come",
          f'=COUNTIFS({_WINDOW},[[Content Calendar.status]],"Scheduled")'
          f'+COUNTIFS({_WINDOW},[[Content Calendar.status]],"Approved")',
          "int", 3, "Approved or scheduled, not yet live")
    s.kpi(row, 9, "Target for the month",
          '=IF(DashMonthName="All year",MonthlyPostTarget*12,'
          'MonthlyPostTarget)', "int", 3,
          "From the Settings tab")
    s.kpi(row, 12, "Against target",
          '=IFERROR(Dash_Published/IF(DashMonthName="All year",'
          'MonthlyPostTarget*12,MonthlyPostTarget),0)', "percent", 3,
          "Published divided by your target")
    s.kpi(row, 15, "Needing attention",
          '=COUNTIFS([[Content Calendar.alert]],"<>OK",'
          '[[Content Calendar.title]],"<>")', "int", 3,
          "Rows with an action in the last column")
    row += 4

    # -------------------------------------------------------- reach tiles
    row = s.section(row, "How it performed", LAST_COL)
    s.kpi(row, 0, "Reach", f"=SUMIFS([[Analytics.reach]],{_AN})", "int", 3,
          "People reached across every account")
    book.define_name("Dash_Reach", f"={quote_sheet(NAME)}!$A${row + 2}")
    s.kpi(row, 3, "Impressions",
          f"=SUMIFS([[Analytics.impressions]],{_AN})", "int", 3,
          "Times your content was shown")
    s.kpi(row, 6, "Engagements",
          f"=SUMIFS([[Analytics.engagements]],{_AN})", "int", 3,
          "Likes, comments, shares and saves")
    book.define_name("Dash_Engagements",
                     f"={quote_sheet(NAME)}!$G${row + 2}")
    s.kpi(row, 9, "Engagement rate",
          "=IFERROR(Dash_Engagements/Dash_Reach,0)", "percent", 3,
          "Engagements divided by reach")
    s.kpi(row, 12, "Target rate", "=TargetEngagement", "percent", 3,
          "Your goal, set on the Settings tab")
    s.kpi(row, 15, "Reach per post",
          "=IFERROR(Dash_Reach/Dash_Published,0)", "int", 3,
          "How far the average post travelled")
    row += 4

    # ----------------------------------------------------- audience tiles
    row = s.section(row, "Audience & results", LAST_COL)
    s.kpi(row, 0, "Followers",
          f"=SUMIFS([[Analytics.followers_end]],{_AN})", "int", 3,
          "Total across every account at the end of the period")
    s.kpi(row, 3, "Follower growth",
          f"=SUMIFS([[Analytics.net_growth]],{_AN})", "int", 3,
          "New followers less unfollows")
    s.kpi(row, 6, "Link clicks", f"=SUMIFS([[Analytics.clicks]],{_AN})",
          "int", 3, "Traffic you sent out")
    s.kpi(row, 9, "Leads", f"=SUMIFS([[Analytics.leads]],{_AN})", "int", 3,
          "Sign-ups, enquiries, waitlist joins")
    s.kpi(row, 12, "Sales", f"=SUMIFS([[Analytics.conversions]],{_AN})",
          "int", 3, "Orders you can trace back to social")
    s.kpi(row, 15, "Revenue", f"=SUMIFS([[Analytics.revenue]],{_AN})",
          "money", 3, "Money attributed to social")
    book.define_name("Dash_Revenue", f"={quote_sheet(NAME)}!$P${row + 2}")
    row += 4

    # --------------------------------------------------------- money tiles
    row = s.section(row, "What it cost", LAST_COL)
    s.kpi(row, 0, "Spend",
          '=SUMIFS([[Budget.actual]],[[Budget.month]],">="&DashStart,'
          '[[Budget.month]],"<="&DashEnd)', "money", 3,
          "Everything on the Budget tab for this period")
    book.define_name("Dash_Spend", f"={quote_sheet(NAME)}!$A${row + 2}")
    s.kpi(row, 3, "Budget", '=IF(DashMonthName="All year",AnnualBudget,'
                            'MonthlyBudget)', "money", 3,
          "From the Settings tab")
    s.kpi(row, 6, "Left to spend",
          '=IF(DashMonthName="All year",AnnualBudget,MonthlyBudget)'
          '-Dash_Spend', "money", 3, "Budget less actual spend")
    s.kpi(row, 9, "Return on spend",
          "=IFERROR(Dash_Revenue/Dash_Spend,0)", "number1", 3,
          "Revenue divided by spend")
    s.kpi(row, 12, "Cost per lead",
          f"=IFERROR(Dash_Spend/SUMIFS([[Analytics.leads]],{_AN}),0)",
          "money", 3, "Spend divided by leads")
    s.kpi(row, 15, "Revenue per post",
          "=IFERROR(Dash_Revenue/Dash_Published,0)", "money", 3,
          "Money per published post")
    row += 4

    # ------------------------------------------------- twelve month table
    row = s.section(row, f"The whole of {cfg.year} month by month", LAST_COL)
    headers = ["Month", "Published", "Reach", "Engagements", "Eng. rate",
               "Followers", "Growth", "Clicks", "Leads", "Sales",
               "Revenue", "Spend", "Net"]
    s.table_head(row, 0, headers)
    first_table_row = row + 1
    keys = ["m_label", "m_published", "m_reach", "m_engagements", "m_er",
            "m_followers", "m_growth", "m_clicks", "m_leads",
            "m_conversions", "m_revenue", "m_spend", "m_profit"]
    styles = ["text", "int", "int", "int", "percent", "int", "int", "int",
              "int", "int", "money", "money", "money"]
    for i in range(MONTH_ROWS):
        r = first_table_row + i
        calc_row = FIRST_CALC_ROW + i
        s.calc_row(r, 0, [f"={_calc(key, calc_row)}" for key in keys],
                   styles)
    last_table_row = first_table_row + MONTH_ROWS - 1

    for col_idx, colour in ((2, "#6350A8"), (3, "#F2665E"), (5, "#2A4A8B"),
                            (10, "#0E6B45")):
        ws.conditional_format(first_table_row, col_idx, last_table_row,
                              col_idx, {"type": "data_bar",
                                        "bar_color": colour})
    ws.conditional_format(first_table_row, 4, last_table_row, 4, {
        "type": "cell", "criteria": ">=", "value": "=TargetEngagement",
        "format": fmts.cf("good")})

    # sparkline strip under the table
    spark_row = last_table_row + 1
    s.write(spark_row, 0, "Trend", fmts.chrome("legend_label"))
    for col_idx, colour in ((1, "#6350A8"), (2, "#3B2A6B"), (3, "#F2665E"),
                            (5, "#2A4A8B"), (10, "#0E6B45")):
        letter = col_letter(col_idx)
        try:
            ws.add_sparkline(spark_row, col_idx, {
                "range": (f"{quote_sheet(NAME)}!{letter}"
                          f"{first_table_row + 1}:{letter}"
                          f"{last_table_row + 1}"),
                "type": "column",
                "series_color": colour,
            })
        except Exception:        # pragma: no cover - library differences
            pass
    for col_idx in range(1, len(headers)):
        if col_idx not in (1, 2, 3, 5, 10):
            ws.write_blank(spark_row, col_idx, None,
                           fmts.chrome("legend_calc"))
    row = spark_row + 2

    # ------------------------------------------------ platform comparison
    row = s.section(row, "Platform by platform (selected period)", 8)
    s.table_head(row, 0, ["Platform", "Posts", "Reach", "Engagements",
                          "Eng. rate", "Followers", "Growth", "Clicks",
                          "Revenue"])
    first_platform = row + 1
    platform_rows = 8
    pkeys = ["p_name", "p_published_m", "p_reach_m", "p_eng_m", "p_er_m",
             "p_followers", "p_growth_m", "p_clicks_m", "p_revenue_m"]
    pstyles = ["text", "int", "int", "int", "percent", "int", "int", "int",
               "money"]
    for i in range(platform_rows):
        r = first_platform + i
        calc_row = FIRST_CALC_ROW + i
        s.calc_row(r, 0, [f'=IF({_calc("p_name", calc_row)}="","",'
                          f"{_calc(key, calc_row)})" for key in pkeys],
                   pstyles)
    ws.conditional_format(first_platform, 2, first_platform
                          + platform_rows - 1, 2,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    ws.conditional_format(first_platform, 4, first_platform
                          + platform_rows - 1, 4,
                          {"type": "3_color_scale",
                           "min_color": "#FBD9DD", "mid_color": "#FFF1CF",
                           "max_color": "#D9F2E4"})

    # pillar table alongside
    s.merge(row - 1, 10, LAST_COL, "Content pillars (this year)",
            fmts.chrome("subsection"))
    s.table_head(row, 10, ["Pillar", "Posts", "Reach", "Engagements",
                           "Eng. rate", "Share"])
    pillar_rows = min(platform_rows, cfg.row_count("Content Pillars"))
    for i in range(pillar_rows):
        r = first_platform + i
        calc_row = FIRST_CALC_ROW + i
        s.calc_row(r, 10, [
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_name", calc_row)})',
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_posts", calc_row)})',
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_reach", calc_row)})',
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_eng", calc_row)})',
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_er", calc_row)})',
            f'=IFERROR({_calc("pl_posts", calc_row)}'
            f'/SUM([[Calc.pl_posts]]),"")',
        ], ["text", "int", "int", "int", "percent", "percent"])
    ws.conditional_format(first_platform, 12, first_platform
                          + pillar_rows - 1, 12,
                          {"type": "data_bar", "bar_color": "#F2665E"})
    row = first_platform + platform_rows + 1

    # ------------------------------------------------------- top & alerts
    row = s.section(row, "Your best posts & what needs doing", LAST_COL)
    s.merge(row, 0, 8, "Top 8 posts in this period (by engagement rate)",
            fmts.chrome("subsection"))
    s.merge(row, 10, LAST_COL, "What needs attention right now",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["#", "Post", "Platform", "Pillar", "Date",
                          "Reach", "Eng. rate"])
    s.table_head(row, 10, ["Action", "How many", "", "", "", "", "", ""])
    first_top = row + 1
    for i in range(8):
        r = first_top + i
        n = i + 1
        pick = (f'MATCH(LARGE([[Calc.pf_key_m]],{n}),'
                f'[[Calc.pf_key_m]],0)')
        s.calc_row(r, 0, [
            n,
            f'=IFERROR(INDEX([[Calc.pf_title]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_platform]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_pillar]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_date]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_reach]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_er]],{pick}),"")',
        ], ["int", "text", "text", "text", "date", "int", "percent"])
    ws.conditional_format(first_top, 5, first_top + 7, 5,
                          {"type": "data_bar", "bar_color": "#6350A8"})

    alerts = [
        ("Posts overdue", '=COUNTIF([[Content Calendar.alert]],'
                          '"Overdue*")'),
        ("Drafts running late", '=COUNTIF([[Content Calendar.alert]],'
                                '"Late*")'),
        ("Waiting for approval",
         '=COUNTIF([[Content Calendar.status]],"Awaiting review")'),
        ("Posts with no caption",
         '=COUNTIFS([[Content Calendar.title]],"<>",'
         '[[Content Calendar.caption]],"")'),
        ("Published with no link",
         '=COUNTIF([[Content Calendar.published]],"Add*")'),
        ("Published with no numbers yet",
         '=COUNTIFS([[Content Calendar.status]],"Published",'
         '[[Content Calendar.reach]],"")'),
        ("Ideas worth scheduling",
         '=COUNTIF([[Idea Bank.alert]],"Plan*")'),
        ("Overdue tasks", '=COUNTIF([[Tasks.alert]],"Overdue*")'),
        ("UGC waiting on permission",
         '=COUNTIF([[UGC Tracker.alert]],"Action*")'),
        ("Goals behind pace", '=COUNTIF([[Goals.pace]],"Behind*")'),
        ("Accounts gone quiet",
         '=COUNTIF([[Accounts.alert]],"Check - account*")'),
        ("Trends about to expire",
         '=COUNTIF([[Trends.alert]],"Expired*")'),
    ]
    for i, (label, formula) in enumerate(alerts):
        r = first_top + i
        s.calc_row(r, 10, [label, formula], ["text", "int"])
        ws.conditional_format(r, 11, r, 11, {
            "type": "cell", "criteria": ">", "value": 0,
            "format": fmts.cf("bad")})
    row = first_top + max(8, len(alerts)) + 1

    # ------------------------------------------------------- what is next
    row = s.section(row, "Coming up & running late", LAST_COL)
    s.merge(row, 0, 8, "Next 8 posts due", fmts.chrome("subsection"))
    s.merge(row, 10, LAST_COL, "Overdue - deal with these first",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["Date", "Platform", "Post", "Status", "Owner",
                          "Days"])
    s.table_head(row, 10, ["Date", "Platform", "Post", "Status", "Owner",
                           "Days late"])
    first_next = row + 1
    for i in range(8):
        r = first_next + i
        n = i + 1
        up = (f'MATCH(SMALL([[Content Calendar.queue_key]],{n}),'
              f'[[Content Calendar.queue_key]],0)')
        late = (f'MATCH(SMALL([[Content Calendar.late_key]],{n}),'
                f'[[Content Calendar.late_key]],0)')
        s.calc_row(r, 0, [
            f'=IFERROR(INDEX([[Content Calendar.date]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.platform]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.title]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.status]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.owner]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.days_to_go]],{up}),"")',
        ], ["date", "text", "text", "status", "text", "int"])
        s.calc_row(r, 10, [
            f'=IFERROR(INDEX([[Content Calendar.date]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.platform]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.title]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.status]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.owner]],{late}),"")',
            f'=IFERROR(-INDEX([[Content Calendar.days_to_go]],{late}),"")',
        ], ["date", "text", "text", "status", "text", "int"])
    row = first_next + 9

    s.note(row,
           "Tip: pick 'All year' in the month box to see the whole plan. "
           "The Annual Dashboard compares every month side by side, and "
           "the Reports tab turns the same numbers into something you can "
           "print or send to a client.",
           LAST_COL, height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `social_media_spreadsheet/sheets/annual.py`

`Annual Dashboard` - the whole year on one page. *(348 lines)*

````python
"""`Annual Dashboard` - the whole year on one page."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..writer import SheetWriter

NAME = "Annual Dashboard"
LAST_COL = 17
FIRST_CALC_ROW = 5
MONTH_ROWS = 12

_YEAR = ('[[Content Calendar.date]],">="&DATE(PlanYear,1,1),'
         '[[Content Calendar.date]],"<="&DATE(PlanYear,12,31)')


def _calc(key: str, row: int) -> str:
    return f"[[Calc.{key}#{row}]]"


def _quarter(metric: str, first_month: int) -> str:
    """Sum three month rows of the Calc month series."""
    rows = [FIRST_CALC_ROW + first_month + i for i in range(3)]
    return "=" + "+".join(f"N({_calc(metric, r)})" for r in rows)


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    ws = s.ws

    s.banner("Annual Dashboard",
             f"Everything you planned and published in {cfg.year}, "
             f"quarter by quarter and month by month.",
             LAST_COL)
    for c in range(0, LAST_COL + 1):
        s.set_col(c, c, 12.5)

    # ------------------------------------------------------- year headline
    row = 4
    row = s.section(row, f"{cfg.year} in numbers", LAST_COL)
    s.kpi(row, 0, "Posts planned", f"=COUNTIFS({_YEAR})", "int", 3,
          "Every row on the calendar this year")
    s.kpi(row, 3, "Published",
          f'=COUNTIFS({_YEAR},[[Content Calendar.status]],"Published")',
          "int", 3, "Actually went out")
    book.define_name("Year_Published", f"={quote_sheet(NAME)}!$D${row + 2}")
    s.kpi(row, 6, "Reach", "=SUM([[Calc.m_reach]])", "int", 3,
          "Across every account and month")
    book.define_name("Year_Reach", f"={quote_sheet(NAME)}!$G${row + 2}")
    s.kpi(row, 9, "Engagements", "=SUM([[Calc.m_engagements]])", "int", 3,
          "Likes, comments, shares, saves")
    book.define_name("Year_Eng", f"={quote_sheet(NAME)}!$J${row + 2}")
    s.kpi(row, 12, "Engagement rate", "=IFERROR(Year_Eng/Year_Reach,0)",
          "percent", 3, "The year's average")
    s.kpi(row, 15, "Follower growth", "=SUM([[Calc.m_growth]])", "int", 3,
          "New followers less unfollows")
    row += 4

    s.kpi(row, 0, "Clicks", "=SUM([[Calc.m_clicks]])", "int", 3,
          "Traffic sent out")
    s.kpi(row, 3, "Leads", "=SUM([[Calc.m_leads]])", "int", 3,
          "Enquiries and sign-ups")
    s.kpi(row, 6, "Sales", "=SUM([[Calc.m_conversions]])", "int", 3,
          "Orders traced back to social")
    s.kpi(row, 9, "Revenue", "=SUM([[Calc.m_revenue]])", "money", 3,
          "Money attributed to social")
    book.define_name("Year_Revenue", f"={quote_sheet(NAME)}!$J${row + 2}")
    s.kpi(row, 12, "Spend", "=SUM([[Calc.m_spend]])", "money", 3,
          "Everything on the Budget tab")
    book.define_name("Year_Spend", f"={quote_sheet(NAME)}!$M${row + 2}")
    s.kpi(row, 15, "Return on spend", "=IFERROR(Year_Revenue/Year_Spend,0)",
          "number1", 3, "Revenue divided by spend")
    row += 4

    s.kpi(row, 0, "Posts a week",
          "=IFERROR(Year_Published/52,0)", "number1", 3,
          "Your real publishing rhythm")
    s.kpi(row, 3, "Reach per post",
          "=IFERROR(Year_Reach/Year_Published,0)", "int", 3,
          "How far the average post travelled")
    s.kpi(row, 6, "Revenue per post",
          "=IFERROR(Year_Revenue/Year_Published,0)", "money", 3,
          "Money per published post")
    s.kpi(row, 9, "Best month",
          '=IFERROR(INDEX([[Calc.m_label]],MATCH(MAX([[Calc.m_reach]]),'
          '[[Calc.m_reach]],0)),"")', "text", 3, "By reach")
    s.kpi(row, 12, "Best platform",
          '=IFERROR(INDEX([[Calc.p_name]],MATCH(MAX([[Calc.p_key]]),'
          '[[Calc.p_key]],0)),"")', "text", 3, "By reach this year")
    s.kpi(row, 15, "Best day to post",
          '=IFERROR(INDEX([[Calc.dw_day]],MATCH(MAX([[Calc.dw_key]]),'
          '[[Calc.dw_key]],0)),"")', "text", 3,
          "Highest average engagement rate")
    row += 4

    # ------------------------------------------------------------ quarters
    row = s.section(row, "Quarter by quarter", LAST_COL)
    s.table_head(row, 0, ["Quarter", "Published", "Reach", "Engagements",
                          "Eng. rate", "Followers gained", "Clicks",
                          "Leads", "Sales", "Revenue", "Spend", "Net"])
    first_q = row + 1
    for i, label in enumerate(("Q1", "Q2", "Q3", "Q4")):
        r = first_q + i
        start = i * 3
        pub = _quarter("m_published", start)
        reach = _quarter("m_reach", start)
        eng = _quarter("m_engagements", start)
        rev = _quarter("m_revenue", start)
        spend = _quarter("m_spend", start)
        reach_cell = f"{col_letter(2)}{r + 1}"
        eng_cell = f"{col_letter(3)}{r + 1}"
        rev_cell = f"{col_letter(9)}{r + 1}"
        spend_cell = f"{col_letter(10)}{r + 1}"
        s.calc_row(r, 0, [
            label, pub, reach, eng,
            f"=IFERROR({eng_cell}/{reach_cell},\"\")",
            _quarter("m_growth", start),
            _quarter("m_clicks", start),
            _quarter("m_leads", start),
            _quarter("m_conversions", start),
            rev, spend,
            f"={rev_cell}-{spend_cell}",
        ], ["text", "int", "int", "int", "percent", "int", "int", "int",
            "int", "money", "money", "money"])
    ws.conditional_format(first_q, 2, first_q + 3, 2,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    ws.conditional_format(first_q, 9, first_q + 3, 9,
                          {"type": "data_bar", "bar_color": "#0E6B45"})
    row = first_q + 5

    # ------------------------------------------------------ monthly detail
    row = s.section(row, "Every month in detail", LAST_COL)
    headers = ["Month", "Planned", "Published", "On target?", "Reach",
               "Impressions", "Engagements", "Eng. rate", "Followers",
               "Growth", "Clicks", "Leads", "Sales", "Revenue", "Spend",
               "Net", "Return"]
    s.table_head(row, 0, headers)
    first_month_row = row + 1
    for i in range(MONTH_ROWS):
        r = first_month_row + i
        cr = FIRST_CALC_ROW + i
        rev = f"{col_letter(13)}{r + 1}"
        spend = f"{col_letter(14)}{r + 1}"
        s.calc_row(r, 0, [
            f"={_calc('m_label', cr)}",
            f"={_calc('m_planned', cr)}",
            f"={_calc('m_published', cr)}",
            f"=IFERROR({_calc('m_published', cr)}/"
            f"{_calc('m_target', cr)},\"\")",
            f"={_calc('m_reach', cr)}",
            f"={_calc('m_impressions', cr)}",
            f"={_calc('m_engagements', cr)}",
            f"={_calc('m_er', cr)}",
            f"={_calc('m_followers', cr)}",
            f"={_calc('m_growth', cr)}",
            f"={_calc('m_clicks', cr)}",
            f"={_calc('m_leads', cr)}",
            f"={_calc('m_conversions', cr)}",
            f"={_calc('m_revenue', cr)}",
            f"={_calc('m_spend', cr)}",
            f"={rev}-{spend}",
            f"=IFERROR({rev}/{spend},\"\")",
        ], ["text", "int", "int", "percent", "int", "int", "int",
            "percent", "int", "int", "int", "int", "int", "money",
            "money", "money", "number1"])
    last_month_row = first_month_row + MONTH_ROWS - 1
    for col_idx, colour in ((4, "#6350A8"), (6, "#F2665E"), (8, "#2A4A8B"),
                            (13, "#0E6B45")):
        ws.conditional_format(first_month_row, col_idx, last_month_row,
                              col_idx, {"type": "data_bar",
                                        "bar_color": colour})
    ws.conditional_format(first_month_row, 3, last_month_row, 3, {
        "type": "cell", "criteria": ">=", "value": 1,
        "format": fmts.cf("good")})
    ws.conditional_format(first_month_row, 7, last_month_row, 7, {
        "type": "cell", "criteria": ">=", "value": "=TargetEngagement",
        "format": fmts.cf("good")})

    spark_row = last_month_row + 1
    s.write(spark_row, 0, "Trend", fmts.chrome("legend_label"))
    for col_idx, colour in ((2, "#6350A8"), (4, "#3B2A6B"),
                            (8, "#2A4A8B"), (13, "#0E6B45")):
        letter = col_letter(col_idx)
        try:
            ws.add_sparkline(spark_row, col_idx, {
                "range": (f"{quote_sheet(NAME)}!{letter}"
                          f"{first_month_row + 1}:{letter}"
                          f"{last_month_row + 1}"),
                "type": "column",
                "series_color": colour,
            })
        except Exception:        # pragma: no cover
            pass
    for col_idx in range(1, len(headers)):
        if col_idx not in (2, 4, 8, 13):
            ws.write_blank(spark_row, col_idx, None,
                           fmts.chrome("legend_calc"))
    row = spark_row + 2

    # ------------------------------------------------ platforms and goals
    row = s.section(row, "Platforms, formats and goals", LAST_COL)
    s.merge(row, 0, 6, "Platform totals for the year",
            fmts.chrome("subsection"))
    s.merge(row, 8, LAST_COL, "Goals", fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["Platform", "Posts", "Reach", "Engagements",
                          "Eng. rate", "Growth", "Share of reach"])
    s.table_head(row, 8, ["Goal", "Target", "Now", "Progress", "On track?",
                          "By when"])
    first_p = row + 1
    platform_rows = 8
    for i in range(platform_rows):
        r = first_p + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF({_calc("p_name", cr)}="","",'
        s.calc_row(r, 0, [
            f"={guard}{_calc('p_name', cr)})",
            f"={guard}{_calc('p_posts_y', cr)})",
            f"={guard}{_calc('p_reach_y', cr)})",
            f"={guard}{_calc('p_eng_y', cr)})",
            f"={guard}{_calc('p_er_y', cr)})",
            f"={guard}{_calc('p_growth_y', cr)})",
            f"=IFERROR({_calc('p_reach_y', cr)}/SUM([[Calc.p_reach_y]]),"
            f'"")',
        ], ["text", "int", "int", "int", "percent", "int", "percent"])
    ws.conditional_format(first_p, 2, first_p + platform_rows - 1, 2,
                          {"type": "data_bar", "bar_color": "#6350A8"})

    goal_rows = min(8, cfg.row_count("Goals"))
    for i in range(goal_rows):
        r = first_p + i
        gr = FIRST_CALC_ROW + i
        guard = f'IF([[Goals.name#{gr}]]="","",'
        s.calc_row(r, 8, [
            f"={guard}[[Goals.name#{gr}]])",
            f"={guard}[[Goals.target#{gr}]])",
            f"={guard}[[Goals.current#{gr}]])",
            f"={guard}[[Goals.progress#{gr}]])",
            f"={guard}[[Goals.pace#{gr}]])",
            f"={guard}[[Goals.target_date#{gr}]])",
        ], ["text", "number", "number", "percent", "status", "date"])
    ws.conditional_format(first_p, 11, first_p + goal_rows - 1, 11,
                          {"type": "data_bar", "bar_color": "#0E6B45"})
    row = first_p + max(platform_rows, goal_rows) + 1

    # ------------------------------------------------------- best of year
    row = s.section(row, f"Best of {cfg.year}", LAST_COL)
    s.merge(row, 0, 7, "Top 10 posts of the year (by engagement rate)",
            fmts.chrome("subsection"))
    s.merge(row, 9, LAST_COL, "Where your reach came from",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["#", "Post", "Platform", "Pillar", "Date",
                          "Reach", "Eng. rate"])
    s.table_head(row, 9, ["Format", "Posts", "Reach", "Engagements",
                          "Eng. rate"])
    first_top = row + 1
    for i in range(10):
        r = first_top + i
        n = i + 1
        pick = (f"MATCH(LARGE([[Calc.pf_key_y]],{n}),[[Calc.pf_key_y]],0)")
        s.calc_row(r, 0, [
            n,
            f'=IFERROR(INDEX([[Calc.pf_title]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_platform]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_pillar]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_date]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_reach]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_er]],{pick}),"")',
        ], ["int", "text", "text", "text", "date", "int", "percent"])
    ws.conditional_format(first_top, 5, first_top + 9, 5,
                          {"type": "data_bar", "bar_color": "#6350A8"})

    for i in range(10):
        r = first_top + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF({_calc("fm_name", cr)}="","",'
        s.calc_row(r, 9, [
            f"={guard}{_calc('fm_name', cr)})",
            f"={guard}{_calc('fm_posts', cr)})",
            f"={guard}{_calc('fm_reach', cr)})",
            f"={guard}{_calc('fm_eng', cr)})",
            f"={guard}{_calc('fm_er', cr)})",
        ], ["text", "int", "int", "int", "percent"])
    ws.conditional_format(first_top, 11, first_top + 9, 11,
                          {"type": "data_bar", "bar_color": "#F2665E"})
    row = first_top + 11

    # --------------------------------------------------- timing and review
    row = s.section(row, "When to post & what to do next", LAST_COL)
    s.table_head(row, 0, ["Day", "Posts", "Reach", "Engagements",
                          "Eng. rate"])
    s.table_head(row, 6, ["Time slot", "Posts", "Reach", "Engagements",
                          "Eng. rate"])
    first_day = row + 1
    for i in range(7):
        r = first_day + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF({_calc("dw_day", cr)}="","",'
        s.calc_row(r, 0, [
            f"={guard}{_calc('dw_day', cr)})",
            f"={guard}{_calc('dw_posts', cr)})",
            f"={guard}{_calc('dw_reach', cr)})",
            f"={guard}{_calc('dw_eng', cr)})",
            f"={guard}{_calc('dw_er', cr)})",
        ], ["text", "int", "int", "int", "percent"])
        guard2 = f'IF({_calc("sl_slot", cr)}="","",'
        s.calc_row(r, 6, [
            f"={guard2}{_calc('sl_slot', cr)})",
            f"={guard2}{_calc('sl_posts', cr)})",
            f"={guard2}{_calc('sl_reach', cr)})",
            f"={guard2}{_calc('sl_eng', cr)})",
            f"={guard2}{_calc('sl_er', cr)})",
        ], ["text", "int", "int", "int", "percent"])
    ws.conditional_format(first_day, 4, first_day + 6, 4,
                          {"type": "3_color_scale",
                           "min_color": "#FBD9DD", "mid_color": "#FFF1CF",
                           "max_color": "#D9F2E4"})
    ws.conditional_format(first_day, 10, first_day + 6, 10,
                          {"type": "3_color_scale",
                           "min_color": "#FBD9DD", "mid_color": "#FFF1CF",
                           "max_color": "#D9F2E4"})

    s.merge(first_day - 1, 12, LAST_COL, "Campaigns by revenue",
            fmts.chrome("subsection"))
    s.table_head(first_day, 12, ["Campaign", "Revenue", "Spend", "Return"])
    for i in range(6):
        r = first_day + 1 + i
        n = i + 1
        pick = f"MATCH(LARGE([[Calc.cm_key]],{n}),[[Calc.cm_key]],0)"
        s.calc_row(r, 12, [
            f'=IFERROR(INDEX([[Calc.cm_name]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.cm_revenue]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.cm_spend]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.cm_revenue]],{pick})'
            f'/INDEX([[Calc.cm_spend]],{pick}),"")',
        ], ["text", "money", "money", "number1"])
    row = first_day + 8

    s.note(row,
           "Use this page at the end of the year: look at the best month, "
           "the best day, the best format and your top ten posts, then "
           "build next year's pillars around what actually worked.",
           LAST_COL, height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `social_media_spreadsheet/sheets/month_view.py`

`Month View` - the calendar as a calendar: six weeks across seven days. *(169 lines)*

````python
"""`Month View` - the calendar as a calendar: six weeks across seven days.

Excel 2010 has no FILTER, so each day cell finds its posts with a trick:
the Content Calendar carries a hidden ``day_key`` (the date plus a tiny
fraction of the row number).  Because that key sorts by date, the n-th
post on a given day is ``SMALL(day_key, COUNTIF(day_key,"<"&day) + n)``,
which needs no array formulas at all.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from

NAME = "Month View"
DAYS = 7
WEEKS = 6
PER_DAY = 3              # posts shown in each day cell
COL_WIDTH = 19.0
LAST_COL = DAYS - 1

_KEY = "[[Content Calendar.day_key]]"


def _nth(day_cell: str, n: int, column: str) -> str:
    """Formula for the n-th post on `day_cell`, pulling `column`."""
    pick = f'SMALL({_KEY},COUNTIF({_KEY},"<"&{day_cell})+{n})'
    return (f'=IF({day_cell}="","",IFERROR(IF({pick}>={day_cell}+1,"",'
            f'INDEX([[Content Calendar.{column}]],'
            f'MATCH({pick},{_KEY},0))),""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Month View",
             "The month laid out as a grid, straight from the Content "
             "Calendar. Pick a month, then print it and pin it up.",
             LAST_COL)
    for c in range(DAYS):
        s.set_col(c, c, COL_WIDTH)

    # ------------------------------------------------------ month picker
    row = 4
    s.merge(row, 0, 1, "Month", fmts.chrome("label_bold"))
    s.merge(row, 2, 3, cfg.today.strftime("%B"), fmts.input("text"))
    ws.data_validation(row, 2, row, 2, list_from("MonthList"))
    book.define_name("ViewMonthName", f"={quote_sheet(NAME)}!$C${row + 1}")
    s.merge(row, 4, 4, "First day", fmts.chrome("label_bold"))
    s.merge(row, 5, 6, None, fmts.calc("date"))
    ws.write_formula(
        row, 5,
        '=DATE(PlanYear,MAX(1,IF(ViewMonthName="All year",'
        'MONTH(TODAY()),MATCH(ViewMonthName,MonthList,0)-1)),1)',
        fmts.calc("date"))
    book.define_name("ViewStart", f"={quote_sheet(NAME)}!$F${row + 1}")
    row += 1

    s.note(row,
           "Each box shows up to three posts for that day: platform, "
           "then the title. Days with more than three are marked '+ more' "
           "- the Content Calendar has the full list.",
           LAST_COL, height=24)
    row += 1

    # ----------------------------------------------------- weekday header
    first_cell = ('=ViewStart-MOD(WEEKDAY(ViewStart,IF(WeekStart="Sunday",'
                  '1,2))-1,7)')
    header_row = row
    for d in range(DAYS):
        ws.write_formula(
            header_row, d,
            f'=TEXT({first_cell.lstrip("=")}+{d},"dddd")',
            fmts.chrome("header"))
    s.row_height(header_row, 22)
    row += 1

    # ------------------------------------------------------- the day grid
    date_fmt = fmts.chrome("day_number")
    day_cells = []
    for week in range(WEEKS):
        date_row = row
        s.row_height(date_row, 20)
        for d in range(DAYS):
            offset = week * 7 + d
            ws.write_formula(
                date_row, d,
                f'=IF(MONTH({first_cell.lstrip("=")}+{offset})'
                f'<>MONTH(ViewStart),"",'
                f'{first_cell.lstrip("=")}+{offset})',
                date_fmt)
            day_cells.append((date_row, d))
        for line in range(PER_DAY):
            r = date_row + 1 + line
            s.row_height(r, 15)
            for d in range(DAYS):
                cell = f"{chr(65 + d)}{date_row + 1}"
                if line < PER_DAY - 1:
                    formula = (
                        f'=IF({cell}="","",IFERROR(IF('
                        f'SMALL({_KEY},COUNTIF({_KEY},"<"&{cell})'
                        f'+{line + 1})>={cell}+1,"",'
                        f'INDEX([[Content Calendar.platform]],MATCH(SMALL('
                        f'{_KEY},COUNTIF({_KEY},"<"&{cell})+{line + 1}),'
                        f'{_KEY},0))&" - "&INDEX('
                        f'[[Content Calendar.title]],MATCH(SMALL({_KEY},'
                        f'COUNTIF({_KEY},"<"&{cell})+{line + 1}),'
                        f'{_KEY},0))),""))')
                else:
                    # last line: either the third post or a "+ n more" note
                    formula = (
                        f'=IF({cell}="","",IF(COUNTIF({_KEY},"<"&{cell}+1)'
                        f'-COUNTIF({_KEY},"<"&{cell})>{PER_DAY},'
                        f'"+ "&COUNTIF({_KEY},"<"&{cell}+1)'
                        f'-COUNTIF({_KEY},"<"&{cell})-{PER_DAY - 1}'
                        f'&" more",IFERROR(IF(SMALL({_KEY},'
                        f'COUNTIF({_KEY},"<"&{cell})+{line + 1})'
                        f'>={cell}+1,"",INDEX('
                        f'[[Content Calendar.platform]],MATCH(SMALL({_KEY},'
                        f'COUNTIF({_KEY},"<"&{cell})+{line + 1}),'
                        f'{_KEY},0))&" - "&INDEX('
                        f'[[Content Calendar.title]],MATCH(SMALL({_KEY},'
                        f'COUNTIF({_KEY},"<"&{cell})+{line + 1}),'
                        f'{_KEY},0))),"")))')
                ws.write_formula(r, d, s.reg.render(formula),
                                 fmts.chrome("day_item"))
        ws.conditional_format(
            date_row, 0, date_row + PER_DAY, DAYS - 1,
            {"type": "formula",
             "criteria": f"=A${date_row + 1}=TODAY()",
             "format": fmts.cf("info")})
        row = date_row + PER_DAY + 1


    # ---------------------------------------------------- month summary --
    row += 1
    row = s.section(row, "This month at a glance", LAST_COL)
    s.table_head(row, 0, ["Posts planned", "Published", "Still to go",
                          "Target", "Against target", "Reach",
                          "Engagements"])
    window = ('[[Content Calendar.date]],">="&ViewStart,'
              '[[Content Calendar.date]],"<="&EOMONTH(ViewStart,0)')
    an_window = ('[[Analytics.month]],">="&ViewStart,'
                 '[[Analytics.month]],"<="&EOMONTH(ViewStart,0)')
    s.calc_row(row + 1, 0, [
        f"=COUNTIFS({window})",
        f'=COUNTIFS({window},[[Content Calendar.status]],"Published")',
        f'=COUNTIFS({window})-COUNTIFS({window},'
        f'[[Content Calendar.status]],"Published")',
        "=MonthlyPostTarget",
        f'=IFERROR(COUNTIFS({window},[[Content Calendar.status]],'
        f'"Published")/MonthlyPostTarget,"")',
        f"=SUMIFS([[Analytics.reach]],{an_window})",
        f"=SUMIFS([[Analytics.engagements]],{an_window})",
    ], ["int", "int", "int", "int", "percent", "int", "int"])
    row += 3

    s.note(row,
           "Posts appear here the moment you add them to the Content "
           "Calendar - there is nothing to copy across. Print this page "
           "on one sheet of landscape A4 and it makes a very good "
           "wall planner.",
           LAST_COL, height=26)

    s.finish(freeze_row=header_row + 1, gridlines=False, fit_width=True)
    ws.set_landscape()
    return s
````

### `social_media_spreadsheet/sheets/workflow.py`

`Workflow` - the production pipeline: what is where, and what is late. *(204 lines)*

````python
"""`Workflow` - the production pipeline: what is where, and what is late."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Workflow"
LAST_COL = 15
FIRST_CALC_ROW = 5

STAGES = [
    ("1. Idea", "Idea",
     "Captured in the Idea Bank and scored. Nothing is written yet."),
    ("2. Brief", "Brief",
     "Hook, message, format and call to action agreed. Owner set."),
    ("3. In production", "In production",
     "Shooting, filming, writing or designing."),
    ("4. Draft", "Draft",
     "Caption written, asset ready, hashtag group chosen."),
    ("5. Awaiting review", "Awaiting review",
     "With whoever approves it. Check the facts, links and claims."),
    ("6. Revision needed", "Revision needed",
     "Sent back with notes. Fix and resubmit."),
    ("7. Approved", "Approved",
     "Signed off and ready to be loaded into the scheduler."),
    ("8. Scheduled", "Scheduled",
     "In the scheduler, or an alarm set to post manually."),
    ("9. Published", "Published",
     "Live. Add the link, then the numbers after 48-72 hours."),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Content Production Workflow",
             "Every post moves through the same nine stages. This page "
             "shows how many are sitting at each one, what is late, and "
             "who is holding it.",
             LAST_COL)
    s.set_col(0, 0, 20)
    s.set_col(1, 1, 13)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.section(row, "Where everything is right now", LAST_COL)
    s.table_head(row, 0, ["Stage", "All posts", "This period",
                          "What happens at this stage"])
    first_stage = row + 1
    for i, (label, status, description) in enumerate(STAGES):
        r = first_stage + i
        s.calc_row(r, 0, [
            label,
            f'=COUNTIF([[Content Calendar.status]],"{status}")',
            f'=COUNTIFS([[Content Calendar.status]],"{status}",'
            f'[[Content Calendar.date]],">="&DashStart,'
            f'[[Content Calendar.date]],"<="&DashEnd)',
        ], ["text", "int", "int"])
        s.merge(r, 3, LAST_COL, description, fmts.cell("wrap", "calc"))
    ws.conditional_format(first_stage, 1, first_stage + len(STAGES) - 1, 1,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    row = first_stage + len(STAGES) + 1

    # ------------------------------------------------------------- health
    row = s.section(row, "Pipeline health", LAST_COL)
    s.kpi(row, 0, "In the pipeline",
          '=COUNTIFS([[Content Calendar.title]],"<>",'
          '[[Content Calendar.status]],"<>Published",'
          '[[Content Calendar.status]],"<>Cancelled",'
          '[[Content Calendar.status]],"<>Archived")', "int", 3,
          "Everything not yet live")
    s.kpi(row, 3, "Drafts late",
          '=COUNTIF([[Content Calendar.alert]],"Late*")', "int", 3,
          "Past the draft deadline")
    s.kpi(row, 6, "Needing approval",
          '=COUNTIF([[Content Calendar.status]],"Awaiting review")',
          "int", 3, "Waiting on a yes")
    s.kpi(row, 9, "Due in 7 days",
          '=COUNTIFS([[Content Calendar.date]],">="&TODAY(),'
          '[[Content Calendar.date]],"<="&TODAY()+7,'
          '[[Content Calendar.status]],"<>Published")', "int", 3,
          "Coming up this week")
    s.kpi(row, 12, "Overdue",
          '=COUNTIF([[Content Calendar.alert]],"Overdue*")', "int", 3,
          "Date passed, not published")
    row += 4

    s.kpi(row, 0, "Weeks of content ready",
          '=IFERROR((COUNTIFS([[Content Calendar.status]],"Approved")'
          '+COUNTIFS([[Content Calendar.status]],"Scheduled"))'
          '/MAX(1,PostsPerWeek),0)', "number1", 3,
          "Approved plus scheduled, divided by your weekly target")
    s.kpi(row, 3, "Captions missing",
          '=COUNTIFS([[Content Calendar.title]],"<>",'
          '[[Content Calendar.caption]],"")', "int", 3,
          "Posts with no caption picked")
    s.kpi(row, 6, "Assets missing",
          '=COUNTIFS([[Content Calendar.title]],"<>",'
          '[[Content Calendar.asset]],"")', "int", 3,
          "Posts with no creative attached")
    s.kpi(row, 9, "Hashtag groups missing",
          '=COUNTIFS([[Content Calendar.title]],"<>",'
          '[[Content Calendar.hashtag_group]],"")', "int", 3,
          "Posts with no hashtag group")
    s.kpi(row, 12, "Open tasks",
          '=COUNTIFS([[Tasks.task]],"<>",[[Tasks.status]],"<>Complete")',
          "int", 3, "On the Tasks tab")
    row += 4

    # -------------------------------------------------------- the queues
    row = s.section(row, "The queue", LAST_COL)
    s.merge(row, 0, 6, "Next 12 posts due", fmts.chrome("subsection"))
    s.merge(row, 8, LAST_COL, "Overdue - deal with these first",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["Date", "Platform", "Post", "Status", "Owner",
                          "Days"])
    s.table_head(row, 8, ["Date", "Platform", "Post", "Status", "Owner",
                          "Days late"])
    first_queue = row + 1
    for i in range(12):
        r = first_queue + i
        n = i + 1
        up = (f"MATCH(SMALL([[Content Calendar.queue_key]],{n}),"
              f"[[Content Calendar.queue_key]],0)")
        late = (f"MATCH(SMALL([[Content Calendar.late_key]],{n}),"
                f"[[Content Calendar.late_key]],0)")
        s.calc_row(r, 0, [
            f'=IFERROR(INDEX([[Content Calendar.date]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.platform]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.title]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.status]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.owner]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.days_to_go]],{up}),"")',
        ], ["date", "text", "text", "status", "text", "int"])
        s.calc_row(r, 8, [
            f'=IFERROR(INDEX([[Content Calendar.date]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.platform]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.title]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.status]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.owner]],{late}),"")',
            f'=IFERROR(-INDEX([[Content Calendar.days_to_go]],{late}),"")',
        ], ["date", "text", "text", "status", "text", "int"])
    row = first_queue + 13

    # ------------------------------------------------------------- owners
    row = s.section(row, "Who is holding what", LAST_COL)
    s.table_head(row, 0, ["Owner", "Posts owned", "Still to publish",
                          "Open tasks", "Late posts", "On-time rate"])
    first_owner = row + 1
    owner_rows = min(10, cfg.row_count("Team"))
    for i in range(owner_rows):
        r = first_owner + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF([[Calc.ow_name#{cr}]]="","",'
        s.calc_row(r, 0, [
            f"={guard}[[Calc.ow_name#{cr}]])",
            f"={guard}[[Calc.ow_posts#{cr}]])",
            f"={guard}[[Calc.ow_open#{cr}]])",
            f"={guard}[[Calc.ow_tasks#{cr}]])",
            f"={guard}[[Team.late#{cr}]])",
            f"={guard}[[Team.on_time#{cr}]])",
        ], ["text", "int", "int", "int", "int", "percent"])
    ws.conditional_format(first_owner, 1, first_owner + owner_rows - 1, 1,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    ws.conditional_format(first_owner, 4, first_owner + owner_rows - 1, 4, {
        "type": "cell", "criteria": ">", "value": 0,
        "format": fmts.cf("bad")})
    row = first_owner + owner_rows + 1

    # ---------------------------------------------------------- checklist
    row = s.section(row, "The weekly routine that keeps this working",
                    LAST_COL)
    row = s.bullets(row, [
        "MONDAY - open the Dashboard, clear anything in the overdue list, "
        "then look at the next 12 posts due.",
        "MONDAY - batch the week's captions in the Caption Bank and pick "
        "a hashtag group for each post.",
        "TUESDAY or WEDNESDAY - make the content: shoot, film, design. "
        "Mark each row 'Draft' as it is finished.",
        "THURSDAY - review and approve. Check links, claims, spelling and "
        "the Compliance tab.",
        "FRIDAY - schedule everything, set the status to 'Scheduled' and "
        "add the publish links as they go live.",
        "ANY DAY - add ideas to the Idea Bank as they come to you. Future "
        "you will be grateful.",
        "FIRST OF THE MONTH - fill in the Analytics tab, update the Goals "
        "tab, and read the Monthly Dashboard.",
    ], LAST_COL)
    row += 1

    s.note(row,
           "The two deadline columns on the Content Calendar (draft due "
           "and approval due) are worked out from the publish date using "
           "the production and approval days on the Settings tab - change "
           "those two numbers to match how you really work.",
           LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `social_media_spreadsheet/sheets/brand_voice.py`

`Brand Voice` - how the brand sounds, looks and behaves online. *(165 lines)*

````python
"""`Brand Voice` - how the brand sounds, looks and behaves online."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from

NAME = "Brand Voice"
LAST_COL = 9

BLOCKS = [
    ("Who you are", [
        ("Mission - why the business exists",
         "We make everyday objects worth keeping."),
        ("Vision - where it is going",
         "A handmade piece on every table, not in every cabinet."),
        ("Brand promise - what people can count on",
         "Made slowly, built to be used, replaced if it fails."),
        ("Brand story in one paragraph",
         "Started at a night class, kept going because people kept "
         "asking. Everything is thrown, glazed and packed by hand."),
        ("What makes you different",
         "Small batches, named glazes, and the maker answers the "
         "messages."),
    ]),
    ("How you sound", [
        ("Three words that describe your voice", "Warm, honest, playful"),
        ("Three words you are not", "Corporate, pushy, precious"),
        ("Reading age / tone", "Chatty but clear - write like you talk"),
        ("Point of view", "First person: I made this, we packed this"),
        ("Humour", "Dry and self-deprecating, never at anyone's expense"),
        ("Emoji use", "One or two, never a wall"),
        ("Exclamation marks", "One per post at most"),
    ]),
    ("Words", [
        ("Words and phrases you use",
         "handmade, small batch, slow made, everyday, kiln, glaze"),
        ("Words and phrases you avoid",
         "cheap, luxury, must-have, hurry, limited time only"),
        ("How you describe the product",
         "Stoneware made in small batches, dishwasher and microwave "
         "safe."),
        ("How you describe the customer",
         "People who use the good stuff on a Tuesday."),
        ("Your sign-off / signature line", "Made slowly, used daily."),
    ]),
    ("Rules", [
        ("Always do", "Reply to every comment in the first hour"),
        ("Never do", "Argue in the comments, or repost without asking"),
        ("How you handle a complaint",
         "Reply once publicly, move it to messages, fix it, follow up"),
        ("How you handle a troll", "Hide, do not engage, block if repeated"),
        ("How you handle a mistake",
         "Say it plainly, say what you are doing about it"),
    ]),
    ("How you look", [
        ("Logo file / link", "drive/wildflower/logo"),
        ("Colour palette", "Clay #C8734A, Sand #EAE0D2, Ink #241B38"),
        ("Fonts", "Headlines: Fraunces. Body: Inter."),
        ("Photo style",
         "Daylight, linen, slight mess - real tables, not studios"),
        ("Video style", "Hands, process, no voiceover, captions always"),
        ("Grid / feed style", "Alternating product, process, person"),
        ("Template files", "canva.com/wildflower-templates"),
    ]),
    ("Profile copy", [
        ("Bio - short (150 characters)",
         "Handmade stoneware for everyday use. Made slowly in the "
         "studio. Shop below."),
        ("Bio - long",
         "Hand-thrown ceramics made in small batches. Dishwasher safe, "
         "built to be used, shipped carefully. New glazes monthly."),
        ("Link in bio destination", "shop.wildflowerstudio.com"),
        ("Pinned post / highlight",
         "Start here: how a mug is made, and how to choose a size"),
        ("Signature hashtag", "#wildflowerstudio"),
    ]),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)

    s.banner("Brand Voice & Identity",
             "Fill this in once. Every caption you write, every reply you "
             "send and anyone you hire should match what is on this page.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 26)
    s.set_col(3, LAST_COL, 16)

    row = 4
    for title, fields in BLOCKS:
        row = s.section(row, title, LAST_COL)
        for label, example in fields:
            value = example if cfg.include_sample_data else ""
            row = s.field(row, label, value, label_col=1, value_col=3,
                          value_end=LAST_COL, style="wrap")
        row += 1

    row = s.section(row, "Tone by situation", LAST_COL)
    s.table_head(row, 1, ["Situation", "How you sound", "Example line"],
                 [26, 20, 44])
    situations = [
        ("Launching something", "Confident, warm",
         "It is here, and there are only forty."),
        ("Teaching something", "Patient, plain",
         "Here is how to tell a good glaze from a lazy one."),
        ("Sharing a mistake", "Honest, light",
         "Forty mugs, one wobbly kiln shelf. We live and learn."),
        ("Celebrating a customer", "Generous, specific",
         "Look what Hannah did with the breakfast set."),
        ("Asking for the sale", "Direct, no apology",
         "The restock is live. Link in bio."),
        ("Replying to a complaint", "Calm, quick, human",
         "That should not have happened - I am sorting it today."),
    ]
    first = row + 1
    for i, (situation, tone, example) in enumerate(situations):
        r = first + i
        value_tone = tone if cfg.include_sample_data else ""
        value_example = example if cfg.include_sample_data else ""
        s.write(r, 1, situation, fmts.chrome("label_bold"))
        s.write(r, 2, value_tone, fmts.input("text"))
        s.merge(r, 3, LAST_COL, None, fmts.input("wrap"))
        s.write(r, 3, value_example, fmts.input("wrap"))
    row = first + len(situations) + 1

    row = s.section(row, "Consistency check", LAST_COL)
    s.table_head(row, 1, ["Check", "Your answer"], [40, 20])
    checks = [
        ("Does every profile use the same name and handle?",
         "YesNoList"),
        ("Do all five profile pictures match?", "YesNoList"),
        ("Is the bio the same tone everywhere?", "YesNoList"),
        ("Do you use the same colours and fonts on graphics?",
         "YesNoList"),
        ("Does your signature hashtag appear on every post?",
         "YesNoList"),
        ("Could someone else write a post that sounds like you?",
         "YesNoList"),
    ]
    first = row + 1
    for i, (check, list_name) in enumerate(checks):
        r = first + i
        s.merge(r, 1, 2, check, fmts.chrome("label"))
        s.write(r, 3, "", fmts.input("flag"))
        s.ws.data_validation(r, 3, r, 3, list_from(list_name))
        s.ws.conditional_format(r, 3, r, 3, {
            "type": "text", "criteria": "containing", "value": "No",
            "format": fmts.cf("warn")})
        s.ws.conditional_format(r, 3, r, 3, {
            "type": "text", "criteria": "containing", "value": "Yes",
            "format": fmts.cf("good")})
    row = first + len(checks) + 1

    s.note(row,
           "A brand people recognise is worth more than a brand people "
           "like. Pick a voice you can keep up on a bad week, write it "
           "down here, and stick to it.",
           LAST_COL, first_col=1, height=28)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `social_media_spreadsheet/sheets/reports.py`

`Reports` - pick two dates, print the result or send it to a client. *(267 lines)*

````python
"""`Reports` - pick two dates, print the result or send it to a client."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Reports"
LAST_COL = 13
FIRST_CALC_ROW = 5

_CAL = ('[[Content Calendar.date]],">="&RepFrom,'
        '[[Content Calendar.date]],"<="&RepTo')
_AN = '[[Analytics.month]],">="&RepFrom,[[Analytics.month]],"<="&RepTo'
_PF = '[[Performance.date]],">="&RepFrom,[[Performance.date]],"<="&RepTo'


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Reports & Exports",
             "Choose a date range, then print this page or copy it into "
             "an email. Everything updates from the calendar, the "
             "Analytics tab and the Performance tab.",
             LAST_COL)
    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 14)

    # -------------------------------------------------------- date range
    row = 4
    row = s.section(row, "Report period", LAST_COL)
    s.merge(row, 0, 0, "From", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.today.replace(day=1), fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("RepFrom", f"={quote_sheet(NAME)}!$B${row + 1}")

    s.merge(row, 2, 2, "To", fmts.chrome("label_bold"))
    s.write(row, 3, cfg.today, fmts.input("date"))
    ws.data_validation(row, 3, row, 3, date_check())
    book.define_name("RepTo", f"={quote_sheet(NAME)}!$D${row + 1}")

    s.merge(row, 4, 5, "Days in the range", fmts.chrome("label_bold"))
    s.write(row, 6, "=RepTo-RepFrom+1", fmts.calc("int"))
    s.merge(row, 7, 8, "Prepared on", fmts.chrome("label_bold"))
    s.write(row, 9, "=TODAY()", fmts.calc("date"))
    s.merge(row, 10, LAST_COL, "=BrandName", fmts.calc("text"))
    row += 2

    # ------------------------------------------------------------ summary
    row = s.section(row, "Headline numbers", LAST_COL)
    pairs = [
        ("Posts published",
         f'=COUNTIFS({_CAL},[[Content Calendar.status]],"Published")',
         "int"),
        ("Posts planned", f"=COUNTIFS({_CAL})", "int"),
        ("Reach", f"=SUMIFS([[Analytics.reach]],{_AN})", "int"),
        ("Impressions", f"=SUMIFS([[Analytics.impressions]],{_AN})",
         "int"),
        ("Engagements", f"=SUMIFS([[Analytics.engagements]],{_AN})",
         "int"),
        ("Engagement rate",
         f"=IFERROR(SUMIFS([[Analytics.engagements]],{_AN})"
         f"/SUMIFS([[Analytics.reach]],{_AN}),0)", "percent"),
        ("Followers gained", f"=SUMIFS([[Analytics.gained]],{_AN})", "int"),
        ("Unfollows", f"=SUMIFS([[Analytics.lost]],{_AN})", "int"),
        ("Net follower growth",
         f"=SUMIFS([[Analytics.net_growth]],{_AN})", "int"),
        ("Link clicks", f"=SUMIFS([[Analytics.clicks]],{_AN})", "int"),
        ("Leads", f"=SUMIFS([[Analytics.leads]],{_AN})", "int"),
        ("Sales", f"=SUMIFS([[Analytics.conversions]],{_AN})", "int"),
        ("Revenue", f"=SUMIFS([[Analytics.revenue]],{_AN})", "money"),
        ("Spend",
         '=SUMIFS([[Budget.actual]],[[Budget.month]],">="&RepFrom,'
         '[[Budget.month]],"<="&RepTo)', "money"),
        ("Return on spend",
         f'=IFERROR(SUMIFS([[Analytics.revenue]],{_AN})'
         f'/SUMIFS([[Budget.actual]],[[Budget.month]],">="&RepFrom,'
         f'[[Budget.month]],"<="&RepTo),0)', "number1"),
        ("Cost per lead",
         f'=IFERROR(SUMIFS([[Budget.actual]],[[Budget.month]],">="&RepFrom,'
         f'[[Budget.month]],"<="&RepTo)/SUMIFS([[Analytics.leads]],{_AN}),'
         f"0)", "money"),
    ]
    first_summary = row
    for i, (label, formula, style) in enumerate(pairs):
        r = first_summary + i % 8
        col = 0 if i < 8 else 5
        s.merge(r, col, col + 2, label, fmts.chrome("label_bold"))
        ws.merge_range(r, col + 3, r, col + 4, "", fmts.calc(style))
        ws.write_formula(r, col + 3, s.reg.render(formula),
                         fmts.calc(style))
    row = first_summary + 9

    # --------------------------------------------------------- platforms
    row = s.section(row, "By platform", LAST_COL)
    s.table_head(row, 0, ["Platform", "Posts", "Reach", "Impressions",
                          "Engagements", "Eng. rate", "Clicks", "Leads",
                          "Sales", "Revenue", "Followers gained"])
    first_platform = row + 1
    for i in range(8):
        r = first_platform + i
        cr = FIRST_CALC_ROW + i
        name = f"[[Calc.p_name#{cr}]]"
        guard = f'IF({name}="","",'
        s.calc_row(r, 0, [
            f"={guard}{name})",
            f'={guard}COUNTIFS([[Content Calendar.platform]],{name},'
            f'{_CAL},[[Content Calendar.status]],"Published"))',
            f"={guard}SUMIFS([[Analytics.reach]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.impressions]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.engagements]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}IFERROR(SUMIFS([[Analytics.engagements]],"
            f"[[Analytics.platform]],{name},{_AN})"
            f"/SUMIFS([[Analytics.reach]],[[Analytics.platform]],{name},"
            f'{_AN}),""))',
            f"={guard}SUMIFS([[Analytics.clicks]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.leads]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.conversions]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.revenue]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.net_growth]],"
            f"[[Analytics.platform]],{name},{_AN}))",
        ], ["text", "int", "int", "int", "int", "percent", "int", "int",
            "int", "money", "int"])
    ws.conditional_format(first_platform, 2, first_platform + 7, 2,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    row = first_platform + 9

    # ----------------------------------------------------------- pillars
    row = s.section(row, "By content pillar", LAST_COL)
    s.table_head(row, 0, ["Pillar", "Posts published", "Reach",
                          "Engagements", "Eng. rate", "Clicks", "Sales",
                          "Revenue"])
    first_pillar = row + 1
    pillar_rows = min(10, cfg.row_count("Content Pillars"))
    for i in range(pillar_rows):
        r = first_pillar + i
        cr = FIRST_CALC_ROW + i
        name = f"[[Calc.pl_name#{cr}]]"
        guard = f'IF({name}="","",'
        s.calc_row(r, 0, [
            f"={guard}{name})",
            f'={guard}COUNTIFS([[Content Calendar.pillar]],{name},'
            f'{_CAL},[[Content Calendar.status]],"Published"))',
            f"={guard}SUMIFS([[Performance.reach]],"
            f"[[Performance.pillar]],{name},{_PF}))",
            f"={guard}SUMIFS([[Performance.engagements]],"
            f"[[Performance.pillar]],{name},{_PF}))",
            f"={guard}IFERROR(SUMIFS([[Performance.engagements]],"
            f"[[Performance.pillar]],{name},{_PF})"
            f"/SUMIFS([[Performance.reach]],[[Performance.pillar]],"
            f'{name},{_PF}),""))',
            f"={guard}SUMIFS([[Performance.clicks]],"
            f"[[Performance.pillar]],{name},{_PF}))",
            f"={guard}SUMIFS([[Performance.conversions]],"
            f"[[Performance.pillar]],{name},{_PF}))",
            f"={guard}SUMIFS([[Performance.revenue]],"
            f"[[Performance.pillar]],{name},{_PF}))",
        ], ["text", "int", "int", "int", "percent", "int", "int", "money"])
    ws.conditional_format(first_pillar, 4, first_pillar + pillar_rows - 1,
                          4, {"type": "3_color_scale",
                              "min_color": "#FBD9DD",
                              "mid_color": "#FFF1CF",
                              "max_color": "#D9F2E4"})
    row = first_pillar + pillar_rows + 1

    # -------------------------------------------------------- top posts
    row = s.section(row, "Top 10 posts in this range", LAST_COL)
    s.table_head(row, 0, ["#", "Post", "Platform", "Pillar", "Date",
                          "Reach", "Engagements", "Eng. rate", "Clicks",
                          "Sales", "Revenue"])
    first_top = row + 1
    for i in range(10):
        r = first_top + i
        n = i + 1
        pick = (f"MATCH(LARGE([[Calc.pf_key_r]],{n}),"
                f"[[Calc.pf_key_r]],0)")
        s.calc_row(r, 0, [
            n,
            f'=IFERROR(INDEX([[Calc.pf_title]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_platform]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_pillar]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_date]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_reach]],{pick}),"")',
            f'=IFERROR(INDEX([[Performance.engagements]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_er]],{pick}),"")',
            f'=IFERROR(INDEX([[Performance.clicks]],{pick}),"")',
            f'=IFERROR(INDEX([[Performance.conversions]],{pick}),"")',
            f'=IFERROR(INDEX([[Performance.revenue]],{pick}),"")',
        ], ["int", "text", "text", "text", "date", "int", "int",
            "percent", "int", "int", "money"])
    row = first_top + 11

    # ------------------------------------------------------- campaigns
    row = s.section(row, "Campaigns in this range", LAST_COL)
    s.table_head(row, 0, ["Campaign", "Start", "End", "Posts", "Reach",
                          "Engagements", "Clicks", "Leads", "Sales",
                          "Spend", "Revenue", "Return"])
    first_campaign = row + 1
    campaign_rows = min(10, cfg.row_count("Campaigns"))
    for i in range(campaign_rows):
        r = first_campaign + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF([[Campaigns.name#{cr}]]="","",'
        s.calc_row(r, 0, [
            f"={guard}[[Campaigns.name#{cr}]])",
            f"={guard}[[Campaigns.start#{cr}]])",
            f"={guard}[[Campaigns.end#{cr}]])",
            f"={guard}[[Campaigns.posts#{cr}]])",
            f"={guard}[[Campaigns.reach#{cr}]])",
            f"={guard}[[Campaigns.engagements#{cr}]])",
            f"={guard}[[Campaigns.clicks#{cr}]])",
            f"={guard}[[Campaigns.leads#{cr}]])",
            f"={guard}[[Campaigns.conversions#{cr}]])",
            f"={guard}[[Campaigns.spend#{cr}]])",
            f"={guard}[[Campaigns.revenue#{cr}]])",
            f"={guard}[[Campaigns.roas#{cr}]])",
        ], ["text", "date", "date", "int", "int", "int", "int", "int",
            "int", "money", "money", "number1"])
    row = first_campaign + campaign_rows + 1

    # ----------------------------------------------------- what to say
    row = s.section(row, "What this says, in a sentence", LAST_COL)
    sentence = (
            '=IF(SUMIFS([[Analytics.reach]],' + _AN + ')=0,'
            '"Fill in the Analytics tab to see a summary here.",'
            '"In this period you published "&'
            'TEXT(COUNTIFS(' + _CAL + ',[[Content Calendar.status]],'
            '"Published"),"#,##0")&" posts, reached "&'
            'TEXT(SUMIFS([[Analytics.reach]],' + _AN + '),"#,##0")&'
            '" people, earned "&TEXT(SUMIFS('
            '[[Analytics.engagements]],' + _AN + '),"#,##0")&'
            '" engagements ("&TEXT(IFERROR(SUMIFS('
            '[[Analytics.engagements]],' + _AN + ')/SUMIFS('
            '[[Analytics.reach]],' + _AN + '),0),"0.0%")&'
            '") and gained "&TEXT(SUMIFS([[Analytics.net_growth]],'
            + _AN + '),"#,##0")&" followers.")')
    s.merge(row, 0, LAST_COL, s.reg.render(sentence), fmts.calc("wrap"))
    s.row_height(row, 32)
    row += 2

    row = s.section(row, "How to share this", LAST_COL)
    row = s.bullets(row, [
        "PRINT - File > Print. The page is already set to fit one page "
        "wide; choose landscape for the widest tables.",
        "PDF - File > Save As > PDF, or print to PDF. Perfect for a "
        "client or a monthly review.",
        "COPY - select the area you want, copy, then paste into an email "
        "as a picture so the formatting survives.",
        "EXPORT - File > Save As > CSV exports the active sheet only. "
        "Use it to move data into another tool.",
        "SHARE - keep the file in Drive, OneDrive or Dropbox so you "
        "always open the same copy on every device.",
    ], LAST_COL)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    ws.set_landscape()
    return s
````

### `social_media_spreadsheet/sheets/start_here.py`

`Start Here` - the welcome page and the six step set-up. *(176 lines)*

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
     "Put your brand name, handles, plan year, targets and budget in. "
     "Ten minutes, and it feeds every other tab."),
    ("2", "Accounts & Brand Voice",
     "List every profile you post to, then write down how the brand "
     "sounds so captions stay consistent."),
    ("3", "Personas & Content Pillars",
     "Decide who you are talking to and the five or six themes you post "
     "about. The pillars become drop downs everywhere else."),
    ("4", "Idea Bank, Captions & Hashtags",
     "Dump every idea in, write captions you can reuse and build "
     "hashtag groups of 20-30 tags."),
    ("5", "Content Calendar",
     "Plan the year. One row per post: date, platform, title, pillar, "
     "caption, hashtag group, asset, owner, status."),
    ("6", "Analytics & Performance",
     "On the first of each month enter the account numbers, and after "
     "each post record how it did. The dashboards do the rest."),
]

TABS = [
    ("Dashboard", "This month at a glance, with alerts and the queue"),
    ("Annual Dashboard", "The whole year, quarter by quarter"),
    ("Settings", "Your brand, targets, budget and reminder windows"),
    ("Lists", "Every drop down list - edit these to suit your business"),
    ("Brand Voice", "How the brand sounds, looks and behaves"),
    ("Audience Personas", "Who you are talking to"),
    ("Accounts", "Every profile, with a six point optimisation check"),
    ("Content Pillars", "Your themes and the mix you planned"),
    ("Content Calendar", "THE MAIN TAB - one row per post, all year"),
    ("Month View", "The same calendar laid out as a printable grid"),
    ("Workflow", "The production pipeline: what is where, what is late"),
    ("Idea Bank", "Scored ideas waiting to be scheduled"),
    ("Caption Bank", "Reusable hooks, captions and calls to action"),
    ("Hashtag Bank", "Hashtag groups, keywords and search intent"),
    ("Instagram / TikTok / Pinterest / YouTube / Facebook",
     "Platform specific fields - hooks, sounds, boards, thumbnails"),
    ("Campaigns", "Themed pushes with a budget and a return"),
    ("Launch Planner", "Teasing, launch day and follow-up"),
    ("Seasonal Planner", "Holidays and awareness days, planned early"),
    ("Influencers", "Outreach, deliverables, fees and results"),
    ("UGC Tracker", "Customer content, permission and rights"),
    ("Repurposing", "One good post, used five more times"),
    ("Competitors", "What the people in your niche are doing"),
    ("Trends", "Sounds and formats worth jumping on"),
    ("Analytics", "Monthly numbers per account - type these in"),
    ("Performance", "Post by post results, ranked and scored"),
    ("AB Tests", "Change one thing, measure it, keep the winner"),
    ("Ads", "Paid promotion, cost and return"),
    ("Assets", "The creative library with links and rights"),
    ("Goals", "The numbers you are trying to move"),
    ("Budget", "What social media costs you, month by month"),
    ("Team", "Who does what"),
    ("Tasks", "The jobs behind the posts"),
    ("Compliance", "Disclosure, rights, accessibility and quality"),
    ("Reports", "Pick two dates and print the result"),
    ("Print Brief", "One post, printed as a brief for whoever makes it"),
    ("Print Week", "Seven days on one page for the wall"),
    ("Help", "Every question answered, including the honest limits"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)

    s.banner(cfg.product_name,
             "Plan a year of content, keep every caption and hashtag in "
             "one place, and see what actually worked.",
             LAST_COL)
    s.set_col(0, 0, 4)
    s.set_col(1, 1, 22)
    s.set_col(2, 2, 30)
    for c in range(3, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.note(
        row,
        "Everything in this workbook is linked. Type something once - a "
        "post, a caption, a number from Instagram - and it appears "
        "wherever it is needed: the dashboards, the reports, the printable "
        "brief and the alerts. Cream cells are yours to fill in; grey "
        "cells are formulas and are locked so they cannot be broken by "
        "accident.",
        LAST_COL, first_col=1, height=46)
    row += 1

    row = s.section(row, "Six steps to set it up", LAST_COL)
    s.table_head(row, 1, ["Step", "Where", "What to do"], [8, 24, 70])
    first = row + 1
    for i, (number, where, what) in enumerate(STEPS):
        r = first + i
        s.row_height(r, 30)
        s.write(r, 1, number, fmts.chrome("kpi_label"))
        s.write(r, 2, where, fmts.chrome("label_bold"))
        s.merge(r, 3, LAST_COL, what, fmts.chrome("note"))
    row = first + len(STEPS) + 1

    row = s.section(row, "The weekly rhythm", LAST_COL)
    row = s.bullets(row, [
        "MONDAY - clear the overdue list on the Dashboard, then look at "
        "the next posts due on the Workflow tab.",
        "ONE BATCH DAY - make the week's content in one sitting. It is "
        "always faster than making it one post at a time.",
        "BEFORE YOU POST - check the caption, the hashtag group, the "
        "asset and the link. The calendar flags anything missing.",
        "AFTER YOU POST - paste the link into the calendar. Two days "
        "later, add the numbers on the Performance tab.",
        "FIRST OF THE MONTH - fill in the Analytics tab (ten minutes), "
        "update your Goals, then read the Monthly Dashboard.",
    ], LAST_COL, first_col=1)
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
        "ADDING ROWS - every register has spare rows ready to use. Right "
        "click a row inside the list and choose Insert to add more; the "
        "formulas copy down.",
        "SORTING - use the filter arrows in the header row rather than "
        "re-sorting the sheet, so the content IDs keep matching the "
        "other tabs.",
        "DROP DOWNS - every list lives on the Lists tab. Edit it there "
        "and the drop downs follow.",
        "GOOGLE SHEETS - File > Import works, and everything except the "
        "sparklines and the sheet protection carries over.",
        "YOUR CURRENCY - set the symbol on the Settings tab, or rebuild "
        "the workbook with a different currency.",
    ], LAST_COL, first_col=1, height=28)
    row += 1

    row = s.section(row, "What a spreadsheet can and cannot do", LAST_COL)
    row = s.note(
        row,
        "Being straight with you: this workbook organises your content "
        "and does the maths on the numbers you give it. It cannot pull "
        "your analytics automatically, it cannot publish or schedule "
        "posts, and it does not talk to Instagram, TikTok, Pinterest, "
        "YouTube or Facebook. You copy the headline numbers in once a "
        "month - about ten minutes - and in return you get a plan, a "
        "history and a set of dashboards no scheduling app will give "
        "you. If you want automatic posting, pair this with a scheduler; "
        "if you want automatic analytics, export a CSV from the platform "
        "and paste it in.",
        LAST_COL, first_col=1, height=70)
    row += 1

    s.merge(row, 1, LAST_COL,
            f"{cfg.product_name}  |  {cfg.author}  |  version "
            f"{cfg.version}", fmts.chrome("print_sub"))

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

### `social_media_spreadsheet/sheets/printables.py`

The print-ready tabs: a single content brief and a weekly plan. *(271 lines)*

````python
"""The print-ready tabs: a single content brief and a weekly plan.

Both are driven by a drop down at the top of the page, so you pick the
post (or the week) and press print.
"""

from __future__ import annotations

from typing import Optional, Tuple

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

BRIEF_SHEET = "Print Brief"
WEEK_SHEET = "Print Week"

LAST_COL = 7          # column H - a comfortable portrait page


def _pick(sheet: SheetWriter, fmts, label: str, list_name: str,
          define_as: str, hint: str, default: str = "") -> None:
    """Drop down selector on row 1 (it sits outside the print area)."""
    ws = sheet.ws
    sheet.merge(0, 0, 1, label, fmts.chrome("label_bold"))
    fmt = fmts.input("text", bold=True, font_size=11)
    ws.merge_range(0, 2, 0, 3, default, fmt)
    ws.data_validation(0, 2, 0, 3, {
        "validate": "list", "source": f"={list_name}",
        "input_title": "Choose a record",
        "input_message": hint[:255]})
    sheet.merge(0, 4, LAST_COL, hint, fmts.chrome("note"))
    sheet.row_height(0, 22)
    sheet.book.define_name(define_as,
                           f"={quote_sheet(sheet.name)}!$C$1")


def _line(sheet, fmts, row: int, label: str,
          formula: Optional[str] = None, style: str = "text",
          label_cols: Tuple[int, int] = (0, 1),
          value_cols: Tuple[int, int] = (2, LAST_COL),
          height: Optional[float] = None) -> int:
    sheet.merge(row, label_cols[0], label_cols[1], label,
                fmts.chrome("label_bold"))
    fmt = fmts.calc(style)
    if value_cols[1] > value_cols[0]:
        sheet.ws.merge_range(row, value_cols[0], row, value_cols[1], "",
                             fmt)
    if formula:
        sheet.ws.write_formula(row, value_cols[0],
                               sheet.reg.render(formula), fmt)
    if height:
        sheet.row_height(row, height)
    return row + 1


def _blank_box(sheet, fmts, row: int, label: str, lines: int = 3) -> int:
    sheet.merge(row, 0, 1, label, fmts.chrome("label_bold"))
    sheet.ws.merge_range(row, 2, row + lines - 1, LAST_COL, "",
                         fmts.chrome("form_box"))
    for i in range(lines):
        sheet.row_height(row + i, 16)
    return row + lines


# ---------------------------------------------------------------------------
# Print Brief - one post, everything the maker needs
# ---------------------------------------------------------------------------

def build_brief(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, BRIEF_SHEET,
                    tab_color=Palette.accent)
    ws = s.ws
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 16)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 13)

    default = "C-0001" if cfg.include_sample_data else ""
    _pick(s, fmts, "Content ID", "ContentIDs", "BriefID",
          "Pick a post from the Content Calendar - the whole brief fills "
          "itself in. This row does not print.", default)

    def pull(column: str) -> str:
        return (f'=IF(BriefID="","",IFERROR(INDEX('
                f'[[Content Calendar.{column}]],MATCH(BriefID,'
                f'[[Content Calendar.id]],0)),"Not found"))')

    row = 2
    s.merge(row, 0, LAST_COL, "CONTENT BRIEF", fmts.chrome("print_title"))
    s.row_height(row, 26)
    row += 1
    s.merge(row, 0, LAST_COL, "=BrandName&\"  |  \"&TEXT(TODAY(),\"d mmmm "
                              "yyyy\")", fmts.chrome("print_sub"))
    row += 2

    row = s.section(row, "The post", LAST_COL)
    row = _line(s, fmts, row, "Content ID", '=IF(BriefID="","",BriefID)')
    row = _line(s, fmts, row, "Title", pull("title"))
    row = _line(s, fmts, row, "Publish date", pull("date"), "date")
    row = _line(s, fmts, row, "Time slot", pull("slot"))
    row = _line(s, fmts, row, "Platform", pull("platform"))
    row = _line(s, fmts, row, "Format", pull("format"))
    row = _line(s, fmts, row, "Owner", pull("owner"))
    row = _line(s, fmts, row, "Status", pull("status"), "status")
    row += 1

    row = s.section(row, "Why we are making it", LAST_COL)
    row = _line(s, fmts, row, "Content pillar", pull("pillar"))
    row = _line(s, fmts, row, "Objective", pull("objective"))
    row = _line(s, fmts, row, "Funnel stage", pull("funnel"))
    row = _line(s, fmts, row, "Written for", pull("persona"))
    row = _line(s, fmts, row, "Campaign", pull("campaign"))
    row = _line(s, fmts, row, "What it is about", pull("description"),
                "wrap", height=34)
    row += 1

    row = s.section(row, "What to say", LAST_COL)
    caption = ('=IF(BriefID="","",IFERROR(INDEX([[Caption Bank.%s]],'
               'MATCH(INDEX([[Content Calendar.caption]],MATCH(BriefID,'
               '[[Content Calendar.id]],0)),[[Caption Bank.title]],0)),'
               '"- add a caption on the Caption Bank tab -"))')
    row = _line(s, fmts, row, "Caption name", pull("caption"))
    row = _line(s, fmts, row, "Hook", caption % "hook", "wrap", height=30)
    row = _line(s, fmts, row, "Caption", caption % "full", "wrap",
                height=70)
    row = _line(s, fmts, row, "Call to action", pull("cta"))
    row = _line(s, fmts, row, "Hashtag group", pull("hashtag_group"))
    row = _line(s, fmts, row, "Tags in the group", pull("hashtag_count"),
                "int")
    row += 1

    row = s.section(row, "What to make", LAST_COL)
    asset = ('=IF(BriefID="","",IFERROR(INDEX([[Assets.%s]],MATCH('
             'INDEX([[Content Calendar.asset]],MATCH(BriefID,'
             '[[Content Calendar.id]],0)),[[Assets.id]],0)),'
             '"- no asset linked -"))')
    row = _line(s, fmts, row, "Asset ID", pull("asset"))
    row = _line(s, fmts, row, "Asset name", asset % "title")
    row = _line(s, fmts, row, "File link", asset % "link", "wrap")
    row = _line(s, fmts, row, "Size", asset % "dimensions")
    row = _line(s, fmts, row, "Draft due", pull("draft_due"), "date")
    row = _line(s, fmts, row, "Approval due", pull("approve_due"), "date")
    row += 1

    row = s.section(row, "Notes & shot list (write on the printout)",
                    LAST_COL)
    row = _line(s, fmts, row, "Notes from the calendar", pull("notes"),
                "wrap", height=30)
    row = _blank_box(s, fmts, row, "Shot list", 4)
    row = _blank_box(s, fmts, row, "Props / setup", 3)
    row = _blank_box(s, fmts, row, "Approved by", 2)
    row += 1

    row = s.section(row, "After it goes live", LAST_COL)
    row = _line(s, fmts, row, "Published link", pull("url"), "wrap")
    row = _line(s, fmts, row, "Reach", pull("reach"), "int")
    row = _line(s, fmts, row, "Engagements", pull("engagements"), "int")
    row = _line(s, fmts, row, "Engagement rate", pull("er"), "percent")
    row += 1
    s.note(row,
           "Record the numbers on the Performance tab 48-72 hours after "
           "publishing - they appear here and on the calendar "
           "automatically.",
           LAST_COL, height=26)
    row += 2

    s.finish(freeze_row=None, gridlines=False, fit_width=True)
    ws.print_area(f"A3:{chr(65 + LAST_COL)}{row}")
    ws.set_portrait()
    ws.center_horizontally()
    return s


# ---------------------------------------------------------------------------
# Print Week - seven days of posts on one page
# ---------------------------------------------------------------------------

def build_week(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, WEEK_SHEET,
                    tab_color=Palette.accent)
    ws = s.ws
    s.set_col(0, 0, 13)
    s.set_col(1, 1, 13)
    s.set_col(2, 2, 34)
    s.set_col(3, 3, 16)
    s.set_col(4, 4, 18)
    s.set_col(5, 5, 16)
    s.set_col(6, 6, 14)
    s.set_col(7, 7, 14)

    s.merge(0, 0, 1, "Week starting", fmts.chrome("label_bold"))
    s.write(0, 2, cfg.today, fmts.input("date", bold=True))
    ws.data_validation(0, 2, 0, 2, date_check())
    book.define_name("WeekStartDate",
                     f"={quote_sheet(WEEK_SHEET)}!$C$1")
    s.merge(0, 3, LAST_COL,
            "Pick any date - the sheet prints the seven days that follow. "
            "This row does not print.", fmts.chrome("note"))
    s.row_height(0, 22)

    row = 2
    s.merge(row, 0, LAST_COL, "WEEKLY CONTENT PLAN",
            fmts.chrome("print_title"))
    s.row_height(row, 26)
    row += 1
    s.merge(row, 0, LAST_COL,
            '=BrandName&"  |  "&TEXT(WeekStartDate,"d mmm")&" to "&'
            'TEXT(WeekStartDate+6,"d mmm yyyy")',
            fmts.chrome("print_sub"))
    row += 2

    key = "[[Content Calendar.day_key]]"
    per_day = 4
    for day in range(7):
        day_cell = f"WeekStartDate+{day}"
        row = s.section(row, f'=TEXT({day_cell},"dddd d mmmm")', LAST_COL)
        s.table_head(row, 0, ["Time", "Platform", "Post", "Format",
                              "Pillar", "Owner", "Status", "Done"])
        first = row + 1
        for n in range(1, per_day + 1):
            r = first + n - 1
            pick = (f'SMALL({key},COUNTIF({key},"<"&{day_cell})+{n})')
            base = (f'=IFERROR(IF({pick}>={day_cell}+1,"",INDEX('
                    f'[[Content Calendar.%s]],MATCH({pick},{key},0))),"")')
            s.calc_row(r, 0, [
                base % "slot",
                base % "platform",
                base % "title",
                base % "format",
                base % "pillar",
                base % "owner",
                base % "status",
                "",
            ], ["text", "text", "text", "text", "text", "text", "status",
                "text"])
            ws.write_blank(r, 7, None, fmts.chrome("form_line"))
        row = first + per_day

    row += 1
    row = s.section(row, "This week in numbers", LAST_COL)
    window = ('[[Content Calendar.date]],">="&WeekStartDate,'
              '[[Content Calendar.date]],"<="&WeekStartDate+6')
    s.table_head(row, 0, ["Posts", "Published", "Still to do",
                          "Weekly target", "Against target", "Platforms",
                          "Captions missing", "Assets missing"])
    s.calc_row(row + 1, 0, [
        f"=COUNTIFS({window})",
        f'=COUNTIFS({window},[[Content Calendar.status]],"Published")',
        f'=COUNTIFS({window})-COUNTIFS({window},'
        f'[[Content Calendar.status]],"Published")',
        "=PostsPerWeek",
        f'=IFERROR(COUNTIFS({window})/PostsPerWeek,"")',
        f'=SUMPRODUCT((COUNTIFS([[Content Calendar.platform]],'
        f'[[Calc.p_name]],{window})>0)*1)',
        f'=COUNTIFS({window},[[Content Calendar.caption]],"")',
        f'=COUNTIFS({window},[[Content Calendar.asset]],"")',
    ], ["int", "int", "int", "int", "percent", "int", "int", "int"])
    row += 3

    s.note(row,
           "Print this on a Monday morning, stick it on the wall and tick "
           "the last column as each post goes out.",
           LAST_COL, height=24)
    row += 2

    s.finish(freeze_row=None, gridlines=False, fit_width=True)
    ws.print_area(f"A3:{chr(65 + LAST_COL)}{row}")
    ws.set_landscape()
    ws.center_horizontally()
    return s
````

### `social_media_spreadsheet/sheets/help_tab.py`

`Help` - how everything works, what every metric means, and the limits. *(241 lines)*

````python
"""`Help` - how everything works, what every metric means, and the limits."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Help"
LAST_COL = 11

HOW_IT_FITS = [
    ("Content Calendar",
     "The master list. Every post gets a row and an ID (C-0001). "
     "Planning only - no results are typed here."),
    ("Performance",
     "The results of individual posts, keyed on that same ID. Type the "
     "numbers once here and they appear on the calendar, the platform "
     "planners, the pillar table and every dashboard."),
    ("Analytics",
     "Account level totals, one row per account per month - followers, "
     "reach, clicks, leads, revenue. This is what the monthly and "
     "annual dashboards are built from."),
    ("Platform planners",
     "The extra fields that only matter on one platform: carousel "
     "slides, sounds, board names, thumbnails. Shared metrics are "
     "pulled in, never retyped."),
]

FAQ = [
    ("Where do I start?",
     "Settings, then Accounts, then Content Pillars, then the Content "
     "Calendar. The Start Here tab walks through it in six steps."),
    ("Do I have to fill everything in?",
     "No. The calendar, the caption bank and the analytics tab give you "
     "most of the value. The rest is there when you grow into it."),
    ("Why can I not type in a grey cell?",
     "Grey cells are formulas and are locked so they cannot be broken. "
     "Review > Unprotect Sheet, password '{password}', if you really "
     "want to change one."),
    ("How do I add more rows?",
     "Right click a row inside the list and choose Insert. The formulas "
     "copy down automatically. Every sheet also starts with spare rows."),
    ("Can I change the drop down options?",
     "Yes - they all live on the Lists tab. Edit, add or delete options "
     "there and every drop down updates."),
    ("Why does my post not show on the dashboard?",
     "Check the publish date is inside the month you picked at the top "
     "of the Dashboard, and that the year matches the plan year on the "
     "Settings tab."),
    ("Why is the engagement rate blank?",
     "It needs reach. Add the reach figure on the Performance tab (for "
     "a single post) or the Analytics tab (for a month)."),
    ("Reach or impressions - which should I use?",
     "Either, as long as you are consistent. Reach counts people, "
     "impressions count views. Rates calculated on reach are higher."),
    ("What is a good engagement rate?",
     "It depends on the platform and your size, but 1-3% is normal, "
     "3-6% is good and anything higher is excellent. Set your own "
     "target on the Settings tab."),
    ("How often should I post?",
     "Consistently beats often. Pick a number you can keep up on a bad "
     "week, set it as your weekly target, and keep it."),
    ("Can I use this for a client?",
     "Yes. Use one copy per client, change the brand details on "
     "Settings, and send them the Reports tab as a PDF each month."),
    ("Does it work in Google Sheets?",
     "Yes - File > Import > Upload. Everything works except the "
     "sparklines and the sheet protection, which Google handles "
     "differently."),
    ("Does it work on a phone or tablet?",
     "It opens in the Excel and Google Sheets apps. Entering data is "
     "fiddly on a phone; the dashboards read fine."),
    ("Can I change the currency?",
     "Yes - the currency symbol is on the Settings tab."),
    ("Something looks wrong - how do I check the maths?",
     "Unhide the Calc tab (right click any tab > Unhide). Every "
     "dashboard number is a plain SUMIFS or COUNTIFS on that sheet."),
]

METRICS = [
    ("Reach", "How many different people saw the post."),
    ("Impressions", "How many times it was shown, including repeats."),
    ("Engagements", "Likes + comments + shares + saves."),
    ("Engagement rate", "Engagements divided by reach. The main one."),
    ("Save rate",
     "Saves divided by reach. The strongest signal on most platforms - "
     "a save means 'I will come back to this'."),
    ("Share rate",
     "Shares divided by reach. Shares are how accounts grow."),
    ("Click rate", "Link clicks divided by reach."),
    ("Follower growth", "New followers less unfollows."),
    ("Reach vs followers",
     "Reach divided by your follower count. Over 100% means you went "
     "beyond your own audience."),
    ("Completion rate", "The share of viewers who watched to the end."),
    ("Thumbnail CTR",
     "On YouTube: how many people who saw the thumbnail clicked it. "
     "4-6% is healthy."),
    ("Cost per lead", "Spend divided by leads."),
    ("Return on spend (ROAS)",
     "Revenue divided by spend. 3.0 means three back for every one in."),
    ("Conversion rate", "Sales divided by clicks."),
]

TROUBLE = [
    ("#DIV/0! or a blank rate",
     "Reach is zero or missing. Add the reach figure."),
    ("#N/A in a lookup column",
     "The ID typed does not exist. Pick it from the drop down instead "
     "of typing it."),
    ("'Not found' in a grey column",
     "Same thing: the content ID, caption name or asset ID does not "
     "match anything. Check for a stray space."),
    ("A drop down is empty",
     "Nothing has been entered on the source tab yet - add a team "
     "member, a pillar or a campaign first."),
    ("The month filter shows nothing",
     "The plan year on Settings does not match the dates on your "
     "calendar."),
    ("Dates look like numbers",
     "Format the column as a date: Home > Number > Short Date."),
    ("A number will not type in",
     "The cell is locked because it is calculated. Check you are in a "
     "cream cell, not a grey one."),
    ("Everything recalculates slowly",
     "Formulas > Calculation Options > Manual, then press F9 when you "
     "want to refresh."),
]

LIMITS = [
    "It cannot fetch your analytics automatically. No spreadsheet can - "
    "that needs a platform API connection. You copy about fifteen "
    "numbers per account once a month.",
    "It cannot publish or schedule posts. Use a scheduler for that and "
    "keep this as the plan and the record.",
    "It cannot see comments, messages or mentions.",
    "Revenue attribution is only as good as the numbers you give it. "
    "Use UTM links and your shop's own reports.",
    "It is not legal advice. The Compliance tab is a prompt list, not a "
    "substitute for the rules where you live.",
    "It does not know what will go viral. Nothing does. What it does do "
    "is show you what worked for you, so you can do more of it.",
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.grey_text)

    s.banner("Help & Instructions",
             "How the workbook fits together, what every number means, "
             "and an honest list of what a spreadsheet cannot do.",
             LAST_COL)
    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.section(row, "How the main tabs fit together", LAST_COL)
    s.table_head(row, 1, ["Tab", "What it holds"], [30, 76])
    first = row + 1
    for i, (tab, text) in enumerate(HOW_IT_FITS):
        r = first + i
        s.row_height(r, 34)
        s.write(r, 1, tab, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL, text, fmts.chrome("note"))
    row = first + len(HOW_IT_FITS)
    row = s.note(
        row,
        "The rule: no number is ever typed twice. A post is planned on "
        "the Content Calendar, its results are recorded once on the "
        "Performance tab, and your account totals are recorded once a "
        "month on the Analytics tab. Everything else is a formula.",
        LAST_COL, first_col=1, height=34)
    row += 1

    row = s.section(row, "Questions people ask", LAST_COL)
    s.table_head(row, 1, ["Question", "Answer"], [30, 76])
    first = row + 1
    for i, (question, answer) in enumerate(FAQ):
        r = first + i
        s.row_height(r, 30)
        s.write(r, 1, question, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL,
                answer.replace("{password}", cfg.password),
                fmts.chrome("note"))
    row = first + len(FAQ) + 1

    row = s.section(row, "What every metric means", LAST_COL)
    s.table_head(row, 1, ["Metric", "In plain English"], [30, 76])
    first = row + 1
    for i, (metric, meaning) in enumerate(METRICS):
        r = first + i
        s.row_height(r, 24)
        s.write(r, 1, metric, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL, meaning, fmts.chrome("note"))
    row = first + len(METRICS) + 1

    row = s.section(row, "If something looks wrong", LAST_COL)
    s.table_head(row, 1, ["What you see", "What to do"], [30, 76])
    first = row + 1
    for i, (symptom, fix) in enumerate(TROUBLE):
        r = first + i
        s.row_height(r, 26)
        s.write(r, 1, symptom, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL, fix, fmts.chrome("note"))
    row = first + len(TROUBLE) + 1

    row = s.section(row, "Honest limitations", LAST_COL)
    for text in LIMITS:
        row = s.note(row, text, LAST_COL, first_col=1, height=28)
    row += 1

    row = s.section(row, "Colour code", LAST_COL)
    legend = [
        ("Cream", "legend_input", "Type here - your input"),
        ("Grey", "legend_calc", "Calculated and locked"),
        ("Green", "good", "On target, approved, done"),
        ("Amber", "warn", "Due soon, needs a look"),
        ("Red", "bad", "Overdue, over budget, below target"),
        ("Blue", "info", "Today, or for information"),
    ]
    for i, (label, fmt_key, meaning) in enumerate(legend):
        r = row + i
        fmt = (fmts.chrome(fmt_key) if fmt_key.startswith("legend")
               else fmts.cf(fmt_key))
        s.write(r, 1, label, fmt)
        s.merge(r, 2, LAST_COL, meaning, fmts.chrome("label"))
    row += len(legend) + 1

    row = s.note(
        row,
        f"{cfg.product_name} - created by {cfg.author}. Version "
        f"{cfg.version}. Sheet protection password: '{cfg.password}'. "
        f"Keep a spare copy of the original file before you make big "
        f"changes.",
        LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
````

---

## 7. Quality checks (optional)

Not needed to build the workbook - these are the checks that keep it honest. The first two need `openpyxl`, the third needs `formulas`.

### `tools/validate_social_workbook.py`

Static checks on a generated Social Media Planner workbook. *(297 lines)*

````python
"""Static checks on a generated Social Media Planner workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_social_workbook.py output/planner.xlsx

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
from social_media_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402

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
    if hidden != ["Calc"]:
        problems.append(f"hidden sheets should be ['Calc'], got {hidden}")

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
                    if cell.alignment and cell.protection and \
                            not cell.protection.locked:
                        locked_formula_problem += 1
                        if locked_formula_problem < 25:
                            problems.append(
                                f"{ws.title}!{cell.coordinate}: formula cell "
                                f"is NOT locked")
                    stripped = re.sub(r'"[^"]*"', '""', value)
                    stripped = re.sub(r"'[^']*'!", "SHEET!", stripped)
                    if value.count('"') % 2:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: odd number of "
                            f"quotes -> {value[:80]}")
                    if stripped.count("(") != stripped.count(")"):
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unbalanced "
                            f"brackets -> {value[:80]}")
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
              else "output/Social Media Content Planner.xlsx")
    raise SystemExit(check(target))
````

### `tools/check_social_sample.py`

Cross-check the example records in a generated social media workbook. *(209 lines)*

````python
"""Cross-check the example records in a generated social media workbook.

The demo data has to join up, otherwise the dashboards open full of zeros
and the lookup columns fill with "Not found". This reads the literal
values out of the finished file and makes sure every content ID, caption
name, hashtag group, asset ID, campaign, pillar, persona, account and
owner points at a record that really exists.

    python tools/check_social_sample.py output/test.xlsx
"""

from __future__ import annotations

import sys
from collections import Counter

import openpyxl

HEADER_ROW = 4


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
    for row in range(header_row + 1, ws.max_row + 1):
        value = ws.cell(row=row, column=col).value
        if value is None:
            continue
        if isinstance(value, str):
            value = value.strip()
            if not value or value.startswith("="):
                continue
        out.append(value)
    return out


def id_range(ws, id_header: str, name_header: str,
             header_row: int = HEADER_ROW) -> int:
    """How many example rows a register holds (its IDs are formulas)."""
    return len(column_values(ws, name_header, header_row))


def check(path: str) -> int:
    wb = openpyxl.load_workbook(path, data_only=False)
    problems: list[str] = []
    notes: list[str] = []

    def sheet(name):
        if name not in wb.sheetnames:
            problems.append(f"missing sheet {name}")
            return None
        return wb[name]

    cal = sheet("Content Calendar")
    perf = sheet("Performance")
    if cal is None or perf is None:
        for problem in problems:
            print(f"    - {problem}")
        return 1

    titles = column_values(cal, "Content title")
    notes.append(f"{len(titles)} example posts on the calendar")
    duplicates = [i for i, n in Counter(titles).items() if n > 1]
    if duplicates:
        problems.append(f"duplicate post titles: {duplicates[:5]}")

    def id_link(from_sheet: str, from_col: str, to_sheet: str,
                count_col: str) -> None:
        """Check IDs that are generated by a formula (C-0001, AS-0007).

        The ID column itself holds no literal values, so the check is
        that the number inside the reference is a row that exists.
        """
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
                f"past the {limit} rows on {to_sheet} -> {sorted(set(bad))[:5]}")
        else:
            notes.append(f"{from_sheet}.{from_col} -> {to_sheet} "
                         f"({len(values)} rows, max {limit}) ok")

    def link(from_sheet: str, from_col: str, to_sheet: str, to_col: str,
             header_row: int = HEADER_ROW,
             to_header_row: int = HEADER_ROW) -> None:
        src, dst = sheet(from_sheet), sheet(to_sheet)
        if src is None or dst is None:
            return
        try:
            values = column_values(src, from_col, header_row)
            allowed = set(column_values(dst, to_col, to_header_row))
        except KeyError as error:
            problems.append(str(error))
            return
        missing = sorted({str(v) for v in values} - {str(a)
                                                     for a in allowed})
        if missing:
            problems.append(
                f"{from_sheet}.{from_col}: {len(missing)} value(s) not in "
                f"{to_sheet}.{to_col} -> {missing[:5]}")
        else:
            notes.append(f"{from_sheet}.{from_col} -> {to_sheet}.{to_col} "
                         f"({len(values)} rows) ok")

    # the spine: every result belongs to a planned post
    id_link("Performance", "Content ID", "Content Calendar",
            "Content title")
    id_link("Repurposing", "Original content ID", "Content Calendar",
            "Content title")
    id_link("Content Calendar", "Creative asset", "Assets", "Asset name")
    for platform in ("Instagram", "TikTok", "Pinterest", "YouTube",
                     "Facebook"):
        id_link(platform, "Content ID", "Content Calendar",
                "Content title")

    link("Content Calendar", "Content pillar", "Content Pillars",
         "Pillar name")
    link("Content Calendar", "Caption (from the bank)", "Caption Bank",
         "Caption name")
    link("Content Calendar", "Hashtag group", "Hashtag Bank", "Group")
    link("Content Calendar", "Campaign", "Campaigns", "Campaign name")
    link("Content Calendar", "Owner", "Team", "Name")
    link("Content Calendar", "Written for", "Audience Personas",
         "Persona name")
    link("Caption Bank", "Content pillar", "Content Pillars",
         "Pillar name")
    link("Hashtag Bank", "Content pillar", "Content Pillars",
         "Pillar name")
    link("Analytics", "Account", "Accounts", "Account / profile name")
    link("Tasks", "Owner", "Team", "Name")
    link("Ads", "Linked campaign", "Campaigns", "Campaign name")
    link("Influencers", "Linked campaign", "Campaigns", "Campaign name")
    link("Repurposing", "Owner", "Team", "Name")
    link("Assets", "Campaign", "Campaigns", "Campaign name")

    # the dashboards need published posts with numbers against them
    statuses = Counter(column_values(cal, "Status"))
    if statuses.get("Published", 0) < 5:
        problems.append(f"only {statuses.get('Published', 0)} published "
                        f"posts - the dashboard will look empty")
    else:
        notes.append("status mix: " + ", ".join(
            f"{k} {v}" for k, v in statuses.most_common(5)))

    reach = [v for v in column_values(perf, "Reach")
             if isinstance(v, (int, float))]
    if len(reach) < 10:
        problems.append(f"only {len(reach)} rows of reach data")
    else:
        notes.append(f"{len(reach)} posts with reach, "
                     f"{sum(reach):,} in total")

    analytics = sheet("Analytics")
    if analytics is not None:
        months = column_values(analytics, "Month")
        if len(months) < 6:
            problems.append(f"only {len(months)} monthly analytics rows")
        else:
            notes.append(f"{len(months)} monthly analytics rows")
        for value in months:
            if hasattr(value, "day") and value.day != 1:
                problems.append(f"Analytics month {value} is not the 1st")
                break

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
    targets = argv[1:] or ["output/Social Media Content Planner.xlsx"]
    return max(check(path) for path in targets)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
````

### `tools/recalc_check.py`

Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error (#REF!, #DIV/0!, #VALUE! ...). *(216 lines)*

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

Write every source file of a generator into one markdown document. *(351 lines)*

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

PRODUCTS = {"pet": PET, "social": SOCIAL}


def blurb(path: str) -> str:
    """First sentence of the module docstring, as a one line description."""
    with open(os.path.join(ROOT, path), encoding="utf-8") as handle:
        source = handle.read()
    try:
        text = ast.get_docstring(ast.parse(source)) or ""
    except SyntaxError:                                  # pragma: no cover
        text = ""
    text = " ".join(text.split())
    if not text:
        return ""
    cut = re.search(r"\.(\s|$)", text)
    if cut:
        text = text[:cut.start()]
    return text.strip().rstrip(".")


def anchor(heading: str) -> str:
    slug = heading.lower()
    slug = re.sub(r"[^\w\s-]", "", slug)
    return re.sub(r"\s+", "-", slug.strip())


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--product", choices=sorted(PRODUCTS),
                        default="pet",
                        help="which generator to export (default: pet)")
    parser.add_argument("-o", "--output", default=None)
    args = parser.parse_args(argv)

    product = PRODUCTS[args.product]
    sections = product["sections"]

    lines = [f"# {product['title']} — source code", "",
             product["subtitle"], ""]
    lines += product["intro"]
    lines += ["", "---", "", "## Contents", ""]

    for heading, _intro, paths in sections:
        lines.append(f"* [{heading}](#{anchor(heading)})")
        for path in paths:
            lines.append(f"  * [`{path}`](#{anchor(path)})")
    lines += ["", "---", ""]

    total = 0
    for heading, intro, paths in sections:
        lines += [f"## {heading}", "", intro, ""]
        for path in paths:
            full = os.path.join(ROOT, path)
            with open(full, encoding="utf-8") as handle:
                code = handle.read().rstrip("\n")
            count = len(code.splitlines())
            total += count
            note = blurb(path)
            lines.append(f"### `{path}`")
            lines.append("")
            if note:
                lines.append(f"{note}. *({count} lines)*")
            else:
                lines.append(f"*({count} lines)*")
            # four backticks, so a fence inside a docstring cannot
            # close the block early
            lines += ["", "````python", code, "````", ""]
        lines.append("---")
        lines.append("")

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

Pack a generator into a single zip people can download and run. *(182 lines)*

````python
"""Pack a generator into a single zip people can download and run.

    python tools/make_download_zip.py
    python tools/make_download_zip.py --product social

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

*10,260 lines of Python in 56 files.*
