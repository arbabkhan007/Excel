# Secret Santa & White Elephant Party Tracker - Complete Source Code

**Product:** Secret Santa & White Elephant Party Tracker (Excel workbook generator)
**Brand:** Novality Store  |  **Sheet-protection password:** `premium`
**Version:** 1.0.0  |  **Built:** 18 September 2026  |  **Repo:** `arbabkhan007/Excel`, branch `arena/01a09508-excel`, tag `santa-v1.0.0`
**Generator:** Python 3 + vendored XlsxWriter (no other runtime dependencies)

This document is the *entire* product source: every module that builds the six
shipped workbooks (Premium/Basic x Noel/Arctic themes, blank + EXAMPLE modes),
the command-line entry point, and the QA tools used to prove every build.
Regenerate the whole product set with:

```bash
python3 secret_santa_party_tracker.py --all --outdir products
```

## What the product does

A complete party-management template, not just a name-draw sheet:

* **Secret Santa draw** - a closed-cycle derangement driven by a seed cell:
  nobody can ever draw themselves; change the seed to shuffle; manual
  overrides per pair; lock the draw with a status dropdown; sheet protection
  (password `premium`) freezes it.
* **Exclusion engine** - couple/household, same-team and last-year repeat
  toggles plus a custom no-pair table; the draw flags every breach live
  (Self / Couple / Team / LastYear / Custom).
* **Budget rules** - global min/max, per-gift actual spend, Under / Within /
  Over flags, receipt tracking, party totals and average gift cost.
* **Wishlists** - per-guest ideas with priorities and anonymous claiming.
* **White Elephant** - seed-shuffled seat order and gift numbers, a turn log
  (Pick / Steal / Pass) that drives steal counts, current holders, last
  stolen-from, automatic locks at max steals, final gifts and a highlighted
  "whose turn" row.
* **Dashboard** - participants, assignments, purchases, budgets, rules,
  game state, alerts, readiness bars and live charts.
* **Printable Santa Cards** - cut-out cards so each guest learns their
  recipient privately, plus an in-file private lookup (one name at a time).

## Editions & files shipped

| File | Sheets | Formulas | Dropdowns | Cond. formats | Charts |
|---|---|---|---|---|---|
| PREMIUM Noel / Arctic (blank) | 12 | 753 | 21 | 28 | 4 |
| PREMIUM Noel EXAMPLE | 12 | 753 | 21 | 28 | 4 |
| BASIC Noel / Arctic (blank) | 7 | 407 | 9 | 16 | 3 |
| BASIC Noel EXAMPLE | 7 | 407 | 9 | 16 | 3 |

Premium adds Exclusions & Rules, Wishlists, White Elephant, Game History and
Santa Cards; Basic keeps the draw + budget core.

## Design notes worth knowing

* The draw is `INDEX(names, MOD(n-1+MOD(seed-1,N-1),N)+1)` - a single N-cycle,
  so self-draws are mathematically impossible and every person gives and
  receives exactly once.
* Excel-2016-safe everywhere: no LET/LAMBDA/XLOOKUP/dynamic arrays; k-th and
  last-action lookups use SUMPRODUCT row-index extraction instead of
  CSE formulas.
* White Elephant state is derived purely from the Game History log, so the
  game can be logged from a phone in Google Sheets and still stay correct.
* Formula cells are locked; sheet protection uses the password `premium`
  (documented in the workbook's Start Here and Settings tabs).

## QA performed on every shipped file

* `tools/verify_workbook.py` - structure: hidden _Data, no writes under
  merges, names/DV/CF/chart counts -> 0 problems on all six.
* `tools/calc_check.py` - full recalculation with the `formulas` engine:
  0 error cells and 0 cached-vs-computed mismatches on all six (the single
  reported #REF! is the engine's known OFFSET limitation; Excel and Google
  Sheets resolve those dropdown names normally).
* `tools/layout_check.py` - column-width / row-height clipping scan -> clean.
* `tools/render_preview.py` - pixel renders of Dashboard, Draw and Start Here
  reviewed during development.

---


## 1. Package init

`santa_tracker/__init__.py`

```python
```

## 2. Config - geometry, column maps, caps, lists

`santa_tracker/config.py`

```python
"""Geometry, column maps, caps and list definitions for the
Secret Santa + White Elephant Party Tracker (Novality Store)."""

PRODUCT = "Secret Santa & White Elephant Party Tracker"
PRODUCT_SHORT = "Party Tracker"
TAGLINE = "draws, budgets, wishlists and white-elephant chaos - handled"
AUTHOR = "Novality Store"
VERSION = "1.0.0"

# protection password for locked (formula) cells
PROTECT_PASSWORD = "premium"

# ---------------------------------------------------------------------------
# sheet keys, tab names, editions
# ---------------------------------------------------------------------------
SHEET_NAMES = {
    "data": "_Data",
    "setup": "\u2699\ufe0f Settings & Instructions",
    "dashboard": "\U0001F3E0 Dashboard",
    "participants": "\U0001F465 Participants",
    "draw": "\U0001F385 Secret Santa Draw",
    "rules": "\U0001F6AB Exclusions & Rules",
    "budget": "\U0001F4B0 Budget Tracker",
    "wishlists": "\U0001F381 Wishlists",
    "we": "\U0001F3B2 White Elephant",
    "history": "\U0001F504 Game History",
    "cards": "\U0001F39F\ufe0f Santa Cards",
    "guide": "\U0001F4D6 Start Here",
}
SHEET_SHORT = {
    "data": "Data", "setup": "Settings", "dashboard": "Dashboard",
    "participants": "People", "draw": "Draw", "rules": "Rules",
    "budget": "Budget", "wishlists": "Wishlists", "we": "White Elephant",
    "history": "History", "cards": "Cards", "guide": "Start Here",
}

EDITIONS = {
    "premium": ["data", "setup", "dashboard", "participants", "draw",
                "rules", "budget", "wishlists", "we", "history", "cards",
                "guide"],
    "basic": ["data", "setup", "dashboard", "participants", "draw",
              "budget", "guide"],
}

# ---------------------------------------------------------------------------
# band geometry (rows 1-7) shared by every table sheet
# ---------------------------------------------------------------------------
ROW_SPACER_1 = 1
ROW_TITLE = 2
ROW_SUBTITLE = 3
ROW_SPACER_2 = 4
ROW_STATS = 5
ROW_SPACER_3 = 6
ROW_HEADER = 7
ROW_FIRST = 8


def last_row(key):
    return ROW_FIRST + CAP[key] - 1


def cols_last(key):
    return COLS_LAST[key]


CAP = {
    "participants": 24,
    "draw": 24,
    "rules": 16,
    "budget": 24,
    "wishlists": 40,
    "we": 24,
    "history": 60,
}

COLS = {
    "participants": {"n": "A", "name": "B", "team": "C", "household": "D",
                     "rsvp": "E", "diet": "F", "lastyear": "G",
                     "status": "H", "wishes": "I", "notes": "J"},
    "draw": {"n": "A", "giver": "B", "computed": "C", "override": "D",
             "final": "E", "flag": "F", "status": "G"},
    "rules": {"n": "A", "giver": "B", "cannot": "C", "reason": "D"},
    "budget": {"n": "A", "name": "B", "min": "C", "max": "D", "spent": "E",
               "flag": "F", "receipt": "G", "ref": "H", "notes": "I"},
    "wishlists": {"n": "A", "who": "B", "item": "C", "priority": "D",
                  "link": "E", "claimed": "F", "notes": "G"},
    "we": {"n": "A", "order": "B", "player": "C", "giftnum": "D",
           "desc": "E", "value": "F", "steals": "G", "maxsteals": "H",
           "status": "I", "holder": "J", "stolenfrom": "K", "final": "L",
           "key1": "M", "key2": "N"},
    "history": {"n": "A", "turn": "B", "player": "C", "action": "D",
                "gift": "E", "stolenfrom": "F"},
}
COLS_LAST = {
    "participants": "J", "draw": "G", "rules": "D", "budget": "I",
    "wishlists": "G", "we": "L", "history": "F",
}

WIDTHS = {
    "participants": {"A": 2.2, "B": 20, "C": 12, "D": 12, "E": 9, "F": 16,
                     "G": 16, "H": 15, "I": 8, "J": 34},
    "draw": {"A": 2.2, "B": 20, "C": 20, "D": 20, "E": 20, "F": 22,
             "G": 15},
    "rules": {"A": 2.2, "B": 20, "C": 20, "D": 32},
    "budget": {"A": 2.2, "B": 20, "C": 9, "D": 9, "E": 10, "F": 17,
               "G": 9, "H": 14, "I": 22},
    "wishlists": {"A": 2.2, "B": 20, "C": 30, "D": 13, "E": 22, "F": 9,
                  "G": 20},
    "we": {"A": 2.2, "B": 7, "C": 20, "D": 7, "E": 26, "F": 10, "G": 8,
           "H": 8, "I": 13, "J": 18, "K": 16, "L": 9, "M": 10, "N": 10},
    "history": {"A": 2.2, "B": 7, "C": 20, "D": 10, "E": 9, "F": 18},
    "dashboard": {"A": 2.2, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13,
                  "G": 13, "H": 13, "I": 13, "J": 13, "K": 13, "L": 13,
                  "M": 13, "N": 13},
    "setup": {"A": 2.2, "B": 38, "C": 24, "D": 16, "E": 16, "F": 16,
              "G": 16},
    "cards": {"A": 2.2, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13,
              "G": 13, "H": 13, "I": 13},
    "guide": {"A": 2.2, "B": 4, "C": 15, "D": 15, "E": 24, "F": 16,
              "G": 16, "H": 16, "I": 16, "J": 16},
}

# ---------------------------------------------------------------------------
# settings rows
# ---------------------------------------------------------------------------
SU_BUSINESS = 6        # party name
SU_MESSAGE = 7         # dashboard message
SU_CURRENCY = 10
SU_YEAR = 11           # party year
SU_DATE = 12           # party date
SU_LOCATION = 13
SU_HOST = 14
SU_BMIN = 15
SU_BMAX = 16
SU_STEALS = 17
SU_START = 18
SU_SEED = 19           # secret santa draw seed
SU_WESEED = 20         # white elephant seed
SU_STATUS = 21         # Draft / Final lock flag
SU_LIST_HEADER = 25
SU_LIST_FIRST = 26
SU_LIST_ROWS = 16
SU_FIXED_FIRST = 46

LIST_COLS = {"teams": "C", "diets": "D"}

FIXED_BLOCKS = {
    "Statuses": ("C", 4),     # Not Started / Gift Purchased / Wrapped / Complete
    "RSVP": ("D", 3),         # Yes / No / Maybe
    "Actions": ("E", 2),      # Pick / Steal
    "Priorities": ("F", 3),   # Must-love / Nice to have / Just an idea
    "Tick": ("G", 2),         # ✓ / blank
    "OnOff": ("H", 2),        # On / Off
}

TICK = "\u2713"

# statuses
ST_NOT = "Not Started"
ST_BOUGHT = "Gift Purchased"
ST_WRAPPED = "Wrapped"
ST_DONE = "Complete"

# gift statuses (white elephant)
GS_AVAIL = "\U0001F381 Available"
GS_HELD = "\U0001F932 Held"
GS_STOLEN = "\U0001F504 Stolen"
GS_FINAL = "\U0001F512 Final"

# budget flags
BF_UNDER = "\U0001F534 Under budget"
BF_OK = "\u2705 Within budget"
BF_OVER = "\U0001F6A8 Over budget"

# draw lock
DRAW_DRAFT = "Draft \u2014 re-draw allowed"
DRAW_FINAL = "Final \u2014 locked \U0001F512"

# ---------------------------------------------------------------------------
# _Data layout
# ---------------------------------------------------------------------------
DATA_KPI_FIRST = 2          # KPI label col AH, value col AI
KPI_ROW = {
    "participants": 2,
    "rsvp_yes": 3,
    "rsvp_pending": 4,
    "assigned": 5,
    "purchased": 6,
    "wrapped": 7,
    "complete": 8,
    "total_budget": 9,
    "spent": 10,
    "avg_gift": 11,
    "rules_count": 12,
    "we_players": 13,
    "gifts_in_play": 14,
    "locked_gifts": 15,
    "completion": 16,
    "turns_done": 17,
    "current_turn": 18,
    "violations": 19,
    "over_budget": 20,
    "diets": 21,
}
KPI_FMT = {
    "participants": "#,##0", "rsvp_yes": "#,##0", "rsvp_pending": "#,##0",
    "assigned": "#,##0", "purchased": "#,##0", "wrapped": "#,##0",
    "complete": "#,##0", "total_budget": "#,##0.00", "spent": "#,##0.00",
    "avg_gift": "#,##0.00", "rules_count": "#,##0", "we_players": "#,##0",
    "gifts_in_play": "#,##0", "locked_gifts": "#,##0", "completion": "0%",
    "turns_done": "#,##0", "current_turn": "#,##0", "violations": "#,##0",
    "over_budget": "#,##0", "diets": "#,##0",
}

# dashboard rows
DASH_HERO = 2
DASH_MSG = 5
DASH_CARD1_L = 8
DASH_CARD1_V = 9
DASH_CARD2_L = 10
DASH_CARD2_V = 11
DASH_BAR1 = 12
DASH_BAR2 = 13
DASH_CHART1 = 16
DASH_CHART2 = 32
DASH_CHART3 = 48
DASH_PANEL = 64
DASH_FOOTER = 72
DASH_NAV = 74

# white elephant / draw control rows
WE_CTRL_ROW = 5
HIST_NOTE_ROW = 5
```

## 3. Themes - Noel Classic & Arctic Frost

`santa_tracker/theme.py`

```python
"""
Colour palettes / design systems for the Secret Santa & White Elephant
Party Tracker (Novality Store).

Two themes ship with the product:

  noel    warm cream canvas, deep holly red, pine green, antique gold
  arctic  cool night canvas paper-white, deep navy, ice blue, berry accent
"""


class Theme(object):
    """Small attribute bag so sheet code can read ``th.primary`` etc."""

    def __init__(self, **kw):
        self.__dict__.update(kw)

    def soft(self, key):
        """Return the pale 'tint' partner of a colour key."""
        return self.__dict__[key + "_soft"]


# ===========================================================================
# NOEL - "classic Christmas parlour"
# ===========================================================================
NOEL = Theme(
    key="noel",
    label="Noel",
    bg="#FBF6EE",
    cover_bg="#FBF6EE",
    card="#FFFFFF",
    alt="#FCF8F1",
    border="#E7D9C6",
    border_strong="#C9AE8C",
    primary="#8C1D2C",          # holly red
    primary_2="#1F5C40",        # pine
    primary_2_soft="#E1EDE5",
    primary_soft="#F5E1E2",
    accent="#1F5C40",           # pine green
    accent_soft="#E1EDE5",
    gold="#B98A2E",             # antique gold
    gold_soft="#F6EDD8",
    ink="#33221E",
    muted="#9A8474",
    white="#FFFFFF",
    ok="#1F5C40",
    ok_soft="#E1EDE5",
    warn="#B4761A",
    warn_soft="#FBF0D9",
    bad="#AC2F2F",
    bad_soft="#F7E2E0",
    info="#5E6FA3",
    info_soft="#E5E9F4",
    plum="#8C1D2C",
    plum_soft="#F5E1E2",
    title_font="Georgia",
    body_font="Calibri",
    mono_font="Consolas",
    tabs={
        "data": "#9A8474",
        "setup": "#1F5C40",
        "dashboard": "#8C1D2C",
        "participants": "#1F5C40",
        "draw": "#8C1D2C",
        "rules": "#B4761A",
        "budget": "#B98A2E",
        "wishlists": "#5E6FA3",
        "we": "#1F5C40",
        "history": "#5E6FA3",
        "cards": "#8C1D2C",
        "guide": "#9A8474",
    },
)

# ===========================================================================
# ARCTIC - "midnight frost party"
# ===========================================================================
ARCTIC = Theme(
    key="arctic",
    label="Arctic",
    bg="#F4F7FB",
    cover_bg="#F4F7FB",
    card="#FFFFFF",
    alt="#F8FAFD",
    border="#D8E1EC",
    border_strong="#A9BBD0",
    primary="#1F3A5F",          # deep navy
    primary_2="#5B8DB8",        # ice blue
    primary_2_soft="#E4EDF5",
    primary_soft="#E2E9F2",
    accent="#C4485C",           # winter berry
    accent_soft="#F8E3E6",
    gold="#C9A227",             # starlight gold
    gold_soft="#F7F0DA",
    ink="#1E2833",
    muted="#7C8CA0",
    white="#FFFFFF",
    ok="#2F6D5F",
    ok_soft="#E0EEE9",
    warn="#B4761A",
    warn_soft="#FBF0D9",
    bad="#AC2F2F",
    bad_soft="#F7E2E0",
    info="#5B8DB8",
    info_soft="#E4EDF5",
    plum="#1F3A5F",
    plum_soft="#E2E9F2",
    title_font="Trebuchet MS",
    body_font="Trebuchet MS",
    mono_font="Consolas",
    tabs={
        "data": "#7C8CA0",
        "setup": "#5B8DB8",
        "dashboard": "#1F3A5F",
        "participants": "#2F6D5F",
        "draw": "#C4485C",
        "rules": "#B4761A",
        "budget": "#C9A227",
        "wishlists": "#5B8DB8",
        "we": "#2F6D5F",
        "history": "#5B8DB8",
        "cards": "#C4485C",
        "guide": "#7C8CA0",
    },
)

THEMES = {"noel": NOEL, "arctic": ARCTIC}
```

## 4. Styles - format factory

`santa_tracker/styles.py`

```python
"""
Every cell format used by the workbook, generated from the active theme.

Sheet builders never build formats inline - they ask ``Styles`` for a
semantic style (``S.cell("money", alt)``), which keeps the two themes
visually consistent and keeps the format cache small.
"""

import math


class Styles(object):

    def __init__(self, wb, th):
        self.wb = wb
        self.th = th
        self._cache = {}
        self._build()

    # ------------------------------------------------------------------
    # core
    # ------------------------------------------------------------------
    def f(self, props=None, **kw):
        """Cached ``workbook.add_format`` (accepts a dict or keywords)."""
        p = dict(props or {})
        p.update(kw)
        key = tuple(sorted((k, str(v)) for k, v in p.items()))
        fmt = self._cache.get(key)
        if fmt is None:
            fmt = self.wb.add_format(p)
            self._cache[key] = fmt
        return fmt

    def base(self, **over):
        """Body-text defaults + overrides."""
        p = {"font_name": self.th.body_font, "font_size": 10.5,
             "font_color": self.th.ink, "valign": "vcenter"}
        p.update(over)
        return p

    # ------------------------------------------------------------------
    # canvases, banners, sections
    # ------------------------------------------------------------------
    def _build(self):
        th = self.th
        self.canvas = self.f({"bg_color": th.bg})
        self.canvas_card = self.f({"bg_color": th.card})

        self.hero_title = self.f(self.base(
            font_name=th.title_font, font_size=26, bold=True,
            font_color=th.white, bg_color=th.primary, align="left",
            valign="vcenter", indent=1))
        self.hero_count = self.f(self.base(
            font_name=th.title_font, font_size=17, bold=True,
            font_color=th.gold_soft, bg_color=th.accent, align="left",
            valign="vcenter", indent=1))
        self.hero_meta = self.f(self.base(
            font_size=10, italic=True, font_color=th.ink,
            bg_color=th.gold_soft, align="left", valign="vcenter", indent=1))

        self.sheet_title = self.f(self.base(
            font_name=th.title_font, font_size=20, bold=True,
            font_color=th.primary, bg_color=th.bg, align="left",
            valign="vcenter", indent=1))
        self.sheet_sub = self.f(self.base(
            font_size=10.5, italic=True, font_color=th.muted,
            bg_color=th.bg, align="left", valign="vcenter", indent=1))
        self.home_link = self.f(self.base(
            font_size=10, bold=True, font_color=th.white, bg_color=th.accent,
            align="center", valign="vcenter", border=1,
            border_color=th.accent))

        self.section = self.f(self.base(
            font_name=th.title_font, font_size=13, bold=True,
            font_color=th.white, bg_color=th.primary, align="left",
            valign="vcenter", indent=1))
        self.section_accent = self.f(self.base(
            font_name=th.title_font, font_size=13, bold=True,
            font_color=th.white, bg_color=th.accent, align="left",
            valign="vcenter", indent=1))
        self.section_gold = self.f(self.base(
            font_name=th.title_font, font_size=13, bold=True,
            font_color=th.ink, bg_color=th.gold_soft, align="left",
            valign="vcenter", indent=1, bottom=2, bottom_color=th.gold))
        self.section_soft = self.f(self.base(
            font_name=th.title_font, font_size=12, bold=True,
            font_color=th.primary, bg_color=th.primary_soft, align="left",
            valign="vcenter", indent=1, left=4, left_color=th.gold))

        # table header
        self.thead = self.f(self.base(
            font_size=10, bold=True, font_color=th.white, bg_color=th.primary,
            align="center", valign="vcenter", text_wrap=True, border=1,
            border_color=th.primary))

        self.note = self.f(self.base(
            font_size=10, italic=True, font_color=th.muted, bg_color=th.alt,
            align="left", valign="top", text_wrap=True, border=1,
            border_color=th.border, indent=1, locked=False))
        self.note_plain = self.f(self.base(
            font_size=10, italic=True, font_color=th.muted, bg_color=th.bg,
            align="left", valign="top", text_wrap=True))
        self.footer = self.f(self.base(
            font_size=9, italic=True, font_color=th.muted, bg_color=th.bg,
            align="left", valign="vcenter"))

        self.bar_text = self.f(self.base(
            font_name=th.mono_font, font_size=13, bold=True,
            font_color=th.primary_2, bg_color=th.card, align="left",
            valign="vcenter", border=1, border_color=th.border))
        self.bar_label = self.f(self.base(
            font_size=10, bold=True, font_color=th.ink, bg_color=th.card,
            align="left", valign="vcenter", border=1,
            border_color=th.border, indent=1))
        self.bar_pct = self.f(self.base(
            font_name=th.mono_font, font_size=13, bold=True,
            font_color=th.accent, bg_color=th.card, align="center",
            valign="vcenter", border=1, border_color=th.border))

    # ------------------------------------------------------------------
    # table headers in an accent colour
    # ------------------------------------------------------------------
    def header(self, color=None):
        color = color or self.th.primary
        return self.f(self.base(
            font_size=10, bold=True, font_color=self.th.white,
            bg_color=color, align="center", valign="vcenter", text_wrap=True,
            border=1, border_color=color))

    # ------------------------------------------------------------------
    # data cells
    # ------------------------------------------------------------------
    _KINDS = {
        "text":       {"align": "left", "locked": False},
        "center":     {"align": "center", "locked": False},
        "money":      {"align": "right", "num_format": "#,##0.00",
                       "locked": False},
        "money0":     {"align": "right", "num_format": "#,##0",
                       "locked": False},
        "num":        {"align": "right", "num_format": "#,##0",
                       "locked": False},
        "qty":        {"align": "center", "num_format": "#,##0",
                       "locked": False},
        "qty1":       {"align": "center", "num_format": "#,##0.0",
                       "locked": False},
        "pct":        {"align": "center", "num_format": "0%",
                       "locked": False},
        "date":       {"align": "center", "num_format": "dd mmm yyyy",
                       "locked": False},
        "tick":       {"align": "center", "font_size": 13, "bold": True,
                       "font_color": "#2F7D4F", "locked": False},
        "wrap":       {"align": "left", "text_wrap": True, "valign": "top",
                       "locked": False},
        "link":       {"align": "left", "font_color": "#2A5D8F",
                       "underline": 1, "locked": False},
        "url":        {"align": "center", "bold": True, "font_size": 10},
        # calculated columns (locked - they must not be typed over)
        "calc":       {"align": "left"},
        "calc_c":     {"align": "center"},
        "calc_wrap":  {"align": "left", "text_wrap": True},
        "calc_money": {"align": "right", "num_format": "#,##0.00"},
        "calc_num":   {"align": "right", "num_format": "#,##0"},
        "calc_qty1":  {"align": "center", "num_format": "#,##0.0"},
        "calc_pct":   {"align": "center", "num_format": "0%"},
        "calc_date":  {"align": "center", "num_format": "dd mmm yyyy"},
        "calc_tick":  {"align": "center", "font_size": 13, "bold": True},
    }

    def cell(self, kind, alt=False, **over):
        """Data-cell format.  ``alt`` selects the zebra-stripe background."""
        th = self.th
        props = self.base(**self._KINDS[kind])
        if kind.startswith("calc"):
            props["bg_color"] = th.alt
            props["font_color"] = th.ink if kind in ("calc", "calc_wrap",
                                                     "calc_c") else th.ink
            props["locked"] = True
            if kind in ("calc_money", "calc_num", "calc_pct", "calc_date"):
                props["font_color"] = th.primary
        else:
            props["bg_color"] = th.alt if alt else th.card
        props["border"] = 1
        props["border_color"] = th.border
        props.update(over)
        return self.f(**props)

    def idx(self, alt=False):
        """The little grey row-number column."""
        return self.f(self.base(
            font_size=9, font_color=self.th.muted, align="center",
            valign="vcenter", border=1, border_color=self.th.border,
            bg_color=self.th.alt if alt else self.th.card,
            num_format="0"))

    # ------------------------------------------------------------------
    # dashboard furniture
    # ------------------------------------------------------------------
    def kpi_label(self, color, size=9.5, align="left"):
        return self.f(self.base(
            font_size=size, bold=True, font_color=self.th.white,
            bg_color=color, align=align, valign="vcenter", indent=1,
            border=1, border_color=color))

    def kpi_value(self, color, num_format=None, size=21, align="center",
                  bg=None):
        p = self.base(font_name=self.th.title_font, font_size=size, bold=True,
                      font_color=color, bg_color=bg or self.th.card,
                      align=align, valign="vcenter", border=1,
                      border_color=self.th.border)
        if num_format:
            p["num_format"] = num_format
        return self.f(**p)

    def kpi_text(self, color, size=11, align="left", bg=None, bold=True,
                 wrap=False):
        return self.f(self.base(
            font_size=size, bold=bold, font_color=color,
            bg_color=bg or self.th.card, align=align, valign="vcenter",
            border=1, border_color=self.th.border, indent=1 if align == "left" else 0,
            text_wrap=wrap))

    def panel(self, color=None, bg=None, size=10.5, bold=False, align="left",
              wrap=True, font_color=None, valign="top", border=True):
        p = self.base(font_size=size, bold=bold, align=align, valign=valign,
                      text_wrap=wrap,
                      bg_color=bg if bg is not None else self.th.card,
                      font_color=font_color or self.th.ink)
        if border:
            p["border"] = 1
            p["border_color"] = color or self.th.border
            if color:
                p["left"] = 3
                p["left_color"] = color
        return self.f(**p)

    def pill(self, bg, fg, size=9.5, bold=True, align="center"):
        return self.f(self.base(font_size=size, bold=bold, font_color=fg,
                                bg_color=bg, align=align, valign="vcenter",
                                border=1, border_color=bg, text_wrap=False))

    def nav(self, color, size=10.5):
        return self.f(self.base(font_size=size, bold=True,
                                font_color=self.th.white, bg_color=color,
                                align="center", valign="vcenter", border=1,
                                border_color=color))

    def nav_soft(self, size=10.5):
        th = self.th
        return self.f(self.base(font_size=size, bold=True,
                                font_color=th.primary, bg_color=th.primary_soft,
                                align="center", valign="vcenter", border=1,
                                border_color=th.border))

    # ------------------------------------------------------------------
    # conditional-format formats (dxf sources)
    # ------------------------------------------------------------------
    def cf(self, bg=None, fg=None, bold=False, strike=False, italic=False,
           border=None, size=None, align=None, num_format=None):
        th = self.th
        p = {}
        if bg:
            p["bg_color"] = bg
        if fg:
            p["font_color"] = fg
        if bold:
            p["bold"] = True
        if strike:
            p["font_strikeout"] = True
        if italic:
            p["italic"] = True
        if size:
            p["font_size"] = size
        if align:
            p["align"] = align
        if num_format:
            p["num_format"] = num_format
        if border:
            p["border"] = 1
            p["border_color"] = border
        p["font_name"] = th.body_font
        return self.f(**p)

    def cf_state(self, state):
        """Semantic state colours: ok / warn / bad / info / plum / gold."""
        th = self.th
        return self.cf(bg=th.soft(state), fg=getattr(th, state), bold=True,
                       border=th.border)

    # ------------------------------------------------------------------
    # helpers
    # ------------------------------------------------------------------
    def guide_text(self, size=10.5, bold=False, italic=False, color=None,
                   bg=None, indent=0, align="left"):
        return self.f(self.base(font_size=size, bold=bold, italic=italic,
                                font_color=color or self.th.ink,
                                bg_color=bg or self.th.card, align=align,
                                valign="top", text_wrap=True, indent=indent,
                                border=1, border_color=self.th.border,
                                locked=False))

    def static_height(self, text, width_chars, line_px=14.6, pad=8,
                      minimum=18):
        """Row height needed to show wrapped ``text`` in ``width_chars``."""
        if not text:
            return minimum
        per_line = max(10, int(width_chars * 0.95))
        lines = 0
        for para in str(text).split("\n"):
            lines += max(1, int(math.ceil(len(para) / float(per_line))))
        return max(minimum, lines * line_px + pad)


def wrap_height(text, width_chars, line_px=14.6, pad=8, minimum=18):
    """Module-level twin of :meth:`Styles.static_height`."""
    if not text:
        return minimum
    per_line = max(10, int(width_chars * 0.95))
    lines = 0
    for para in str(text).split("\n"):
        lines += max(1, int(math.ceil(len(para) / float(per_line))))
    return max(minimum, lines * line_px + pad)
```

## 5. Book - workbook engine (names, DV, CF, charts, protection)

`santa_tracker/book.py`

```python
"""
The build context.

``Book`` owns the workbook, the worksheet registry, the defined names and all
of the little helpers that turn a cell address into a formula fragment.  Sheet
builders receive one ``Book`` and never talk to XlsxWriter's raw indices
directly, which is what keeps the cross-sheet formulas correct when a tab is
missing (Basic edition) or when a capacity changes.
"""

import datetime
import types

import xlsxwriter
from xlsxwriter.utility import datetime_to_excel_datetime

from . import config as C
from .styles import Styles, wrap_height

_XLWorksheet = xlsxwriter.worksheet.Worksheet


def _coerce(value):
    """Cached formula results must be a number, a bool or a string.

    XlsxWriter writes whatever it is given straight into ``<v>``; a
    ``datetime.date`` would produce invalid XML that Excel refuses to open.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, (datetime.datetime, datetime.date)):
        return datetime_to_excel_datetime(value, False, False)
    if isinstance(value, (str, int, float)):
        return value
    return str(value)


def _safe_write_formula(self, row, col, formula, cell_format=None, value=0):
    """``write_formula`` that also coerces the cached result."""
    return _XLWorksheet.write_formula(self, row, col, formula, cell_format,
                                      _coerce(value))


def r(n):
    """1-indexed spreadsheet row -> 0-indexed XlsxWriter row."""
    return n - 1


def cellref(col, row):
    return "%s%d" % (col, row)


class Book(object):

    def __init__(self, path, theme, edition="premium", mode="blank",
                 protect=None, images=True, demo=None):
        self.path = path
        self.th = theme
        self.edition = edition
        self.mode = mode            # "blank" | "demo"
        self.protect = protect
        self.images = images
        self.demo = demo

        self.order = C.EDITIONS[edition]
        self.wb = xlsxwriter.Workbook(path)
        self.S = Styles(self.wb, theme)
        self.sheets = {}
        self.stats = {"formulas": 0, "validations": 0, "cond_formats": 0,
                      "charts": 0, "links": 0, "cells": 0}

        self._create_sheets()
        self._define_names()

        self.wb.set_properties({
            "title": C.PRODUCT,
            "subject": C.TAGLINE,
            "author": C.AUTHOR,
            "manager": C.AUTHOR,
            "company": C.AUTHOR,
            "category": "Crochet Craft Fair Business Spreadsheet",
            "keywords": ("crochet business, craft fair tracker, product "
                         "catalog, yarn inventory, inventory spreadsheet, "
                         "profit dashboard, pricing calculator, sales log, "
                         "event profit, reorder list, excel template, "
                         "google sheets, etsy spreadsheet"),
            "comments": ("%s v%s - %s. Works in Excel 2016+ and Google "
                         "Sheets. No macros, nothing to install."
                         % (C.PRODUCT, C.VERSION, C.TAGLINE)),
        })
        self.wb.set_calc_mode("auto")

    # ------------------------------------------------------------------
    # sheet registry
    # ------------------------------------------------------------------
    def _create_sheets(self):
        for key in self.order:
            ws = self.wb.add_worksheet(C.SHEET_NAMES[key])
            ws.write_formula = types.MethodType(_safe_write_formula, ws)
            ws.set_tab_color(self.th.tabs[key])
            ws.hide_gridlines(2)
            self.sheets[key] = ws
        if "data" in self.sheets:
            self.sheets["data"].hide()
        self.sheets["dashboard"].set_first_sheet()
        self.sheets["dashboard"].activate()
        self.sheets["dashboard"].set_zoom(90)

    def ws(self, key):
        return self.sheets[key]

    def has(self, key):
        return key in self.sheets and key != "data"

    def name(self, key):
        return C.SHEET_NAMES[key]

    def q(self, key):
        """Quoted sheet name for use inside formulas."""
        return "'%s'" % C.SHEET_NAMES[key]

    # ------------------------------------------------------------------
    # defined names
    # ------------------------------------------------------------------
    def _define_names(self):
        su = self.q("setup")
        simple = {
            "PartyName": "%s!$C$%d" % (su, C.SU_BUSINESS),
            "Message": "%s!$C$%d" % (su, C.SU_MESSAGE),
            "Currency": "%s!$C$%d" % (su, C.SU_CURRENCY),
            "PartyYear": "%s!$C$%d" % (su, C.SU_YEAR),
            "PartyDate": "%s!$C$%d" % (su, C.SU_DATE),
            "PartyLocation": "%s!$C$%d" % (su, C.SU_LOCATION),
            "PartyHost": "%s!$C$%d" % (su, C.SU_HOST),
            "BudgetMin": "%s!$C$%d" % (su, C.SU_BMIN),
            "BudgetMax": "%s!$C$%d" % (su, C.SU_BMAX),
            "MaxSteals": "%s!$C$%d" % (su, C.SU_STEALS),
            "StartPlayer": "%s!$C$%d" % (su, C.SU_START),
            "DrawSeed": "%s!$C$%d" % (su, C.SU_SEED),
            "WESeed": "%s!$C$%d" % (su, C.SU_WESEED),
            "DrawStatus": "%s!$C$%d" % (su, C.SU_STATUS),
        }
        for nm, ref in simple.items():
            self.wb.define_name(nm, "=" + ref)

        first = C.SU_LIST_FIRST
        lastc = C.SU_LIST_FIRST + C.SU_LIST_ROWS - 1
        for key, col in sorted(C.LIST_COLS.items()):
            rng = "%s!$%s$%d:$%s$%d" % (su, col, first, col, lastc)
            self.wb.define_name(
                _name_for_list(key),
                "=OFFSET(%s!$%s$%d,0,0,MAX(1,COUNTA(%s)),1)" % (su, col, first, rng))

        # Fixed (non-customisable) status lists, in the locked block.
        for nm, (col, n) in sorted(C.FIXED_BLOCKS.items()):
            self.wb.define_name(
                nm, "=%s!$%s$%d:$%s$%d"
                % (su, col, C.SU_FIXED_FIRST, col, C.SU_FIXED_FIRST + n - 1))

        if self.has("rules"):
            ru = self.q("rules")
            for nm, row in (("CoupleRule", 8), ("TeamRule", 9),
                            ("LastYearRule", 10)):
                self.wb.define_name(nm, "=%s!$C$%d" % (ru, row))

        # Dynamic lists that live in the working sheets (grow with the data).
        dyn = {"ParticipantsList": ("participants", "B"),
               "GiftNums": ("we", "D")}
        for nm, (sheet, colL) in sorted(dyn.items()):
            if not self.has(sheet):
                continue
            rng = "%s!$%s$%d:$%s$%d" % (self.q(sheet), colL, C.ROW_FIRST,
                                        colL, C.last_row(sheet))
            base = rng.rsplit("$", 1)[0].rsplit(":", 1)[0]
            self.wb.define_name(
                nm, "=OFFSET(%s,0,0,MAX(1,COUNTA(%s)),1)" % (base, rng))

    # ------------------------------------------------------------------
    # reference helpers
    # ------------------------------------------------------------------
    def col(self, key, field):
        return C.COLS[key][field]

    def rng(self, key, field, first=None, last=None, quoted=True):
        """Absolute range for a column of a tracker sheet."""
        col = self.col(key, field)
        first = C.ROW_FIRST if first is None else first
        last = C.last_row(key) if last is None else last
        ref = "%s!$%s$%d:$%s$%d" % (self.q(key) if quoted else self.name(key),
                                    col, first, col, last)
        return ref

    def cell(self, key, field, row):
        return "%s!$%s$%d" % (self.q(key), self.col(key, field), row)

    def kpi(self, key):
        return "%s!$AI$%d" % (self.q("data"), C.KPI_ROW[key])

    def kpi_fmt(self, key):
        return C.KPI_FMT[key]

    def data_rng(self, col, first, last):
        return "%s!$%s$%d:$%s$%d" % (self.q("data"), col, first, col, last)

    def data_cell(self, col, row):
        return "%s!$%s$%d" % (self.q("data"), col, row)

    def listname(self, key):
        return _name_for_list(key)

    # ------------------------------------------------------------------
    # cached values (so previews look right before Excel recalculates)
    # ------------------------------------------------------------------
    def cached(self, key, default=0):
        if self.demo and key in self.demo.agg:
            return self.demo.agg[key]
        return default

    # ------------------------------------------------------------------
    # money / text formula helpers
    # ------------------------------------------------------------------
    def money(self, value_formula, decimals="#,##0"):
        """Currency-prefixed text, driven by the Currency setting."""
        return 'Currency&TEXT(%s,"%s")' % (value_formula, decimals)

    def bar(self, numerator, denominator, blocks=18):
        """Text progress bar built with REPT()."""
        pct = "MIN(1,IFERROR((%s)/(%s),0))" % (numerator, denominator)
        filled = "ROUND(%s*%d,0)" % (pct, blocks)
        return ('=IFERROR(REPT("\u2588",%s)&REPT("\u2591",%d-%s),"%s")'
                % (filled, blocks, filled, "\u2591" * blocks))

    def bar_static(self, numerator, denominator, blocks=18):
        """Python twin of :meth:`bar` used for cached values."""
        try:
            pct = min(1.0, float(numerator) / float(denominator)) if denominator else 0.0
        except (TypeError, ValueError, ZeroDivisionError):
            pct = 0.0
        filled = int(round(pct * blocks))
        return "\u2588" * filled + "\u2591" * (blocks - filled)

    # ------------------------------------------------------------------
    # page furniture
    # ------------------------------------------------------------------
    def page(self, key, last_col, last_row, landscape=True, freeze=None,
             fit=True, zoom=90, title_rows=None, paper=9):
        ws = self.sheets[key]
        ws.set_zoom(zoom)
        if landscape:
            ws.set_landscape()
        if fit:
            ws.fit_to_pages(1, 0)
        ws.set_margins(0.4, 0.4, 0.5, 0.5)
        ws.set_paper(paper)
        ws.set_header("&L&\"%s,Bold\"&11%s&R&\"%s,Italic\"&9Page &P of &N"
                      % (self.th.body_font, C.PRODUCT_SHORT,
                         self.th.body_font))
        ws.set_footer("&C&\"%s,Italic\"&8&A  \u2022  %s v%s  \u2022  "
                      "personal licence"
                      % (self.th.body_font, C.PRODUCT, C.VERSION))
        ws.print_area("A1:%s%d" % (last_col, last_row))
        if title_rows:
            ws.repeat_rows(title_rows[0], title_rows[1])
        if freeze:
            ws.freeze_panes(*freeze)
        if self.protect:
            ws.protect(self.protect, {
                "objects": True, "scenarios": True, "format_cells": True,
                "format_columns": True, "format_rows": True,
                "insert_rows": True, "insert_columns": True,
                "insert_hyperlinks": True, "delete_rows": True,
                "delete_columns": True, "sort": True, "autofilter": True,
                "select_locked_cells": True, "select_unlocked_cells": True,
            })
        return ws

    def widths(self, key, extra=None):
        ws = self.sheets[key]
        for col, width in sorted((C.WIDTHS.get(key, {}) or {}).items()):
            ws.set_column("%s:%s" % (col, col), width)
        for col, width in sorted((extra or {}).items()):
            ws.set_column("%s:%s" % (col, col), width)
        return ws

    def title_block(self, key, title, subtitle, last_col, home=True):
        """Standard 6-row header band used by every tracker tab."""
        ws = self.sheets[key]
        S = self.S
        ws.set_row(r(C.ROW_SPACER_1), 7)
        ws.set_row(r(C.ROW_TITLE), 30)
        ws.set_row(r(C.ROW_SUBTITLE), 18)
        ws.set_row(r(C.ROW_SPACER_2), 7)
        ws.set_row(r(C.ROW_STATS), 26)
        ws.set_row(r(C.ROW_SPACER_3), 7)
        ws.merge_range(r(C.ROW_TITLE), 1, r(C.ROW_TITLE), _ci(last_col) - 3,
                       title, S.sheet_title)
        for c in range(_ci(last_col) - 2, _ci(last_col) + 1):
            ws.write_blank(r(C.ROW_TITLE), c, None, S.canvas)
        ws.merge_range(r(C.ROW_SUBTITLE), 1, r(C.ROW_SUBTITLE), _ci(last_col) - 3,
                       subtitle, S.sheet_sub)
        if home:
            ws.merge_range(r(C.ROW_SUBTITLE), _ci(last_col) - 2,
                           r(C.ROW_SUBTITLE), _ci(last_col),
                           "", S.home_link)
            ws.write_url(r(C.ROW_SUBTITLE), _ci(last_col) - 2,
                         "internal:%s!A1" % self.q("dashboard"), S.home_link,
                         "\U0001F3E0  Back to Dashboard")
            self.stats["links"] += 1
        else:
            ws.merge_range(r(C.ROW_SUBTITLE), _ci(last_col) - 2,
                           r(C.ROW_SUBTITLE), _ci(last_col), "", S.canvas)
        return ws

    def stats_strip(self, key, items, row=None, first_col=1, span=2):
        """Row of small formula chips under the sheet title.

        ``items`` is a list of ``(formula_or_text, colour_key, cached)``.
        """
        ws = self.sheets[key]
        row = C.ROW_STATS if row is None else row
        col = first_col
        for formula, color, cached in items:
            fmt = self.S.pill(self.th.soft(color), getattr(self.th, color),
                              size=10.5, bold=True, align="left")
            ws.merge_range(r(row), col, r(row), col + span - 1, "", fmt)
            if formula.startswith("="):
                ws.write_formula(r(row), col, formula, fmt,
                                 cached if cached is not None else 0)
                self.stats["formulas"] += 1
            else:
                ws.write(r(row), col, formula, fmt)
            col += span
        return ws

    def nav_row(self, key, row, first_col=1, span=2, max_col="N"):
        """Bottom navigation buttons (one per visible tab)."""
        ws = self.sheets[key]
        palette = ["primary", "accent", "gold", "info", "plum", "ok",
                   "primary_2", "warn", "accent", "info"]
        col = first_col
        n = 0
        for target in self.order:
            if target == "data":
                continue
            if col + span - 1 > _ci(max_col):
                row += 1
                col = first_col
                ws.set_row(r(row), 24)
            color = getattr(self.th, palette[n % len(palette)])
            fmt = self.S.nav(color)
            if span > 1:
                ws.merge_range(r(row), col, r(row), col + span - 1, "", fmt)
            label = C.SHEET_SHORT[target]
            if target == key:
                label = "\u25B6 " + label
            ws.write_url(r(row), col, "internal:%s!A1" % self.q(target),
                         fmt, label)
            self.stats["links"] += 1
            col += span
            n += 1
        return row

    # ------------------------------------------------------------------
    # painting / writing helpers
    # ------------------------------------------------------------------
    def paint(self, key, row1, col1, row2, col2, fmt):
        ws = self.sheets[key]
        for rr in range(row1, row2 + 1):
            for cc in range(col1, col2 + 1):
                ws.write_blank(rr, cc, None, fmt)

    def band(self, key, row, col1, col2, fmt, height=None):
        ws = self.sheets[key]
        ws.merge_range(r(row), col1, r(row), col2, "", fmt)
        if height:
            ws.set_row(r(row), height)

    def write(self, key, row, col, value, fmt=None, cached=None):
        ws = self.sheets[key]
        self.stats["cells"] += 1
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(r(row), col, value, fmt,
                             cached if cached is not None else 0)
            self.stats["formulas"] += 1
        elif value is None or value == "":
            ws.write_blank(r(row), col, None, fmt)
        else:
            ws.write(r(row), col, value, fmt)
        return ws

    def merge(self, key, row1, col1, row2, col2, value, fmt=None,
              cached=None):
        ws = self.sheets[key]
        self.stats["cells"] += 1
        if isinstance(value, str) and value.startswith("="):
            ws.merge_range(r(row1), col1, r(row2), col2, "", fmt)
            ws.write_formula(r(row1), col1, value, fmt,
                             cached if cached is not None else 0)
            self.stats["formulas"] += 1
        else:
            ws.merge_range(r(row1), col1, r(row2), col2, value, fmt)
        return ws

    def para(self, key, row, col1, col2, text, fmt, width_chars=None,
             minimum=18):
        """Merged, wrapped paragraph with a row height that actually fits."""
        if width_chars is None:
            width_chars = sum(
                (C.WIDTHS.get(key, {}) or {}).get(_cl(c), 10)
                for c in range(col1, col2 + 1))
        h = wrap_height(text, width_chars, minimum=minimum)
        self.merge(key, row, col1, row, col2, text, fmt)
        self.sheets[key].set_row(r(row), h)
        return row + 1

    def validate(self, key, row1, col1, row2, col2, source, title=None,
                 message=None, error=None, error_type="stop"):
        ws = self.sheets[key]
        opts = {"validate": "list", "source": source, "ignore_blank": True,
                "show_input": bool(title or message),
                "show_error": bool(error)}
        if title:
            opts["input_title"] = title[:32]
        if message:
            opts["input_message"] = message[:255]
        if error:
            opts["error_title"] = "Pick from the list"[:32]
            opts["error_message"] = error[:255]
            opts["error_type"] = error_type
        res = ws.data_validation(r(row1), col1, r(row2), col2, opts)
        if res == 0:
            self.stats["validations"] += 1
        return res

    def cond(self, key, row1, col1, row2, col2, opts):
        ws = self.sheets[key]
        res = ws.conditional_format(r(row1), col1, r(row2), col2, opts)
        if res == 0:
            self.stats["cond_formats"] += 1
        return res

    def chart(self, ctype, **opts):
        ch = self.wb.add_chart(dict({"type": ctype}, **opts))
        ch.show_hidden_data()          # data lives on the hidden _Data sheet
        ch.show_blanks_as("gap")
        self.stats["charts"] += 1
        return ch

    def close(self):
        self.wb.close()


# ----------------------------------------------------------------------
# small utilities
# ----------------------------------------------------------------------
def _ci(letter):
    """Column letter -> 0-indexed column number."""
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def _cl(index):
    """0-indexed column number -> column letter."""
    s = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        s = chr(65 + rem) + s
    return s


_LIST_NAMES = {
    "teams": "Teams",
    "diets": "Diets",
    "statuses": "Statuses",
    "rsvp": "RSVP",
    "actions": "Actions",
    "priorities": "Priorities",
    "onoff": "OnOff",
    "participants": "ParticipantsList",
    "giftnums": "GiftNums",
}


def _name_for_list(key):
    return _LIST_NAMES[key]


def name_for_list(key):
    return _LIST_NAMES[key]


def ci(letter):
    return _ci(letter)


def cl(index):
    return _cl(index)
```

## 6. Demo model - Novality Store holiday party (with White Elephant simulation)

`santa_tracker/demo.py`

```python
"""Fictional demo party: the Novality Store holiday party, December 2026.

Feeds cached values into every formula cell of the EXAMPLE workbooks so the
file opens looking alive (and so calc_check can prove the caches match what
the formulas compute).
"""

import datetime


def today():
    return datetime.date(2026, 9, 18)


# ---------------------------------------------------------------------------
# cast
# ---------------------------------------------------------------------------
PEOPLE = [
    # name, team, household, rsvp, diet, last year gave to, status, notes
    ("Ava Novak", "Studio", "Novak home", "Yes", "", "Jack Byrne",
     "Complete", "Hosting helper - arrives early"),
    ("Ben Carter", "Shop", "", "Yes", "Vegetarian", "Ava Novak",
     "Gift Purchased", "Owes Kia a bet from 2025"),
    ("Chloe Diaz", "Online", "", "Yes", "Vegan", "Grace Liu",
     "Wrapped", ""),
    ("Dan Novak", "Studio", "Novak home", "Yes", "", "Elena Petrov",
     "Gift Purchased", "Ava's husband - couple rule applies"),
    ("Elena Petrov", "Shop", "", "Maybe", "Gluten-free", "Farid Khan",
     "Not Started", "Waiting on shift rota"),
    ("Farid Khan", "Shop", "", "Yes", "Halal", "Chloe Diaz",
     "Wrapped", ""),
    ("Grace Liu", "Online", "Flat 4B", "Yes", "", "Hugo Silva",
     "Complete", "Flat 4B - couple rule applies too"),
    ("Hugo Silva", "Online", "Flat 4B", "Yes", "", "Grace Liu",
     "Gift Purchased", ""),
    ("Isla Moore", "Studio", "", "No", "", "Leo Fontaine",
     "Not Started", "Away in December - sad face"),
    ("Jack Byrne", "Shop", "", "Yes", "Nut allergy", "Ava Novak",
     "Gift Purchased", "Espresso-machine feud with Leo"),
    ("Kia Osei", "Online", "", "Yes", "", "Ben Carter",
     "Wrapped", ""),
    ("Leo Fontaine", "Studio", "", "Yes", "Vegetarian", "Isla Moore",
     "Gift Purchased", ""),
]

CUSTOM_PAIRS = [
    ("Jack Byrne", "Leo Fontaine", "Espresso-machine feud of 2025"),
    ("Kia Osei", "Ben Carter", "Last year's regrettable candle"),
]

RULES = {"couple": "On", "team": "Off", "lastyear": "On"}

SPENT = [32.50, 24.00, 28.75, 31.20, 18.00, 26.40, 33.10, 42.00,
         0.00, 29.95, 27.30, 25.60]
RECEIPTS = ["\u2713", "", "\u2713", "", "", "\u2713", "", "\u2713", "",
            "\u2713", "", "\u2713"]
REFS = ["NV-88121", "", "AMZ-4071", "", "", "ETS-2210", "", "NV-88203",
        "", "AMZ-4188", "", "ETS-2355"]

WISHES = [
    ("Ava Novak", "Chunky knit throw blanket", "Must-love",
     "https://example.com/throw", "\u2713", "Any colour but grey"),
    ("Ava Novak", "Ceramic mug tree", "Nice to have", "", "", ""),
    ("Ben Carter", "Single-origin coffee beans 1kg", "Must-love",
     "https://example.com/beans", "\u2713", "No dark roasts"),
    ("Chloe Diaz", "Vegan candle trio", "Nice to have", "", "", ""),
    ("Chloe Diaz", "Linen tea towel set", "Just an idea", "", "", ""),
    ("Dan Novak", "Whisky tasting set", "Must-love",
     "https://example.com/whisky", "", "Peated, please"),
    ("Elena Petrov", "Gluten-free chocolate box", "Must-love", "", "", ""),
    ("Farid Khan", "Desk plant (low light)", "Nice to have", "", "\u2713", ""),
    ("Grace Liu", "Silk scrunchie set", "Just an idea", "", "", ""),
    ("Grace Liu", "Matcha ceremony kit", "Must-love",
     "https://example.com/matcha", "", ""),
    ("Hugo Silva", "Vinyl: jazz classics", "Nice to have", "", "", ""),
    ("Jack Byrne", "Espresso tamp mat", "Must-love", "", "", "Obvious reasons"),
    ("Kia Osei", "Botanical print A4", "Nice to have", "", "\u2713", ""),
    ("Leo Fontaine", "Sourdough starter kit", "Just an idea", "", "", ""),
    ("Leo Fontaine", "Wooden bread lame", "Nice to have", "", "", ""),
    ("Isla Moore", "Travel journal", "Nice to have", "", "", ""),
]

WE_GIFTS = [
    ("Giant pickle plushie", 12.00),
    ("Disco-ball bauble set", 15.50),
    ("Mystery sealed box", 20.00),
    ("Heated mug coaster", 18.25),
    ("Llama Christmas jumper", 22.00),
    ("Singing fish plaque", 14.75),
    ("Fancy hot chocolate tin", 16.40),
    ("Desktop zen garden", 19.90),
    ("Reindeer antler headband", 11.30),
    ("Marble cheese board", 24.50),
]

# round-1 turns then round-2 catch-up turns for the stolen-from players
HISTORY = [
    (1, "Pick", 4, ""),
    (2, "Pick", 7, ""),
    (3, "Steal", 4, None),     # from order 1
    (4, "Pick", 2, ""),
    (5, "Steal", 4, None),     # from order 3 -> gift 4 locks (2 steals)
    (6, "Pick", 9, ""),
    (7, "Steal", 9, None),     # from order 6
    (8, "Pick", 1, ""),
    (9, "Steal", 1, None),     # from order 8
    (10, "Pick", 6, ""),
    (1, "Pick", 3, ""),        # round 2: Ava replaces her stolen gift
    (2, "Pass", 7, ""),
    (3, "Pick", 5, ""),        # Chloe replaces hers
    (4, "Pass", 2, ""),
]


def _offset_for(people, pairs, rules):
    n = len(people)
    for off in range(1, n):
        ok = True
        for i in range(n):
            rec = people[(i + off) % n][0]
            g = people[i]
            if rec == g[0]:
                ok = False
            if rules["couple"] == "On" and g[2] and \
                    people[(i + off) % n][2] == g[2]:
                ok = False
            if rules["lastyear"] == "On" and rec == g[6 - 1]:
                ok = False
            if (g[0], rec) in pairs:
                ok = False
        if ok:
            return off
    return 1


def _rank(keys):
    """1-based rank, largest first (matches Excel RANK default)."""
    order = sorted(range(len(keys)), key=lambda i: -keys[i])
    out = [0] * len(keys)
    for pos, i in enumerate(order):
        out[i] = pos + 1
    return out


class Demo(object):
    def __init__(self):
        self.people = [dict(zip(
            ("name", "team", "household", "rsvp", "diet", "lastyear",
             "status", "notes"), p)) for p in PEOPLE]
        self.pairs = [(a, b) for a, b, _ in CUSTOM_PAIRS]
        self.pair_reason = {(a, b): r for a, b, r in CUSTOM_PAIRS}
        self.rules = dict(RULES)
        n = len(self.people)
        self.offset = _offset_for(PEOPLE, self.pairs, RULES)
        self.assign = [self.people[(i + self.offset) % n]["name"]
                       for i in range(n)]
        self.spent = list(SPENT)
        self.receipts = list(RECEIPTS)
        self.refs = list(REFS)
        self.wishes = WISHES
        self.settings = {
            "party": "Novality Store Holiday Party",
            "message": "19 Dec - bring a wrapped gift and a stealing face!",
            "currency": "$",
            "year": 2026,
            "date": datetime.date(2026, 12, 19),
            "location": "Novality Loft, 12 Market Street",
            "host": "The Novality Team",
            "bmin": 20.0, "bmax": 35.0, "steals": 2, "start": 1,
            "seed": self.offset, "weseed": 7,
            "status": "Final \u2014 locked \U0001F512",
        }
        self._build_we()
        self._agg()

    # ------------------------------------------------------------------
    def _build_we(self):
        s = self.settings["weseed"]
        players = [p["name"] for p in self.people if p["rsvp"] != "No"][:10]
        n = len(players)
        k1 = [((s * 7919 * (i + 2) + i * 104729) % 999983) * 100 + i
              for i in range(n)]
        k2 = [((s * 6967 * (i + 5) + i * 1299709) % 999983) * 100 + i
              for i in range(n)]
        seats = _rank(k1)              # seat order per player row
        nums = _rank(k2)               # gift number per player row
        self.we = []
        for i, pl in enumerate(players):
            self.we.append({
                "player": pl, "seat": seats[i], "giftnum": nums[i],
                "desc": WE_GIFTS[i][0], "value": WE_GIFTS[i][1],
            })
        by_seat = {w["seat"]: w for w in self.we}
        by_num = {w["giftnum"]: w for w in self.we}
        self.history = []
        holder = {}
        steals = {}
        for turn, (seat, action, gnum, _) in enumerate(HISTORY, 1):
            pl = by_seat[seat]["player"]
            frm = ""
            if action == "Steal":
                frm = holder.get(gnum, "")
                holder[gnum] = pl
                steals[gnum] = steals.get(gnum, 0) + 1
            elif action == "Pick":
                holder[gnum] = pl
            self.history.append((turn, pl, action, gnum, frm))
        self.holder = holder
        self.steals = steals
        self.last_action = {}
        for turn, pl, action, gnum, frm in self.history:
            self.last_action[gnum] = action
        for w in self.we:
            g = w["giftnum"]
            w["steals"] = steals.get(g, 0)
            w["holder"] = holder.get(g, "")
            w["last"] = self.last_action.get(g, "")
            complete = len(set(holder.values())) == len(self.we)
            if w["steals"] >= self.settings["steals"] or complete:
                w["status"] = "\U0001F512 Final"
            elif g not in holder:
                w["status"] = "\U0001F381 Available"
            elif w["last"] == "Steal":
                w["status"] = "\U0001F504 Stolen"
            else:
                w["status"] = "\U0001F932 Held"
            frm = ""
            for turn, pl, action, gn, f in reversed(self.history):
                if gn == g and action == "Steal":
                    frm = f
                    break
            w["stolenfrom"] = frm if w["status"] == "\U0001F504 Stolen" else ""
            w["final"] = ""
        self.we_players = players
        self.by_num = by_num

    # ------------------------------------------------------------------
    def money(self, value, decimals=2):
        cur = self.settings["currency"]
        if decimals:
            return "%s%.*f" % (cur, decimals, value)
        return "%s%d" % (cur, value)

    def _agg(self):
        n = len(self.people)
        st = [p["status"] for p in self.people]
        bought = sum(1 for s in st if s in ("Gift Purchased", "Wrapped",
                                            "Complete"))
        wrapped = sum(1 for s in st if s in ("Wrapped", "Complete"))
        done = sum(1 for s in st if s == "Complete")
        spent = sum(self.spent)
        rules = sum(1 for v in self.rules.values() if v == "On") + \
            len(self.pairs)
        in_play = sum(1 for w in self.we if w["status"] in
                      ("\U0001F932 Held", "\U0001F504 Stolen"))
        locked = sum(1 for w in self.we if w["status"] == "\U0001F512 Final")
        comp = (bought + wrapped + done) / (3.0 * n)
        self.agg = {
            "participants": n,
            "rsvp_yes": sum(1 for p in self.people if p["rsvp"] == "Yes"),
            "rsvp_pending": sum(1 for p in self.people
                                if p["rsvp"] == "Maybe"),
            "assigned": n,
            "purchased": bought,
            "wrapped": wrapped,
            "complete": done,
            "total_budget": n * self.settings["bmax"],
            "spent": spent,
            "avg_gift": spent / bought,
            "rules_count": rules,
            "we_players": len(self.we),
            "gifts_in_play": in_play,
            "locked_gifts": locked,
            "completion": comp,
            "turns_done": len(self.history),
            "current_turn": (self.settings["start"] - 1 +
                             len(self.history)) % len(self.we) + 1,
            "violations": 0,
            "over_budget": sum(1 for x in self.spent
                               if x > self.settings["bmax"]),
            "diets": sum(1 for p in self.people if p["diet"]),
        }
```

## 7. Workbook orchestration

`santa_tracker/workbook.py`

```python
"""Build orchestration: editions, themes, modes and product filenames."""

import os

from . import config as C
from . import theme as T
from .book import Book
from .sheets import (budget, cards, dashboard, data, draw, guide, history,
                     participants, rules, setup, we, wishlists)

BUILDERS = {
    "data": data.build,
    "setup": setup.build,
    "dashboard": dashboard.build,
    "participants": participants.build,
    "draw": draw.build,
    "rules": rules.build,
    "budget": budget.build,
    "wishlists": wishlists.build,
    "we": we.build,
    "history": history.build,
    "cards": cards.build,
    "guide": guide.build,
}


def product_filename(edition, theme_name, mode):
    return "Secret_Santa_White_Elephant_Tracker_%s_%s%s.xlsx" % (
        edition.upper(), T.THEMES[theme_name].label,
        "_EXAMPLE" if mode == "demo" else "")


def build_workbook(path, edition="premium", theme_name="noel",
                   mode="blank", protect=C.PROTECT_PASSWORD, images=True):
    """Write one workbook file and return its build stats."""
    from . import demo as D
    if edition not in ("basic", "premium"):
        raise ValueError("edition must be 'basic' or 'premium'")
    if mode not in ("blank", "demo"):
        raise ValueError("mode must be 'blank' or 'demo'")
    model = D.Demo() if mode == "demo" else None
    bk = Book(path, T.THEMES[theme_name], edition=edition, mode=mode,
              protect=protect, images=images, demo=model)
    for key in bk.order:
        BUILDERS[key](bk)
    bk.close()
    stats = dict(bk.stats)
    stats["size"] = os.path.getsize(path)
    return stats


def build_all(outdir, protect=C.PROTECT_PASSWORD, images=True):
    """The curated six-file Etsy product set."""
    os.makedirs(outdir, exist_ok=True)
    combos = [("premium", "noel", "demo"), ("premium", "noel", "blank"),
              ("premium", "arctic", "blank"), ("basic", "noel", "blank"),
              ("basic", "arctic", "blank"), ("basic", "noel", "demo")]
    out = []
    for edition, theme_name, mode in combos:
        name = product_filename(edition, theme_name, mode)
        path = os.path.join(outdir, name)
        stats = build_workbook(path, edition, theme_name, mode,
                               protect=protect, images=images)
        out.append((name, stats))
    return out
```

## 8. Sheets package init

`santa_tracker/sheets/__init__.py`

```python
```

## 9. Shared sheet helpers

`santa_tracker/sheets/common.py`

```python
"""Shared worksheet furniture: headers, data rows, totals, validations."""

from datetime import date, datetime

from .. import config as C
from ..book import r, ci, cl


# ---------------------------------------------------------------------------
# geometry helpers
# ---------------------------------------------------------------------------
def first_row():
    return C.ROW_FIRST


def last_row(key):
    return C.last_row(key)


def n_rows(key):
    return C.CAP[key]


def alt(rownum):
    """Zebra stripe: even spreadsheet rows get the tinted background."""
    return (rownum % 2) == 0


# ---------------------------------------------------------------------------
# headers
# ---------------------------------------------------------------------------
def header_row(bk, key, columns, row=None, height=36):
    """``columns`` = list of (field, label, kind, colour_key|None)."""
    ws = bk.ws(key)
    row = C.ROW_HEADER if row is None else row
    for field, label, kind, color in columns:
        col = ci(bk.col(key, field))
        ws.write(r(row), col, label,
                 bk.S.header(getattr(bk.th, color) if color else None))
    ws.set_row(r(row), height)
    return row


def data_rows(bk, key, height=20):
    ws = bk.ws(key)
    for i in range(n_rows(key)):
        ws.set_row(r(C.ROW_FIRST + i), height)


def table_frame(bk, key, columns, height=20):
    """Write the header and pre-format every empty data cell."""
    header_row(bk, key, columns)
    ws = bk.ws(key)
    data_rows(bk, key, height)
    for i in range(n_rows(key)):
        rownum = C.ROW_FIRST + i
        a = alt(rownum)
        for field, label, kind, color in columns:
            col = ci(bk.col(key, field))
            if kind == "idx":
                fmt = bk.S.idx(a)
            elif kind.startswith("calc"):
                fmt = bk.S.cell(kind)
            else:
                fmt = bk.S.cell(kind, a)
            ws.write_blank(r(rownum), col, None, fmt)
    return ws


# ---------------------------------------------------------------------------
# writing a single row of a tracker table
# ---------------------------------------------------------------------------
def write_row(bk, key, columns, rownum, values, cached=None):
    ws = bk.ws(key)
    a = alt(rownum)
    cached = cached or {}
    for field, label, kind, color in columns:
        col = ci(bk.col(key, field))
        value = values.get(field)
        if kind == "idx":
            fmt = bk.S.idx(a)
        elif kind.startswith("calc"):
            fmt = bk.S.cell(kind)
        else:
            fmt = bk.S.cell(kind, a)
        if value is None or value == "":
            ws.write_blank(r(rownum), col, None, fmt)
        elif isinstance(value, str) and value.startswith("="):
            ws.write_formula(r(rownum), col, value, fmt,
                             cached.get(field, 0))
            bk.stats["formulas"] += 1
            bk.stats["cells"] += 1
        elif isinstance(value, (date, datetime)):
            ws.write_datetime(r(rownum), col, value, fmt)
            bk.stats["cells"] += 1
        else:
            ws.write(r(rownum), col, value, fmt)
            bk.stats["cells"] += 1
    return ws


# ---------------------------------------------------------------------------
# validation presets
# ---------------------------------------------------------------------------
def list_dv(bk, key, field, list_key, title=None, message=None, error=None,
            first=None, last=None):
    return bk.validate(key, first or C.ROW_FIRST, ci(bk.col(key, field)),
                       last or C.last_row(key), ci(bk.col(key, field)),
                       "=" + bk.listname(list_key), title=title,
                       message=message, error=error)


def fixed_dv(bk, key, field, name, title=None, message=None, error=None):
    return bk.validate(key, C.ROW_FIRST, ci(bk.col(key, field)),
                       C.last_row(key), ci(bk.col(key, field)), "=" + name,
                       title=title, message=message, error=error)


def tick_dv(bk, key, fields, message="Pick \u2713 from the dropdown to tick "
                                     "it off (leave blank for not yet)."):
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        bk.validate(key, C.ROW_FIRST, ci(bk.col(key, field)),
                    C.last_row(key), ci(bk.col(key, field)), "=Tick",
                    title="Tick it off", message=message)


def money_dv(bk, key, fields, label="money"):
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        bk.ws(key).data_validation(
            r(C.ROW_FIRST), ci(bk.col(key, field)),
            r(C.last_row(key)), ci(bk.col(key, field)),
            {"validate": "decimal", "criteria": ">=", "value": 0,
             "ignore_blank": True, "show_error": True,
             "error_title": "Enter %s" % label,
             "error_message": "Please enter a positive number (no currency "
                              "symbol) - the symbol comes from Setup.",
             "error_type": "warning"})
        bk.stats["validations"] += 1


def date_dv(bk, key, fields, message="Type a date, or pick one from the "
                                     "calendar picker.", first=None,
            last=None):
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        bk.ws(key).data_validation(
            r(first or C.ROW_FIRST), ci(bk.col(key, field)),
            r(last or C.last_row(key)), ci(bk.col(key, field)),
            {"validate": "date", "criteria": "between",
             "minimum": date(2000, 1, 1), "maximum": date(2100, 12, 31),
             "ignore_blank": True, "show_input": True, "input_title": "Date",
             "input_message": message, "show_error": True,
             "error_title": "That's not a date",
             "error_message": "Enter a date between 2000 and 2100.",
             "error_type": "warning"})
        bk.stats["validations"] += 1


def whole_dv(bk, key, fields, minimum=0, maximum=9999):
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        bk.ws(key).data_validation(
            r(C.ROW_FIRST), ci(bk.col(key, field)),
            r(C.last_row(key)), ci(bk.col(key, field)),
            {"validate": "whole", "criteria": "between", "minimum": minimum,
             "maximum": maximum, "ignore_blank": True, "show_error": True,
             "error_title": "Whole numbers only",
             "error_message": "Please enter a whole number.",
             "error_type": "warning"})
        bk.stats["validations"] += 1


# ---------------------------------------------------------------------------
# conditional formatting presets
# ---------------------------------------------------------------------------
def tick_cf(bk, key, fields):
    """Green glow on any ticked box."""
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        col = ci(bk.col(key, field))
        L = bk.col(key, field)
        bk.cond(key, C.ROW_FIRST, col, C.last_row(key), col, {
            "type": "formula", "criteria": '=$%s%d="%s"' % (L, C.ROW_FIRST,
                                                            C.TICK),
            "format": bk.S.cf(bg=bk.th.ok_soft, fg=bk.th.ok, bold=True,
                              size=13, border=bk.th.border)})


def secret_cf(bk, key, field):
    """Secret Mode: make the hiding-spot text invisible (white on white)."""
    col = ci(bk.col(key, field))
    bk.cond(key, C.ROW_FIRST, col, C.last_row(key), col, {
        "type": "formula", "criteria": '=SecretMode="Yes"',
        "format": bk.S.cf(bg=bk.th.card, fg=bk.th.card)})


def status_cf(bk, key, field, mapping, first=None, last=None):
    """Colour a status column.  ``mapping`` = {cell text: (bg, fg)}."""
    L = bk.col(key, field)
    col = ci(L)
    first = first or C.ROW_FIRST
    last = last or C.last_row(key)
    for text, (bg, fg) in mapping.items():
        bk.cond(key, first, col, last, col, {
            "type": "formula",
            "criteria": '=$%s%d="%s"' % (L, first, text),
            "format": bk.S.cf(bg=bg, fg=fg, bold=True, border=bk.th.border)})


def deadline_cf(bk, key, date_field, tick_field=None):
    """Overdue = red, due inside the DueSoonDays window = amber."""
    L = bk.col(key, date_field)
    col = ci(L)
    first, last = C.ROW_FIRST, C.last_row(key)
    done = ""
    if tick_field:
        done = '*($%s%d<>"%s")' % (bk.col(key, tick_field), first, C.TICK)
    bk.cond(key, first, col, last, col, {
        "type": "formula",
        "criteria": '=AND($%s%d<>"",$%s%d<TODAY()%s)' % (L, first, L, first,
                                                         done),
        "format": bk.S.cf(bg=bk.th.bad_soft, fg=bk.th.bad, bold=True)})
    bk.cond(key, first, col, last, col, {
        "type": "formula",
        "criteria": '=AND($%s%d>=TODAY(),$%s%d-TODAY()<=DueSoonDays%s)'
                   % (L, first, L, first, done),
        "format": bk.S.cf(bg=bk.th.warn_soft, fg=bk.th.warn, bold=True)})


def databar(bk, key, field, color=None, first=None, last=None):
    col = ci(bk.col(key, field))
    bk.cond(key, first or C.ROW_FIRST, col, last or C.last_row(key), col, {
        "type": "data_bar", "bar_color": color or bk.th.primary_2,
        "bar_solid": True, "min_type": "num", "min_value": 0,
        "max_type": "num", "max_value": 1})


# ---------------------------------------------------------------------------
# totals row
# ---------------------------------------------------------------------------
def totals_row(bk, key, row, cells, first_col="B", last_col=None,
               label="TOTALS", label_span=None):
    """``cells`` = {field: (formula, kind, cached)}."""
    ws = bk.ws(key)
    th = bk.th
    lab_fmt = bk.S.f(**bk.S.base(font_size=11, bold=True, font_color=th.white,
                                 bg_color=th.primary, align="right",
                                 valign="vcenter", border=1,
                                 border_color=th.primary, indent=1))
    if label_span:
        c1, c2 = label_span
        if c1 == c2:
            ws.write(r(row), ci(c1), label, lab_fmt)
        else:
            ws.merge_range(r(row), ci(c1), r(row), ci(c2), label, lab_fmt)
    else:
        ws.write(r(row), ci(first_col), label, lab_fmt)
    for field, (formula, kind, cached) in cells.items():
        col = ci(bk.col(key, field))
        fmt = bk.S.f(**bk.S.base(
            font_size=11, bold=True, font_color=th.white, bg_color=th.primary,
            align="right" if kind != "center" else "center",
            valign="vcenter", border=1, border_color=th.primary,
            num_format=kind))
        if isinstance(formula, str) and formula.startswith("="):
            ws.write_formula(r(row), col, formula, fmt, cached or 0)
            bk.stats["formulas"] += 1
        else:
            ws.write(r(row), col, formula, fmt)
    if last_col:
        span_end = ci(label_span[1]) if label_span else ci(first_col)
        for c in range(ci(first_col), ci(last_col) + 1):
            letter = cl(c)
            if letter not in [bk.col(key, f) for f in cells] and \
                    c > span_end:
                ws.write_blank(r(row), c, None, bk.S.f(**bk.S.base(
                    bg_color=th.primary, border=1, border_color=th.primary)))
    ws.set_row(r(row), 24)
    return row


def blank_row(bk, key, row, first_col="A", last_col="N", height=8, fmt=None):
    ws = bk.ws(key)
    fmt = fmt or bk.S.canvas
    ws.set_row(r(row), height)
    for c in range(ci(first_col), ci(last_col) + 1):
        ws.write_blank(r(row), c, None, fmt)
    return row


def note_block(bk, key, row, first_col, last_col, lines, title=None):
    """A bordered, wrapped tip box."""
    ws = bk.ws(key)
    width = sum((C.WIDTHS.get(key, {}) or {}).get(cl(c), 10)
                for c in range(ci(first_col), ci(last_col) + 1))
    if title:
        ws.merge_range(r(row), ci(first_col), r(row), ci(last_col), title,
                       bk.S.section_soft)
        ws.set_row(r(row), 22)
        row += 1
    text = "\n".join(lines)
    from ..styles import wrap_height
    h = wrap_height(text, width, minimum=20)
    ws.merge_range(r(row), ci(first_col), r(row), ci(last_col), text,
                   bk.S.note)
    ws.set_row(r(row), h)
    return row + 1


# ===========================================================================
# shared page furniture (crochet product)
# ===========================================================================
def sheet_head(bk, key, last_col, title, sub):
    """Canvas, widths, title band and the home link."""
    from .. import config as _C
    S, th = bk.S, bk.th
    ws = bk.ws(key)
    bk.widths(key, _C.WIDTHS[key])
    try:
        depth = _C.last_row(key) + 10
    except KeyError:
        depth = 60
    bk.paint(key, 0, 0, depth, ci(last_col), S.canvas)
    ws.set_row(r(_C.ROW_SPACER_1), 7)
    ws.set_row(r(_C.ROW_TITLE), 32)
    ws.set_row(r(_C.ROW_SUBTITLE), 18)
    ws.set_row(r(_C.ROW_SPACER_2), 8)
    ws.set_row(r(_C.ROW_STATS), 22)
    ws.set_row(r(_C.ROW_SPACER_3), 8)
    ws.merge_range(r(_C.ROW_TITLE), 1, r(_C.ROW_TITLE), ci(last_col) - 3,
                   title, S.sheet_title)
    ws.merge_range(r(_C.ROW_SUBTITLE), 1, r(_C.ROW_SUBTITLE),
                   ci(last_col) - 3, "  " + sub, S.sheet_sub)
    ws.merge_range(r(_C.ROW_SUBTITLE), ci(last_col) - 2,
                   r(_C.ROW_SUBTITLE), ci(last_col), "", S.home_link)
    ws.write_url(r(_C.ROW_SUBTITLE), ci(last_col) - 2,
                 "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                 "\U0001F3E0  Dashboard")


def chips(bk, key, chips, last_col=None):
    """Stats pills on ROW_STATS. chips = (formula, colour, cached, span)."""
    from .. import config as _C
    S, th = bk.S, bk.th
    ws = bk.ws(key)
    last_col = last_col or _C.cols_last(key)
    col = 1
    for formula, color, cached, span in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(_C.ROW_STATS), col, r(_C.ROW_STATS),
                       col + span - 1, "", fmt)
        ws.write_formula(r(_C.ROW_STATS), col,
                         formula if formula.startswith("=")
                         else '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += span
    if col <= ci(last_col):
        ws.merge_range(r(_C.ROW_STATS), col, r(_C.ROW_STATS),
                       ci(last_col), "", S.canvas)


def footer_nav(bk, key, row, last_col, landscape=True, zoom=90, tip=None):
    """Gold tip strip, nav grid and print setup."""
    S, th = bk.S, bk.th
    ws = bk.ws(key)
    ws.set_row(r(row), 30)
    ws.merge_range(r(row), 1, r(row), ci(last_col),
                   tip or "  \U0001F4A1  Cream cells are yours to type in; "
                          "white cells calculate.  Dropdowns keep every "
                          "column honest.",
                   S.f(**S.base(font_size=10, italic=True,
                                font_color=th.ink, bg_color=th.gold_soft,
                                align="left", valign="vcenter", indent=1)))
    nav = row + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(key, nav, max_col=last_col)
    bk.page(key, last_col, nav + 1, landscape=landscape, zoom=zoom)
```

## 10. _Data - hidden KPI engine

`santa_tracker/sheets/data.py`

```python
"""_Data - the hidden calculation engine (KPIs for the dashboard)."""

from .. import config as C
from ..book import r, ci


def build(bk):
    key = "data"
    ws = bk.ws(key)
    for col, width in (("A", 3), ("AH", 26), ("AI", 12)):
        ws.set_column(ci(col), ci(col), width)
    m = bk.demo
    premium = bk.has("we")
    has_rules = bk.has("rules")

    hdr = bk.S.f(**bk.S.base(font_size=9, bold=True, font_color=bk.th.white,
                             bg_color=bk.th.muted, align="center",
                             valign="vcenter"))
    ws.write(r(1), ci("AH"), "Metric", hdr)
    ws.write(r(1), ci("AI"), "Value", hdr)

    names = bk.rng("participants", "name")
    rsvp = bk.rng("participants", "rsvp")
    diet = bk.rng("participants", "diet")
    status = bk.rng("participants", "status")
    dfinal = bk.rng("draw", "final")
    dflag = bk.rng("draw", "flag")
    bspent = bk.rng("budget", "spent")
    bflag = bk.rng("budget", "flag")

    F = {}
    F["participants"] = ("=COUNTA(%s)" % names, "#,##0")
    F["rsvp_yes"] = ('=COUNTIF(%s,"Yes")' % rsvp, "#,##0")
    F["rsvp_pending"] = ('=COUNTIF(%s,"Maybe")' % rsvp, "#,##0")
    F["assigned"] = ('=SUMPRODUCT(--(%s<>""))' % dfinal, "#,##0")
    F["purchased"] = ('=COUNTIF(%s,"%s")+COUNTIF(%s,"%s")+COUNTIF(%s,"%s")'
                      % (status, C.ST_BOUGHT, status, C.ST_WRAPPED,
                         status, C.ST_DONE), "#,##0")
    F["wrapped"] = ('=COUNTIF(%s,"%s")+COUNTIF(%s,"%s")'
                    % (status, C.ST_WRAPPED, status, C.ST_DONE), "#,##0")
    F["complete"] = ('=COUNTIF(%s,"%s")' % (status, C.ST_DONE), "#,##0")
    F["total_budget"] = ("=$AI$%d*BudgetMax" % C.KPI_ROW["participants"],
                         "#,##0.00")
    F["spent"] = ("=SUM(%s)" % bspent, "#,##0.00")
    F["avg_gift"] = ('=IFERROR($AI$%d/$AI$%d,0)'
                     % (C.KPI_ROW["spent"], C.KPI_ROW["purchased"]),
                     "#,##0.00")
    rules = ('SUMPRODUCT(--(%s<>""))' % bk.rng("rules", "giver")) \
        if has_rules else "0"
    F["rules_count"] = (
        '=(CoupleRule="On")+(TeamRule="On")+(LastYearRule="On")+%s'
        % rules if has_rules else "=0", "#,##0")
    if premium:
        wplayer = bk.rng("we", "player")
        wstatus = bk.rng("we", "status")
        hturn = bk.rng("history", "turn")
        F["we_players"] = ("=COUNTA(%s)" % wplayer, "#,##0")
        F["gifts_in_play"] = ('=COUNTIF(%s,"%s")+COUNTIF(%s,"%s")'
                              % (wstatus, C.GS_HELD, wstatus, C.GS_STOLEN),
                              "#,##0")
        F["locked_gifts"] = ('=COUNTIF(%s,"%s")' % (wstatus, C.GS_FINAL),
                             "#,##0")
        F["turns_done"] = ("=COUNT(%s)" % hturn, "#,##0")
        F["current_turn"] = ('=IF($AI$%d=0,"",MOD(StartPlayer-1+$AI$%d,'
                             '$AI$%d)+1)'
                             % (C.KPI_ROW["we_players"],
                                C.KPI_ROW["turns_done"],
                                C.KPI_ROW["we_players"]), "#,##0")
    else:
        for k in ("we_players", "gifts_in_play", "locked_gifts",
                  "turns_done", "current_turn"):
            F[k] = ("=0", "#,##0")
    F["completion"] = ('=IFERROR(($AI$%d+$AI$%d+$AI$%d)/(3*$AI$%d),0)'
                       % (C.KPI_ROW["purchased"], C.KPI_ROW["wrapped"],
                          C.KPI_ROW["complete"],
                          C.KPI_ROW["participants"]), "0%")
    F["violations"] = ('=COUNTIF(%s,"?*")-COUNTIF(%s,"%s OK")'
                       % (dflag, dflag, "\u2705"), "#,##0")
    F["over_budget"] = ('=COUNTIF(%s,"%s")' % (bflag, C.BF_OVER), "#,##0")
    F["diets"] = ("=COUNTA(%s)" % diet, "#,##0")

    # small pools for dashboard charts (status + rsvp mixes)
    pool = bk.S.f(**bk.S.base(font_size=9, font_color=bk.th.muted,
                              align="left", valign="vcenter"))
    pval = bk.S.f(**bk.S.base(font_size=9, font_color=bk.th.ink,
                              align="right", valign="vcenter",
                              num_format="#,##0"))
    for i, stt in enumerate((C.ST_NOT, C.ST_BOUGHT, C.ST_WRAPPED,
                             C.ST_DONE)):
        ws.write(r(40 + i), ci("AH"), stt, pool)
        ws.write_formula(r(40 + i), ci("AI"), '=COUNTIF(%s,"%s")'
                         % (status, stt), pval,
                         (sum(1 for x in (m.people if m else [])
                              if x["status"] == stt) if m else 0))
        bk.stats["formulas"] += 1
    for i, rv in enumerate(("Yes", "No", "Maybe")):
        ws.write(r(45 + i), ci("AH"), rv, pool)
        ws.write_formula(r(45 + i), ci("AI"), '=COUNTIF(%s,"%s")'
                         % (rsvp, rv), pval,
                         (sum(1 for x in (m.people if m else [])
                              if x["rsvp"] == rv) if m else 0))
        bk.stats["formulas"] += 1

    lab = bk.S.f(**bk.S.base(font_size=9.5, font_color=bk.th.muted,
                             align="left", valign="vcenter"))
    val = bk.S.f(**bk.S.base(font_size=9.5, font_color=bk.th.ink,
                             align="right", valign="vcenter"))
    for name, row in sorted(C.KPI_ROW.items()):
        formula, fmt = F[name]
        ws.write(r(row), ci("AH"), name, lab)
        cached = None
        if m:
            cached = m.agg.get(name, 0)
        ws.write_formula(r(row), ci("AI"), formula,
                         bk.S.f(**bk.S.base(font_size=9.5,
                                            font_color=bk.th.ink,
                                            align="right", valign="vcenter",
                                            num_format=fmt)),
                         cached)
        bk.stats["formulas"] += 1
```

## 11. Settings & Instructions

`santa_tracker/sheets/setup.py`

```python
"""⚙️ Settings & Instructions - party info, rules of the game, lists."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "setup"
LAST_COL = "H"


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\u2699\ufe0f  Settings & Instructions",
                 "Everything the party needs to know: dates, budgets, "
                 "steal limits and your own dropdown lists.")
    bk.paint(KEY, 0, 0, 70, ci(LAST_COL), bk.S.canvas)

    label = bk.S.f(**bk.S.base(font_size=10.5, bold=True, font_color=th.ink,
                               bg_color=th.card, align="left",
                               valign="vcenter", indent=1, border=1,
                               border_color=th.border))
    text = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                              bg_color=th.gold_soft, border=1,
                              border_color=th.border_strong, align="left",
                              valign="vcenter", indent=1, locked=False))
    num = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                             bg_color=th.gold_soft, border=1,
                             border_color=th.border_strong, align="right",
                             valign="vcenter", locked=False,
                             num_format="#,##0.00"))
    intf = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                              bg_color=th.gold_soft, border=1,
                              border_color=th.border_strong, align="right",
                              valign="vcenter", locked=False,
                              num_format="#,##0"))
    datef = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                               bg_color=th.gold_soft, border=1,
                               border_color=th.border_strong, align="right",
                               valign="vcenter", locked=False,
                               num_format="dd mmm yyyy"))
    pick = bk.S.f(**bk.S.base(font_size=10.5, bold=True,
                              font_color=th.primary, bg_color=th.gold_soft,
                              border=1, border_color=th.border_strong,
                              align="left", valign="vcenter", indent=1,
                              locked=False))

    def setting(row, text_, value, fmt, cached=None):
        ws.set_row(r(row), 22)
        ws.write(r(row), 1, text_, label)
        ws.write(r(row), 2, value if cached is None else cached, fmt)
        if cached is not None:
            ws.write(r(row), 2, value, fmt)

    ws.merge_range(r(4), 1, r(4), ci(LAST_COL),
                   "  \U0001F389  THE PARTY", bk.S.section_soft)
    setting(C.SU_BUSINESS, "Party name", "Your party name", pick,
            m.settings["party"] if m else None)
    setting(C.SU_MESSAGE, "Dashboard message of the day",
            "Type a message for the dashboard banner", pick,
            m.settings["message"] if m else None)
    ws.merge_range(r(9), 1, r(9), ci(LAST_COL),
                   "  \U0001F39B\ufe0f  RULES OF THE GAME", bk.S.section_soft)
    setting(C.SU_CURRENCY, "Currency symbol", "$", text,
            m.settings["currency"] if m else None)
    setting(C.SU_YEAR, "Party year (used everywhere)", 2026, intf,
            m.settings["year"] if m else None)
    setting(C.SU_DATE, "Party date", None, datef)
    if m:
        ws.write(r(C.SU_DATE), 2, m.settings["date"], datef)
    setting(C.SU_LOCATION, "Location", "Where is the party?", text,
            m.settings["location"] if m else None)
    setting(C.SU_HOST, "Host", "Who is hosting?", text,
            m.settings["host"] if m else None)
    setting(C.SU_BMIN, "Minimum gift budget", 20, num,
            m.settings["bmin"] if m else None)
    setting(C.SU_BMAX, "Maximum gift budget", 35, num,
            m.settings["bmax"] if m else None)
    setting(C.SU_STEALS, "Max steals per gift (White Elephant)", 2, intf,
            m.settings["steals"] if m else None)
    setting(C.SU_START, "Starting player seat (White Elephant)", 1, intf,
            m.settings["start"] if m else None)
    setting(C.SU_SEED, "Secret Santa draw seed (change = re-draw)", 5, intf,
            m.settings["seed"] if m else None)
    setting(C.SU_WESEED, "White Elephant seat seed", 7, intf,
            m.settings["weseed"] if m else None)
    setting(C.SU_STATUS, "Draw status", C.DRAW_DRAFT, pick,
            m.settings["status"] if m else None)
    bk.validate(KEY, C.SU_STATUS, 2, C.SU_STATUS, 2,
                '="%s,%s"' % (C.DRAW_DRAFT, C.DRAW_FINAL),
                title="Draw status")

    # ------------------------------------------------------------------
    ws.merge_range(r(C.SU_LIST_HEADER), 1, r(C.SU_LIST_HEADER), ci(LAST_COL),
                   "  \U0001F4DD  YOUR LISTS - rename or add, every dropdown "
                   "follows", bk.S.section_soft)
    ws.set_row(r(C.SU_LIST_HEADER), 20)
    for col, title in ((ci("C"), "Teams / departments"),
                       (ci("D"), "Dietary restrictions")):
        ws.write(r(C.SU_LIST_HEADER + 1), col, title, bk.S.thead)
    list_fmt = bk.S.f(**bk.S.base(font_size=10, font_color=th.ink,
                                  bg_color=th.card, border=1,
                                  border_color=th.border, locked=False,
                                  align="left", valign="vcenter", indent=1))
    defaults = {"C": ["Studio", "Shop", "Online", "Friends", "Family"],
                "D": ["Vegetarian", "Vegan", "Gluten-free", "Halal",
                      "Nut allergy"]}
    for i in range(C.SU_LIST_ROWS):
        row = C.SU_LIST_FIRST + i
        ws.set_row(r(row), 18)
        for col in ("C", "D"):
            val = defaults[col][i] if i < len(defaults[col]) else ""
            ws.write(r(row), ci(col), val, list_fmt)

    # fixed lists (locked)
    ws.merge_range(r(C.SU_FIXED_FIRST - 2), 1, r(C.SU_FIXED_FIRST - 2),
                   ci(LAST_COL),
                   "  \U0001F512  BUILT-IN LISTS (part of the engine - "
                   "do not edit)", bk.S.section_soft)
    fixed_fmt = bk.S.f(**bk.S.base(font_size=9.5, font_color=th.muted,
                                   align="left", valign="vcenter",
                                   indent=1))
    blocks = {"Statuses": [C.ST_NOT, C.ST_BOUGHT, C.ST_WRAPPED, C.ST_DONE],
              "RSVP": ["Yes", "No", "Maybe"],
              "Actions": ["Pick", "Steal", "Pass"],
              "Priorities": ["Must-love", "Nice to have", "Just an idea"],
              "Tick": [C.TICK, ""],
              "OnOff": ["On", "Off"]}
    for nm, (col, n) in sorted(C.FIXED_BLOCKS.items()):
        ws.write(r(C.SU_FIXED_FIRST - 1), ci(col), nm, bk.S.thead)
        for i, val in enumerate(blocks[nm]):
            ws.write(r(C.SU_FIXED_FIRST + i), ci(col), val, fixed_fmt)

    K.note_block(bk, KEY, C.SU_FIXED_FIRST + 6, "B", "H", [
        "Formula cells are locked so nobody breaks the maths. The sheet "
        "password is \u201cpremium\u201d - unlock via Review \u2192 Unprotect "
        "Sheet if you ever need to change a seed or a formula.",
        "Re-draw the Secret Santa any time: unlock, change the draw seed on "
        "this tab, watch \U0001F385 Secret Santa Draw shuffle, then set Draw "
        "status to Final and re-protect.",
        "White Elephant: log every Pick / Steal / Pass on \U0001F504 Game "
        "History - the \U0001F3B2 White Elephant tab updates holders, steal "
        "counts, locks and final gifts by itself.",
    ], title="\U0001F4D6  HOW TO RUN THE PARTY")
    K.footer_nav(bk, KEY, C.SU_FIXED_FIRST + 12, LAST_COL, landscape=False)
```

## 12. Dashboard

`santa_tracker/sheets/dashboard.py`

```python
"""🏠 Dashboard - party command centre."""

from .. import config as C
from ..book import r, ci
from . import common as K

import datetime as _dt

KEY = "dashboard"
LAST_COL = "N"


def _today():
    from .. import demo as D
    return D.today()


def build(bk):
    th = bk.th
    S = bk.S
    ws = bk.ws(KEY)
    m = bk.demo
    agg = m.agg if m else {}
    bk.widths(KEY, C.WIDTHS[KEY])
    bk.paint(KEY, 0, 0, C.DASH_NAV + 2, ci(LAST_COL), S.canvas)

    # ------------------------------------------------------------------
    # hero
    # ------------------------------------------------------------------
    ws.set_row(r(1), 8)
    ws.set_row(r(2), 34)
    ws.set_row(r(3), 26)
    ws.set_row(r(4), 20)
    ws.set_row(r(5), 22)
    ws.merge_range(r(2), 1, r(2), ci(LAST_COL), "", S.hero_title)
    ws.write_formula(
        r(2), 1,
        '=IF(PartyName="","\U0001F385  Secret Santa & White Elephant '
        'Command Centre","\U0001F385  "&PartyName&"   \u2022   Secret '
        'Santa & White Elephant Command Centre")', S.hero_title,
        "\U0001F385  %s   \u2022   Secret Santa & White Elephant Command "
        "Centre" % (m.settings["party"] if m else ""))
    bk.stats["formulas"] += 1
    ws.merge_range(r(3), 1, r(3), ci(LAST_COL), "", S.hero_count)
    ws.write_formula(
        r(3), 1,
        '=IF(PartyDate="","", "\U0001F384  "&TEXT(PartyDate,"dddd, dd '
        'mmmm yyyy")&"   \u2022   "&PartyLocation&"   \u2022   hosted by '
        '"&PartyHost&"   \u2022   "&MAX(0,PartyDate-TODAY())&" days to '
        'go!")', S.hero_count,
        "\U0001F384  Saturday, 19 December 2026   \u2022   %s   \u2022   "
        "hosted by %s   \u2022   %d days to go!"
        % (m.settings["location"], m.settings["host"],
           (m.settings["date"] - _today()).days) if m else "")
    bk.stats["formulas"] += 1
    ws.merge_range(r(4), 1, r(4), ci(LAST_COL), "", S.hero_meta)
    ws.write_formula(
        r(4), 1,
        '="\U0001F465 "&%s&" guests   \u2022   "&%s&" in the draw   '
        '\u2022   "&%s&" white-elephant players   \u2022   draw: "&'
        'DrawStatus' % (bk.kpi("participants"), bk.kpi("assigned"),
                        bk.kpi("we_players")),
        S.hero_meta,
        "\U0001F465 %d guests   \u2022   %d in the draw   \u2022   %d "
        "white-elephant players   \u2022   draw: %s"
        % (agg.get("participants", 0), agg.get("assigned", 0),
           agg.get("we_players", 0),
           m.settings["status"] if m else ""))
    bk.stats["formulas"] += 1
    msgfmt = S.f(**S.base(bg_color=th.gold, font_color=th.white, bold=True,
                          font_size=10.5, align="left", valign="vcenter",
                          indent=1))
    ws.merge_range(r(5), 1, r(5), ci(LAST_COL), "", msgfmt)
    ws.write_formula(r(5), 1, '="\U0001F4AC  "&Message', msgfmt,
                     "\U0001F4AC  %s" % (m.settings["message"] if m
                                          else ""))
    bk.stats["formulas"] += 1

    ws.merge_range(r(7), 1, r(7), ci(LAST_COL),
                   "  \U0001F389  PARTY AT A GLANCE", S.section)
    ws.set_row(r(7), 22)

    cards1 = [
        ("PARTICIPANTS", "primary", "=%s" % bk.kpi("participants"),
         agg.get("participants", 0)),
        ("SANTA ASSIGNED", "ok", "=%s&\" of \"&%s"
         % (bk.kpi("assigned"), bk.kpi("participants")),
         "%d of %d" % (agg.get("assigned", 0), agg.get("participants", 0))),
        ("GIFTS PURCHASED", "info", "=%s" % bk.kpi("purchased"),
         agg.get("purchased", 0)),
        ("TOTAL BUDGET", "gold", bk.money(bk.kpi("total_budget"),
                                          "#,##0"),
         agg.get("total_budget", 0)),
        ("ACTUAL SPEND", "accent", bk.money(bk.kpi("spent"), "#,##0"),
         agg.get("spent", 0)),
        ("AVG GIFT COST", "primary_2", bk.money(bk.kpi("avg_gift"),
                                                "#,##0.00"),
         agg.get("avg_gift", 0)),
    ]
    cards2 = [
        ("EXCLUSION RULES", "primary", "=%s" % bk.kpi("rules_count"),
         agg.get("rules_count", 0)),
        ("W-E PLAYERS", "info", "=%s" % bk.kpi("we_players"),
         agg.get("we_players", 0)),
        ("GIFTS IN PLAY", "accent", "=%s" % bk.kpi("gifts_in_play"),
         agg.get("gifts_in_play", 0)),
        ("LOCKED GIFTS", "bad", "=%s" % bk.kpi("locked_gifts"),
         agg.get("locked_gifts", 0)),
        ("RSVP YES", "ok", "=%s&\" / \"&%s"
         % (bk.kpi("rsvp_yes"), bk.kpi("participants")),
         "%d / %d" % (agg.get("rsvp_yes", 0), agg.get("participants", 0))),
        ("COMPLETION", "gold", "=TEXT(%s,\"0%%\")" % bk.kpi("completion"),
         "%.0f%%" % (100 * agg.get("completion", 0))),
    ]
    _cards(bk, cards1, C.DASH_CARD1_L, C.DASH_CARD1_V)
    _cards(bk, cards2, C.DASH_CARD2_L, C.DASH_CARD2_V)

    _bar(bk, C.DASH_BAR1, "Gift readiness", bk.kpi("purchased"),
         bk.kpi("participants"), agg.get("participants", 0) and
         agg["purchased"] / float(agg["participants"]), "pct")
    _bar(bk, C.DASH_BAR2, "RSVP confirmed", bk.kpi("rsvp_yes"),
         bk.kpi("participants"), agg.get("participants", 0) and
         agg["rsvp_yes"] / float(agg["participants"]), "pct")

    _charts(bk)

    # ------------------------------------------------------------------
    # panels
    # ------------------------------------------------------------------
    row = C.DASH_PANEL
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  \U0001F3B2  GAME STATE   \u2022   \u26A0\ufe0F  PARTY "
                   "ALERTS", S.section)
    ws.set_row(r(row), 22)
    lab = S.f(**S.base(font_size=10, bold=True, font_color=th.ink,
                       bg_color=th.card, align="left", indent=1,
                       valign="vcenter", border=1, border_color=th.border))
    val = S.f(**S.base(font_size=10.5, bold=True, font_color=th.primary,
                       bg_color=th.card, align="left", indent=1,
                       valign="vcenter", border=1, border_color=th.border))
    left = [
        ("Turns logged", "=%s" % bk.kpi("turns_done"),
         agg.get("turns_done", 0)),
        ("Now playing (seat)", "=%s" % bk.kpi("current_turn"),
         agg.get("current_turn", 0)),
        ("Gifts still stealable", "=%s" % bk.kpi("gifts_in_play"),
         agg.get("gifts_in_play", 0)),
        ("Gifts locked \U0001F512", "=%s" % bk.kpi("locked_gifts"),
         agg.get("locked_gifts", 0)),
    ]
    right = [
        ("RSVP still maybe", "=%s" % bk.kpi("rsvp_pending"),
         agg.get("rsvp_pending", 0)),
        ("Over-budget gifts", "=%s" % bk.kpi("over_budget"),
         agg.get("over_budget", 0)),
        ("Draw rule hits", "=%s" % bk.kpi("violations"),
         agg.get("violations", 0)),
        ("Dietary needs noted", "=%s" % bk.kpi("diets"),
         agg.get("diets", 0)),
    ]
    for i, ((lt, lf, lc), (rt_, rf, rc)) in enumerate(zip(left, right)):
        rr = row + 1 + i
        ws.set_row(r(rr), 20)
        ws.merge_range(r(rr), 1, r(rr), 4, "  " + lt, lab)
        ws.merge_range(r(rr), 5, r(rr), 7, "", val)
        ws.write_formula(r(rr), 5, lf, val, lc)
        bk.stats["formulas"] += 1
        ws.merge_range(r(rr), 8, r(rr), 11, "  " + rt_, lab)
        ws.merge_range(r(rr), 12, r(rr), ci(LAST_COL), "", val)
        ws.write_formula(r(rr), 12, rf, val, rc)
        bk.stats["formulas"] += 1

    K.footer_nav(bk, KEY, C.DASH_NAV, LAST_COL, landscape=True, tip=
                 "  \U0001F4AC  %s" % "Novality Store wishes you a loud, "
                 "glorious, steal-happy party.")


# ===========================================================================
def _cards(bk, cards, row_l, row_v):
    S, th = bk.S, bk.th
    ws = bk.ws("dashboard")
    pairs = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]
    for (label, colour, formula, cached), (c1, c2) in zip(cards, pairs):
        bg = getattr(th, colour)
        lbl = S.f(**S.base(bold=True, font_size=9.5, font_color=th.white,
                           bg_color=bg, align="left", valign="vcenter",
                           indent=1))
        vfmt = S.f(**S.base(bold=True, font_size=15,
                            font_color=bg if colour != "gold" else th.gold,
                            bg_color=th.card, align="center",
                            valign="vcenter"))
        ws.merge_range(r(row_l), c1, r(row_l), c2, "  " + label, lbl)
        ws.merge_range(r(row_v), c1, r(row_v), c2, "", vfmt)
        ws.write_formula(r(row_v), c1,
                         formula if formula.startswith("=")
                         else "=" + formula, vfmt, cached)
        bk.stats["formulas"] += 1
    ws.set_row(r(row_l), 18)
    ws.set_row(r(row_v), 26)


def _bar(bk, row, label, num_k, den_k, cached_pct, kind=None):
    S, th = bk.S, bk.th
    ws = bk.ws("dashboard")
    ws.set_row(r(row), 20)
    ws.merge_range(r(row), 1, r(row), 2, "  " + label, S.bar_label)
    ws.merge_range(r(row), 3, r(row), 9, "", S.bar_text)
    ws.write_formula(r(row), 3, bk.bar(num_k, den_k, 34), S.bar_text,
                     _bar_cached(cached_pct, 34))
    bk.stats["formulas"] += 1
    ws.merge_range(r(row), 10, r(row), 12, "", S.bar_pct)
    ws.write_formula(r(row), 10,
                     '=TEXT(IFERROR(%s/%s,0),"0.0%%")' % (num_k, den_k),
                     S.bar_pct, "%.1f%%" % (100.0 * float(cached_pct or 0)))
    bk.stats["formulas"] += 1


def _bar_cached(pct, blocks):
    filled = int(round(min(1.0, max(0.0, float(pct or 0))) * blocks))
    return "\u2588" * filled + "\u2591" * (blocks - filled)


# ===========================================================================
def _charts(bk):
    S, th = bk.S, bk.th
    wb = bk.wb
    d = bk.q("data").strip("'")
    premium = bk.has("we")

    def base_chart(ctype, title):
        ch = wb.add_chart({"type": ctype})
        ch.set_title({"name": title, "name_font": {"size": 11,
                                                   "color": th.primary,
                                                   "bold": True}})
        ch.set_size({"width": 470, "height": 250})
        ch.set_chartarea({"border": {"color": th.border},
                          "fill": {"color": th.card}})
        ch.set_legend({"position": "none"})
        ch.show_hidden_data()
        return ch

    ch1 = base_chart("column", "Actual spend per giver")
    ch1.add_series({
        "name": "Spent",
        "categories": "=%s!$B$%d:$B$%d" % (bk.q("participants").strip("'"),
                                           C.ROW_FIRST,
                                           C.last_row("participants")),
        "values": "=%s!$E$%d:$E$%d" % (bk.q("budget").strip("'"),
                                       C.ROW_FIRST, C.last_row("budget")),
        "fill": {"color": th.accent},
        "border": {"color": th.border_strong},
    })
    ws = bk.ws("dashboard")
    ws.insert_chart("C%d" % (C.DASH_CHART1 + 1), ch1)

    ch2 = base_chart("doughnut", "Gift readiness mix")
    ch2.add_series({
        "name": "Status",
        "categories": "=%s!$AH$40:$AH$43" % d,
        "values": "=%s!$AI$40:$AI$43" % d,
        "points": [{"fill": {"color": th.bad}},
                   {"fill": {"color": th.warn}},
                   {"fill": {"color": th.info}},
                   {"fill": {"color": th.ok}}],
    })
    ch2.set_legend({"position": "bottom", "font": {"size": 9,
                                                   "color": th.muted}})
    ws.insert_chart("I%d" % (C.DASH_CHART1 + 1), ch2)

    if premium:
        ch3 = base_chart("bar", "Steals per gift")
        wq = bk.q("we").strip("'")
        ch3.add_series({
            "name": "Steals",
            "categories": "=%s!$D$%d:$D$%d" % (wq, C.ROW_FIRST,
                                               C.last_row("we")),
            "values": "=%s!$G$%d:$G$%d" % (wq, C.ROW_FIRST,
                                           C.last_row("we")),
            "fill": {"color": th.primary_2},
        })
        ws.insert_chart("C%d" % (C.DASH_CHART2 + 1), ch3)

        ch4 = base_chart("column", "Gift values brought")
        ch4.add_series({
            "name": "Value",
            "categories": "=%s!$D$%d:$D$%d" % (wq, C.ROW_FIRST,
                                               C.last_row("we")),
            "values": "=%s!$F$%d:$F$%d" % (wq, C.ROW_FIRST,
                                           C.last_row("we")),
            "fill": {"color": th.gold},
        })
        ws.insert_chart("I%d" % (C.DASH_CHART2 + 1), ch4)
    else:
        ch3 = base_chart("doughnut", "RSVP mix")
        ch3.add_series({
            "name": "RSVP",
            "categories": "=%s!$AH$45:$AH$47" % d,
            "values": "=%s!$AI$45:$AI$47" % d,
            "points": [{"fill": {"color": th.ok}},
                       {"fill": {"color": th.bad}},
                       {"fill": {"color": th.warn}}],
        })
        ch3.set_legend({"position": "bottom", "font": {"size": 9,
                                                       "color": th.muted}})
        ws.insert_chart("C%d" % (C.DASH_CHART2 + 1), ch3)
```

## 13. Participants

`santa_tracker/sheets/participants.py`

```python
"""👥 Participants - the guest list with RSVP, teams and status."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "participants"
LAST_COL = "J"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Participant", "text", None),
    ("team", "Team / dept", "center", None),
    ("household", "Household / couple", "center", None),
    ("rsvp", "RSVP", "center", None),
    ("diet", "Dietary needs", "center", None),
    ("lastyear", "Gave to last year", "center", None),
    ("status", "Gift status", "center", None),
    ("wishes", "Wishes", "calc_num", None),
    ("notes", "Notes", "wrap", None),
]


def build(bk):
    th = bk.th
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F465  Participants",
                 "Everyone in the draw - RSVP, household and team feed the "
                 "exclusion rules automatically.")
    K.table_frame(bk, KEY, COLUMNS, height=20)
    if m:
        for i, p in enumerate(m.people):
            row = C.ROW_FIRST + i
            K.write_row(bk, KEY, COLUMNS, row,
                        {"n": i + 1, "name": p["name"], "team": p["team"],
                         "household": p["household"], "rsvp": p["rsvp"],
                         "diet": p["diet"], "lastyear": p["lastyear"],
                         "status": p["status"], "notes": p["notes"] or ""},
                        {"wishes": sum(1 for w in m.wishes
                                       if w[0] == p["name"])})
    for i in range(len(m.people) if m else 0, C.CAP[KEY]):
        K.write_row(bk, KEY, COLUMNS, C.ROW_FIRST + i,
                    {"n": i + 1, "name": "", "team": "", "household": "",
                     "rsvp": "", "diet": "", "lastyear": "", "status": "",
                     "wishes": "", "notes": ""})
    # wishes count formula per row
    wrng = bk.rng("wishlists", "who") if bk.has("wishlists") else None
    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        if wrng:
            bk.ws(KEY).write_formula(
                r(row), ci(bk.col(KEY, "wishes")),
                '=IF($B%d="","",COUNTIF(%s,$B%d))' % (row, wrng, row),
                bk.S.f(**bk.S.base(font_size=10.5, font_color=th.primary,
                                   bg_color=th.alt, align="center",
                                   valign="vcenter", border=1,
                                   border_color=th.border,
                                   num_format="#,##0")),
                sum(1 for w in (m.wishes if m else [])
                    if w[0] == (m.people[i]["name"] if i < len(m.people)
                                else "")) or "")
            bk.stats["formulas"] += 1

    K.list_dv(bk, KEY, "team", "teams")
    K.list_dv(bk, KEY, "diet", "diets")
    K.fixed_dv(bk, KEY, "rsvp", "RSVP")
    K.fixed_dv(bk, KEY, "status", "Statuses")
    K.status_cf(bk, KEY, "status", {
        C.ST_NOT: (th.bad_soft, th.bad),
        C.ST_BOUGHT: (th.warn_soft, th.warn),
        C.ST_WRAPPED: (th.info_soft, th.info),
        C.ST_DONE: (th.ok_soft, th.ok)})
    K.status_cf(bk, KEY, "rsvp", {
        "Yes": (th.ok_soft, th.ok),
        "No": (th.bad_soft, th.bad),
        "Maybe": (th.warn_soft, th.warn)})

    names = "$B$%d:$B$%d" % (C.ROW_FIRST, C.last_row(KEY))
    st = "$H$%d:$H$%d" % (C.ROW_FIRST, C.last_row(KEY))
    rv = "$E$%d:$E$%d" % (C.ROW_FIRST, C.last_row(KEY))
    chips = [
        ('="\U0001F465 In the draw: "&COUNTA(%s)' % names, "primary",
         "In the draw: %d" % (m.agg["participants"] if m else 0), 3),
        ('="\u2705 RSVP yes: "&COUNTIF(%s,"Yes")' % rv, "ok",
         "RSVP yes: %d" % (m.agg["rsvp_yes"] if m else 0), 3),
        ('="\u23F3 Awaiting: "&COUNTIF(%s,"Maybe")+COUNTIF(%s,"No")'
         % (rv, rv), "warn",
         "Awaiting: %d" % ((m.agg["rsvp_pending"] if m else 0)), 3),
        ('="\U0001F381 Ready (wrapped+done): "&COUNTIF(%s,"%s")+COUNTIF(%s,'
         '"%s")' % (st, C.ST_WRAPPED, st, C.ST_DONE), "gold",
         "Ready: %d" % (m.agg["wrapped"] if m else 0), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, C.last_row(KEY) + 2, LAST_COL)
```

## 14. Secret Santa Draw (seed derangement + rules validation + private lookup)

`santa_tracker/sheets/draw.py`

```python
"""🎅 Secret Santa Draw - seed-driven derangement with rules validation."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "draw"
LAST_COL = "G"

COLUMNS = [
    ("n", "#", "idx", None),
    ("giver", "Giver", "text", None),
    ("computed", "Drawn (auto)", "calc_c", None),
    ("override", "Manual override", "text", None),
    ("final", "FINAL recipient", "calc_c", None),
    ("flag", "Rules check", "calc_wrap", None),
    ("status", "Gift status", "calc_c", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    has_rules = bk.has("rules")
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F385  Secret Santa Draw",
                 "Change the seed on \u2699\ufe0f Settings to shuffle the "
                 "draw - nobody ever draws themselves.")
    K.table_frame(bk, KEY, COLUMNS, height=22)

    p = bk.q("participants")
    names = "%s!$B$%d:$B$%d" % (p, C.ROW_FIRST, C.last_row("participants"))
    hh = "%s!$D$%d:$D$%d" % (p, C.ROW_FIRST, C.last_row("participants"))
    tm = "%s!$C$%d:$C$%d" % (p, C.ROW_FIRST, C.last_row("participants"))
    ly = "%s!$G$%d:$G$%d" % (p, C.ROW_FIRST, C.last_row("participants"))
    pst = "%s!$H$%d:$H$%d" % (p, C.ROW_FIRST, C.last_row("participants"))
    xg = bk.rng("rules", "giver") if has_rules else None
    xc = bk.rng("rules", "cannot") if has_rules else None

    calc = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.primary,
                              bg_color=th.alt, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border))
    bold = bk.S.f(**bk.S.base(font_size=11, bold=True, font_color=th.ok,
                              bg_color=th.ok_soft, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border))
    flagf = bk.S.f(**bk.S.base(font_size=9.5, font_color=th.ink,
                               bg_color=th.card, align="center",
                               valign="vcenter", border=1,
                               border_color=th.border, text_wrap=True))
    statf = bk.S.f(**bk.S.base(font_size=10, font_color=th.ink,
                               bg_color=th.card, align="center",
                               valign="vcenter", border=1,
                               border_color=th.border))

    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        n = "COUNTA(%s)" % names
        inner = ('IF($E{r}=$B{r},"Self ","")'
                 + ('&IF(AND(CoupleRule="On",IFERROR(INDEX({hh},MATCH($B{r},'
                  '{nm},0)),"")<>"",INDEX({hh},MATCH($B{r},{nm},0))=INDEX('
                  '{hh},MATCH($E{r},{nm},0))),"Couple ","")'
                  + '&IF(AND(TeamRule="On",IFERROR(INDEX({tm},MATCH($B{r},'
                   '{nm},0)),"")<>"",INDEX({tm},MATCH($B{r},{nm},0))=INDEX('
                   '{tm},MATCH($E{r},{nm},0))),"Team ","")' if has_rules
                  else '')
                 + ('&IF(AND(LastYearRule="On",IFERROR(INDEX({ly},MATCH('
                    '$B{r},{nm},0)),"")=$E{r}),"LastYear ","")'
                   if has_rules else '')
                 + ('&IF(COUNTIFS({xg},$B{r},{xc},$E{r})>0,"Custom ","")'
                    if has_rules else '')).format(
            r=row, hh=hh, tm=tm, ly=ly, nm=names, xg=xg, xc=xc)
        values = {
            "n": i + 1,
            "giver": '=IF(%s!$B$%d="","",%s!$B$%d)'
                     % (p, row, p, row),
            "computed": '=IF($B{r}="","",INDEX({nm},MOD($A{r}-1+MOD('
                        'DrawSeed-1,MAX(1,{n}-1)),{n})+1))'.format(
                            r=row, nm=names, n=n),
            "override": "",
            "final": '=IF($B{r}="","",IF($D{r}<>"",$D{r},$C{r}))'.format(
                r=row),
            "flag": '=IF($E{r}="","",IF(TRIM({inner})="","'
                    '\u2705 OK",TRIM({inner})))'.format(r=row, inner=inner),
            "status": '=IF($B{r}="","",INDEX({st},MATCH($B{r},{nm},0)))'
                      .format(r=row, st=pst, nm=names),
        }
        cached = {}
        if m and i < len(m.people):
            cached = {"giver": m.people[i]["name"],
                      "computed": m.assign[i], "final": m.assign[i],
                      "flag": "\u2705 OK",
                      "status": m.people[i]["status"]}
        elif m:
            cached = {"giver": "", "computed": "", "final": "", "flag": "",
                      "status": ""}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)
        # restyle the computed/final/flag/status cells (kinds are generic)
        ws.write_formula(r(row), ci("C"), values["computed"], calc,
                         cached.get("computed", ""))
        ws.write_formula(r(row), ci("E"), values["final"], bold,
                         cached.get("final", ""))
        ws.write_formula(r(row), ci("F"), values["flag"], flagf,
                         cached.get("flag", ""))
        ws.write_formula(r(row), ci("G"), values["status"], statf,
                         cached.get("status", ""))
        bk.stats["formulas"] += 4

    K.list_dv(bk, KEY, "override", "participants",
              title="Manual override",
              message="Leave blank to keep the automatic draw.")
    K.status_cf(bk, KEY, "status", {
        C.ST_NOT: (th.bad_soft, th.bad),
        C.ST_BOUGHT: (th.warn_soft, th.warn),
        C.ST_WRAPPED: (th.info_soft, th.info),
        C.ST_DONE: (th.ok_soft, th.ok)})
    bk.cond(KEY, C.ROW_FIRST, ci("F"), C.last_row(KEY), ci("F"), {
        "type": "formula",
        "criteria": '=AND($F%d<>"",$F%d<>"\u2705 OK")'
                    % (C.ROW_FIRST, C.ROW_FIRST),
        "format": bk.S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})

    # ------------------------------------------------------------------
    row = C.last_row(KEY) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  \U0001FDD1  DRAW CONTROLS", bk.S.section_soft)
    ws.set_row(r(row), 22)
    lab = bk.S.f(**bk.S.base(font_size=10.5, bold=True, font_color=th.ink,
                             bg_color=th.card, align="left", indent=1,
                             valign="vcenter", border=1,
                             border_color=th.border))
    seedf = bk.S.f(**bk.S.base(font_size=12, bold=True,
                               font_color=th.primary, bg_color=th.gold_soft,
                               border=1, border_color=th.border_strong,
                               align="center", valign="vcenter",
                               locked=False, num_format="#,##0"))
    stat = bk.S.f(**bk.S.base(font_size=10.5, bold=True,
                              font_color=th.primary, bg_color=th.gold_soft,
                              border=1, border_color=th.border_strong,
                              align="center", valign="vcenter"))
    ws.set_row(r(row + 1), 24)
    ws.merge_range(r(row + 1), 1, r(row + 1), 2, "  Draw seed (change to "
                                                "re-draw)", lab)
    ws.write_formula(r(row + 1), 3, "=DrawSeed", seedf,
                     m.settings["seed"] if m else "")
    bk.stats["formulas"] += 1
    ws.merge_range(r(row + 1), 4, r(row + 1), 5, "  Draw status", lab)
    ws.write_formula(r(row + 1), 6, "=DrawStatus", stat,
                     m.settings["status"] if m else "")
    bk.stats["formulas"] += 1

    row += 3
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  \U0001F575\ufe0f  PRIVATE LOOKUP - type a name, see "
                   "only their recipient", bk.S.section_soft)
    ws.set_row(r(row), 22)
    ws.set_row(r(row + 1), 26)
    inp = bk.S.f(**bk.S.base(font_size=11, bold=True, font_color=th.ink,
                             bg_color=th.card, border=1,
                             border_color=th.border_strong, align="left",
                             valign="vcenter", indent=1, locked=False))
    out = bk.S.f(**bk.S.base(font_size=11.5, bold=True,
                             font_color=th.primary, bg_color=th.card,
                             align="left", valign="vcenter", indent=1))
    ws.merge_range(r(row + 1), 1, r(row + 1), 2, "  I am...", lab)
    ws.merge_range(r(row + 1), 3, r(row + 1), 4, "", inp)
    bk.validate(KEY, row + 1, 3, row + 1, 3, "=" + bk.listname("participants"),
                title="Your name")
    ws.merge_range(r(row + 1), 5, r(row + 1), ci(LAST_COL), "", out)
    ws.write_formula(
        r(row + 1), 5,
        '=IF($C%d="","Type your name to peek at your recipient\u2026",'
        'IFERROR("\U0001F385  You are drawing:  "&INDEX($E$%d:$E$%d,MATCH('
        '$C%d,$B$%d:$B$%d,0))&"   \u2022   budget "&Currency&TEXT(BudgetMin,'
        '"#,##0")&"\u2013"&Currency&TEXT(BudgetMax,"#,##0")&"   \u2022   "'
        '&TEXT(PartyDate,"dd mmm yyyy")&"  \u2014  tell no one!",'
        '"Name not found - check the spelling on \U0001F465 Participants."))'
        % (row + 1, C.ROW_FIRST, C.last_row(KEY), row + 1, C.ROW_FIRST,
           C.last_row(KEY)),
        out, "")
    bk.stats["formulas"] += 1

    chips = [
        ('="\U0001F385 Assigned: "&SUMPRODUCT(--($E$%d:$E$%d<>""))&" of "&'
         'COUNTA(%s!$B$%d:$B$%d)'
         % (C.ROW_FIRST, C.last_row(KEY), p, C.ROW_FIRST,
            C.last_row("participants")), "primary",
         "Assigned: %d of %d" % ((m.agg["assigned"], m.agg["participants"])
                                 if m else (0, 0)), 4),
        ('="\u2705 Clean pairs: "&COUNTIF($F$%d:$F$%d,"\u2705 OK")'
         % (C.ROW_FIRST, C.last_row(KEY)), "ok",
         "Clean pairs: %d" % (m.agg["participants"] if m else 0), 3),
        ('="\u26A0\ufe0f Rule hits: "&%s' % bk.kpi("violations"), "bad",
         "Rule hits: %d" % (m.agg["violations"] if m else 0), 3),
        ('="\U0001F512 Status: "&DrawStatus', "gold",
         "Status: %s" % (m.settings["status"] if m else C.DRAW_DRAFT), 4),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, row + 3, LAST_COL, tip=
                 "  \U0001F4A1  Re-draw: unlock the sheet (password "
                 "\u201cpremium\u201d), change the seed, re-protect.  "
                 "Overrides beat the auto-draw pair by pair.")
```

## 15. Exclusions & Rules

`santa_tracker/sheets/rules.py`

```python
"""🚫 Exclusions & Rules - toggles plus custom no-pair rules."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "rules"
LAST_COL = "D"

COLUMNS = [
    ("n", "#", "idx", None),
    ("giver", "This person\u2026", "text", None),
    ("cannot", "\u2026must never draw", "text", None),
    ("reason", "Why", "wrap", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F6AB  Exclusions & Rules",
                 "Switch the house rules on or off and list custom no-pairs "
                 "- the draw validates itself against all of them.")
    bk.paint(KEY, 0, 0, C.ROW_FIRST + C.CAP[KEY] + 12, ci(LAST_COL),
             bk.S.canvas)

    lab = bk.S.f(**bk.S.base(font_size=10.5, bold=True, font_color=th.ink,
                             bg_color=th.card, align="left", indent=1,
                             valign="vcenter", border=1,
                             border_color=th.border))
    tog = bk.S.f(**bk.S.base(font_size=11, bold=True, font_color=th.primary,
                             bg_color=th.gold_soft, border=1,
                             border_color=th.border_strong, align="center",
                             valign="vcenter", locked=False))
    ws.merge_range(r(7), 1, r(7), ci(LAST_COL),
                   "  \U0001F39B\ufe0f  HOUSE RULES", bk.S.section_soft)
    ws.set_row(r(7), 22)
    toggles = [
        (8, "Couples / households cannot draw each other", "couple"),
        (9, "Same team / department cannot draw each other", "team"),
        (10, "No repeat of last year's pairing", "lastyear"),
    ]
    for row, text, key in toggles:
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), 2, "  " + text, lab)
        ws.write(r(row), 3, m.rules[key] if m else "On", tog)
    ws.merge_range(r(11), 1, r(11), ci(LAST_COL),
                   "  \u26D4  (Self-draws are always impossible - the draw "
                   "is a closed cycle.)", bk.S.sheet_sub)

    head = 13
    ws.merge_range(r(head), 1, r(head), ci(LAST_COL),
                   "  \u270D\ufe0f  CUSTOM NO-PAIR RULES", bk.S.section_soft)
    ws.set_row(r(head), 22)
    # shift the table down: header row 14, data from 15
    K.header_row(bk, KEY, COLUMNS, row=head + 1, height=24)
    first = head + 2
    last = first + C.CAP[KEY] - 1
    from .common import data_rows
    ws = bk.ws(KEY)
    for i in range(C.CAP[KEY]):
        ws.set_row(r(first + i), 20)
    if m:
        for i, (a, b) in enumerate(m.pairs):
            K.write_row(bk, KEY, COLUMNS, first + i,
                        {"n": i + 1, "giver": a, "cannot": b,
                         "reason": m.pair_reason[(a, b)]})
    for i in range(len(m.pairs) if m else 0, C.CAP[KEY]):
        K.write_row(bk, KEY, COLUMNS, first + i,
                    {"n": i + 1, "giver": "", "cannot": "", "reason": ""})
    K.list_dv(bk, KEY, "giver", "participants", first=first, last=last)
    K.list_dv(bk, KEY, "cannot", "participants", first=first, last=last)

    row = last + 2
    chips = [
        ('="\U0001F6AB Rules armed: "&%s' % bk.kpi("rules_count"),
         "primary", "Rules armed: %d" % (m.agg["rules_count"] if m else 0),
         4),
        ('="\u26A0\ufe0f Draw violations: "&%s' % bk.kpi("violations"),
         "bad", "Draw violations: %d" % (m.agg["violations"] if m else 0),
         4),
        ('="\u270D\ufe0f Custom pairs: "&SUMPRODUCT(--($B$%d:$B$%d<>""))'
         % (first, last), "gold",
         "Custom pairs: %d" % (len(m.pairs) if m else 0), 4),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.note_block(bk, KEY, row + 5, "B", "D", [
        "Every rule you arm here is checked live on the \U0001F385 Secret "
        "Santa Draw tab - any offending pair shows exactly which rule it "
        "breaks (Self / Couple / Team / LastYear / Custom).",
        "Fix a flagged pair with a manual override on the Draw tab, or "
        "re-draw with a new seed until the board is clean.",
    ], title="\U0001F4D0  HOW VALIDATION WORKS")
    K.footer_nav(bk, KEY, row + 9, LAST_COL)
```

## 16. Budget Tracker

`santa_tracker/sheets/budget.py`

```python
"""💰 Budget Tracker - min/max rules, actual spend, receipts."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "budget"
LAST_COL = "I"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Gift for (recipient)", "text", None),
    ("min", "Min", "calc_money", None),
    ("max", "Max", "calc_money", None),
    ("spent", "Actual spent", "money", None),
    ("flag", "Budget check", "calc_wrap", None),
    ("receipt", "Receipt?", "tick", None),
    ("ref", "Receipt / order ref", "center", None),
    ("notes", "Notes", "wrap", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F4B0  Budget Tracker",
                 "One row per giver: what they spent on their Secret Santa "
                 "gift, checked against the party's min and max.")
    K.table_frame(bk, KEY, COLUMNS, height=20)
    flagf = bk.S.f(**bk.S.base(font_size=9.5, font_color=th.ink,
                               bg_color=th.alt, align="center",
                               valign="vcenter", border=1,
                               border_color=th.border, text_wrap=True))
    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "name": '=IF(%s!$C$%d="","",%s!$C$%d)'
                    % (bk.q("draw"), row, bk.q("draw"), row),
            "min": "=BudgetMin",
            "max": "=BudgetMax",
            "spent": "",
            "flag": '=IF($E{r}="","",IF($E{r}<$C{r},"%s",IF($E{r}>$D{r},"%s",'
                    '"%s")))' % (C.BF_UNDER, C.BF_OVER, C.BF_OK),
            "receipt": "",
            "ref": "",
            "notes": "",
        }
        cached = {}
        if m and i < len(m.people):
            cached = {"name": m.assign[i], "min": m.settings["bmin"],
                      "max": m.settings["bmax"],
                      "flag": (C.BF_UNDER if m.spent[i] < m.settings["bmin"]
                               else C.BF_OVER if m.spent[i]
                               > m.settings["bmax"] else C.BF_OK)}
            values["spent"] = m.spent[i] or ""
            values["receipt"] = m.receipts[i]
            values["ref"] = m.refs[i]
        elif m:
            cached = {"name": "", "min": m.settings["bmin"],
                      "max": m.settings["bmax"], "flag": ""}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)
        ws.write_formula(r(row), ci("F"),
                         values["flag"].format(r=row), flagf,
                         cached.get("flag", ""))
        bk.stats["formulas"] += 1

    K.money_dv(bk, KEY, ("spent",))
    K.tick_dv(bk, KEY, ("receipt",))
    K.tick_cf(bk, KEY, ("receipt",))
    bk.cond(KEY, C.ROW_FIRST, ci("F"), C.last_row(KEY), ci("F"), {
        "type": "formula",
        "criteria": '=$F%d="%s"' % (C.ROW_FIRST, C.BF_OVER),
        "format": bk.S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.cond(KEY, C.ROW_FIRST, ci("F"), C.last_row(KEY), ci("F"), {
        "type": "formula",
        "criteria": '=$F%d="%s"' % (C.ROW_FIRST, C.BF_UNDER),
        "format": bk.S.cf(bg=th.warn_soft, fg=th.warn, bold=True)})
    bk.cond(KEY, C.ROW_FIRST, ci("F"), C.last_row(KEY), ci("F"), {
        "type": "formula",
        "criteria": '=$F%d="%s"' % (C.ROW_FIRST, C.BF_OK),
        "format": bk.S.cf(bg=th.ok_soft, fg=th.ok)})

    tot = C.last_row(KEY) + 1
    K.totals_row(bk, KEY, tot, {
        "spent": ("=SUM(E%d:E%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", m.agg["spent"] if m else 0)},
        first_col="B", last_col=LAST_COL)

    sp = "$E$%d:$E$%d" % (C.ROW_FIRST, C.last_row(KEY))
    fl = "$F$%d:$F$%d" % (C.ROW_FIRST, C.last_row(KEY))
    chips = [
        ('="\U0001F4B8 Total spent: "&Currency&TEXT(SUM(%s),"#,##0.00")'
         % sp, "primary",
         "Total spent: %s" % (m.money(m.agg["spent"]) if m else "$0.00"), 4),
        ('="\U0001F4CA Average gift: "&Currency&TEXT(%s,"#,##0.00")'
         % bk.kpi("avg_gift"), "info",
         "Average gift: %s" % (m.money(m.agg["avg_gift"]) if m else "$0.00"),
         4),
        ('="\U0001F6A8 Over budget: "&COUNTIF(%s,"%s")' % (fl, C.BF_OVER),
         "bad", "Over budget: %d" % (m.agg["over_budget"] if m else 0), 3),
        ('="\U0001F9FE Receipts filed: "&COUNTIF($G$%d:$G$%d,"%s")'
         % (C.ROW_FIRST, C.last_row(KEY), C.TICK), "ok",
         "Receipts filed: %d" % (sum(1 for x in (m.receipts if m else [])
                                    if x) ), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, tot + 2, LAST_COL)
```

## 17. Wishlists

`santa_tracker/sheets/wishlists.py`

```python
"""🎁 Wishlists - ideas per participant, claimable by their Santa."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "wishlists"
LAST_COL = "G"

COLUMNS = [
    ("n", "#", "idx", None),
    ("who", "Participant", "text", None),
    ("item", "Wish", "wrap", None),
    ("priority", "Priority", "center", None),
    ("link", "Link / shop", "link", None),
    ("claimed", "Claimed", "tick", None),
    ("notes", "Notes", "wrap", None),
]


def build(bk):
    th = bk.th
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F381  Wishlists",
                 "Ideas each participant dropped - tick Claimed once their "
                 "Santa commits (no names shown, no spoilers).")
    K.table_frame(bk, KEY, COLUMNS, height=20)
    if m:
        for i, w in enumerate(m.wishes):
            K.write_row(bk, KEY, COLUMNS, C.ROW_FIRST + i,
                        {"n": i + 1, "who": w[0], "item": w[1],
                         "priority": w[2], "link": w[3], "claimed": w[4],
                         "notes": w[5]})
    for i in range(len(m.wishes) if m else 0, C.CAP[KEY]):
        K.write_row(bk, KEY, COLUMNS, C.ROW_FIRST + i,
                    {"n": i + 1, "who": "", "item": "", "priority": "",
                     "link": "", "claimed": "", "notes": ""})
    K.list_dv(bk, KEY, "who", "participants")
    K.fixed_dv(bk, KEY, "priority", "Priorities")
    K.tick_dv(bk, KEY, ("claimed",))
    K.tick_cf(bk, KEY, ("claimed",))
    K.status_cf(bk, KEY, "priority", {
        "Must-love": (th.accent_soft, th.accent),
        "Nice to have": (th.info_soft, th.info),
        "Just an idea": (th.card, th.muted)})

    cl = "$F$%d:$F$%d" % (C.ROW_FIRST, C.last_row(KEY))
    pr = "$D$%d:$D$%d" % (C.ROW_FIRST, C.last_row(KEY))
    chips = [
        ('="\U0001F381 Wishes logged: "&SUMPRODUCT(--($C$%d:$C$%d<>""))'
         % (C.ROW_FIRST, C.last_row(KEY)), "primary",
         "Wishes logged: %d" % (len(m.wishes) if m else 0), 4),
        ('="\u2705 Claimed: "&COUNTIF(%s,"%s")' % (cl, C.TICK), "ok",
         "Claimed: %d" % (sum(1 for w in (m.wishes if m else []) if w[4]),),
         4),
        ('="\u2764\ufe0f Must-love open: "&SUMPRODUCT((%s="Must-love")*'
         '(%s<>"%s"))' % (pr, cl, C.TICK), "bad",
         "Must-love open: %d" % (sum(1 for w in (m.wishes if m else [])
                                     if w[2] == "Must-love" and not w[4]),),
         4),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, C.last_row(KEY) + 2, LAST_COL)
```

## 18. White Elephant engine

`santa_tracker/sheets/we.py`

```python
"""🎲 White Elephant - seat order, gifts, steals, locks and holders."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "we"
LAST_COL = "L"

COLUMNS = [
    ("n", "#", "idx", None),
    ("order", "Seat", "calc_num", None),
    ("player", "Player", "text", None),
    ("giftnum", "Gift #", "calc_num", None),
    ("desc", "Gift description", "wrap", None),
    ("value", "Gift value", "money", None),
    ("steals", "Steals", "calc_num", None),
    ("maxsteals", "Max", "calc_num", None),
    ("status", "Gift status", "calc_c", None),
    ("holder", "Current holder", "calc_c", None),
    ("stolenfrom", "Last stolen from", "calc_c", None),
    ("final", "Final?", "calc_tick", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F3B2  White Elephant",
                 "Seats and gift numbers shuffle from the seed; holders, "
                 "steal counts and locks follow \U0001F504 Game History.")
    K.table_frame(bk, KEY, COLUMNS, height=22)
    ws.set_column(ci("M"), ci("N"), 10, None, {"hidden": True})

    h = bk.q("history")
    hgift = "%s!$E$%d:$E$%d" % (h, C.ROW_FIRST, C.last_row("history"))
    hturn = "%s!$G$%d:$G$%d" % (h, C.ROW_FIRST, C.last_row("history"))
    hplayer = "%s!$C$%d:$C$%d" % (h, C.ROW_FIRST, C.last_row("history"))
    haction = "%s!$D$%d:$D$%d" % (h, C.ROW_FIRST, C.last_row("history"))
    hfrom = "%s!$F$%d:$F$%d" % (h, C.ROW_FIRST, C.last_row("history"))

    statf = bk.S.f(**bk.S.base(font_size=10, bold=True, font_color=th.ink,
                               bg_color=th.alt, align="center",
                               valign="vcenter", border=1,
                               border_color=th.border))
    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        lastt = 'SUMPRODUCT(MAX((%s=$D%d)*%s))' % (hgift, row, hturn)
        lastidx = ('SUMPRODUCT((%s=$D%d)*(%s=%s)*(ROW(%s)-%d))'
                   % (hgift, row, hturn, lastt, hgift,
                      C.ROW_FIRST - 1))
        values = {
            "n": i + 1,
            "order": '=IF($C%d="","",1+SUMPRODUCT(--($M$%d:$M$%d>$M%d)))'
                     % (row, C.ROW_FIRST, C.last_row(KEY), row),
            "player": "",
            "giftnum": '=IF($C%d="","",1+SUMPRODUCT(--($N$%d:$N$%d>$N%d)))'
                       % (row, C.ROW_FIRST, C.last_row(KEY), row),
            "desc": "",
            "value": "",
            "steals": '=IF($D%d="","",COUNTIFS(%s,$D%d,%s,"Steal"))'
                      % (row, hgift, row, haction),
            "maxsteals": '=IF($C%d="","",MaxSteals)' % row,
            "status": '=IF($D%d="","",IF($G%d>=$H%d,"%s",IF(COUNTIFS(%s,'
                      '$D%d)=0,"%s",IF(INDEX(%s,%s)="Steal","%s","%s"))))'
                      % (row, row, row, C.GS_FINAL, hgift, row, C.GS_AVAIL,
                         haction, lastidx, C.GS_STOLEN, C.GS_HELD),
            "holder": '=IF($D%d="","",IF(COUNTIFS(%s,$D%d)=0,"",INDEX('
                      '%s,%s)))' % (row, hgift, row, hplayer, lastidx),
            "stolenfrom": '=IF($I%d="%s",IFERROR(INDEX(%s,%s),""),"")'
                          % (row, C.GS_STOLEN, hfrom, lastidx),
            "final": '=IF($C%d="","",IF(COUNTA($J$%d:$J$%d)=COUNTA($C$%d:'
                     '$C$%d),IFERROR(INDEX($D$%d:$D$%d,MATCH($C%d,$J$%d:$J$%d,'
                     '0)),""),"-"))'
                     % (row, C.ROW_FIRST, C.last_row(KEY), C.ROW_FIRST,
                        C.last_row(KEY), C.ROW_FIRST, C.last_row(KEY), row,
                        C.ROW_FIRST, C.last_row(KEY)),
            "key1": '=IF($C%d="",0,MOD(WESeed*7919*(ROW()-7)+(ROW()-8)*'
                    '104729,999983)*100+(ROW()-8))' % row,
            "key2": '=IF($C%d="",0,MOD(WESeed*6967*(ROW()-3)+(ROW()-8)*'
                    '1299709,999983)*100+(ROW()-8))' % row,
        }
        cached = {}
        if m and i < len(m.we):
            w = m.we[i]
            cached = {"order": w["seat"], "giftnum": w["giftnum"],
                      "steals": w["steals"], "maxsteals": m.settings["steals"],
                      "status": w["status"], "holder": w["holder"],
                      "stolenfrom": w["stolenfrom"], "final": "",
                      "key1": 0, "key2": 0}
            values["player"] = w["player"]
            values["desc"] = w["desc"]
            values["value"] = w["value"]
        elif m:
            cached = {"order": "", "giftnum": "", "steals": "",
                      "maxsteals": "", "status": "", "holder": "",
                      "stolenfrom": "", "final": "", "key1": "", "key2": ""}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)
        ws.write_formula(r(row), ci("I"), values["status"], statf,
                         cached.get("status", ""))
        bk.stats["formulas"] += 1

    K.list_dv(bk, KEY, "player", "participants")
    K.money_dv(bk, KEY, ("value",))
    K.status_cf(bk, KEY, "status", {
        C.GS_AVAIL: (th.card, th.muted),
        C.GS_HELD: (th.ok_soft, th.ok),
        C.GS_STOLEN: (th.warn_soft, th.warn),
        C.GS_FINAL: (th.bad_soft, th.bad)})
    # highlight the seat whose turn it is
    bk.cond(KEY, C.ROW_FIRST, 1, C.last_row(KEY), ci(LAST_COL), {
        "type": "formula",
        "criteria": '=AND($C%d<>"",$B%d=%s)'
                    % (C.ROW_FIRST, C.ROW_FIRST, bk.kpi("current_turn")),
        "format": bk.S.cf(bg=th.gold_soft, fg=th.primary, bold=True)})

    chips = [
        ('="\U0001F3B2 Players: "&COUNTA($C$%d:$C$%d)'
         % (C.ROW_FIRST, C.last_row(KEY)), "primary",
         "Players: %d" % (m.agg["we_players"] if m else 0), 3),
        ('="\U0001F381 In play: "&%s' % bk.kpi("gifts_in_play"), "ok",
         "In play: %d" % (m.agg["gifts_in_play"] if m else 0), 3),
        ('="\U0001F512 Locked: "&%s' % bk.kpi("locked_gifts"), "bad",
         "Locked: %d" % (m.agg["locked_gifts"] if m else 0), 3),
        ('="\U0001F449 Now: seat "&%s&" ("&IFERROR(INDEX($C$%d:$C$%d,MATCH('
         '%s,$B$%d:$B$%d,0)),"game over")&")"'
         % (bk.kpi("current_turn"), C.ROW_FIRST, C.last_row(KEY),
            bk.kpi("current_turn"), C.ROW_FIRST, C.last_row(KEY)), "gold",
         "Now: seat %d (%s)" % ((m.agg["current_turn"],
                                 m.we_players[m.agg["current_turn"] - 1])
                                if m else (0, "")), 5),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, C.last_row(KEY) + 2, LAST_COL, tip=
                 "  \U0001F4A1  House rule: one turn per seat per round; if "
                 "your gift was stolen, pick again on your next round turn. "
                 " A gift locks at max steals.")
```

## 19. Game History log

`santa_tracker/sheets/history.py`

```python
"""🔄 Game History - the white elephant turn log that drives everything."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "history"
LAST_COL = "F"

COLUMNS = [
    ("n", "#", "idx", None),
    ("turn", "Turn", "num", None),
    ("player", "Player", "text", None),
    ("action", "Action", "center", None),
    ("gift", "Gift #", "num", None),
    ("stolenfrom", "Stolen from", "center", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F504  Game History",
                 "Log every Pick, Steal and Pass in order - the White "
                 "Elephant tab reads this log live.")
    K.table_frame(bk, KEY, COLUMNS, height=20)
    if m:
        for i, (turn, pl, action, gnum, frm) in enumerate(m.history):
            K.write_row(bk, KEY, COLUMNS, C.ROW_FIRST + i,
                        {"n": i + 1, "turn": turn, "player": pl,
                         "action": action, "gift": gnum,
                         "stolenfrom": frm})
    for i in range(len(m.history) if m else 0, C.CAP[KEY]):
        K.write_row(bk, KEY, COLUMNS, C.ROW_FIRST + i,
                    {"n": i + 1, "turn": "", "player": "", "action": "",
                     "gift": "", "stolenfrom": ""})
    ws.set_column(ci("G"), ci("G"), 8, None, {"hidden": True})
    hf = bk.S.f(**bk.S.base(font_size=9, font_color=bk.th.muted,
                            align="right", valign="vcenter",
                            num_format="#,##0"))
    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        ws.write_formula(r(row), ci("G"), '=IF($B%d="",0,$B%d)' % (row, row),
                         hf, (m.history[i][0] if m and i < len(m.history)
                              else 0))
        bk.stats["formulas"] += 1
    K.whole_dv(bk, KEY, ("turn",), minimum=1, maximum=200)
    K.list_dv(bk, KEY, "player", "participants")
    K.fixed_dv(bk, KEY, "action", "Actions")
    K.list_dv(bk, KEY, "gift", "giftnums")
    K.list_dv(bk, KEY, "stolenfrom", "participants")
    K.status_cf(bk, KEY, "action", {
        "Pick": (th.ok_soft, th.ok),
        "Steal": (th.bad_soft, th.bad),
        "Pass": (th.card, th.muted)})

    ac = "$D$%d:$D$%d" % (C.ROW_FIRST, C.last_row(KEY))
    chips = [
        ('="\U0001F504 Turns logged: "&COUNT($B$%d:$B$%d)'
         % (C.ROW_FIRST, C.last_row(KEY)), "primary",
         "Turns logged: %d" % (m.agg["turns_done"] if m else 0), 4),
        ('="\U0001F932 Steals: "&COUNTIF(%s,"Steal")' % ac, "bad",
         "Steals: %d" % (sum(1 for h in (m.history if m else [])
                             if h[2] == "Steal"),), 4),
        ('="\U0001F381 Picks: "&COUNTIF(%s,"Pick")' % ac, "ok",
         "Picks: %d" % (sum(1 for h in (m.history if m else [])
                            if h[2] == "Pick"),), 3),
        ('="\u23ED\ufe0f Passes: "&COUNTIF(%s,"Pass")' % ac, "info",
         "Passes: %d" % (sum(1 for h in (m.history if m else [])
                             if h[2] == "Pass"),), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.note_block(bk, KEY, C.last_row(KEY) + 3, "B", "F", [
        "Turn numbers run 1, 2, 3\u2026 in seat order each round; extra "
        "rounds continue the count.",
        "On a Steal row, always say who the gift was stolen from - the "
        "\u201clast stolen from\u201d column on the game tab reads it.",
        "A gift reaches \U0001F512 Final automatically once its steal count "
        "hits the max set on \u2699\ufe0f Settings.",
    ], title="\U0001F4D6  LOGGING THE GAME")
    K.footer_nav(bk, KEY, C.last_row(KEY) + 8, LAST_COL)
```

## 20. Printable Santa Cards

`santa_tracker/sheets/cards.py`

```python
"""🎟️ Santa Cards - printable cut-out cards, one per participant."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "cards"
LAST_COL = "I"


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F39F\ufe0f  Santa Cards",
                 "Print this tab, cut along the cards, hand one to each "
                 "guest - their recipient stays secret until then.")
    bk.paint(KEY, 0, 0, 46, ci(LAST_COL), bk.S.canvas)

    p = bk.q("participants")
    d = bk.q("draw")
    names = "%s!$B$%d:$B$%d" % (p, C.ROW_FIRST, C.last_row("participants"))
    finals = "%s!$E$%d:$E$%d" % (d, C.ROW_FIRST, C.last_row("draw"))

    frame = bk.S.f(**bk.S.base(bg_color=th.card, border=1,
                               border_color=th.border_strong))
    head = bk.S.f(**bk.S.base(font_size=10, bold=True, font_color=th.white,
                              bg_color=th.primary, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border_strong))
    giver = bk.S.f(**bk.S.base(font_size=15, bold=True,
                               font_color=th.primary, bg_color=th.card,
                               align="center", valign="vcenter", border=1,
                               border_color=th.border_strong))
    lab = bk.S.f(**bk.S.base(font_size=9, italic=True, font_color=th.muted,
                             bg_color=th.card, align="center",
                             valign="vcenter", border=1,
                             border_color=th.border_strong))
    rec = bk.S.f(**bk.S.base(font_size=14, bold=True, font_color=th.accent,
                             bg_color=th.accent_soft, align="center",
                             valign="vcenter", border=1,
                             border_color=th.border_strong))
    meta = bk.S.f(**bk.S.base(font_size=9, font_color=th.ink,
                              bg_color=th.card, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border_strong))

    for k in range(1, 9):
        col0 = 1 if k % 2 == 1 else 6
        row0 = 8 + ((k - 1) // 2) * 8
        ws.set_row(r(row0), 20)
        ws.set_row(r(row0 + 1), 26)
        ws.set_row(r(row0 + 2), 14)
        ws.set_row(r(row0 + 3), 24)
        ws.set_row(r(row0 + 4), 14)
        ws.set_row(r(row0 + 5), 14)
        ws.merge_range(r(row0), col0, r(row0), col0 + 3,
                       "  \U0001F385  SECRET SANTA CARD  \u2022  %d" % k,
                       head)
        ws.merge_range(r(row0 + 1), col0, r(row0 + 1), col0 + 3, "", giver)
        ws.write_formula(r(row0 + 1), col0,
                         '=IFERROR(INDEX(%s,%d),"")' % (names, k), giver,
                         (m.people[k - 1]["name"]
                          if m and k <= len(m.people) else ""))
        bk.stats["formulas"] += 1
        ws.merge_range(r(row0 + 2), col0, r(row0 + 2), col0 + 3,
                       "you are secretly drawing\u2026", lab)
        ws.merge_range(r(row0 + 3), col0, r(row0 + 3), col0 + 3, "", rec)
        ws.write_formula(r(row0 + 3), col0,
                         '=IFERROR(INDEX(%s,%d),"")' % (finals, k), rec,
                         (m.assign[k - 1] if m and k <= len(m.people)
                          else ""))
        bk.stats["formulas"] += 1
        ws.merge_range(r(row0 + 4), col0, r(row0 + 4), col0 + 3, "", meta)
        ws.write_formula(
            r(row0 + 4), col0,
            '=IF(COUNTA(%s)=0,"","budget "&Currency&TEXT('
            'BudgetMin,"#,##0")&"\u2013"&Currency&TEXT(BudgetMax,'
            '"#,##0")&"   \u2022   bring it by "&TEXT(PartyDate,'
            '"dd mmm"))' % names, meta,
            ("budget $20\u2013$35   \u2022   bring it by 19 Dec"
             if m else ""))
        bk.stats["formulas"] += 1
        ws.merge_range(r(row0 + 5), col0, r(row0 + 5), col0 + 3, "", meta)
        ws.write_formula(
            r(row0 + 5), col0,
            '=IF(COUNTA(%s)=0,"",PartyName&"  \u2022  "&PartyLocation)'
            % names, meta,
            ("%s  \u2022  %s" % (m.settings["party"],
                                  m.settings["location"]) if m else ""))
        bk.stats["formulas"] += 1
    K.footer_nav(bk, KEY, 42, LAST_COL, landscape=True, tip=
                 "  \u2702\ufe0f  Print landscape, cut along the borders, "
                 "fold once - secret inside.")
```

## 21. Start Here guide

`santa_tracker/sheets/guide.py`

```python
"""📖 Start Here - quick start, tab tour, tips and the friendly rules."""

import os

from ..book import r, ci

LAST_COL = "J"


def build(bk):
    key = "guide"
    ws = bk.ws(key)
    S, th = bk.S, bk.th

    bk.widths(key, {"A": 2.2, "B": 4, "C": 15, "D": 15, "E": 15, "F": 15,
                    "G": 15, "H": 15, "I": 15, "J": 15})
    bk.paint(key, 0, 0, 90, ci(LAST_COL), S.canvas)

    body = S.f(**S.base(font_size=10.5, bg_color=th.bg, align="left",
                        valign="vcenter", text_wrap=True))
    body_soft = S.f(**S.base(font_size=10.5, bg_color=th.card, align="left",
                             valign="vcenter", text_wrap=True,
                             border=1, border_color=th.border))
    badge = S.f(**S.base(bold=True, font_size=12, font_color=th.white,
                         bg_color=th.accent, align="center", valign="vcenter"))
    badge2 = S.f(**S.base(bold=True, font_size=12, font_color=th.white,
                          bg_color=th.info, align="center", valign="vcenter"))
    step_title = S.f(**S.base(bold=True, font_size=11.5,
                              font_color=th.primary, bg_color=th.bg,
                              align="left", valign="vcenter", indent=1))
    mark = S.f(**S.base(font_size=10, font_color=th.muted, bg_color=th.bg,
                        align="center", valign="vcenter"))

    def sec(row, emoji, text):
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                       "  %s  %s" % (emoji, text), S.section)

    # ------------------------------------------------------------------
    # cover banner + overlaid title (plain title band when no artwork)
    # ------------------------------------------------------------------
    image = _banner_path(bk)
    row = 1
    ws.set_row(r(row), 7)
    row += 1
    if image:
        w, h = _png_size(image)
        scale = 860.0 / float(w)
        span = int(h * scale / 20.0) + 1
        cover = S.f(**S.base(bg_color=th.cover_bg))
        for i in range(span):
            ws.set_row(r(row + i), 20)
            for col in range(1, ci(LAST_COL) + 1):
                ws.write(r(row + i), col, "", cover)
        ws.insert_image(r(row), 1, image,
                        {"x_scale": scale, "y_scale": scale,
                         "object_position": 1})
        t_row = row + max(1, span // 2 - 1)
        ws.merge_range(r(t_row), ci("C"), r(t_row), ci("G"),
                       "\U0001F4D6  START HERE",
                       S.f(**S.base(font_name=th.title_font, font_size=26,
                                    bold=True, font_color=th.primary,
                                    bg_color=th.cover_bg, align="center",
                                    valign="vcenter")))
        ws.merge_range(r(t_row + 1), ci("C"), r(t_row + 1), ci("G"),
                       "  your 5-minute tour of the Secret Santa & White "
                       "Elephant Party Tracker",
                       S.f(**S.base(font_size=11.5, italic=True,
                                    font_color=th.muted,
                                    bg_color=th.cover_bg, align="center",
                                    valign="vcenter")))
        row += span + 1
        ws.set_row(r(row), 18)
        ws.merge_range(r(row), ci(LAST_COL) - 2, r(row), ci(LAST_COL), "",
                       S.home_link)
        ws.write_url(r(row), ci(LAST_COL) - 2,
                     "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                     "\U0001F3E0  Back to Dashboard")
        row += 2
    else:
        ws.set_row(r(row), 32)
        ws.set_row(r(row + 1), 18)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL) - 3,
                       "\U0001F4D6  Start Here \u2014 your 5-minute setup",
                       S.sheet_title)
        ws.merge_range(r(row + 1), 1, r(row + 1), ci(LAST_COL) - 3,
                       "  Everything is wired together: log a sale once "
                       "and the fairs, reorder list, monthly summary and "
                       "dashboard follow.", S.sheet_sub)
        ws.merge_range(r(row + 1), ci(LAST_COL) - 2, r(row + 1), ci(LAST_COL),
                       "", S.home_link)
        ws.write_url(r(row + 1), ci(LAST_COL) - 2,
                     "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                     "\U0001F3E0  Back to Dashboard")
        row += 3
    # ------------------------------------------------------------------
    sec(row, "\U0001F680", "QUICK START \u2014 FIVE STEPS TO LIVE")
    row += 1
    steps = [
        ("1", "Open \u2699\ufe0f Settings & Instructions",
         "Type the party name, date, location, host, currency, gift budget "
         "min/max, max steals and the two seeds.  Every tab reads these."),
        ("2", "List your guests on \U0001F465 Participants",
         "Names plus team, household/couple, RSVP, dietary needs and who "
         "they drew last year - the exclusion engine feeds off this."),
        ("3", "Draw on \U0001F385 Secret Santa Draw",
         "The draw is a closed cycle: nobody can ever pick themselves.  "
         "Change the seed to shuffle, override any pair by hand, then lock "
         "the draw with the status dropdown."),
        ("4", "Arm your rules on \U0001F6AB Exclusions & Rules",
         "Couples, same-team and last-year repeats toggle on/off; custom "
         "no-pairs go in the table.  The draw flags every breach live."),
        ("5", "Run the party",
         "\U0001F4B0 Budget Tracker watches every spend, \U0001F381 "
         "Wishlists keeps ideas claimable, and \U0001F3B2 White Elephant "
         "plus \U0001F504 Game History run the steal-fest turn by turn."),
    ]
    for num, title, text in steps:
        ws.set_row(r(row), 16)
        ws.set_row(r(row + 1), 16)
        ws.merge_range(r(row), 1, r(row + 1), 1, num, badge)
        ws.merge_range(r(row), 2, r(row), 3, title, step_title)
        ws.merge_range(r(row), 4, r(row + 1), ci(LAST_COL), text, body)
        row += 2
    row += 1

    # ------------------------------------------------------------------
    sec(row, "\U0001F5FA\uFE0F", "THE TAB TOUR")
    row += 1
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  Premium edition: every tab below.  Basic edition: the "
                   "tabs marked \u25CF only.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.muted,
                                bg_color=th.bg, align="left", valign="vcenter",
                                indent=1)))
    row += 1
    tour = [
        ("\u25CF", "\U0001F3E0 Dashboard", "Participants, spend, readiness, "
         "game state, alerts and three live charts."),
        ("\u25CF", "\U0001F465 Participants", "Guest list with RSVP, team, "
         "household, dietary needs and gift status."),
        ("\u25CF", "\U0001F385 Secret Santa Draw", "Auto draw, overrides, "
         "rules check and the private recipient lookup."),
        ("\u25CF", "\U0001F4B0 Budget Tracker", "Min/max rules, actual "
         "spend, under/over flags and receipt tracking."),
        ("\u25CF", "\U0001F3B2 White Elephant", "Seats, gift numbers, steal "
         "counts, locks, holders and final gifts."),
        ("\u25CF", "\U0001F4D6 Start Here", "This page."),
        ("\u25CB", "\U0001F6AB Exclusions & Rules", "House-rule toggles and "
         "custom no-pair list."),
        ("\u25CB", "\U0001F381 Wishlists", "Per-guest ideas with priorities "
         "and anonymous claiming."),
        ("\u25CB", "\U0001F504 Game History", "The turn log: every Pick, "
         "Steal and Pass in order."),
        ("\u25CB", "\U0001F39F\ufe0f Santa Cards", "Printable cut-out "
         "cards so each guest learns their recipient privately."),
    ]
    for mark_txt, tab, blurb in tour:
        ws.set_row(r(row), 18)
        ws.merge_range(r(row), 2, r(row), 3, mark_txt + "  " + tab, step_title)
        ws.merge_range(r(row), 4, r(row), ci(LAST_COL), blurb, body)
        row += 1
    row += 1

    # ------------------------------------------------------------------
    sec(row, "\U0001F4A1", "TEN TIPS FROM WORKING CATERERS")
    row += 1
    tips = [
        "Set the gift budget min AND max - too-cheap gifts hurt feelings "
        "just as much as too-pricey ones.",
        "Collect wishlists before the draw; three wishes each is plenty.",
        "Households in the same couple get the same Household value - the "
        "couple rule then keeps them apart automatically.",
        "Fill in \u201cgave to last year\u201d and arm the repeat rule: "
        "two years running the same pair feels lazy.",
        "Print the Santa Cards, fold them, and let guests pick a card at "
        "random at the door - zero spoilers.",
        "White Elephant: cap steals at 2.  Three or more and the game "
        "stalls; one and nobody gets naughty.",
        "Log steals on Game History as they happen (phone in Sheets) and "
        "the board stays honest without a whiteboard.",
        "The private lookup shows one recipient at a time - perfect for "
        "passing the laptop round without spoilers.",
        "Formula cells are protected with the password \u201cpremium\u201d; "
        "unlock only when you need to change a seed.",
        "Reuse the file every year: bump the party year and date, clear the "
        "logs, keep your lists.",
    ]
    for i, tip in enumerate(tips, 1):
        ws.set_row(r(row), 16)
        ws.set_row(r(row + 1), 16)
        ws.merge_range(r(row), 1, r(row + 1), 1, str(i), badge2)
        ws.merge_range(r(row), 2, r(row + 1), ci(LAST_COL), tip, body)
        row += 2
    row += 1

    # ------------------------------------------------------------------
    sec(row, "\U0001F512", "THE GENTLE RULES")
    row += 1
    rules = [
        "Cream-filled cells are yours to type in.  White cells are formulas "
        "\u2014 leave them alone and they will keep working for you.",
        "Sheets are protected to stop accidental edits.  Review \u2192 "
        "Unprotect Sheet (no password) if you ever need to restructure "
        "something.",
        "Add rows by copying an existing data row and inserting below it "
        "\u2014 the formulas, dropdowns and colours travel with the copy.",
        "This file opens in Excel 2016+ and in Google Sheets (upload to "
        "Drive \u2192 open with Sheets).  A few chart styles look slightly "
        "different in Sheets; every number still calculates.",
        "The EXAMPLE edition is loaded with a fictional office party so "
        "you can see it working.  The blank edition is the one you keep.",
    ]
    for rule in rules:
        ws.set_row(r(row), 16)
        ws.set_row(r(row + 1), 16)
        ws.merge_range(r(row), 1, r(row + 1), ci(LAST_COL),
                       "\u2022  " + rule, body_soft)
        row += 2
    row += 1

    # ------------------------------------------------------------------
    sec(row, "\U0001F91D", "SUPPORT & GOOD WISHES")
    row += 1
    ws.set_row(r(row), 16)
    ws.set_row(r(row + 1), 16)
    ws.merge_range(r(row), 1, r(row + 1), ci(LAST_COL),
                   "  Thank you for buying this template \u2014 it was "
                   "built by Novality Store, people who have hosted one "
                   "too many white-elephant stampedes and know what a "
                   "spreadsheet owes you in December.  Message the shop "
                   "any time and a human will help quickly.", body_soft)
    row += 3

    # ------------------------------------------------------------------
    # footer + nav
    # ------------------------------------------------------------------
    ws.set_row(r(row), 30)
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  \U0001F4AC  Questions about this file?  Your Etsy "
                   "shop message reaches a human, usually the same day.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav = row + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(key, nav, max_col=LAST_COL)
    bk.page(key, LAST_COL, nav + 1, landscape=False, zoom=90)


# ---------------------------------------------------------------------------
def _banner_path(bk):
    """Watercolour cover art for the Start Here tab, if it exists."""
    if not bk.images:
        return None
    here = os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))))
    for ext in (".png", ".jpg", ".jpeg"):
        path = os.path.join(here, "assets", "banner_%s%s" % (bk.th.key, ext))
        if os.path.exists(path):
            return path
    return None


def _png_size(path):
    """Pixel size of a PNG or JPEG, without needing Pillow."""
    with open(path, "rb") as fh:
        data = fh.read(32)
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        import struct
        w, h = struct.unpack(">II", data[16:24])
        return w, h
    return 1600, 300
```

## 22. Command line interface

`secret_santa_party_tracker.py`

```python
#!/usr/bin/env python3
"""Command line entry point for the Secret Santa & White Elephant tracker."""

import argparse
import sys


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Build Secret Santa & White Elephant Party Tracker "
                    "workbooks (Novality Store).")
    p.add_argument("--edition", choices=("premium", "basic"),
                   default="premium")
    p.add_argument("--theme", choices=("noel", "arctic"), default="noel")
    p.add_argument("--mode", choices=("blank", "demo"), default="blank")
    p.add_argument("--out", default=None)
    p.add_argument("--outdir", default="products")
    p.add_argument("--all", action="store_true",
                   help="build the curated six-file product set")
    p.add_argument("--no-protect", dest="protect", action="store_false",
                   default=None)
    args = p.parse_args(argv)

    sys.path.insert(0, ".")
    from santa_tracker import workbook as W
    if args.all:
        rows = W.build_all(args.outdir, protect=args.protect)
    else:
        from santa_tracker.workbook import product_filename
        out = args.out or product_filename(args.edition, args.theme,
                                           args.mode)
        stats = W.build_workbook(out, args.edition, args.theme, args.mode,
                                 protect=args.protect)
        rows = [(out, stats)]
    print("=" * 78)
    print("BUILD COMPLETE")
    print("=" * 78)
    for name, stats in rows:
        print("  %-58s %7.1f KB" % (name, stats["size"] / 1024.0))
        print("      formulas %5d   validations %3d   cond. formats %3d   "
              "charts %d" % (stats["formulas"], stats["validations"],
                             stats["cond_formats"], stats["charts"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

## 23. QA: structural verifier

`tools/verify_workbook.py`

```python
#!/usr/bin/env python3
"""
Audit a generated .xlsx without needing Excel.

Checks performed
    1. the workbook opens (openpyxl) and every expected tab exists / is hidden
    2. every sheet-qualified reference in every formula points at a real sheet
    3. no modern-only function leaked in without the ``_xlfn.`` prefix older
       Excel versions need
    4. no cell contains the literal strings "None", "#REF!" or "nan"
    5. nothing was written *under* a merged range (Excel only ever shows the
       top-left cell of a merge, so such writes would be invisible)

Usage
    python3 tools/verify_workbook.py products/*.xlsx
"""

import glob
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils import (column_index_from_string, get_column_letter,
                            range_boundaries)

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Functions that did not exist in Excel 2007 and therefore must be written as
# _xlfn.NAME by a file generator (XlsxWriter does this automatically for the
# names it knows about).
MODERN = {
    "TEXTJOIN", "CONCAT", "IFS", "SWITCH", "MAXIFS", "MINIFS", "XLOOKUP",
    "XMATCH", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "LET",
    "LAMBDA", "RANDARRAY", "IMAGE", "TEXTSPLIT", "VSTACK", "HSTACK", "TOCOL",
    "TOROW", "TAKE", "DROP", "CHOOSEROWS", "CHOOSECOLS", "WRAPROWS",
    "WRAPCOLS", "GROUPBY", "PIVOTBY", "REGEXTEST", "REGEXEXTRACT",
    "REGEXREPLACE",
}

FUNC_RE = re.compile(r"(?<![A-Za-z0-9_.])([A-Z][A-Z0-9_.]*)\s*\(")
SHEET_REF_RE = re.compile(r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!")


def sheet_xml(path):
    """{sheet name: worksheet xml} for low level checks."""
    zf = zipfile.ZipFile(path)
    wbxml = zf.read("xl/workbook.xml").decode()
    rels = dict(re.findall(
        r'Id="(rId\d+)"[^>]*Target="([^"]*worksheets/sheet\d+\.xml)"',
        zf.read("xl/_rels/workbook.xml.rels").decode()))
    out = {}
    for name, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"',
                                wbxml):
        out[name] = zf.read("xl/" + rels[rid].lstrip("/")).decode()
    return out, [n for n in zipfile.ZipFile(path).namelist()]


def hidden_writes(path):
    """Cells written inside a merged range but not at its top-left."""
    xmls, _ = sheet_xml(path)
    out = []
    for name, xml in xmls.items():
        covered = {}
        for m in re.findall(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"/>', xml):
            c1, r1, c2, r2 = range_boundaries(m)
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    if (rr, cc) != (r1, c1):
                        covered[(rr, cc)] = m
        root = ET.fromstring(xml)
        for row in root.iter(NS + "row"):
            for c in row.iter(NS + "c"):
                m = re.match(r"([A-Z]+)(\d+)", c.get("r"))
                key = (int(m.group(2)), column_index_from_string(m.group(1)))
                if key not in covered:
                    continue
                v = c.find(NS + "v")
                if c.find(NS + "f") is not None or (
                        v is not None and (v.text or "").strip()):
                    out.append((name, c.get("r"), covered[key]))
    return out


def check_formulas(path, problems):
    wb = openpyxl.load_workbook(path, data_only=False)
    sheetnames = set(wb.sheetnames)
    n_formula = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                n_formula += 1
                if v in ("None", "nan"):
                    problems.append("%s!%s contains %r"
                                    % (ws.title, cell.coordinate, v))
                body = re.sub(r'"[^"]*"', '""', v)     # drop string literals
                for m in SHEET_REF_RE.finditer(body):
                    ref = m.group(1) or m.group(2)
                    if ref not in sheetnames:
                        problems.append("%s!%s references unknown sheet %r"
                                        % (ws.title, cell.coordinate, ref))
                for m in FUNC_RE.finditer(body):
                    fn = m.group(1).split(".")[-1]
                    if fn in MODERN and "_xlfn." not in m.group(1):
                        problems.append(
                            "%s!%s uses modern function %s without _xlfn."
                            % (ws.title, cell.coordinate, fn))
    return n_formula


def audit(path):
    problems = []
    wb = openpyxl.load_workbook(path, data_only=False)
    n_formula = check_formulas(path, problems)

    n_dv = sum(len(ws.data_validations.dataValidation)
               for ws in wb.worksheets)
    n_cf = sum(len(ws.conditional_formatting._cf_rules)
               for ws in wb.worksheets)
    _, names = sheet_xml(path)
    charts = len([n for n in names if n.startswith("xl/charts/chart")])

    for sheet, ref, merge in hidden_writes(path):
        problems.append("%s!%s is written inside merge %s and would never "
                        "be visible" % (sheet, ref, merge))

    return {
        "path": path,
        "sheets": wb.sheetnames,
        "hidden": [s.title for s in wb.worksheets
                   if s.sheet_state != "visible"],
        "formulas": n_formula,
        "validations": n_dv,
        "cond_formats": n_cf,
        "charts": charts,
        "problems": problems,
    }


def main(argv):
    paths = []
    for a in argv:
        paths.extend(sorted(glob.glob(a)) if any(c in a for c in "*?[") else [a])
    if not paths:
        print(__doc__)
        return 1
    bad = 0
    for path in paths:
        r = audit(path)
        print("=" * 78)
        print(path)
        print("  sheets   : %s" % ", ".join(r["sheets"]))
        print("  hidden   : %s" % (", ".join(r["hidden"]) or "-"))
        print("  formulas : %-6d validations %-4d cond. formats %-4d "
              "charts %d" % (r["formulas"], r["validations"],
                             r["cond_formats"], r["charts"]))
        if r["problems"]:
            bad += 1
            print("  PROBLEMS (%d):" % len(r["problems"]))
            for p in r["problems"][:25]:
                print("    - %s" % p)
            if len(r["problems"]) > 25:
                print("    ... and %d more" % (len(r["problems"]) - 25))
        else:
            print("  PROBLEMS : none \u2713")
    print("=" * 78)
    print("clean" if not bad else "%d file(s) need attention" % bad)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

## 24. QA: formula recalculation checker

`tools/calc_check.py`

```python
#!/usr/bin/env python3
"""
Recalculate a generated workbook with the ``formulas`` engine and report every
cell that evaluates to an Excel error (#REF!, #NAME?, #VALUE!, #DIV/0! ...).

This is the closest thing to opening the file in Excel without Excel: every
formula is parsed and evaluated for real, so broken references, typos and
unsupported functions show up here instead of in a customer's hands.

    /tmp/venv/bin/pip install formulas
    python3 tools/calc_check.py products/Christmas_Gift_Tracker_PREMIUM_Festive_EXAMPLE.xlsx

Known limitation: the ``formulas`` package cannot evaluate OFFSET-based defined
names (the auto-expanding dropdown lists), so those report #REF! here while
Excel and Google Sheets handle them normally.  They are listed separately as
"engine limitations" rather than failures.
"""

import glob
import logging
import re
import sys
import warnings

logging.disable(logging.CRITICAL)
warnings.simplefilter("ignore")

import formulas                                     # noqa: E402
import openpyxl                                     # noqa: E402


def check(path, verbose=True, compare=True):
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()

    wb = openpyxl.load_workbook(path)
    wbv = openpyxl.load_workbook(path, data_only=True)
    book = path.split("/")[-1]
    book = re.escape("[%s]" % path.split("/")[-1].upper())

    def formula_of(key):
        m = re.match(r"^'\[.*\](.*)'!([A-Z]+\d+)$", key)
        if not m:
            return None, None
        sheet, coord = m.group(1), m.group(2)
        for ws in wb.worksheets:
            if ws.title.upper() == sheet:
                return ws.title, ws[coord].value
        return sheet, None

    mismatches = []
    if compare:
        for ws in wb.worksheets:
            wsv = wbv[ws.title]
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue
                    key = "'[%s]%s'!%s" % (book, ws.title.upper(),
                                           cell.coordinate)
                    got = sol.get(key)
                    if got is None:
                        continue
                    try:
                        calc = got.value[0, 0]
                    except Exception:
                        calc = got
                    cached = wsv[cell.coordinate].value
                    if calc is None and cached is None:
                        continue
                    if isinstance(calc, (int, float)) and \
                            isinstance(cached, (int, float)):
                        if abs(float(calc) - float(cached)) > 1e-6:
                            mismatches.append((ws.title, cell.coordinate,
                                               calc, cached))
                    elif str(calc).strip() != str(cached).strip():
                        mismatches.append((ws.title, cell.coordinate,
                                           calc, cached))

    errors, limits = [], []
    for key, value in sol.items():
        try:
            val = value.value[0, 0] if hasattr(value, "value") else value
        except Exception:
            continue
        text = str(val).strip()
        if not text.startswith("#"):
            continue
        if text == "#":            # a literal "#" header cell, not an error
            continue
        sheet, f = formula_of(key)
        if sheet is None:          # workbook-level defined name, not a cell
            limits.append((text, key, sheet, f))
            continue
        # things the engine cannot do but Excel and Google Sheets can:
        #   OFFSET() defined names (the auto-expanding dropdown lists)
        #   HYPERLINK()  (not implemented in the formulas package)
        if "OFFSET" in key.upper() or (f and "HYPERLINK(" in str(f).upper()):
            limits.append((text, key, sheet, f))
            continue
        errors.append((text, key, sheet, f))

    if verbose:
        print("=" * 78)
        print(path)
        print("  cells evaluated : %d" % len(sol))
        print("  engine limitations (fine in Excel / Sheets) : %d %s"
              % (len(limits), sorted({t for t, k, s2, f in limits})))
        if mismatches:
            print("  CACHED VALUE OUT OF DATE (%d):" % len(mismatches))
            for sheet, coord, calc, cached in mismatches[:15]:
                print("      %-20s %-6s recalcs to %r, file says %r"
                      % (sheet, coord, str(calc)[:30], str(cached)[:30]))
            if len(mismatches) > 15:
                print("      ... and %d more" % (len(mismatches) - 15))
        if errors:
            print("  ERRORS (%d):" % len(errors))
            for text, key, sheet, f in errors[:30]:
                print("    %-9s %-28s %s" % (text, (sheet or "?"),
                                             str(f)[:90]))
            if len(errors) > 30:
                print("    ... and %d more" % (len(errors) - 30))
        else:
            print("  ERRORS : none \u2713")
    return errors + mismatches, limits


def main(argv):
    paths = []
    for a in argv:
        paths.extend(sorted(glob.glob(a)) if any(c in a for c in "*?[") else [a])
    if not paths:
        print(__doc__)
        return 1
    total = 0
    for p in paths:
        errors, _ = check(p)
        total += len(errors)
    print("=" * 78)
    print("clean" if not total else "%d error cell(s) across %d file(s)"
          % (total, len(paths)))
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

## 25. QA: layout / clipping checker

`tools/layout_check.py`

```python
#!/usr/bin/env python3
"""
Layout audit: finds text that will be clipped or overflow in Excel.

Excel never grows a merged cell to fit its text, so wrapped text in a merged
range needs an explicit row height.  This tool re-measures every cell in a
generated workbook the same way the builder did and reports:

    TOO TALL  wrapped text needs more height than the row(s) provide
    OVERFLOW  unwrapped text is wider than its cell and the neighbour is busy
    TINY ROW  an explicit row height smaller than one line of its font

    python3 tools/layout_check.py products/*.xlsx
"""

import glob
import math
import sys

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

DEFAULT_COL_WIDTH = 8.43
CHARS_PER_UNIT = 1.05        # measured against Calibri 11 in Excel


def width_map(ws):
    """Column widths, expanding grouped ranges such as ``set_column("C:L")``."""
    widths = {}
    for dim in ws.column_dimensions.values():
        if not dim.width:
            continue
        lo = dim.min or 1
        hi = dim.max or lo
        for idx in range(lo, hi + 1):
            widths[idx] = dim.width
    return widths


def col_width(ws, idx, widths=None):
    if widths is None:
        widths = width_map(ws)
    return widths.get(idx, DEFAULT_COL_WIDTH)


def row_height(ws, idx, default=15.0):
    dim = ws.row_dimensions.get(idx)
    if dim is not None and dim.height:
        return dim.height
    return default


def check(path, verbose=True):
    wb = openpyxl.load_workbook(path)
    problems = []
    for ws in wb.worksheets:
        widths = width_map(ws)
        merged = {}
        for mr in ws.merged_cells.ranges:
            c1, r1, c2, r2 = range_boundaries(str(mr))
            merged[(r1, c1)] = (r1, c1, r2, c2)
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    if (rr, cc) != (r1, c1):
                        merged[(rr, cc)] = None      # swallowed cell

        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.strip():
                    continue
                if v.startswith("="):
                    continue                       # formula: length unknown
                key = (cell.row, cell.column)
                if key in merged and merged[key] is None:
                    continue                       # hidden by a merge
                blocked = True
                if key in merged:
                    r1, c1, r2, c2 = merged[key]
                    width = sum(col_width(ws, c, widths)
                                for c in range(c1, c2 + 1))
                    height = sum(row_height(ws, r) for r in range(r1, r2 + 1))
                else:
                    width = col_width(ws, cell.column, widths)
                    height = row_height(ws, cell.row)
                    nxt = ws.cell(row=cell.row, column=cell.column + 1)
                    blocked = (nxt.value not in (None, "")
                               or (cell.row, cell.column + 1) in merged)

                size = cell.font.size or 11
                room = width * CHARS_PER_UNIT * (11.0 / float(size))
                text = v.strip()
                wrap = bool(cell.alignment.wrap_text)

                if wrap:
                    lines = 0
                    for part in text.split("\n"):
                        lines += max(1, int(math.ceil(len(part) / max(room, 1))))
                    line_h = size * 1.32 + 2.2
                    need = lines * line_h
                    if need > height + 1.5:
                        problems.append(
                            ("TOO TALL", ws.title, cell.coordinate,
                             "needs ~%.0fpt for %d line(s), has %.0fpt: %r"
                             % (need, lines, height, text[:60])))
                else:
                    if len(text) > room * 1.02 and key not in merged:
                        if blocked:
                            # a pasted web address is always longer than its
                            # column: the clickable button next to it is the
                            # intended way in, so this is by design
                            kind = ("CLIPPED LINK"
                                    if text.lower().startswith("http")
                                    else "OVERFLOW")
                            problems.append(
                                (kind, ws.title, cell.coordinate,
                                 "%d chars in %.0f wide cell: %r"
                                 % (len(text), width, text[:60])))
                if height < size * 1.15 and key not in merged:
                    problems.append(
                        ("TINY ROW", ws.title, cell.coordinate,
                         "row %.0fpt for %spt font" % (height, size)))
    if verbose:
        print("=" * 78)
        print(path)
        if problems:
            by_kind = {}
            for kind, sheet, coord, msg in problems:
                by_kind.setdefault(kind, []).append((sheet, coord, msg))
            for kind, items in sorted(by_kind.items()):
                print("  %-9s %d" % (kind, len(items)))
                for sheet, coord, msg in items[:12]:
                    print("      %-22s %-6s %s" % (sheet, coord, msg))
                if len(items) > 12:
                    print("      ... and %d more" % (len(items) - 12))
        else:
            print("  layout: no clipping or overflow found \u2713")
    return problems


def main(argv):
    paths = []
    for a in argv:
        paths.extend(sorted(glob.glob(a)) if any(c in a for c in "*?[") else [a])
    if not paths:
        print(__doc__)
        return 1
    total = 0
    for p in paths:
        total += len([x for x in check(p) if x[0] != "CLIPPED LINK"])
    print("=" * 78)
    print("clean" if not total else "%d layout issue(s)" % total)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

## 26. QA: sheet renderer (visual proof)

`tools/render_preview.py`

```python
#!/usr/bin/env python3
"""
Render a sheet of a generated .xlsx to a PNG so the design can be eyeballed
without Excel (there is no Excel or LibreOffice in this sandbox).

It is an approximation of Excel's rendering: cached formula values, solid
fills, borders, merged ranges, wrapped text, horizontal/vertical alignment,
number formats and text that spills into empty neighbours.  Emoji are drawn as
coloured tiles because no emoji font is installed - their position and size are
still visible, which is what matters for layout checking.

    /tmp/venv/bin/python tools/render_preview.py products/x.xlsx "🎄 Dashboard" \\
        /tmp/dash.png --rows 1-90
"""

import argparse
import datetime
import re
import sys
import unicodedata

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = "/usr/share/fonts/truetype/dejavu/"
PX_PER_WIDTH = 7.0
PT_TO_PX = 96.0 / 72.0
DEFAULT_ROW_PX = 20
EMOJI_RE = re.compile(
    "[\U0001F000-\U0001FAFF\u2600-\u27BF\u2B00-\u2BFF\uFE0F\u2190-\u21FF"
    "\u2700-\u27BF\u2764\u2699\u26A1\u2705\u274C\u2753\u2764]")


def is_emoji(ch):
    if ord(ch) < 0x2100:
        return False
    if ord(ch) in (0xFE0F, 0x2022, 0x2014, 0x2013, 0x201C, 0x201D, 0x2019,
                   0x2026, 0x00B7):
        return False
    cat = unicodedata.category(ch)
    return cat in ("So", "Sk", "Cs") or ord(ch) > 0x2500


def font(size, bold=False, italic=False):
    name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
    return ImageFont.truetype(FONT_DIR + name, max(6, int(round(size))))


def color_of(c, default=None):
    if c is None:
        return default
    rgb = getattr(c, "rgb", None)
    if isinstance(rgb, str) and len(rgb) >= 6:
        rgb = rgb[-6:]
        try:
            return tuple(int(rgb[i:i + 2], 16) for i in (0, 2, 4))
        except ValueError:
            return default
    return default


def fmt_value(value, num_format):
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        if value.hour or value.minute:
            return value.strftime("%d %b %Y %H:%M")
        return value.strftime("%d %b %Y")
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    nf = (num_format or "").lower()
    if isinstance(value, (int, float)):
        if "yy" in nf or "dd" in nf or "mmmm" in nf or "hh" in nf:
            try:
                dt = (datetime.datetime(1899, 12, 30)
                      + datetime.timedelta(days=float(value)))
            except (OverflowError, ValueError):
                return str(value)
            if "yyyy" in nf and "d" not in nf:
                return dt.strftime("%Y")
            if "dddd" in nf:
                return dt.strftime("%A %d %B %Y")
            if "ddd" in nf or "dd mmm" in nf or "d mmm" in nf:
                return dt.strftime("%d %b %Y")
            if "mm" in nf or "yy" in nf:
                return dt.strftime("%d/%m/%Y")
            return dt.strftime("%d %b")
        if "%" in nf:
            dp = 2 if "0.00%" in nf else (1 if "0.0%" in nf else 0)
            return ("%%.%df%%%%" % dp) % (value * 100)
        if "#,##0.00" in nf:
            return "{:,.2f}".format(value)
        if "#,##0" in nf or nf in ("0", "#,##0.0"):
            return "{:,.0f}".format(value) if "#,##0.0" not in nf \
                else "{:,.1f}".format(value)
        if value == int(value):
            return str(int(value))
        return "%.2f" % value
    return str(value)


def render(path, sheet_title, out, rows=None, cols=None, zoom=1.0):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_title]

    # --- geometry -------------------------------------------------------
    widths = {}
    for dim in ws.column_dimensions.values():
        if not dim.width:
            continue
        for idx in range(dim.min or 1, (dim.max or dim.min or 1) + 1):
            widths[idx] = dim.width
    heights = {i: d.height for i, d in ws.row_dimensions.items() if d.height}

    r1, r2 = (rows or (1, ws.max_row))
    c1, c2 = (cols or (1, ws.max_column))

    def col_px(i):
        return int(round(widths.get(i, 8.43) * PX_PER_WIDTH)) + 5

    def row_px(i):
        return int(round(heights.get(i, 15.0) * PT_TO_PX)) or DEFAULT_ROW_PX

    xs = {}
    x = 0
    for i in range(c1, c2 + 1):
        xs[i] = x
        x += col_px(i)
    ys = {}
    y = 0
    for i in range(r1, r2 + 1):
        ys[i] = y
        y += row_px(i)
    W, H = int(x * zoom), int(y * zoom)
    img = Image.new("RGB", (max(W, 1), max(H, 1)), (255, 255, 255))
    d = ImageDraw.Draw(img)

    merges = {}
    swallowed = set()
    for mr in ws.merged_cells.ranges:
        mc1, mr1, mc2, mr2 = range_boundaries(str(mr))
        merges[(mr1, mc1)] = (mr1, mc1, mr2, mc2)
        for rr in range(mr1, mr2 + 1):
            for cc in range(mc1, mc2 + 1):
                if (rr, cc) != (mr1, mc1):
                    swallowed.add((rr, cc))

    def cell_box(rr, cc):
        if (rr, cc) in merges:
            a, b, c_, e = merges[(rr, cc)]
            x1, y1 = xs.get(b, x), ys.get(a, y)
            x2 = xs.get(e, x) + col_px(e)
            y2 = ys.get(c_, y) + row_px(c_)
        else:
            x1, y1 = xs.get(cc, x), ys.get(rr, y)
            x2, y2 = x1 + col_px(cc), y1 + row_px(rr)
        return x1, y1, x2, y2

    def occupied(rr, cc):
        c = ws.cell(row=rr, column=cc)
        return c.value not in (None, "")

    # --- paint ----------------------------------------------------------
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            if (rr, cc) in swallowed:
                continue
            cell = ws.cell(row=rr, column=cc)
            x1, y1, x2, y2 = cell_box(rr, cc)
            if x2 <= 0 or y2 <= 0 or x1 >= x or y1 >= y:
                continue
            fill = color_of(cell.fill.start_color) if (
                cell.fill and cell.fill.patternType == "solid") else None
            if fill:
                d.rectangle([x1, y1, x2 - 1, y2 - 1], fill=fill)
            b = cell.border
            for side, edges in ((b.left, (x1, y1, x1, y2)),
                                (b.right, (x2 - 1, y1, x2 - 1, y2)),
                                (b.top, (x1, y1, x2, y1)),
                                (b.bottom, (x1, y2 - 1, x2, y2 - 1))):
                if side is not None and side.style:
                    col = color_of(side.color, (200, 200, 200))
                    thick = 2 if side.style in ("medium", "thick", "double") \
                        else 1
                    d.line(list(edges), fill=col, width=thick)

    # --- text -----------------------------------------------------------
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            if (rr, cc) in swallowed:
                continue
            cell = ws.cell(row=rr, column=cc)
            raw = fmt_value(cell.value, cell.number_format)
            if not raw:
                continue
            x1, y1, x2, y2 = cell_box(rr, cc)
            if x2 <= 0 or y2 <= 0 or x1 >= x or y1 >= y:
                continue
            al = cell.alignment
            fnt = font((cell.font.size or 11) * 1.02,
                       bool(cell.font.bold), bool(cell.font.italic))
            col = color_of(cell.font.color, (0, 0, 0))
            indent = int((al.indent or 0) * 8)
            pad = 4 + indent

            # spill into empty neighbours when the text is not wrapped
            box_w = x2 - x1
            if not al.wrap_text:
                cc2 = cc + 1
                while (cc2 <= c2 and not occupied(rr, cc2)
                       and (rr, cc2) not in swallowed
                       and d.textlength(raw, font=fnt) > box_w - 2 * pad):
                    box_w += col_px(cc2)
                    cc2 += 1

            if al.wrap_text:
                words, lines, cur = raw.replace("\n", " \n ").split(), [], ""
                for w in words:
                    trial = (cur + " " + w).strip()
                    if d.textlength(trial, font=fnt) <= box_w - 2 * pad \
                            or not cur:
                        cur = trial
                    else:
                        lines.append(cur)
                        cur = w
                    if w == "\n":
                        lines.append(cur)
                        cur = ""
                if cur:
                    lines.append(cur)
            else:
                lines = raw.split("\n")

            lh = int((cell.font.size or 11) * 1.42)
            total_h = lh * len(lines)
            va = al.vertical or "bottom"
            if va == "center":
                ty = y1 + max(0, (y2 - y1 - total_h) // 2)
            elif va == "top":
                ty = y1 + 3
            else:
                ty = y1 + max(0, (y2 - y1 - total_h) // 2)

            for line in lines:
                drawn, cx = [], x1 + pad
                for ch in line:
                    if is_emoji(ch):
                        drawn.append(("tile", ch))
                    else:
                        drawn.append(("text", ch))
                # measure
                widths_px, tiles = [], []
                for kind, ch in drawn:
                    if kind == "tile":
                        wpx = int(lh * 0.95)
                        tiles.append(wpx)
                    else:
                        wpx = d.textlength(ch, font=fnt)
                    widths_px.append(wpx)
                line_w = sum(widths_px)
                ha = al.horizontal or ("right" if isinstance(
                    cell.value, (int, float)) and not cell.font.bold
                    else "left")
                if ha == "center":
                    cx = x1 + max(pad, (box_w - line_w) // 2)
                elif ha == "right":
                    cx = x1 + max(pad, box_w - line_w - pad)
                for (kind, ch), wpx in zip(drawn, widths_px):
                    if kind == "tile":
                        hue = (ord(ch) * 47) % 360
                        import colorsys
                        rgb = tuple(int(255 * v) for v in
                                    colorsys.hsv_to_rgb(hue / 360.0, 0.55, 0.85))
                        d.rounded_rectangle([cx, ty + 1, cx + wpx - 2,
                                             ty + lh - 2], radius=3, fill=rgb)
                    else:
                        d.text((cx, ty), ch, font=fnt, fill=col)
                    cx += wpx
                ty += lh
    # --- inserted pictures ---------------------------------------------
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, \
        TwoCellAnchor, AbsoluteAnchor
    for pic in getattr(ws, "_images", []):
        anc = pic.anchor
        frm = getattr(anc, "_from", None)
        if frm is None:
            continue
        col, row = frm.col + 1, frm.row + 1
        if not (c1 <= col <= c2 and r1 <= row <= r2):
            continue
        x1 = xs.get(col, 0) + int((frm.colOff or 0) / 9525.0)
        y1 = ys.get(row, 0) + int((frm.rowOff or 0) / 9525.0)
        to = getattr(anc, "to", None)
        if to is not None:
            x2 = xs.get(to.col + 1, x) + int((to.colOff or 0) / 9525.0)
            y2 = ys.get(to.row + 1, y) + int((to.rowOff or 0) / 9525.0)
            wpx, hpx = x2 - x1, y2 - y1
        else:
            try:
                ext = anc.ext
                wpx, hpx = int(ext.cx / 9525.0), int(ext.cy / 9525.0)
            except AttributeError:
                wpx, hpx = int(pic.width), int(pic.height)
        try:
            import io as _io
            sub = Image.open(_io.BytesIO(pic._data())).convert("RGBA")
            sub = sub.resize((max(wpx, 1), max(hpx, 1)), Image.LANCZOS)
            img.paste(sub, (x1, y1), sub)
        except Exception as exc:
            print("  (image skipped: %s)" % exc)

    if zoom != 1.0:
        img = img.resize((W, H), Image.LANCZOS)
    img.save(out)
    return out, img.size


def main(argv):
    p = argparse.ArgumentParser()
    p.add_argument("workbook")
    p.add_argument("sheet")
    p.add_argument("out")
    p.add_argument("--rows", default=None, help="e.g. 1-90")
    p.add_argument("--cols", default=None, help="e.g. A-M")
    p.add_argument("--zoom", type=float, default=1.0)
    a = p.parse_args(argv)
    rows = tuple(int(x) for x in a.rows.split("-")) if a.rows else None
    cols = None
    if a.cols:
        lo, hi = a.cols.split("-")
        cols = (openpyxl.utils.column_index_from_string(lo),
                openpyxl.utils.column_index_from_string(hi))
    out, size = render(a.workbook, a.sheet, a.out, rows, cols, a.zoom)
    print("%s -> %s (%dx%d)" % (a.sheet, out, size[0], size[1]))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

## 27. Banner alpha pipeline

`tools/make_banner_alpha.py`

```python
#!/usr/bin/env python3
"""
Give the cover banners a transparent middle.

Excel draws inserted pictures *above* the cell grid, so text typed in the cells
underneath an opaque picture can never be seen.  The Start Here tab puts its
title in cells under the banner; for that to work the flat cream/white middle
of the artwork has to be see-through.

This script converts ``assets/banner_<theme>.jpg`` (flat background) into
``assets/banner_<theme>.png`` with an alpha channel: everything inside the
title window that matches the background colour becomes transparent, with a
one pixel feather so the watercolour edges stay soft.

    /tmp/venv/bin/python tools/make_banner_alpha.py
"""

import os

from PIL import Image, ImageFilter

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(HERE, "assets")

# title window as a fraction of the image (matches the C:H cell overlay)
WIN = (0.06, 0.02, 0.68, 0.98)
TOLERANCE = 26


def background_color(im):
    w, h = im.size
    px = im.load()
    samples = [px[int(w * 0.5), int(h * f)] for f in (0.02, 0.5, 0.98)]
    r = sum(s[0] for s in samples) // 3
    g = sum(s[1] for s in samples) // 3
    b = sum(s[2] for s in samples) // 3
    return r, g, b


def convert(src, dst):
    im = Image.open(src).convert("RGB")
    w, h = im.size
    bg = background_color(im)
    alpha = Image.new("L", (w, h), 255)
    ap = alpha.load()
    px = im.load()
    x1, y1, x2, y2 = (int(w * WIN[0]), int(h * WIN[1]),
                      int(w * WIN[2]), int(h * WIN[3]))
    for y in range(y1, y2):
        for x in range(x1, x2):
            r, g, b = px[x, y]
            if abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2]) < TOLERANCE * 3:
                ap[x, y] = 0
    alpha = alpha.filter(ImageFilter.GaussianBlur(1.1))
    out = im.convert("RGBA")
    out.putalpha(alpha)
    out.save(dst, optimize=True)
    return out.size, os.path.getsize(dst) // 1024


def main():
    import sys
    stems = sys.argv[1:] or ("banner_festive", "banner_minimal",
                             "banner_classic", "banner_fresh")
    for stem in stems:
        src = os.path.join(ASSETS, stem + ".jpg")
        dst = os.path.join(ASSETS, stem + ".png")
        if not os.path.exists(src):
            print("missing", src)
            continue
        size, kb = convert(src, dst)
        print("%-16s %s -> %s (%d KB)" % (stem, size, dst, kb))
        os.remove(src)


if __name__ == "__main__":
    main()
```

---

## Appendix A - build & QA recipe

```bash
python3 -m venv /tmp/venv && /tmp/venv/bin/pip install openpyxl Pillow formulas
python3 secret_santa_party_tracker.py --all --outdir products
for f in products/Secret_Santa_*.xlsx; do
    /tmp/venv/bin/python tools/verify_workbook.py "$f"
    /tmp/venv/bin/python tools/calc_check.py "$f"
    /tmp/venv/bin/python tools/layout_check.py "$f"
done
```

## Appendix B - changelog

* **1.0.0 (2026-09-18)** - first release: 12-sheet Premium / 7-sheet Basic,
  Noel Classic + Arctic Frost themes, blank and EXAMPLE modes, seed-driven
  derangement draw with exclusion validation, White Elephant engine driven
  by a turn log, printable Santa Cards, Novality Store branding, protection
  password `premium`, full QA pass on all six shipped files.
