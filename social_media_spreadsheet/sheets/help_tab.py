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
