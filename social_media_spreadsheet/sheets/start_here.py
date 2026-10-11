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
