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
