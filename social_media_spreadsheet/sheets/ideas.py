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
