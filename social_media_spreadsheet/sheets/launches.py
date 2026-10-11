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
