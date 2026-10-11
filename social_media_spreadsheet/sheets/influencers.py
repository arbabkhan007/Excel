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
