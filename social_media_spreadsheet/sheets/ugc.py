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
