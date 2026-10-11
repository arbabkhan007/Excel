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
