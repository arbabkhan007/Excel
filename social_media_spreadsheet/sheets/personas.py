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
