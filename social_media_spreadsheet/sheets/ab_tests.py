"""`A/B Tests` - change one thing, measure it, keep the winner."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, status_cf

NAME = "AB Tests"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"AB-"', digits=3, width=9),
        Col("name", "Test name", 24, "text"),
        Col("variable", "What you changed", 16, "text",
            validate=list_from("TestVariableList"),
            note="Change one thing only, or you will not know what "
                 "caused the difference."),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("hypothesis", "Hypothesis", 34, "wrap",
            note="'If I ... then ... because ...'"),
        Col("version_a", "Version A", 26, "wrap"),
        Col("version_b", "Version B", 26, "wrap"),
        Col("audience", "Kept the same", 24, "wrap"),
        Col("metric", "Main measure", 16, "text"),
        Col("secondary", "Second measure", 15, "text"),
        Col("start", "Started", 12, "date", validate=date_check()),
        Col("end", "Finished", 12, "date", validate=date_check()),
        Col("a_result", "A result", 11, "percent",
            validate=number_check(0, 10)),
        Col("b_result", "B result", 11, "percent",
            validate=number_check(0, 10)),
        Col("a_clicks", "A clicks", 10, "int",
            validate=number_check(0, 1e8, decimal=False)),
        Col("b_clicks", "B clicks", 10, "int",
            validate=number_check(0, 1e8, decimal=False)),
        Col("a_conversions", "A sales", 9, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("b_conversions", "B sales", 9, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("status", "Status", 13, "status",
            validate=list_from("TestStatusList"), cf=status_cf()),
        Col("difference", "Difference", 11, "percent",
            formula='=IF(OR({a_result}{r}="",{b_result}{r}=""),"",'
                    '{b_result}{r}-{a_result}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("uplift", "Uplift", 11, "percent",
            formula='=IF(OR({a_result}{r}="",{a_result}{r}=0,'
                    '{b_result}{r}=""),"",{b_result}{r}/{a_result}{r}-1)'),
        Col("winner", "Winner", 11, "text",
            formula='=IF(OR({a_result}{r}="",{b_result}{r}=""),"",'
                    'IF(ABS({uplift}{r})<0.1,"Too close",'
                    'IF({b_result}{r}>{a_result}{r},"B","A")))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Too close", "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "B",
                 "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "A",
                 "fmt": "good"}]),
        Col("sample_size", "Total clicks", 11, "int",
            formula='=IF(AND({a_clicks}{r}="",{b_clicks}{r}=""),"",'
                    'N({a_clicks}{r})+N({b_clicks}{r}))'),
        Col("trust", "How much to trust it", 22, "status",
            formula='=IF({winner}{r}="","",IF({sample_size}{r}<100,'
                    '"Low - too small a sample",IF(ABS({uplift}{r})<0.1,'
                    '"Low - the gap is noise",IF({sample_size}{r}<400,'
                    '"Fair - worth repeating","Good - act on it"))))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Good", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Fair", "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "Low",
                 "fmt": "bad"}]),
        Col("days", "Days run", 9, "int",
            formula='=IF(OR({start}{r}="",{end}{r}=""),"",'
                    '{end}{r}-{start}{r}+1)'),
        Col("confidence", "Your read on it", 26, "wrap"),
        Col("next_test", "Next test to run", 26, "wrap"),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "name",
                'IF(AND({status}{r}="Running",{end}{r}<>"",'
                '{end}{r}<TODAY()),"Due - test finished, record the '
                'result",IF(AND({status}{r}="Complete",'
                '{next_test}{r}=""),"Add - what will you test next?",'
                'IF({trust}{r}="Low - too small a sample",'
                '"Check - run it for longer","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="A/B Testing Tracker",
        subtitle="Two versions, one difference, one measure. Small "
                 "samples lie, so the sheet tells you how much to trust "
                 "each result.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).tests,
        help_text="Test one variable at a time - hook, cover, posting "
                  "time or call to action - and run it long enough to "
                  "collect a few hundred interactions.",
    )
