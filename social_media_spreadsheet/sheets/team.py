"""`Team` - who does what, and how much of it."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col, yesno)

NAME = "Team"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"TM-"', digits=2, width=8),
        Col("name", "Name", 20, "text",
            note="These names fill the 'Owner' drop downs across the "
                 "workbook."),
        Col("role", "Role", 20, "text", validate=list_from("RoleList")),
        Col("contact", "Email / contact", 26, "text"),
        Col("platform", "Looks after", 20, "text"),
        Col("hours", "Hours a week", 11, "number1",
            validate=number_check(0, 168), cf=bar_cf()),
        Col("rate", "Hourly rate", 12, "money",
            validate=number_check(0, 1e6)),
        Col("reviews", "Reviews content?", 12, "flag", validate=yesno(),
            cf=flag_yes_no_cf()),
        Col("approves", "Final approval?", 12, "flag", validate=yesno(),
            cf=flag_yes_no_cf()),
        Col("weekly_cost", "Cost a week", 12, "money",
            formula='=IF(OR({hours}{r}="",{rate}{r}=""),"",'
                    '{hours}{r}*{rate}{r})'),
        Col("monthly_cost", "Cost a month", 12, "money",
            formula='=IF({weekly_cost}{r}="","",'
                    '{weekly_cost}{r}*52/12)'),
        Col("posts", "Posts owned", 11, "int",
            formula=blank("name",
                          'COUNTIF([[Content Calendar.owner]],{name}{r})'),
            cf=bar_cf()),
        Col("published", "Published", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.owner]],{name}{r},'
                          '[[Content Calendar.status]],"Published")')),
        Col("late", "Late posts", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.owner]],{name}{r},'
                          '[[Content Calendar.alert]],"Overdue*")'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}]),
        Col("tasks_open", "Open tasks", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Tasks.owner]],{name}{r},'
                          '[[Tasks.status]],"<>Complete")')),
        Col("tasks_late", "Overdue tasks", 11, "int",
            formula=blank("name",
                          'COUNTIFS([[Tasks.owner]],{name}{r},'
                          '[[Tasks.status]],"<>Complete",'
                          '[[Tasks.due]],"<"&TODAY())'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("on_time", "On-time rate", 11, "percent",
            formula='=IF(OR({posts}{r}="",{posts}{r}=0),"",'
                    '1-{late}{r}/{posts}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": 0.9,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.75,
                 "fmt": "warn"}]),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "name",
                'IF({tasks_late}{r}>0,"Action - overdue tasks",'
                'IF({late}{r}>0,"Check - late content",'
                'IF({posts}{r}=0,"Add - nothing assigned","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Team & Responsibilities",
        subtitle="Even a team of one. Set who owns what, and the sheet "
                 "counts the posts and tasks sitting with each person.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).team,
        help_text="If it is just you, add one row with your name - every "
                  "owner drop down in the workbook reads this list.",
    )
