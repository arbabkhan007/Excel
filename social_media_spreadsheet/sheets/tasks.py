"""`Tasks` - the jobs behind the posts."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, id_col, lookup_content, status_cf,
                     week_col)

NAME = "Tasks"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("task", '"TSK-"', digits=3, width=9),
        Col("task", "Task", 34, "text"),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("due", "Due date", 12, "date", validate=date_check()),
        week_col("due"),
        Col("priority", "Priority", 11, "text",
            validate=list_from("PriorityList"),
            cf=[{"type": "text", "criteria": "containing", "value": "High",
                 "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Medium", "fmt": "warn"},
                {"type": "text", "criteria": "containing", "value": "Low",
                 "fmt": "muted"}]),
        Col("status", "Status", 14, "status",
            validate=list_from("TaskStatusList"), cf=status_cf()),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("content", "Linked content", 13, "text",
            validate=list_from("ContentIDs")),
        Col("content_title", "Which post", 28, "text",
            formula=lookup_content("title")),
        Col("days_to_go", "Days to go", 10, "int",
            formula='=IF(OR({due}{r}="",{status}{r}="Complete"),"",'
                    '{due}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 2, "fmt": "warn"}]),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "task",
                'IF({status}{r}="Complete","OK",'
                'IF(AND({due}{r}<>"",{due}{r}<TODAY()),"Overdue",'
                'IF({due}{r}=TODAY(),"Today - due now",'
                'IF({status}{r}="Blocked","Action - unblock this",'
                'IF(AND({days_to_go}{r}<>"",{days_to_go}{r}<=2),'
                '"Due - this week","OK")))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Task Management",
        subtitle="Shoots, edits, approvals, replies and the monthly "
                 "analytics entry - everything that is not a post but "
                 "still has to happen.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).tasks,
        help_text="Filter on 'Overdue' in the action column each Monday "
                  "morning - that is your week sorted.",
    )
