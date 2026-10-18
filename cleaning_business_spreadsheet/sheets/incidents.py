"""`Incidents` - accidents, near misses and damage.

Short, dull and important. Cleaning involves wet floors, chemicals,
ladders and other people's property, so things occasionally go wrong.
A written record protects your staff, your client and you.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Incidents"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"INC-"', digits=3, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("time", "Time", 9, "text"),
        Col("type", "What happened", 22, "text",
            validate=list_from("IncidentTypeList")),
        Col("employee", "Person involved", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs")),
        Col("property", "Property ID", 11, "id",
            validate=list_from("PropertyIDs")),
        Col("severity", "Severity", 11, "status",
            validate=list_from("SeverityList"), cf=status_cf()),
        Col("description", "What happened, in full", 40, "wrap",
            note="Write it the same day, in plain language, "
                 "including what was being done at the time."),
        Col("injury", "Injury?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("first_aid", "First aid given?", 10, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf()),
        Col("medical", "Medical help needed?", 10, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf()),
        Col("damage", "Property damaged", 26, "wrap"),
        Col("cost", "Cost", 11, "money", validate=number_check(0, 1000000)),
        Col("insurance", "Insurance claim?", 10, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf()),
        Col("reported_to", "Reported to", 18, "text",
            note="Client, insurer, the authority your country "
                 "requires. Write down who and when."),
        Col("reported_date", "Reported on", 11, "date",
            validate=date_check()),
        Col("action", "Action taken", 34, "wrap"),
        Col("prevent", "How we prevent a repeat", 32, "wrap"),
        Col("status", "Status", 13, "status",
            validate=list_from("ComplaintStatusList"), cf=status_cf()),
        Col("closed_date", "Closed on", 11, "date", validate=date_check()),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("days_open", "Days open", 10, "int",
            formula='=IF({date}{r}="","",IF({closed_date}{r}<>"",'
                    '{closed_date}{r}-{date}{r},TodayDate-{date}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 14,
                 "fmt": "bad"}]),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "date",
                'IF({closed_date}{r}<>"","OK - closed",'
                'IF(AND({injury}{r}="Yes",{reported_date}{r}=""),'
                '"Action - injury not reported yet",'
                'IF({severity}{r}="Critical","Action - critical",'
                'IF({prevent}{r}="","Add - what stops it happening '
                'again?","Action - open"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Incidents & Near Misses",
        subtitle="Accidents, damage and the things that almost went "
                 "badly wrong.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.grey_text,
        freeze_cols=2,
        samples=sample(cfg).incidents,
        help_text="Record it the same day, including near misses. "
                  "The near misses are the free lessons.",
        extra_notes=[
            "Reporting duties differ from country to country and "
            "some injuries must be reported to an authority within "
            "days. Find out what applies where you are - this "
            "workbook keeps the record, it does not know your law.",
            "Damage to a client's property is better reported by "
            "you, with a photo, than discovered by them a week "
            "later.",
        ],
    )
