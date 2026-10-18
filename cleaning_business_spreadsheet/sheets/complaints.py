"""`Complaints` - what went wrong and what you did about it.

Nobody enjoys this tab. It is also the one that keeps clients: a
complaint handled in two days is a loyal customer, and a complaint
nobody wrote down happens again next month.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_client,
                     month_col, score_validation, status_cf)

NAME = "Complaints"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "ComplaintPrefix", digits=4, width=11),
        Col("date", "Date raised", 11, "date", validate=date_check()),
        month_col("date"),
        Col("client", "Client ID", 11, "id", validate=list_from("ClientIDs")),
        Col("client_name", "Client", 20, "text",
            formula=lookup_client("name")),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Which visit it relates to, if you know."),
        Col("employee", "Cleaner on the job", 16, "text",
            validate=list_from("EmployeeNames"),
            note="Record it, but treat a single complaint as "
                 "information rather than evidence."),
        Col("type", "What about", 20, "text",
            validate=list_from("ComplaintTypeList")),
        Col("severity", "Severity", 11, "status",
            validate=list_from("SeverityList"), cf=status_cf()),
        Col("description", "What they said", 38, "wrap",
            note="In their words where you can. It reads very "
                 "differently from yours."),
        Col("status", "Status", 15, "status",
            validate=list_from("ComplaintStatusList"), cf=status_cf()),
        Col("owner", "Being handled by", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("action", "What we did", 36, "wrap"),
        Col("resolution", "Resolution", 18, "text",
            validate=list_from("ResolutionList")),
        Col("reclean_job", "Re-clean job", 11, "id",
            validate=list_from("JobIDs"),
            note="If you went back, book it as a job with type "
                 "'Re-clean (free)' and link it here."),
        Col("cost", "Cost to put right", 11, "money",
            validate=number_check(0, 1000000),
            note="Refund, free re-clean, replacement. Worth knowing "
                 "- complaints are a real line in your costs."),
        Col("resolved_date", "Resolved on", 11, "date",
            validate=date_check()),
        Col("kept", "Client kept?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("rating_after", "Rating afterwards", 10, "rating",
            validate=score_validation(5)),
        Col("prevent", "How we stop it happening again", 34, "wrap",
            note="The most valuable column on the tab, and the one "
                 "everybody leaves blank."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("days_open", "Days open", 10, "int",
            formula='=IF(OR({date}{r}="",{client}{r}=""),"",'
                    'IF({resolved_date}{r}<>"",'
                    '{resolved_date}{r}-{date}{r},TodayDate-{date}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 7,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 2,
                 "fmt": "warn"}]),
        Col("open", "Still open?", 9, "flag",
            formula=blank("date",
                          'IF(OR({status}{r}="Resolved",'
                          '{status}{r}="Closed - no action",'
                          '{status}{r}="Refunded"),"No","Yes")'),
            cf=flag_yes_no_cf(), hidden=True),
        Col("repeat", "Repeat issue?", 10, "flag",
            formula=blank(
                "date",
                'IF(COUNTIFS({client_first}:{client_last},{client}{r},'
                '{type_first}:{type_last},{type}{r})>1,"Yes","No")'),
            cf=flag_yes_no_cf(),
            note="The same client complaining about the same thing "
                 "twice. Fix the cause, not the visit."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({open}{r}="No","OK - closed",'
                'IF({days_open}{r}>7,"Overdue - open "&{days_open}{r}&'
                '" days",'
                'IF({severity}{r}="Critical","Action - critical, today",'
                'IF({owner}{r}="","Assign - nobody owns this",'
                'IF({repeat}{r}="Yes","Check - repeat issue",'
                '"Action - open")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Complaints & Quality Control",
        subtitle="What went wrong, how fast you fixed it and whether "
                 "you kept the client.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).complaints,
        help_text="Log every complaint, even the small ones you "
                  "sorted out on the phone. The pattern is the point, "
                  "not the individual row.",
        extra_notes=[
            "Days open is the number that matters. Research after "
            "research says the same thing: it is not the mistake "
            "that loses the client, it is the silence afterwards.",
            "A free re-clean has a real cost - the wages, the "
            "products and the hour you could have sold to somebody "
            "else. Put it in the cost column and the picture gets "
            "clearer.",
            "Treat the cleaner column carefully. One complaint is "
            "noise; three in a month against the same person is a "
            "training conversation.",
        ],
    )
