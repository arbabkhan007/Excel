"""`Allergy & Medical` - conditions, plans, consents and the dose log.

Two kinds of row live here, and the record type tells them apart.

A **profile** row is a standing fact: this child has a peanut allergy,
this child has an inhaler, this child must not have dairy.  It carries
the severity, the reaction, what to do, who prescribed it and when the
plan or the medicine runs out.

An **administration** row is an event: at 2:15pm on Tuesday, this
member of staff gave this dose.  Keeping both in one register means a
dose is always filed next to the plan it came from, and the expiry
dates are never in a different file from the medicine.

This is a working record for your own staff.  It is not a medical
record, it does not replace a medication protocol, and it should be
stored somewhere only the people who need it can open.
"""

from __future__ import annotations

from ..config import MEDICAL_EVENT_TYPES, Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, time_check
from .common import (alert_cf, blank, id_col, lookup_child,
                     lookup_staff, seq_key, severity_cf, status_cf)

NAME = "Allergy & Medical"


def spec(cfg: Config) -> SheetSpec:
    event_test = "+".join(f'({{record_type}}{{r}}="{t}")'
                          for t in MEDICAL_EVENT_TYPES)

    columns = [
        id_col("child", "MedicalPrefix", digits=4, width=11,
               header="Record ID"),
        Col("child", "Child", 11, "id", validate=list_from("ChildList"),
            note="Which child this record belongs to. Everything else "
                 "on the row describes them."),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("room", "Room", 13, "text",
            formula=lookup_child("room", missing=""), hidden=True),
        *seq_key("child"),
        Col("record_type", "Record type", 17, "text",
            validate=list_from("MedicalRecordList"), cf=status_cf(),
            note="Allergy, condition, medication, immunisation - or "
                 "'Medication given' to log a single dose."),
        Col("is_event", "Dose log?", 8, "flag",
            formula=blank("record_type",
                          f'IF(({event_test})>0,"Yes","No")'),
            hidden=True,
            note="Yes on the rows that record one dose rather than a "
                 "standing instruction."),
        Col("detail", "Detail", 24, "text",
            note="Peanuts. Asthma. Lactose. Amoxicillin. Whatever the "
                 "record is actually about - one thing per row."),
        Col("severity", "Severity", 15, "status",
            validate=list_from("AllergySeverityList"), cf=severity_cf(),
            note="Drives the warning on the child's record, on every "
                 "meal and on the dashboard. 'Anaphylaxis risk' turns "
                 "everything red, which is the point."),
        Col("reaction", "What happens", 28, "wrap",
            note="What a reaction looks like, in the words a new "
                 "member of staff would recognise it by."),
        Col("action", "What to do", 34, "wrap",
            note="The steps, in order, including when to call an "
                 "ambulance. Short sentences. This is read in a hurry."),
        Col("medication", "Medication", 20, "text"),
        Col("dose", "Dose & route", 20, "text",
            note="Exactly as the prescription says it. Never as "
                 "anybody remembers it."),
        Col("stored", "Stored where", 18, "text",
            note="Red box in the office. Fridge, top shelf. The "
                 "sentence that saves two minutes in an emergency."),
        Col("trained", "Staff trained", 18, "text", hidden=True,
            note="Who is trained to give this. Names, not job titles."),
        # ---------------------------------------------- authorisation ---
        Col("consent", "Parent consent", 12, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}],
            note="Written consent on file. Without it, nothing is "
                 "given - and the Action column says so."),
        Col("consent_date", "Consent dated", 12, "date",
            validate=date_check(), hidden=True),
        Col("prescriber", "Prescribed by", 18, "text", hidden=True),
        Col("doctor", "Doctor / clinic", 20, "text"),
        Col("doctor_phone", "Doctor's phone", 14, "text"),
        Col("reviewed", "Plan dated", 11, "date",
            validate=date_check(),
            note="When the plan was last written or reviewed with the "
                 "parent."),
        Col("expires", "Expires", 11, "date", validate=date_check(),
            note="The medicine's expiry date, or the date the plan "
                 "needs reviewing. Whichever comes first."),
        Col("days_left", "Days left", 10, "int",
            formula=blank(
                "expires", '{expires}{r}-TodayDate'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocUrgentDays", "fmt": "bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocWarningDays", "fmt": "warn"}]),
        # ------------------------------------------------- the dose log -
        Col("given_date", "Given on", 11, "date",
            validate=date_check(),
            note="Only on a 'Medication given' row. Leave blank on a "
                 "standing record."),
        Col("given_time", "Given at", 10, "time",
            validate=time_check()),
        Col("given_by", "Given by", 11, "id",
            validate=list_from("StaffList"),
            note="The member of staff who administered it. Required "
                 "on every dose row - that is the whole point of the "
                 "log."),
        Col("given_name", "Staff", 16, "text",
            formula=lookup_staff("name", "given_by")),
        Col("witness", "Witnessed by", 14, "text", hidden=True,
            note="Many settings require a second pair of eyes. If "
                 "yours does, this is where it goes."),
        Col("doses_logged", "Doses logged", 11, "int",
            formula=blank(
                "child",
                'IF({is_event}{r}="Yes","",'
                'COUNTIFS({child_first}:{child_last},{child}{r},'
                '{detail_first}:{detail_last},{detail}{r},'
                '{is_event_first}:{is_event_last},"Yes"))'),
            hidden=True,
            note="On a standing record, how many dose rows exist for "
                 "the same child and the same medicine."),
        Col("last_dose", "Last dose", 11, "date",
            formula=blank(
                "child",
                'IF({is_event}{r}="Yes","",'
                'IF({doses_logged}{r}=0,"",'
                'SUMPRODUCT(MAX(([[Allergy & Medical.child]]='
                '{child}{r})*([[Allergy & Medical.detail]]='
                '{detail}{r})*([[Allergy & Medical.is_event]]="Yes")*'
                '[[Allergy & Medical.given_date]]))))'),
            hidden=True),
        Col("notes", "Notes", 28, "wrap"),
        Col("alert", "Action", 34, "status",
            formula=blank(
                "child",
                'IF({child_name}{r}="Not in Child Enrollment",'
                '"Check - child ID not found",'
                'IF({record_type}{r}="","Add - what kind of record is '
                'this?",'
                'IF(IF({is_event}{r}<>"Yes",FALSE,{given_by}{r}=""),'
                '"Missing - no staff name against this dose",'
                'IF(IF({is_event}{r}<>"Yes",FALSE,{given_date}{r}=""),'
                '"Missing - no date against this dose",'
                'IF(IF({medication}{r}="",FALSE,{consent}{r}<>"Yes"),'
                '"Not authorised - no written parent consent",'
                'IF(IF({expires}{r}="",FALSE,{days_left}{r}<0),'
                '"Expired - renew it before it is needed",'
                'IF(IF({expires}{r}="",FALSE,'
                '{days_left}{r}<=DocUrgentDays),'
                '"Expiring - "&{days_left}{r}&" days left",'
                'IF(IF(({severity}{r}="Severe")+'
                '({severity}{r}="Anaphylaxis risk")>0,'
                '{action}{r}="",FALSE),'
                '"Missing - severe allergy with no written plan",'
                'IF(IF({expires}{r}="",FALSE,'
                '{days_left}{r}<=DocWarningDays),'
                '"Due - review the plan with the parent","OK")))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Allergy, Medical & Medication Records",
        subtitle="Conditions, allergy plans, consents, expiry dates "
                 "and the dose log - one child's record in one place.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.bad_text,
        freeze_cols=3,
        samples=sample(cfg).medical,
        help_text="One row per condition, allergy or medication. To "
                  "log a dose, add a row with record type 'Medication "
                  "given' - it files itself next to the plan.",
        extra_notes=[
            "This is a working note for your staff. It is not a "
            "medical record and it does not replace your medication "
            "protocol, your allergy plans or your emergency "
            "procedures. Keep the signed originals where your policy "
            "says they live.",
            "Severity is the column that travels. 'Severe' or "
            "'Anaphylaxis risk' here turns the child's enrollment "
            "record red and puts a warning on every meal they are "
            "served.",
            "This tab holds sensitive personal data about children. "
            "Restrict who can open this file, keep it encrypted or "
            "access controlled, and delete records in line with your "
            "retention policy.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Medical register summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Medical register at a glance", last)
    sheet.table_head(row, 0, [
        "Records", "Children with a record", "Allergies",
        "Anaphylaxis risk", "Severe", "On medication",
        "Missing consent", "Plans expired", "Expiring soon",
        "Doses logged"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Allergy & Medical.child]],"<>")',
        '=COUNTIFS([[Child Enrollment.allergy_count]],">0")',
        '=COUNTIFS([[Allergy & Medical.record_type]],"Allergy")',
        '=COUNTIFS([[Allergy & Medical.severity]],"Anaphylaxis risk")',
        '=COUNTIFS([[Allergy & Medical.severity]],"Severe")',
        '=COUNTIF([[Allergy & Medical.medication]],"<>")',
        '=COUNTIFS([[Allergy & Medical.medication]],"<>",'
        '[[Allergy & Medical.consent]],"<>Yes")',
        '=COUNTIFS([[Allergy & Medical.days_left]],"<0")',
        '=COUNTIFS([[Allergy & Medical.days_left]],">=0",'
        '[[Allergy & Medical.days_left]],"<="&DocWarningDays)',
        '=COUNTIFS([[Allergy & Medical.is_event]],"Yes")',
    ], ["int", "int", "int", "int", "int", "int", "int", "int", "int",
        "int"])
