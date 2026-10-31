"""`Document Tracker` - the paperwork, and the day each piece expires.

Children have enrollment agreements and immunisation records.  Staff
have background checks and first aid certificates.  The business has a
licence, an insurance policy and a fire inspection.  All of it expires,
and all of it is the same problem: a date, a holder and a status.

So it is one register.  The holder column says whether the row is
about a child, a family, a member of staff or the business itself, and
the name resolves itself from whichever register the ID belongs to.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, id_col, status_cf)

NAME = "Document Tracker"


def spec(cfg: Config) -> SheetSpec:
    # The ID can belong to any of three registers, so the name lookup
    # tries each in turn and falls back to whatever was typed.
    who_name = (
        'IFERROR(INDEX([[Child Enrollment.name]],'
        'MATCH({who}{r},[[Child Enrollment.id]],0))&"",'
        'IFERROR(INDEX([[Staff Database.name]],'
        'MATCH({who}{r},[[Staff Database.id]],0))&"",'
        'IFERROR(INDEX([[Parent Database.family_name]],'
        'MATCH({who}{r},[[Parent Database.id]],0))&"",'
        '{who}{r})))')

    columns = [
        id_col("doc_type", "DocumentPrefix", digits=4, width=10,
               header="Doc ID"),
        Col("holder", "Belongs to", 13, "text",
            validate=list_from("DocumentHolderList"), cf=status_cf(),
            note="Child, family, staff or the business itself."),
        Col("who", "ID", 11, "id",
            note="The child ID, staff ID or family ID. For a business "
                 "document - a licence, an insurance policy - leave "
                 "it blank or type the setting's name."),
        Col("who_name", "Name", 22, "text",
            formula=blank("who", who_name),
            note="Looked up from whichever register the ID belongs "
                 "to. Nothing to type."),
        Col("doc_type", "Document", 26, "text",
            validate=list_from("DocumentTypeList"), cf=status_cf(),
            note="Pick from the list. Add your own at the bottom of "
                 "the list on Settings if your regulator wants "
                 "something this one does not have."),
        Col("required", "Required?", 10, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "info"}],
            note="Only required documents are counted in the "
                 "'missing paperwork' figures on the child and staff "
                 "tabs."),
        Col("status", "Status", 13, "status",
            validate=list_from("DocumentStatusList"), cf=status_cf()),
        Col("requested", "Requested", 11, "date",
            validate=date_check(), hidden=True),
        Col("received", "Received", 11, "date", validate=date_check()),
        Col("expires", "Expires", 11, "date", validate=date_check(),
            note="Leave blank for anything that does not expire - an "
                 "enrollment agreement, a birth certificate sighting."),
        Col("days_left", "Days left", 9, "int",
            formula=blank("expires", '{expires}{r}-TodayDate'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocUrgentDays", "fmt": "bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocWarningDays", "fmt": "warn"},
                {"type": "cell", "criteria": ">",
                 "value": "=DocWarningDays", "fmt": "good"}]),
        Col("state", "Where it stands", 14, "status",
            formula=blank(
                "doc_type",
                'IF({required}{r}<>"Yes","Not required",'
                'IF(({status}{r}="On file")+({status}{r}="Signed")+'
                '({status}{r}="Received")=0,"Missing",'
                'IF({expires}{r}="","On file",'
                'IF({days_left}{r}<0,"Expired",'
                'IF({days_left}{r}<=DocWarningDays,"Expiring",'
                '"On file")))))'),
            cf=status_cf(),
            note="One word for the whole row: Missing, Expired, "
                 "Expiring, On file or Not required. This is the "
                 "column the dashboard counts."),
        Col("renewed", "Last renewed", 11, "date",
            validate=date_check(), hidden=True),
        Col("renew_every", "Renew every (months)", 12, "int",
            hidden=True,
            note="For anything on a cycle. Only here so you can "
                 "filter by it - the workbook does not auto-renew."),
        Col("owner", "Chased by", 10, "id",
            validate=list_from("StaffList"), hidden=True,
            note="Whose job it is to get this one back."),
        Col("location", "Filed where", 20, "text",
            note="Cabinet 2, child files. Shared drive, Staff/HR. "
                 "A pointer, not the document."),
        Col("reference", "Reference", 14, "text", hidden=True,
            note="Policy number, certificate number, licence number."),
        Col("notes", "Notes", 26, "wrap"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "doc_type",
                'IF({required}{r}<>"Yes","",'
                'IF({state}{r}="Missing",'
                'IF({status}{r}="Chased","Chase - asked for it '
                'already, ask again",'
                'IF({status}{r}="Requested","Chase - requested, not '
                'back yet","Missing - request it")),'
                'IF({state}{r}="Expired","Expired - '
                '"&ABS({days_left}{r})&" days ago, renew now",'
                'IF(IF({expires}{r}="",FALSE,'
                '{days_left}{r}<=DocUrgentDays),'
                '"Expiring - "&{days_left}{r}&" days left",'
                'IF(IF({expires}{r}="",FALSE,'
                '{days_left}{r}<=DocWarningDays),'
                '"Due - book the renewal",'
                'IF(IF({who}{r}="",FALSE,{who_name}{r}={who}{r}),'
                '"Check - that ID is not in any register",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Document & Renewal Tracker",
        subtitle="Children's forms, staff certificates and the "
                 "setting's own licences - with the day each one "
                 "runs out.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).documents,
        help_text="One row per document. Say who it belongs to, "
                  "whether it is required and when it expires - the "
                  "status, the warnings and the missing-paperwork "
                  "counts all follow.",
        extra_notes=[
            "The name column resolves the ID against the children, "
            "the staff and the families in turn, so one register "
            "holds every kind of document without you having to "
            "choose a tab.",
            "Only rows marked 'Required - Yes' are counted in the "
            "missing paperwork figures on the child and staff "
            "records. Mark the optional ones No and they stop "
            "nagging.",
            "This tracks your paperwork. It does not make you "
            "compliant with anybody's licensing rules, and the "
            "document list is a starting point to edit rather than a "
            "legal checklist.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Paperwork summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Paperwork at a glance", last)
    sheet.table_head(row, 0, [
        "Documents tracked", "Required", "On file", "Missing",
        "Expired", "Expiring soon", "Child documents",
        "Staff documents", "Business documents", "Not required"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Document Tracker.doc_type]],"<>")',
        '=COUNTIFS([[Document Tracker.required]],"Yes")',
        '=COUNTIFS([[Document Tracker.state]],"On file")',
        '=COUNTIFS([[Document Tracker.state]],"Missing")',
        '=COUNTIFS([[Document Tracker.state]],"Expired")',
        '=COUNTIFS([[Document Tracker.state]],"Expiring")',
        '=COUNTIFS([[Document Tracker.holder]],"Child")',
        '=COUNTIFS([[Document Tracker.holder]],"Staff")',
        '=COUNTIFS([[Document Tracker.holder]],"Business")',
        '=COUNTIFS([[Document Tracker.state]],"Not required")',
    ], ["int", "int", "int", "int", "int", "int", "int", "int", "int",
        "int"])
