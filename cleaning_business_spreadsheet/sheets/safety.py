"""`Safety & Compliance` - the paperwork that has to be in date.

Insurance certificates, background checks, training, risk assessments,
chemical data sheets. Each one has an expiry date and a quiet habit of
lapsing in the week a new commercial client asks to see it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Safety"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("document", '"DOC-"', digits=3, width=10),
        Col("document", "Document", 26, "text",
            validate=list_from("SafetyDocList"),
            note="Pick the type, or type your own. One row per "
                 "document per person."),
        Col("applies_to", "Applies to", 18, "text",
            note="A person, a property, a chemical, or the business "
                 "as a whole."),
        Col("employee", "Person", 16, "text",
            validate=list_from("EmployeeNames"),
            note="Leave blank for a business-wide document."),
        Col("reference", "Reference / number", 20, "text"),
        Col("issuer", "Issued by", 20, "text"),
        Col("issued", "Issued on", 11, "date", validate=date_check()),
        Col("expires", "Expires", 11, "date", validate=date_check(),
            note="The whole point of the tab. Leave blank for a "
                 "document that does not expire."),
        Col("cost", "Cost to renew", 10, "money",
            validate=number_check(0, 1000000)),
        Col("holder", "Where it is kept", 22, "text",
            note="Folder, drive, filing cabinet. 'Somewhere in the "
                 "van' is not an answer."),
        Col("link", "Link", 24, "link",
            note="A link to the scan beats a photocopy in a drawer."),
        Col("required", "Required?", 9, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Some documents are legally required, some are "
                 "asked for by clients, some are neither."),
        Col("reviewed", "Last reviewed", 11, "date",
            validate=date_check(),
            note="Risk assessments and method statements should be "
                 "looked at again every year, not just filed."),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("days_left", "Days to expiry", 11, "int",
            formula='=IF(OR({document}{r}="",{expires}{r}=""),"",'
                    '{expires}{r}-TodayDate)',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=CertWarningDays", "fmt": "warn"}]),
        Col("status", "Status", 14, "status",
            formula=blank(
                "document",
                'IF({expires}{r}="","Valid",'
                'IF({days_left}{r}<0,"Expired",'
                'IF({days_left}{r}<=CertWarningDays,"Due soon",'
                '"Valid")))'),
            cf=status_cf()),
        Col("renew_by", "Start renewing", 11, "date",
            formula='=IF({expires}{r}="","",'
                    '{expires}{r}-CertWarningDays)',
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "document",
                'IF({status}{r}="Expired","Expired - renew it now",'
                'IF({status}{r}="Due soon","Due - renew within "&'
                '{days_left}{r}&" days",'
                'IF(AND({required}{r}="Yes",{link}{r}="",'
                '{holder}{r}=""),"Add - where is this kept?",'
                'IF(AND({reviewed}{r}<>"",TodayDate-{reviewed}{r}>365),'
                '"Due - review it, over a year old","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Safety & Compliance",
        subtitle="Certificates, checks and assessments, with the "
                 "expiry dates watched for you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.grey_text,
        freeze_cols=2,
        samples=sample(cfg).safety,
        help_text="One row per document. Put the expiry date in and "
                  "the workbook warns you well before a commercial "
                  "client asks to see it.",
        extra_notes=[
            "What you are legally required to hold depends "
            "entirely on where you are and who you employ - "
            "insurance, background checks for work around "
            "children, chemical data sheets, electrical testing. "
            "This tab keeps the record; it does not know your "
            "country's rules.",
            "The warning window is on the Settings tab. Six weeks "
            "is usually enough to renew anything without "
            "panicking.",
            "Risk assessments are not a one-off. The review column "
            "flags anything you have not looked at for a year.",
        ],
    )
