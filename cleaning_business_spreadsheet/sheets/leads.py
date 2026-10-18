"""`Leads` - every enquiry, from first phone call to won or lost.

The pipeline tab. Most cleaning businesses lose more work to a quote
nobody chased than to a price that was too high, so this tab counts the
days an enquiry has been sitting and tells you who to ring today.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, status_cf

NAME = "Leads"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "LeadPrefix", digits=4, width=10),
        Col("date", "Enquiry date", 12, "date", validate=date_check(),
            note="The day they got in touch. The clock starts here."),
        Col("name", "Name / company", 24, "text"),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 24, "text"),
        Col("address", "Address or area", 28, "wrap"),
        Col("type", "Client type", 16, "text",
            validate=list_from("ClientTypeList")),
        Col("source", "Came from", 17, "text",
            validate=list_from("LeadSourceList"),
            note="Fill this in every single time. It is the only way "
                 "to find out which advertising actually works."),
        Col("service", "Interested in", 22, "text",
            validate=list_from("ServiceNames")),
        Col("frequency", "Frequency wanted", 14, "text",
            validate=list_from("FrequencyList")),
        Col("stage", "Stage", 14, "status",
            validate=list_from("LeadStageList"), cf=status_cf(),
            note="Move it along as you go. Won creates a client; Lost "
                 "wants a reason in the next column."),
        Col("value", "Estimated value a year", 13, "money",
            validate=number_check(0, 1000000),
            note="Rough annual value. Sort on it to decide who to "
                 "chase first."),
        Col("quote", "Quote sent", 11, "id", validate=list_from("QuoteIDs"),
            note="Link the quote once you have raised it on the "
                 "Quotes tab."),
        Col("visit_date", "Visit booked", 11, "date",
            validate=date_check()),
        Col("contacted", "Last contacted", 12, "date",
            validate=date_check(),
            note="Update it every time you speak to them. The chase "
                 "list counts from this date."),
        Col("next_action", "Next action", 24, "text",
            note="'Call Tuesday after 6', 'send revised price'. "
                 "Specific beats vague."),
        Col("next_date", "Next action due", 12, "date",
            validate=date_check()),
        Col("client", "Became client", 11, "id",
            validate=list_from("ClientIDs"),
            note="Once they book, add them to the Clients tab and put "
                 "their ID here."),
        Col("lost_reason", "If lost, why", 20, "text",
            validate=list_from("LostReasonList"),
            note="Be honest with yourself here. Three 'price too "
                 "high' in a row is information."),
        Col("notes", "Notes", 36, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("days_open", "Days open", 10, "int",
            formula='=IF(OR({name}{r}="",{date}{r}=""),"",'
                    'IF(OR({stage}{r}="Won",{stage}{r}="Lost",'
                    '{stage}{r}="Not a fit"),"",TodayDate-{date}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 14,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 7,
                 "fmt": "warn"}]),
        Col("days_quiet", "Days since anything happened", 11, "int",
            formula='=IF({name}{r}="","",TodayDate-MAX('
                    'IF({date}{r}="",0,{date}{r}),'
                    'IF({visit_date}{r}="",0,{visit_date}{r}),'
                    'IF({next_date}{r}="",0,IF({next_date}{r}>TodayDate,'
                    '0,{next_date}{r})),'
                    'IF({quote}{r}="",0,IFERROR(INDEX('
                    '[[Quotes.sent_date]],MATCH({quote}{r},'
                    '[[Quotes.id]],0)),0))))',
            hidden=True,
            note="Days since the enquiry, the visit, the quote going "
                 "out or the last action - whichever is latest."),
        Col("quote_value", "Quote total", 12, "money",
            formula='=IF({quote}{r}="","",IFERROR(INDEX('
                    '[[Quotes.total]],MATCH({quote}{r},[[Quotes.id]],0)),'
                    '""))'),
        Col("quote_status", "Quote status", 13, "status",
            formula='=IF({quote}{r}="","",IFERROR(INDEX('
                    '[[Quotes.status]],MATCH({quote}{r},[[Quotes.id]],0)),'
                    '""))',
            cf=status_cf()),
        Col("won", "Won?", 8, "flag",
            formula=blank("name",
                          'IF({stage}{r}="Won","Yes",'
                          'IF({stage}{r}="Lost","No",""))'),
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({stage}{r}="Won","OK - won",'
                'IF(OR({stage}{r}="Lost",{stage}{r}="Not a fit"),'
                '"Closed",'
                'IF(AND({next_date}{r}<>"",{next_date}{r}<TodayDate),'
                '"Overdue - action was due "&TEXT({next_date}{r},'
                '"d mmm"),'
                'IF(AND({next_date}{r}<>"",{next_date}{r}=TodayDate),'
                '"Today - "&{next_action}{r},'
                'IF(AND({quote}{r}<>"",{days_quiet}{r}<>"",'
                '{days_quiet}{r}>=QuoteFollowupDays),'
                '"Chase - quote sent, no answer",'
                'IF(AND({quote}{r}="",{days_open}{r}<>"",'
                '{days_open}{r}>=2),"Action - no quote sent yet",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Leads & Sales Pipeline",
        subtitle="Every enquiry, how long it has been waiting and who "
                 "needs a phone call today.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).leads,
        help_text="Log the enquiry the day it arrives, even the ones "
                  "you are not sure about. Set a next action and a "
                  "date, and the Action column becomes your call list.",
        extra_notes=[
            "The win rate on the Reports tab is counted from the "
            "Stage column, so keep it honest - a lead parked at "
            "'Following up' for three months is a Lost.",
            "'Came from' is worth more than it looks. After fifty "
            "enquiries you will know exactly which of your adverts "
            "is paying for itself.",
            "Nothing here emails or texts the client. The workbook "
            "tells you who to contact; the phone call is still "
            "yours to make.",
        ],
    )
