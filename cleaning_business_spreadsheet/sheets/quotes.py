"""`Quotes` - one row per quote, totalled from its lines.

You write the lines on the Quote Lines tab; this tab adds them up,
applies the discount and the tax, works out the deposit, counts the
days since you sent it and tells you when to follow up. The win rate
on the Reports tab is counted from the Status column here.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, month_col,
                     status_cf)

NAME = "Quotes"

_LINES = "[[Quote Lines.quote]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "QuotePrefix", digits=4, width=10,
               note="Quote number. Put it on the Quote Lines tab to "
                    "add lines to this quote."),
        Col("date", "Quote date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("client", "Client ID", 11, "id", validate=list_from("ClientIDs"),
            note="An existing client. For a brand new enquiry, leave "
                 "it blank and type the name in the next column."),
        Col("client_name", "Client", 22, "text",
            formula='=IF({client}{r}<>"",IFERROR(INDEX([[Clients.name]],'
                    'MATCH({client}{r},[[Clients.id]],0)),'
                    '{prospect}{r}),{prospect}{r})',
            note="The client's name, or the prospect name if they are "
                 "not on the books yet."),
        Col("prospect", "Prospect name", 20, "text",
            note="For a quote to somebody who is not a client yet."),
        Col("property", "Property ID", 11, "id",
            validate=list_from("PropertyIDs")),
        Col("address", "Address", 30, "wrap"),
        Col("service", "Main service", 22, "text",
            validate=list_from("ServiceNames")),
        Col("frequency", "Frequency", 14, "text",
            validate=list_from("FrequencyList")),
        Col("status", "Status", 13, "status",
            validate=list_from("QuoteStatusList"), cf=status_cf(),
            note="Draft, Sent, Followed up, Accepted, Declined. The "
                 "win rate counts Accepted against everything sent."),
        Col("sent_date", "Sent on", 11, "date", validate=date_check(),
            note="The follow-up clock starts here, not on the quote "
                 "date."),
        Col("followed_up", "Last followed up", 12, "date",
            validate=date_check()),
        Col("decision_date", "Accepted / declined on", 12, "date",
            validate=date_check()),
        Col("discount", "Discount", 10, "percent0",
            validate=number_check(0, 1),
            note="A percentage off the lines. First clean discounts "
                 "win a lot of recurring work."),
        Col("deposit_input", "Deposit asked for", 11, "money",
            validate=number_check(0, 1000000),
            note="Blank uses the deposit percentage from Settings."),
        Col("valid_input", "Valid until", 11, "date",
            validate=date_check(),
            note="Blank adds the standard validity from Settings to "
                 "the quote date."),
        Col("job", "Job booked", 11, "id", validate=list_from("JobIDs"),
            note="Once they say yes, book the job and put its ID "
                 "here."),
        Col("decline_reason", "If declined, why", 18, "text",
            validate=list_from("LostReasonList")),
        Col("notes", "Notes", 34, "wrap"),
        Col("terms", "Terms shown", 34, "wrap",
            note="What is included, what is not, cancellation terms. "
                 "This is what prints on the quote."),
        # ------------------------------------------------- the automatics --
        Col("lines", "Lines", 8, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _LINES + ',{id}{r}))',
            hidden=True),
        Col("hours", "Hours quoted", 11, "number1",
            formula='=IF({id}{r}="","",SUMIFS([[Quote Lines.hours]],' +
                    _LINES + ',{id}{r}))'),
        Col("subtotal", "Subtotal", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Quote Lines.total]],' +
                    _LINES + ',{id}{r}))',
            note="Added up from the Quote Lines tab."),
        Col("discount_value", "Less discount", 11, "money",
            formula='=IF({id}{r}="","",ROUND({subtotal}{r}*'
                    'IF({discount}{r}="",0,{discount}{r}),2))'),
        Col("net", "Net", 12, "money",
            formula='=IF({id}{r}="","",{subtotal}{r}-'
                    '{discount_value}{r})'),
        Col("tax", "Tax", 10, "money",
            formula='=IF({id}{r}="","",ROUND({net}{r}*TaxRate,2))'),
        Col("total", "Total", 12, "money",
            formula='=IF({id}{r}="","",{net}{r}+{tax}{r})',
            note="What the client sees. Price the lines, not this "
                 "cell.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("deposit", "Deposit", 11, "money",
            formula='=IF({id}{r}="","",IF({deposit_input}{r}<>"",'
                    '{deposit_input}{r},ROUND({total}{r}*'
                    'DepositPercent,2)))'),
        Col("cost", "Cost to deliver", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Quote Lines.cost]],' +
                    _LINES + ',{id}{r}))',
            hidden=True),
        Col("profit", "Profit", 11, "money",
            formula='=IF({id}{r}="","",{net}{r}-{cost}{r})'),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({id}{r}="",{net}{r}=0),"",'
                    '{profit}{r}/{net}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}],
            note="Check this before you send it. A discount that "
                 "drops the margin under the floor is work you "
                 "cannot afford to win."),
        Col("valid_until", "Valid to", 11, "date",
            formula='=IF(OR({id}{r}="",{date}{r}=""),"",'
                    'IF({valid_input}{r}<>"",{valid_input}{r},'
                    '{date}{r}+QuoteValidDays))'),
        Col("days_waiting", "Days waiting", 11, "int",
            formula='=IF(OR({id}{r}="",{sent_date}{r}=""),"",'
                    'IF(OR({status}{r}="Accepted",{status}{r}="Declined"),'
                    '"",TodayDate-{sent_date}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 7,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 3,
                 "fmt": "warn"}]),
        Col("expired", "Expired?", 9, "flag",
            formula='=IF(OR({id}{r}="",{valid_until}{r}=""),"",'
                    'IF(AND({valid_until}{r}<TodayDate,{status}{r}<>'
                    '"Accepted",{status}{r}<>"Declined"),"Yes","No"))',
            cf=flag_yes_no_cf(), hidden=True),
        Col("won", "Won?", 8, "flag",
            formula=blank("date",
                          'IF({status}{r}="Accepted","Yes",'
                          'IF({status}{r}="Declined","No",""))'),
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({status}{r}="Accepted",IF({job}{r}="",'
                '"Action - accepted, book the job","OK - won"),'
                'IF({status}{r}="Declined","Closed",'
                'IF({lines}{r}=0,"Add - no lines on this quote yet",'
                'IF({status}{r}="Draft","Draft - not sent",'
                'IF({expired}{r}="Yes","Expired - re-quote or close",'
                'IF(AND({days_waiting}{r}<>"",'
                '{days_waiting}{r}>=QuoteFollowupDays),'
                '"Chase - "&{days_waiting}{r}&" days with no answer",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Quotes & Estimates",
        subtitle="Every quote you have sent, what it is worth, what "
                 "it makes you and when to chase it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).quotes,
        help_text="Raise the quote here, then list what is included on "
                  "the Quote Lines tab using this quote's ID. The "
                  "totals, the deposit and the margin all appear by "
                  "themselves.",
        extra_notes=[
            "Use the Pricing & Quote Calculator tab to work out the "
            "price first, then write the quote here. The calculator "
            "shows you the margin before you commit to a number.",
            "A quote with no lines totals zero - that is what the "
            "'Add - no lines' warning means.",
            "A spreadsheet cannot collect a signature. Print the "
            "quote, email it, and record the acceptance date here "
            "when they reply; the email itself is your evidence.",
        ],
    )
