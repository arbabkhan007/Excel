"""`Invoices` - what you have billed and what is still owed.

One row per invoice, totalled from its lines, with the due date worked
out from the client's terms and the amount received pulled from the
Payments tab. The overdue flag is the one that gets you paid.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_client,
                     month_col, status_cf)

NAME = "Invoices"

_LINES = "[[Invoice Lines.invoice]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "InvoicePrefix", digits=4, width=11,
               note="The invoice number. Put it on the Invoice Lines "
                    "tab to add lines, and on any payment you "
                    "receive."),
        Col("date", "Invoice date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("client", "Client ID", 11, "id", validate=list_from("ClientIDs"),
            note="Everything else about them - name, address, terms - "
                 "comes from the Clients tab."),
        Col("client_name", "Client", 22, "text",
            formula=lookup_client("name")),
        Col("address", "Billing address", 30, "text",
            formula=lookup_client("address", missing=""), hidden=True),
        Col("period", "Covers", 20, "text",
            note="'Visits in March', 'End of tenancy, 12B Lakeview'. "
                 "What the client will recognise."),
        Col("status", "Status", 13, "status",
            validate=list_from("InvoiceStatusList"), cf=status_cf(),
            note="Draft until you send it. Paid is worked out for "
                 "you from the payments, but you can override it."),
        Col("sent_date", "Sent on", 11, "date", validate=date_check(),
            note="The due date counts from the invoice date, but it "
                 "is worth recording when it really went out."),
        Col("terms_input", "Terms (days)", 10, "int",
            validate=number_check(0, 180, decimal=False),
            note="Blank uses the client's terms, or the business "
                 "default from Settings."),
        Col("discount", "Discount", 9, "percent0",
            validate=number_check(0, 1)),
        Col("method", "Pay by", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("po", "PO / reference", 14, "text",
            note="Commercial clients often will not pay without "
                 "their own reference on the invoice."),
        Col("reminders", "Reminders sent", 10, "int",
            validate=number_check(0, 50, decimal=False)),
        Col("last_reminder", "Last reminder", 11, "date",
            validate=date_check()),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("lines", "Lines", 7, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _LINES + ',{id}{r}))',
            hidden=True),
        Col("subtotal", "Subtotal", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Invoice Lines.total]],' +
                    _LINES + ',{id}{r}))'),
        Col("discount_value", "Less discount", 11, "money",
            formula='=IF({id}{r}="","",ROUND({subtotal}{r}*'
                    'IF({discount}{r}="",0,{discount}{r}),2))',
            hidden=True),
        Col("net", "Net", 12, "money",
            formula='=IF({id}{r}="","",{subtotal}{r}-'
                    '{discount_value}{r})'),
        Col("tax", "Tax", 10, "money",
            formula='=IF({id}{r}="","",ROUND({net}{r}*TaxRate,2))'),
        Col("total", "Total", 12, "money",
            formula='=IF({id}{r}="","",{net}{r}+{tax}{r})',
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}],
            note="What the client owes. Edit the lines, not this "
                 "cell."),
        Col("paid", "Received", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Payments.amount]],'
                    '[[Payments.invoice]],{id}{r}))',
            note="Added up from the Payments tab."),
        Col("outstanding", "Still owed", 12, "money",
            formula='=IF({id}{r}="","",ROUND({total}{r}-{paid}{r},2))',
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}]),
        Col("terms_used", "Terms used", 9, "int",
            formula=blank(
                "date",
                'IF({terms_input}{r}<>"",{terms_input}{r},'
                'IF(IFERROR(INDEX([[Clients.terms]],MATCH({client}{r},'
                '[[Clients.id]],0)),"")<>"",INDEX([[Clients.terms]],'
                'MATCH({client}{r},[[Clients.id]],0)),PaymentTerms))'),
            hidden=True),
        Col("due_date", "Due", 11, "date",
            formula='=IF(OR({id}{r}="",{date}{r}=""),"",'
                    '{date}{r}+{terms_used}{r})'),
        Col("days_late", "Days late", 9, "int",
            formula='=IF(OR({id}{r}="",{due_date}{r}=""),"",'
                    'IF({outstanding}{r}<=0.004,"",'
                    'MAX(0,TodayDate-{due_date}{r}-OverdueGrace)))',
            cf=[{"type": "cell", "criteria": ">", "value": 30,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("age", "Age bracket", 12, "status",
            formula=blank(
                "date",
                'IF({outstanding}{r}<=0.004,"Paid",'
                'IF({days_late}{r}="","Not due yet",'
                'IF({days_late}{r}=0,"Due now",'
                'IF({days_late}{r}<=30,"1-30 days",'
                'IF({days_late}{r}<=60,"31-60 days",'
                'IF({days_late}{r}<=90,"61-90 days","90+ days"))))))'),
            cf=status_cf(),
            note="The ageing buckets the Reports tab counts."),
        Col("late_fee", "Late fee due", 10, "money",
            formula='=IF(OR({id}{r}="",{days_late}{r}=""),"",'
                    'IF({days_late}{r}<=0,0,ROUND({outstanding}{r}*'
                    'LateFeePercent,2)))',
            hidden=True,
            note="Only charge it if your terms say you can."),
        Col("fully_paid", "Settled?", 9, "flag",
            formula=blank("date",
                          'IF({outstanding}{r}<=0.004,"Yes","No")'),
            cf=flag_yes_no_cf(), hidden=True),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "date",
                'IF({lines}{r}=0,"Add - no lines on this invoice yet",'
                'IF({status}{r}="Written off","Written off",'
                'IF({fully_paid}{r}="Yes","OK - paid",'
                'IF({status}{r}="Draft","Draft - not sent yet",'
                'IF(AND({days_late}{r}<>"",{days_late}{r}>30),'
                '"Overdue - "&{days_late}{r}&" days, call them",'
                'IF(AND({days_late}{r}<>"",{days_late}{r}>0),'
                '"Overdue - send a reminder",'
                'IF({paid}{r}>0,"Part paid","Due - waiting for '
                'payment")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Invoices",
        subtitle="What you have billed, what has come in and what is "
                 "late.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).invoices,
        help_text="Raise the invoice here, list the jobs on the "
                  "Invoice Lines tab, and record the money on the "
                  "Payments tab. The due date, the balance and the "
                  "chase list look after themselves.",
        extra_notes=[
            "The fastest way to invoice a month of visits: filter "
            "the Jobs tab to one client and status Completed, then "
            "add one invoice line per job with the job ID.",
            "Once a job appears on an invoice line, the Jobs tab "
            "shows the invoice number against it automatically - so "
            "nothing gets billed twice or forgotten.",
            "The workbook cannot take a payment or email a "
            "reminder. It tells you exactly who to chase and for "
            "how much; the sending is still yours.",
        ],
    )
