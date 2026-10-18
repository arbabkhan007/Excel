"""`Payments` - the money that actually arrived.

Every payment against an invoice. The invoice balance, the client's
account and the cash figures on the dashboard all come from here, so
one row per payment - including part payments - keeps everything
honest.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, month_col

NAME = "Payments"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "PaymentPrefix", digits=4, width=11),
        Col("date", "Date received", 12, "date", validate=date_check(),
            note="The day the money landed, not the day you invoiced."),
        month_col("date"),
        Col("invoice", "Invoice ID", 11, "id",
            validate=list_from("InvoiceIDs"),
            note="Which invoice this pays. Part payments get their "
                 "own row against the same invoice."),
        Col("client", "Client ID", 11, "id",
            formula='=IF({invoice}{r}="","",IFERROR(INDEX('
                    '[[Invoices.client]],MATCH({invoice}{r},'
                    '[[Invoices.id]],0)),""))',
            note="Pulled from the invoice."),
        Col("client_name", "Client", 22, "text",
            formula='=IF({invoice}{r}="","",IFERROR(INDEX('
                    '[[Invoices.client_name]],MATCH({invoice}{r},'
                    '[[Invoices.id]],0)),""))'),
        Col("amount", "Amount", 12, "money",
            validate=number_check(0, 10000000),
            note="What arrived in the bank. If a card processor took "
                 "a fee, put the fee in the next column and this "
                 "stays the gross amount."),
        Col("fee", "Processing fee", 10, "money",
            validate=number_check(0, 100000),
            note="Card or platform fee. It is a real cost - the "
                 "Expenses tab has a category for it too."),
        Col("method", "Method", 15, "text",
            validate=list_from("PaymentMethodList")),
        Col("reference", "Reference", 18, "text",
            note="Bank reference, cheque number, transaction ID. "
                 "Worth having when a client says they paid."),
        Col("deposit", "Is this a deposit?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("banked", "Banked?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Cash in a jacket pocket is not banked. This column "
                 "is how you find it again."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("net", "Net received", 11, "money",
            formula=blank("date",
                          'IF({amount}{r}="",0,{amount}{r})-'
                          'IF({fee}{r}="",0,{fee}{r})'),
            note="After the processing fee."),
        Col("invoice_total", "Invoice total", 11, "money",
            formula='=IF({invoice}{r}="","",IFERROR(INDEX('
                    '[[Invoices.total]],MATCH({invoice}{r},'
                    '[[Invoices.id]],0)),""))',
            hidden=True),
        Col("invoice_paid", "Paid on invoice", 11, "money",
            formula='=IF({invoice}{r}="","",IFERROR(INDEX('
                    '[[Invoices.paid]],MATCH({invoice}{r},'
                    '[[Invoices.id]],0)),""))',
            hidden=True),
        Col("balance", "Invoice balance after", 12, "money",
            formula='=IF(OR({invoice}{r}="",{invoice_total}{r}=""),"",'
                    'ROUND({invoice_total}{r}-{invoice_paid}{r},2))',
            note="What is still owed on that invoice once every "
                 "payment against it is counted.",
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": -0.004,
                 "fmt": "bad"}]),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "date",
                'IF({invoice}{r}="","Add - which invoice?",'
                'IF(OR({amount}{r}="",{amount}{r}=0),'
                '"Add - how much came in?",'
                'IF({balance}{r}<-0.004,"Check - overpaid",'
                'IF({banked}{r}="No","Due - not banked yet",'
                'IF({balance}{r}>0.004,"Part paid","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Payment Tracking",
        subtitle="Every payment received, matched to its invoice.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).payments,
        help_text="One row per payment. Part payments get a row each "
                  "against the same invoice, and the invoice balance "
                  "updates itself.",
        extra_notes=[
            "Record the gross amount and the fee separately. A 2.9% "
            "card fee on $60,000 of turnover is $1,740 you would "
            "otherwise never see.",
            "The 'Banked?' column exists because cash jobs are how "
            "small cleaning businesses lose money without noticing.",
        ],
    )
