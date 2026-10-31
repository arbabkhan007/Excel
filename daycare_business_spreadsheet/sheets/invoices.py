"""`Invoices & Receipts` - one ledger for everything that moves money.

Every row is one movement on a family's account:

    Invoice      you asked them for money        account goes up
    Payment      they gave you some              account goes down
    Credit note  you cancelled part of a bill    account goes down
    Write-off    you gave up on a bill           account goes down
    Refund       you gave money back             account goes up

That is the whole model, and it is why partial payments, instalments,
deposits, credits and refunds all work without a single extra column.
Pay an invoice in four goes and you get four Payment rows pointing at
it; the invoice knows how much is left.

Two figures are kept deliberately separate.  **Billed** is revenue -
what you charged, less credits and write-offs.  **Collected** is cash -
what actually arrived, less refunds.  They are different numbers, they
answer different questions, and mixing them up is the single most
common way a childcare business talks itself into trouble.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, card_fee, lookup_family,
                     month_col, owing_cf, status_cf)
from .common import ROW_OFFSET

NAME = "Invoices & Receipts"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        # An invoice gets an invoice number and a receipt gets a
        # receipt number, from the same row - which is exactly what a
        # parent expects to see on the piece of paper.
        Col("id", "Reference", 12, "id",
            formula=('=IF({date}{r}="","",'
                     'IF({type}{r}="","",'
                     'IF(({type}{r}="Payment")+({type}{r}="Refund")>0,'
                     'ReceiptPrefix,InvoicePrefix)&'
                     f'TEXT(ROW()-{ROW_OFFSET},"00000")))'),
            note="Automatic. Invoice rows get an invoice number, "
                 "payment and refund rows get a receipt number. Both "
                 "come from the row, so they never collide."),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("type", "Type", 12, "status",
            validate=list_from("LedgerTypeList"), cf=status_cf(),
            note="Invoice, Payment, Credit note, Refund or Write-off. "
                 "This one cell decides what the row does to the "
                 "family's balance, to revenue and to cash."),
        Col("family", "Family", 11, "id",
            validate=list_from("FamilyList"),
            note="Whose account this touches. Invoices go to the "
                 "family, not the child, so siblings share one bill."),
        Col("family_name", "Family name", 18, "text",
            formula=lookup_family("family_name")),
        Col("child", "Child (optional)", 11, "id",
            validate=list_from("ChildList"), hidden=True,
            note="Only when you want one child's name on the "
                 "document. The money still belongs to the family."),
        Col("period", "Period", 12, "text",
            note="What the invoice covers in plain words - 'October "
                 "2026' or 'w/c 12 Oct'. Printed on the invoice."),
        Col("description", "Description", 30, "text"),
        # --------------------------------------------------- amounts ----
        Col("charges", "From charge lines", 14, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice",0,'
                'ROUND(SUMIFS([[Tuition & Fees.net]],'
                '[[Tuition & Fees.invoice]],{id}{r}),2))'),
            note="Everything on the Tuition & Fees tab tagged with "
                 "this invoice number. Add the lines there and this "
                 "fills itself in."),
        Col("lines", "Lines", 7, "int",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",'
                'COUNTIFS([[Tuition & Fees.invoice]],{id}{r}))'),
            hidden=True),
        Col("amount", "Amount typed", 12, "money",
            validate=number_check(0, 1000000),
            note="For payments, refunds, credit notes and write-offs, "
                 "this is the amount. For an invoice, leave it blank "
                 "if you are using charge lines - type a figure only "
                 "for a one-off bill with no lines behind it."),
        Col("total", "Total", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({type}{r}="Invoice",'
                'IF({lines}{r}>0,{charges}{r},{amount}{r}),'
                '{amount}{r}),2)'),
            cf=bar_cf(Palette.accent),
            note="The figure on the document. An invoice uses its "
                 "charge lines when it has any, and the typed amount "
                 "when it does not."),
        # ------------------------------------------------- settlement ---
        Col("against", "Pays invoice", 12, "text",
            validate=list_from("InvoiceList"),
            note="On a payment, refund, credit note or write-off, the "
                 "invoice number it applies to. A payment with nothing "
                 "here still reduces the family's balance - it just "
                 "is not matched to a bill."),
        Col("method", "Method", 15, "text",
            validate=list_from("PaymentMethodList")),
        Col("reference", "Reference", 14, "text", hidden=True,
            note="Cheque number, transaction ID, the reference on the "
                 "bank statement."),
        Col("terms", "Terms (days)", 10, "int",
            validate=number_check(0, 365, decimal=False), hidden=True,
            note="Leave blank to use the standard payment terms from "
                 "Settings."),
        Col("due_date", "Due", 11, "date",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",{date}{r}+'
                'IF({terms}{r}="",PaymentTermsDays,{terms}{r}))'),
            note="Invoice date plus your payment terms."),
        Col("due_num", "Due (number)", 10, "number",
            # the due date as a plain number in every row, including
            # the empty ones. Other sheets search this column instead
            # of the dated one: multiplying a blank by a number is an
            # error in Excel, and a column of zeros never is.
            formula='=IF({date}{r}="",0,IF({due_date}{r}="",0,'
                    '{due_date}{r}))',
            hidden=True),
        Col("paid", "Paid so far", 12, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",'
                'ROUND(SUMIFS({total_first}:{total_last},'
                '{type_first}:{type_last},"Payment",'
                '{against_first}:{against_last},{id}{r})-'
                'SUMIFS({total_first}:{total_last},'
                '{type_first}:{type_last},"Refund",'
                '{against_first}:{against_last},{id}{r}),2))'),
            note="Every payment pointing at this invoice, less any "
                 "refund. Partial payments just work."),
        Col("credited", "Credited", 11, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",'
                'ROUND(SUMIFS({total_first}:{total_last},'
                '{type_first}:{type_last},"Credit note",'
                '{against_first}:{against_last},{id}{r})+'
                'SUMIFS({total_first}:{total_last},'
                '{type_first}:{type_last},"Write-off",'
                '{against_first}:{against_last},{id}{r}),2))'),
            hidden=True),
        Col("balance", "Still owed", 12, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",'
                'ROUND({total}{r}-{paid}{r}-{credited}{r},2))'),
            cf=owing_cf(),
            note="What is left on this invoice."),
        Col("status", "Status", 12, "status",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice",{type}{r},'
                'IF({sent}{r}<>"Yes","Draft",'
                'IF({balance}{r}<=0.004,"Paid",'
                'IF(IF({due_date}{r}="",FALSE,'
                '{due_date}{r}<TodayDate),"Overdue",'
                'IF({paid}{r}>0.004,"Part paid","Sent")))))'),
            cf=status_cf(),
            note="Draft until you mark it sent, then Sent, Part paid, "
                 "Paid or Overdue - worked out, never typed."),
        Col("sent", "Sent to parent?", 11, "flag",
            validate=list_from("YesNoList"),
            note="An invoice you have not sent is not overdue, so the "
                 "workbook will not call it overdue."),
        Col("days_overdue", "Days overdue", 11, "int",
            formula=blank(
                "date",
                'IF({status}{r}<>"Overdue","",'
                'TodayDate-{due_date}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": 30,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}]),
        # ------------------------------------- what it does to the books
        Col("billed", "Revenue effect", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({type}{r}="Invoice",{total}{r},'
                'IF({type}{r}="Credit note",-{total}{r},'
                'IF({type}{r}="Write-off",-{total}{r},0))),2)'),
            note="Revenue. Invoices add, credit notes and write-offs "
                 "take away. Payments are not revenue - they are the "
                 "same money arriving later."),
        Col("collected", "Cash effect", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({type}{r}="Payment",{total}{r},'
                'IF({type}{r}="Refund",-{total}{r},0)),2)'),
            note="Cash. Payments in, refunds out. Never counted as "
                 "revenue a second time."),
        Col("owed", "Account effect", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF(({type}{r}="Invoice")+({type}{r}="Refund")'
                '>0,{total}{r},-{total}{r}),2)'),
            hidden=True,
            note="What this row does to the family's balance. The "
                 "parent tab adds the column up."),
        Col("fee", "Processing fee", 11, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Payment",0,' +
                card_fee("{total}{r}", "{method}{r}") + ")"),
            hidden=True,
            note="What a card payment costs you. Set the rates on "
                 "Settings."),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "date",
                'IF({type}{r}="","Add - pick a row type",'
                'IF({family}{r}="","Add - no family on this row",'
                'IF({family_name}{r}="Not in Parent Database",'
                '"Check - family ID not found",'
                'IF(IF({type}{r}<>"Invoice",FALSE,'
                'IF({lines}{r}>0,FALSE,{amount}{r}="")),'
                '"Add - no charge lines and no amount",'
                'IF(IF({type}{r}="Invoice",FALSE,{amount}{r}=""),'
                '"Add - no amount entered",'
                'IF(IF({type}{r}="Invoice",FALSE,'
                'IF({type}{r}="Payment",FALSE,{against}{r}="")),'
                '"Check - not matched to an invoice",'
                'IF({status}{r}="Overdue","Overdue - "&'
                '{days_overdue}{r}&" days, chase it",'
                'IF(IF({type}{r}<>"Invoice",FALSE,{sent}{r}<>"Yes"),'
                '"Action - invoice still a draft",'
                'IF(IF({type}{r}<>"Payment",FALSE,{method}{r}=""),'
                '"Add - how did they pay?","OK")))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Invoices, Payments & Receipts",
        subtitle="One ledger for invoices, payments, credits, refunds "
                 "and write-offs. Revenue and cash kept firmly apart.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).ledger,
        help_text="Add an Invoice row, tag the charge lines on "
                  "Tuition & Fees with its number, then add a Payment "
                  "row each time money arrives. Everything else is "
                  "worked out.",
        extra_notes=[
            "To take a payment in instalments, add one Payment row "
            "per instalment, all pointing at the same invoice number. "
            "The invoice shows what is left after each one.",
            "A deposit is just a Payment with no invoice behind it "
            "yet. It sits on the family's account as a credit until "
            "you raise the bill it belongs to - which is why it can "
            "never be counted as income twice.",
            "'Revenue effect' and 'Cash effect' are two different "
            "columns on purpose. The profit figures use revenue; the "
            "collection rate uses cash.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Ledger summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The ledger", last)
    sheet.table_head(row, 0, [
        "Rows", "Invoices raised", "Invoiced", "Credits & write-offs",
        "Revenue", "Collected", "Outstanding", "Overdue",
        "Overdue invoices", "Collection rate", "Processing fees",
        "This month revenue", "This month collected"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Invoices & Receipts.date]],"<>")',
        '=COUNTIFS([[Invoices & Receipts.type]],"Invoice")',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.type]],"Invoice"),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.type]],"Credit note")+'
        'SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.type]],"Write-off"),2)',
        '=ROUND(SUM([[Invoices & Receipts.billed]]),2)',
        '=ROUND(SUM([[Invoices & Receipts.collected]]),2)',
        '=ROUND(SUM([[Invoices & Receipts.billed]])-'
        'SUM([[Invoices & Receipts.collected]]),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.balance]],'
        '[[Invoices & Receipts.status]],"Overdue"),2)',
        '=COUNTIFS([[Invoices & Receipts.status]],"Overdue")',
        '=IF(SUM([[Invoices & Receipts.billed]])<=0,"",'
        'SUM([[Invoices & Receipts.collected]])/'
        'SUM([[Invoices & Receipts.billed]]))',
        '=ROUND(SUM([[Invoices & Receipts.fee]]),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.billed]],'
        '[[Invoices & Receipts.month]],ThisMonthKey),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.collected]],'
        '[[Invoices & Receipts.month]],ThisMonthKey),2)',
    ], ["int", "int", "money", "money", "money", "money", "money",
        "money", "int", "percent", "money", "money", "money"])
