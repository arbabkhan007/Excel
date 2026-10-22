"""`Payments` - money in and money back out, with the deposit rule.

A deposit is cash in the bank and nothing more. It is logged here the
day it is taken, it shows against the appointment, and it is applied to
the bill when the sale is rung up. It is never revenue, and it is never
counted twice: once a payment carries a sale ID it stops being an
unapplied deposit.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_client, lookup_staff,
                     month_col, status_cf, week_col, yesno)

NAME = "Payments"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "PaymentPrefix", digits=4, width=11),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day the money moved."),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList")),
        Col("client_name", "Client", 20, "text",
            formula=lookup_client("name", missing="Walk-in"),
            hidden=True),
        Col("kind", "What for", 16, "text",
            validate=list_from("PaymentKindList"),
            note="Deposit, Balance, Full payment, Retail only, Refund, "
                 "Gift certificate sale, Package sale or Tip. This "
                 "column decides how the money is treated everywhere "
                 "else."),
        Col("method", "Method", 15, "text",
            validate=list_from("PaymentMethodList"),
            note="Card payments pick up the processing fee from "
                 "Settings automatically."),
        Col("amount", "Amount", 11, "money",
            validate=number_check(0, 1000000),
            note="Always a positive number, even for a refund - the "
                 "kind column says which way it went."),
        Col("sale", "Sale ID", 11, "id",
            note="The till line this settles. Fill it in and the "
                 "payment is applied to that sale."),
        Col("appointment", "Appointment", 12, "id",
            note="For a deposit taken before there is a sale. When the "
                 "client comes in, the deposit is applied to the first "
                 "sale line of that appointment on its own."),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList"),
            note="Who the tip belongs to, or who took the payment."),
        Col("staff_name", "Stylist", 16, "text",
            formula=lookup_staff("name", missing=""), hidden=True),
        Col("reference", "Reference", 16, "text",
            note="Receipt number, last four digits, transfer "
                 "reference - whatever you would need to find it "
                 "again."),
        Col("banked", "Reconciled?", 12, "flag", validate=yesno(),
            note="Tick it off when you have seen it on the bank "
                 "statement or in the till count."),
        Col("notes", "Notes", 24, "wrap"),
        # ---------------------------------------------------- worked out --
        month_col("date"),
        week_col("date"),
        Col("counts", "Settles a bill?", 12, "flag",
            formula=blank(
                "date",
                'IF({kind}{r}="Tip","No",IF({kind}{r}="Refund","No",'
                '"Yes"))'),
            hidden=True,
            note="Tips belong to the stylist and refunds go the other "
                 "way, so neither of them pays off a bill."),
        Col("direction", "In or out", 9, "int",
            formula=blank(
                "date",
                'IF({kind}{r}="Refund",-1,1)'),
            hidden=True),
        Col("fee", "Processing fee", 12, "money",
            formula=blank(
                "date",
                'IF({amount}{r}="",0,IF({method}{r}="Card",'
                'ROUND({amount}{r}*CardFeePercent+CardFeeFixed,2),'
                'IF({method}{r}="Online deposit",'
                'ROUND({amount}{r}*CardFeePercent+CardFeeFixed,2),0)))'),
            note="What the card machine keeps. Set the percentage and "
                 "the fixed part on Settings; it never touches cash or "
                 "bank transfers."),
        Col("net_cash", "Into the bank", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({amount}{r}="",0,{amount}{r})*'
                '{direction}{r}-{fee}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}],
            note="Amount, minus the fee, negative for a refund. Add "
                 "this column up and you have your cash flow."),
        Col("is_tip", "Tip?", 8, "flag",
            formula=blank("date", 'IF({kind}{r}="Tip","Yes","No")'),
            hidden=True),
        Col("applied_to", "Applied to", 14, "text",
            formula=blank(
                "date",
                'IF({sale}{r}<>"","Sale "&{sale}{r},'
                'IF({appointment}{r}<>"","Booking "&{appointment}{r},'
                'IF({kind}{r}="Tip","Stylist","Unapplied")))'),
            cf=status_cf(),
            note="Unapplied is not wrong - a gift certificate sale or "
                 "a general payment on account belongs here."),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "date",
                'IF({amount}{r}="","Add - no amount yet",'
                'IF({kind}{r}="","Add - what was this for?",'
                'IF(IF({sale}{r}="",FALSE,'
                'COUNTIF([[Sales.id]],{sale}{r})=0),'
                '"Missing - no sale with that ID",'
                'IF(IF({appointment}{r}="",FALSE,'
                'COUNTIF([[Appointments.id]],{appointment}{r})=0),'
                '"Missing - no booking with that ID",'
                'IF(IF({kind}{r}="Tip",{staff}{r}="",FALSE),'
                '"Assign - whose tip is it?",'
                'IF(IF({kind}{r}="Deposit",'
                'IF({sale}{r}="",{appointment}{r}="",FALSE),FALSE),'
                '"Check - deposit not tied to a booking",'
                'IF({banked}{r}="No","Due - not reconciled yet",'
                '"OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Payments & Deposits",
        subtitle="Every payment, deposit, tip and refund - what it "
                 "settles, what the card machine took, what reached "
                 "the bank.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).payments,
        help_text="Log the money here; log what was sold on the Sales "
                  "tab. Keeping them apart is what lets a deposit sit "
                  "in the bank for three weeks without pretending to "
                  "be revenue.",
        extra_notes=[
            "Deposits: log one with kind Deposit and the appointment "
            "ID. It shows on the booking straight away and is applied "
            "to the first sale line of that appointment when the "
            "client comes in. Once a payment has a sale ID it is no "
            "longer treated as an unapplied deposit, so nothing is "
            "ever counted twice.",
            "Tips are the stylist's money. They are kept out of sales, "
            "out of the profit and loss, and carried through to the "
            "payroll tab instead.",
            "Refunds are logged as a positive amount with kind Refund. "
            "The 'into the bank' column turns them negative; the sale "
            "they relate to should also carry the refunded amount so "
            "that net sales drops.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Money summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Money in, money out", last)
    sheet.table_head(row, 0, [
        "Payments logged", "Taken", "Refunded", "Tips", "Card fees",
        "Into the bank", "Deposits held", "Not reconciled"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Payments.date]],"<>")',
        '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],'
        '"<>Refund"),2)',
        '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],'
        '"Refund"),2)',
        '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],"Tip"),2)',
        '=ROUND(SUM([[Payments.fee]]),2)',
        '=ROUND(SUM([[Payments.net_cash]]),2)',
        '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],"Deposit",'
        '[[Payments.sale]],""),2)',
        '=COUNTIFS([[Payments.banked]],"No")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "int"])
