"""`Tuition & Fees` - every charge and every credit, one row at a time.

This is the charge register, not the invoice.  Tuition for a week, a
registration fee, a late pickup, a field trip, the sibling discount, a
subsidy credit - each one is a line here, and each one carries the
invoice number it was billed on.  Invoices & Receipts then adds the
lines up.

Keeping charges and invoices apart is what makes partial billing,
mid-month changes and credits possible without any of the arithmetic
going wrong.  Discounts and subsidies are entered as positive numbers
and subtracted, so a parent never has to read a minus sign.
"""

from __future__ import annotations

from ..config import CREDIT_FEE_TYPES, Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, lookup_child,
                     month_col, seq_key, status_cf, tax_on)

NAME = "Tuition & Fees"


def spec(cfg: Config) -> SheetSpec:
    credit_test = "+".join(f'({{fee_type}}{{r}}="{t}")'
                           for t in CREDIT_FEE_TYPES)

    columns = [
        id_col("date", "FeePrefix", digits=5, width=11,
               header="Charge ID"),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The date the charge belongs to. Monthly reports "
                 "group on it."),
        month_col("date"),
        Col("child", "Child", 11, "id", validate=list_from("ChildList"),
            note="Which child the charge is for. Leave blank only for "
                 "a family-level charge such as a registration fee "
                 "covering siblings."),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("family", "Family", 11, "id",
            formula=blank(
                "child",
                'IFERROR(INDEX([[Child Enrollment.family]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"","")'),
            note="From the child's record. The invoice goes to this "
                 "family."),
        Col("family_set", "Family override", 11, "id",
            validate=list_from("FamilyList"), hidden=True,
            note="Only needed for a charge that has no child against "
                 "it at all."),
        Col("family_used", "Billed to", 11, "id",
            formula=blank(
                "date",
                'IF({family}{r}<>"",{family}{r},{family_set}{r})'),
            hidden=True),
        Col("fee_type", "Charge type", 17, "text",
            validate=list_from("FeeTypeList"), cf=status_cf(),
            note="Discounts, subsidies and credits are in this list "
                 "too. Enter them as a positive amount - the workbook "
                 "knows to take them off."),
        Col("is_credit", "Credit?", 8, "flag",
            formula=blank("fee_type",
                          f'IF(({credit_test})>0,"Yes","No")'),
            hidden=True),
        Col("description", "Description", 30, "text",
            note="What a parent would need to read to recognise the "
                 "charge. 'Tuition, week beginning 5 Oct' beats "
                 "'fees'."),
        Col("period_from", "Covers from", 11, "date",
            validate=date_check(), hidden=True),
        Col("period_to", "Covers to", 11, "date",
            validate=date_check(), hidden=True),
        Col("qty", "Quantity", 9, "number1",
            validate=number_check(0, 10000),
            note="Weeks, days, hours, sessions or trips - whatever the "
                 "rate is per. Leave blank and it counts as 1."),
        Col("rate", "Rate", 11, "money",
            validate=number_check(-100000, 100000),
            note="The unit price. For a flat charge, put the whole "
                 "amount here and leave the quantity blank."),
        Col("gross", "Gross", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({qty}{r}="",1,{qty}{r})*{rate}{r},2)')),
        Col("discount_pct", "Discount %", 10, "percent",
            validate=number_check(0, 1), hidden=True),
        Col("discount", "Discount", 11, "money",
            formula=blank(
                "date",
                'ROUND({gross}{r}*{discount_pct}{r},2)'),
            hidden=True),
        Col("subtotal", "Subtotal", 11, "money",
            formula=blank("date",
                          'ROUND({gross}{r}-{discount}{r},2)')),
        Col("tax", "Tax", 10, "money",
            formula=blank(
                "date",
                tax_on("{subtotal}{r}", '{fee_type}{r}="Tuition"')),
            note="Zero unless you switch tax on for tuition or extras "
                 "on the Settings tab. Childcare is exempt in a great "
                 "many places, so nothing is assumed."),
        Col("net", "Amount", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({is_credit}{r}="Yes",-1,1)*'
                '({subtotal}{r}+{tax}{r}),2)'),
            cf=bar_cf(Palette.accent),
            note="What this line adds to the bill. Discounts, "
                 "subsidies and credits come out negative even though "
                 "you typed a positive number."),
        # -------------------------------------------------- invoicing --
        Col("invoice", "Invoice", 12, "text",
            validate=list_from("InvoiceList"),
            note="The invoice number this line was billed on. Leave it "
                 "blank until you raise the invoice and the Action "
                 "column will remind you."),
        Col("billed", "Billed?", 9, "flag",
            formula=blank("date",
                          'IF({invoice}{r}="","No","Yes")'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"}]),
        Col("invoice_ok", "Invoice found?", 11, "flag",
            formula=blank(
                "invoice",
                'IF(COUNTIF([[Invoices & Receipts.id]],'
                '{invoice}{r})>0,"Yes","No")'),
            hidden=True),
        # Line 1, 2, 3 ... within one invoice, and the `INV-00012|3`
        # key the printable invoice uses to pull a line without an
        # array formula.
        *seq_key("invoice"),
        Col("auto_source", "Raised by", 13, "text",
            validate=list_from("ChargeSourceList"), hidden=True,
            note="Where the charge came from - the schedule, a late "
                 "pickup, an extra day. Useful when you are checking "
                 "a bill a parent has queried."),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF(IF({child}{r}="",{family_set}{r}="",FALSE),'
                '"Add - no child or family on this charge",'
                'IF({fee_type}{r}="","Add - pick a charge type",'
                'IF({rate}{r}="","Add - no rate entered",'
                'IF(IF({invoice}{r}="",FALSE,{invoice_ok}{r}="No"),'
                '"Check - that invoice number does not exist",'
                'IF(IF({invoice}{r}<>"",FALSE,'
                'TodayDate-{date}{r}>PaymentTermsDays),'
                '"Invoice - charge is '
                '"&(TodayDate-{date}{r})&" days old and unbilled",'
                'IF({invoice}{r}="","Invoice - not billed yet",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Tuition & Fee Charges",
        subtitle="Every charge and every credit, ready to be pulled "
                 "onto an invoice. Discounts and subsidies are typed "
                 "positive and come off automatically.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).fees,
        help_text="Add a line for anything you want to bill. Put the "
                  "invoice number against it once it is billed, and "
                  "Invoices & Receipts adds the lines up for you.",
        extra_notes=[
            "The rate card lives on the Child Enrollment tab: basis, "
            "agreed rate, discount and subsidy per child. This tab is "
            "the actual charges that came out of it.",
            "Late pickup fees are calculated on the Attendance tab "
            "from each child's own contracted collection time. Copy "
            "the amount here to charge it, then mark it as charged "
            "over there.",
            "A charge with no invoice number against it is a charge "
            "nobody has been asked to pay. That is what the Action "
            "column is watching for.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Charge register summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Charges raised", last)
    sheet.table_head(row, 0, [
        "Charge lines", "Gross charges", "Discounts & credits",
        "Tax", "Net charged", "Tuition", "Extras & fees",
        "Billed", "Not yet billed", "This month"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Tuition & Fees.date]],"<>")',
        '=ROUND(SUMIFS([[Tuition & Fees.subtotal]],'
        '[[Tuition & Fees.is_credit]],"No"),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.subtotal]],'
        '[[Tuition & Fees.is_credit]],"Yes"),2)',
        '=ROUND(SUM([[Tuition & Fees.tax]]),2)',
        '=ROUND(SUM([[Tuition & Fees.net]]),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.fee_type]],"Tuition"),2)',
        '=ROUND(SUM([[Tuition & Fees.net]])-'
        'SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.fee_type]],"Tuition"),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.billed]],"Yes"),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.billed]],"No"),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.month]],ThisMonthKey),2)',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "money", "money", "money"])
