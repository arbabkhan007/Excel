"""The three printable forms.

* `Print Quote` - a quote on one side of A4 with your letterhead, the
  lines, the total and an acceptance box the client can sign.
* `Print Invoice` - the same for an invoice, with payment details, the
  due date and what is still owed.
* `Print Job Sheet` - the sheet the cleaner takes to the property:
  access details, the checklist, time in and out, and a sign-off.

Pick an ID at the top of each page and the whole form fills itself in
from the registers. Nothing is typed twice.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from

QUOTE_NAME = "Print Quote"
INVOICE_NAME = "Print Invoice"
JOB_NAME = "Print Job Sheet"

QUOTE_LINES = 12
INVOICE_LINES = 12
TASK_LINES = 18


def _letterhead(s: SheetWriter, fmts, row: int, last_col: int,
                heading: str) -> int:
    """The business name, address and contact block."""
    ws = s.ws
    s.merge(row, 0, 3, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0, "=BizName", fmts.chrome("print_title"))
    s.merge(row, 4, last_col, heading, fmts.chrome("print_title"))
    row += 1
    s.merge(row, 0, 3, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0,
                     '=BizAddress&IF(BizArea=""," "," - "&BizArea)',
                     fmts.chrome("print_sub"))
    s.merge(row, 4, last_col, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 4,
                     '=BizPhone&"   "&BizEmail&IF(BizWebsite="","",'
                     '"   "&BizWebsite)', fmts.chrome("print_sub"))
    row += 1
    s.merge(row, 0, last_col, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0,
                     '=IF(BizRegNo="","","Registration "&BizRegNo&'
                     '"   ")&IF(BizInsurance="","","Insurance: "&'
                     'BizInsurance)', fmts.chrome("print_sub"))
    return row + 2


def _pick(s: SheetWriter, fmts, book, row: int, label: str,
          default: str, list_name: str, define: str,
          help_text: str) -> int:
    """The 'which one are we printing' picker."""
    ws = s.ws
    s.write(row, 0, label, fmts.chrome("label_bold"))
    s.write(row, 1, default, fmts.input("id"))
    ws.data_validation(row, 1, row, 1, list_from(list_name))
    book.define_name(define,
                     f"={quote_sheet(s.name)}!$B${row + 1}")
    s.merge(row, 2, 7, help_text, fmts.chrome("note"))
    return row + 2


# ---------------------------------------------------------------------------
# Print Quote
# ---------------------------------------------------------------------------

def build_quote(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = QUOTE_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Quote", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 16)
    s.set_col(1, 1, 24)
    s.set_col(2, 2, 26)
    s.set_col(3, 3, 10)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 12)
    s.set_col(7, 7, 14)

    row = 4
    row = _pick(s, fmts, book, row, "Quote to print", "Q-0001",
                "QuoteIDs", "PrintQuote",
                "Pick a quote. Everything below fills itself in. "
                "Print with File, Print - the page is already sized "
                "for A4.")
    row = _letterhead(s, fmts, row, LAST_COL, "QUOTE")

    def q(key: str) -> str:
        return s.reg.render(f'=IFERROR(INDEX([[Quotes.{key}]],'
                            f'MATCH(PrintQuote,[[Quotes.id]],0)),"")')

    # --------------------------------------------------- the two blocks --
    s.write(row, 0, "Quote for", fmts.chrome("label_bold"))
    s.write(row, 4, "Quote number", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, "=PrintQuote", fmts.cell("id", "plain"))
    row += 1
    s.merge(row, 0, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 0,
                     s.reg.render('=IFERROR(INDEX([[Quotes.client_name]],'
                                  'MATCH(PrintQuote,[[Quotes.id]],0)),"")'
                                  ), fmts.cell("text", "plain"))
    s.write(row, 4, "Date", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, q("date"), fmts.cell("date", "plain"))
    row += 1
    s.merge(row, 0, 3, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0, q("address"), fmts.cell("wrap", "plain"))
    s.write(row, 4, "Valid until", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, q("valid_until"), fmts.cell("date", "plain"))
    s.row_height(row, 32)
    row += 1
    s.write(row, 0, "Service", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 1, q("service"), fmts.cell("text", "plain"))
    s.write(row, 4, "How often", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, q("frequency"), fmts.cell("text", "plain"))
    row += 2

    # ------------------------------------------------------- the lines --
    row = s.table_head(row, 0, ["Item", "What is included", "Qty",
                                "Price", "Line total", "Optional?"],
                       [18, 36, 8, 12, 13, 11])
    first_line = row
    for n in range(1, QUOTE_LINES + 1):
        key = f'PrintQuote&"|{n}"'
        for col, (field, style) in enumerate([
                ("service", "text"), ("wording", "wrap"),
                ("qty", "number1"), ("price_used", "money"),
                ("total", "money"), ("optional", "text")]):
            if col == 1:
                s.merge(row, 1, 2, None, fmts.cell("wrap", "plain"))
            target = col if col == 0 else col + 1
            ws.write_formula(
                row, target,
                s.reg.render(f'=IFERROR(INDEX([[Quote Lines.{field}]],'
                             f'MATCH({key},[[Quote Lines.key]],0)),"")'),
                fmts.cell(style, "plain"))
        row += 1
    last_line = row - 1
    ws.conditional_format(first_line, 6, last_line, 6, {
        "type": "text", "criteria": "containing", "value": "Yes",
        "format": fmts.cf("info")})
    row += 1

    # ------------------------------------------------------ the totals --
    totals = [("Subtotal", q("subtotal"), "money"),
              ("Discount", q("discount_value"), "money"),
              (cfg.tax_label, q("tax"), "money"),
              ("Total", q("total"), "money"),
              ("Deposit to confirm", q("deposit"), "money")]
    for i, (label, formula, style) in enumerate(totals):
        s.merge(row, 3, 4, label,
                fmts.chrome("total_label" if i == 3 else "label_right"))
        ws.write_formula(
            row, 5, formula,
            fmts.chrome("total_value", num_format=cfg.money_format)
            if i == 3 else fmts.cell(style, "plain"))
        if i == 0:
            s.merge(row, 0, 2,
                    "Prices hold until the valid-until date above.",
                    fmts.chrome("note"))
        if i == 4:
            s.merge(row, 0, 2,
                    "The balance is due on the terms below.",
                    fmts.chrome("note"))
        row += 1
    row += 1

    # -------------------------------------------------- what's included --
    row = s.section(row, "What is included", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0, q("notes"), fmts.cell("wrap", "plain"))
    s.row_height(row, 46)
    row += 1
    row = s.section(row, "Terms", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        s.reg.render('=IF(IFERROR(INDEX([[Quotes.terms]],MATCH('
                     'PrintQuote,[[Quotes.id]],0)),"")="",'
                     '"Payment within "&PaymentTerms&" days of '
                     'invoice. Cancellations inside 24 hours may be '
                     'charged. We bring our own supplies and '
                     'equipment unless agreed otherwise.",'
                     'INDEX([[Quotes.terms]],MATCH(PrintQuote,'
                     '[[Quotes.id]],0)))'),
        fmts.cell("wrap", "plain"))
    s.row_height(row, 46)
    row += 2

    # -------------------------------------------------- the acceptance --
    row = s.section(row, "Acceptance", LAST_COL)
    s.merge(row, 0, LAST_COL,
            "I accept this quote and agree to the terms above.",
            fmts.chrome("note"))
    row += 1
    for label, span in [("Signed", 3), ("Print name", 3)]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, span, None, fmts.chrome("form_line"))
        s.write(row, 4, "Date", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 26)
        row += 1
    row += 1
    s.merge(row, 0, LAST_COL,
            "A signature on a printed page is not an electronic "
            "signature. If you need one that stands up legally, use "
            "a proper e-signature service - then record the date "
            "here.", fmts.chrome("note"))
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(row, 0,
                     '="Quote produced by "&BizName&" on "&'
                     'TEXT(TodayDate,"d mmmm yyyy")',
                     fmts.chrome("note"))

    s.finish(freeze_row=None, print_area=f"A1:H{row + 2}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Invoice
# ---------------------------------------------------------------------------

def build_invoice(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = INVOICE_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Invoice", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 16)
    s.set_col(1, 1, 24)
    s.set_col(2, 2, 26)
    s.set_col(3, 3, 10)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 12)
    s.set_col(7, 7, 14)

    row = 4
    row = _pick(s, fmts, book, row, "Invoice to print", "INV-0001",
                "InvoiceIDs", "PrintInvoice",
                "Pick an invoice. Check the totals against the "
                "Invoices tab before you send it.")
    row = _letterhead(s, fmts, row, LAST_COL, "INVOICE")

    def inv(key: str) -> str:
        return s.reg.render(f'=IFERROR(INDEX([[Invoices.{key}]],'
                            f'MATCH(PrintInvoice,[[Invoices.id]],0)),"")')

    s.write(row, 0, "Invoice to", fmts.chrome("label_bold"))
    s.write(row, 4, "Invoice number", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, "=PrintInvoice", fmts.cell("id", "plain"))
    row += 1
    s.merge(row, 0, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 0, inv("client_name"),
                     fmts.cell("text", "plain"))
    s.write(row, 4, "Date", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, inv("date"), fmts.cell("date", "plain"))
    row += 1
    s.merge(row, 0, 3, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0, inv("address"), fmts.cell("wrap", "plain"))
    s.write(row, 4, "Due date", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, inv("due_date"), fmts.cell("date", "plain"))
    s.row_height(row, 32)
    row += 1
    s.write(row, 0, "Period", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.cell("text", "plain"))
    ws.write_formula(row, 1, inv("period"), fmts.cell("text", "plain"))
    s.write(row, 4, "Your reference", fmts.chrome("label_bold"))
    ws.write_formula(row, 5, inv("po"), fmts.cell("text", "plain"))
    row += 2

    row = s.table_head(row, 0, ["Date", "What we did", "Qty", "Price",
                                "Line total", "Tax?"],
                       [12, 42, 8, 12, 13, 9])
    first_line = row
    for n in range(1, INVOICE_LINES + 1):
        key = f'PrintInvoice&"|{n}"'
        fields = [("date", "date"), ("wording", "wrap"),
                  ("qty", "number1"), ("price_used", "money"),
                  ("total", "money"), ("taxable", "text")]
        for col, (field, style) in enumerate(fields):
            if col == 1:
                s.merge(row, 1, 2, None, fmts.cell("wrap", "plain"))
            target = col if col == 0 else col + 1
            ws.write_formula(
                row, target,
                s.reg.render(f'=IFERROR(INDEX([[Invoice Lines.{field}]],'
                             f'MATCH({key},[[Invoice Lines.key]],0)),"")'),
                fmts.cell(style, "plain"))
        row += 1
    last_line = row - 1
    ws.conditional_format(first_line, 5, last_line, 5,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("info")})
    row += 1

    totals = [("Subtotal", inv("subtotal")),
              ("Discount", inv("discount_value")),
              (cfg.tax_label, inv("tax")),
              ("Invoice total", inv("total")),
              ("Paid so far", inv("paid")),
              ("Amount due", inv("outstanding"))]
    for i, (label, formula) in enumerate(totals):
        big = label in ("Invoice total", "Amount due")
        s.merge(row, 3, 4, label,
                fmts.chrome("total_label" if big else "label_right"))
        ws.write_formula(
            row, 5, formula,
            fmts.chrome("total_value", num_format=cfg.money_format)
            if big else fmts.cell("money", "plain"))
        row += 1
    row += 1

    row = s.section(row, "How to pay", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        '="Payment within "&PaymentTerms&" days. "&'
        'IF(BizPayDetails="","Bank details on request.",BizPayDetails)&'
        'IF(LateFeePercent=0,"","  Late payments may carry a fee of "&'
        'TEXT(LateFeePercent,"0%")&".")',
        fmts.cell("wrap", "plain"))
    s.row_height(row, 40)
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0, inv("notes"), fmts.cell("wrap", "plain"))
    s.row_height(row, 32)
    row += 2

    s.merge(row, 0, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(
        row, 0,
        '="Thank you for your business. Any questions, call "&'
        'BizPhone&" or email "&BizEmail&"."', fmts.chrome("note"))
    row += 1
    s.merge(row, 0, LAST_COL,
            "A spreadsheet cannot take card payments. Put a payment "
            "link in your payment details on the Settings tab and it "
            "prints here automatically.", fmts.chrome("note"))

    s.finish(freeze_row=None, print_area=f"A1:H{row + 2}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Job Sheet
# ---------------------------------------------------------------------------

def build_job(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = JOB_NAME, 7
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Printable Job Sheet", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 16)
    s.set_col(1, 1, 22)
    s.set_col(2, 2, 30)
    s.set_col(3, 3, 10)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 12)
    s.set_col(6, 6, 10)
    s.set_col(7, 7, 14)

    row = 4
    row = _pick(s, fmts, book, row, "Job to print", "JOB-0001",
                "JobIDs", "PrintJob",
                "Pick a job. Print it for the cleaner, or print "
                "several by changing the ID and printing again.")
    # Which checklist to print: the service booked, falling back to the
    # one set against the property.
    helper = row - 1
    ws.write_formula(
        helper, 0,
        s.reg.render('=IFERROR(IF(COUNTIF([[Checklists.template]],'
                     'INDEX([[Jobs.service]],MATCH(PrintJob,'
                     '[[Jobs.id]],0)))>0,INDEX([[Jobs.service]],'
                     'MATCH(PrintJob,[[Jobs.id]],0)),INDEX('
                     '[[Properties.checklist]],MATCH(INDEX('
                     '[[Jobs.property]],MATCH(PrintJob,[[Jobs.id]],0)),'
                     '[[Properties.id]],0))),"")'),
        fmts.chrome("note"))
    book.define_name("PrintJobChecklist",
                     f"={quote_sheet(NAME)}!$A${helper + 1}")
    row = _letterhead(s, fmts, row, LAST_COL, "JOB SHEET")

    def job(key: str) -> str:
        return s.reg.render(f'=IFERROR(INDEX([[Jobs.{key}]],'
                            f'MATCH(PrintJob,[[Jobs.id]],0)),"")')

    def prop(key: str) -> str:
        return s.reg.render(
            f'=IFERROR(INDEX([[Properties.{key}]],MATCH(INDEX('
            f'[[Jobs.property]],MATCH(PrintJob,[[Jobs.id]],0)),'
            f'[[Properties.id]],0)),"")')

    pairs = [
        ("Job number", "=PrintJob", "id", "Date", job("date"), "date"),
        ("Client", job("client_name"), "text", "Start", job("start"),
         "time"),
        ("Property", job("property_name"), "text", "Finish",
         job("finish"), "time"),
        ("Service", job("service"), "text", "Hours on site",
         job("hours_each"), "number1"),
        ("Cleaner", job("employee"), "text", "Second cleaner",
         job("employee2"), "text"),
        ("Frequency", job("frequency"), "text", "Cleaners",
         job("crew"), "int"),
    ]
    for left, lf, lstyle, right, rf, rstyle in pairs:
        s.write(row, 0, left, fmts.chrome("label_bold"))
        s.merge(row, 1, 2, None, fmts.cell(lstyle, "plain"))
        ws.write_formula(row, 1, lf, fmts.cell(lstyle, "plain"))
        s.write(row, 4, right, fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.cell(rstyle, "plain"))
        ws.write_formula(row, 5, rf, fmts.cell(rstyle, "plain"))
        row += 1
    row += 1

    # ---------------------------------------------------- getting in ----
    row = s.section(row, "Getting in", LAST_COL)
    for label, formula, height in [
            ("Address", prop("address"), 30),
            ("Access", prop("access"), 18),
            ("Keys", prop("key_status"), 18),
            ("Entry notes", prop("entry_notes"), 40),
            ("Parking", prop("parking"), 24),
            ("Pets", prop("pets"), 18)]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, None, fmts.cell("wrap", "plain"))
        ws.write_formula(row, 1, formula, fmts.cell("wrap", "plain"))
        s.row_height(row, height)
        row += 1
    row += 1

    # ------------------------------------------------- the checklist ----
    row = s.section(row, "The clean", LAST_COL)
    s.merge(row, 0, 5, None, fmts.chrome("note"))
    ws.write_formula(
        row, 0,
        '="Checklist: "&IF(PrintJobChecklist="","none set yet - build '
        'one on the Checklists tab",PrintJobChecklist)&"   -   tick '
        'each line as you finish it"', fmts.chrome("note"))
    s.merge(row, 6, LAST_COL, None, fmts.chrome("note"))
    ws.write_formula(
        row, 6,
        s.reg.render('=IF(PrintJobChecklist="","",COUNTIF('
                     '[[Checklists.template]],PrintJobChecklist)&'
                     '" tasks")'), fmts.chrome("note"))
    row += 1
    row = s.table_head(row, 0, ["Area", "Task", "What good looks like",
                                "Mins", "Done", "Issue?"],
                       [16, 22, 30, 8, 9, 10])
    template = "PrintJobChecklist"
    first_task = row
    for n in range(1, TASK_LINES + 1):
        key = f'{template}&"|{n}"'
        for col, (field, style) in enumerate([
                ("area", "text"), ("task", "text"), ("detail", "wrap"),
                ("minutes", "int")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Checklists.{field}]],'
                             f'MATCH({key},[[Checklists.key]],0)),"")'),
                fmts.cell(style, "plain"))
        s.write(row, 4, None, fmts.chrome("form_box"))
        s.merge(row, 5, 6, None, fmts.chrome("form_box"))
        s.row_height(row, 20)
        row += 1
    last_task = row - 1
    ws.conditional_format(first_task, 0, last_task, 0, {
        "type": "text", "criteria": "containing", "value": "Kitchen",
        "format": fmts.cf("info")})
    s.merge(row, 0, 2, "Planned minutes", fmts.chrome("total_label"))
    ws.write_formula(row, 3, f"=SUM(D{first_task + 1}:D{last_task + 1})",
                     fmts.chrome("total_value", num_format="0"))
    s.merge(row, 4, LAST_COL,
            "If the clean took much longer than this, say so below - "
            "it usually means the price is wrong, not the cleaner.",
            fmts.chrome("note"))
    row += 2

    # --------------------------------------------------- supplies -------
    row = s.section(row, "Supplies and equipment", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(
        row, 0,
        s.reg.render('="Supplies kept on site: "&IFERROR(INDEX('
                     '[[Properties.supplies_onsite]],MATCH(INDEX('
                     '[[Jobs.property]],MATCH(PrintJob,[[Jobs.id]],0)),'
                     '[[Properties.id]],0)),"unknown")&"      Bring '
                     'anything else you need and log it on the Supply '
                     'Usage tab."'), fmts.cell("wrap", "plain"))
    s.row_height(row, 30)
    row += 2

    # ---------------------------------------------------- sign off ------
    row = s.section(row, "Sign off", LAST_COL)
    for label in ["Time in", "Time out", "Total hours"]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 2, None, fmts.chrome("form_line"))
        s.write(row, 4, "Notes", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 24)
        row += 1
    row += 1
    for label in ["Cleaner signature", "Client signature"]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, 3, None, fmts.chrome("form_line"))
        s.write(row, 4, "Date", fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 28)
        row += 1
    row += 1
    s.merge(row, 0, LAST_COL,
            "Anything broken, missing or not right? Write it here, "
            "photograph it, and tell the office the same day.",
            fmts.chrome("note"))
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("form_box"))
    s.row_height(row, 60)
    row += 1

    s.finish(freeze_row=None, print_area=f"A1:H{row + 1}",
             gridlines=False, fit_width=True)
    return s
