"""`Invoice Lines` - the detail on each invoice.

Usually one line per job: pick the invoice, pick the job, and the
date, the description and the price come through from the scheduler.
Add a line without a job for anything else you are charging for.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, id_col, lookup_job

NAME = "Invoice Lines"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("invoice", '"IL-"', digits=4, width=10),
        Col("invoice", "Invoice ID", 11, "id",
            validate=list_from("InvoiceIDs"),
            note="Which invoice this line belongs to."),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Link the visit you are billing. Leave blank for a "
                 "line that is not a job - a late fee, materials, a "
                 "deposit."),
        Col("date", "Job date", 11, "date",
            formula=lookup_job("date", missing="")),
        Col("description", "Description", 40, "wrap",
            note="Blank uses the service and the date from the job."),
        Col("qty", "Quantity", 9, "number1",
            validate=number_check(0, 100000),
            note="Blank counts as one."),
        Col("price_input", "Price each (override)", 12, "money",
            validate=number_check(0, 1000000),
            note="Blank uses the price on the job."),
        Col("taxable", "Taxable?", 9, "flag",
            validate=list_from("YesNoList"),
            note="Only matters if you charge tax. The invoice total "
                 "applies the rate from Settings to the whole net "
                 "amount."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("job_price", "Job price", 11, "money",
            formula=lookup_job("price", missing="0"), hidden=True),
        Col("job_service", "Service", 20, "text",
            formula=lookup_job("service", missing=""), hidden=True),
        Col("wording", "Wording used", 40, "wrap",
            formula=blank(
                "invoice",
                'IF({description}{r}<>"",{description}{r},'
                'IF({job}{r}="","",{job_service}{r}&" - "&'
                'TEXT({date}{r},"d mmm yyyy")))'),
            hidden=True,
            note="What the printed invoice says for this line."),
        Col("price_used", "Price each", 11, "money",
            formula=blank("invoice",
                          'IF({price_input}{r}<>"",{price_input}{r},'
                          'IFERROR({job_price}{r}*1,0))')),
        Col("total", "Line total", 12, "money",
            formula=blank("invoice",
                          'ROUND(IF({qty}{r}="",1,{qty}{r})*'
                          '{price_used}{r},2)'),
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("duplicate", "Billed twice?", 11, "status",
            formula=blank(
                "invoice",
                'IF({job}{r}="","",IF(COUNTIFS({job_first}:{job_last},'
                '{job}{r})>1,"Duplicate","OK"))'),
            cf=alert_cf(),
            note="Catches the same job appearing on two invoice "
                 "lines - the most common billing mistake there is."),
        Col("seq", "Line number", 9, "int",
            formula='=IF(OR({invoice}{r}="",{description}{r}=""),"",'
                    'COUNTIFS({invoice_first}:{invoice}{r},'
                    '{invoice}{r}))',
            hidden=True,
            note="1, 2, 3 within each invoice. The printable form "
                 "reads it."),
        Col("key", "Lookup key", 12, "text",
            formula='=IF({seq}{r}="","",{invoice}{r}&"|"&{seq}{r})',
            hidden=True),
        Col("alert", "Check", 24, "status",
            formula=blank(
                "invoice",
                'IF({duplicate}{r}="Duplicate",'
                '"Duplicate - job billed twice",'
                'IF({price_used}{r}=0,"Add - no price on this line",'
                'IF(AND({job}{r}="",{description}{r}=""),'
                '"Add - what is this line for?","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Invoice Lines",
        subtitle="One row per job (or per extra) on each invoice.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).invoice_lines,
        help_text="Pick the invoice, pick the job. The description "
                  "and the price come from the scheduler, and the "
                  "total goes straight back onto the invoice.",
        extra_notes=[
            "The 'Billed twice?' column is worth its place on its "
            "own. Billing the same visit on two invoices is both "
            "embarrassing and expensive.",
            "Charging a late fee or a deposit? Leave the job blank, "
            "write your own description and type the amount.",
        ],
    )
