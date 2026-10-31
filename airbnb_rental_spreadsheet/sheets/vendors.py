"""`Staff & Vendors` - the people who keep the properties running.

Cleaners, the laundry, a handyman, the plumber you only ring in
February, the co-host who covers you in August.  One row each, with
what they charge, where they work and how much work they have actually
done for you.

The job counts and the money are pulled from the Cleaning Schedule and
the Maintenance Log, so this tab is never out of date and never has to
be filled in twice.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec, max_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col,
                     money_cf, status_cf, yesno)

NAME = "Staff & Vendors"


def spec(cfg: Config) -> SheetSpec:
    cl_cleaner = "[[Cleaning Schedule.cleaner]]"
    cl_cost = "[[Cleaning Schedule.clean_cost]]"
    cl_hours = "[[Cleaning Schedule.hours]]"
    cl_paid = "[[Cleaning Schedule.paid]]"
    cl_comp = "[[Cleaning Schedule.completion]]"
    mt_vendor = "[[Maintenance Log.vendor]]"
    mt_cost = "[[Maintenance Log.cost_used]]"
    mt_open = "[[Maintenance Log.is_open]]"
    mt_paid = "[[Maintenance Log.paid]]"

    _is_me = cl_cleaner + "={id}{r}"
    _last_clean = max_if("[[Cleaning Schedule.clean_num]]", _is_me)

    columns = [
        id_col("name", "VendorPrefix", digits=cfg.digits("vendor"),
               width=10, header="Vendor ID"),
        Col("name", "Name", 22, "text",
            note="A person or a company. Cleaning jobs and maintenance "
                 "jobs both point at this ID.",
            group="Who they are"),
        Col("role", "Role", 18, "text",
            validate=list_from("StaffRoleList"), group="Who they are"),
        Col("company", "Company", 20, "text", group="Who they are"),
        Col("phone", "Phone", 15, "text",
            note="Printed on the turnover checklist, so a cleaner "
                 "standing in a flat with a broken boiler has the "
                 "plumber's number.",
            group="Who they are"),
        Col("email", "Email", 24, "text", group="Who they are"),
        Col("properties", "Covers", 22, "text",
            note="Which properties they work on. Free text - 'all', or "
                 "the names.",
            group="Who they are"),
        Col("availability", "Available", 15, "text",
            validate=list_from("AvailabilityList"), group="Who they are"),

        # ----------------------------------------------------- the deal --
        Col("pay_basis", "Paid", 18, "text",
            validate=list_from("PayBasisList"), group="The deal"),
        Col("rate", "Rate", 11, "money",
            validate=number_check(0, 100000, "Rate"),
            note="Whatever the basis says - per turnover, per hour, per "
                 "month. New cleaning jobs do not read this "
                 "automatically; the property's cleaning cost does "
                 "that, so a cleaner who charges differently per "
                 "property still works.",
            group="The deal"),
        Col("contract", "Contract", 14, "status",
            validate=list_from("ContractStatusList"),
            cf=status_cf(), group="The deal"),
        Col("started", "Working with you since", 14, "date",
            validate=date_check(), hidden=True, group="The deal"),
        Col("insured", "Insured", 9, "flag",
            validate=yesno("Do they carry their own liability cover?"),
            cf=flag_yes_no_cf(),
            note="Public liability cover. Worth asking for, and worth "
                 "writing down when you do.",
            group="The deal"),
        Col("insurance_expiry", "Cover expires", 13, "date",
            validate=date_check(), group="The deal"),
        Col("cover_days", "Days left", 10, "int",
            formula=blank("name",
                          'IF({insurance_expiry}{r}="","",'
                          '{insurance_expiry}{r}-TodayDate)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 30,
                 "fmt": "warn"}],
            hidden=True, group="The deal"),
        Col("keyholder", "Holds keys", 11, "flag",
            validate=yesno(), cf=flag_yes_no_cf(), group="The deal"),

        # ------------------------------------------------- the workload --
        Col("cleans", "Turnovers done", 13, "int",
            formula=blank("name", f'COUNTIFS({cl_cleaner},{{id}}{{r}})'),
            cf=bar_cf(), group="Workload"),
        Col("cleans_done", "Finished properly", 14, "int",
            formula=blank("name", f'COUNTIFS({cl_cleaner},{{id}}{{r}},'
                                  f'{cl_comp},1)'),
            hidden=True, group="Workload"),
        Col("clean_quality", "Completed in full", 14, "percent",
            formula=blank("name",
                          'IF({cleans}{r}<=0,"",'
                          '{cleans_done}{r}/{cleans}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": 0.98,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.9,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0.98,
                 "fmt": "warn"}],
            note="How often every one of the nine turnover tasks got "
                 "ticked. The most useful number on this tab.",
            group="Workload"),
        Col("clean_hours", "Hours on site", 12, "number1",
            formula=blank("name", f'ROUND(SUMIFS({cl_hours},'
                                  f'{cl_cleaner},{{id}}{{r}}),1)'),
            hidden=True, group="Workload"),
        Col("avg_hours", "Avg a turnover", 13, "number1",
            formula=blank("name",
                          'IF({clean_hours}{r}<=0,"",'
                          'ROUND({clean_hours}{r}/MAX(1,{cleans}{r}),1))'),
            hidden=True, group="Workload"),
        Col("jobs", "Repairs done", 12, "int",
            formula=blank("name", f'COUNTIFS({mt_vendor},{{id}}{{r}})'),
            group="Workload"),
        Col("jobs_open", "Repairs open", 12, "int",
            formula=blank("name", f'COUNTIFS({mt_vendor},{{id}}{{r}},'
                                  f'{mt_open},"Yes")'),
            group="Workload"),
        Col("last_job", "Last turnover", 13, "date",
            formula=blank("name",
                          f'IF({{cleans}}{{r}}<=0,"",{_last_clean})'),
            group="Workload"),

        # ---------------------------------------------------- the money --
        Col("paid_cleaning", "Paid for cleaning", 14, "money",
            formula=blank("name", f'ROUND(SUMIFS({cl_cost},'
                                  f'{cl_cleaner},{{id}}{{r}}),2)'),
            group="Money"),
        Col("paid_repairs", "Paid for repairs", 14, "money",
            formula=blank("name", f'ROUND(SUMIFS({mt_cost},'
                                  f'{mt_vendor},{{id}}{{r}}),2)'),
            group="Money"),
        Col("paid_total", "Total paid", 12, "money",
            formula=blank("name",
                          'ROUND({paid_cleaning}{r}+{paid_repairs}{r},2)'),
            cf=money_cf(), group="Money"),
        Col("owed", "Still owed", 11, "money",
            formula=blank(
                "name",
                f'ROUND(SUMIFS({cl_cost},{cl_cleaner},{{id}}{{r}},'
                f'{cl_paid},"<>Yes")+SUMIFS({mt_cost},'
                f'{mt_vendor},{{id}}{{r}},{mt_paid},"<>Yes"),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "warn"}],
            note="Cleaning jobs and repairs that are not marked paid.",
            group="Money"),
        Col("spend_rank", "Rank", 7, "int",
            formula=('=IF({name}{r}="","",COUNTIFS({paid_total_first}:'
                     '{paid_total_last},">"&{paid_total}{r})+'
                     'COUNTIFS({paid_total_first}:{paid_total}{r},'
                     '{paid_total}{r}))'),
            hidden=True, group="Money"),

        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({contract}{r}="Ended","",'
                'IF({contract}{r}="Do not use","Check - marked do not '
                'use",'
                'IF({owed}{r}>0.004,"Unpaid - "&TEXT({owed}{r},"0")&'
                '" outstanding",'
                'IF({cover_days}{r}="",'
                'IF({insured}{r}="Yes","Add - no insurance expiry date",'
                '"OK"),'
                'IF({cover_days}{r}<0,"Expired - insurance cover ran '
                'out",'
                'IF({cover_days}{r}<30,"Due - insurance expires soon",'
                'IF({jobs_open}{r}>0,"Follow up - "&{jobs_open}{r}&'
                '" open repair(s)",'
                'IF({clean_quality}{r}="","OK",'
                'IF({clean_quality}{r}<0.9,"Check - turnovers often '
                'left incomplete","OK")))))))))'),
            cf=alert_cf(), group="Money"),
        Col("notes", "Notes", 28, "wrap", group="Money"),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff & Vendors",
        subtitle="Cleaners, contractors and co-hosts - what they charge "
                 "and what they have actually done",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("vendors", cfg),
        help_text="One row per person or company. The turnover count, "
                  "the repair count and everything you have paid them "
                  "are read straight from the Cleaning Schedule and "
                  "the Maintenance Log, so this tab never needs "
                  "updating by hand.",
        extra_notes=[
            "'Completed in full' is how often a cleaner ticked all "
            "nine turnover tasks. It is the closest thing to a quality "
            "score you will get without standing in the flat yourself.",
            "The rate here is for reference. New cleaning jobs price "
            "themselves from the property's cleaning cost instead, so "
            "a cleaner who charges more for the big house still comes "
            "out right.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Your team", last)
    row = sheet.table_head(row, 0, [
        "On the books", "Active", "Cleaners", "Contractors", "Key holders",
        "Insurance expiring", "Total paid", "Still owed"],
        [12, 9, 10, 12, 12, 17, 12, 11])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Staff & Vendors.name]],"?*")',
        '=COUNTIFS([[Staff & Vendors.contract]],"Active")',
        '=COUNTIFS([[Staff & Vendors.role]],"Cleaner")+'
        'COUNTIFS([[Staff & Vendors.role]],"Head cleaner")',
        '=COUNTIFS([[Staff & Vendors.jobs]],">0")',
        '=COUNTIFS([[Staff & Vendors.keyholder]],"Yes")',
        '=COUNTIFS([[Staff & Vendors.cover_days]],"<30",'
        '[[Staff & Vendors.cover_days]],">=-3650")',
        '=ROUND(SUM([[Staff & Vendors.paid_total]]),2)',
        '=ROUND(SUM([[Staff & Vendors.owed]]),2)',
    ], ["int", "int", "int", "int", "int", "int", "money", "money"])
