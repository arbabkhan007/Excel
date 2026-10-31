"""`Maintenance Log` - what is broken, who is fixing it, what it cost.

Short-term rentals wear out roughly five times faster than a home,
because a hundred different people a year use the shower, the sofa and
the front door.  The jobs that sink a property are rarely the big ones;
they are the twelve small ones nobody wrote down.

Every row carries a priority, a contractor, an estimate and an actual.
The gap between the estimate and the actual is shown, because that gap
is the most useful thing you will ever learn about a tradesperson.
"""

from __future__ import annotations

from ..config import MAINTENANCE_CATEGORIES, Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col,
                     lookup_property, lookup_vendor, money_cf, month_col,
                     owing_cf, priority_cf, status_cf, year_col, yesno)

NAME = "Maintenance Log"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("report_date", "MaintenancePrefix",
               digits=cfg.digits("maintenance"), width=10, header="Job ID"),
        Col("report_date", "Reported", 12, "date",
            validate=date_check(),
            note="When it was noticed. Every row starts here.",
            group="The job"),
        month_col("report_date"),
        year_col("report_date"),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"), group="The job"),
        Col("property_name", "Property name", 19, "text",
            formula=lookup_property("name"), group="The job"),
        Col("category", "Category", 17, "text",
            validate=list_from("MaintenanceCategoryList"),
            group="The job"),
        Col("description", "What is wrong", 34, "wrap", group="The job"),
        Col("priority", "Priority", 12, "status",
            validate=list_from("PriorityList"),
            cf=priority_cf(),
            note="Urgent means a guest cannot use the property. Be "
                 "strict about it or the word stops meaning anything.",
            group="The job"),
        Col("reported_by", "Reported by", 16, "text",
            note="The guest, the cleaner, you, the neighbour.",
            group="The job"),
        Col("booking", "During booking", 12, "text",
            validate=list_from("BookingList"),
            note="Optional. If a particular stay caused it, say so - "
                 "it is the first thing you will want if you ever claim "
                 "on a deposit.",
            hidden=True, group="The job"),

        # ---------------------------------------------------- who fixes --
        Col("vendor", "Contractor", 11, "text",
            validate=list_from("VendorList"), group="Who fixes it"),
        Col("vendor_name", "Contractor name", 18, "text",
            formula=lookup_vendor("name"), group="Who fixes it"),
        Col("vendor_phone", "Phone", 14, "text",
            formula=lookup_vendor("phone"), hidden=True,
            group="Who fixes it"),
        Col("scheduled", "Booked in for", 12, "date",
            validate=date_check(), group="Who fixes it"),
        Col("completed", "Finished", 12, "date",
            validate=date_check(), group="Who fixes it"),
        Col("status", "Status", 16, "status",
            validate=list_from("MaintenanceStatusList"),
            cf=status_cf(), group="Who fixes it"),
        Col("blocks_calendar", "Blocks the calendar", 15, "flag",
            validate=yesno("Does this stop you letting the property?"),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"}],
            note="Yes means you cannot take a booking until it is done. "
                 "Put a Maintenance block on the Booking Tracker to "
                 "hold the dates.",
            group="Who fixes it"),

        # ----------------------------------------------------- the cost --
        Col("est_cost", "Estimate", 11, "money",
            validate=number_check(0, 1000000, "Estimated cost"),
            group="Cost"),
        Col("act_cost", "Actual", 11, "money",
            validate=number_check(0, 1000000, "Actual cost"),
            group="Cost"),
        Col("cost_used", "Cost counted", 12, "money",
            formula=blank(
                "report_date",
                'IF({act_cost}{r}<>"",{act_cost}{r},'
                'IF({status}{r}="Complete",0,0))'),
            cf=money_cf(),
            note="Only the actual figure counts towards profit. An "
                 "estimate is a plan, not a cost.",
            group="Cost"),
        Col("overrun", "Over the estimate", 13, "money",
            formula=blank(
                "report_date",
                'IF({act_cost}{r}="","",IF({est_cost}{r}="","",'
                'ROUND({act_cost}{r}-{est_cost}{r},2)))'),
            cf=owing_cf(),
            note="Actual less estimate. Worth watching per contractor.",
            hidden=True, group="Cost"),
        Col("invoice_ref", "Invoice ref", 13, "text", group="Cost"),
        Col("paid", "Paid", 8, "flag",
            validate=yesno(), cf=flag_yes_no_cf(), group="Cost"),
        Col("warranty_until", "Under warranty to", 14, "date",
            validate=date_check(),
            note="For a new appliance or a job with a guarantee. Saves "
                 "you paying twice.",
            hidden=True, group="Cost"),

        # ----------------------------------------------- where it stands --
        Col("is_open", "Still open", 10, "flag",
            formula=blank(
                "report_date",
                'IF({completed}{r}<>"","No",'
                'IF({status}{r}="Complete","No",'
                'IF({status}{r}="Cancelled","No","Yes")))'),
            cf=flag_yes_no_cf(), hidden=True, group="Where it stands"),
        Col("days_open", "Days open", 10, "int",
            formula=blank(
                "report_date",
                'IF({is_open}{r}="Yes",MAX(0,TodayDate-{report_date}{r}),'
                'IF({completed}{r}="","",'
                'MAX(0,{completed}{r}-{report_date}{r})))'),
            cf=[{"type": "cell", "criteria": ">", "value": 30,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 14,
                 "fmt": "warn"}],
            group="Where it stands"),
        Col("due_in", "Due in", 9, "int",
            formula=blank(
                "report_date",
                'IF({scheduled}{r}="","",IF({is_open}{r}<>"Yes","",'
                '{scheduled}{r}-TodayDate))'),
            hidden=True, group="Where it stands"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "report_date",
                'IF({status}{r}="Cancelled","",'
                'IF({is_open}{r}="No",'
                'IF({paid}{r}<>"Yes",'
                'IF({act_cost}{r}="","Add - no final cost",'
                '"Unpaid - invoice not settled"),"OK"),'
                'IF({priority}{r}="Urgent","Urgent - guest cannot use '
                'the property",'
                'IF({vendor}{r}="","Action - nobody assigned",'
                'IF({scheduled}{r}="","Schedule - no date booked in",'
                'IF({due_in}{r}<0,"Overdue - booked in for '
                '"&TEXT({scheduled}{r},"d mmm"),'
                'IF({due_in}{r}<=MaintenanceWarnDays,"Due - '
                '"&TEXT({scheduled}{r},"d mmm"),'
                'IF({days_open}{r}>30,"Chase - open over a month",'
                '"OK - scheduled"))))))))'),
            cf=alert_cf(), group="Where it stands"),
        Col("notes", "Notes", 30, "wrap", group="Where it stands"),
    ]

    return SheetSpec(
        name=NAME,
        title="Maintenance Log",
        subtitle="What is broken, who is fixing it, when, and what it "
                 "actually cost",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("maintenance", cfg),
        help_text="One row per job, from the moment somebody notices it. "
                  "Only the actual cost counts towards profit - an "
                  "estimate is a plan. If a job stops you letting the "
                  "property, mark it as blocking and put a Maintenance "
                  "block on the Booking Tracker so the dates come off "
                  "the calendar.",
        extra_notes=[
            "Days open counts from the report date, and keeps counting "
            "until the job is finished. Anything over a month goes red, "
            "because a job open that long is usually a job nobody owns.",
            "Keep the Urgent priority for things that stop a guest "
            "using the property. If everything is urgent, nothing is.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "The job list", last)
    row = sheet.table_head(row, 0, [
        "Jobs logged", "Still open", "Urgent open", "Blocking lettings",
        "Avg days to fix", "Estimated", "Actually spent", "Over estimate",
        "Unpaid invoices"],
        [12, 11, 12, 16, 15, 12, 14, 13, 14])
    sheet.calc_row(row, 0, [
        '=COUNT([[Maintenance Log.report_date]])',
        '=COUNTIFS([[Maintenance Log.is_open]],"Yes")',
        '=COUNTIFS([[Maintenance Log.is_open]],"Yes",'
        '[[Maintenance Log.priority]],"Urgent")',
        '=COUNTIFS([[Maintenance Log.is_open]],"Yes",'
        '[[Maintenance Log.blocks_calendar]],"Yes")',
        '=IF(COUNTIFS([[Maintenance Log.is_open]],"No")<=0,"",'
        'ROUND(SUMIFS([[Maintenance Log.days_open]],'
        '[[Maintenance Log.is_open]],"No")/'
        'COUNTIFS([[Maintenance Log.is_open]],"No"),1))',
        '=ROUND(SUM([[Maintenance Log.est_cost]]),2)',
        '=ROUND(SUM([[Maintenance Log.cost_used]]),2)',
        '=ROUND(SUM([[Maintenance Log.overrun]]),2)',
        '=COUNTIFS([[Maintenance Log.is_open]],"No",'
        '[[Maintenance Log.paid]],"<>Yes")',
    ], ["int", "int", "int", "int", "number1", "money", "money", "money",
        "int"])
    row += 2

    row = sheet.section(row, "What keeps breaking", last)
    row = sheet.table_head(row, 0, [
        "Category", "Jobs", "Open", "Spent", "Avg cost", "Avg days",
        "Share of spend"],
        [20, 8, 8, 12, 11, 10, 13])
    total = 'SUM([[Maintenance Log.cost_used]])'
    for name in MAINTENANCE_CATEGORIES:
        key = f'"{name}"'
        cat = f'[[Maintenance Log.category]],{key}'
        spend = f'SUMIFS([[Maintenance Log.cost_used]],{cat})'
        jobs = f'COUNTIFS({cat})'
        closed = f'COUNTIFS({cat},[[Maintenance Log.is_open]],"No")'
        sheet.calc_row(row, 0, [
            name,
            f"={jobs}",
            f'=COUNTIFS({cat},[[Maintenance Log.is_open]],"Yes")',
            f"=ROUND({spend},2)",
            f'=IF({jobs}<=0,"",ROUND({spend}/{jobs},2))',
            f'=IF({closed}<=0,"",ROUND(SUMIFS('
            f'[[Maintenance Log.days_open]],{cat},'
            f'[[Maintenance Log.is_open]],"No")/{closed},1))',
            f'=IF({total}<=0,"",{spend}/{total})',
        ], ["text", "int", "int", "money", "money", "number1", "percent"])
        row += 1
