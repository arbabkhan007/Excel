"""`Clients` - the client book, and who is quietly drifting away.

Every salon knows its regulars. What a salon rarely knows is which of
last spring's regulars has not been in since, and that is the column
this tab exists for: last visit, days since, due to rebook, lapsing,
gone. Spend, visit count and average ticket come along for the ride.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if, min_if
from ..sample_data import get as sample
from ..writer import list_from, date_check
from .common import (alert_cf, blank, id_col, lookup_staff, owing_cf,
                     rank_desc, status_cf, yesno)

NAME = "Clients"

#: 31-Dec-9999 as a serial - what `min_if` hands back when nothing matched.
NEVER = 2958465


def spec(cfg: Config) -> SheetSpec:
    appt_client = "[[Appointments.client]]"
    appt_date = "[[Appointments.date]]"
    appt_status = "[[Appointments.status]]"

    columns = [
        id_col("name", "ClientPrefix", digits=4, width=11,
               note="Automatic. Bookings, sales and payments all point "
                    "at this ID."),
        Col("name", "Client", 24, "text",
            note="First and last name. Keep one row per person - "
                 "duplicates split their history in two, and the check "
                 "column tells you when you have made one."),
        Col("phone", "Phone", 16, "text"),
        Col("email", "Email", 26, "text"),
        Col("contact", "Contact by", 12, "text",
            validate=list_from("ContactList")),
        Col("joined", "First seen", 11, "date", validate=date_check(),
            note="When they first came in. Leave it blank and the "
                 "workbook uses their first completed appointment."),
        Col("source", "Found you by", 18, "text",
            validate=list_from("SourceList"),
            note="Where this client came from. The marketing tab "
                 "counts new clients by this column."),
        Col("referred_by", "Referred by", 18, "text",
            note="The client ID of whoever sent them. The loyalty tab "
                 "pays the referral credit from it."),
        Col("staff", "Usual stylist", 12, "id",
            validate=list_from("StaffList"),
            note="Staff ID. Used for the rebooking reminders and the "
                 "'my clients' counts on the staff tab."),
        Col("staff_name", "Stylist", 16, "text",
            formula=lookup_staff("name"), hidden=True),
        Col("birthday", "Birthday", 11, "date", validate=date_check(),
            note="Day and month are what matter - the year is up to "
                 "the client. Drives the birthday list on the "
                 "marketing tab."),
        Col("preferences", "Formula & preferences", 34, "wrap",
            note="Colour formula, timings, how they take their coffee, "
                 "the fringe conversation. The thing that makes a "
                 "client feel known."),
        Col("allergies", "Allergies & notes", 26, "wrap"),
        Col("patch_test", "Patch test", 11, "date", validate=date_check(),
            note="Last patch test date. The action column flags it "
                 "when it is older than the period on Settings."),
        Col("marketing_ok", "Marketing OK?", 12, "flag", validate=yesno(),
            note="Only email or text the clients who said yes."),
        Col("photos_ok", "Photos OK?", 11, "flag", validate=yesno()),
        Col("tags", "Tags", 18, "text",
            note="Colour client, bridal, student, blow-dry regular - "
                 "whatever you want to filter on."),
        Col("set_status", "Status", 12, "text",
            validate=list_from("ClientStatusList"), cf=status_cf(),
            note="Leave as Active for a normal client. Archived keeps "
                 "the history but drops them out of the chase lists."),
        Col("notes", "Notes", 26, "wrap"),
        # ----------------------------------------------- visit behaviour --
        Col("visits", "Visits", 8, "int",
            formula=blank(
                "name",
                f'COUNTIFS({appt_client},{{id}}{{r}},'
                f'{appt_status},"Completed")'),
            note="Completed appointments. Cancellations and no shows "
                 "are counted separately."),
        Col("visits_year", "Visits, 12m", 10, "int",
            formula=blank(
                "name",
                f'COUNTIFS({appt_client},{{id}}{{r}},'
                f'{appt_status},"Completed",'
                f'{appt_date},">="&TodayDate-365)'),
            hidden=True),
        Col("cancels", "Cancelled", 10, "int",
            formula=blank(
                "name",
                f'COUNTIFS({appt_client},{{id}}{{r}},'
                f'{appt_status},"Cancelled")'),
            hidden=True),
        Col("no_shows", "No shows", 10, "int",
            formula=blank(
                "name",
                f'COUNTIFS({appt_client},{{id}}{{r}},'
                f'{appt_status},"No show")'),
            cf=[{"type": "cell", "criteria": ">=", "value": 2,
                 "fmt": "bad"}],
            note="Two or more is the point at which most salons start "
                 "asking for a deposit."),
        Col("first_visit", "First visit", 11, "date",
            formula=blank(
                "name",
                'IF({joined}{r}<>"",{joined}{r},'
                'IF({visits}{r}=0,"",' +
                min_if(appt_date, f"{appt_client}={{id}}{{r}}",
                       f'{appt_status}="Completed"') + "))")),
        Col("last_visit", "Last visit", 11, "date",
            formula=blank(
                "name",
                'IF({visits}{r}=0,"",' +
                max_if(appt_date, f"{appt_client}={{id}}{{r}}",
                       f'{appt_status}="Completed"') + ")")),
        Col("days_since", "Days since", 10, "int",
            formula=blank(
                "name",
                'IF({last_visit}{r}="","",TodayDate-{last_visit}{r})'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=InactiveDays", "fmt": "bad"},
                {"type": "cell", "criteria": ">=",
                 "value": "=LapsingDays", "fmt": "warn"}]),
        Col("next_raw", "Next raw", 11, "number",
            formula=blank(
                "name",
                min_if(appt_date, f"{appt_client}={{id}}{{r}}",
                       f'{appt_status}<>"Cancelled"',
                       f'{appt_status}<>"Completed"',
                       f'{appt_status}<>"No show"',
                       f"{appt_date}>=TodayDate")),
            hidden=True),
        Col("next_visit", "Next booked", 11, "date",
            formula=blank(
                "name",
                f'IF({{next_raw}}{{r}}>={NEVER},"",{{next_raw}}{{r}})'),
            note="Their next appointment that has not been cancelled. "
                 "Blank here is the whole point of the rebooking list."),
        Col("gap_days", "Usual gap", 10, "int",
            formula=blank(
                "name",
                'IF({visits}{r}<2,"",ROUND(({last_visit}{r}-'
                '{first_visit}{r})/MAX(1,{visits}{r}-1),0))'),
            note="Average days between visits. A client whose gap is "
                 "stretching is a client on the way out."),
        Col("due_date", "Rebook by", 11, "date",
            formula=blank(
                "name",
                'IF({last_visit}{r}="","",{last_visit}{r}+'
                'IF({gap_days}{r}="",RebookWeeks*7,'
                'MAX(14,{gap_days}{r})))')),
        # ------------------------------------------------------- the money --
        Col("spend", "Lifetime spend", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],'
                '{id}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Net of discounts and refunds, services and retail "
                 "together."),
        Col("spend_service", "On services", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],{id}{r},'
                '[[Sales.kind]],"Service"),2)'),
            hidden=True),
        Col("spend_retail", "On retail", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],{id}{r},'
                '[[Sales.kind]],"Retail"),2)'),
            hidden=True),
        Col("spend_year", "Spend, 12m", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],{id}{r},'
                '[[Sales.date]],">="&TodayDate-365),2)')),
        Col("avg_ticket", "Average ticket", 12, "money",
            formula=blank(
                "name",
                'IF({visits}{r}=0,"",ROUND({spend}{r}/{visits}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetTicket", "fmt": "good"}]),
        Col("yearly_value", "A year is worth", 12, "money",
            formula=blank(
                "name",
                'IF({avg_ticket}{r}="","",IF({gap_days}{r}="","",'
                'ROUND({avg_ticket}{r}*(365/MAX(14,{gap_days}{r})),2)))'),
            note="Average ticket x how many times a year they come. "
                 "This is what losing them costs."),
        Col("balance", "Balance owed", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.balance]],[[Sales.client]],'
                '{id}{r}),2)'),
            cf=owing_cf(),
            note="Everything billed less everything paid, deposits "
                 "included. Positive means they owe you."),
        Col("points", "Loyalty points", 12, "int",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Loyalty.balance]],'
                'MATCH({id}{r},[[Loyalty.client]],0))*1,0)'),
            note="From the loyalty tab. Zero means they are not on the "
                 "scheme yet."),
        # -------------------------------------------------- where they sit --
        Col("segment", "Segment", 13, "status",
            formula=blank(
                "name",
                'IF({set_status}{r}="Archived","Archived",'
                'IF({visits}{r}=0,"Not been yet",'
                'IF({days_since}{r}>=InactiveDays,"Inactive",'
                'IF({days_since}{r}>=LapsingDays,"Lapsing",'
                'IF({visits}{r}=1,"New",'
                'IF({spend_year}{r}>=VipSpend,"VIP",'
                'IF({visits_year}{r}>=4,"Regular","Active")))))))'),
            cf=status_cf(),
            note="Worked out from the visits and the spend, using the "
                 "thresholds on Settings. Nothing to type."),
        Col("retained", "Came back?", 11, "flag",
            formula=blank("name",
                          'IF({visits}{r}>=2,"Yes","No")'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"}],
            hidden=True),
        rank_desc("spend", header="Spend rank"),
        rank_desc("visits", header="Visit rank"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "name",
                'IF(COUNTIF({name_first}:{name_last},{name}{r})>1,'
                '"Duplicate - same name twice",'
                'IF({set_status}{r}="Archived","",'
                'IF({balance}{r}>0.004,"Chase - balance outstanding",'
                'IF({segment}{r}="Inactive","Gone - win-back offer",'
                'IF({segment}{r}="Lapsing","Chase - lapsing, call them",'
                'IF(IF({patch_test}{r}="",FALSE,'
                'TodayDate-{patch_test}{r}>PatchTestDays),'
                '"Due - patch test out of date",'
                'IF(IF({due_date}{r}="",FALSE,'
                'IF({next_visit}{r}<>"",FALSE,'
                '{due_date}{r}<=TodayDate+14)),'
                '"Rebook - due and nothing booked",'
                'IF(IF({birthday}{r}="",FALSE,'
                'ABS(DATE(YEAR(TodayDate),MONTH({birthday}{r}),'
                'DAY({birthday}{r}))-TodayDate)<=7),'
                '"Birthday - send the treat","OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Client CRM",
        subtitle="Who they are, what they have spent, when they were "
                 "last in - and which of them you are about to lose.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).clients,
        help_text="Type the name and the contact details. Visits, spend, "
                  "last visit and the chase list all build themselves "
                  "from the appointments and sales tabs.",
        extra_notes=[
            "The action column is the one to filter on each Monday. "
            "Lapsing, due a rebook, balance outstanding, patch test out "
            "of date - that is your call list for the week.",
            "A year is worth is the number to quote when somebody says "
            "losing one client does not matter. It is the average "
            "ticket times how often they come.",
            "Segment is worked out, not typed. Change the lapsing and "
            "VIP thresholds on Settings and every client re-sorts.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Client book summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The client book", last)
    sheet.table_head(row, 0, [
        "Clients on file", "Active", "New this year", "Lapsing",
        "Inactive", "Average lifetime spend", "Average visits",
        "Owing you money", "Marketing opt-ins"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Clients.name]],"<>")',
        '=COUNTIFS([[Clients.segment]],"Active")+'
        'COUNTIFS([[Clients.segment]],"Regular")+'
        'COUNTIFS([[Clients.segment]],"VIP")+'
        'COUNTIFS([[Clients.segment]],"New")',
        '=COUNTIFS([[Clients.first_visit]],">="&TodayDate-365)',
        '=COUNTIFS([[Clients.segment]],"Lapsing")',
        '=COUNTIFS([[Clients.segment]],"Inactive")',
        '=IFERROR(ROUND(SUM([[Clients.spend]])/'
        'MAX(1,COUNTIF([[Clients.name]],"<>")),2),0)',
        '=IFERROR(ROUND(SUM([[Clients.visits]])/'
        'MAX(1,COUNTIF([[Clients.name]],"<>")),1),0)',
        '=COUNTIFS([[Clients.balance]],">0.004")',
        '=COUNTIFS([[Clients.marketing_ok]],"Yes")',
    ], ["int", "int", "int", "int", "int", "money", "number1", "int",
        "int"])
