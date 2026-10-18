"""`Clients` - the customer database every other tab points at.

One row per client. You type who they are and how they like to be
cleaned for; the workbook adds up what they have spent, when you last
saw them, what their work actually earns you and whether they are
quietly drifting away.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if, min_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, rank_desc,
                     status_cf)

NAME = "Clients"

_JOBS = "[[Jobs.client]]"
_INV = "[[Invoices.client]]"

#: the next visit in the diary for this client
_NEXT = min_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
               "[[Jobs.date]]>=TodayDate",
               '[[Jobs.status]]<>"Cancelled"')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ClientPrefix", digits=4, width=10,
               note="The ID the properties, jobs, quotes and invoices all "
                    "refer to. Generated for you - never type over it."),
        Col("name", "Client", 26, "text",
            note="A household name or a company name. One row per "
                 "paying customer; their addresses go on the Properties "
                 "tab."),
        Col("type", "Type", 16, "text", validate=list_from("ClientTypeList"),
            note="Residential and commercial work behave very "
                 "differently - the reports split on this."),
        Col("contact", "Main contact", 18, "text"),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 26, "text"),
        Col("address", "Billing address", 34, "wrap",
            note="Where the invoice goes. The place you clean goes on "
                 "the Properties tab - often not the same address."),
        Col("status", "Status", 13, "status",
            validate=list_from("ClientStatusList"), cf=status_cf(),
            note="Prospect until they book. Mark them Paused rather "
                 "than deleting them - the history is worth keeping."),
        Col("source", "Came from", 16, "text",
            validate=list_from("LeadSourceList"),
            note="Where they heard about you. Fill this in every time "
                 "and the Reports tab tells you where to spend your "
                 "marketing money."),
        Col("since", "Client since", 12, "date", validate=date_check()),
        Col("frequency", "Usual frequency", 14, "text",
            validate=list_from("FrequencyList"),
            note="Drives the recurring booking dates and the value of "
                 "the client over a year."),
        Col("pref_day", "Preferred day", 12, "text",
            validate=list_from("WeekdayList")),
        Col("pref_time", "Preferred time", 14, "text",
            validate=list_from("TimeSlotList")),
        Col("pref_cleaner", "Preferred cleaner", 16, "text",
            validate=list_from("EmployeeNames"),
            note="The scheduler flags a job given to anybody else, so "
                 "you can at least tell the client."),
        Col("rate", "Agreed hourly rate", 13, "money",
            validate=number_check(0, 10000),
            note="Leave blank to use the standard rate from Settings. "
                 "Fill it in for a client on a special price."),
        Col("terms", "Payment terms (days)", 12, "int",
            validate=number_check(0, 180, decimal=False),
            note="Blank uses the default terms from the Settings tab."),
        Col("payment_method", "Pays by", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("contract", "Contract", 14, "status",
            validate=list_from("ContractTermList"), cf=status_cf()),
        Col("contract_end", "Contract ends", 12, "date",
            validate=date_check(),
            note="The renewal alert counts down to this date."),
        Col("vip", "VIP?", 7, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("marketing", "Happy to be contacted?", 11, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Keep a record of consent before you add anyone to a "
                 "mailing list."),
        Col("notes", "Notes", 40, "wrap",
            note="Allergies, pets, gate codes, the mug they do not want "
                 "moved. The things that keep a client."),
        # ------------------------------------------------- what they do --
        Col("properties", "Properties", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Properties.client]],'
                    '{id}{r}))',
            note="Counted from the Properties tab."),
        Col("jobs_booked", "Jobs booked", 11, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _JOBS + ',{id}{r}))'),
        Col("jobs_done", "Jobs done", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _JOBS + ',{id}{r},'
                    '[[Jobs.status]],"Completed")+COUNTIFS(' + _JOBS +
                    ',{id}{r},[[Jobs.status]],"Invoiced"))'),
        Col("cancelled", "Cancelled", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _JOBS + ',{id}{r},'
                    '[[Jobs.status]],"Cancelled")+COUNTIFS(' + _JOBS +
                    ',{id}{r},[[Jobs.status]],"No access"))',
            cf=[{"type": "cell", "criteria": ">=", "value": 3,
                 "fmt": "warn"}]),
        Col("last_job", "Last clean", 12, "date",
            formula='=IF({id}{r}="","",IF(' +
                    max_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
                           '([[Jobs.status]]="Completed")+'
                           '([[Jobs.status]]="Invoiced")') +
                    '=0,"",' +
                    max_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
                           '([[Jobs.status]]="Completed")+'
                           '([[Jobs.status]]="Invoiced")') + '))',
            note="The last visit that actually happened."),
        Col("days_since", "Days since", 10, "int",
            formula='=IF({last_job}{r}="","",TodayDate-{last_job}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": "=InactiveClientDays",
                 "fmt": "bad"}]),
        Col("next_job", "Next booking", 12, "date",
            formula='=IF({id}{r}="","",IF(' + _NEXT + '=2958465,"",' +
                    _NEXT + '))',
            note="The next visit in the diary. Empty means nothing is "
                 "booked - that is the list to work from."),
        # ------------------------------------------------- what they pay --
        Col("invoiced", "Invoiced to date", 14, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Invoices.total]],' + _INV +
                    ',{id}{r}))',
            note="Everything you have billed them, lifetime."),
        Col("received", "Paid to date", 13, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Invoices.paid]],' + _INV +
                    ',{id}{r}))'),
        Col("outstanding", "Owing now", 12, "money",
            formula='=IF({id}{r}="","",{invoiced}{r}-{received}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("avg_job", "Average job", 12, "money",
            formula='=IF(OR({id}{r}="",{jobs_done}{r}=0),"",'
                    '{invoiced}{r}/{jobs_done}{r})'),
        Col("profit", "Profit to date", 13, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Job Costing.profit]],'
                    '[[Job Costing.client]],{id}{r}))',
            note="Revenue less wages, supplies and travel, from the Job "
                 "Costing tab. The number that tells you whether a "
                 "client is worth keeping.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({id}{r}="",{invoiced}{r}=0),"",'
                    '{profit}{r}/{invoiced}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}]),
        Col("year_value", "Value a year", 13, "money",
            formula='=IF(OR({id}{r}="",{frequency}{r}="",{avg_job}{r}=""),'
                    '"",{avg_job}{r}*IFERROR(INDEX(FreqPerYear,'
                    'MATCH({frequency}{r},FreqNames,0)),1))',
            note="What this client is worth over twelve months at their "
                 "usual frequency. Losing one of these is worth a phone "
                 "call to keep."),
        Col("rating", "Average rating", 11, "number1",
            formula='=IF({id}{r}="","",IFERROR(AVERAGEIFS('
                    '[[Jobs.rating]],' + _JOBS + ',{id}{r},'
                    '[[Jobs.rating]],">0"),""))',
            cf=[{"type": "cell", "criteria": "<", "value": 3.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 4.5,
                 "fmt": "good"}]),
        Col("complaints", "Complaints", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Complaints.client]],'
                    '{id}{r}))',
            cf=[{"type": "cell", "criteria": ">=", "value": 2,
                 "fmt": "warn"}]),
        rank_desc("invoiced", guard='{name}{r}=""',
                  note="1 = your biggest client by money billed."),
        rank_desc("profit", guard='{name}{r}=""',
                  note="1 = your most profitable client."),
        rank_desc("year_value", guard='{name}{r}=""',
                  note="1 = the client worth most over a year."),
        Col("retention", "Standing", 12, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Lost","Lost",'
                'IF({status}{r}="Inactive","Dormant",'
                'IF({last_job}{r}="","Dormant",'
                'IF({days_since}{r}>InactiveClientDays*2,"Dormant",'
                'IF({days_since}{r}>InactiveClientDays,"At risk",'
                'IF({next_job}{r}="","Slipping",'
                'IF({jobs_done}{r}>=10,"Loyal","Steady")))))))'),
            cf=status_cf(),
            note="Worked out from how long it has been and whether "
                 "anything is in the diary. 'Slipping' means they are "
                 "still a client but nothing is booked."),
        Col("risk", "Risk score", 9, "number1",
            formula='=IF(OR({name}{r}="",{retention}{r}="Lost",'
                    '{retention}{r}="Dormant"),"",'
                    'IF(OR({retention}{r}="At risk",'
                    '{retention}{r}="Slipping"),'
                    'IF({days_since}{r}="",999,{days_since}{r}),""))',
            hidden=True,
            note="Days since the last clean, but only for clients "
                 "who are drifting. The Retention tab ranks on it."),
        Col("risk_rank", "Risk order", 9, "int",
            formula='=IF({risk}{r}="","",COUNTIFS({risk_first}:'
                    '{risk_last},">"&{risk}{r})+COUNTIFS('
                    '{risk_first}:{risk}{r},{risk}{r}))',
            hidden=True, note="1 = the client you should ring first."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({outstanding}{r}>0.5,"Chase - "&TEXT({outstanding}{r},'
                '"0")&" owing",'
                'IF(AND({status}{r}<>"Lost",{status}{r}<>"Inactive",'
                '{status}{r}<>"Paused",{next_job}{r}="",'
                '{frequency}{r}<>"One-off"),"Book - nothing in the diary",'
                'IF(AND({contract_end}{r}<>"",{contract_end}{r}-TodayDate'
                '<=RenewalWarningDays,{contract_end}{r}>=TodayDate),'
                '"Due - contract renews soon",'
                'IF({retention}{r}="At risk","Call - going quiet",'
                'IF({complaints}{r}>=2,"Check - complaints logged",'
                'IF(AND({margin}{r}<>"",{margin}{r}<MarginFloor,'
                '{jobs_done}{r}>0),"Check - thin margin","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Client Database",
        subtitle="Every customer, what they have spent, what they earn "
                 "you and when you last saw them.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).clients,
        help_text="Add the client here first, then their address on the "
                  "Properties tab, then book the work on the Jobs tab. "
                  "Everything downstream keys off the client ID in "
                  "column A.",
        extra_notes=[
            "The grey columns are all worked out for you - jobs done, "
            "money owing, profit, average rating and whether the "
            "client is drifting. There is nothing to update by hand.",
            "'Value a year' multiplies the average job by how often "
            "they book. Sort on it and you can see at a glance which "
            "five clients your business actually rests on.",
            "Client records hold names, addresses, phone numbers and "
            "sometimes keys or alarm codes. Keep the file somewhere "
            "sensible, and delete what you no longer need.",
        ],
    )
