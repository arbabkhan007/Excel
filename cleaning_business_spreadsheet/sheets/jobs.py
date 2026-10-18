"""`Jobs` - the job scheduler, and the spine of the whole workbook.

One row per visit. Pick the client, the property and the service and
the rest fills itself in: how long it should take, what it should cost
the client, who is free, when the next visit in the series falls and
whether you have double-booked anybody.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_client,
                     lookup_property, month_col, score_validation,
                     status_cf, week_col, weekday_col)

NAME = "Jobs"

# ---------------------------------------------------------------- lookups --
_SVC_HOURS = ('IFERROR(INDEX([[Services.hours]],MATCH({service}{r},'
              '[[Services.name]],0)),0)')
_SVC_MODEL = ('IFERROR(INDEX([[Services.model]],MATCH({service}{r},'
              '[[Services.name]],0)),"")')
_SVC_PRICE = ('IFERROR(INDEX([[Services.standard_price]],MATCH('
              '{service}{r},[[Services.name]],0)),0)')
_PROP_HOURS = ('IFERROR(INDEX([[Properties.est_hours]],MATCH('
               '{property}{r},[[Properties.id]],0)),0)')
_CLIENT_RATE = ('IFERROR(INDEX([[Clients.rate]],MATCH({client}{r},'
                '[[Clients.id]],0)),0)')
_FREQ_FACTOR = ('IFERROR(INDEX(FreqFactors,MATCH({frequency}{r},'
                'FreqNames,0)),1)')
_FREQ_DAYS = ('IFERROR(INDEX(FreqDays,MATCH({frequency}{r},FreqNames,0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("client", "JobPrefix", digits=4, width=11,
               note="The job reference. Checklists, timesheets, "
                    "invoices and the costing tab all key off it."),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day of the visit. Everything in the calendar, "
                 "the rota and the reminders reads this column."),
        weekday_col("date"),
        week_col("date"),
        month_col("date"),
        Col("start", "Start", 9, "time", validate=time_check(),
            note="Type 9:00 AM or 14:30 - either works."),
        Col("client", "Client ID", 11, "id", validate=list_from("ClientIDs"),
            note="Pick from the drop down. The client's name, their "
                 "rate and their terms all follow."),
        Col("client_name", "Client", 22, "text",
            formula=lookup_client("name")),
        Col("client_type", "Client type", 16, "text",
            formula=lookup_client("type", missing=""), hidden=True,
            note="Residential or commercial, pulled from the client. "
                 "The Reports tab splits the money by it."),
        Col("property", "Property ID", 11, "id",
            validate=list_from("PropertyIDs"),
            note="Which address. Leave blank for a client with only "
                 "one property and the price still works."),
        Col("property_name", "Property", 24, "text",
            formula=lookup_property("name", missing="")),
        Col("address", "Address", 30, "text",
            formula=lookup_property("address", missing=""),
            note="Pulled through for the job sheet and the cleaner's "
                 "list."),
        Col("service", "Service", 22, "text",
            validate=list_from("ServiceNames"),
            note="From the Services tab. The price and the time come "
                 "from there unless you override them."),
        Col("type", "Job type", 16, "text",
            validate=list_from("JobTypeList")),
        Col("frequency", "Frequency", 14, "text",
            validate=list_from("FrequencyList"),
            note="One-off, weekly, fortnightly. Drives the discount "
                 "and the next visit date."),
        Col("status", "Status", 13, "status",
            validate=list_from("JobStatusList"), cf=status_cf(),
            note="Move it along as the day goes: Booked, Confirmed, "
                 "In progress, Completed. Invoiced happens when you "
                 "raise the invoice."),
        Col("priority", "Priority", 11, "text",
            validate=list_from("PriorityList")),
        Col("employee", "Cleaner", 18, "text",
            validate=list_from("EmployeeNames"),
            note="The lead cleaner. A second name goes in the next "
                 "column; anyone else goes on the timesheet."),
        Col("employee2", "Second cleaner", 18, "text",
            validate=list_from("EmployeeNames")),
        Col("cleaners", "Cleaners", 9, "int",
            validate=number_check(1, 20, decimal=False),
            note="Blank uses the standard crew size from the Settings "
                 "tab."),
        Col("hours_input", "Planned hours (override)", 12, "number1",
            validate=number_check(0, 100),
            note="Total cleaner-hours. Leave blank and the workbook "
                 "uses the service or the property estimate."),
        Col("price_input", "Price (override)", 12, "money",
            validate=number_check(0, 1000000),
            note="Leave blank to let the calculator price it. Type a "
                 "number to fix the price for this visit only."),
        Col("travel", f"Travel ({cfg.distance_unit}, round trip)", 11,
            "number1", validate=number_check(0, 10000)),
        Col("recurring", "Recurring?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Yes means the workbook should suggest the next "
                 "visit date when this one is done."),
        Col("confirmed", "Confirmed with client?", 11, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("actual_input", "Actual hours (if no timesheet)", 12,
            "number1", validate=number_check(0, 100),
            note="Only needed if you are not filling in timesheets. "
                 "Timesheet hours always win."),
        Col("rating", "Client rating", 10, "rating",
            validate=score_validation(5),
            note="1 to 5, filled in after the visit. Feeds the client "
                 "and employee averages."),
        Col("feedback", "Feedback", 30, "wrap"),
        Col("notes", "Job notes", 34, "wrap",
            note="What is different about this visit: the spare room "
                 "is out of bounds, the dog is at the vet."),
        # ------------------------------------------------- the automatics --
        Col("est_hours", "Hours (total)", 11, "number1",
            formula=blank(
                "client",
                'IF({hours_input}{r}<>"",{hours_input}{r},'
                'IF(AND(' + _SVC_MODEL + '="Per hour",' + _PROP_HOURS +
                '>0),' + _PROP_HOURS + ',IF(' + _SVC_HOURS + '>0,' +
                _SVC_HOURS + ',' + _PROP_HOURS + ')))'),
            note="Total cleaner-hours the visit should take. Hourly "
                 "services use the property estimate; fixed price "
                 "services use the time on the Services tab."),
        Col("crew", "Crew used", 8, "int",
            formula=blank("client",
                          'IF({cleaners}{r}="",CleanersPerJob,'
                          'MAX(1,{cleaners}{r}))'),
            hidden=True),
        Col("hours_each", "Hours on site", 11, "number1",
            formula=blank("client",
                          'ROUND({est_hours}{r}/{crew}{r},2)'),
            note="How long the visit lasts: total hours shared "
                 "between the crew. This is what blocks out the "
                 "diary.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("finish", "Finish", 9, "time",
            formula='=IF(OR({client}{r}="",{start}{r}="",'
                    '{hours_each}{r}=""),"",{start}{r}+'
                    '{hours_each}{r}/24)',
            note="Start plus the time on site. Add the travel "
                 "allowance from Settings before you book the next "
                 "one."),
        Col("rate_used", "Rate used", 10, "money",
            formula=blank("client",
                          'IF(' + _CLIENT_RATE + '>0,' + _CLIENT_RATE +
                          ',HourlyRate)'),
            hidden=True,
            note="The client's agreed rate if they have one, "
                 "otherwise the standard rate."),
        Col("base_price", "Base price", 11, "money",
            formula=blank(
                "client",
                'IF(AND({service}{r}<>"",' + _SVC_PRICE + '>0,' +
                _SVC_MODEL + '<>"Per hour"),' + _SVC_PRICE + ','
                '{est_hours}{r}*{rate_used}{r})'),
            hidden=True),
        Col("price", "Price", 12, "money",
            formula=blank(
                "client",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF({status}{r}="Cancelled",0,'
                'MAX(MinimumCharge,ROUND({base_price}{r}*' +
                _FREQ_FACTOR + ',2))))'),
            note="What the client pays. Worked out from the service, "
                 "the hours, the client's rate and the frequency "
                 "discount - unless you typed an override.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("actual_hours", "Actual hours", 11, "number1",
            formula=blank(
                "client",
                'IF(SUMIFS([[Timesheets.hours]],[[Timesheets.job]],'
                '{id}{r})>0,SUMIFS([[Timesheets.hours]],'
                '[[Timesheets.job]],{id}{r}),'
                'IF({actual_input}{r}="","",{actual_input}{r}))'),
            note="Timesheet hours for this job if there are any, "
                 "otherwise what you typed in the override column."),
        Col("over_run", "Over / under", 10, "number1",
            formula='=IF(OR({actual_hours}{r}="",{est_hours}{r}=""),"",'
                    '{actual_hours}{r}-{est_hours}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 0.25,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}],
            note="Hours over the estimate. Consistently positive "
                 "means the price is wrong, not the cleaner."),
        Col("tasks", "Checklist", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Job Tasks.job]],'
                    '{id}{r}))',
            hidden=True),
        Col("tasks_done", "Tasks done", 10, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Job Tasks.job]],'
                    '{id}{r},[[Job Tasks.status]],"Done"))',
            hidden=True),
        Col("checklist_pct", "Checklist complete", 11, "percent0",
            formula='=IF(OR({id}{r}="",{tasks}{r}=0),"",'
                    '{tasks_done}{r}/{tasks}{r})',
            cf=[{"type": "data_bar", "bar_color": "#4E8A63"}],
            note="Filled in from the Job Tasks tab as the cleaner "
                 "ticks things off."),
        Col("invoice", "Invoice", 11, "id",
            formula='=IF({id}{r}="","",IFERROR(INDEX('
                    '[[Invoice Lines.invoice]],MATCH({id}{r},'
                    '[[Invoice Lines.job]],0)),""))',
            note="Found automatically once the job appears on an "
                 "invoice line."),
        Col("next_due", "Next visit due", 12, "date",
            formula='=IF(OR({client}{r}="",{date}{r}="",'
                    '{recurring}{r}<>"Yes"),"",IF(' + _FREQ_DAYS +
                    '=0,"",{date}{r}+' + _FREQ_DAYS + '))',
            note="Date plus the gap for this frequency. Copy the row "
                 "down and paste this date in to book the next one.",
            cf=[{"type": "cell", "criteria": "<", "value": "=TodayDate",
                 "fmt": "warn"}]),
        Col("clash", "Diary", 16, "status",
            formula=blank(
                "client",
                'IF({employee}{r}="","Unassigned",'
                'IF(COUNTIFS({employee_first}:{employee_last},'
                '{employee}{r},{date_first}:{date_last},{date}{r},'
                '{start_first}:{start_last},{start}{r})>1,"Clash",'
                'IF(SUMIFS({hours_each_first}:{hours_each_last},'
                '{employee_first}:{employee_last},{employee}{r},'
                '{date_first}:{date_last},{date}{r})>ProductiveHours,'
                '"Long day","OK")))'),
            cf=alert_cf(),
            note="Two jobs at the same time for the same cleaner is a "
                 "Clash. More than a full day's work is a Long day."),
        Col("future", "Upcoming date", 11, "date",
            formula='=IF(OR({client}{r}="",{date}{r}=""),"",'
                    'IF(AND({date}{r}>=TodayDate,{status}{r}<>'
                    '"Cancelled"),{date}{r},""))',
            hidden=True,
            note="The date, but only for visits still to come. The "
                 "calendar and the dashboard read it."),
        Col("future_rank", "Diary order", 9, "int",
            formula='=IF({future}{r}="","",COUNTIFS({future_first}:'
                    '{future_last},"<"&{future}{r})+COUNTIFS('
                    '{future_first}:{future}{r},{future}{r}))',
            hidden=True,
            note="1 = the next job in the diary. No array formulas "
                 "anywhere in this workbook."),
        Col("unassigned", "Needs a cleaner", 11, "date",
            formula='=IF(OR({client}{r}="",{date}{r}="",'
                    '{employee}{r}<>""),"",IF(AND({date}{r}>=TodayDate,'
                    '{status}{r}<>"Cancelled"),{date}{r},""))',
            hidden=True),
        Col("unassigned_rank", "Assign order", 9, "int",
            formula='=IF({unassigned}{r}="","",COUNTIFS('
                    '{unassigned_first}:{unassigned_last},"<"&'
                    '{unassigned}{r})+COUNTIFS({unassigned_first}:'
                    '{unassigned}{r},{unassigned}{r}))',
            hidden=True,
            note="1 = the soonest job with nobody on it."),
        Col("days_away", "Days away", 9, "int",
            formula='=IF(OR({client}{r}="",{date}{r}=""),"",'
                    '{date}{r}-TodayDate)',
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "client",
                'IF({status}{r}="Cancelled","Cancelled",'
                'IF({status}{r}="No access","Action - no access, rebook",'
                'IF(AND({date}{r}<>"",{days_away}{r}<0,'
                'OR({status}{r}="Booked",{status}{r}="Confirmed",'
                '{status}{r}="In progress")),"Late - mark it completed",'
                'IF(AND({status}{r}="Completed",{invoice}{r}=""),'
                '"Action - not invoiced yet",'
                'IF(AND({days_away}{r}>=0,{days_away}{r}<=ReminderDays,'
                '{confirmed}{r}<>"Yes"),"Confirm - visit coming up",'
                'IF({employee}{r}="","Assign - no cleaner booked",'
                'IF(AND({status}{r}="Completed",{checklist_pct}{r}<>"",'
                '{checklist_pct}{r}<1),"Check - checklist unfinished",'
                '"OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Job Scheduler",
        subtitle="Every visit you have booked, priced automatically "
                 "and checked for clashes.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).jobs,
        help_text="Pick a client, a property and a service, give it a "
                  "date and a cleaner, and the hours, the finish time "
                  "and the price work themselves out. The Calendar tab "
                  "shows the same rows as a diary.",
        extra_notes=[
            "To book a recurring client, fill one row in, set "
            "'Recurring? = Yes' and read the next visit date in the "
            "last columns - copy the row down and paste that date in. "
            "A spreadsheet cannot create rows on its own, so this is "
            "the honest way to do it.",
            "The Diary column is a real double-booking check: it "
            "looks for the same cleaner, the same date and the same "
            "start time anywhere on this tab.",
            "A job only becomes revenue when its status is Completed "
            "or Invoiced. Cancelled jobs are priced at zero and drop "
            "out of every total.",
        ],
    )
