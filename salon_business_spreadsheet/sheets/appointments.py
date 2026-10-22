"""`Appointments` - the diary, with a double-booking alarm.

One row per booking. Type the date, the time, the client, the service
and the stylist; the workbook fills in the name, the price, the
duration, the finish time, the deposit position and - the column that
earns its keep - whether that stylist is already booked at that moment.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, blank, id_col, lookup_client, lookup_service,
                     lookup_staff, month_col, status_cf, week_col,
                     weekday_col)

NAME = "Appointments"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "AppointmentPrefix", digits=4, width=11,
               note="Automatic. Payments and sales can quote this ID so "
                    "a deposit lands against the right booking."),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day of the appointment. Everything on this row "
                 "waits for it."),
        weekday_col("date"),
        Col("start", "Start", 9, "time", validate=time_check(),
            note="Type 9:30 or 2:15 PM - Excel stores it as a time."),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList"),
            note="Pick the client ID. A walk-in with no record can be "
                 "left blank - the sale still counts."),
        Col("client_name", "Client", 20, "text",
            formula=lookup_client("name", missing="Walk-in"),
            note="Looked up from the Clients tab."),
        Col("service", "Service ID", 11, "id",
            validate=list_from("ServiceList"),
            note="Pick from the service menu. The price and the "
                 "minutes come with it."),
        Col("service_name", "Service", 24, "text",
            formula=lookup_service("name")),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList")),
        Col("staff_name", "Stylist", 16, "text",
            formula=lookup_staff("name", missing="Unassigned")),
        Col("minutes", "Minutes", 9, "int",
            validate=number_check(0, 1440, decimal=False),
            note="Only fill this in when the booking is longer or "
                 "shorter than the menu says. Otherwise leave it "
                 "blank."),
        Col("price", "Price", 10, "money",
            validate=number_check(0, 100000),
            note="Overrides the menu price - a senior stylist "
                 "surcharge, a quoted colour correction, a friend "
                 "rate. Blank means use the menu."),
        Col("status", "Status", 13, "text",
            validate=list_from("AppointmentStatusList"), cf=status_cf(),
            note="Enquiry, Booked, Confirmed, Arrived, Completed, "
                 "Cancelled, No show, Rescheduled. Only the last three "
                 "and Completed count as a final answer."),
        Col("source", "Booked via", 16, "text",
            validate=list_from("SourceList")),
        Col("deposit_due", "Deposit asked", 12, "money",
            validate=number_check(0, 100000),
            note="Leave blank to use the deposit rule on Settings. "
                 "Type 0 to waive it."),
        Col("notes", "Notes", 28, "wrap"),
        Col("cancel_reason", "If cancelled, why", 22, "text",
            validate=list_from("CancelReasonList")),
        Col("fee_charged", "Late fee charged", 13, "money",
            validate=number_check(0, 10000),
            note="A no-show or late cancellation fee you actually "
                 "collected. Log the money itself on Payments."),
        # ------------------------------------------------- worked out ----
        month_col("date"),
        week_col("date"),
        Col("minutes_used", "Minutes used", 11, "int",
            formula=blank(
                "date",
                'IF({minutes}{r}<>"",{minutes}{r},IFERROR(INDEX('
                '[[Services.duration]],MATCH({service}{r},'
                '[[Services.id]],0))*1,0)+IFERROR(INDEX('
                '[[Services.buffer]],MATCH({service}{r},'
                '[[Services.id]],0))*1,0))'),
            hidden=True),
        Col("finish", "Finishes", 9, "time",
            formula=blank(
                "date",
                'IF({start}{r}="","",{start}{r}+{minutes_used}{r}/1440)'),
            note="Start plus the service time and the turnaround."),
        Col("hour", "Hour", 7, "int",
            formula=blank(
                "date",
                'IF({start}{r}="","",HOUR({start}{r}))'),
            hidden=True,
            note="The hour the appointment starts, for the "
                 "busiest-time analysis."),
        Col("hours", "Chair hours", 10, "number1",
            formula=blank("date", 'ROUND({minutes_used}{r}/60,2)'),
            hidden=True),
        Col("price_used", "Price charged", 11, "money",
            formula=blank(
                "date",
                'IF({price}{r}<>"",{price}{r},IFERROR(INDEX('
                '[[Services.price]],MATCH({service}{r},'
                '[[Services.id]],0))*1,0))'),
            note="Your override if there is one, otherwise the menu "
                 "price. This is what the diary is worth."),
        Col("deposit_wanted", "Deposit wanted", 12, "money",
            formula=blank(
                "date",
                'IF({deposit_due}{r}<>"",{deposit_due}{r},'
                'IF({price_used}{r}>=DepositThreshold,'
                'ROUND({price_used}{r}*DepositPercent,2),0))'),
            hidden=True),
        Col("deposit_paid", "Deposit paid", 12, "money",
            formula=blank(
                "date",
                'ROUND(SUMIFS([[Payments.amount]],'
                '[[Payments.appointment]],{id}{r},'
                '[[Payments.kind]],"Deposit"),2)'),
            note="Deposits taken against this booking. They are money "
                 "received, not revenue - the revenue appears when the "
                 "sale is rung up."),
        Col("billed", "Billed", 11, "money",
            formula=blank(
                "date",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.appointment]],'
                '{id}{r}),2)'),
            note="What was actually rung through the till against this "
                 "appointment, retail included."),
        Col("is_final", "Final?", 8, "flag",
            formula=blank(
                "date",
                'IF({status}{r}="Completed","Yes",'
                'IF({status}{r}="Cancelled","Yes",'
                'IF({status}{r}="No show","Yes","No")))'),
            hidden=True,
            note="Completed, Cancelled and No show are final. "
                 "Everything else is still in play, so it is left out "
                 "of the completion and cancellation rates."),
        Col("lead_days", "Booked ahead", 11, "int",
            formula=blank(
                "date",
                'IF({logged}{r}="","",{date}{r}-{logged}{r})'),
            hidden=True),
        Col("logged", "Booked on", 11, "date", validate=date_check(),
            note="When the booking was made. Optional - it feeds the "
                 "lead-time figure on the analytics tab."),
        Col("rebooked", "Rebooked?", 11, "flag",
            formula=blank(
                "date",
                'IF({status}{r}<>"Completed","",'
                'IF({client}{r}="","No",'
                'IF(COUNTIFS([[Appointments.client]],{client}{r},'
                '[[Appointments.date]],">"&{date}{r},'
                '[[Appointments.status]],"<>Cancelled")>0,"Yes","No")))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"}],
            note="Did this client have anything else in the book after "
                 "this visit? The rebooking rate on the dashboard is "
                 "built from this column."),
        Col("clash", "Clash?", 9, "flag",
            formula=blank(
                "date",
                'IF({staff}{r}="","",IF({start}{r}="","",'
                'IF(COUNTIFS({staff_first}:{staff_last},{staff}{r},'
                '{date_first}:{date_last},{date}{r},'
                '{start_first}:{start_last},"<"&{finish}{r},'
                '{finish_first}:{finish_last},">"&{start}{r},'
                '{status_first}:{status_last},"<>Cancelled")>1,'
                '"Clash","")))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Clash", "fmt": "bad"}],
            note="Counts any other booking for the same stylist on the "
                 "same day whose time overlaps this one. Cancelled "
                 "bookings are ignored."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({clash}{r}="Clash","Clash - stylist double booked",'
                'IF({staff}{r}="","Assign - no stylist yet",'
                'IF({service}{r}="","Add - no service chosen",'
                'IF({status}{r}="No show","No show - charge the fee?",'
                'IF({status}{r}="Cancelled","",'
                'IF(IF({status}{r}="Completed",'
                '{billed}{r}<{price_used}{r}-0.004,FALSE),'
                '"Unpaid - nothing rung up for this visit",'
                'IF(IF({date}{r}<TodayDate,{status}{r}="Booked",FALSE),'
                '"Overdue - past date, still only Booked",'
                'IF(IF({date}{r}=TodayDate,TRUE,FALSE),"Today",'
                'IF(IF({deposit_wanted}{r}>0,'
                '{deposit_paid}{r}<{deposit_wanted}{r}-0.004,FALSE),'
                '"Due - deposit not taken",'
                'IF(IF({date}{r}>TodayDate,'
                '{date}{r}<=TodayDate+2,FALSE),'
                '"Confirm - coming up","OK"))))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Appointment Planner",
        subtitle="The diary: who, with whom, how long, how much - and "
                 "a double-booking alarm that checks every row.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).appointments,
        help_text="Date, start time, client, service, stylist, status. "
                  "The finish time, the price and the clash check look "
                  "after themselves.",
        extra_notes=[
            "Status is the engine of this workbook. Completed means it "
            "happened and can be charged; Cancelled and No show are "
            "the two ways it did not. Enquiry, Booked, Confirmed and "
            "Arrived are all still open, so they are left out of the "
            "completion and cancellation rates rather than counted as "
            "failures.",
            "A deposit is money received, never revenue. It sits on "
            "the Payments tab as kind Deposit, shows here as deposit "
            "paid, and is subtracted from the balance when the sale is "
            "rung up - so it can never be counted twice.",
            "Sorting the diary is safe: every formula finds its own "
            "row. Sort by date and stylist whenever the week gets "
            "messy.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Diary summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The diary in numbers", last)
    sheet.table_head(row, 0, [
        "Appointments logged", "Completed", "Cancelled", "No shows",
        "Completion rate", "Cancellation rate", "No-show rate",
        "Booked value, completed", "Chair hours, completed",
        "Clashes to fix"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Appointments.date]],"<>")',
        '=COUNTIFS([[Appointments.status]],"Completed")',
        '=COUNTIFS([[Appointments.status]],"Cancelled")',
        '=COUNTIFS([[Appointments.status]],"No show")',
        '=IFERROR(COUNTIFS([[Appointments.status]],"Completed")/'
        'MAX(1,COUNTIFS([[Appointments.is_final]],"Yes")),0)',
        '=IFERROR(COUNTIFS([[Appointments.status]],"Cancelled")/'
        'MAX(1,COUNTIFS([[Appointments.is_final]],"Yes")),0)',
        '=IFERROR(COUNTIFS([[Appointments.status]],"No show")/'
        'MAX(1,COUNTIFS([[Appointments.is_final]],"Yes")),0)',
        '=ROUND(SUMIFS([[Appointments.price_used]],'
        '[[Appointments.status]],"Completed"),2)',
        '=ROUND(SUMIFS([[Appointments.hours]],'
        '[[Appointments.status]],"Completed"),1)',
        '=COUNTIFS([[Appointments.clash]],"Clash")',
    ], ["int", "int", "int", "int", "percent", "percent", "percent",
        "money", "number1", "int"])
