"""`Waitlist` - the enquiries you have not turned into enrollments yet.

A childcare waitlist is a sales pipeline with nappies.  Somebody calls,
you send information, they tour, you offer, they accept or they do not.
Every stage has a date, every stage has a drop-out rate, and the
difference between a setting that fills its rooms and one that does not
is usually nothing more than following the list up.

The position number ranks the people still waiting by priority first
and then by how long they have waited, which is how most settings say
they allocate places - and now they can prove it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (age_months, age_years_text, alert_cf, blank,
                     id_col, status_cf)

NAME = "Waitlist"

#: Statuses that mean the enquiry is finished with, one way or another.
CLOSED = ["Enrolled", "Offer declined", "Lost", "Withdrawn"]


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("child_name", "WaitlistPrefix", digits=4, width=10,
               header="Enquiry ID"),
        Col("enquiry_date", "Enquiry date", 12, "date",
            validate=date_check(),
            note="The day they first got in touch. The waiting time "
                 "and the position in the queue both measure from "
                 "here."),
        Col("child_name", "Child's name", 20, "text"),
        Col("dob", "Date of birth", 12, "date", validate=date_check(),
            note="Needed to work out which room they would go into, "
                 "and when."),
        Col("age_months", "Age (months)", 10, "int",
            formula="=" + age_months("{dob}{r}"), hidden=True),
        Col("age", "Age now", 9, "text",
            formula="=" + age_years_text("{dob}{r}")),
        Col("age_at_start", "Age at start", 11, "text",
            formula=blank(
                "start_wanted",
                age_years_text("{dob}{r}", "{start_wanted}{r}")),
            hidden=True),
        Col("band_wanted", "Age band then", 13, "text",
            formula=blank(
                "dob",
                'IFERROR(INDEX(AgeBandList,MATCH('
                'DATEDIF({dob}{r},MAX({dob}{r},'
                'IF({start_wanted}{r}="",TodayDate,'
                '{start_wanted}{r})),"m"),AgeBandFrom,1))&"","")'),
            note="The band the child will be in on the day they want "
                 "to start - which is not always the band they are in "
                 "today."),
        Col("parent_name", "Parent / guardian", 20, "text"),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 24, "text"),
        Col("source", "Heard about you via", 18, "text",
            validate=list_from("InquirySourceList"),
            note="The single most useful marketing column in the "
                 "workbook. Fill it in every time."),
        # -------------------------------------------------- what they want
        Col("program", "Programme wanted", 15, "text",
            validate=list_from("ProgramList")),
        Col("room", "Room wanted", 16, "text",
            validate=list_from("ClassroomList"),
            note="Pick a room and the workbook tells you whether there "
                 "is a seat in it right now."),
        Col("days_wanted", "Days a week", 10, "number1",
            validate=number_check(0, 7)),
        Col("start_wanted", "Wants to start", 12, "date",
            validate=date_check()),
        Col("priority", "Priority", 13, "text",
            validate=list_from("WaitlistPriorityList"),
            note="Siblings and staff children usually come first. "
                 "Whatever your policy is, record it here and the "
                 "queue position follows it."),
        Col("priority_rank", "Priority order", 10, "int",
            formula=blank(
                "child_name",
                'IF({priority}{r}="Sibling",1,'
                'IF({priority}{r}="Staff child",2,'
                'IF({priority}{r}="Returning family",3,'
                'IF({priority}{r}="Low",5,4))))'),
            hidden=True),
        Col("sibling", "Sibling already here", 11, "id",
            validate=list_from("ChildList"),
            note="The child ID of a brother or sister already with "
                 "you. Links the enquiry to the family you know."),
        Col("sibling_name", "Sibling", 16, "text",
            formula=blank(
                "sibling",
                'IFERROR(INDEX([[Child Enrollment.name]],'
                'MATCH({sibling}{r},[[Child Enrollment.id]],0))&"","")'),
            hidden=True),
        # ------------------------------------------------------ pipeline
        Col("status", "Status", 15, "status",
            validate=list_from("WaitlistStatusList"), cf=status_cf()),
        Col("info_sent", "Info sent", 11, "date", validate=date_check(),
            hidden=True),
        Col("tour_date", "Tour", 11, "date", validate=date_check(),
            note="Settings that get families through the door convert "
                 "two to three times better. Book the tour."),
        Col("offered_date", "Offered", 11, "date",
            validate=date_check()),
        Col("response_date", "They replied", 11, "date",
            validate=date_check(), hidden=True),
        Col("enrolled_child", "Became child", 11, "id",
            validate=list_from("ChildList"),
            note="Once they enroll, put the new child ID here. It "
                 "closes the loop between marketing and revenue."),
        Col("closed", "Closed?", 9, "flag",
            formula=blank(
                "child_name",
                'IF(' + "+".join(f'({{status}}{{r}}="{s}")'
                                 for s in CLOSED) + '>0,"Yes","No")'),
            hidden=True),
        Col("converted", "Enrolled?", 10, "flag",
            formula=blank(
                "child_name",
                'IF({status}{r}="Enrolled","Yes",'
                'IF({enrolled_child}{r}<>"","Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("days_waiting", "Days waiting", 11, "int",
            formula=blank(
                "enquiry_date",
                'IF({closed}{r}="Yes",'
                'IF({response_date}{r}<>"",'
                '{response_date}{r}-{enquiry_date}{r},'
                'IF({offered_date}{r}<>"",'
                '{offered_date}{r}-{enquiry_date}{r},'
                'TodayDate-{enquiry_date}{r})),'
                'TodayDate-{enquiry_date}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 90,
                 "fmt": "warn"}]),
        Col("days_quiet", "Days since contact", 12, "int",
            formula=blank(
                "enquiry_date",
                'IF({closed}{r}="Yes","",'
                'TodayDate-MAX({enquiry_date}{r},{info_sent}{r},'
                '{tour_date}{r},{offered_date}{r},'
                '{response_date}{r}))'),
            cf=[{"type": "cell", "criteria": ">=", "value": 14,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 7,
                 "fmt": "warn"}],
            note="How long since anything happened on this enquiry. "
                 "Over a week and it is going cold."),
        Col("position", "Queue position", 12, "int",
            formula=blank(
                "child_name",
                'IF({closed}{r}="Yes","",'
                'COUNTIFS({closed_first}:{closed_last},"No",'
                '{room_first}:{room_last},{room}{r},'
                '{priority_rank_first}:{priority_rank_last},'
                '"<"&{priority_rank}{r})+'
                'COUNTIFS({closed_first}:{closed_last},"No",'
                '{room_first}:{room_last},{room}{r},'
                '{priority_rank_first}:{priority_rank_last},'
                '{priority_rank}{r},'
                '{enquiry_date_first}:{enquiry_date_last},'
                '"<="&{enquiry_date}{r}))'),
            note="Where they stand in the queue for the room they "
                 "want: priority first, then how long they have "
                 "waited. Nothing to type."),
        Col("seats_now", "Seats free now", 11, "int",
            formula=blank(
                "room",
                'IFERROR(INDEX([[Classroom Management.seats]],'
                'MATCH({room}{r},[[Classroom Management.name]],0))*1,'
                '0)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}],
            note="Live, from the classroom tab. A free seat and "
                 "somebody waiting for it is money sitting on the "
                 "table."),
        Col("value_week", "Worth a week", 12, "money",
            formula=blank(
                "days_wanted",
                'ROUND(IF({days_wanted}{r}>=5,FullTimeWeek,'
                '{days_wanted}{r}*PartTimeDay),2)'),
            note="Rough value at your standard rates, so you can see "
                 "what the waiting list is worth."),
        Col("followed_up", "Last follow-up", 12, "date",
            validate=date_check(), hidden=True),
        Col("notes", "Notes", 30, "wrap"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "child_name",
                'IF({closed}{r}="Yes","",'
                'IF({room}{r}="","Add - which room do they want?",'
                'IF(IF({seats_now}{r}<=0,FALSE,'
                '({status}{r}="Waiting - no space")+'
                '({status}{r}="Toured")+'
                '({status}{r}="New enquiry")+'
                '({status}{r}="Information sent")>0),'
                '"Space - a seat just opened, offer it",'
                'IF({days_quiet}{r}>=14,"Chase - two weeks with no '
                'contact",'
                'IF(IF({offered_date}{r}="",FALSE,'
                'IF({response_date}{r}<>"",FALSE,'
                'TodayDate-{offered_date}{r}>7)),'
                '"Chase - offer made, no reply",'
                'IF({status}{r}="Toured","Follow up - toured, now '
                'offer them a place",'
                'IF({tour_date}{r}="","Confirm - book them a tour",'
                'IF({days_quiet}{r}>=7,"Follow up - going quiet",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Waitlist & Enquiries",
        subtitle="Every enquiry, where it got to, how long they have "
                 "waited and which of them you can say yes to today.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=3,
        samples=sample(cfg).waitlist,
        help_text="Log every enquiry the day it arrives. The queue "
                  "position, the waiting time, the free seats and the "
                  "follow-up list all calculate themselves.",
        extra_notes=[
            "The queue position only counts enquiries that are still "
            "open and still want the same room, so it is a real "
            "position rather than a row number.",
            "'Seats free now' reads the classroom tab live. When it "
            "goes above zero the Action column tells you to offer the "
            "place - that is the one alert on this tab that makes "
            "money.",
            "Fill in 'Heard about you via' every single time. After "
            "thirty enquiries it tells you which half of your "
            "marketing budget to stop spending.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Pipeline summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The pipeline", last)
    sheet.table_head(row, 0, [
        "Enquiries logged", "Still open", "Tours booked", "Toured",
        "Offers out", "Enrolled", "Lost or declined",
        "Enquiry to enrollment", "Tour to enrollment",
        "Average days waiting", "Waiting list worth a week"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Waitlist.child_name]],"<>")',
        '=COUNTIFS([[Waitlist.closed]],"No")',
        '=COUNTIFS([[Waitlist.tour_date]],"<>")',
        '=COUNTIFS([[Waitlist.status]],"Toured")',
        '=COUNTIFS([[Waitlist.offered_date]],"<>")',
        '=COUNTIFS([[Waitlist.converted]],"Yes")',
        '=COUNTIFS([[Waitlist.status]],"Offer declined")+'
        'COUNTIFS([[Waitlist.status]],"Lost")+'
        'COUNTIFS([[Waitlist.status]],"Withdrawn")',
        '=IF(COUNTIF([[Waitlist.child_name]],"<>")<=0,"",'
        'COUNTIFS([[Waitlist.converted]],"Yes")/'
        'COUNTIF([[Waitlist.child_name]],"<>"))',
        '=IF(COUNTIFS([[Waitlist.tour_date]],"<>")<=0,"",'
        'COUNTIFS([[Waitlist.tour_date]],"<>",'
        '[[Waitlist.converted]],"Yes")/'
        'COUNTIFS([[Waitlist.tour_date]],"<>"))',
        '=IFERROR(ROUND(AVERAGEIF([[Waitlist.days_waiting]],">=0"),0),'
        '0)',
        '=ROUND(SUMIFS([[Waitlist.value_week]],'
        '[[Waitlist.closed]],"No"),2)',
    ], ["int", "int", "int", "int", "int", "int", "int", "percent",
        "percent", "int", "money"])
