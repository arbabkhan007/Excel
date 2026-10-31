"""`Child Enrollment` - the register every other tab hangs off.

One row per child: who they are, how old they are today, which room
they are in, which family pays, what that family is contracted to pay,
how often the child actually turns up, and what is still outstanding on
their paperwork and their account.

Age is calculated, never typed.  So is the age band, the weekly fee,
the attendance percentage, the balance and the list of missing
documents.  The only things you type are the things only you know.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (age_months, age_years_text, alert_cf, bar_cf, blank,
                     count_absent, count_present, id_col, lookup_family,
                     next_birthday, owing_cf, rank_desc, status_cf,
                     yesno)

NAME = "Child Enrollment"


def spec(cfg: Config) -> SheetSpec:
    med_child = "[[Allergy & Medical.child]]"
    med_sev = "[[Allergy & Medical.severity]]"
    doc_who = "[[Document Tracker.who]]"
    doc_req = "[[Document Tracker.required]]"
    doc_status = "[[Document Tracker.status]]"

    def severity_count(level: str) -> str:
        return f'COUNTIFS({med_child},{{id}}{{r}},{med_sev},"{level}")'

    columns = [
        id_col("name", "ChildPrefix", digits=4, width=11,
               header="Child ID",
               note="Automatic. Attendance, fees, meals, daily reports "
                    "and incidents all point at this ID."),
        Col("name", "Child's full name", 24, "text",
            note="As it appears on the birth certificate and the "
                 "enrollment agreement. One row per child, always."),
        Col("preferred", "Known as", 14, "text",
            note="What you actually call them. Printed on the daily "
                 "report."),
        Col("dob", "Date of birth", 12, "date", validate=date_check(),
            note="Everything age related in this workbook is worked "
                 "out from this one cell - age, age band, room "
                 "suitability, birthdays, moving-up dates."),
        Col("age_months", "Age (months)", 11, "int",
            formula="=" + age_months("{dob}{r}"), hidden=True),
        Col("age", "Age", 10, "text",
            formula="=" + age_years_text("{dob}{r}"),
            note="Calculated. Recalculates every time you open the "
                 "file."),
        Col("age_band", "Age band", 14, "text",
            formula=blank(
                "dob",
                'IFERROR(INDEX(AgeBandList,MATCH({age_months}{r},'
                'AgeBandFrom,1))&"","")'),
            note="From the age band table on Settings. Tells you when "
                 "a child has grown out of the room they are in."),
        Col("gender", "Gender", 11, "text",
            validate=list_from("GenderList"),
            note="Optional. Some settings record it, some deliberately "
                 "do not - leave it blank if you do not."),
        Col("family", "Family ID", 11, "id",
            validate=list_from("FamilyList"),
            note="The family that pays. Two siblings share one family "
                 "ID - that is what makes one invoice per family and "
                 "the sibling discount possible."),
        Col("family_name", "Family", 18, "text",
            formula=lookup_family("family_name")),
        Col("parent_name", "Main contact", 20, "text",
            formula=lookup_family("parent1")),
        Col("parent_phone", "Phone", 15, "text",
            formula=lookup_family("phone1")),
        Col("parent_email", "Email", 24, "text",
            formula=lookup_family("email1"), hidden=True),
        Col("family_seq", "Sibling no.", 9, "int",
            formula=blank(
                "family",
                'COUNTIFS({family_first}:{family}{r},{family}{r})'),
            hidden=True,
            note="1 for the first child of a family, 2 for the next. "
                 "The sibling discount reads this."),
        Col("family_key", "Family key", 12, "text",
            formula=blank("family",
                          '{family}{r}&"|"&{family_seq}{r}'),
            hidden=True),
        # -------------------------------------------------- the place ---
        Col("program", "Programme", 14, "text",
            validate=list_from("ProgramList"), cf=status_cf()),
        Col("room", "Classroom", 16, "text",
            validate=list_from("ClassroomList"),
            note="Pick from the rooms on Classroom Management. The "
                 "capacity and ratio figures count this column."),
        Col("teacher", "Key teacher", 18, "text",
            formula=blank(
                "room",
                'IFERROR(INDEX([[Classroom Management.teacher_name]],'
                'MATCH({room}{r},[[Classroom Management.name]],0))&"",'
                '"")'),
            note="The lead teacher of whichever room the child is in."),
        Col("schedule", "Schedule", 16, "text",
            validate=list_from("ScheduleTypeList")),
        Col("days_week", "Days a week", 11, "number1",
            validate=number_check(0, 7),
            note="Contracted sessions a week. Daily and per-session "
                 "fees multiply by this."),
        Col("hours_day", "Hours a day", 11, "number1",
            validate=number_check(0, 24),
            note="Contracted hours. Hourly fees multiply by this and "
                 "by the days a week."),
        Col("usual_in", "Usual drop-off", 12, "time",
            validate=time_check(),
            note="The contracted arrival time. Attendance measures "
                 "late arrivals against it."),
        Col("usual_out", "Usual pickup", 12, "time",
            validate=time_check(),
            note="The contracted collection time. The late pickup fee "
                 "is calculated from this, not from closing time."),
        Col("enrolled_date", "Enrolled on", 12, "date",
            validate=date_check(),
            note="The day the paperwork was signed."),
        Col("start_date", "Starts", 11, "date", validate=date_check(),
            note="First session. A start date in the future with "
                 "status 'Pending start' is a place held, not a child "
                 "in the room."),
        Col("end_date", "Last day", 11, "date", validate=date_check(),
            note="Leave blank while the child is still with you."),
        Col("leaving_reason", "Leaving because", 18, "text",
            validate=list_from("LeavingReasonList"), hidden=True,
            note="Worth filling in. Six of these in a row tells you "
                 "something an individual conversation never will."),
        Col("status", "Status", 13, "status",
            validate=list_from("EnrollmentStatusList"), cf=status_cf(),
            note="Active, On hold and Notice given all occupy a place. "
                 "Withdrawn and Graduated free it up."),
        Col("holds_place", "Holds a place?", 11, "flag",
            formula=blank(
                "name",
                'IF(({status}{r}="Active")+({status}{r}="On hold")+'
                '({status}{r}="Notice given")>0,"Yes","No")'),
            hidden=True),
        Col("months_with_us", "Months with us", 12, "int",
            formula=blank(
                "start_date",
                'DATEDIF(MIN({start_date}{r},TodayDate),'
                'IF({end_date}{r}="",TodayDate,{end_date}{r}),"m")'),
            hidden=True),
        # ------------------------------------------------- the money ----
        Col("basis", "Billing basis", 13, "text",
            validate=list_from("BillingBasisList"),
            note="How this family's tuition is worked out. Different "
                 "children can be on different bases - a full-timer on "
                 "a monthly rate next to a drop-in on an hourly one."),
        Col("rate", "Agreed rate", 12, "money",
            validate=number_check(0, 100000),
            note="The rate for the basis you picked: the weekly rate "
                 "if weekly, the hourly rate if hourly, and so on."),
        Col("auto_discount", "Sibling %", 9, "percent",
            formula=blank(
                "name",
                'IF({family_seq}{r}>1,SiblingDiscount,0)'),
            hidden=True),
        Col("discount_pct", "Discount", 10, "percent",
            validate=number_check(0, 1),
            note="Leave blank and the sibling discount applies "
                 "automatically to the second and later child of a "
                 "family. Type a figure here to override it."),
        Col("discount_used", "Discount applied", 12, "percent",
            formula=blank(
                "name",
                'IF({discount_pct}{r}<>"",{discount_pct}{r},'
                '{auto_discount}{r})'),
            note="What is actually coming off. Typed beats automatic."),
        Col("subsidy_week", "Subsidy a week", 12, "money",
            validate=number_check(0, 100000),
            note="Government or employer help paid on this child's "
                 "behalf. It comes off what the family owes, and it is "
                 "still revenue to you."),
        Col("gross_week", "Full rate a week", 13, "money",
            formula=blank(
                "name",
                'ROUND(IF({basis}{r}="Weekly rate",{rate}{r},'
                'IF({basis}{r}="Fortnightly rate",{rate}{r}/2,'
                'IF({basis}{r}="Monthly rate",{rate}{r}*12/52,'
                'IF({basis}{r}="Term rate",{rate}{r}/MAX(1,TermWeeks),'
                'IF({basis}{r}="Daily rate",{rate}{r}*{days_week}{r},'
                'IF({basis}{r}="Per session",{rate}{r}*{days_week}{r},'
                'IF({basis}{r}="Hourly rate",'
                '{rate}{r}*{hours_day}{r}*{days_week}{r},0))))))),2)'),
            note="Every basis converted to one comparable weekly "
                 "number, before any discount."),
        Col("weekly_net", "Weekly fee", 12, "money",
            formula=blank(
                "name",
                'ROUND(MAX(0,{gross_week}{r}*(1-{discount_used}{r})-'
                '{subsidy_week}{r}),2)'),
            cf=bar_cf(Palette.accent),
            note="After the discount and after the subsidy - what the "
                 "family actually pays in a normal week."),
        Col("monthly_net", "Monthly fee", 12, "money",
            formula=blank(
                "name",
                'ROUND({weekly_net}{r}*52/12,2)'),
            note="The weekly fee annualised and divided by twelve, so "
                 "that every month is the same amount."),
        Col("annual_net", "A year", 12, "money",
            formula=blank("name",
                          'ROUND({weekly_net}{r}*WeeksOpen,2)'),
            hidden=True,
            note="Weekly fee x the weeks you open. What this child is "
                 "worth over a year."),
        Col("billed", "Billed to date", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Tuition & Fees.net]],'
                '[[Tuition & Fees.child]],{id}{r}),2)'),
            note="Every charge raised against this child on the "
                 "Tuition & Fees tab, discounts and credits included."),
        Col("balance", "Family balance", 12, "money",
            formula=blank(
                "family",
                'IFERROR(INDEX([[Parent Database.balance]],'
                'MATCH({family}{r},[[Parent Database.id]],0))*1,0)'),
            cf=owing_cf(),
            note="The whole family's balance, not just this child's - "
                 "because you invoice the family, not the child."),
        # ---------------------------------------------- how often in ----
        Col("sessions", "Sessions booked", 13, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Attendance.child]],{id}{r})-'
                '(' + "+".join(
                    f'COUNTIFS([[Attendance.child]],{{id}}{{r}},'
                    f'[[Attendance.status]],"{s}")'
                    for s in ("Public holiday", "Centre closed",
                              "Snow day", "Not scheduled")) + ')'),
            hidden=True,
            note="Days a session was expected. Closures and days the "
                 "child was never booked for do not count."),
        Col("days_in", "Days attended", 12, "int",
            formula=blank("name", count_present("{id}{r}")),
            cf=bar_cf()),
        Col("days_out", "Days absent", 11, "int",
            formula=blank("name", count_absent("{id}{r}")),
            cf=[{"type": "cell", "criteria": ">=", "value": 10,
                 "fmt": "warn"}]),
        Col("days_sick", "Sick days", 10, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Attendance.child]],{id}{r},'
                '[[Attendance.status]],"Absent - sick")'),
            hidden=True),
        Col("attendance_pct", "Attendance", 11, "percent",
            formula=blank(
                "name",
                'IF({sessions}{r}<=0,"",{days_in}{r}/{sessions}{r})'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetAttendance", "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.7,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetAttendance", "fmt": "warn"}],
            note="Days attended divided by sessions expected. "
                 "Closures and unbooked days are excluded, so this is "
                 "a fair figure."),
        Col("last_in", "Last attended", 12, "date",
            formula=blank(
                "name",
                'IF({days_in}{r}=0,"",' +
                max_if("[[Attendance.date]]",
                       "[[Attendance.child]]={id}{r}",
                       '[[Attendance.present]]=1') + ")")),
        Col("days_since", "Days since", 10, "int",
            formula=blank(
                "name",
                'IF({last_in}{r}="","",TodayDate-{last_in}{r})'),
            hidden=True),
        Col("late_count", "Late arrivals", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Attendance.child]],{id}{r},'
                '[[Attendance.status]],"Late")'),
            hidden=True),
        Col("late_pickups", "Late pickups", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Attendance.child]],{id}{r},'
                '[[Attendance.late_out]],">0")'),
            cf=[{"type": "cell", "criteria": ">=", "value": 3,
                 "fmt": "warn"}],
            note="Collections after the contracted pickup time. Three "
                 "of them is a conversation, not a fee."),
        # -------------------------------------------------- care needs --
        Col("diet", "Diet", 14, "text",
            validate=list_from("DietList"),
            note="Feeds the meal tracker, where it sits next to every "
                 "plate you serve."),
        Col("allergy_level", "Allergy risk", 13, "status",
            formula=blank(
                "name",
                f'IF({severity_count("Anaphylaxis risk")}>0,'
                f'"Anaphylaxis risk",'
                f'IF({severity_count("Severe")}>0,"Severe",'
                f'IF({severity_count("Moderate")}>0,"Moderate",'
                f'IF({severity_count("Mild")}>0,"Mild",'
                f'"None recorded"))))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Anaphylaxis", "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Severe", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Moderate", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "None", "fmt": "muted"}],
            note="The worst severity on the Allergy & Medical tab for "
                 "this child. Nothing to type - record the allergy "
                 "there and it appears here, on the meal tracker and "
                 "on the daily report."),
        Col("allergy_count", "Allergies", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Allergy & Medical.child]],{id}{r},'
                '[[Allergy & Medical.record_type]],"Allergy")'),
            hidden=True),
        Col("first_allergy", "Main allergy", 18, "text",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Allergy & Medical.detail]],'
                'MATCH({id}{r}&"|1",[[Allergy & Medical.key]],0))&"",'
                '"")'),
            hidden=True),
        Col("medical_notes", "Care notes", 28, "wrap",
            note="Asthma inhaler in the red box, needs a nap at 12, "
                 "frightened of the hand dryer. The things a cover "
                 "member of staff has to know."),
        Col("pickup_list", "Authorised to collect", 30, "wrap",
            note="Names of everybody allowed to collect this child. "
                 "Attendance checks what you type at pickup against "
                 "this list and flags anything it cannot find."),
        Col("pickup_password", "Pickup password", 14, "text",
            note="A word only the family knows, for the day somebody "
                 "unexpected arrives."),
        Col("photo_ok", "Photos OK?", 10, "flag", validate=yesno(),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"}],
            note="Check this before anything goes on social media."),
        Col("sunscreen_ok", "Sunscreen OK?", 11, "flag",
            validate=yesno(), hidden=True),
        Col("outings_ok", "Outings OK?", 10, "flag", validate=yesno(),
            hidden=True),
        # ------------------------------------------------- paperwork ----
        Col("docs_required", "Docs required", 11, "int",
            formula=blank(
                "name",
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes")'),
            hidden=True),
        Col("docs_ok", "Docs on file", 11, "int",
            formula=blank(
                "name",
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"On file")+'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"Signed")+'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"Received")'),
            hidden=True),
        Col("docs_missing", "Docs missing", 11, "int",
            formula=blank(
                "name",
                '{docs_required}{r}-{docs_ok}{r}'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "=", "value": 0,
                 "fmt": "good"}],
            note="Required documents for this child that are not yet "
                 "received, signed or on file. Set them up on the "
                 "Document Tracker."),
        Col("contract", "Agreement", 12, "status",
            validate=list_from("DocumentStatusList"), cf=status_cf(),
            note="The enrollment agreement and consent forms. "
                 "Shorthand for the full picture on the Document "
                 "Tracker."),
        # -------------------------------------------------- birthdays ---
        Col("next_bday", "Next birthday", 12, "date",
            formula="=" + next_birthday("{dob}{r}")),
        Col("bday_in", "Days to birthday", 12, "int",
            formula=blank("dob", '{next_bday}{r}-TodayDate'),
            hidden=True),
        Col("turning", "Turning", 8, "int",
            formula=blank(
                "dob", 'YEAR({next_bday}{r})-YEAR({dob}{r})'),
            hidden=True),
        Col("anniversary", "Years with us", 11, "number1",
            formula=blank(
                "start_date",
                'ROUND((IF({end_date}{r}="",TodayDate,{end_date}{r})-'
                '{start_date}{r})/365,1)'),
            hidden=True),
        Col("moves_up", "Moves up on", 12, "date",
            formula=blank(
                "dob",
                'IF({age_band}{r}="","",'
                'IFERROR(EDATE({dob}{r},INDEX(AgeBandTo,'
                'MATCH({age_months}{r},AgeBandFrom,1))),""))'),
            note="The day the child ages out of their band. Plan the "
                 "room move before the parent asks about it."),
        Col("notes", "Notes", 28, "wrap"),
        rank_desc("annual_net", header="Value rank"),
        Col("alert", "Action", 34, "status",
            formula=blank(
                "name",
                'IF({dob}{r}="","Add - no date of birth",'
                'IF({family}{r}="","Add - no family linked",'
                'IF(({status}{r}="Withdrawn")+({status}{r}="Graduated")'
                '>0,"",'
                'IF({docs_missing}{r}>0,"Missing - "&{docs_missing}{r}&'
                '" document(s) outstanding",'
                'IF({balance}{r}>0.004,"Chase - family owes money",'
                'IF({room}{r}="","Add - no classroom",'
                'IF({basis}{r}="","Add - no billing basis",'
                'IF(IF({weekly_net}{r}=0,{basis}{r}<>"No charge",FALSE),'
                '"Check - fee works out at zero",'
                'IF(IF({moves_up}{r}="",FALSE,'
                '{moves_up}{r}<=TodayDate+30),'
                '"Due - ages out of the band soon",'
                'IF(IF({bday_in}{r}="",FALSE,'
                '{bday_in}{r}<=BirthdayWindow),'
                '"Birthday - turning "&{turning}{r}&" in "&'
                '{bday_in}{r}&" days",'
                'IF(IF({attendance_pct}{r}="",FALSE,'
                '{attendance_pct}{r}<0.7),'
                '"Check - attendance below 70%","OK")))))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Child Enrollment & Registration",
        subtitle="Every child, their family, their room, their fee and "
                 "the paperwork still outstanding.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).children,
        help_text="Type the name, the date of birth and the family ID. "
                  "Age, age band, the weekly fee, attendance, the "
                  "balance and the missing paperwork all calculate "
                  "themselves.",
        extra_notes=[
            "The sibling discount is automatic: the first child of a "
            "family pays full price, the second and later ones get "
            "the percentage set on Settings. Type a figure in the "
            "Discount column to override it for one child.",
            "'Holds a place' is what the capacity figures count. A "
            "child on notice still occupies a seat until their last "
            "day - which is exactly when you need to know about it.",
            "Filter the Action column every Monday morning. Missing "
            "documents, outstanding balances and children about to "
            "age out of their room are the three things that cost "
            "money if you find them late.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Enrollment summary strip."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Enrollment at a glance", last)
    sheet.table_head(row, 0, [
        "On the register", "Active", "Pending start", "On notice",
        "Withdrawn", "Average age (months)", "Weekly contracted fees",
        "Average weekly fee", "Average attendance",
        "Children with documents missing", "Allergy plans on file"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Child Enrollment.name]],"<>")',
        '=COUNTIFS([[Child Enrollment.status]],"Active")',
        '=COUNTIFS([[Child Enrollment.status]],"Pending start")',
        '=COUNTIFS([[Child Enrollment.status]],"Notice given")',
        '=COUNTIFS([[Child Enrollment.status]],"Withdrawn")',
        '=IFERROR(ROUND(AVERAGEIF([[Child Enrollment.age_months]],'
        '">0"),0),0)',
        '=ROUND(SUMIFS([[Child Enrollment.weekly_net]],'
        '[[Child Enrollment.holds_place]],"Yes"),2)',
        '=IFERROR(ROUND(SUMIFS([[Child Enrollment.weekly_net]],'
        '[[Child Enrollment.holds_place]],"Yes")/'
        'MAX(1,COUNTIFS([[Child Enrollment.holds_place]],"Yes")),2),0)',
        '=IFERROR(AVERAGEIF([[Child Enrollment.attendance_pct]],">0"),'
        '"")',
        '=COUNTIFS([[Child Enrollment.docs_missing]],">0")',
        '=COUNTIFS([[Child Enrollment.allergy_count]],">0")',
    ], ["int", "int", "int", "int", "int", "int", "money", "money",
        "percent", "int", "int"])
