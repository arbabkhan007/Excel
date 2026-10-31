"""`Reviews Tracker` - the seven scores, and which one is costing you.

An overall rating is nearly useless on its own.  A 4.6 made of straight
4.6s is a property that needs re-photographing; a 4.6 made of six 4.9s
and a 3.4 for cleanliness is one cleaner having a bad month.  This tab
keeps all seven category scores side by side so the weakest one is
obvious, per property and per platform.

It also tracks the thing hosts forget: whether you asked for the review
at all, and whether you replied to it.
"""

from __future__ import annotations

from ..config import RATING_CATEGORIES, Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_booking,
                     month_col, rating_cf, score_validation, status_cf,
                     year_col, yesno)

NAME = "Reviews Tracker"


def spec(cfg: Config) -> SheetSpec:
    score_keys = [k for k, _h in RATING_CATEGORIES if k != "rate_overall"]
    detail_sum = "+".join(f'IFERROR({{{k}}}{{r}}*1,0)' for k in score_keys)
    detail_count = "+".join(f'({{{k}}}{{r}}<>"")' for k in score_keys)

    columns = [
        id_col("booking", "ReviewPrefix", digits=cfg.digits("review"),
               width=10, header="Review ID"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="The stay being reviewed. One row per review.",
            group="The stay"),
        Col("property", "Property", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="The stay"),
        Col("property_name", "Property", 19, "text",
            formula=lookup_booking("property_name"), group="The stay"),
        Col("guest", "Guest", 11, "text",
            formula=lookup_booking("guest"), hidden=True,
            group="The stay"),
        Col("guest_name", "Guest", 18, "text",
            formula=lookup_booking("guest_name"), group="The stay"),
        Col("platform", "Platform", 14, "text",
            formula=lookup_booking("platform"), group="The stay"),
        Col("checkout", "Checked out", 12, "date",
            formula=lookup_booking("check_out"), group="The stay"),

        # ------------------------------------------------- the process --
        Col("requested", "Review asked for", 13, "date",
            validate=date_check(),
            note="The date you asked. Most platforms prompt the guest "
                 "automatically, but a short personal message still "
                 "moves the needle.",
            group="The process"),
        Col("status", "Status", 15, "status",
            validate=list_from("ReviewStatusList"),
            cf=status_cf(), group="The process"),
        Col("review_date", "Review received", 13, "date",
            validate=date_check(), group="The process"),
        month_col("review_date"),
        year_col("review_date"),
        Col("days_to_review", "Days to review", 13, "int",
            formula=blank(
                "booking",
                'IF({review_date}{r}="","",IF({checkout}{r}="","",'
                'MAX(0,{review_date}{r}-{checkout}{r})))'),
            hidden=True, group="The process"),
    ]

    for key, header in RATING_CATEGORIES:
        columns.append(
            Col(key, header, 11, "number1",
                validate=score_validation(5),
                cf=rating_cf() if key == "rate_overall" else [],
                note=f"{header} out of 5. Leave blank if the platform "
                     f"does not score it.",
                group="The scores"))

    columns += [
        Col("has_score", "Scored", 8, "flag",
            formula=blank("booking",
                          'IF({rate_overall}{r}="","No","Yes")'),
            cf=flag_yes_no_cf(), hidden=True, group="The scores"),
        Col("detail_avg", "Category average", 14, "number",
            formula=blank(
                "booking",
                f'IF(({detail_count})<=0,"",'
                f'ROUND(({detail_sum})/({detail_count}),2))'),
            cf=rating_cf(),
            note="The six category scores averaged. Compare it with the "
                 "overall - a gap means the guest liked the place more "
                 "than the sum of its parts, or a lot less.",
            group="The scores"),
        Col("weakest", "Weakest category", 16, "text",
            formula=blank(
                "booking",
                'IF({rate_clean}{r}="","",'
                'IF({rate_clean}{r}<=MIN(IFERROR({rate_accuracy}{r}*1,5),'
                'IFERROR({rate_checkin}{r}*1,5),'
                'IFERROR({rate_comm}{r}*1,5),'
                'IFERROR({rate_location}{r}*1,5),'
                'IFERROR({rate_value}{r}*1,5)),"Cleanliness",'
                'IF(IFERROR({rate_value}{r}*1,5)<=MIN('
                'IFERROR({rate_accuracy}{r}*1,5),'
                'IFERROR({rate_checkin}{r}*1,5),'
                'IFERROR({rate_comm}{r}*1,5),'
                'IFERROR({rate_location}{r}*1,5)),"Value",'
                'IF(IFERROR({rate_accuracy}{r}*1,5)<=MIN('
                'IFERROR({rate_checkin}{r}*1,5),'
                'IFERROR({rate_comm}{r}*1,5),'
                'IFERROR({rate_location}{r}*1,5)),"Accuracy",'
                'IF(IFERROR({rate_checkin}{r}*1,5)<=MIN('
                'IFERROR({rate_comm}{r}*1,5),'
                'IFERROR({rate_location}{r}*1,5)),"Check-in",'
                'IF(IFERROR({rate_comm}{r}*1,5)<='
                'IFERROR({rate_location}{r}*1,5),"Communication",'
                '"Location"))))))'),
            note="Which of the six scored categories came out lowest on "
                 "this review. The column to filter when a rating "
                 "starts sliding.",
            group="The scores"),
        Col("below_target", "Under target", 11, "flag",
            formula=blank(
                "booking",
                'IF({rate_overall}{r}="","No",'
                'IF({rate_overall}{r}<TargetRating,"Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"}],
            hidden=True, group="The scores"),

        # ------------------------------------------------- what was said --
        Col("public_review", "What they wrote", 36, "wrap",
            group="What was said"),
        Col("private_note", "Private feedback", 30, "wrap",
            note="The bit only you see. Usually more useful than the "
                 "public review.",
            group="What was said"),
        Col("responded", "You replied", 11, "flag",
            validate=yesno("Have you replied publicly?"),
            cf=flag_yes_no_cf(),
            note="Always reply to a poor review. Future guests read "
                 "the reply, not the complaint.",
            group="What was said"),
        Col("response_date", "Replied on", 12, "date",
            validate=date_check(), hidden=True,
            group="What was said"),
        Col("response", "Your reply", 32, "wrap",
            group="What was said"),
        Col("action_taken", "What you changed", 28, "wrap",
            note="The point of a bad review. New mattress, earlier "
                 "cleaner, clearer directions.",
            group="What was said"),

        Col("alert", "Action", 28, "status",
            formula=blank(
                "booking",
                'IF({checkout}{r}="","Check - booking not found",'
                'IF({has_score}{r}="No",'
                'IF({status}{r}="No review left","",'
                'IF(TodayDate-{checkout}{r}>ReviewChaseDays,'
                '"Chase - no review yet","Due - review window open")),'
                'IF({below_target}{r}="Yes",'
                'IF({responded}{r}<>"Yes","Action - reply to this '
                'review","OK - replied"),'
                'IF({responded}{r}<>"Yes","Follow up - thank them",'
                '"OK"))))'),
            cf=alert_cf(), group="What was said"),
    ]

    return SheetSpec(
        name=NAME,
        title="Reviews Tracker",
        subtitle="All seven scores, per property and per platform - so "
                 "you can see which one is actually costing you",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("reviews", cfg),
        help_text="One row per review. Pick the booking and the "
                  "property, guest and platform come with it. Fill in "
                  "whichever of the seven scores the platform gives "
                  "you - the weakest category is worked out for you, "
                  "and that is the column worth sorting on when a "
                  "rating starts to slide.",
        extra_notes=[
            "Always reply, and always reply to the bad ones. Future "
            "guests read the reply far more carefully than the "
            "complaint.",
            "A review with no score is still worth a row. The 'No "
            "review left' status stops the chase without pretending "
            "they gave you five stars.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    scored = 'COUNTIFS([[Reviews Tracker.has_score]],"Yes")'

    row = sheet.section(row, "How you are doing", last)
    row = sheet.table_head(row, 0, [
        "Reviews logged", "Scored", "Review rate", "Average overall",
        "Under target", "5 stars", "Replied to", "Avg days to review"],
        [13, 9, 12, 14, 12, 9, 11, 16])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Reviews Tracker.booking]],"?*")',
        f"={scored}",
        '=IF(COUNTIF([[Reviews Tracker.booking]],"?*")<=0,"",'
        f'{scored}/COUNTIF([[Reviews Tracker.booking]],"?*"))',
        f'=IF({scored}<=0,"",ROUND(SUMIFS('
        '[[Reviews Tracker.rate_overall]],'
        f'[[Reviews Tracker.has_score]],"Yes")/{scored},2))',
        '=COUNTIFS([[Reviews Tracker.below_target]],"Yes")',
        '=COUNTIFS([[Reviews Tracker.rate_overall]],5)',
        '=COUNTIFS([[Reviews Tracker.responded]],"Yes")',
        f'=IF({scored}<=0,"",ROUND(SUM('
        f'[[Reviews Tracker.days_to_review]])/{scored},1))',
    ], ["int", "int", "percent", "number", "int", "int", "int", "number1"])
    row += 2

    row = sheet.section(row, "Where the stars are going", last)
    headers = ["Category", "Reviews scored", "Average", "Under 4",
               "5 stars", "Weakest on"]
    row = sheet.table_head(row, 0, headers, [20, 14, 11, 10, 9, 12])
    for key, label in RATING_CATEGORIES:
        col = f"[[Reviews Tracker.{key}]]"
        count = f'COUNT({col})'
        sheet.calc_row(row, 0, [
            label,
            f"={count}",
            f'=IF({count}<=0,"",ROUND(SUM({col})/{count},2))',
            f'=COUNTIFS({col},"<4",{col},">0")',
            f'=COUNTIFS({col},5)',
            f'=COUNTIFS([[Reviews Tracker.weakest]],"{label}")',
        ], ["text", "int", "number", "int", "int", "int"])
        row += 1

    row += 1
    sheet.note(row, "'Weakest on' counts how many reviews put this "
                    "category bottom. One category holding the bottom "
                    "spot again and again is the thing to fix - and it "
                    "is almost always cleanliness or value, both of "
                    "which you control.", last)
