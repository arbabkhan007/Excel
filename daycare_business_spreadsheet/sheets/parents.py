"""`Parent Database` - one row per family, however many children they have.

The mistake most childcare templates make is one row per parent, which
means a family with two children appears twice, gets two invoices and
shows two half-balances.  Here a family is a family: one row, one
balance, one invoice, as many children as they like linked to it.

Emergency contacts, authorised collectors and the preferred way of
being contacted all live here too, because those are family facts
rather than child facts.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if, min_if
from ..sample_data import get as sample
from ..writer import list_from, date_check
from .common import (alert_cf, bar_cf, blank, id_col, owing_cf,
                     rank_desc, status_cf)

NAME = "Parent Database"


def spec(cfg: Config) -> SheetSpec:
    child_family = "[[Child Enrollment.family]]"
    child_status = "[[Child Enrollment.status]]"
    led_family = "[[Invoices & Receipts.family]]"

    # the earliest due date among this family's overdue invoices, and
    # the sentinel min_if() hands back when there are none
    _oldest_due = min_if("[[Invoices & Receipts.due_num]]",
                         f"{led_family}={{id}}{{r}}",
                         '[[Invoices & Receipts.status]]="Overdue"')
    led_type = "[[Invoices & Receipts.type]]"

    def nth_child(n: int) -> str:
        return (f'IFERROR(INDEX([[Child Enrollment.name]],'
                f'MATCH({{id}}{{r}}&"|{n}",'
                f'[[Child Enrollment.family_key]],0))&"","")')

    columns = [
        id_col("parent1", "FamilyPrefix", digits=4, width=11,
               header="Family ID",
               note="Automatic. Children, invoices, payments and "
                    "messages all point at this ID. One family, one "
                    "ID, however many children."),
        Col("family_name", "Family name", 16, "text",
            note="The surname you would write on an envelope. Used on "
                 "invoices and in every list that groups by family."),
        # ------------------------------------------------ first parent --
        Col("parent1", "Parent / guardian 1", 22, "text",
            note="The main contact. Filling this in is what creates "
                 "the family ID."),
        Col("relation1", "Relationship", 13, "text",
            validate=list_from("RelationshipList")),
        Col("phone1", "Phone", 15, "text"),
        Col("email1", "Email", 26, "text"),
        Col("employer1", "Employer", 18, "text", hidden=True,
            note="Optional. Useful when an employer pays part of the "
                 "fees or runs a childcare voucher scheme."),
        Col("work_phone1", "Work phone", 14, "text", hidden=True),
        # ----------------------------------------------- second parent --
        Col("parent2", "Parent / guardian 2", 22, "text"),
        Col("relation2", "Relationship", 13, "text",
            validate=list_from("RelationshipList")),
        Col("phone2", "Phone", 15, "text"),
        Col("email2", "Email", 26, "text", hidden=True),
        # --------------------------------------------------- where ------
        Col("address", "Address", 30, "text"),
        Col("city", "Town / city", 16, "text"),
        Col("postcode", "Postcode / ZIP", 12, "text"),
        Col("contact_by", "Contact by", 15, "text",
            validate=list_from("ContactMethodList"),
            note="How this family actually wants to hear from you. "
                 "Respect it and you will get replies."),
        Col("language", "Main language", 14, "text", hidden=True,
            note="Worth recording if you ever send a notice home."),
        # --------------------------------------------- emergency ------
        Col("em1_name", "Emergency contact 1", 20, "text",
            note="Somebody who is not parent 1 or 2, and who can get "
                 "here quickly."),
        Col("em1_phone", "Phone", 15, "text"),
        Col("em1_relation", "Relationship", 13, "text",
            validate=list_from("RelationshipList"), hidden=True),
        Col("em2_name", "Emergency contact 2", 20, "text",
            hidden=True),
        Col("em2_phone", "Phone", 15, "text", hidden=True),
        Col("doctor", "Family doctor", 20, "text", hidden=True),
        Col("doctor_phone", "Doctor's phone", 15, "text", hidden=True),
        Col("pickup_list", "Authorised to collect", 32, "wrap",
            note="Everybody this family has authorised, separated by "
                 "commas. Individual children can have their own list "
                 "on the enrollment tab - the attendance check looks "
                 "at both."),
        Col("pickup_ok", "Pickup list signed?", 13, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bad"}],
            note="Has the family signed the pickup authorisation form? "
                 "Until they have, only the parents collect."),
        Col("custody_notes", "Custody / legal notes", 28, "wrap",
            hidden=True,
            note="Court orders and collection restrictions. Sensitive - "
                 "restrict who can see this file."),
        # ---------------------------------------------- their children --
        Col("children", "Children", 9, "int",
            formula=blank(
                "parent1",
                f'COUNTIFS({child_family},{{id}}{{r}})'),
            cf=bar_cf()),
        Col("children_active", "In now", 8, "int",
            formula=blank(
                "parent1",
                f'COUNTIFS({child_family},{{id}}{{r}},'
                f'{child_status},"Active")'),
            note="Children currently attending. A family with none is "
                 "an old family, not a lost one - they may be back."),
        Col("child1", "First child", 18, "text",
            formula=blank("parent1", nth_child(1))),
        Col("child2", "Second child", 18, "text",
            formula=blank("parent1", nth_child(2))),
        Col("child3", "Third child", 18, "text",
            formula=blank("parent1", nth_child(3)), hidden=True),
        Col("more_children", "And others", 11, "text",
            formula=blank(
                "parent1",
                'IF({children}{r}>3,"+"&({children}{r}-3)&" more","")'),
            hidden=True),
        Col("weekly_fees", "Fees a week", 12, "money",
            formula=blank(
                "parent1",
                'ROUND(SUMIFS([[Child Enrollment.weekly_net]],'
                f'{child_family},{{id}}{{r}},'
                '[[Child Enrollment.holds_place]],"Yes"),2)'),
            cf=bar_cf(Palette.accent),
            note="Everything this family is contracted to pay in a "
                 "normal week, across all their children."),
        Col("monthly_fees", "Fees a month", 12, "money",
            formula=blank("parent1",
                          'ROUND({weekly_fees}{r}*52/12,2)'),
            hidden=True),
        # -------------------------------------------------- the money ---
        Col("billed", "Billed", 12, "money",
            formula=blank(
                "parent1",
                f'ROUND(SUMIFS([[Invoices & Receipts.owed]],'
                f'{led_family},{{id}}{{r}}),2)'),
            hidden=True,
            note="Net effect of every invoice, credit note, write-off "
                 "and refund on this family's account."),
        Col("invoiced", "Invoiced", 12, "money",
            formula=blank(
                "parent1",
                f'ROUND(SUMIFS([[Invoices & Receipts.total]],'
                f'{led_family},{{id}}{{r}},{led_type},"Invoice"),2)')),
        Col("paid", "Paid", 12, "money",
            formula=blank(
                "parent1",
                f'ROUND(SUMIFS([[Invoices & Receipts.collected]],'
                f'{led_family},{{id}}{{r}}),2)'),
            note="Payments received less any refunds given back."),
        Col("balance", "Balance", 12, "money",
            formula=blank(
                "parent1",
                'ROUND({billed}{r}-{paid}{r},2)'),
            cf=owing_cf(),
            note="Positive means the family owes you. Negative means "
                 "they are in credit - usually a deposit waiting to be "
                 "used."),
        Col("overdue", "Overdue", 12, "money",
            formula=blank(
                "parent1",
                'ROUND(SUMIFS([[Invoices & Receipts.balance]],'
                f'{led_family},{{id}}{{r}},'
                '[[Invoices & Receipts.status]],"Overdue"),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "bold_bad"}],
            note="Only the invoices that are past their due date. "
                 "This is the chase list."),
        Col("oldest_due", "Oldest unpaid", 12, "date",
            # the earliest due date still unpaid - the one that has
            # been outstanding longest. Reads the numeric mirror of
            # the due date, never the dated column, which is blank on
            # payments and on empty rows.
            formula=blank(
                "parent1",
                'IF({overdue}{r}<=0.004,"",IF(' +
                _oldest_due +
                '>=2958465,"",' + _oldest_due + '))'),
            hidden=True),
        Col("last_paid", "Last payment", 12, "date",
            formula=blank(
                "parent1",
                'IF({paid}{r}=0,"",' +
                max_if("[[Invoices & Receipts.date]]",
                       f"{led_family}={{id}}{{r}}",
                       f'{led_type}="Payment"') + ")")),
        Col("days_since_paid", "Days since", 9, "int",
            formula=blank(
                "parent1",
                'IF({last_paid}{r}="","",TodayDate-{last_paid}{r})'),
            hidden=True),
        Col("subsidy", "On subsidy?", 11, "flag",
            formula=blank(
                "parent1",
                'IF(SUMIFS([[Child Enrollment.subsidy_week]],'
                f'{child_family},{{id}}{{r}})>0,"Yes","No")'),
            hidden=True),
        # ------------------------------------------------ relationship --
        Col("joined", "With us since", 12, "date",
            validate=date_check(),
            note="Leave blank and nothing breaks - it is here for the "
                 "families who have been with you for years."),
        Col("messages", "Messages logged", 12, "int",
            formula=blank(
                "parent1",
                'COUNTIFS([[Parent Communication.family]],{id}{r})'),
            hidden=True),
        Col("open_items", "Open items", 10, "int",
            formula=blank(
                "parent1",
                'COUNTIFS([[Parent Communication.family]],{id}{r},'
                '[[Parent Communication.status]],"Open")+'
                'COUNTIFS([[Parent Communication.family]],{id}{r},'
                '[[Parent Communication.status]],"Waiting on us")'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Conversations waiting on you. Not on the parent - "
                 "on you."),
        Col("status", "Family status", 13, "status",
            formula=blank(
                "parent1",
                'IF({children}{r}=0,"No children linked",'
                'IF({overdue}{r}>0.004,"Overdue",'
                'IF({children_active}{r}=0,"Inactive",'
                'IF({balance}{r}>0.004,"Owing","Active"))))'),
            cf=status_cf()),
        Col("notes", "Notes", 28, "wrap"),
        rank_desc("weekly_fees", guard='{parent1}{r}=""',
                  header="Value rank"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "parent1",
                'IF(COUNTIF({parent1_first}:{parent1_last},'
                '{parent1}{r})>1,"Duplicate - same contact twice",'
                'IF({children}{r}=0,"Add - no children linked yet",'
                'IF({overdue}{r}>0.004,"Overdue - '
                + str(cfg.currency_symbol) +
                '"&TEXT({overdue}{r},"#,##0.00")&" past due",'
                'IF({phone1}{r}="","Missing - no phone number",'
                'IF({em1_name}{r}="","Missing - no emergency contact",'
                'IF({pickup_ok}{r}="No","Not authorised - pickup list '
                'unsigned",'
                'IF({open_items}{r}>0,"Follow up - "&{open_items}{r}&'
                '" message(s) waiting on you",'
                'IF({balance}{r}>0.004,"Chase - balance outstanding",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Parent & Guardian Database",
        subtitle="One row per family. Siblings share it, invoices "
                 "follow it, and the balance is the family's, not the "
                 "child's.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).parents,
        help_text="Type the family's details once. Link each child to "
                  "the family ID on the Child Enrollment tab and the "
                  "children, the weekly fees and the balance appear "
                  "here on their own.",
        extra_notes=[
            "Children, fees, invoices, payments, balance and overdue "
            "amounts are all calculated. The only money column you "
            "ever type on this tab is nothing at all - it comes from "
            "Invoices & Receipts.",
            "Emergency contacts and the authorised collection list "
            "are the two columns worth auditing termly. The "
            "attendance tab checks collections against them.",
            "Custody notes and anything else sensitive should be a "
            "pointer to a secure file, not the detail itself.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Family book summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The family book", last)
    sheet.table_head(row, 0, [
        "Families", "With a child attending", "Siblings families",
        "Weekly contracted fees", "Total invoiced", "Total paid",
        "Outstanding", "Overdue", "Families owing", "On subsidy",
        "Missing emergency contact"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Parent Database.parent1]],"<>")',
        '=COUNTIFS([[Parent Database.children_active]],">0")',
        '=COUNTIFS([[Parent Database.children]],">1")',
        '=ROUND(SUM([[Parent Database.weekly_fees]]),2)',
        '=ROUND(SUM([[Parent Database.invoiced]]),2)',
        '=ROUND(SUM([[Parent Database.paid]]),2)',
        '=ROUND(SUM([[Parent Database.balance]]),2)',
        '=ROUND(SUM([[Parent Database.overdue]]),2)',
        '=COUNTIFS([[Parent Database.balance]],">0.004")',
        '=COUNTIFS([[Parent Database.subsidy]],"Yes")',
        '=COUNTIFS([[Parent Database.parent1]],"<>")-'
        'COUNTIFS([[Parent Database.em1_name]],"<>")',
    ], ["int", "int", "int", "money", "money", "money", "money",
        "money", "int", "int", "int"])
