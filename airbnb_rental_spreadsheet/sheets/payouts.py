"""`Platform Revenue & Fees` - what the platform said, and what arrived.

Every host has had the moment where the money in the bank does not
match the money on the booking.  A cancelled night, a currency
conversion, a resolution centre adjustment, a fee that changed when the
account moved plan - none of it is visible from the reservation screen.

This is the reconciliation ledger.  One row per payout: what the
booking said you should get, what actually landed, and the difference
between the two, with the difference shown in money and as a
percentage, because a consistent two per cent gap is a fee rate that
needs correcting on Settings rather than a mystery.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_booking,
                     lookup_booking_num, money_cf, month_col, owing_cf,
                     status_cf, year_col)

NAME = "Platform Revenue & Fees"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("booking", "PayoutPrefix", digits=cfg.digits("payout"),
               width=10, header="Payout ID"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="The reservation this payout belongs to. One row per "
                 "payout - split a part payment across two rows.",
            group="The payout"),
        Col("property", "Property", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="The payout"),
        Col("property_name", "Property", 19, "text",
            formula=lookup_booking("property_name"), group="The payout"),
        Col("guest_name", "Guest", 18, "text",
            formula=lookup_booking("guest_name"), group="The payout"),
        Col("platform", "Platform", 15, "text",
            formula=lookup_booking("platform"), group="The payout"),
        Col("check_out", "Check-out", 12, "date",
            formula=lookup_booking("check_out"), hidden=True,
            group="The payout"),

        # -------------------------------------------- what was expected --
        Col("exp_gross", "Gross booked", 12, "money",
            formula=lookup_booking_num("gross"),
            note="Straight from the booking - everything the guest paid.",
            group="Expected"),
        Col("exp_fee", "Fees expected", 12, "money",
            formula=lookup_booking_num("platform_fee_total"),
            group="Expected"),
        Col("exp_net", "Payout expected", 13, "money",
            formula=lookup_booking_num("net_revenue"),
            note="What the booking tab says should reach your account.",
            group="Expected"),

        # --------------------------------------------- what turned up --
        Col("payout_date", "Paid on", 12, "date",
            validate=date_check(),
            note="The date the money actually reached your account.",
            group="Received"),
        month_col("payout_date"),
        year_col("payout_date"),
        Col("amount_in", "Amount received", 14, "money",
            validate=number_check(-1000000, 1000000, "Amount received"),
            note="What landed. Take it from the bank statement, not "
                 "from the platform dashboard.",
            group="Received"),
        Col("fee_actual", "Fees actually taken", 14, "money",
            validate=number_check(0, 1000000, "Fees taken"),
            note="Optional. If the payout statement itemises the fee, "
                 "put it here and the fee-rate column below will tell "
                 "you whether your Settings rate is right.",
            group="Received"),
        Col("refund", "Refund given", 12, "money",
            validate=number_check(0, 1000000, "Refund"),
            note="Money you gave back - a cancellation, a goodwill "
                 "gesture, a resolution centre award. Positive number.",
            group="Received"),
        Col("adjustment", "Adjustment", 12, "money",
            validate=number_check(-1000000, 1000000, "Adjustment"),
            note="Anything else: a damage claim paid to you, a currency "
                 "difference, a host fee refund. Can be negative.",
            group="Received"),
        Col("received", "Net received", 13, "money",
            formula=blank(
                "booking",
                'ROUND({amount_in}{r}-{refund}{r}+{adjustment}{r},2)'),
            cf=money_cf(),
            note="Amount in, less refunds, plus adjustments. This is the "
                 "figure the financial reports treat as cash.",
            group="Received"),
        Col("status", "Status", 14, "status",
            validate=list_from("PayoutStatusList"),
            cf=status_cf(), group="Received"),
        Col("reference", "Statement ref", 14, "text", group="Received"),

        # -------------------------------------------------- the gap ---
        Col("variance", "Difference", 12, "money",
            formula=blank(
                "booking",
                'IF({payout_date}{r}="","",'
                'ROUND({received}{r}-{exp_net}{r},2))'),
            cf=owing_cf(),
            note="Received less expected. Amber when the platform kept "
                 "more than you thought, blue when it kept less.",
            group="The gap"),
        Col("variance_pct", "Difference %", 11, "percent",
            formula=blank(
                "booking",
                'IF({payout_date}{r}="","",IF({exp_net}{r}<=0,"",'
                '{variance}{r}/{exp_net}{r}))'),
            hidden=True, group="The gap"),
        Col("fee_rate_actual", "Real fee rate", 12, "percent",
            formula=blank(
                "booking",
                'IF({fee_actual}{r}="","",IF({exp_gross}{r}<=0,"",'
                '{fee_actual}{r}/{exp_gross}{r}))'),
            note="What the platform really charged, as a percentage of "
                 "the gross. If this keeps disagreeing with your "
                 "Settings rate, change the Settings rate.",
            group="The gap"),
        Col("days_to_pay", "Days to pay", 11, "int",
            formula=blank(
                "booking",
                'IF({payout_date}{r}="","",IF({check_out}{r}="","",'
                '{payout_date}{r}-{check_out}{r}))'),
            group="The gap"),
        Col("is_paid", "Money in", 9, "flag",
            formula=blank("booking",
                          'IF({payout_date}{r}="","No","Yes")'),
            cf=flag_yes_no_cf(), hidden=True, group="The gap"),

        Col("alert", "Action", 28, "status",
            formula=blank(
                "booking",
                'IF({exp_net}{r}="","Check - booking not found",'
                'IF({payout_date}{r}="",'
                'IF({check_out}{r}="","Check - no check-out date",'
                'IF(TodayDate-{check_out}{r}>PayoutChaseDays,'
                '"Chase - payout is overdue","Due - payout expected '
                'shortly")),'
                'IF({status}{r}="Disputed","Action - disputed payout",'
                'IF(ABS({variance}{r})<=0.5,"OK",'
                'IF({variance}{r}<0,"Check - less arrived than expected",'
                '"Check - more arrived than expected")))))'),
            cf=alert_cf(), group="The gap"),
        Col("notes", "Notes", 30, "wrap", group="The gap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Platform Revenue & Fees",
        subtitle="What each platform said it would pay you, and what "
                 "actually landed",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("payouts", cfg),
        help_text="One row per payout. Pick the booking and the expected "
                  "figures arrive from the booking tab - then type what "
                  "really hit your bank account. The Difference column "
                  "is the point of this tab: small and random is "
                  "normal, consistent and one-sided means a fee rate "
                  "on Settings needs correcting.",
        extra_notes=[
            "Take the amount received from your bank statement rather "
            "than the platform dashboard. The dashboard shows what the "
            "platform intends to send; the statement shows what "
            "survived the currency conversion.",
            "A payout with no date is chased automatically once it is "
            "more days past check-out than the chase window on "
            "Settings.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Reconciliation summary, then a per-platform fee check."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Reconciliation", last)
    row = sheet.table_head(row, 0, [
        "Payouts logged", "Received", "Still expected", "Expected total",
        "Actually received", "Refunds", "Adjustments", "Difference",
        "Avg days to pay"],
        [13, 11, 13, 13, 15, 11, 12, 12, 14])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Platform Revenue & Fees.booking]],"?*")',
        '=COUNTIFS([[Platform Revenue & Fees.is_paid]],"Yes")',
        '=COUNTIFS([[Platform Revenue & Fees.is_paid]],"No")',
        '=ROUND(SUM([[Platform Revenue & Fees.exp_net]]),2)',
        '=ROUND(SUM([[Platform Revenue & Fees.received]]),2)',
        '=ROUND(SUM([[Platform Revenue & Fees.refund]]),2)',
        '=ROUND(SUM([[Platform Revenue & Fees.adjustment]]),2)',
        '=ROUND(SUM([[Platform Revenue & Fees.received]])-'
        'SUMIFS([[Platform Revenue & Fees.exp_net]],'
        '[[Platform Revenue & Fees.is_paid]],"Yes"),2)',
        '=IF(COUNTIFS([[Platform Revenue & Fees.is_paid]],"Yes")<=0,"",'
        'ROUND(SUM([[Platform Revenue & Fees.days_to_pay]])/'
        'COUNTIFS([[Platform Revenue & Fees.is_paid]],"Yes"),1))',
    ], ["int", "int", "int", "money", "money", "money", "money", "money",
        "number1"])
    row += 2

    row = sheet.section(row, "Is your fee rate right?", last)
    row = sheet.table_head(row, 0, [
        "Platform", "Payouts", "Gross booked", "Fees you assumed",
        "Fees really taken", "Assumed rate", "Real rate", "Difference"],
        [18, 10, 13, 15, 16, 12, 11, 12])
    for name, _commission, _payment, _counts, _note in cfg.platform_table:
        key = f'"{name}"'
        plat = f'[[Platform Revenue & Fees.platform]],{key}'
        gross = f'SUMIFS([[Platform Revenue & Fees.exp_gross]],{plat})'
        assumed = f'SUMIFS([[Platform Revenue & Fees.exp_fee]],{plat})'
        real = (f'SUMIFS([[Platform Revenue & Fees.fee_actual]],{plat})')
        with_real = (f'SUMIFS([[Platform Revenue & Fees.exp_gross]],{plat},'
                     f'[[Platform Revenue & Fees.fee_actual]],">0")')
        sheet.calc_row(row, 0, [
            name,
            f"=COUNTIFS({plat})",
            f"=ROUND({gross},2)",
            f"=ROUND({assumed},2)",
            f"=ROUND({real},2)",
            f'=IF({gross}<=0,"",{assumed}/{gross})',
            f'=IF({with_real}<=0,"",{real}/{with_real})',
            f'=IF({with_real}<=0,"",ROUND({real}-'
            f'SUMIFS([[Platform Revenue & Fees.exp_fee]],{plat},'
            f'[[Platform Revenue & Fees.fee_actual]],">0"),2))',
        ], ["text", "int", "money", "money", "money", "percent",
            "percent", "money"])
        row += 1

    row += 1
    sheet.note(row, "The real rate only counts payouts where you typed "
                    "the fee the statement showed, so it is blank until "
                    "you have filled a few in. When it settles a long "
                    "way from your assumed rate, change the rate on "
                    "Settings and every future booking prices itself "
                    "correctly.", last)
