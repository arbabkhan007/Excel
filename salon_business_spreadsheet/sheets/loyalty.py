"""`Loyalty` - points, rewards and referrals, one row per client.

Points are earned on spend at whatever rate you set, redeemed when
you say so, and the balance is simply the difference. Referrals are
counted from the Clients tab, so the credit you owe somebody for
sending a friend is never a matter of memory.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_client, rank_desc,
                     status_cf, yesno)

NAME = "Loyalty"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("client", "LoyaltyPrefix", digits=4, width=11),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList"),
            note="One row per client on the scheme."),
        Col("name", "Client", 22, "text",
            formula=lookup_client("name", key="client")),
        Col("joined", "Joined", 11, "date", validate=date_check()),
        Col("card", "Card number", 14, "text",
            note="Only if you use physical cards. Otherwise leave it "
                 "blank - the client ID is enough."),
        Col("bonus_points", "Bonus points", 12, "int",
            validate=number_check(0, 1000000),
            note="Points you have handed out by hand - a birthday "
                 "gift, an apology, a launch promotion."),
        Col("redeemed", "Points redeemed", 14, "int",
            validate=number_check(0, 1000000),
            note="Total points spent so far. Add to this number each "
                 "time a reward is taken."),
        Col("last_reward", "Last reward", 11, "date",
            validate=date_check()),
        Col("reward_type", "Usual reward", 20, "text",
            validate=list_from("LoyaltyRewardList")),
        Col("referral_paid", "Referral credit paid", 15, "money",
            validate=number_check(0, 100000),
            note="What you have already given them for sending "
                 "people in."),
        Col("active", "On the scheme?", 13, "flag", validate=yesno()),
        Col("notes", "Notes", 24, "wrap"),
        # ---------------------------------------------------- worked out --
        Col("spend", "Spend counted", 12, "money",
            formula=blank(
                "client",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],'
                '{client}{r},[[Sales.date]],">="&IF({joined}{r}="",0,'
                '{joined}{r})),2)'),
            note="Everything they have spent since joining - services "
                 "and retail."),
        Col("visits", "Visits", 8, "int",
            formula=blank(
                "client",
                'IFERROR(INDEX([[Clients.visits]],MATCH({client}{r},'
                '[[Clients.id]],0))*1,0)'),
            hidden=True),
        Col("earned", "Points earned", 12, "int",
            formula=blank(
                "client",
                'ROUND({spend}{r}*PointsPerUnit+IF({bonus_points}{r}='
                '"",0,{bonus_points}{r}),0)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Spend x the points rate on Settings, plus any "
                 "bonus points."),
        Col("balance", "Points balance", 12, "int",
            formula=blank(
                "client",
                '{earned}{r}-IF({redeemed}{r}="",0,{redeemed}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": "=RewardAt",
                 "fmt": "good"}],
            note="Earned less redeemed. This is the number the client "
                 "cares about."),
        Col("rewards_due", "Rewards due", 11, "int",
            formula=blank(
                "client",
                'IF(RewardAt<=0,0,ROUNDDOWN({balance}{r}/RewardAt,0))'),
            hidden=True),
        Col("reward_value", "Reward owed", 12, "money",
            formula=blank(
                "client",
                'ROUND({rewards_due}{r}*RewardValue,2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="What those points are worth if they cash them in "
                 "today. Worth knowing before you run a double "
                 "points week."),
        Col("to_next", "Points to next", 12, "int",
            formula=blank(
                "client",
                'IF(RewardAt<=0,"",MAX(0,RewardAt*({rewards_due}{r}+1)-'
                '{balance}{r}))'),
            note="The number to tell them at the desk. 'Forty more "
                 "points and your next treatment is on us.'"),
        Col("referrals", "People referred", 13, "int",
            formula=blank(
                "client",
                'COUNTIFS([[Clients.referred_by]],{client}{r})'),
            note="Counted from the 'referred by' column on the "
                 "Clients tab."),
        Col("referral_owed", "Referral credit owed", 15, "money",
            formula=blank(
                "client",
                'ROUND({referrals}{r}*ReferralReward-'
                'IF({referral_paid}{r}="",0,{referral_paid}{r}),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("segment", "Segment", 13, "text",
            formula=blank(
                "client",
                'IFERROR(INDEX([[Clients.segment]],MATCH({client}{r},'
                '[[Clients.id]],0))&"","")'),
            cf=status_cf(),
            hidden=True),
        rank_desc("balance", guard='{client}{r}=""', header="Points rank"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "client",
                'IF({active}{r}="No","",'
                'IF(COUNTIF({client_first}:{client_last},{client}{r})>1,'
                '"Duplicate - this client is on twice",'
                'IF({rewards_due}{r}>0,"Action - reward ready to give",'
                'IF({referral_owed}{r}>0.004,'
                '"Action - referral credit owed",'
                'IF({segment}{r}="Lapsing",'
                '"Chase - points are a reason to call",'
                'IF({to_next}{r}<=RewardAt*0.2,'
                '"Rebook - nearly at a reward","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Client Loyalty",
        subtitle="Points, rewards and referral credit - counted from "
                 "the spend you have already logged.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).loyalty,
        help_text="Add a row per client on the scheme. Points build "
                  "themselves from the Sales tab; you only record "
                  "what has been redeemed.",
        extra_notes=[
            "Set the points rate, the reward threshold and the reward "
            "value on Settings. One point per unit of spend with a "
            "reward at 200 points is a 10% scheme - worth knowing "
            "before you promise it.",
            "The reward owed column is a liability, not a decoration. "
            "It is what you would have to hand over if every client "
            "cashed in tomorrow.",
            "Referral credit is counted from the 'referred by' column "
            "on the Clients tab, so the person who sent three friends "
            "in gets what they were promised without anybody having "
            "to remember.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Loyalty summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The scheme", last)
    sheet.table_head(row, 0, [
        "Members", "Active", "Points earned", "Points redeemed",
        "Points outstanding", "Rewards ready", "Reward value owed",
        "Referrals", "Referral credit owed"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Loyalty.client]],"<>")',
        '=COUNTIFS([[Loyalty.active]],"Yes")',
        '=SUM([[Loyalty.earned]])',
        '=SUM([[Loyalty.redeemed]])',
        '=SUM([[Loyalty.balance]])',
        '=SUM([[Loyalty.rewards_due]])',
        '=ROUND(SUM([[Loyalty.reward_value]]),2)',
        '=SUM([[Loyalty.referrals]])',
        '=ROUND(SUM([[Loyalty.referral_owed]]),2)',
    ], ["int", "int", "int", "int", "int", "int", "money", "int",
        "money"])
