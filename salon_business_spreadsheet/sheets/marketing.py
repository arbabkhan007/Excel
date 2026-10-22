"""`Marketing` - campaigns, offers and what they actually brought in.

A promo code on a sale is the link. Type the code here, use the same
code on the till, and this tab counts the bookings, the revenue, the
discount given away and the new clients - then divides by what the
campaign cost you.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, month_col, rank_desc,
                     status_cf)

NAME = "Marketing"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "CampaignPrefix", digits=3, width=10),
        Col("name", "Campaign", 26, "text",
            note="What you would call it out loud - 'January colour "
                 "offer', 'refer a friend', 'Tuesday blow dry'."),
        Col("channel", "Channel", 18, "text",
            validate=list_from("MarketingChannelList")),
        Col("offer", "The offer", 28, "text",
            note="20% off a first colour, free treatment with a cut, "
                 "two for one on Tuesdays."),
        Col("code", "Promo code", 13, "text",
            note="The code you type on the Sales tab. This is what "
                 "links the money back to the campaign - keep it "
                 "short and keep it unique."),
        Col("start", "Starts", 11, "date", validate=date_check()),
        Col("end", "Ends", 11, "date", validate=date_check()),
        Col("spend", "Cost", 11, "money", validate=number_check(0, 1000000),
            note="Ads, print, the influencer's treatment at cost, the "
                 "hour you spent making the reel."),
        Col("target_rev", "Revenue target", 13, "money",
            validate=number_check(0, 1000000)),
        Col("status", "Status", 12, "text",
            validate=list_from("CampaignStatusList"), cf=status_cf()),
        Col("notes", "What happened", 28, "wrap",
            note="Worth a sentence while it is fresh. Next January "
                 "you will want to know whether it was worth "
                 "repeating."),
        # ---------------------------------------------------- worked out --
        month_col("start"),
        Col("uses", "Times used", 11, "int",
            formula=blank(
                "name",
                'IF({code}{r}="",0,COUNTIFS([[Sales.promo]],'
                '{code}{r}))'),
            note="Sales lines carrying this code."),
        Col("revenue", "Revenue", 12, "money",
            formula=blank(
                "name",
                'IF({code}{r}="",0,ROUND(SUMIFS([[Sales.net]],'
                '[[Sales.promo]],{code}{r}),2))'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Net sales on those lines - after the discount the "
                 "offer gave away."),
        Col("given", "Discount given", 13, "money",
            formula=blank(
                "name",
                'IF({code}{r}="",0,ROUND(SUMIFS([[Sales.discount]],'
                '[[Sales.promo]],{code}{r}),2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("clients", "Clients used it", 12, "int",
            formula=blank(
                "name",
                'IF({code}{r}="",0,SUMPRODUCT(([[Sales.promo]]='
                '{code}{r})*([[Sales.first_of_appt]]="Yes")))'),
            hidden=True),
        Col("new_clients", "New clients", 11, "int",
            formula=blank(
                "name",
                'IF({start}{r}="",0,COUNTIFS([[Clients.source]],'
                '{channel}{r},[[Clients.first_visit]],">="&{start}{r},'
                '[[Clients.first_visit]],"<="&IF({end}{r}="",'
                'TodayDate,{end}{r})))'),
            note="Clients whose 'found you by' matches this channel "
                 "and whose first visit falls inside the dates."),
        Col("bookings", "Bookings", 10, "int",
            formula=blank(
                "name",
                'IF({start}{r}="",0,COUNTIFS([[Appointments.source]],'
                '{channel}{r},[[Appointments.date]],">="&{start}{r},'
                '[[Appointments.date]],"<="&IF({end}{r}="",TodayDate,'
                '{end}{r})))'),
            hidden=True),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({spend}{r}="",0,{spend}{r})+{given}{r},2)'),
            note="What you paid out plus what you gave away. Both are "
                 "the cost of the offer."),
        Col("profit", "Left over", 11, "money",
            formula=blank(
                "name",
                'ROUND({revenue}{r}-{total_cost}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}]),
        Col("roi", "Return", 10, "percent",
            formula=blank(
                "name",
                'IF({total_cost}{r}<=0,"",'
                '{profit}{r}/{total_cost}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"}],
            note="What you got back for every pound or dollar the "
                 "campaign cost. 100% means it doubled its money."),
        Col("cost_per_client", "Cost per new client", 14, "money",
            formula=blank(
                "name",
                'IF({new_clients}{r}=0,"",'
                'ROUND({total_cost}{r}/{new_clients}{r},2))'),
            cf=[{"type": "cell", "criteria": ">", "value": "=TargetTicket",
                 "fmt": "warn"}],
            note="Over the price of an average visit means the "
                 "campaign only pays if they come back - which is "
                 "exactly what the retention tab is for."),
        Col("vs_target", "Against target", 12, "money",
            formula=blank(
                "name",
                'IF({target_rev}{r}="","",'
                'ROUND({revenue}{r}-{target_rev}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "warn"}],
            hidden=True),
        rank_desc("revenue", header="Best earner"),
        Col("alert", "Verdict", 28, "status",
            formula=blank(
                "name",
                'IF({code}{r}="","Add - no code, nothing to measure",'
                'IF({status}{r}="Planned","Upcoming",'
                'IF({uses}{r}=0,"Check - code never used",'
                'IF({profit}{r}<0,"Loss - cost more than it made",'
                'IF({roi}{r}>=1,"OK - doubled its money",'
                'IF({roi}{r}>0,"OK - made money","Check - review it"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Marketing & Promotions",
        subtitle="Offers, codes and campaigns, measured against what "
                 "they cost - money out and discount given away.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).marketing,
        help_text="Give every offer a code and use that code on the "
                  "Sales tab. Without the code, nothing can be traced "
                  "back here.",
        extra_notes=[
            "The discount you gave away is counted as part of the "
            "cost. A full column of half-price blow dries is not free "
            "marketing, and this tab will not pretend it was.",
            "New clients are matched by channel and date, so keep the "
            "'found you by' column on the Clients tab filled in. It "
            "takes two seconds at the basin and it is the only way "
            "the number means anything.",
            "Cost per new client against your average ticket is the "
            "honest test of an offer. If it costs more to win them "
            "than they spend on the first visit, the campaign is a "
            "bet on them coming back.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Marketing summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "What the marketing did", last)
    sheet.table_head(row, 0, [
        "Campaigns", "Running now", "Spent", "Discount given",
        "Revenue traced", "Left over", "New clients", "Average return",
        "Cost per new client"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Marketing.name]],"<>")',
        '=COUNTIFS([[Marketing.status]],"Running")',
        '=ROUND(SUM([[Marketing.spend]]),2)',
        '=ROUND(SUM([[Marketing.given]]),2)',
        '=ROUND(SUM([[Marketing.revenue]]),2)',
        '=ROUND(SUM([[Marketing.profit]]),2)',
        '=SUM([[Marketing.new_clients]])',
        '=IFERROR(SUM([[Marketing.profit]])/'
        'MAX(0.01,SUM([[Marketing.total_cost]])),0)',
        '=IFERROR(ROUND(SUM([[Marketing.total_cost]])/'
        'MAX(1,SUM([[Marketing.new_clients]])),2),0)',
    ], ["int", "int", "money", "money", "money", "money", "int",
        "percent", "money"])
