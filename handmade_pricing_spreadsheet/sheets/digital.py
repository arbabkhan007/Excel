"""`Digital` - patterns and downloads: all the cost is up front.

A PDF pattern costs nothing to make a second copy of, so the whole
question is how many you will sell and whether the design time is
ever recovered. This tab spreads the design, testing, photography and
writing hours over the sales you honestly expect, adds the per sale
costs that do exist, and tells you how many you have to sell before
the file has paid for itself.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Digital"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Download", 28, "text",
            note="'Cable beanie crochet pattern', 'Farmhouse SVG "
                 "bundle'."),
        Col("kind", "Type", 20, "text",
            validate=list_from("DigitalTypeList")),
        Col("formats", "Formats", 16, "text",
            validate=list_from("FileFormatList")),
        Col("licence", "Licence", 20, "text",
            validate=list_from("LicenceList"),
            note="A commercial licence is a different product at a "
                 "different price, not a favour."),
        Col("design_hours", "Design hours", 11, "number",
            validate=number_check(0, 10000),
            note="Making the thing in the first place - the "
                 "sample, the grading, the drawing."),
        Col("write_hours", "Writing hours", 11, "number",
            validate=number_check(0, 10000),
            note="Writing the pattern up, charts, diagrams, "
                 "layout. Usually longer than making it."),
        Col("test_hours", "Testing hours", 11, "number",
            validate=number_check(0, 10000),
            note="Your time running a test round and correcting "
                 "it. Tester payments go in the fees column."),
        Col("photo_hours", "Photography hours", 13, "number",
            validate=number_check(0, 10000)),
        Col("tester_cost", "Tester & editor fees", 14, "money",
            validate=number_check(0, 100000),
            note="Free patterns for testers, a tech editor, a "
                 "proofreader. Real money, paid once."),
        Col("software_cost", "Software & stock", 12, "money",
            validate=number_check(0, 100000),
            note="The share of a subscription or stock licence "
                 "that belongs to this file."),
        Col("expected_sales", "Sales expected", 12, "int",
            validate=number_check(1, 1000000, decimal=False),
            note="Over the life of the listing. Be honest - this "
                 "single number decides whether the design time "
                 "is ever recovered."),
        Col("support_mins", "Support minutes a sale", 15, "number1",
            validate=number_check(0, 1000),
            note="Questions, errata, 'it will not open'. Two "
                 "minutes a sale is a real cost at volume."),
        Col("update_hours", "Updates a year", 11, "number",
            validate=number_check(0, 10000),
            note="Revisions, new sizes, a format nobody could "
                 "open. Spread over a year of sales."),
        Col("sold_so_far", "Sold so far", 11, "int",
            validate=number_check(0, 10000000, decimal=False),
            note="Lifetime sales of this file, for the recovery "
                 "figure."),
        Col("listed", "Listed on", 11, "date", validate=date_check()),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("dev_hours", "Development hours", 13, "number",
            formula=blank(
                "name",
                'IF({design_hours}{r}="",0,{design_hours}{r})+'
                'IF({write_hours}{r}="",0,{write_hours}{r})+'
                'IF({test_hours}{r}="",0,{test_hours}{r})+'
                'IF({photo_hours}{r}="",0,{photo_hours}{r})+'
                'IF({update_hours}{r}="",0,{update_hours}{r})')),
        Col("dev_cost", "Development cost", 13, "money",
            formula=blank(
                "name",
                'ROUND({dev_hours}{r}*{rate}{r}+'
                'IF({tester_cost}{r}="",0,{tester_cost}{r}),2)'),
            note="Everything the file cost to exist, paid once, "
                 "before a single sale."),
        Col("expected", "Sales used", 10, "int",
            formula=blank(
                "name",
                'MAX(1,IF({expected_sales}{r}="",1,'
                '{expected_sales}{r}))'),
            hidden=True),
        Col("mat_cost", "Materials", 10, "money",
            formula=blank("name", '0'),
            hidden=True,
            note="A download has no materials - which is exactly "
                 "why the development time has to be taken "
                 "seriously."),
        Col("minutes", "Minutes a sale", 12, "number1",
            formula=blank(
                "name",
                'ROUND({dev_hours}{r}*60/{expected}{r}+'
                'IF({support_mins}{r}="",0,{support_mins}{r}),1)')),
        Col("labor_cost", "Labour a sale", 12, "money",
            formula=blank(
                "name",
                'ROUND({dev_cost}{r}/{expected}{r}+'
                'IF({support_mins}{r}="",0,{support_mins}{r})/60*'
                '{rate}{r},2)'),
            note="The development cost spread over the sales you "
                 "expect, plus the support each sale really takes."),
        Col("other_cost", "Extras a sale", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({software_cost}{r}="",0,{software_cost}{r})/'
                '{expected}{r},2)'),
            hidden=True),
        Col("recovered", "Recovered so far", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({sold_so_far}{r}="",0,{sold_so_far}{r})*'
                'MAX(0,{profit}{r}+{labor_cost}{r}-'
                'IF({support_mins}{r}="",0,{support_mins}{r})/60*'
                '{rate}{r}),2)'),
            note="What the sales so far have put back against the "
                 "development cost."),
        Col("recovery", "Development recovered", 13, "percent",
            formula=blank(
                "name",
                'IF({dev_cost}{r}=0,"",{recovered}{r}/{dev_cost}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "warn"}]),
        Col("break_even", "Sales to break even", 14, "int",
            formula=blank(
                "name",
                'IF(OR({price}{r}="",{price}{r}-{other_cost}{r}-'
                '({price}{r}*{fee_pct}{r}+{fee_fixed}{r})-'
                'IF({support_mins}{r}="",0,{support_mins}{r})/60*'
                '{rate}{r}<=0),"Never at this price",'
                'ROUNDUP({dev_cost}{r}/({price}{r}-{other_cost}{r}-'
                '({price}{r}*{fee_pct}{r}+{fee_fixed}{r})-'
                'IF({support_mins}{r}="",0,{support_mins}{r})/60*'
                '{rate}{r}),0))'),
            note="How many copies pay for the design time. If that "
                 "number is bigger than the sales you expect, the "
                 "price is too low or the file is too ambitious."),
    ] + craft_tail() + [craft_check(needs="labor_cost")]

    return SheetSpec(
        name=NAME,
        title="Digital Product Calculator",
        subtitle="Design time spread over the sales you expect, and "
                 "the day it pays for itself.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).digital,
        help_text="All the cost is up front. Enter the hours "
                  "honestly and the sales you expect "
                  "conservatively, and the per sale cost falls out.",
        extra_notes=[
            "Writing a pattern up takes longer than making the "
            "thing. Twelve hours of design and sixteen of writing "
            "at $25 is $700 before a single sale - over 200 copies "
            "that is $3.50 each.",
            "A download that sells steadily is the best margin in "
            "handmade. One that sells twice is an expensive way to "
            "spend a fortnight. Expect the second, be delighted by "
            "the first.",
            "Support time is real. Two minutes a sale at $25 an "
            "hour is 83 cents - on a $6 pattern that is 14% of the "
            "price.",
        ],
    )
