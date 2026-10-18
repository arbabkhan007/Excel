"""`Packaging` - the box, the tissue, the card, the sticker.

Packaging is the cost makers forget and buyers judge you on. List
every component once, say which of the four packaging profiles it
belongs to - Budget, Standard, Gift, Premium - and the workbook adds
up what each profile costs and weighs. A product then picks a profile
and inherits the lot.

The totals at the bottom are published as PackBudget, PackStandard,
PackGift and PackPremium (and the matching weights), which is how the
catalog and both price calculators read them.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..layout import quote_sheet
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf

NAME = "Packaging"

PROFILES = [
    ("budget", "Budget", "PackBudget", "PackWeightBudget",
     "Plain mailer, a sticker, done. For low priced items where the "
     "packaging cannot be more than the profit."),
    ("standard", "Standard", "PackStandard", "PackWeightStandard",
     "What most of your orders go out in."),
    ("gift", "Gift", "PackGift", "PackWeightGift",
     "Ribbon, tissue, a hand written card - the one people "
     "photograph."),
    ("premium", "Premium", "PackPremium", "PackWeightPremium",
     "Rigid box, insert, branded sleeve. For the pieces that carry "
     "it."),
]


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("name", "Component", 26, "text",
            note="Mailer, tissue, thank you card, twine, fragile "
                 "sticker - one row each."),
        Col("type", "Type", 18, "text", validate=list_from("PackTypeList")),
        Col("supplier", "Supplier", 16, "text"),
        Col("pack_qty", "Units in a pack", 12, "number1",
            validate=number_check(0, 1000000),
            note="500 stickers on a roll, 100 mailers in a box."),
        Col("pack_cost", "Pack price", 11, "money",
            validate=number_check(0, 100000),
            note="Including delivery, as always."),
        Col("qty", "Used per order", 11, "number1",
            validate=number_check(0, 10000),
            note="Usually 1. Two sheets of tissue is 2."),
        Col("weight", "Weight each (g)", 12, "number1",
            validate=number_check(0, 100000),
            note="Packaging weight pushes orders into the next "
                 "postage band more often than the product does."),
        Col("budget", "Budget", 9, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("standard", "Standard", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("gift", "Gift", 9, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("premium", "Premium", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("reusable", "Branded?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Branded packaging is marketing as well as cost - "
                 "worth knowing how much of the spend it is."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("unit_cost", "Cost each", 10, "money4",
            formula=blank(
                "name",
                'IF(OR({pack_cost}{r}="",{pack_qty}{r}="",'
                '{pack_qty}{r}=0),"",ROUND({pack_cost}{r}/'
                '{pack_qty}{r},4))')),
        Col("line_cost", "Cost per order", 12, "money4",
            formula=blank(
                "name",
                'IF(OR({unit_cost}{r}="",{qty}{r}=""),"",'
                'ROUND({unit_cost}{r}*{qty}{r},4))')),
        Col("line_weight", "Weight per order", 13, "number1",
            formula=blank(
                "name",
                'IF(OR({weight}{r}="",{qty}{r}=""),"",'
                'ROUND({weight}{r}*{qty}{r},1))')),
        Col("c_budget", "Budget cost", 10, "money4",
            formula=blank(
                "name",
                'IF(AND({budget}{r}="Yes",{line_cost}{r}<>""),'
                '{line_cost}{r},0)'),
            hidden=True),
        Col("c_standard", "Standard cost", 10, "money4",
            formula=blank(
                "name",
                'IF(AND({standard}{r}="Yes",{line_cost}{r}<>""),'
                '{line_cost}{r},0)'),
            hidden=True),
        Col("c_gift", "Gift cost", 10, "money4",
            formula=blank(
                "name",
                'IF(AND({gift}{r}="Yes",{line_cost}{r}<>""),'
                '{line_cost}{r},0)'),
            hidden=True),
        Col("c_premium", "Premium cost", 10, "money4",
            formula=blank(
                "name",
                'IF(AND({premium}{r}="Yes",{line_cost}{r}<>""),'
                '{line_cost}{r},0)'),
            hidden=True),
        Col("w_budget", "Budget weight", 10, "number1",
            formula=blank(
                "name",
                'IF(AND({budget}{r}="Yes",{line_weight}{r}<>""),'
                '{line_weight}{r},0)'),
            hidden=True),
        Col("w_standard", "Standard weight", 10, "number1",
            formula=blank(
                "name",
                'IF(AND({standard}{r}="Yes",{line_weight}{r}<>""),'
                '{line_weight}{r},0)'),
            hidden=True),
        Col("w_gift", "Gift weight", 10, "number1",
            formula=blank(
                "name",
                'IF(AND({gift}{r}="Yes",{line_weight}{r}<>""),'
                '{line_weight}{r},0)'),
            hidden=True),
        Col("w_premium", "Premium weight", 10, "number1",
            formula=blank(
                "name",
                'IF(AND({premium}{r}="Yes",{line_weight}{r}<>""),'
                '{line_weight}{r},0)'),
            hidden=True),
        Col("profiles", "In profiles", 11, "int",
            formula=blank(
                "name",
                'COUNTIF({budget}{r}:{premium}{r},"Yes")'),
            note="How many of the four profiles include this "
                 "component."),
        Col("check", "Check", 24, "status",
            formula=blank(
                "name",
                'IF(OR({pack_cost}{r}="",{pack_qty}{r}=""),'
                '"Add - price and pack size please",'
                'IF({profiles}{r}=0,"Check - in no profile",'
                'IF({qty}{r}="","Add - how many per order?",'
                'IF({weight}{r}="","Check - no weight, postage may be '
                'wrong","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Packaging Costs",
        subtitle="Four packaging profiles, built from the components "
                 "you actually buy.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).packaging,
        help_text="Tick the profiles each component belongs to. The "
                  "four totals at the bottom are what products "
                  "inherit when they pick a profile.",
        extra_notes=[
            "Branded packaging is a marketing decision with a price "
            "tag. Know the price tag before you decide.",
            "Weight matters as much as cost: 40 g of box and tissue "
            "can push a 180 g order into the next postage band and "
            "cost you more than the tissue did.",
            "A gift profile that costs $3.40 on a $12 product is "
            "eating a quarter of the price. Either charge for gift "
            "wrap or keep it for the pieces that can carry it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Profile totals, published as defined names."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    book = sheet.book
    ref = quote_sheet(NAME)

    row = sheet.section(row, "What each packaging profile costs", last)
    row = sheet.note(
        row,
        "These four totals are the figures the product catalog, both "
        "price calculators and the order forms read. A product picks a "
        "profile by name and inherits the cost and the weight.",
        last, height=30)
    row += 1

    row = sheet.table_head(row, 0, ["Profile", "Components", "Cost",
                                    "Weight (g)", "What it is for"],
                           [16, 13, 13, 13, 58])
    for key, label, cost_name, weight_name, blurb in PROFILES:
        sheet.write(row, 0, label, fmts.chrome("label_bold"))
        sheet.calc_row(row, 1, [
            f'=COUNTIFS([[Packaging.{key}]],"Yes")',
            f'=ROUND(SUM([[Packaging.c_{key}]]),2)',
            f'=ROUND(SUM([[Packaging.w_{key}]]),1)',
        ], ["int", "money", "number1"])
        sheet.write(row, 4, blurb, fmts.chrome("note"))
        book.define_name(cost_name, f"={ref}!$C${row + 1}")
        book.define_name(weight_name, f"={ref}!$D${row + 1}")
        row += 1

    row += 1
    sheet.note(
        row,
        "No profile that suits a product? Add a component, tick the "
        "profile it belongs to, and these totals move - along with "
        "every price in the workbook that uses them.",
        last, height=26)
