"""`Crochet` - yarn by the gram, because that is how it is used.

Nobody uses a whole skein. This tab costs a crochet or knitting piece
the way it is actually made: so many grams of this yarn, a few of
that, a couple of buttons, the pattern you bought spread over the
pieces you will make from it, and the hours it really takes.

Quantities are in the material's base unit - grams for yarn - and the
unit table on the Lists tab converts whatever the skein was sold in.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Crochet"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')
_PACK = ('IFERROR(INDEX([[Materials.pack_qty]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')
_FACTOR = ('IFERROR(INDEX([[Materials.base_factor]],MATCH({key}{r},'
           '[[Materials.id]],0)),1)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs"),
            note="Link it to the catalog and the costs below flow "
                 "straight into the product's price. Leave it blank "
                 "to cost a piece you have not listed yet."),
        Col("name", "Piece", 28, "text",
            note="What you are making: 'Cable beanie, adult', "
                 "'Granny square blanket, throw'."),
        Col("pattern", "Pattern", 22, "text"),
        Col("size", "Size", 14, "text", validate=list_from("SizeNames"),
            note="Stretches both the yarn and the time by the size "
                 "factors on the Lists tab."),
        Col("yarn_weight", "Yarn weight", 16, "text",
            validate=list_from("YarnWeightList")),
        Col("hook", "Hook", 10, "text", validate=list_from("HookSizeList")),
        Col("yarn", "Main yarn", 12, "id", validate=list_from("MaterialIDs"),
            note="The material ID from the Materials tab."),
        Col("grams", "Grams used", 11, "number1",
            validate=number_check(0, 100000),
            note="Weigh the finished piece, or weigh the skein "
                 "before and after. This single number is the "
                 "difference between a guess and a price."),
        Col("yarn2", "Second yarn", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("grams2", "Grams used", 11, "number1",
            validate=number_check(0, 100000)),
        Col("waste", "Yarn waste", 10, "percent",
            validate=number_check(0, 1),
            note="Tails, swatches, the bit you frogged. Blank uses "
                 "the default on Settings."),
        Col("notions", "Buttons, zips, eyes", 13, "money",
            validate=number_check(0, 10000)),
        Col("stuffing", "Stuffing / lining", 12, "money",
            validate=number_check(0, 10000)),
        Col("pattern_cost", "Pattern cost", 11, "money",
            validate=number_check(0, 10000),
            note="What you paid for the pattern or the licence."),
        Col("pattern_uses", "Pieces from it", 11, "int",
            validate=number_check(1, 100000, decimal=False),
            note="How many pieces you expect to make from that "
                 "pattern. The cost is spread over them."),
        Col("hours_input", "Hours hooking", 11, "number",
            validate=number_check(0, 1000),
            note="Actual hook time. Most makers are surprised the "
                 "first time they measure it honestly."),
        Col("finish_mins", "Finishing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Weaving in ends, blocking, sewing up, buttons. "
                 "The part everybody forgets to charge for."),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("size_mat", "Size factor", 9, "number",
            formula=blank(
                "name",
                'IFERROR(INDEX(SizeMaterialFactors,MATCH({size}{r},'
                'SizeNames,0)),1)'),
            hidden=True),
        Col("size_time", "Time factor", 9, "number",
            formula=blank(
                "name",
                'IFERROR(INDEX(SizeTimeFactors,MATCH({size}{r},'
                'SizeNames,0)),1)'),
            hidden=True),
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "name", 'IF({waste}{r}="",DefaultWaste,{waste}{r})'),
            hidden=True),
        Col("grams_total", "Grams, with waste", 12, "number1",
            formula=blank(
                "name",
                'ROUND((IF({grams}{r}="",0,{grams}{r})+'
                'IF({grams2}{r}="",0,{grams2}{r}))*{size_mat}{r}*'
                '(1+{waste_rate}{r}),1)'),
            note="Everything the piece really consumes, size and "
                 "waste included."),
        Col("yarn_cost", "Main yarn cost", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({grams}{r}="",0,{grams}{r})*{size_mat}{r}*'
                '(1+{waste_rate}{r})*' + _BASE.replace("{key}", "{yarn}") + ',2)')),
        Col("yarn2_cost", "Second yarn cost", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({grams2}{r}="",0,{grams2}{r})*{size_mat}{r}*'
                '(1+{waste_rate}{r})*' + _BASE.replace("{key}", "{yarn2}") +
                ',2)')),
        Col("skeins", "Skeins needed", 11, "number1",
            formula=blank(
                "name",
                'IF(' + _PACK.replace("{key}", "{yarn}") + '=0,"",'
                'ROUNDUP(IF({grams}{r}="",0,{grams}{r})*{size_mat}{r}*'
                '(1+{waste_rate}{r})/(' + _PACK.replace("{key}", "{yarn}") + '*' +
                _FACTOR.replace("{key}", "{yarn}") + '),0))'),
            note="How many balls you have to buy to make one. Buy "
                 "whole skeins, cost the grams - but know the "
                 "difference, because the leftovers are real money."),
        Col("leftover", "Grams left over", 12, "number1",
            formula=blank(
                "name",
                'IF({skeins}{r}="","",ROUND({skeins}{r}*' +
                _PACK.replace("{key}", "{yarn}") + '*' +
                _FACTOR.replace("{key}", "{yarn}") + '-IF({grams}{r}="",0,'
                '{grams}{r})*{size_mat}{r}*(1+{waste_rate}{r}),1))'),
            hidden=True),
        Col("pattern_share", "Pattern per piece", 12, "money",
            formula=blank(
                "name",
                'IF({pattern_cost}{r}="",0,ROUND({pattern_cost}{r}/'
                'MAX(1,IF({pattern_uses}{r}="",1,{pattern_uses}{r})),2))'),
            hidden=True),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({yarn_cost}{r}+{yarn2_cost}{r}+'
                'IF({notions}{r}="",0,{notions}{r})+'
                'IF({stuffing}{r}="",0,{stuffing}{r})+'
                '{pattern_share}{r},2)'),
            note="Yarn, notions, stuffing and the pattern's share. "
                 "This is what flows into the catalog."),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'ROUND((IF({hours_input}{r}="",0,{hours_input}{r})*60+'
                'IF({finish_mins}{r}="",0,{finish_mins}{r}))*'
                '{size_time}{r},1)')),
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank(
                "name",
                'ROUND({minutes}{r}/60*{rate}{r},2)'),
            note="Minutes / 60 x your hourly rate. A blanket that "
                 "takes 22 hours is not a $45 blanket."),
        Col("other_cost", "Extras", 10, "money",
            formula=blank(
                "name",
                'ROUND({minutes}{r}/60*SuppliesPerHour,2)'),
            hidden=True),
        Col("cost_per_gram", "Cost a gram", 11, "money4",
            formula=blank(
                "name",
                'IF(OR({grams_total}{r}="",{grams_total}{r}=0),"",'
                'ROUND({mat_cost}{r}/{grams_total}{r},4))'),
            hidden=True),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Crochet & Knitting Calculator",
        subtitle="Yarn by the gram, notions, the pattern's share and "
                 "the hours it really takes.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).crochet,
        help_text="Weigh the piece, enter the grams, and the yarn "
                  "costs itself. Link the row to a product ID and "
                  "the catalog picks the whole lot up.",
        extra_notes=[
            "Weigh the finished piece on kitchen scales. A 94 g "
            "beanie from a $9.40 100 g skein used $8.84 of yarn - "
            "no estimate gets that close.",
            "Hours are the real cost of crochet. At $25 an hour a "
            "six hour shawl carries $150 of labour before a single "
            "gram of yarn, which is why handmade crochet cannot "
            "compete with machine made and should never try.",
            "Use this tab OR the bill of materials for a given "
            "product, not both - the catalog adds them together.",
        ],
    )
