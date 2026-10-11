"""`Leftovers` - the half a lasagne in the fridge.

Log what is left after a meal, where you put it and how long it is good
for. The dashboard nags you about it, the planner lets you drop it into
tomorrow's lunch slot, and anything that does get thrown away rolls
through to the Food Waste tab with a price on it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, lookup_recipe, status_cf

NAME = "Leftovers"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("item", "LeftoverPrefix", digits=4, width=10),
        Col("item", "What it is", 26, "text",
            note="'Half the chilli', 'roast chicken - two portions'."),
        Col("recipe", "From which recipe", 11, "text",
            validate=list_from("RecipeIDs"),
            note="Optional. Links it back so the cost per portion is "
                 "worked out for you."),
        Col("recipe_name", "Recipe", 22, "text",
            formula=lookup_recipe("name"), hidden=True),
        Col("cooked", "Cooked on", 11, "date", validate=date_check()),
        Col("portions", "Portions left", 10, "number1",
            validate=number_check(0, 100)),
        Col("location", "Stored in", 12, "text",
            validate=list_from("StorageList"),
            note="Fridge or freezer - it changes how long you have."),
        Col("container", "Container", 14, "text",
            note="'Big glass box', 'blue lid' - so you can find it."),
        Col("window", "Days it keeps", 10, "int",
            formula='=IF({location}{r}="","",IF({location}{r}='
                    '"Freezer",FreezerDays,LeftoverDays))',
            hidden=True),
        Col("use_by", "Eat it by", 11, "date",
            formula='=IF(OR({cooked}{r}="",{window}{r}=""),"",'
                    '{cooked}{r}+{window}{r})',
            note="Cooked date plus the fridge or freezer window from "
                 "the Settings tab."),
        Col("days_left", "Days left", 9, "int",
            formula='=IF({use_by}{r}="","",{use_by}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 1, "fmt": "warn"}]),
        Col("cost_portion", "Cost a portion", 11, "money",
            formula=lookup_recipe("cost_serving"), hidden=True),
        Col("value", "Worth", 11, "money",
            formula='=IF(OR({portions}{r}="",{cost_portion}{r}="",'
                    'NOT(ISNUMBER({cost_portion}{r}))),"",'
                    '{portions}{r}*{cost_portion}{r})',
            note="What you would pay to make it again - the number "
                 "that makes leftovers feel worth eating."),
        Col("planned_for", "Planned for", 11, "date",
            validate=date_check(),
            note="When you intend to eat it. Add the meal to the plan "
                 "as well and it shows up in the week."),
        Col("eaten_by", "Who is having it", 14, "text",
            validate=list_from("MemberNames")),
        Col("queue_rank", "Queue rank", 9, "int",
            formula='=IF(OR({item}{r}="",{status}{r}="Eaten",'
                    '{status}{r}="Binned",NOT(ISNUMBER('
                    '{days_left}{r}))),"",COUNTIFS({status_first}:'
                    '{status_last},"<>Eaten",{status_first}:'
                    '{status_last},"<>Binned",{days_left_first}:'
                    '{days_left_last},"<"&{days_left}{r})+'
                    'COUNTIFS({status_first}:{status}{r},"<>Eaten",'
                    '{status_first}:{status}{r},"<>Binned",'
                    '{days_left_first}:{days_left}{r},{days_left}{r}))',
            hidden=True,
            note="1 = eat this one next. Drives the Fridge & Freezer "
                 "queue and the printable week."),
        Col("status", "Status", 13, "status",
            validate=list_from("LeftoverStatusList"), cf=status_cf()),
        Col("outcome", "What happened", 26, "status",
            formula='=IF({item}{r}="","",'
                    'IF({status}{r}="Eaten","OK - eaten",'
                    'IF({status}{r}="Binned","Wasted - log the cost",'
                    'IF({status}{r}="Frozen","OK - in the freezer",'
                    'IF(AND({days_left}{r}<>"",{days_left}{r}<0),'
                    '"Expired - check before eating",'
                    'IF(AND({days_left}{r}<>"",{days_left}{r}<=1),'
                    '"Today - eat it now","OK"))))))',
            cf=alert_cf()),
        Col("notes", "Notes", 22, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Leftovers Tracker",
        subtitle="What is in the fridge from last night, what it is "
                 "worth and how long you have to eat it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_soft,
        freeze_cols=2,
        samples=sample(cfg).leftovers,
        help_text="Log it while you are putting it in the fridge - it "
                  "takes ten seconds and turns into a free lunch.",
        extra_notes=[
            "The fridge and freezer windows come from the Settings "
            "tab (fridge "
            f"{cfg.leftover_safe_days} days, freezer "
            f"{cfg.freezer_review_days} "
            "days by default). They are guidance, not a guarantee - "
            "look, smell and use your judgement.",
            "Mark a portion 'Eaten' and it counts towards the money "
            "you saved; mark it 'Binned' and it shows up as waste.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Leftovers scoreboard", last)
    row = s.table_head(row, 0, ["Portions in the house", "Worth",
                                "Eat today or tomorrow", "Past their date",
                                "Eaten so far", "Binned so far",
                                "Money saved by eating them"],
                       [18, 12, 19, 16, 13, 13, 24])
    row = s.calc_row(row, 0, [
        '=SUMIFS([[Leftovers.portions]],[[Leftovers.status]],'
        '"<>Eaten",[[Leftovers.status]],"<>Binned")',
        '=SUMIFS([[Leftovers.value]],[[Leftovers.status]],"<>Eaten",'
        '[[Leftovers.status]],"<>Binned")',
        '=COUNTIFS([[Leftovers.days_left]],">=0",'
        '[[Leftovers.days_left]],"<=1")',
        '=COUNTIFS([[Leftovers.days_left]],"<0",'
        '[[Leftovers.status]],"<>Eaten")',
        '=COUNTIF([[Leftovers.status]],"Eaten")',
        '=COUNTIF([[Leftovers.status]],"Binned")',
        '=SUMIFS([[Leftovers.value]],[[Leftovers.status]],"Eaten")',
    ], ["number1", "money", "int", "int", "int", "int", "money"])
