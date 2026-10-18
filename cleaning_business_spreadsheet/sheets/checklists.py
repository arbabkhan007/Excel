"""`Checklists` - the master task lists behind every clean.

Build one checklist per service (and, where it matters, per property)
and you have a standard. The cleaner works through it on the Job Tasks
tab; the minutes here add up to a realistic time for the visit, which
is what turns a guess into a price.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Checklists"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("task", '"CHK-"', digits=4, width=10),
        Col("template", "Checklist for", 24, "text",
            validate=list_from("ServiceNames"),
            note="Which service this task belongs to. The job sheet "
                 "prints every task for the service you booked."),
        Col("property", "Only at this property", 13, "id",
            validate=list_from("PropertyIDs"),
            note="Leave blank for a task that applies everywhere. "
                 "Fill it in for the one client who wants the bins "
                 "put out."),
        Col("area", "Area", 18, "text", validate=list_from("TaskAreaList"),
            note="Group the tasks by room so the cleaner works "
                 "through the property in a sensible order."),
        Col("order", "Order", 7, "int",
            validate=number_check(1, 999, decimal=False),
            note="Sort on Area then Order to print the checklist in "
                 "the order you want it done."),
        Col("task", "Task", 40, "text",
            note="One job, written as an instruction: 'Descale the "
                 "shower head', not 'bathroom'."),
        Col("detail", "How it should look", 36, "wrap",
            note="The standard. This is what settles an argument "
                 "about whether something was done properly."),
        Col("minutes", "Minutes", 9, "int",
            validate=number_check(0, 600, decimal=False),
            note="Honest minutes for one cleaner. The total at the "
                 "bottom of this tab is how long the service really "
                 "takes."),
        Col("frequency", "How often", 14, "text",
            validate=list_from("FrequencyList"),
            note="Every visit, or only weekly / monthly / quarterly. "
                 "Deep tasks do not need doing every time."),
        Col("required", "Required?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Required tasks are the ones that make the job "
                 "'complete'. Extras are nice to have."),
        Col("photo", "Photo needed?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Useful for end of tenancy and changeover work - "
                 "proof is cheaper than an argument."),
        Col("products", "Products / kit", 22, "text"),
        Col("active", "In use?", 8, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("hours", "Hours", 8, "number1",
            formula=blank("task",
                          'ROUND(IF({minutes}{r}="",0,{minutes}{r})/60,2)'),
            hidden=True),
        Col("service_total", "Total for this checklist", 12, "number1",
            formula='=IF({template}{r}="","",ROUND(SUMIFS('
                    '{minutes_first}:{minutes_last},'
                    '{template_first}:{template_last},{template}{r},'
                    '{active_first}:{active_last},"Yes")/60,2))',
            note="Every active task for this service, in hours. "
                 "Compare it with the hours on the Services tab - if "
                 "they disagree, one of them is wrong.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("service_hours", "Service allows", 11, "number1",
            formula='=IF({template}{r}="","",IFERROR(INDEX('
                    '[[Services.hours]],MATCH({template}{r},'
                    '[[Services.name]],0)),""))',
            hidden=True),
        Col("times_done", "Times ticked off", 11, "int",
            formula='=IF({task}{r}="","",COUNTIFS([[Job Tasks.task]],'
                    '{task}{r},[[Job Tasks.status]],"Done"))'),
        Col("times_skipped", "Times skipped", 11, "int",
            formula='=IF({task}{r}="","",COUNTIFS([[Job Tasks.task]],'
                    '{task}{r},[[Job Tasks.status]],"Skipped"))',
            cf=[{"type": "cell", "criteria": ">=", "value": 3,
                 "fmt": "warn"}],
            note="A task skipped again and again is either not "
                 "needed or not possible in the time you allow."),
        Col("issues", "Issues found", 10, "int",
            formula='=IF({task}{r}="","",COUNTIFS([[Job Tasks.task]],'
                    '{task}{r},[[Job Tasks.status]],"Issue found"))'),
        Col("status", "Standard", 13, "status",
            formula=blank(
                "task",
                'IF({times_done}{r}+{times_skipped}{r}=0,"New",'
                'IF({times_skipped}{r}/({times_done}{r}+'
                '{times_skipped}{r})>0.25,"Often skipped","Working"))'),
            cf=status_cf(), hidden=True),
        Col("seq", "Line number", 9, "int",
            formula='=IF(OR({template}{r}="",{task}{r}=""),"",'
                    'COUNTIFS({template_first}:{template}{r},'
                    '{template}{r}))',
            hidden=True,
            note="1, 2, 3 within each checklist. The printable form "
                 "reads it."),
        Col("key", "Lookup key", 12, "text",
            formula='=IF({seq}{r}="","",{template}{r}&"|"&{seq}{r})',
            hidden=True),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "task",
                'IF({template}{r}="","Add - which service?",'
                'IF({minutes}{r}="","Add - how many minutes?",'
                'IF(AND({service_hours}{r}<>"",{service_total}{r}>'
                '{service_hours}{r}+0.25),'
                '"Check - checklist is longer than the service allows",'
                'IF({times_skipped}{r}>=3,"Check - often skipped",'
                '"OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Cleaning Checklists",
        subtitle="The standard for every service, room by room, with "
                 "honest minutes against each task.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).checklists,
        help_text="Write the checklist once per service. Add a "
                  "property ID only for the tasks that apply to one "
                  "address. The Print Job Sheet tab turns it into "
                  "something the cleaner can carry.",
        extra_notes=[
            "Adding up the minutes is the single most useful thing "
            "on this tab. Most underpriced cleans are underpriced "
            "because nobody ever counted the tasks.",
            "The 'Checklist is longer than the service allows' "
            "warning means your price and your standard disagree. "
            "Change one of them.",
            "Tasks ticked off on the Job Tasks tab feed back here, "
            "so you can see which parts of your standard actually "
            "get done.",
        ],
    )
