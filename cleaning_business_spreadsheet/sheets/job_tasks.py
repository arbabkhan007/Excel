"""`Job Tasks` - the checklist as it was actually done, job by job.

Copy the tasks for the service onto a job, tick them off as you go,
and the job's completion percentage, the quality score and the list of
issues found all fill themselves in. This is the tab that turns "it was
cleaned" into something you can show a client.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_job,
                     status_cf)

NAME = "Job Tasks"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("job", '"JT-"', digits=4, width=10),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Which visit. Everything about the job is pulled "
                 "through from the Jobs tab."),
        Col("date", "Date", 11, "date",
            formula=lookup_job("date", missing="")),
        Col("client_name", "Client", 20, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))'),
        Col("property_name", "Property", 22, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.property_name]],MATCH({job}{r},[[Jobs.id]],'
                    '0)),""))'),
        Col("area", "Area", 16, "text", validate=list_from("TaskAreaList")),
        Col("task", "Task", 38, "text",
            note="Copy these from the Checklists tab for the service "
                 "you are doing - filter, copy, paste."),
        Col("status", "Status", 13, "status",
            validate=list_from("TaskStatusList"), cf=status_cf(),
            note="Done, Skipped, Issue found. The job's completion "
                 "percentage counts the Done ones."),
        Col("employee", "Done by", 16, "text",
            validate=list_from("EmployeeNames"),
            note="Blank means the lead cleaner on the job."),
        Col("minutes", "Minutes taken", 10, "int",
            validate=number_check(0, 600, decimal=False)),
        Col("quality", "Quality", 13, "status",
            validate=list_from("QualityList"), cf=status_cf(),
            note="Filled in on a spot check, not on every task. Only "
                 "the ones you look at."),
        Col("issue", "Issue found", 32, "wrap",
            note="A dripping tap, a stain that will not shift, a "
                 "light bulb gone. Tell the client before they find "
                 "it."),
        Col("reported", "Told the client?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("photo", "Photo link", 20, "link",
            note="A spreadsheet cannot hold the photo itself tidily - "
                 "paste a cloud link."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("required", "Required?", 9, "flag",
            formula='=IF({task}{r}="","",IFERROR(INDEX('
                    '[[Checklists.required]],MATCH({task}{r},'
                    '[[Checklists.task]],0)),"Yes"))',
            hidden=True),
        Col("planned", "Planned minutes", 10, "int",
            formula='=IF({task}{r}="","",IFERROR(INDEX('
                    '[[Checklists.minutes]],MATCH({task}{r},'
                    '[[Checklists.task]],0)),""))',
            hidden=True),
        Col("over", "Over / under (min)", 10, "int",
            formula='=IF(OR({minutes}{r}="",{planned}{r}=""),"",'
                    '{minutes}{r}-{planned}{r})',
            hidden=True),
        Col("done", "Complete?", 9, "flag",
            formula=blank("job",
                          'IF({status}{r}="Done","Yes","No")'),
            cf=flag_yes_no_cf(), hidden=True),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "job",
                'IF({task}{r}="","Add - what was the task?",'
                'IF({status}{r}="Issue found",IF({reported}{r}="Yes",'
                '"OK - issue reported","Action - tell the client"),'
                'IF({status}{r}="Needs rework","Action - rework needed",'
                'IF(AND({status}{r}="Skipped",{required}{r}="Yes"),'
                '"Check - required task skipped",'
                'IF({status}{r}="Done","OK","Due")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Job Tasks & Completion",
        subtitle="The checklist as it was really done - ticked off, "
                 "timed and with the problems written down.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).job_tasks,
        help_text="Filter the Checklists tab to the service you are "
                  "doing, copy the tasks, paste them here against the "
                  "job ID, and tick them off. The job's completion "
                  "percentage updates as you go.",
        extra_notes=[
            "You do not have to use this tab for every visit. Most "
            "businesses use it for deep cleans, end of tenancy work "
            "and anywhere a client has complained before.",
            "'Issue found' plus a photo link is the cheapest "
            "insurance there is. It proves the damage was already "
            "there.",
            "Minutes taken against planned minutes is how you find "
            "out that the oven really takes 90 minutes, not 60.",
        ],
    )
