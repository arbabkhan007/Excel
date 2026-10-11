"""`Workflow` - the production pipeline: what is where, and what is late."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Workflow"
LAST_COL = 15
FIRST_CALC_ROW = 5

STAGES = [
    ("1. Idea", "Idea",
     "Captured in the Idea Bank and scored. Nothing is written yet."),
    ("2. Brief", "Brief",
     "Hook, message, format and call to action agreed. Owner set."),
    ("3. In production", "In production",
     "Shooting, filming, writing or designing."),
    ("4. Draft", "Draft",
     "Caption written, asset ready, hashtag group chosen."),
    ("5. Awaiting review", "Awaiting review",
     "With whoever approves it. Check the facts, links and claims."),
    ("6. Revision needed", "Revision needed",
     "Sent back with notes. Fix and resubmit."),
    ("7. Approved", "Approved",
     "Signed off and ready to be loaded into the scheduler."),
    ("8. Scheduled", "Scheduled",
     "In the scheduler, or an alarm set to post manually."),
    ("9. Published", "Published",
     "Live. Add the link, then the numbers after 48-72 hours."),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Content Production Workflow",
             "Every post moves through the same nine stages. This page "
             "shows how many are sitting at each one, what is late, and "
             "who is holding it.",
             LAST_COL)
    s.set_col(0, 0, 20)
    s.set_col(1, 1, 13)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.section(row, "Where everything is right now", LAST_COL)
    s.table_head(row, 0, ["Stage", "All posts", "This period",
                          "What happens at this stage"])
    first_stage = row + 1
    for i, (label, status, description) in enumerate(STAGES):
        r = first_stage + i
        s.calc_row(r, 0, [
            label,
            f'=COUNTIF([[Content Calendar.status]],"{status}")',
            f'=COUNTIFS([[Content Calendar.status]],"{status}",'
            f'[[Content Calendar.date]],">="&DashStart,'
            f'[[Content Calendar.date]],"<="&DashEnd)',
        ], ["text", "int", "int"])
        s.merge(r, 3, LAST_COL, description, fmts.cell("wrap", "calc"))
    ws.conditional_format(first_stage, 1, first_stage + len(STAGES) - 1, 1,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    row = first_stage + len(STAGES) + 1

    # ------------------------------------------------------------- health
    row = s.section(row, "Pipeline health", LAST_COL)
    s.kpi(row, 0, "In the pipeline",
          '=COUNTIFS([[Content Calendar.title]],"<>",'
          '[[Content Calendar.status]],"<>Published",'
          '[[Content Calendar.status]],"<>Cancelled",'
          '[[Content Calendar.status]],"<>Archived")', "int", 3,
          "Everything not yet live")
    s.kpi(row, 3, "Drafts late",
          '=COUNTIF([[Content Calendar.alert]],"Late*")', "int", 3,
          "Past the draft deadline")
    s.kpi(row, 6, "Needing approval",
          '=COUNTIF([[Content Calendar.status]],"Awaiting review")',
          "int", 3, "Waiting on a yes")
    s.kpi(row, 9, "Due in 7 days",
          '=COUNTIFS([[Content Calendar.date]],">="&TODAY(),'
          '[[Content Calendar.date]],"<="&TODAY()+7,'
          '[[Content Calendar.status]],"<>Published")', "int", 3,
          "Coming up this week")
    s.kpi(row, 12, "Overdue",
          '=COUNTIF([[Content Calendar.alert]],"Overdue*")', "int", 3,
          "Date passed, not published")
    row += 4

    s.kpi(row, 0, "Weeks of content ready",
          '=IFERROR((COUNTIFS([[Content Calendar.status]],"Approved")'
          '+COUNTIFS([[Content Calendar.status]],"Scheduled"))'
          '/MAX(1,PostsPerWeek),0)', "number1", 3,
          "Approved plus scheduled, divided by your weekly target")
    s.kpi(row, 3, "Captions missing",
          '=COUNTIFS([[Content Calendar.title]],"<>",'
          '[[Content Calendar.caption]],"")', "int", 3,
          "Posts with no caption picked")
    s.kpi(row, 6, "Assets missing",
          '=COUNTIFS([[Content Calendar.title]],"<>",'
          '[[Content Calendar.asset]],"")', "int", 3,
          "Posts with no creative attached")
    s.kpi(row, 9, "Hashtag groups missing",
          '=COUNTIFS([[Content Calendar.title]],"<>",'
          '[[Content Calendar.hashtag_group]],"")', "int", 3,
          "Posts with no hashtag group")
    s.kpi(row, 12, "Open tasks",
          '=COUNTIFS([[Tasks.task]],"<>",[[Tasks.status]],"<>Complete")',
          "int", 3, "On the Tasks tab")
    row += 4

    # -------------------------------------------------------- the queues
    row = s.section(row, "The queue", LAST_COL)
    s.merge(row, 0, 6, "Next 12 posts due", fmts.chrome("subsection"))
    s.merge(row, 8, LAST_COL, "Overdue - deal with these first",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["Date", "Platform", "Post", "Status", "Owner",
                          "Days"])
    s.table_head(row, 8, ["Date", "Platform", "Post", "Status", "Owner",
                          "Days late"])
    first_queue = row + 1
    for i in range(12):
        r = first_queue + i
        n = i + 1
        up = (f"MATCH(SMALL([[Content Calendar.queue_key]],{n}),"
              f"[[Content Calendar.queue_key]],0)")
        late = (f"MATCH(SMALL([[Content Calendar.late_key]],{n}),"
                f"[[Content Calendar.late_key]],0)")
        s.calc_row(r, 0, [
            f'=IFERROR(INDEX([[Content Calendar.date]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.platform]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.title]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.status]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.owner]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.days_to_go]],{up}),"")',
        ], ["date", "text", "text", "status", "text", "int"])
        s.calc_row(r, 8, [
            f'=IFERROR(INDEX([[Content Calendar.date]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.platform]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.title]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.status]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.owner]],{late}),"")',
            f'=IFERROR(-INDEX([[Content Calendar.days_to_go]],{late}),"")',
        ], ["date", "text", "text", "status", "text", "int"])
    row = first_queue + 13

    # ------------------------------------------------------------- owners
    row = s.section(row, "Who is holding what", LAST_COL)
    s.table_head(row, 0, ["Owner", "Posts owned", "Still to publish",
                          "Open tasks", "Late posts", "On-time rate"])
    first_owner = row + 1
    owner_rows = min(10, cfg.row_count("Team"))
    for i in range(owner_rows):
        r = first_owner + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF([[Calc.ow_name#{cr}]]="","",'
        s.calc_row(r, 0, [
            f"={guard}[[Calc.ow_name#{cr}]])",
            f"={guard}[[Calc.ow_posts#{cr}]])",
            f"={guard}[[Calc.ow_open#{cr}]])",
            f"={guard}[[Calc.ow_tasks#{cr}]])",
            f"={guard}[[Team.late#{cr}]])",
            f"={guard}[[Team.on_time#{cr}]])",
        ], ["text", "int", "int", "int", "int", "percent"])
    ws.conditional_format(first_owner, 1, first_owner + owner_rows - 1, 1,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    ws.conditional_format(first_owner, 4, first_owner + owner_rows - 1, 4, {
        "type": "cell", "criteria": ">", "value": 0,
        "format": fmts.cf("bad")})
    row = first_owner + owner_rows + 1

    # ---------------------------------------------------------- checklist
    row = s.section(row, "The weekly routine that keeps this working",
                    LAST_COL)
    row = s.bullets(row, [
        "MONDAY - open the Dashboard, clear anything in the overdue list, "
        "then look at the next 12 posts due.",
        "MONDAY - batch the week's captions in the Caption Bank and pick "
        "a hashtag group for each post.",
        "TUESDAY or WEDNESDAY - make the content: shoot, film, design. "
        "Mark each row 'Draft' as it is finished.",
        "THURSDAY - review and approve. Check links, claims, spelling and "
        "the Compliance tab.",
        "FRIDAY - schedule everything, set the status to 'Scheduled' and "
        "add the publish links as they go live.",
        "ANY DAY - add ideas to the Idea Bank as they come to you. Future "
        "you will be grateful.",
        "FIRST OF THE MONTH - fill in the Analytics tab, update the Goals "
        "tab, and read the Monthly Dashboard.",
    ], LAST_COL)
    row += 1

    s.note(row,
           "The two deadline columns on the Content Calendar (draft due "
           "and approval due) are worked out from the publish date using "
           "the production and approval days on the Settings tab - change "
           "those two numbers to match how you really work.",
           LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
