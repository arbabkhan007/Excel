"""`Reports & Analytics` - the questions a cleaning business should ask.

Where the money comes from, which services earn their keep, which
advertising works, how many quotes turn into work, what the team
costs and who owes you money. Every table reads the registers, so the
answers change the moment you enter a job.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Reports"
LAST_COL = 11
TOP = 8

_REV = ('SUMIFS([[Jobs.price]],{cond},[[Jobs.status]],"Completed")+'
        'SUMIFS([[Jobs.price]],{cond},[[Jobs.status]],"Invoiced")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Reports & Business Analytics",
             "Where the work comes from, what it earns and what it "
             "costs - all from the tabs you have already filled in.",
             LAST_COL)

    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.note(row,
                 "Every figure is live. Filter or sort the registers "
                 "as much as you like - these tables add up the whole "
                 "column either way.", LAST_COL, height=22)
    row += 1

    # ------------------------------------------------- by service -------
    row = s.section(row, "Which services earn their keep", LAST_COL)
    row = s.table_head(row, 0, ["Service", "Times booked", "Revenue",
                                "Hours", "Earns an hour", "Margin",
                                "Times quoted", "Conversion"],
                       [26, 12, 13, 10, 13, 10, 12, 12])
    first_svc = row
    for n in range(1, TOP + 1):
        pick = f'MATCH({n},[[Services.revenue_rank]],0)'
        cell = f"$A${row + 1}"
        for col, (key, style) in enumerate([
                ("name", "text"), ("times_booked", "int"),
                ("revenue", "money0")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Services.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF({cell}="","",SUMIFS([[Jobs.est_hours]],'
                         f'[[Jobs.service]],{cell},[[Jobs.status]],'
                         f'"Completed")+SUMIFS([[Jobs.est_hours]],'
                         f'[[Jobs.service]],{cell},[[Jobs.status]],'
                         f'"Invoiced"))'), fmts.calc("number1"))
        for col, key, style in [(4, "hourly_yield", "money"),
                                (5, "margin", "percent0"),
                                (6, "times_quoted", "int")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Services.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 7,
            f'=IF(OR({cell}="",G{row + 1}=0),"",B{row + 1}/G{row + 1})',
            fmts.calc("percent0"))
        row += 1
    ws.conditional_format(first_svc, 2, row - 1, 2,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    ws.conditional_format(first_svc, 4, row - 1, 4,
                          {"type": "3_color_scale",
                           "min_color": "#FBD9DD",
                           "mid_color": "#FFF0CE",
                           "max_color": "#D7F2E3"})
    row += 2

    # ------------------------------------------------- by client type ---
    row = s.section(row, "Residential or commercial?", LAST_COL)
    row = s.table_head(row, 0, ["Client type", "Clients", "Jobs",
                                "Revenue", "Average job",
                                "Share of revenue"],
                       [24, 11, 10, 14, 13, 15])
    types = ["Residential", "Commercial office", "Retail",
             "Restaurant & cafe", "Holiday let / Airbnb", "Landlord",
             "Letting agent", "Property manager", "Medical & dental",
             "School & nursery", "Gym & studio"]
    first_type = row
    for kind in types:
        cond = f'[[Jobs.client_type]],"{kind}"'
        s.write(row, 0, kind, fmts.chrome("label_bold"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=COUNTIFS([[Clients.type]],"{kind}")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=COUNTIFS({cond},[[Jobs.status]],'
                         f'"<>Cancelled")'), fmts.calc("int"))
        ws.write_formula(row, 3,
                         s.reg.render("=" + _REV.format(cond=cond)),
                         fmts.calc("money0"))
        ws.write_formula(
            row, 4, f'=IF(C{row + 1}=0,"",D{row + 1}/C{row + 1})',
            fmts.calc("money"))
        ws.write_formula(
            row, 5,
            s.reg.render('=IFERROR(D{r}/('
                         + _REV.format(cond='[[Jobs.client]],"<>"')
                         + '),"")').replace("{r}", str(row + 1)),
            fmts.calc("percent0"))
        row += 1
    ws.conditional_format(first_type, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    row += 2

    # ---------------------------------------------- marketing ----------
    row = s.section(row, "Which advertising actually works", LAST_COL)
    row = s.table_head(row, 0, ["Came from", "Enquiries", "Won", "Lost",
                                "Win rate", "Clients now",
                                "Revenue from them"],
                       [24, 11, 9, 9, 10, 12, 16])
    sources = ["Google search", "Google Business Profile", "Facebook",
               "Instagram", "Nextdoor", "Yelp", "Directory listing",
               "Referral - client", "Referral - staff",
               "Flyer / leaflet", "Van signage", "Word of mouth",
               "Letting agent"]
    first_src = row
    for src in sources:
        s.write(row, 0, src, fmts.chrome("label_bold"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=COUNTIFS([[Leads.source]],"{src}")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=COUNTIFS([[Leads.source]],"{src}",'
                         f'[[Leads.stage]],"Won")'), fmts.calc("int"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=COUNTIFS([[Leads.source]],"{src}",'
                         f'[[Leads.stage]],"Lost")'), fmts.calc("int"))
        ws.write_formula(
            row, 4,
            f'=IF(C{row + 1}+D{row + 1}=0,"",C{row + 1}/'
            f'(C{row + 1}+D{row + 1}))', fmts.calc("percent0"))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=COUNTIFS([[Clients.source]],"{src}")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=SUMIFS([[Clients.invoiced]],'
                         f'[[Clients.source]],"{src}")'),
            fmts.calc("money0"))
        row += 1
    ws.conditional_format(first_src, 6, row - 1, 6,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    ws.conditional_format(first_src, 4, row - 1, 4,
                          {"type": "cell", "criteria": ">=", "value": 0.5,
                           "format": fmts.cf("good")})
    row += 1
    s.merge(row, 0, LAST_COL,
            "Referrals almost always win this table. If an advert has "
            "produced ten enquiries and no clients, it is not an "
            "advert, it is a subscription.", fmts.chrome("note"))
    row += 2

    # ---------------------------------------------- quotes -------------
    row = s.section(row, "Quoting and converting", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "Value", "What it means"],
                       [26, 14, 58])
    quote_rows = [
        ("Quotes sent", '=COUNTIF([[Quotes.won]],"Yes")+'
                        'COUNTIF([[Quotes.won]],"No")+'
                        'COUNTIF([[Quotes.status]],"Sent")+'
                        'COUNTIF([[Quotes.status]],"Followed up")',
         "int", '="Everything that left your desk with a price on it"'),
        ("Quotes won", '=COUNTIF([[Quotes.won]],"Yes")', "int",
         '="Accepted by the client"'),
        ("Win rate",
         '=IFERROR(COUNTIF([[Quotes.won]],"Yes")/(COUNTIF('
         '[[Quotes.won]],"Yes")+COUNTIF([[Quotes.won]],"No")),"")',
         "percent0",
         '="Target "&TEXT(QuoteWinTarget,"0%")&". Winning every quote '
         'means you are too cheap"'),
        ("Average quote",
         '=IFERROR(AVERAGEIFS([[Quotes.total]],[[Quotes.total]],">0"),'
         '"")', "money", '="Across every quote you have raised"'),
        ("Average won quote",
         '=IFERROR(AVERAGEIFS([[Quotes.total]],[[Quotes.won]],"Yes"),'
         '"")', "money",
         '="Compare it with the line above - if won quotes are much '
         'smaller, you are losing the big ones"'),
        ("Value still out there",
         '=SUMIFS([[Quotes.total]],[[Quotes.status]],"Sent")+'
         'SUMIFS([[Quotes.total]],[[Quotes.status]],"Followed up")',
         "money0", '="Quotes with no answer yet"'),
        ("Average days to answer",
         '=IFERROR(AVERAGEIFS([[Quotes.days_waiting]],'
         '[[Quotes.days_waiting]],">0"),"")', "number1",
         '="How long clients take to come back to you"'),
        ("Expired, never chased",
         '=COUNTIF([[Quotes.expired]],"Yes")', "int",
         '="Quotes that ran out of time. Every one is money you '
         'already did the work for"'),
    ]
    for label, formula, style, meaning in quote_rows:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula), fmts.calc(style))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(meaning), fmts.calc("text"))
        row += 1
    row += 1

    # ---------------------------------------------- operations ---------
    row = s.section(row, "How the work goes", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "Value", "What it means"],
                       [26, 14, 58])
    ops_rows = [
        ("Jobs completed",
         '=COUNTIF([[Jobs.status]],"Completed")+'
         'COUNTIF([[Jobs.status]],"Invoiced")', "int",
         '="Visits that happened"'),
        ("Cancelled",
         '=COUNTIF([[Jobs.status]],"Cancelled")', "int",
         '="Every cancellation is an empty slot you could have sold"'),
        ("No access",
         '=COUNTIF([[Jobs.status]],"No access")', "int",
         '="Turned up, could not get in. Charge for these or they '
         'repeat"'),
        ("Average rating",
         '=IFERROR(AVERAGEIFS([[Jobs.rating]],[[Jobs.rating]],">0"),"")',
         "number1", '="Target "&TEXT(SatisfactionTarget,"0.0")&" out '
                    'of 5"'),
        ("Checklists finished",
         '=IFERROR(AVERAGEIFS([[Jobs.checklist_pct]],'
         '[[Jobs.checklist_pct]],">0"),"")', "percent0",
         '="Across every job with a checklist against it"'),
        ("Jobs that ran over",
         '=COUNTIF([[Job Costing.alert]],"Check - ran over*")', "int",
         '="More hours than the estimate. A pricing problem, not a '
         'people problem"'),
        ("Loss-making jobs",
         '=COUNTIF([[Job Costing.alert]],"Loss*")', "int",
         '="Cost more than they earned"'),
        ("Average profit a job",
         '=IFERROR(AVERAGEIFS([[Job Costing.profit]],'
         '[[Job Costing.revenue]],">0"),"")', "money",
         '="After wages, supplies and travel"'),
        ("Average profit an hour",
         '=IFERROR(AVERAGEIFS([[Job Costing.per_hour]],'
         '[[Job Costing.per_hour]],">0"),"")', "money",
         '="The number to compare every job against"'),
        ("Team utilisation",
         '=IFERROR(SUMIFS([[Timesheets.hours]],[[Timesheets.billable]],'
         '"Yes")/SUM([[Timesheets.hours]]),"")', "percent0",
         '="Billed hours as a share of paid hours. Target "&'
         'TEXT(UtilisationTarget,"0%")'),
        ("Complaints per hundred jobs",
         '=IFERROR(COUNTIF([[Complaints.date]],">0")/(COUNTIF('
         '[[Jobs.status]],"Completed")+COUNTIF([[Jobs.status]],'
         '"Invoiced"))*100,"")', "number1",
         '="Under two is good. Under one is exceptional"'),
    ]
    for label, formula, style, meaning in ops_rows:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula), fmts.calc(style))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(meaning), fmts.calc("text"))
        row += 1
    row += 1

    # ---------------------------------------------- expenses -----------
    row = s.section(row, "Where the money goes", LAST_COL)
    row = s.table_head(row, 0, ["Category", "This month", "This year",
                                "Share", "Monthly average"],
                       [26, 14, 14, 10, 15])
    categories = ["Cleaning supplies", "Equipment purchase",
                  "Equipment repair", "Vehicle fuel",
                  "Vehicle maintenance", "Business insurance",
                  "Wages & payroll", "Subcontractors",
                  "Marketing & advertising", "Software & apps",
                  "Phone & internet", "Office & rent", "Uniforms",
                  "Training & certification", "Accounting & legal",
                  "Bank & card fees", "Laundry", "Waste disposal",
                  "Travel & parking", "Other"]
    first_cat = row
    for cat in categories:
        s.write(row, 0, cat, fmts.chrome("label_bold"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=SUMIFS([[Expenses.total]],'
                         f'[[Expenses.category]],"{cat}",'
                         f'[[Expenses.month]],DashKey)'),
            fmts.calc("money"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=SUMIFS([[Expenses.total]],'
                         f'[[Expenses.category]],"{cat}",'
                         f'[[Expenses.date]],">="&DATE(PlanYear,1,1),'
                         f'[[Expenses.date]],"<="&DATE(PlanYear,12,31))'),
            fmts.calc("money0"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IFERROR(C{row + 1}/SUMIFS('
                         f'[[Expenses.total]],[[Expenses.date]],'
                         f'">="&DATE(PlanYear,1,1),[[Expenses.date]],'
                         f'"<="&DATE(PlanYear,12,31)),"")'),
            fmts.calc("percent0"))
        ws.write_formula(row, 4, f'=C{row + 1}/12', fmts.calc("money"))
        row += 1
    last_cat = row - 1
    ws.conditional_format(first_cat, 2, last_cat, 2,
                          {"type": "data_bar", "bar_color": "#C86B6B"})
    s.write(row, 0, "Total", fmts.chrome("label_bold"))
    for col, letter in [(1, "B"), (2, "C")]:
        ws.write_formula(
            row, col,
            f"=SUM({letter}{first_cat + 1}:{letter}{last_cat + 1})",
            fmts.chrome("total_value", num_format=cfg.money_format_0))
    row += 2

    # ---------------------------------------------- debtors ------------
    row = s.section(row, "Who owes you money", LAST_COL)
    row = s.table_head(row, 0, ["Age", "Invoices", "Amount",
                                "What to do"],
                       [18, 11, 14, 56])
    buckets = [
        ("Not due yet", "Keep an eye on it"),
        ("Due now", "Today is the day to send the gentle reminder"),
        ("1-30 days", "Phone, do not email. It works far better"),
        ("31-60 days", "Stop work for this client until they pay"),
        ("61-90 days", "Final demand, in writing, with your terms"),
        ("90+ days", "Decide: formal recovery, or write it off and "
                     "learn from it"),
    ]
    first_b = row
    for label, advice in buckets:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=COUNTIF([[Invoices.age]],"{label}")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=SUMIFS([[Invoices.outstanding]],'
                         f'[[Invoices.age]],"{label}")'),
            fmts.calc("money"))
        s.merge(row, 3, LAST_COL, advice, fmts.chrome("note"))
        row += 1
    ws.conditional_format(first_b + 2, 2, row - 1, 2,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("bad")})
    s.write(row, 0, "Total owed", fmts.chrome("label_bold"))
    ws.write_formula(row, 2, f"=SUM(C{first_b + 1}:C{row})",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    s.note(row,
           "Three tables on this page pay for the workbook on their "
           "own: the advertising table (stop paying for what does not "
           "work), the service table (do more of what earns most an "
           "hour) and the debtors table (chase at thirty days, not "
           "ninety).", LAST_COL, height=38)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
