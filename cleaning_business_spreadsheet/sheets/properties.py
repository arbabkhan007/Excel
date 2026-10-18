"""`Properties` - the places you actually clean.

A client can have one address or twenty. This tab holds the detail that
decides how long a job takes and how much it should cost: size, number
of bathrooms, what state it is in, how you get in and where to park.
The estimated hours and the suggested price are worked out from those,
using the rates on the Settings tab.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_client,
                     status_cf)

NAME = "Properties"

_JOBS = "[[Jobs.property]]"


def spec(cfg: Config) -> SheetSpec:
    area_header = f"Area ({cfg.area_unit})"
    columns = [
        id_col("name", "PropertyPrefix", digits=4, width=10,
               note="The ID the jobs, quotes and checklists refer to."),
        Col("name", "Property", 28, "text",
            note="A name you will recognise on a job sheet: "
                 "'Harper - family home', 'Willow Park - Block A'."),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientIDs"),
            note="Pick the client from the drop down. Everything else "
                 "about them is pulled through."),
        Col("client_name", "Client", 22, "text",
            formula=lookup_client("name"),
            note="Automatic, from the Clients tab."),
        Col("type", "Property type", 15, "text",
            validate=list_from("PropertyTypeList")),
        Col("address", "Address", 36, "wrap"),
        Col("area", area_header, 11, "number",
            validate=number_check(0, 1000000),
            note="Rough floor area. Used by the area based pricing "
                 "method - pace it out once and write it down."),
        Col("beds", "Bedrooms", 9, "int",
            validate=number_check(0, 50, decimal=False)),
        Col("baths", "Bathrooms", 10, "int",
            validate=number_check(0, 50, decimal=False),
            note="Bathrooms and kitchens are where the time goes."),
        Col("rooms", "Other rooms", 11, "int",
            validate=number_check(0, 200, decimal=False)),
        Col("floors", "Floor types", 14, "text",
            note="Carpet, hardwood, tile, mixed - it changes what you "
                 "take with you."),
        Col("condition", "Condition", 20, "text",
            validate=list_from("ConditionList"),
            note="The honest answer. It multiplies the estimated hours "
                 "- the factors are on the Lists tab."),
        Col("pets", "Pets", 12, "text",
            note="Pet hair is real extra time, and the cleaner needs "
                 "to know before they open the door."),
        Col("access", "How we get in", 20, "text",
            validate=list_from("AccessList")),
        Col("key_status", "Key", 14, "status",
            validate=list_from("KeyStatusList"), cf=status_cf(),
            note="If you hold keys, keep this up to date. It is the "
                 "first thing anyone asks when a key goes missing."),
        Col("entry_notes", "Entry instructions", 32, "wrap",
            note="Lockbox position, alarm panel, which door sticks. "
                 "Treat codes as sensitive - see the note below the "
                 "table."),
        Col("parking", "Parking", 16, "text",
            validate=list_from("ParkingList"),
            note="Nothing wastes a morning like a permit you did not "
                 "know about."),
        Col("distance", f"Distance ({cfg.distance_unit})", 10, "number1",
            validate=number_check(0, 10000),
            note="One way, from your base. The travel cost on a quote "
                 "doubles it for the round trip."),
        Col("checklist", "Checklist", 18, "text",
            validate=list_from("ServiceNames"),
            note="Which checklist the cleaner should work through at "
                 "this property. Set up on the Checklists tab."),
        Col("supplies_onsite", "Client supplies?", 11, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Yes if the client provides the products - it comes "
                 "off the job cost."),
        Col("active", "Active?", 8, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("notes", "Notes", 38, "wrap"),
        # -------------------------------------------------- the estimate --
        Col("size_hours", "Size hours", 10, "number1",
            formula=blank(
                "name",
                'MAX(IF({beds}{r}="",0,{beds}{r})*0.45+'
                'IF({baths}{r}="",0,{baths}{r})*0.6+'
                'IF({rooms}{r}="",0,{rooms}{r})*0.3,'
                'IF({area}{r}="",0,{area}{r})/650)'),
            hidden=True,
            note="Rooms and bathrooms, or floor area - whichever "
                 "suggests the bigger job."),
        Col("est_hours", "Estimated hours", 13, "number1",
            formula=blank(
                "name",
                'ROUND(MAX(1,{size_hours}{r})*IFERROR(INDEX('
                'ConditionFactors,MATCH({condition}{r},ConditionNames,0)),'
                '1),2)'),
            note="How long one cleaner would take. Split it between "
                 "two cleaners and you have the time on site.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("visit_hours", "Hours on site", 11, "number1",
            formula=blank(
                "name",
                'ROUND({est_hours}{r}/MAX(1,CleanersPerJob),2)'),
            note="Estimated hours shared between the standard number "
                 "of cleaners from the Settings tab."),
        Col("suggested", "Suggested price", 13, "money",
            formula=blank(
                "name",
                'MAX(MinimumCharge,ROUND({est_hours}{r}*'
                'IF(IFERROR(INDEX([[Clients.rate]],MATCH({client}{r},'
                '[[Clients.id]],0)),0)>0,INDEX([[Clients.rate]],'
                'MATCH({client}{r},[[Clients.id]],0)),HourlyRate)+'
                'IF({pets}{r}="",0,IF({pets}{r}="None",0,PetSurcharge)),'
                '2))'),
            note="Hours times the rate, with the pet surcharge and the "
                 "minimum charge applied. A starting point for the "
                 "quote, not a contract.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("labour_cost", "Labour cost", 11, "money",
            formula=blank("name", 'ROUND({est_hours}{r}*CostPerHour,2)'),
            hidden=True),
        Col("est_margin", "Estimated margin", 12, "percent0",
            formula=blank(
                "name",
                'IF({suggested}{r}=0,"",({suggested}{r}-{labour_cost}{r}-'
                '{est_hours}{r}*SuppliesPerHour)/{suggested}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}]),
        # ------------------------------------------------- the real world --
        Col("jobs", "Jobs here", 9, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _JOBS + ',{id}{r}))'),
        Col("last_visit", "Last visit", 12, "date",
            formula='=IF({id}{r}="","",IF(' +
                    max_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
                           '([[Jobs.status]]="Completed")+'
                           '([[Jobs.status]]="Invoiced")') +
                    '=0,"",' +
                    max_if("[[Jobs.date]]", f"{_JOBS}={{id}}{{r}}",
                           '([[Jobs.status]]="Completed")+'
                           '([[Jobs.status]]="Invoiced")') + '))'),
        Col("actual_hours", "Average actual hours", 13, "number1",
            formula='=IF({id}{r}="","",IFERROR(AVERAGEIFS('
                    '[[Jobs.actual_hours]],' + _JOBS + ',{id}{r},'
                    '[[Jobs.actual_hours]],">0"),""))',
            note="What the job really takes, averaged over every visit. "
                 "Compare it with the estimate - that gap is your "
                 "profit or your loss."),
        Col("variance", "Estimate vs actual", 12, "number1",
            formula='=IF(OR({actual_hours}{r}="",{est_hours}{r}=""),"",'
                    '{actual_hours}{r}-{est_hours}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 0.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}],
            note="Positive means the job takes longer than you price "
                 "for. Re-quote it or re-scope it."),
        Col("revenue", "Revenue here", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Jobs.price]],' + _JOBS +
                    ',{id}{r},[[Jobs.status]],"Invoiced")+SUMIFS('
                    '[[Jobs.price]],' + _JOBS + ',{id}{r},'
                    '[[Jobs.status]],"Completed"))'),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({client}{r}="","Add - no client linked",'
                'IF(IFERROR(INDEX([[Clients.name]],MATCH({client}{r},'
                '[[Clients.id]],0)),"")="","Check - client ID not found",'
                'IF(AND({area}{r}="",{beds}{r}=""),'
                '"Add - size needed to price it",'
                'IF(AND({variance}{r}<>"",{variance}{r}>0.5),'
                '"Check - runs over the estimate",'
                'IF({access}{r}="","Add - how do we get in?","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Property & Site Database",
        subtitle="The places you clean, how long they take and what "
                 "they should be worth.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).properties,
        help_text="One row per address. Fill in the size, the condition "
                  "and how you get in, and the workbook estimates the "
                  "hours and suggests a price straight away.",
        extra_notes=[
            "The estimate is deliberately conservative: it takes the "
            "larger of 'rooms and bathrooms' or 'floor area divided by "
            "650', then multiplies by the condition factor. Compare it "
            "with the actual hours column after a month and adjust the "
            "factors on the Lists tab until they match your reality.",
            "Entry instructions, alarm codes and key locations are "
            "sensitive. Anyone you send this file to can read them - "
            "share printouts of a single job sheet instead, and keep "
            "codes out of the file altogether if you can.",
            "A property that keeps running over its estimate is not a "
            "bad cleaner, it is a bad price. The variance column tells "
            "you which ones to re-quote.",
        ],
    )
