"""`Purchases` - stock orders, and the supplier book underneath them.

Top half: every order you place, what it cost, whether it turned up.
Bottom half: an editable supplier directory that feeds the supplier
drop-down on this tab and on the inventory tab. Add a wholesaler
there and it appears in the lists straight away.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, dynamic_name, quote_sheet
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_product, month_col,
                     status_cf, yesno)

NAME = "Purchases"

#: Supplier directory written under the order grid.
SUPPLIER_COLUMNS = [
    ("Supplier", 24, "text"),
    ("Contact", 18, "text"),
    ("Phone", 16, "text"),
    ("Email", 26, "text"),
    ("Account number", 16, "text"),
    ("Minimum order", 14, "money"),
    ("Lead days", 11, "int"),
    ("Pays by", 14, "text"),
    ("Notes", 30, "text"),
]
SUPPLIER_ROWS = 24


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "PurchasePrefix", digits=4, width=11),
        Col("date", "Ordered", 11, "date", validate=date_check()),
        Col("supplier", "Supplier", 22, "text",
            validate=list_from("SupplierList"),
            note="From the supplier book at the bottom of this tab."),
        Col("product", "Product ID", 11, "id",
            validate=list_from("ProductList"),
            note="The inventory ID. One line per product, as on the "
                 "invoice."),
        Col("product_name", "Product", 24, "text",
            formula=lookup_product("name")),
        Col("qty", "Quantity", 10, "number1",
            validate=number_check(0, 100000)),
        Col("unit_cost", "Cost each", 11, "money",
            validate=number_check(0, 100000),
            note="What this order charged. Leave blank to use the "
                 "cost on the inventory tab."),
        Col("shipping", "Delivery", 10, "money",
            validate=number_check(0, 100000)),
        Col("tax", "Tax", 10, "money", validate=number_check(0, 100000)),
        Col("status", "Status", 14, "text",
            validate=list_from("PurchaseStatusList"), cf=status_cf(),
            note="Nothing counts into stock until this says Received "
                 "or Part received."),
        Col("expected", "Expected", 11, "date", validate=date_check()),
        Col("received_on", "Received", 11, "date", validate=date_check()),
        Col("qty_received", "Quantity in", 11, "number1",
            validate=number_check(0, 100000),
            note="Only needed when a delivery is short. Blank means "
                 "all of it arrived."),
        Col("invoice", "Invoice ref", 14, "text"),
        Col("paid", "Paid?", 9, "flag", validate=yesno()),
        Col("notes", "Notes", 22, "wrap"),
        # ---------------------------------------------------- worked out --
        month_col("date"),
        Col("cost_used", "Cost used", 11, "money",
            formula=blank(
                "date",
                'IF({unit_cost}{r}<>"",{unit_cost}{r},'
                'IFERROR(INDEX([[Inventory.unit_cost]],MATCH('
                '{product}{r},[[Inventory.id]],0))*1,0))'),
            hidden=True),
        Col("goods", "Goods", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({qty}{r}="",0,{qty}{r})*{cost_used}{r},2)')),
        Col("total", "Order total", 11, "money",
            formula=blank(
                "date",
                'ROUND({goods}{r}+IF({shipping}{r}="",0,'
                '{shipping}{r})+IF({tax}{r}="",0,{tax}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}]),
        Col("qty_in", "Into stock", 11, "number1",
            formula=blank(
                "date",
                'IF({status}{r}="Received",'
                'IF({qty_received}{r}<>"",{qty_received}{r},'
                'IF({qty}{r}="",0,{qty}{r})),'
                'IF({status}{r}="Part received",'
                'IF({qty_received}{r}="",0,{qty_received}{r}),0))'),
            note="What the inventory tab counts as received. Draft "
                 "and ordered lines add nothing until they land."),
        Col("short", "Short by", 10, "number1",
            formula=blank(
                "date",
                'IF({status}{r}="Ordered","",'
                'IF({status}{r}="Draft","",'
                'ROUND(IF({qty}{r}="",0,{qty}{r})-{qty_in}{r},2)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            hidden=True),
        Col("lead_days", "Days to arrive", 12, "int",
            formula=blank(
                "date",
                'IF({received_on}{r}="","",{received_on}{r}-{date}{r})'),
            hidden=True),
        Col("price_change", "Cost moved", 11, "money",
            formula=blank(
                "date",
                'IF({unit_cost}{r}="","",IFERROR(ROUND({unit_cost}{r}-'
                'INDEX([[Inventory.unit_cost]],MATCH({product}{r},'
                '[[Inventory.id]],0)),2),""))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "good"}],
            note="This order against the cost on file. A positive "
                 "number is a price rise you may not have noticed."),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "date",
                'IF({product}{r}="","Add - which product?",'
                'IF({status}{r}="","Add - no status yet",'
                'IF(IF({status}{r}="Ordered",IF({expected}{r}="",FALSE,'
                '{expected}{r}<TodayDate),FALSE),'
                '"Late - expected and not arrived",'
                'IF({short}{r}>0,"Check - delivery came up short",'
                'IF(IF({price_change}{r}="",FALSE,'
                '{price_change}{r}>0),"Check - cost has gone up",'
                'IF({paid}{r}<>"Yes","Due - invoice not paid",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Suppliers & Purchases",
        subtitle="What you ordered, what it cost, whether it turned "
                 "up - and the supplier book underneath.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).purchases,
        help_text="Log the order when you place it and set the status "
                  "to Received when it lands. Only received quantities "
                  "reach the stock figures.",
        extra_notes=[
            "Cost moved is the column that pays for this workbook. "
            "Wholesale prices creep up quietly; this says by how "
            "much, on which line, on which order.",
            "Short deliveries: set the status to Part received and "
            "type what actually arrived in 'quantity in'. The "
            "difference shows as short, and the stock stays honest.",
            "The supplier book at the bottom of this tab is a plain "
            "editable list. Whatever you type in the first column "
            "appears in the supplier drop-downs here and on the "
            "inventory tab.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Order summary, then the editable supplier directory."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Ordering, all time", last)
    sheet.table_head(row, 0, [
        "Orders", "Goods", "Delivery", "Order total", "Received",
        "Still on order", "Unpaid invoices", "Late deliveries",
        "Lines where cost rose"])
    row = sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Purchases.date]],"<>")',
        '=ROUND(SUM([[Purchases.goods]]),2)',
        '=ROUND(SUM([[Purchases.shipping]]),2)',
        '=ROUND(SUM([[Purchases.total]]),2)',
        '=ROUND(SUMIFS([[Purchases.total]],[[Purchases.status]],'
        '"Received"),2)',
        '=ROUND(SUMIFS([[Purchases.total]],[[Purchases.status]],'
        '"Ordered"),2)',
        '=ROUND(SUMIFS([[Purchases.total]],[[Purchases.paid]],'
        '"<>Yes"),2)',
        '=COUNTIFS([[Purchases.alert]],"Late*")',
        '=COUNTIFS([[Purchases.price_change]],">0")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "int", "int"])

    # ------------------------------------------------- supplier directory --
    row += 1
    row = sheet.section(row, "Supplier book", last)
    row = sheet.note(
        row,
        "Type your wholesalers here. The first column feeds the supplier "
        "drop-down on this tab and on the inventory tab, so add a new "
        "one and it is immediately available everywhere. Blank rows are "
        "ignored.",
        last, height=30)
    row += 1

    headers = [h for h, _, _ in SUPPLIER_COLUMNS]
    widths = [w for _, w, _ in SUPPLIER_COLUMNS]
    sheet.table_head(row, 0, headers, widths)
    first = row + 1                      # 0 based first input row
    rows = sample(cfg).suppliers if sample(cfg).suppliers else []
    for offset in range(SUPPLIER_ROWS):
        data = rows[offset] if offset < len(rows) else {}
        for index, (header, _width, style) in enumerate(SUPPLIER_COLUMNS):
            key = header.lower().replace(" ", "_")
            sheet.write(first + offset, index, data.get(key),
                        fmts.input(style))

    sheet_ref = quote_sheet(NAME)
    sheet.book.define_name(
        "SupplierList",
        dynamic_name(NAME, "A", first + 1, first + SUPPLIER_ROWS))
    sheet.book.define_name(
        "SupplierLeadDays",
        f"={sheet_ref}!$G${first + 1}:$G${first + SUPPLIER_ROWS}")

    # spend per supplier, so the book is worth reading
    row = first + SUPPLIER_ROWS + 1
    sheet.table_head(row, 0, ["Suppliers listed", "Spend, all time",
                              "Spend this month", "Biggest supplier"])
    sheet.calc_row(row + 1, 0, [
        f'=COUNTIF({sheet_ref}!$A${first + 1}:$A${first + SUPPLIER_ROWS},'
        f'"?*")',
        '=ROUND(SUM([[Purchases.total]]),2)',
        '=ROUND(SUMIFS([[Purchases.total]],[[Purchases.month]],'
        'ThisMonthKey),2)',
        '=IFERROR(INDEX([[Purchases.supplier]],MATCH(MAX('
        '[[Purchases.total]]),[[Purchases.total]],0)),"")',
    ], ["int", "money", "money", "text"])
