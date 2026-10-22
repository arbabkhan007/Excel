"""`Inventory` - retail stock and backbar, counted the same way.

Stock on hand = opening + received - used in services - sold - waste
+ adjustments. Every one of those columns is visible, so when the
count is wrong you can see which part of the story is wrong rather
than typing over the answer.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, rank_desc, round_price,
                     status_cf, yesno)

NAME = "Inventory"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ProductPrefix", digits=3, width=10),
        Col("name", "Product", 28, "text"),
        Col("brand", "Brand", 16, "text"),
        Col("category", "Category", 24, "text",
            validate=list_from("ProductCategoryList"),
            note="Retail lines and backbar lines are both welcome "
                 "here - the category keeps them apart."),
        Col("size", "Size", 10, "text",
            note="250ml, 100g, a tube, a box of 50."),
        Col("sku", "SKU / barcode", 15, "text"),
        Col("supplier", "Supplier", 20, "text",
            validate=list_from("SupplierList"),
            note="From the supplier list at the bottom of the "
                 "Suppliers & Purchases tab."),
        Col("unit_cost", "Cost each", 11, "money",
            validate=number_check(0, 100000),
            note="What you pay, per unit, delivered."),
        Col("retail_price", "Retail price", 11, "money",
            validate=number_check(0, 100000),
            note="What the client pays. Leave blank for backbar "
                 "stock that is never sold."),
        Col("opening", "Opening stock", 12, "number1",
            validate=number_check(0, 1000000),
            note="What was on the shelf when you started using this "
                 "workbook. Set it once; after that let the movements "
                 "do the work."),
        Col("used", "Used in services", 13, "number1",
            validate=number_check(0, 1000000),
            note="Backbar usage - tubes of colour, lash trays, wax "
                 "pots. Count it monthly rather than per client."),
        Col("waste", "Waste", 9, "number1",
            validate=number_check(0, 1000000),
            note="Spilt, expired, broken, given away."),
        Col("adjust", "Adjustment", 11, "number1",
            validate=number_check(-1000000, 1000000),
            note="The honest column. Plus or minus whatever it takes "
                 "to match the real count, with a note saying why."),
        Col("reorder_level", "Reorder at", 11, "number1",
            validate=number_check(0, 1000000),
            note="The point at which you want to be told. Blank "
                 "means the workbook works one out from how fast it "
                 "sells."),
        Col("reorder_qty", "Order quantity", 12, "number1",
            validate=number_check(0, 1000000)),
        Col("counted_on", "Last counted", 12, "date",
            validate=date_check()),
        Col("expires", "Expires", 11, "date", validate=date_check(),
            note="For developer, lash adhesive, anything with a "
                 "shelf life."),
        Col("retail_item", "Sold to clients?", 13, "flag",
            validate=yesno(),
            note="No for pure backbar stock, so it stays out of the "
                 "retail figures."),
        Col("notes", "Notes", 22, "wrap"),
        # ---------------------------------------------------- worked out --
        Col("received", "Received", 10, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Purchases.qty_in]],'
                '[[Purchases.product]],{id}{r}),2)'),
            note="Everything booked in on the purchases tab."),
        Col("sold", "Sold", 9, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.qty_used]],[[Sales.item]],'
                '{id}{r},[[Sales.kind]],"Retail"),2)'),
            note="Units that went out through the till."),
        Col("stock", "Stock on hand", 12, "number1",
            formula=blank(
                "name",
                'ROUND(IF({opening}{r}="",0,{opening}{r})+'
                '{received}{r}-IF({used}{r}="",0,{used}{r})-'
                '{sold}{r}-IF({waste}{r}="",0,{waste}{r})+'
                'IF({adjust}{r}="",0,{adjust}{r}),2)'),
            cf=[{"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "bad"}],
            note="Opening + received - used - sold - waste + "
                 "adjustments. Nothing hidden."),
        Col("weekly_use", "Used a week", 11, "number1",
            formula=blank(
                "name",
                'ROUND(({sold}{r}+IF({used}{r}="",0,{used}{r}))/'
                'MAX(4,TrackingWeeks),2)'),
            hidden=True,
            note="Average weekly movement since you started tracking."),
        Col("weeks_cover", "Weeks left", 10, "number1",
            formula=blank(
                "name",
                'IF({weekly_use}{r}<=0,"",'
                'ROUND({stock}{r}/{weekly_use}{r},1))'),
            cf=[{"type": "cell", "criteria": "<",
                 "value": "=StockWarningWeeks", "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 6,
                 "fmt": "warn"}],
            note="At the current rate. Under the warning level on "
                 "Settings and it goes red."),
        Col("reorder_used", "Reorder point", 11, "number1",
            formula=blank(
                "name",
                'IF({reorder_level}{r}<>"",{reorder_level}{r},'
                'ROUND({weekly_use}{r}*StockWarningWeeks,1))'),
            hidden=True),
        Col("status", "Stock status", 12, "status",
            formula=blank(
                "name",
                'IF({stock}{r}<=0,"Out of stock",'
                'IF({stock}{r}<={reorder_used}{r},"Low stock",'
                'IF({retail_item}{r}="No","Backbar","In stock")))'),
            cf=status_cf()),
        Col("order_qty", "Order this many", 12, "number1",
            formula=blank(
                "name",
                'IF({status}{r}="In stock","",IF({status}{r}="Backbar",'
                'IF({stock}{r}<={reorder_used}{r},'
                'MAX(1,IF({reorder_qty}{r}="",'
                'ROUND({weekly_use}{r}*ReorderWeeks,0),'
                '{reorder_qty}{r})),""),'
                'MAX(1,IF({reorder_qty}{r}="",'
                'ROUND({weekly_use}{r}*ReorderWeeks,0),'
                '{reorder_qty}{r}))))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Your order quantity if you set one, otherwise "
                 "enough to cover the reorder window on Settings."),
        Col("value_cost", "Value at cost", 12, "money",
            formula=blank(
                "name",
                'ROUND({stock}{r}*IF({unit_cost}{r}="",0,'
                '{unit_cost}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="What is sitting on the shelf, in money. Usually a "
                 "bigger number than anybody expects."),
        Col("value_retail", "Value at retail", 12, "money",
            formula=blank(
                "name",
                'ROUND({stock}{r}*IF({retail_price}{r}="",0,'
                '{retail_price}{r}),2)'),
            hidden=True),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF(IF({retail_price}{r}="",0,{retail_price}{r})=0,"",'
                '({retail_price}{r}-IF({unit_cost}{r}="",0,'
                '{unit_cost}{r}))/{retail_price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0.3,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": 0.4,
                 "fmt": "good"}]),
        Col("suggested_raw", "Suggested raw", 11, "money",
            formula=blank(
                "name",
                'IF({unit_cost}{r}="","",{unit_cost}{r}*RetailMarkup)'),
            hidden=True),
        Col("suggested", "Suggested price", 12, "money",
            formula=blank(
                "name",
                'IF({suggested_raw}{r}="","",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="Cost x the retail markup on Settings, rounded by "
                 "your rounding rule."),
        Col("revenue", "Retail revenue", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.item]],{id}{r},'
                '[[Sales.kind]],"Retail"),2)')),
        Col("profit", "Retail profit", 11, "money",
            formula=blank(
                "name",
                'ROUND({revenue}{r}-{sold}{r}*IF({unit_cost}{r}="",0,'
                '{unit_cost}{r}),2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}],
            hidden=True),
        rank_desc("revenue", header="Seller rank"),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({unit_cost}{r}="","Add - no cost, no stock value",'
                'IF({status}{r}="Out of stock",'
                '"Out of stock - order today",'
                'IF({status}{r}="Low stock","Order - running low",'
                'IF(IF({expires}{r}="",FALSE,'
                '{expires}{r}<=TodayDate+30),'
                '"Check - expiring within a month",'
                'IF(IF({retail_item}{r}="Yes",'
                '{margin}{r}<0.25,FALSE),'
                '"Below - thin margin for a retail line",'
                'IF(IF({counted_on}{r}="",TRUE,'
                'TodayDate-{counted_on}{r}>90),'
                '"Due - not counted in three months","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Product Inventory",
        subtitle="Retail and backbar stock, what is left, what it is "
                 "worth and what to order.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).inventory,
        help_text="Set the opening stock once. After that, purchases "
                  "and sales move the numbers; you only type the "
                  "backbar usage, the waste and the odd adjustment.",
        extra_notes=[
            "Backbar stock is tracked the same way as retail, it just "
            "leaves by a different door: 'used in services' instead "
            "of 'sold'. Count it once a month and the weeks-left "
            "column does the rest.",
            "The adjustment column exists because real counts never "
            "match. Use it, write why in the notes, and keep the "
            "movement columns honest.",
            "Order quantity is a suggestion, not an order. Nothing is "
            "sent anywhere - this workbook has no connection to your "
            "supplier.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Stock summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The stock room", last)
    sheet.table_head(row, 0, [
        "Products listed", "Value at cost", "Value at retail",
        "Out of stock", "Low stock", "To reorder", "Retail revenue",
        "Average retail margin", "Expiring in 30 days"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Inventory.name]],"<>")',
        '=ROUND(SUM([[Inventory.value_cost]]),2)',
        '=ROUND(SUM([[Inventory.value_retail]]),2)',
        '=COUNTIFS([[Inventory.status]],"Out of stock")',
        '=COUNTIFS([[Inventory.status]],"Low stock")',
        '=ROUND(SUM([[Inventory.order_qty]]),0)',
        '=ROUND(SUM([[Inventory.revenue]]),2)',
        '=IFERROR(SUMPRODUCT(IFERROR([[Inventory.margin]]*1,0))/'
        'MAX(1,COUNTIFS([[Inventory.retail_item]],"Yes")),0)',
        '=COUNTIFS([[Inventory.expires]],">0",'
        '[[Inventory.expires]],"<="&TodayDate+30)',
    ], ["int", "money", "money", "int", "int", "number1", "money",
        "percent", "int"])
