"""`Bundle Lines` - what is in each bundle.

One row per product per bundle. The costs and the retail prices come
straight from the catalog, so a bundle re-prices itself the moment a
component does.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, lookup_product

NAME = "Bundle Lines"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("bundle", "Bundle", 12, "id", validate=list_from("BundleIDs"),
            note="The bundle ID from the Bundles tab."),
        Col("bundle_name", "Bundle name", 24, "text",
            formula=blank(
                "bundle",
                'IFERROR(INDEX([[Bundles.name]],MATCH({bundle}{r},'
                '[[Bundles.id]],0)),"Not found")')),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs")),
        Col("product_name", "Product name", 26, "text",
            formula=lookup_product("name")),
        Col("qty", "Quantity", 10, "int",
            validate=number_check(1, 10000, decimal=False)),
        Col("swap", "Substitutable?", 12, "flag",
            validate=list_from("YesNoList"),
            note="Yes if a buyer can pick a colour or scent. "
                 "Substitutes should cost about the same or the "
                 "bundle margin moves with every order."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("unit_cost", "Cost each", 11, "money",
            formula=lookup_product("total_cost", missing="")),
        Col("unit_price", "Retail each", 11, "money",
            formula=lookup_product("price", missing="")),
        Col("line_cost", "Line cost", 11, "money",
            formula=blank(
                "product",
                'IF(OR({qty}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({qty}{r}*{unit_cost}{r},2))')),
        Col("line_price", "Line retail", 11, "money",
            formula=blank(
                "product",
                'IF(OR({qty}{r}="",{unit_price}{r}=""),"",'
                'ROUND({qty}{r}*{unit_price}{r},2))')),
        Col("seq", "Line", 7, "int",
            formula=blank(
                "bundle",
                'COUNTIFS({bundle_first}:{bundle}{r},{bundle}{r})'),
            hidden=True),
        Col("key", "Key", 14, "text",
            formula=blank("bundle", '{bundle}{r}&"|"&{seq}{r}'),
            hidden=True),
        Col("check", "Check", 24, "status",
            formula=(
                '=IF(AND({bundle}{r}="",{product}{r}=""),"",'
                'IF({bundle}{r}="","Add - which bundle?",'
                'IF({product}{r}="","Add - which product?",'
                'IF({product_name}{r}="Not found",'
                '"Check - no such product ID",'
                'IF({qty}{r}="","Add - how many?",'
                'IF({unit_price}{r}="",'
                '"Check - that product has no price","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Bundle Contents",
        subtitle="What goes in each set, costed from the catalog.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).bundle_lines,
        help_text="One row per product per bundle. Change a "
                  "component's price in the catalog and every "
                  "bundle that contains it follows.",
        extra_notes=[
            "Put the quantity in rather than listing the same "
            "product twice - the bundle totals count quantities, "
            "not rows.",
            "If a buyer can choose a substitute, make sure the "
            "alternatives cost about the same. A free upgrade "
            "chosen by every customer is a price cut.",
        ],
    )
