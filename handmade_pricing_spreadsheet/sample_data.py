"""Demo data: a season in the life of Willow & Thread Handmade.

A one-maker studio in Asheville that crochets, pours candles, solders a
little silver, sews a few bags and turns out the occasional walnut
board - plus a pattern shop that sells while she sleeps. Every tab
opens with believable numbers so the dashboard, the price calculators
and the reports mean something the moment the file is opened, and
`--blank` strips the lot out for your own data.

Everything is dated relative to the build date, so the workbook never
looks stale.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .config import Config


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def pid(n: int) -> str:
    return f"P-{n:04d}"


def mid(n: int) -> str:
    return f"M-{n:04d}"


def bid(n: int) -> str:
    return f"BN-{n:04d}"


def btid(n: int) -> str:
    return f"BT-{n:04d}"


def wcid(n: int) -> str:
    return f"WC-{n:04d}"


def woid(n: int) -> str:
    return f"WO-{n:04d}"


@dataclass
class Sample:
    products: List[Dict[str, Any]] = field(default_factory=list)
    materials: List[Dict[str, Any]] = field(default_factory=list)
    material_usage: List[Dict[str, Any]] = field(default_factory=list)
    labor: List[Dict[str, Any]] = field(default_factory=list)
    packaging: List[Dict[str, Any]] = field(default_factory=list)
    shipping: List[Dict[str, Any]] = field(default_factory=list)
    overhead: List[Dict[str, Any]] = field(default_factory=list)
    channels: List[Dict[str, Any]] = field(default_factory=list)
    crochet: List[Dict[str, Any]] = field(default_factory=list)
    jewelry: List[Dict[str, Any]] = field(default_factory=list)
    candles: List[Dict[str, Any]] = field(default_factory=list)
    clothing: List[Dict[str, Any]] = field(default_factory=list)
    woodworking: List[Dict[str, Any]] = field(default_factory=list)
    digital: List[Dict[str, Any]] = field(default_factory=list)
    bundles: List[Dict[str, Any]] = field(default_factory=list)
    bundle_lines: List[Dict[str, Any]] = field(default_factory=list)
    variations: List[Dict[str, Any]] = field(default_factory=list)
    batches: List[Dict[str, Any]] = field(default_factory=list)
    sales: List[Dict[str, Any]] = field(default_factory=list)
    promotions: List[Dict[str, Any]] = field(default_factory=list)
    price_history: List[Dict[str, Any]] = field(default_factory=list)
    competitors: List[Dict[str, Any]] = field(default_factory=list)
    waste: List[Dict[str, Any]] = field(default_factory=list)
    development: List[Dict[str, Any]] = field(default_factory=list)
    events: List[Dict[str, Any]] = field(default_factory=list)
    wholesale_customers: List[Dict[str, Any]] = field(default_factory=list)
    wholesale_orders: List[Dict[str, Any]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------

_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg) if cfg.include_sample_data else Sample()
    return _CACHE[key]


def _build(cfg: Config) -> Sample:
    t = cfg.today
    s = Sample()

    # ---------------------------------------------------------- materials --
    # pack_qty is the number of `unit`s in the pack, so the cost per gram,
    # per centimetre or per item falls out of the pack price by itself.
    s.materials = [
        dict(name="Merino worsted yarn, natural", type="Yarn & thread",
             detail="100 g skein, 200 m", supplier="Blue Ridge Fibre Co",
             code="BRF-MW-NAT", pack_qty=100, unit="g", pack_cost=9.40,
             pack_ship=0.00, discount=0.00, waste=0.04, stock=1850,
             reorder=600, last_cost=8.60, priced_on=d(t, -24),
             notes="Up 9% this spring. Hats and cowls."),
        dict(name="Merino worsted yarn, moss", type="Yarn & thread",
             detail="100 g skein, 200 m", supplier="Blue Ridge Fibre Co",
             code="BRF-MW-MSS", pack_qty=100, unit="g", pack_cost=9.40,
             pack_ship=0.00, discount=0.00, waste=0.04, stock=700,
             reorder=600, last_cost=8.60, priced_on=d(t, -24)),
        dict(name="Cotton 8/4 twine", type="Yarn & thread",
             detail="250 g cone", supplier="Carolina Cotton",
             code="CC-84-ECR", pack_qty=250, unit="g", pack_cost=12.80,
             pack_ship=0.00, discount=0.05, waste=0.03, stock=1100,
             reorder=500, last_cost=12.80, priced_on=d(t, -96),
             notes="Market bags. 5% trade discount over 10 cones."),
        dict(name="Cotton DK, mixed colours", type="Yarn & thread",
             detail="50 g ball", supplier="Blue Ridge Fibre Co",
             code="BRF-DK-MIX", pack_qty=50, unit="g", pack_cost=3.20,
             pack_ship=0.00, waste=0.08, stock=640, reorder=300,
             last_cost=3.05, priced_on=d(t, -62),
             notes="Amigurumi. Waste is high - colour changes."),
        dict(name="Polyester toy stuffing", type="Other supply",
             detail="500 g bag", supplier="Craft Warehouse",
             code="CW-STF-500", pack_qty=500, unit="g", pack_cost=8.90,
             pack_ship=0.00, waste=0.02, stock=1400, reorder=500,
             last_cost=8.90, priced_on=d(t, -120)),
        dict(name="Safety eyes 10 mm", type="Elastic & notions",
             detail="Pack of 20 pairs", supplier="Craft Warehouse",
             code="CW-EYE-10", pack_qty=20, unit="pair", pack_cost=4.60,
             pack_ship=0.00, waste=0.00, stock=34, reorder=10,
             last_cost=4.60, priced_on=d(t, -120)),
        dict(name="Sterling silver wire 1.5 mm", type="Metal & wire",
             detail="Half hard, 30 g coil", supplier="Rio Metals",
             code="RIO-SS-15", pack_qty=30, unit="g", pack_cost=42.00,
             pack_ship=4.50, discount=0.00, waste=0.06, stock=74,
             reorder=30, last_cost=36.00, priced_on=d(t, -18),
             notes="Silver moved 17% in six months. Re-price rings."),
        dict(name="Sterling silver chain 1.2 mm", type="Metal & wire",
             detail="Fine cable, sold by the metre",
             supplier="Rio Metals", code="RIO-CH-12", pack_qty=100,
             unit="cm", pack_cost=14.00, pack_ship=0.00, waste=0.05,
             stock=420, reorder=150, last_cost=12.40,
             priced_on=d(t, -18)),
        dict(name="Sea glass cabochons", type="Stones & crystals",
             detail="Tumbled, 15-20 mm", supplier="Tidewater Finds",
             code="TW-SG-MIX", pack_qty=10, unit="each", pack_cost=18.00,
             pack_ship=3.20, waste=0.05, stock=22, reorder=10,
             last_cost=18.00, priced_on=d(t, -55)),
        dict(name="Jump rings & clasps, sterling",
             type="Beads & findings", detail="Mixed pack of 50",
             supplier="Rio Metals", code="RIO-FND-50", pack_qty=50,
             unit="each", pack_cost=9.50, pack_ship=0.00, waste=0.02,
             stock=180, reorder=60, last_cost=8.90,
             priced_on=d(t, -18)),
        dict(name="Coconut soy wax", type="Wax",
             detail="5 kg box, 464 blend", supplier="Blue Ridge Candle",
             code="BRC-CS-5", pack_qty=5, unit="kg", pack_cost=38.00,
             pack_ship=6.00, discount=0.00, waste=0.03, stock=11,
             reorder=4, last_cost=34.50, priced_on=d(t, -31),
             notes="Box price includes freight - that is why pack "
                   "delivery is filled in."),
        dict(name="Amber & fir fragrance oil", type="Fragrance & oils",
             detail="500 ml bottle", supplier="Blue Ridge Candle",
             code="BRC-FO-AF", pack_qty=500, unit="ml", pack_cost=42.00,
             pack_ship=0.00, waste=0.02, stock=720, reorder=250,
             last_cost=39.00, priced_on=d(t, -31)),
        dict(name="Lavender fragrance oil", type="Fragrance & oils",
             detail="250 ml bottle", supplier="Blue Ridge Candle",
             code="BRC-FO-LAV", pack_qty=250, unit="ml", pack_cost=26.00,
             pack_ship=0.00, waste=0.02, stock=180, reorder=100,
             last_cost=26.00, priced_on=d(t, -31)),
        dict(name="Cotton wick CD-12, pretabbed", type="Wicks",
             detail="Pack of 100", supplier="Blue Ridge Candle",
             code="BRC-WK-CD12", pack_qty=100, unit="each",
             pack_cost=12.00, pack_ship=0.00, waste=0.02, stock=260,
             reorder=100, last_cost=12.00, priced_on=d(t, -31)),
        dict(name="Amber glass jar 220 ml", type="Vessels & jars",
             detail="With bamboo lid, case of 24",
             supplier="Glassware Direct", code="GD-AM-220", pack_qty=24,
             unit="each", pack_cost=33.60, pack_ship=8.40, waste=0.02,
             stock=46, reorder=24, last_cost=31.20,
             priced_on=d(t, -40),
             notes="Breakage in transit is why the waste allowance "
                   "is not zero."),
        dict(name="Candle tin 120 ml", type="Vessels & jars",
             detail="Seamless, case of 48", supplier="Glassware Direct",
             code="GD-TIN-120", pack_qty=48, unit="each",
             pack_cost=38.40, pack_ship=0.00, waste=0.01, stock=96,
             reorder=48, last_cost=38.40, priced_on=d(t, -40)),
        dict(name="Washed linen, flax", type="Fabric",
             detail="150 cm wide, sold by the metre",
             supplier="Mill End Fabrics", code="MEF-LIN-FLX",
             pack_qty=300, unit="cm", pack_cost=54.00, pack_ship=0.00,
             waste=0.12, stock=640, reorder=200, last_cost=49.50,
             priced_on=d(t, -27),
             notes="Quantities in centimetres of length at full "
                   "width. 12% layout waste on the bias."),
        dict(name="Cotton lining, natural", type="Fabric",
             detail="112 cm wide", supplier="Mill End Fabrics",
             code="MEF-COT-NAT", pack_qty=500, unit="cm",
             pack_cost=32.00, pack_ship=0.00, waste=0.10, stock=380,
             reorder=150, last_cost=32.00, priced_on=d(t, -27)),
        dict(name="Zip, webbing & hardware set",
             type="Hardware & fixings", detail="Antique brass, per bag",
             supplier="Bag Makers Supply", code="BMS-HW-AB",
             pack_qty=10, unit="each", pack_cost=16.00, pack_ship=0.00,
             waste=0.00, stock=18, reorder=10, last_cost=14.80,
             priced_on=d(t, -47)),
        dict(name="Woven label, folded", type="Labels & tags",
             detail="Pack of 100", supplier="Label Lab",
             code="LL-WV-100", pack_qty=100, unit="each",
             pack_cost=24.00, pack_ship=0.00, waste=0.01, stock=210,
             reorder=50, last_cost=24.00, priced_on=d(t, -150)),
        dict(name="Walnut board, 25 mm", type="Timber & boards",
             detail="Kiln dried, 180 cm x 20 cm",
             supplier="Appalachian Timber", code="AT-WAL-25",
             pack_qty=180, unit="cm", pack_cost=76.00, pack_ship=0.00,
             waste=0.25, stock=240, reorder=180, last_cost=68.00,
             priced_on=d(t, -21),
             notes="A quarter of every board is offcut and sawdust. "
                   "Pretending otherwise is why boards feel "
                   "unprofitable."),
        dict(name="White oak board, 20 mm", type="Timber & boards",
             detail="240 cm x 15 cm", supplier="Appalachian Timber",
             code="AT-OAK-20", pack_qty=240, unit="cm", pack_cost=68.00,
             pack_ship=0.00, waste=0.22, stock=180, reorder=120,
             last_cost=61.00, priced_on=d(t, -21)),
        dict(name="Hard wax oil, satin", type="Finish & glue",
             detail="500 ml tin", supplier="Appalachian Timber",
             code="AT-HWO-500", pack_qty=500, unit="ml",
             pack_cost=32.00, pack_ship=0.00, waste=0.05, stock=310,
             reorder=150, last_cost=32.00, priced_on=d(t, -21)),
        dict(name="Sanding discs & glue", type="Tools & consumables",
             detail="Mixed grits, per project pack",
             supplier="Appalachian Timber", code="AT-ABR-MIX",
             pack_qty=25, unit="each", pack_cost=18.50, pack_ship=0.00,
             waste=0.00, stock=14, reorder=10, last_cost=18.50,
             priced_on=d(t, -70)),
    ]

    # ----------------------------------------------------------- products --
    s.products = [
        dict(name="Willow cable beanie", sku_input="", category="Crochet & knitting",
             status="Active", season="Autumn", difficulty="Moderate",
             channel="Etsy", pack_profile="Standard",
             ship_option="First class - Domestic to 250g", weight=140,
             stock=9, price_input=78.00, made=d(t, -12),
             launched=d(t, -400),
             url="willowandthread.co/beanie",
             notes="Costed on the Crochet tab, not the bill of "
                   "materials."),
        dict(name="Chunky market bag", sku_input="",
             category="Bags & leather", status="Active",
             season="All year", difficulty="Simple", channel="Etsy",
             pack_profile="Budget",
             ship_option="First class - Domestic to 500g", weight=320,
             stock=6, price_input=75.00, made=d(t, -20),
             launched=d(t, -520)),
        dict(name="Woodland fox amigurumi", sku_input="",
             category="Toys & amigurumi", status="Active",
             season="Christmas", difficulty="Moderate", channel="Etsy",
             pack_profile="Gift",
             ship_option="First class - Domestic to 500g", weight=180,
             stock=4, price_input=78.00, made=d(t, -7),
             launched=d(t, -300)),
        dict(name="Hammered silver stacking ring", sku_input="",
             category="Jewelry", status="Active", season="All year",
             difficulty="Moderate", channel="Etsy",
             pack_profile="Gift",
             ship_option="First class - Domestic to 100g", weight=25,
             stock=14, price_input=58.00, made=d(t, -9),
             launched=d(t, -620)),
        dict(name="Sea glass pendant", sku_input="",
             category="Jewelry", status="Active", season="Summer",
             difficulty="Simple", channel="Etsy", pack_profile="Gift",
             ship_option="First class - Domestic to 100g", weight=30,
             stock=11, price_input=52.00, made=d(t, -15),
             launched=d(t, -420)),
        dict(name="Amber & fir candle 220 ml", sku_input="",
             category="Candles & home fragrance", status="Active",
             season="Winter", difficulty="Simple", channel="Etsy",
             pack_profile="Standard",
             ship_option="Priority - Domestic to 1,000g", weight=520,
             stock=22, price_input=32.00, made=d(t, -5),
             launched=d(t, -260)),
        dict(name="Lavender tin candle 120 ml", sku_input="",
             category="Candles & home fragrance", status="Active",
             season="All year", difficulty="Very simple",
             channel="Etsy", pack_profile="Budget",
             ship_option="First class - Domestic to 500g", weight=260,
             stock=31, price_input=19.00, made=d(t, -5),
             launched=d(t, -260)),
        dict(name="Linen crossbody bag", sku_input="",
             category="Bags & leather", status="Active",
             season="All year", difficulty="Complex", channel="Etsy",
             pack_profile="Standard",
             ship_option="Priority - Domestic to 1,000g", weight=380,
             stock=3, price_input=195.00, made=d(t, -30),
             launched=d(t, -180)),
        dict(name="Linen pinafore apron", sku_input="",
             category="Clothing & accessories", status="Active",
             season="All year", difficulty="Moderate",
             channel="Own website", pack_profile="Standard",
             ship_option="Priority - Domestic to 1,000g", weight=420,
             stock=5, price_input=148.00, made=d(t, -40),
             launched=d(t, -210)),
        dict(name="Walnut serving board", sku_input="",
             category="Woodworking", status="Active", season="All year",
             difficulty="Moderate", channel="Own website",
             pack_profile="Standard",
             ship_option="Courier - Domestic to 5,000g", weight=1400,
             stock=4, price_input=135.00, made=d(t, -18),
             launched=d(t, -340)),
        dict(name="Oak floating shelf 60 cm", sku_input="",
             category="Woodworking", status="Made to order",
             season="All year", difficulty="Simple",
             channel="Own website", pack_profile="Budget",
             ship_option="Courier - Domestic to 5,000g", weight=2100,
             stock=0, price_input=None, made=d(t, -60),
             launched=d(t, -150),
             notes="No price yet - the dashboard is right to "
                   "complain about it."),
        dict(name="Cable beanie crochet pattern", sku_input="",
             category="Digital download", status="Active",
             season="Autumn", difficulty="Very simple", channel="Etsy",
             weight=0, stock=0, price_input=8.50,
             launched=d(t, -380),
             notes="Instant download - no packaging profile, so "
                   "nothing is charged for a box that is never "
                   "posted."),
        dict(name="Woodland fox pattern", sku_input="",
             category="Digital download", status="Active",
             season="Christmas", difficulty="Very simple",
             channel="Etsy", weight=0, stock=0,
             price_input=9.50, launched=d(t, -120)),
        dict(name="Beeswax wrap set of three", sku_input="",
             category="Other handmade", status="Prototype",
             season="All year", difficulty="Simple", channel="Etsy",
             pack_profile="Budget",
             ship_option="First class - Domestic to 250g", weight=120,
             stock=0, price_input=None, launched=None,
             notes="Still costing it. No materials entered yet."),
    ]

    # ---------------------------------------------------- bill of materials
    # Only the products that are costed the generic way have BOM lines -
    # anything on a specialist tab is costed there instead, which is what
    # the check column on those tabs is watching for.
    s.material_usage = [
        dict(product=pid(2), material=mid(3), qty=240, waste=0.03,
             optional="No", notes="Weighed off the cone."),
        dict(product=pid(2), material=mid(20), qty=1, optional="No",
             notes="Woven label in the side seam."),
        dict(product=pid(3), material=mid(4), qty=85, waste=0.08,
             optional="No", notes="Three colours, so more waste."),
        dict(product=pid(3), material=mid(5), qty=60, optional="No"),
        dict(product=pid(3), material=mid(6), qty=1, optional="No",
             notes="One pair of safety eyes."),
        dict(product=pid(3), material=mid(20), qty=1, optional="No"),
        dict(product=pid(8), material=mid(17), qty=70, waste=0.12,
             optional="No", notes="Outer, cut on the straight grain."),
        dict(product=pid(8), material=mid(18), qty=60, waste=0.08,
             optional="No", notes="Lining and inside pocket."),
        dict(product=pid(8), material=mid(19), qty=1, optional="No",
             notes="Zip, webbing, slider and two rings."),
        dict(product=pid(8), material=mid(20), qty=1, optional="No"),
        dict(product=pid(10), material=mid(21), qty=55, waste=0.25,
             optional="No",
             notes="A quarter of the board is offcut - that is the "
                   "honest number."),
        dict(product=pid(10), material=mid(23), qty=18, waste=0.05,
             optional="No", notes="Three coats of hard wax oil."),
        dict(product=pid(10), material=mid(24), qty=1.5, optional="No",
             notes="Discs die fast on walnut."),
        dict(product=pid(10), material=mid(20), qty=1, optional="Yes",
             notes="Branded tag, only on gift orders."),
    ]

    # ----------------------------------------------------------- the hours
    s.labor = [
        dict(product=pid(2), stage="Making", who="Me", minutes=85,
             basis="Per item", skilled="No",
             notes="Base, sides and handles."),
        dict(product=pid(2), stage="Weaving in", who="Me", minutes=10,
             basis="Per item", skilled="No"),
        dict(product=pid(2), stage="Packing", who="Me", minutes=6,
             basis="Per item", skilled="No"),
        dict(product=pid(2), stage="Photography", who="Me", minutes=25,
             basis="Per batch", batch_qty=6, skilled="No",
             notes="One shoot covers the batch."),
        dict(product=pid(3), stage="Making", who="Me", minutes=90,
             basis="Per item", skilled="Yes",
             notes="Body, head, four legs, tail and ears."),
        dict(product=pid(3), stage="Assembly", who="Me", minutes=18,
             basis="Per item", skilled="Yes",
             notes="Sewing up and getting the face right."),
        dict(product=pid(3), stage="Finishing", who="Me", minutes=8,
             basis="Per item", skilled="No"),
        dict(product=pid(3), stage="Packing", who="Me", minutes=4,
             basis="Per item", skilled="No"),
        dict(product=pid(8), stage="Cutting & prep", who="Me",
             minutes=30, basis="Per item", skilled="Yes"),
        dict(product=pid(8), stage="Sewing", who="Me", minutes=80,
             basis="Per item", skilled="Yes",
             notes="Zip, lining, strap, topstitching."),
        dict(product=pid(8), stage="Finishing", who="Me", minutes=15,
             basis="Per item", skilled="No"),
        dict(product=pid(8), stage="Packing", who="Me", minutes=8,
             basis="Per item", skilled="No"),
        dict(product=pid(8), stage="Photography", who="Me", minutes=30,
             basis="Per batch", batch_qty=4, skilled="No"),
        dict(product=pid(10), stage="Cutting & prep", who="Me",
             minutes=25, basis="Per item", skilled="Yes",
             notes="Milling, thicknessing, cutting to length."),
        dict(product=pid(10), stage="Sanding & smoothing", who="Me",
             minutes=35, basis="Per item", skilled="No",
             notes="Always longer than anybody admits."),
        dict(product=pid(10), stage="Painting & staining", who="Me",
             minutes=20, basis="Per item", skilled="No",
             notes="Three coats of oil, with waiting in between."),
        dict(product=pid(10), stage="Packing", who="Me", minutes=12,
             basis="Per item", skilled="No",
             notes="Boards need real protection in the post."),
        dict(product=pid(10), stage="Photography", who="Me",
             minutes=25, basis="Per batch", batch_qty=3, skilled="No"),
    ]

    # --------------------------------------------------------- packaging --
    s.packaging = [
        dict(name="Kraft mailer box, small", type="Mailer & box",
             supplier="Packmart", pack_qty=50, pack_cost=34.50, qty=1,
             weight=45, budget="Yes", standard="Yes", gift="Yes",
             premium="Yes", reusable="No"),
        dict(name="Poly mailer", type="Polybag", supplier="Packmart",
             pack_qty=100, pack_cost=14.00, qty=1, weight=12,
             budget="Yes", standard="No", gift="No", premium="No",
             reusable="No", notes="Patterns and anything soft."),
        dict(name="Printed tissue paper", type="Tissue & wrap",
             supplier="Noissue", pack_qty=500, pack_cost=42.00, qty=2,
             weight=4, budget="No", standard="Yes", gift="Yes",
             premium="Yes", reusable="No"),
        dict(name="Kraft sticker seal", type="Label & sticker",
             supplier="Noissue", pack_qty=500, pack_cost=22.00, qty=2,
             weight=1, budget="Yes", standard="Yes", gift="Yes",
             premium="Yes", reusable="No"),
        dict(name="Thank you card", type="Thank you card",
             supplier="Moo", pack_qty=250, pack_cost=38.00, qty=1,
             weight=5, budget="No", standard="Yes", gift="Yes",
             premium="Yes", reusable="No",
             notes="Handwritten. Worth more than it costs."),
        dict(name="Care card", type="Care card", supplier="Moo",
             pack_qty=250, pack_cost=26.00, qty=1, weight=3,
             budget="No", standard="Yes", gift="Yes", premium="Yes",
             reusable="No", notes="Washing, burning and storing."),
        dict(name="Baker's twine", type="Ribbon & twine",
             supplier="Packmart", pack_qty=80, pack_cost=9.00, qty=1,
             weight=2, budget="No", standard="No", gift="Yes",
             premium="Yes", reusable="No"),
        dict(name="Rigid gift box", type="Gift box",
             supplier="Packmart", pack_qty=50, pack_cost=76.00, qty=1,
             weight=95, budget="No", standard="No", gift="Yes",
             premium="Yes", reusable="Yes",
             notes="The thing people keep. Worth the money on "
                   "jewelry."),
        dict(name="Bubble wrap sheet", type="Fragile protection",
             supplier="Packmart", pack_qty=100, pack_cost=19.00, qty=1,
             weight=15, budget="No", standard="Yes", gift="Yes",
             premium="Yes", reusable="No",
             notes="Candles and boards."),
        dict(name="Wax seal & ribbon", type="Ribbon & twine",
             supplier="Etsy supplier", pack_qty=100, pack_cost=28.00,
             qty=1, weight=4, budget="No", standard="No", gift="No",
             premium="Yes", reusable="No"),
    ]

    # ---------------------------------------------------------- shipping --
    s.shipping = [
        dict(service="First class", zone="Domestic", max_weight=100,
             cost=4.25, packaging_extra=0.00, charge=4.95,
             payer="Buyer pays", tracked="No", days=3, carrier="USPS",
             notes="Jewelry in a padded envelope."),
        dict(service="First class", zone="Domestic", max_weight=250,
             cost=5.40, packaging_extra=0.00, charge=5.95,
             payer="Buyer pays", tracked="Yes", days=3, carrier="USPS"),
        dict(service="First class", zone="Domestic", max_weight=500,
             cost=7.10, packaging_extra=0.00, charge=6.95,
             payer="Buyer pays", tracked="Yes", days=3, carrier="USPS",
             notes="You absorb 15 cents on every one of these."),
        dict(service="Priority", zone="Domestic", max_weight=1000,
             cost=9.85, packaging_extra=0.35, charge=8.95,
             payer="Buyer pays", tracked="Yes", days=2, carrier="USPS",
             notes="Candles and bags."),
        dict(service="Priority", zone="Domestic", max_weight=2000,
             cost=12.40, packaging_extra=0.35, charge=11.95,
             payer="Buyer pays", tracked="Yes", days=2, carrier="USPS"),
        dict(service="Courier", zone="Domestic", max_weight=5000,
             cost=18.60, packaging_extra=1.20, charge=16.00,
             payer="Flat rate", tracked="Yes", days=4, carrier="UPS",
             notes="Boards. The flat rate loses a little on every "
                   "order and sells a lot more of them."),
        dict(service="International tracked", zone="Rest of world",
             max_weight=500, cost=24.50, packaging_extra=0.00,
             charge=24.50, payer="Buyer pays", tracked="Yes", days=12,
             carrier="USPS", notes="Customs form every time."),
        dict(service="Collection / pickup", zone="Local", cost=0.00,
             charge=0.00, payer="Collection only", tracked="No",
             days=0, carrier="Studio",
             notes="Local buyers collect from the studio."),
    ]

    # ----------------------------------------------------------- overhead --
    s.overhead = [
        dict(name="Studio rent", type="Studio rent",
             supplier="River Arts District", amount=260.00,
             frequency="Monthly", start=d(t, -700), essential="Yes",
             notes="Shared bench space, 24 hour access."),
        dict(name="Electricity & heat", type="Utilities",
             supplier="Duke Energy", amount=48.00, frequency="Monthly",
             start=d(t, -700), essential="Yes",
             notes="The wax melter is not cheap to run."),
        dict(name="Maker's insurance", type="Insurance",
             supplier="ACT Insurance", amount=288.00,
             frequency="Yearly", start=d(t, -200), essential="Yes",
             notes="Product liability - candles need it."),
        dict(name="Etsy Plus", type="Shop subscription",
             supplier="Etsy", amount=10.00, frequency="Monthly",
             start=d(t, -400), essential="No"),
        dict(name="Website & domain", type="Website & domain",
             supplier="Squarespace", amount=192.00, frequency="Yearly",
             start=d(t, -300), essential="Yes"),
        dict(name="Design software", type="Software",
             supplier="Adobe", amount=14.99, frequency="Monthly",
             start=d(t, -500), essential="No",
             notes="Patterns and listing photos."),
        dict(name="Bookkeeping", type="Accounting",
             supplier="Hill & Co", amount=240.00, frequency="Yearly",
             start=d(t, -300), essential="Yes"),
        dict(name="Phone & internet (business share)",
             type="Phone & internet", supplier="Spectrum",
             amount=35.00, frequency="Monthly", start=d(t, -700),
             essential="Yes", notes="Half the household bill."),
        dict(name="Tool replacement fund", type="Tool replacement",
             supplier="-", amount=360.00, frequency="Yearly",
             start=d(t, -365), essential="Yes",
             notes="Hooks, blades, sanding belts, the next wax "
                   "melter. Put it aside before you need it."),
        dict(name="Photography props & backdrops",
             type="Photography props", supplier="Various",
             amount=120.00, frequency="Yearly", start=d(t, -200),
             essential="No"),
        dict(name="Guild membership", type="Memberships",
             supplier="Southern Highland Craft Guild", amount=60.00,
             frequency="Yearly", start=d(t, -250), essential="No",
             notes="Pays for itself in fair bookings."),
    ]

    # ----------------------------------------------------------- channels --
    reviewed = d(t, -38)
    s.channels = [
        dict(name="Etsy", kind="Percent + fixed", listing_fee=0.20,
             listing_sales=2, transaction=0.065, processing_pct=0.03,
             processing_fixed=0.25, ads_fee=0.15, ads_share=0.20,
             regulatory=0.0040, currency_fee=0.00, other_pct=0.00,
             other_fixed=0.00, subscription=10.00,
             fees_on_shipping="Yes", reviewed=reviewed,
             url="etsy.com/legal/fees",
             notes="Offsite ads are 15% but only on the orders that "
                   "come through them - about one in five here. "
                   "Verify the current official fee schedule for "
                   "your country before you rely on these."),
        dict(name="Own website", kind="Percent + fixed",
             listing_fee=0.00, listing_sales=1, transaction=0.00,
             processing_pct=0.029, processing_fixed=0.30, ads_fee=0.00,
             ads_share=0.00, regulatory=0.00, currency_fee=0.00,
             other_pct=0.00, other_fixed=0.00, subscription=16.00,
             fees_on_shipping="Yes", reviewed=reviewed,
             url="squarespace.com/pricing",
             notes="Card processing only. The subscription belongs "
                   "in overhead, not in the price."),
        dict(name="Craft fair", kind="Percent + fixed",
             listing_fee=0.00, listing_sales=1, transaction=0.00,
             processing_pct=0.026, processing_fixed=0.10, ads_fee=0.00,
             ads_share=0.00, regulatory=0.00, currency_fee=0.00,
             other_pct=0.00, other_fixed=0.00, subscription=0.00,
             fees_on_shipping="No", reviewed=reviewed,
             url="squareup.com/us/en/pricing",
             notes="Card reader rate. The stall fee is on the "
                   "Events tab, where it belongs."),
        dict(name="Wholesale", kind="None", listing_fee=0.00,
             listing_sales=1, transaction=0.00, processing_pct=0.00,
             processing_fixed=0.00, ads_fee=0.00, ads_share=0.00,
             regulatory=0.00, currency_fee=0.00, other_pct=0.00,
             other_fixed=0.00, subscription=0.00,
             fees_on_shipping="No", reviewed=reviewed,
             notes="No fees - the discount is the cost."),
        dict(name="Instagram / direct", kind="Percent + fixed",
             listing_fee=0.00, listing_sales=1, transaction=0.00,
             processing_pct=0.029, processing_fixed=0.30, ads_fee=0.00,
             ads_share=0.00, regulatory=0.00, currency_fee=0.00,
             other_pct=0.00, other_fixed=0.00, subscription=0.00,
             fees_on_shipping="Yes", reviewed=reviewed,
             notes="A payment link. The cheapest place you sell."),
        dict(name="Amazon Handmade", kind="Percent of order",
             listing_fee=0.00, listing_sales=1, transaction=0.15,
             processing_pct=0.00, processing_fixed=0.00, ads_fee=0.00,
             ads_share=0.00, regulatory=0.00, currency_fee=0.00,
             other_pct=0.00, other_fixed=0.00, subscription=0.00,
             fees_on_shipping="Yes", reviewed=reviewed,
             url="sell.amazon.com/handmade",
             notes="15% referral fee. Nothing listed here yet - "
                   "the margin would have to be built in first."),
        dict(name="Other", kind="None", listing_fee=0.00,
             listing_sales=1, transaction=0.00, processing_pct=0.00,
             processing_fixed=0.00, subscription=0.00,
             fees_on_shipping="No", reviewed=reviewed,
             notes="Gifts, swaps and anything that does not fit."),
    ]

    # ------------------------------------------------------ craft tabs --
    # Each specialist tab costs the pieces that belong to it.  A piece
    # with a product ID feeds the catalog; a piece without one is a
    # costing you are still thinking about.
    s.crochet = [
        dict(product=pid(1), name="Willow cable beanie",
             pattern="Willow cable (my own)", size="Standard",
             yarn_weight="Worsted / Aran", hook="5.0 mm", yarn=mid(1),
             grams=95, waste=0.05, notions=0.25, stuffing=0.00,
             pattern_cost=0.00, hours_input=1.25, finish_mins=10,
             price_input=78.00,
             notes="Weighed before and after. 95 g is the real "
                   "number, not the 80 g the pattern claims."),
        dict(name="Moss chunky cowl", pattern="Moss cowl (my own)",
             size="Standard", yarn_weight="Bulky / chunky",
             hook="8.0 mm", yarn=mid(2), grams=150, waste=0.05,
             notions=0.25, hours_input=1.8, finish_mins=12,
             price_input=78.00,
             notes="No product ID yet - costing it before I list it."),
        dict(name="Cotton baby blanket, 90 cm",
             pattern="Granny square blanket", size="Large",
             yarn_weight="DK", hook="4.0 mm", yarn=mid(4), grams=480,
             waste=0.06, notions=0.25, hours_input=9.5,
             finish_mins=35, price_input=215.00,
             notes="Nine and a half hours. Look at what the hourly "
                   "return column says before listing this."),
    ]

    s.jewelry = [
        dict(product=pid(4), name="Hammered silver stacking ring",
             metal="Sterling silver", metal_id=mid(7), metal_g=3.8,
             scrap=0.15, recovery=0.60, stone="None", stone_count=0,
             stone_cost=0.00, findings=0.00, solder=0.45, plating=0.00,
             hallmark=0.00, design_mins=0, bench_mins=32,
             polish_mins=10, price_input=58.00,
             notes="Scrap is real on a soldered ring - 15% of it, "
                   "and I re-melt about 60%."),
        dict(product=pid(5), name="Sea glass pendant",
             metal="Sterling silver", metal_id=mid(7), metal_g=1.6,
             scrap=0.10, recovery=0.60, stone="Sea glass",
             stone_count=1, stone_cost=2.12, chain_cm=45,
             chain_id=mid(8), findings=0.57, solder=0.25,
             plating=0.00, hallmark=0.00, design_mins=0,
             bench_mins=26, polish_mins=8, price_input=52.00,
             notes="Chain is 45 cm of the fine cable, costed by the "
                   "centimetre."),
        dict(name="Copper cuff, hammered", metal="Copper",
             metal_rate_input=0.0620, metal_g=28, scrap=0.08,
             recovery=0.00, stone="None", stone_count=0,
             findings=0.00, solder=0.20, design_mins=15,
             bench_mins=35, polish_mins=10, price_input=52.00,
             notes="No material row for copper yet, so the rate a "
                   "gram is typed in instead."),
    ]

    s.candles = [
        dict(product=pid(6), name="Amber & fir candle 220 ml",
             wax_type="Coconut soy", wax_id=mid(11),
             vessel="Amber jar", vessel_ml=220, fill=0.85,
             density=0.90, melt_loss=0.05, fragrance_id=mid(12),
             fragrance_load=0.09, wick="Cotton", wick_count=1,
             wick_cost=0.12, vessel_cost=1.75, dye_cost=0.00,
             label_cost=0.24, batch_size=12, batch_mins=55,
             item_mins=6, cure_days=10, price_input=32.00,
             notes="Fill to 85% of the jar. Nobody pours to the "
                   "brim."),
        dict(product=pid(7), name="Lavender tin candle 120 ml",
             wax_type="Coconut soy", wax_id=mid(11), vessel="Tin",
             vessel_ml=120, fill=0.88, density=0.90, melt_loss=0.05,
             fragrance_id=mid(13), fragrance_load=0.08, wick="Cotton",
             wick_count=1, wick_cost=0.12, vessel_cost=0.80,
             dye_cost=0.00, label_cost=0.24, batch_size=24,
             batch_mins=65, item_mins=4, cure_days=10,
             price_input=19.00,
             notes="The batch minutes are spread over 24 tins, "
                   "which is why these look cheap to make."),
        dict(name="Wax melt clamshell", wax_type="Coconut soy",
             wax_id=mid(11), vessel="Clamshell", vessel_ml=70,
             fill=0.95, density=0.90, melt_loss=0.04,
             fragrance_id=mid(13), fragrance_load=0.10, wick="Cotton",
             wick_count=0, wick_cost=0.00, vessel_cost=0.38,
             label_cost=0.18, batch_size=36, batch_mins=50,
             item_mins=2, cure_days=7, price_input=8.50,
             notes="No wick at all - wick count zero, and the maths "
                   "copes."),
    ]

    s.clothing = [
        dict(product=pid(9), name="Linen pinafore apron",
             garment_type="Apron", size="M", fabric_id=mid(17),
             fabric_cm=180, layout_waste=0.12, interfacing=0.40,
             thread=0.35, notions=1.20, labels=0.24, cut_mins=25,
             sew_mins=55, finish_mins=20, press_mins=10, fit_mins=0,
             price_input=148.00,
             notes="Size M is the base. The size factor scales the "
                   "fabric and the time for everything else."),
        dict(name="Toddler pinafore, age 3", garment_type="Pinafore",
             size="Toddler", size_factor=0.55, fabric_id=mid(17),
             fabric_cm=95, layout_waste=0.15, lining_id=mid(18),
             lining_cm=60, interfacing=0.20, thread=0.25,
             notions=0.90, labels=0.24, cut_mins=20, sew_mins=48,
             finish_mins=15, press_mins=8, fit_mins=0,
             price_input=85.00,
             notes="Half the fabric, nearly all of the time. This "
                   "is why children's clothes are hard to price."),
    ]

    s.woodworking = [
        dict(product=pid(11), name="Oak floating shelf 60 cm",
             species="Oak", wood_id=mid(22), wood_qty=65,
             offcut_waste=0.22, hardware=6.40, glue=0.60,
             finish_type="Hard wax oil", finish_ml=22,
             finish_rate=0.0640, abrasives=1.80, machine_mins=18,
             machine_rate=8.00, mill_mins=15, build_mins=25,
             sand_mins=40, finish_mins=25, dry_days=2,
             notes="No price typed in yet - the suggested price on "
                   "this row is the one to beat."),
        dict(name="Live edge walnut shelf 90 cm", species="Walnut",
             wood_id=mid(21), wood_qty=95, offcut_waste=0.30,
             hardware=11.80, glue=0.80, finish_type="Hard wax oil",
             finish_ml=34, finish_rate=0.0640, abrasives=2.60,
             machine_mins=26, machine_rate=8.00, mill_mins=25,
             build_mins=35, sand_mins=55, finish_mins=35, dry_days=2,
             price_input=210.00,
             notes="Thirty percent offcut on a live edge slab. "
                   "Anybody quoting less has not measured."),
    ]

    s.digital = [
        dict(product=pid(12), name="Cable beanie crochet pattern",
             kind="Crochet pattern", formats="PDF",
             licence="Personal use", design_hours=6, write_hours=11,
             test_hours=4, photo_hours=3, tester_cost=0.00,
             software_cost=18.00, expected_sales=220,
             support_mins=1.5, update_hours=2, sold_so_far=143,
             listed=d(t, -380), price_input=8.50,
             notes="Two testers worked for a free pattern, so the "
                   "tester fee is zero and the hours are mine."),
        dict(product=pid(13), name="Woodland fox pattern",
             kind="Crochet pattern", formats="PDF, Video",
             licence="Small commercial", design_hours=9,
             write_hours=14, test_hours=5, photo_hours=4,
             tester_cost=45.00, software_cost=18.00,
             expected_sales=250, support_mins=2.5, update_hours=2,
             sold_so_far=38, listed=d(t, -120), price_input=9.50,
             notes="Amigurumi patterns generate questions. Two and "
                   "a half minutes of support a sale is honest."),
        dict(name="Christmas ornament SVG bundle",
             kind="SVG cut file", formats="SVG, PNG, DXF",
             licence="Small commercial", design_hours=12,
             write_hours=3, test_hours=2, photo_hours=2,
             tester_cost=0.00, software_cost=24.00,
             expected_sales=300, support_mins=0.5, update_hours=1,
             sold_so_far=0, listed=None, price_input=12.00,
             notes="Not listed yet. The break-even sales column "
                   "says how many it has to sell to pay for the "
                   "design time."),
    ]

    # ------------------------------------------------------- bundles --
    s.bundles = [
        dict(name="Cosy night in gift set", kind="Gift set",
             channel="Etsy", pack_profile="Gift", assembly_mins=12,
             discount=0.12, status="Active", season="Winter",
             notes="Beanie and the amber candle. The discount is "
                   "what the set gives away against buying both."),
        dict(name="Candle duo", kind="Multi-buy", channel="Etsy",
             pack_profile="Standard", assembly_mins=6, discount=0.10,
             status="Active", season="All year",
             notes="Buyer picks the two scents."),
        dict(name="New baby keepsake set", kind="Gift set",
             channel="Own website", pack_profile="Premium",
             assembly_mins=15, discount=0.08, price_input=92.00,
             status="Active", season="All year",
             notes="Priced by hand at 92 - the suggested price "
                   "column says whether that was brave or silly."),
    ]

    s.bundle_lines = [
        dict(bundle=bid(1), product=pid(1), qty=1, swap="No"),
        dict(bundle=bid(1), product=pid(6), qty=1, swap="Yes",
             notes="Scent of their choice."),
        dict(bundle=bid(2), product=pid(6), qty=1, swap="Yes"),
        dict(bundle=bid(2), product=pid(7), qty=1, swap="Yes"),
        dict(bundle=bid(3), product=pid(3), qty=1, swap="No",
             notes="The fox is the piece they remember."),
        dict(bundle=bid(3), product=pid(7), qty=1, swap="No"),
    ]

    # ----------------------------------------------------- variations --
    s.variations = [
        dict(product=pid(1), name="Beanie, extra large", kind="Size",
             size="Large", extra_mins=25, rush="Standard",
             channel="Etsy", offered="Yes",
             notes="More yarn and more rounds - the size factor "
                   "handles both."),
        dict(product=pid(1), name="Beanie with faux fur pom",
             kind="Personalisation", extra_material=3.40,
             extra_mins=8, rush="Standard", channel="Etsy",
             price_input=86.00, offered="Yes"),
        dict(product=pid(3), name="Fox with embroidered name",
             kind="Embroidery", extra_mins=22, personalisation=8.00,
             rush="Standard", channel="Etsy", offered="Yes",
             notes="Eight dollars for twenty two minutes. Look at "
                   "the uplift column before saying yes to this."),
        dict(product=pid(4), name="Ring, rush order",
             kind="Rush order", rush="Same week", channel="Etsy",
             offered="Yes",
             notes="Half again on the labour, because everything "
                   "else waits."),
        dict(product=pid(9), name="Apron, made to measure",
             kind="Size", size="Custom size", extra_mins=35,
             rush="Standard", channel="Own website", offered="Yes",
             notes="Fitting, a toile and a second pressing."),
        dict(product=pid(6), name="Candle, custom scent",
             kind="Scent", extra_material=1.20, extra_mins=10,
             rush="Standard", channel="Instagram / direct",
             offered="No",
             notes="Costed and declined - a ten minute change on a "
                   "32 dollar candle does not pay."),
    ]

    # -------------------------------------------------------- batches --
    s.batches = [
        dict(product=pid(6), started=d(t, -40), finished=d(t, -38),
             planned=12, made=12, seconds=2, sold_seconds=24.00,
             material_actual=64.80, hours_actual=2.6, extra_cost=0.00,
             status="Complete",
             notes="Two with sink holes - sold as seconds at the "
                   "fair."),
        dict(product=pid(7), started=d(t, -40), finished=d(t, -39),
             planned=24, made=24, seconds=1, sold_seconds=9.00,
             material_actual=71.50, hours_actual=3.1, extra_cost=0.00,
             status="Complete"),
        dict(product=pid(1), started=d(t, -25), finished=d(t, -18),
             planned=6, made=6, seconds=0, material_actual=58.40,
             hours_actual=10.8, extra_cost=0.00, status="Complete",
             notes="Ten point eight hours for six. That is 1.8 "
                   "each, not the 1.57 the costing assumes - and "
                   "the costing is the one that needs changing."),
        dict(product=pid(3), started=d(t, -14), finished=d(t, -7),
             planned=4, made=4, seconds=1, sold_seconds=22.00,
             material_actual=31.20, hours_actual=9.6,
             extra_cost=0.00, status="Complete",
             notes="One fox has a wonky ear. Sold at the pop-up."),
        dict(product=pid(10), started=d(t, -22), finished=d(t, -18),
             planned=3, made=3, seconds=0, material_actual=92.00,
             hours_actual=5.6, extra_cost=4.50, status="Complete",
             notes="Extra cost is a replacement sanding belt."),
        dict(product=pid(2), started=d(t, -6), planned=8,
             status="In progress",
             notes="Half crocheted. Nothing recorded as made yet, "
                   "which is why the row is flagged."),
    ]

    # ---------------------------------------------------- promotions --
    s.promotions = [
        dict(code="WELCOME10", name="Welcome to the shop",
             kind="Percent off", value=0.10, scope="Everything",
             start=d(t, -200), min_spend=0.00, base_margin=0.30,
             notes="Permanent, on the shop banner."),
        dict(code="AUTUMN20", name="Autumn sale",
             kind="Percent off", value=0.20, scope="Everything",
             start=d(t, -95), ends=d(t, -80), min_spend=0.00,
             base_margin=0.30,
             notes="Twenty percent off ate most of the margin. "
                   "Look at what it left before running it again."),
        dict(code="FREESHIP50", name="Free postage over 50",
             kind="Free shipping", value=7.10, scope="Order",
             start=d(t, -120), min_spend=50.00, base_margin=0.30,
             notes="Value is the postage you absorb on an average "
                   "qualifying order."),
        dict(code="BUNDLE5", name="Five off any set",
             kind="Amount off", value=5.00, scope="Bundles",
             start=d(t, -60), min_spend=60.00, limit=100,
             base_margin=0.26),
    ]

    # ------------------------------------------------- price history --
    s.price_history = [
        dict(date=d(t, -400), product=pid(1), old_price=48.00,
             new_price=65.00, cost_then=52.00, reason="First price",
             channel="Etsy", sold_before=0, sold_after=18,
             notes="Opening price was set by looking at other shops. "
                   "Then it was costed properly."),
        dict(date=d(t, -26), product=pid(1), old_price=65.00,
             new_price=78.00, cost_then=59.40,
             reason="Material cost rose", channel="Etsy",
             sold_before=11, sold_after=9,
             notes="Merino went up 9%. Sales did not move."),
        dict(date=d(t, -200), product=pid(2), old_price=62.00,
             new_price=75.00, cost_then=68.00,
             reason="Competitor move", channel="Etsy",
             sold_before=14, sold_after=16,
             notes="Still does not cover the hours. The bag is on "
                   "the list to redesign or retire."),
        dict(date=d(t, -30), product=pid(3), old_price=62.00,
             new_price=78.00, cost_then=88.00,
             reason="Underpriced", channel="Etsy",
             sold_before=6, sold_after=3,
             notes="Sixteen dollars more and it still loses money. "
                   "A two hour toy is a two hour toy."),
        dict(date=d(t, -18), product=pid(4), old_price=52.00,
             new_price=58.00, cost_then=33.90,
             reason="Material cost rose", channel="Etsy",
             sold_before=7, sold_after=6,
             notes="Silver moved 17%. This one had to move."),
        dict(date=d(t, -45), product=pid(5), old_price=44.00,
             new_price=52.00, cost_then=31.40,
             reason="Underpriced", channel="Etsy",
             sold_before=9, sold_after=7,
             notes="The margin column said nine percent. Eight "
                   "dollars fixed it and nobody complained."),
        dict(date=d(t, -120), product=pid(6), old_price=28.00,
             new_price=32.00, cost_then=12.80,
             reason="Rebrand / repositioning", channel="Etsy",
             sold_before=22, sold_after=26,
             notes="New label, better photos, higher price - and "
                   "it sold more, not less."),
        dict(date=d(t, -120), product=pid(7), old_price=17.00,
             new_price=19.00, cost_then=6.80, reason="Rounding tidy-up",
             channel="Etsy", sold_before=30, sold_after=28),
        dict(date=d(t, -180), product=pid(8), old_price=165.00,
             new_price=195.00, cost_then=128.00,
             reason="Labour rate change", channel="Etsy",
             sold_before=5, sold_after=4,
             notes="Put my own rate up from 20 to 25."),
        dict(date=d(t, -60), product=pid(9), old_price=128.00,
             new_price=148.00, cost_then=114.00,
             reason="Underpriced", channel="Own website",
             sold_before=4, sold_after=3,
             notes="Linen went up twice in a year and the price "
                   "had not moved once."),
        dict(date=d(t, -70), product=pid(10), old_price=118.00,
             new_price=135.00, cost_then=98.00,
             reason="Material cost rose", channel="Own website",
             sold_before=6, sold_after=5,
             notes="Walnut up 12%. The old price never covered the "
                   "sanding either."),
        dict(date=d(t, -90), product=pid(12), old_price=7.50,
             new_price=8.50, cost_then=5.00,
             reason="Testing a price", channel="Etsy",
             sold_before=18, sold_after=17,
             notes="A dollar more, the same number of sales. The "
                   "test paid for itself."),
        dict(date=d(t, -50), product=pid(13), old_price=6.50,
             new_price=9.50, cost_then=6.20,
             reason="Underpriced", channel="Etsy",
             sold_before=12, sold_after=9,
             notes="Two and a half minutes of support a sale has "
                   "to be paid for by somebody."),
    ]

    # -------------------------------------------------- competitors --
    s.competitors = [
        dict(product=pid(1), competitor="Highland Knits",
             shop="Etsy star seller", channel="Etsy", price=84.00,
             shipping=5.95, quality="Similar", handmade="Yes",
             reviews=1240, checked=d(t, -20),
             notes="Same yarn weight, no cable. Charging more."),
        dict(product=pid(1), competitor="FastFashionCo",
             shop="Amazon", channel="Amazon Handmade", price=19.99,
             shipping=0.00, quality="Mass produced", handmade="No",
             reviews=8600, checked=d(t, -20),
             notes="Machine made in a factory. Not the same "
                   "product and not your customer."),
        dict(product=pid(3), competitor="Thistle & Thread Toys",
             shop="Etsy", channel="Etsy", price=92.00, shipping=6.50,
             quality="Above mine", handmade="Yes", reviews=430,
             checked=d(t, -32),
             notes="Better photographs. That is the gap, not the "
                   "crochet."),
        dict(product=pid(4), competitor="Asheville Silverworks",
             shop="Own website", channel="Own website", price=72.00,
             shipping=0.00, quality="Above mine", handmade="Yes",
             reviews=96, checked=d(t, -14),
             notes="Hallmarked and boxed. Room to move up."),
        dict(product=pid(5), competitor="Tidewater Jewelry",
             shop="Etsy", channel="Etsy", price=46.00, shipping=4.50,
             quality="Similar", handmade="Yes", reviews=610,
             checked=d(t, -14)),
        dict(product=pid(6), competitor="Smoky Mountain Candle",
             shop="Etsy", channel="Etsy", price=28.00, shipping=7.95,
             quality="Similar", handmade="Yes", reviews=2100,
             checked=d(t, -26),
             notes="Cheaper jar, same wax. Their postage is "
                   "dearer, so the delivered price is level."),
        dict(product=pid(7), competitor="Smoky Mountain Candle",
             shop="Etsy", channel="Etsy", price=16.00, shipping=7.95,
             quality="Similar", handmade="Yes", reviews=2100,
             checked=d(t, -26)),
        dict(product=pid(8), competitor="Foothills Canvas",
             shop="Own website", channel="Own website", price=215.00,
             shipping=0.00, quality="Above mine", handmade="Yes",
             reviews=58, checked=d(t, -45),
             notes="Leather trim and free postage."),
        dict(product=pid(10), competitor="Blue Ridge Boards",
             shop="Craft fair", channel="Craft fair", price=112.00,
             shipping=0.00, quality="Below mine", handmade="Yes",
             reviews=0, checked=d(t, -35),
             notes="Thinner stock, machine oiled. Undercuts you "
                   "at every fair."),
        dict(product=pid(12), competitor="Loop & Lane Patterns",
             shop="Etsy", channel="Etsy", price=7.00, shipping=0.00,
             quality="Similar", handmade="N/A", reviews=3400,
             checked=d(t, -60),
             notes="Patterns are a volume game - the price is set "
                   "by the market, not by your hours."),
    ]

    # ------------------------------------------------------- waste --
    s.waste = [
        dict(date=d(t, -8), material=mid(1), qty=22, reason="Offcut",
             product=pid(1), avoidable="No",
             notes="Tail ends from six hats."),
        dict(date=d(t, -12), material=mid(21), qty=38,
             reason="Miscut", product=pid(10), avoidable="Yes",
             notes="Measured once. Expensive lesson in walnut."),
        dict(date=d(t, -15), material=mid(11), qty=0.42,
             reason="Failed pour", product=pid(6), avoidable="Yes",
             notes="Poured too hot, two jars with sink holes."),
        dict(date=d(t, -19), material=mid(15), qty=2,
             reason="Breakage in post", product=pid(6),
             avoidable="No", notes="Courier. Claim submitted."),
        dict(date=d(t, -24), material=mid(17), qty=45,
             reason="Normal trim waste", product=pid(8),
             avoidable="No"),
        dict(date=d(t, -30), material=mid(4), qty=36,
             reason="Mistake / frogged", product=pid(3),
             avoidable="Yes",
             notes="Frogged a whole fox body - yarn is reusable "
                   "but tired."),
        dict(date=d(t, -44), material=mid(12), qty=18,
             reason="Spoiled in storage", avoidable="Yes",
             notes="Left the cap off. Fragrance flashes off fast."),
        dict(date=d(t, -52), material=mid(22), qty=26,
             reason="Warped", product=pid(11), avoidable="No",
             notes="Board cupped after milling."),
        dict(date=d(t, -66), material=mid(3), qty=30,
             reason="Dye lot mismatch", product=pid(2),
             avoidable="No"),
        dict(date=d(t, -80), material=mid(9), qty=1,
             reason="Cracked", product=pid(5), avoidable="No",
             notes="Sea glass split while setting."),
    ]

    # ------------------------------------------------- development --
    s.development = [
        dict(name="Willow cable beanie", product=pid(1),
             stage="Recovered cost", started=d(t, -430),
             design_hours=6, sample_hours=5, photo_hours=3,
             sample_cost=18.60, other_cost=0.00, samples_sold=2,
             decided=d(t, -400),
             notes="Three samples before the crown sat right. Paid "
                   "for itself in the first season."),
        dict(name="Woodland fox amigurumi", product=pid(3),
             stage="Listed", started=d(t, -330), design_hours=11,
             sample_hours=9, photo_hours=4, sample_cost=14.20,
             other_cost=0.00, samples_sold=1, decided=d(t, -300),
             notes="Twenty four hours of development. The break "
                   "even column says how many foxes that is."),
        dict(name="Linen crossbody bag", product=pid(8),
             stage="Listed", started=d(t, -220), design_hours=9,
             sample_hours=12, photo_hours=4, sample_cost=42.80,
             other_cost=16.00, samples_sold=1, decided=d(t, -180),
             notes="Two toiles and a pattern draft."),
        dict(name="Beeswax wrap set", product=pid(14),
             stage="Sampling", started=d(t, -40), design_hours=3,
             sample_hours=4, photo_hours=0, sample_cost=22.40,
             other_cost=0.00, samples_sold=0,
             notes="Still deciding. Nothing listed, nothing "
                   "recovered."),
        dict(name="Ceramic wick holder", stage="Shelved",
             started=d(t, -150), design_hours=5, sample_hours=6,
             photo_hours=0, sample_cost=38.00, other_cost=0.00,
             samples_sold=0, decided=d(t, -120),
             notes="Shelved - kiln time costs more than the "
                   "holders would ever sell for. Written off on "
                   "purpose, and visible."),
        dict(name="Christmas ornament SVG bundle", stage="Testing",
             started=d(t, -25), design_hours=12, sample_hours=2,
             photo_hours=2, sample_cost=0.00, other_cost=24.00,
             samples_sold=0,
             notes="Digital, so the only cost is time and the "
                   "stock licence."),
    ]

    # ------------------------------------------------------ events --
    s.events = [
        dict(name="Asheville Spring Market", kind="Craft fair",
             date=d(t, -70), days=1, stall_fee=45.00, travel=18.00,
             accommodation=0.00, other_cost=22.00, prep_hours=9,
             stall_hours=8, target=350.00, footfall=900,
             notes="First fair of the year. Good weather, slow "
                   "crowd."),
        dict(name="River Arts Night Market", kind="Farmers market",
             date=d(t, -35), days=1, stall_fee=30.00, travel=12.00,
             accommodation=0.00, other_cost=8.00, prep_hours=5,
             stall_hours=6, target=240.00, footfall=520,
             notes="Evening market. Candles sell under lights."),
        dict(name="Holiday Pop-up at The Mill", kind="Pop-up shop",
             date=d(t, -8), days=1, stall_fee=60.00, travel=15.00,
             accommodation=0.00, other_cost=10.00, prep_hours=6,
             stall_hours=7, target=320.00, footfall=700),
        dict(name="Winter Makers Fair", kind="Christmas market",
             date=d(t, 24), days=2, stall_fee=120.00, travel=40.00,
             accommodation=95.00, other_cost=35.00, prep_hours=16,
             stall_hours=14, target=1400.00, footfall=2600,
             notes="Booked. The stock needed row says what has to "
                   "be made before then."),
    ]

    # ------------------------------------------ wholesale customers --
    s.wholesale_customers = [
        dict(name="The Mill Mercantile", kind="Independent shop",
             contact="Dana Pryor", email="dana@millmercantile.com",
             phone="828 555 0142", city="Asheville NC",
             terms="Net 30", discount=0.50, min_order=250.00,
             first_order=d(t, -240), stockist_since=d(t, -240),
             exclusive="No", active="Yes",
             notes="Reorders candles every six weeks."),
        dict(name="Blue Ridge Gift Co", kind="Gift shop",
             contact="Hal Winters", email="orders@brgift.co",
             phone="828 555 0198", city="Black Mountain NC",
             terms="50% deposit, 50% on despatch", discount=0.45,
             min_order=200.00, first_order=d(t, -150),
             stockist_since=d(t, -150), exclusive="No", active="Yes",
             notes="Takes 45% rather than 50 - small orders."),
        dict(name="Pinecone Boutique", kind="Boutique",
             contact="Rae Olsen", email="rae@pineconeboutique.com",
             phone="828 555 0177", city="Asheville NC",
             terms="Net 14", discount=0.50, min_order=250.00,
             first_order=d(t, -90), stockist_since=d(t, -90),
             exclusive="No", active="Yes"),
        dict(name="Folk Art Gallery", kind="Gallery",
             contact="Marianne Shaw", email="hello@folkartgallery.org",
             phone="828 555 0121", city="Asheville NC",
             terms="Consignment", discount=0.40, min_order=0.00,
             first_order=d(t, -300), stockist_since=d(t, -300),
             exclusive="No", active="No",
             notes="Consignment, and gone quiet. Chase or drop."),
        dict(name="Hearthside Subscription Box",
             kind="Subscription box", contact="Toby Lin",
             email="buying@hearthsidebox.com", phone="404 555 0160",
             city="Atlanta GA", terms="Payment with order",
             discount=0.55, min_order=600.00, first_order=d(t, -40),
             stockist_since=d(t, -40), exclusive="No", active="Yes",
             notes="55% off, 200 units at a time. Volume only "
                   "works if the maths works - check the margin."),
    ]

    s.wholesale_orders = [
        dict(date=d(t, -240), customer=wcid(1), po_ref="MM-1041",
             product=pid(6), qty=24, due=d(t, -226),
             despatched=d(t, -228), carriage=18.00,
             carriage_cost=16.40, status="Paid", paid="Yes"),
        dict(date=d(t, -150), customer=wcid(2), po_ref="BRG-220",
             product=pid(7), qty=36, due=d(t, -136),
             despatched=d(t, -138), carriage=14.00,
             carriage_cost=12.80, status="Paid", paid="Yes"),
        dict(date=d(t, -90), customer=wcid(3), po_ref="PB-0012",
             product=pid(1), qty=12, due=d(t, -76),
             despatched=d(t, -78), carriage=12.00,
             carriage_cost=11.20, status="Paid", paid="Yes",
             notes="Sold through in three weeks - and still the worst "
                   "order in the book. Half of 78 does not cover a "
                   "beanie. The margin column said so first."),
        dict(date=d(t, -62), customer=wcid(1), po_ref="MM-1108",
             product=pid(6), qty=24, due=d(t, -48),
             despatched=d(t, -50), carriage=18.00,
             carriage_cost=16.40, status="Paid", paid="Yes"),
        dict(date=d(t, -40), customer=wcid(5), po_ref="HSB-A417",
             product=pid(7), qty=200, price_input=7.60,
             due=d(t, -12), despatched=d(t, -14), carriage=48.00,
             carriage_cost=52.00, status="Paid", paid="Yes",
             notes="Price typed in by hand at 7.60 because the "
                   "55% discount off 19 would have been 8.55 and "
                   "they would not go above 7.60. Check the "
                   "margin column - this is the order to watch."),
        dict(date=d(t, -34), customer=wcid(3), po_ref="PB-0019",
             product=pid(5), qty=8, due=d(t, -20),
             despatched=d(t, -22), carriage=9.00, carriage_cost=8.40,
             status="Invoiced", paid="No",
             notes="Net 14 and now overdue. Chase it."),
        dict(date=d(t, -21), customer=wcid(2), po_ref="BRG-268",
             product=pid(6), qty=18, due=d(t, -7),
             despatched=d(t, -9), carriage=14.00,
             carriage_cost=13.60, status="Invoiced", paid="No"),
        dict(date=d(t, -12), customer=wcid(1), po_ref="MM-1155",
             product=pid(1), qty=10, due=d(t, 2), carriage=12.00,
             carriage_cost=11.20, status="In production", paid="No",
             notes="Ten beanies, and the last trade order of them at "
                   "this price. Candles wholesale, crochet does not."),
        dict(date=d(t, -6), customer=wcid(5), po_ref="HSB-A512",
             product=pid(7), qty=200, price_input=7.60,
             due=d(t, 22), carriage=48.00, carriage_cost=52.00,
             status="Confirmed", paid="No",
             notes="Second box order. Same price, same question."),
        dict(date=d(t, -2), customer=wcid(4), po_ref="FAG-77",
             product=pid(4), qty=6, due=d(t, 26), carriage=0.00,
             carriage_cost=0.00, status="Enquiry", paid="No",
             notes="Consignment enquiry - they pay when it sells, "
                   "if it sells."),
    ]

    # -------------------------------------------------------- sales --
    # Five months of orders: mostly Etsy, three fairs, a handful of
    # direct sales, one breakage and one gift.  Postage charged and
    # postage paid are both recorded, because the difference between
    # them is real money and most sellers never add it up.
    def sale(days: int, product: int, qty: int, channel: str,
             price: float, **kw: Any) -> Dict[str, Any]:
        row: Dict[str, Any] = dict(
            date=d(t, -days), product=pid(product), qty=qty,
            channel=channel, unit_price=price,
            status=kw.pop("status", "Delivered"))
        row.update(kw)
        return row

    fair1 = "Asheville Spring Market"
    fair2 = "River Arts Night Market"
    fair3 = "Holiday Pop-up at The Mill"

    s.sales = [
        sale(1, 8, 1, "Own website", 195.00, ship_charged=8.95,
             ship_paid=9.85, status="Awaiting payment",
             customer="D. Fairweather",
             notes="Invoice sent, not paid yet."),
        sale(2, 6, 1, "Etsy", 32.00, ship_charged=8.95, ship_paid=9.85,
             status="Paid", customer="R. Mendez"),
        sale(3, 12, 1, "Etsy", 8.50, status="Delivered",
             customer="S. Udall",
             notes="Instant download. No postage, no packing, no "
                   "hours after the first sale."),
        sale(4, 4, 1, "Etsy", 58.00, ship_charged=4.95, ship_paid=4.25,
             fees_input=8.12, customer="J. Whitcomb",
             notes="Fee typed in from the statement - compare it "
                   "with the estimate in the next column."),
        sale(5, 7, 2, "Etsy", 19.00, ship_charged=6.95, ship_paid=7.10,
             customer="A. Fenwick"),
        sale(6, 1, 1, "Etsy", 78.00, discount=7.80, promo="WELCOME10",
             ship_charged=5.95, ship_paid=5.40, customer="N. Brodie"),
        sale(8, 6, 2, "Craft fair", 32.00, event=fair3,
             customer="Walk-in"),
        sale(8, 7, 3, "Craft fair", 19.00, event=fair3,
             customer="Walk-in"),
        sale(8, 1, 1, "Craft fair", 78.00, event=fair3,
             customer="Walk-in"),
        sale(8, 5, 1, "Craft fair", 52.00, event=fair3,
             customer="Walk-in"),
        sale(9, 13, 1, "Etsy", 9.50, customer="M. Kaur"),
        sale(10, 10, 1, "Own website", 135.00, ship_charged=16.00,
             ship_paid=18.60, customer="Hearthside Interiors",
             notes="Flat rate postage. The courier cost 2.60 more "
                   "than the buyer paid."),
        sale(11, 7, 1, "Other", 0.00, status="Gift",
             customer="Blue Ridge Living",
             notes="Sample to a magazine. Costed, not charged - "
                   "which is exactly how it should appear."),
        sale(12, 2, 1, "Etsy", 75.00, ship_charged=6.95,
             ship_paid=7.10, customer="P. Lasseter"),
        sale(13, 5, 1, "Etsy", 52.00, ship_charged=4.95,
             ship_paid=4.25, customer="C. Ngata"),
        sale(15, 9, 1, "Own website", 148.00, ship_charged=8.95,
             ship_paid=9.85, customer="T. Alcorn"),
        sale(16, 12, 2, "Etsy", 8.50, customer="E. Rossiter"),
        sale(18, 3, 1, "Etsy", 78.00, ship_charged=6.95,
             ship_paid=7.10, customer="H. Delacroix"),
        sale(19, 6, 1, "Instagram / direct", 32.00, ship_charged=8.95,
             ship_paid=9.85, customer="L. Byrne",
             notes="A payment link. Cheapest fees you pay."),
        sale(21, 6, 1, "Etsy", 32.00, ship_charged=8.95,
             ship_paid=9.85, refund=40.95, status="Refunded",
             customer="G. Hollis",
             notes="Jar cracked in the post. Refunded in full, "
                   "postage as well - and the wax is gone too."),
        sale(23, 8, 1, "Etsy", 195.00, ship_charged=8.95,
             ship_paid=9.85, customer="V. Amadi"),
        sale(24, 1, 1, "Etsy", 78.00, ship_charged=5.95,
             ship_paid=5.40, fees_input=21.20, customer="B. Soriano",
             notes="Came through offsite ads, so the fee was 15% "
                   "on top. One order in five looks like this."),
        sale(26, 7, 2, "Own website", 19.00, ship_charged=6.95,
             ship_paid=7.10, customer="F. Okonkwo"),
        sale(28, 13, 1, "Etsy", 9.50, customer="I. Petrova"),
        sale(30, 5, 1, "Etsy", 52.00, ship_charged=4.95,
             ship_paid=4.25, customer="D. Mwangi"),
        sale(31, 2, 1, "Etsy", 75.00, ship_charged=6.95,
             ship_paid=7.10, customer="K. Lindqvist"),
        sale(33, 12, 1, "Etsy", 8.50, customer="O. Reyes"),
        sale(35, 7, 4, "Craft fair", 19.00, event=fair2,
             customer="Walk-in"),
        sale(35, 6, 2, "Craft fair", 32.00, event=fair2,
             customer="Walk-in"),
        sale(35, 1, 2, "Craft fair", 65.00, event=fair2,
             customer="Walk-in"),
        sale(38, 3, 1, "Etsy", 62.00, ship_charged=6.95,
             ship_paid=7.10, customer="W. Castellano"),
        sale(40, 9, 1, "Instagram / direct", 148.00, ship_charged=8.95,
             ship_paid=9.85, customer="K. Osei"),
        sale(42, 12, 2, "Etsy", 8.50, customer="R. Bannerjee"),
        sale(44, 4, 1, "Etsy", 52.00, ship_charged=4.95,
             ship_paid=4.25, customer="Y. Tanaka",
             notes="Sold at the old price, before silver moved."),
        sale(46, 10, 1, "Own website", 135.00, ship_charged=16.00,
             ship_paid=18.60, customer="M. Quintero"),
        sale(48, 6, 1, "Etsy", 32.00, ship_charged=8.95,
             ship_paid=9.85, customer="S. Ellery"),
        sale(52, 1, 1, "Etsy", 65.00, ship_charged=5.95,
             ship_paid=5.40, customer="J. Pembroke",
             notes="Before the autumn price rise."),
        sale(55, 13, 1, "Etsy", 6.50, customer="A. Laurent"),
        sale(58, 5, 1, "Etsy", 44.00, ship_charged=4.95,
             ship_paid=4.25, refund=8.00, status="Partly refunded",
             customer="N. Oyelaran",
             notes="Chain was shorter than listed. Eight dollars "
                   "back rather than a return."),
        sale(60, 2, 1, "Etsy", 75.00, ship_charged=6.95,
             ship_paid=7.10, customer="T. Halloran"),
        sale(63, 7, 1, "Etsy", 19.00, ship_charged=6.95,
             ship_paid=7.10, customer="E. Vasquez"),
        sale(66, 12, 1, "Etsy", 8.50, customer="L. Marchetti"),
        sale(70, 5, 2, "Craft fair", 44.00, event=fair1,
             customer="Walk-in"),
        sale(70, 2, 2, "Craft fair", 75.00, event=fair1,
             customer="Walk-in"),
        sale(70, 7, 3, "Craft fair", 19.00, event=fair1,
             customer="Walk-in"),
        sale(70, 1, 1, "Craft fair", 65.00, event=fair1,
             customer="Walk-in"),
        sale(74, 8, 1, "Etsy", 195.00, ship_charged=8.95,
             ship_paid=9.85, customer="P. Ashworth"),
        sale(78, 12, 3, "Etsy", 8.50, customer="Three buyers",
             notes="A pattern round-up blog linked it."),
        sale(82, 4, 1, "Etsy", 52.00, ship_charged=4.95,
             ship_paid=4.25, customer="R. Stenberg"),
        sale(88, 6, 1, "Etsy", 32.00, discount=6.40, promo="AUTUMN20",
             ship_charged=8.95, ship_paid=9.85, customer="J. Farrow",
             notes="Twenty percent off. Look at what was left."),
        sale(92, 3, 1, "Own website", 62.00, ship_charged=6.95,
             ship_paid=7.10, customer="U. Bergstrom"),
        sale(96, 9, 1, "Etsy", 128.00, ship_charged=8.95,
             ship_paid=9.85, customer="C. Devereux"),
        sale(104, 5, 1, "Etsy", 44.00, ship_charged=24.50,
             ship_paid=24.50, customer="A. Kowalski",
             notes="International. Customs form, twelve days, "
                   "and the fees are charged on the postage too."),
        sale(112, 1, 1, "Etsy", 65.00, ship_charged=5.95,
             ship_paid=5.40, customer="M. Thorne"),
        sale(120, 7, 2, "Etsy", 19.00, ship_charged=6.95,
             ship_paid=7.10, customer="S. Nakamura"),
        sale(130, 12, 1, "Etsy", 7.50, customer="B. Ferreira",
             notes="Pattern was 7.50 before the price test."),
        sale(140, 10, 1, "Own website", 118.00, ship_charged=16.00,
             ship_paid=18.60, customer="Cottonwood Interiors"),
        sale(150, 2, 1, "Etsy", 75.00, ship_charged=6.95,
             ship_paid=7.10, customer="G. Vuković"),
    ]

    return s
