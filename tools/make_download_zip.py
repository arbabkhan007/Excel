"""Pack a generator into a single zip people can download and run.

    python tools/make_download_zip.py
    python tools/make_download_zip.py --product social

Writes `download/<product>-source.zip` containing the generator, the
quality tools, the documentation and a `requirements.txt`, all under one
folder so it unpacks tidily.

The vendored copy of XlsxWriter in this repository is deliberately left
out: the zip is meant to be unpacked anywhere and run against a normal
`pip install XlsxWriter`.  Timestamps are fixed so rebuilding the zip
without changing any source produces exactly the same bytes.
"""

from __future__ import annotations

import argparse
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#: Same date for every entry, so the zip only changes when the code does.
STAMP = (2026, 1, 1, 0, 0, 0)

SHARED_TOOLS = ["tools/recalc_check.py", "tools/export_source_md.py",
                "tools/make_download_zip.py"]

REQUIREMENTS = """\
# The generator needs one package.
XlsxWriter>=3.0

# Only for the optional checks in tools/
# openpyxl>=3.1     the validator and the sample data check
# formulas>=1.3     recalc_check.py
"""

PET_HOW_TO_RUN = """\
Pet Sitting & Dog Walking Business Spreadsheet
==============================================

    pip install XlsxWriter
    python build_pet_sitting_workbook.py

The workbook appears in:

    output/Pet Sitting Business Spreadsheet.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --business "My Name"    your business name on the cover and invoices
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --unit miles            mileage in miles instead of km
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 1000             more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

README.md explains how it is built, what is on each of the 27 tabs and
how to change things.  SOURCE_CODE.md is the whole program in one
readable document.
"""

SOCIAL_HOW_TO_RUN = """\
Social Media Content Calendar & Analytics Planner
=================================================

    pip install XlsxWriter
    python build_social_media_workbook.py

The workbook appears in:

    output/Social Media Content Planner.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --brand "My Brand"      your brand name on the cover and dashboards
    --year 2027             the year the 12 month calendar is built for
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 600              more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

Start on the "Start Here" tab inside the workbook.  README_SOCIAL.md
explains how it is built and what is on each of the 43 tabs;
SOURCE_CODE_SOCIAL.md is the whole program in one readable document.
"""

PRODUCTS = {
    "pet": {
        "folder": "pet-sitting-spreadsheet",
        "output": os.path.join("download",
                               "pet-sitting-spreadsheet-source.zip"),
        "files": ["build_pet_sitting_workbook.py", "README.md",
                  "SOURCE_CODE.md"],
        "trees": ["pet_sitting_spreadsheet"],
        "tools": ["tools/validate_workbook.py", "tools/check_sample_data.py"]
                 + SHARED_TOOLS,
        "how_to_run": PET_HOW_TO_RUN,
    },
    "social": {
        "folder": "social-media-content-planner",
        "output": os.path.join(
            "download", "social-media-content-planner-source.zip"),
        "files": ["build_social_media_workbook.py", "README_SOCIAL.md",
                  "SOURCE_CODE_SOCIAL.md"],
        "trees": ["social_media_spreadsheet"],
        "tools": ["tools/validate_social_workbook.py",
                  "tools/check_social_sample.py"] + SHARED_TOOLS,
        "how_to_run": SOCIAL_HOW_TO_RUN,
    },
}


def gather(product: dict) -> list:
    """`[(archive name, text)]` for everything that goes in the zip."""
    items = []
    for name in product["files"]:
        with open(os.path.join(ROOT, name), encoding="utf-8") as handle:
            items.append((name, handle.read()))
    for tree in product["trees"]:
        for folder, dirs, names in os.walk(os.path.join(ROOT, tree)):
            dirs[:] = sorted(d for d in dirs if d != "__pycache__")
            for name in sorted(names):
                if not name.endswith(".py"):
                    continue
                full = os.path.join(folder, name)
                rel = os.path.relpath(full, ROOT).replace(os.sep, "/")
                with open(full, encoding="utf-8") as handle:
                    items.append((rel, handle.read()))
    for name in product["tools"]:
        with open(os.path.join(ROOT, name), encoding="utf-8") as handle:
            items.append((name, handle.read()))
    items.append(("requirements.txt", REQUIREMENTS))
    items.append(("HOW TO RUN.txt", product["how_to_run"]))
    return items


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--product", choices=sorted(PRODUCTS),
                        default="pet",
                        help="which generator to pack (default: pet)")
    parser.add_argument("-o", "--output", default=None)
    args = parser.parse_args(argv)

    product = PRODUCTS[args.product]
    out = args.output or product["output"]
    folder = os.path.dirname(os.path.abspath(out))
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)

    items = gather(product)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, text in items:
            info = zipfile.ZipInfo(f"{product['folder']}/{name}",
                                   date_time=STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            zf.writestr(info, text)

    size = os.path.getsize(out)
    lines = sum(len(text.splitlines()) for name, text in items
                if name.endswith(".py"))
    print(f"{out}  ({len(items)} files, {lines:,} lines of Python, "
          f"{size / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
