#!/usr/bin/env python3
"""Validate the Airbnb product's example data before it is built.

Two checks, both of which have caught real bugs in this repo:

1. every key in a sample row is an *input* column on that register -
   writing a sample value into a formula column silently destroys the
   formula for that row;
2. every value that feeds a drop-down exists in that drop-down's list -
   otherwise Excel shows the cell with a validation error the moment a
   buyer clicks it.

    python3 tools/check_airbnb_sample.py
"""

from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from airbnb_rental_spreadsheet.config import Config          # noqa: E402
from airbnb_rental_spreadsheet.build import REGISTER_MODULES  # noqa: E402
from airbnb_rental_spreadsheet.sheets.settings import LISTS   # noqa: E402

#: lists that are built as a table on Settings rather than a plain column
TABLE_LISTS = {
    "PlatformList": lambda cfg: [name for name, *_ in cfg.platform_table],
}

#: names that resolve to a live column of IDs on another register
ID_LISTS = {"PropertyList", "BookingList", "GuestList", "VendorList",
            "SupplyList", "CleanTaskList", "MaintenanceList",
            "PropertyNameList", "GuestNameList", "VendorNameList",
            "SupplyNameList"}

ID_PATTERN = re.compile(r"^[A-Z]{3}-\d+$")


def main() -> int:
    cfg = Config()
    lists = {name: set(values) for name, values in LISTS}
    for name, builder in TABLE_LISTS.items():
        lists[name] = set(builder(cfg))

    problems = 0
    total_rows = 0
    for module in REGISTER_MODULES:
        spec = module.spec(cfg)
        inputs = {c.key for c in spec.columns if not c.formula}
        by_key = {c.key: c for c in spec.columns}
        rows = spec.samples
        total_rows += len(rows)

        if len(rows) > spec.rows:
            problems += 1
            print(f"{spec.name}: {len(rows)} sample rows will not fit "
                  f"in {spec.rows}")

        for index, row in enumerate(rows):
            extra = sorted(set(row) - inputs)
            if extra:
                problems += 1
                print(f"{spec.name} row {index}: not input columns "
                      f"-> {extra}")
            for key, value in row.items():
                column = by_key.get(key)
                if column is None or not column.validate:
                    continue
                source = column.validate.get("source")
                if not isinstance(source, str) or not source.startswith("="):
                    continue
                name = source.lstrip("=")
                if name in lists:
                    if value != "" and value not in lists[name]:
                        problems += 1
                        print(f"{spec.name}.{key} row {index}: "
                              f"{value!r} is not in {name}")
                elif name in ID_LISTS:
                    if value != "" and not ID_PATTERN.match(str(value)):
                        problems += 1
                        print(f"{spec.name}.{key} row {index}: "
                              f"{value!r} does not look like an ID "
                              f"({name})")

        print(f"  {spec.name:26s} {len(rows):4d} rows, "
              f"{len(inputs)} input columns")

    print(f"\n{total_rows} sample rows checked, {problems} problems")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
