"""Ultimate Airbnb & Short-Term Rental Management Spreadsheet.

Build the workbook with::

    python -m airbnb_rental_spreadsheet -o "output/Airbnb.xlsx"

or from Python::

    from airbnb_rental_spreadsheet import build_workbook
    build_workbook("Airbnb Rental Management Spreadsheet.xlsx")
"""

from .build import build_workbook
from .config import (AUTHOR, LISTING_TITLE, PRODUCT_NAME,
                     PRODUCT_SUBTITLE, VERSION, Config)

__all__ = ["build_workbook", "Config", "AUTHOR", "PRODUCT_NAME",
           "PRODUCT_SUBTITLE", "LISTING_TITLE", "VERSION"]
__version__ = VERSION
