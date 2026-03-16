"""
Package-wide layout constants for Excel rendering.

These values control the whitespace inserted between adjacent tables
when writing a sheet. Modify here to change spacing globally.
"""

TABLE_MARGIN_ROWS: int = 1
"""Number of blank rows inserted between table rows."""

TABLE_MARGIN_COLS: int = 1
"""Number of blank columns inserted between tables within the same row."""