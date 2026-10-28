
"""
Reusable UI helper functions for the Partlist Validation app.
"""


def selected_row_index(selection):
    """
    Return the selected row index from a Streamlit dataframe event.

    Supports Streamlit selection objects with:
    - selection.rows
    - selection.cells

    Returns:
        int: Selected row index, or None if no row is selected.
    """
    if selection is None:
        return None

    # Handle Streamlit dataframe row selection.
    rows = getattr(selection, "rows", None)

    if rows:
        return int(rows[0])

    # Handle Streamlit dataframe cell selection.
    cells = getattr(selection, "cells", None)

    if cells:
        first_cell = cells[0]

        if isinstance(first_cell, (tuple, list)) and first_cell:
            return int(first_cell[0])

    return None
