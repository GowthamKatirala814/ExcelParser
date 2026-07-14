"""
Builds a size-reduced projection of the raw workbook JSON for the LLM call
only - this is never what's stored/downloaded as the Raw JSON, which stays
full-fidelity. Kept: value, fill color (the signal most likely to carry
business meaning - e.g. a status legend), comment, formula, merges, and -
critically - each cell's own reference ("B12"), so the LLM can locate every
value exactly and the merge/validation step downstream can look any of it
back up. Dropped for this payload only:

  - Cells that are genuinely blank (no value, color, comment, formula, or
    merge) - they carry no information for structuring.
  - Border style and font color - visual-only noise for this task; still
    present in the full raw.json download.
  - The verbose per-color object (resolved/hex/theme_ref/reason_unresolved)
    collapses to a short string.
  - The `structured.tables[].rows[].cells` dense grid, which duplicates
    every cell already present in the flat `rows` list - only the table's
    positional metadata (row range, header row, date row, column index) is
    kept, since the LLM can look up any referenced cell by its `ref`.
"""


def _compact_color(color):
    if not color:
        return None
    if color.get("resolved") and color.get("hex"):
        return f"#{color['hex']}"
    if color.get("theme_ref"):
        return f"unresolved({color['theme_ref']})"
    if color.get("reason_unresolved"):
        return f"unresolved({color['reason_unresolved']})"
    return None


def _compact_cell(row):
    compact = {"ref": row.get("cell_ref")}

    if row.get("value") is not None:
        compact["value"] = row["value"]

    fill = _compact_color(row.get("color"))
    if fill:
        compact["fill"] = fill

    if row.get("comment"):
        compact["comment"] = row["comment"]
    if row.get("formula"):
        compact["formula"] = row["formula"]
    if row.get("merged_with"):
        compact["merged_with"] = row["merged_with"]
    if row.get("human_label"):
        compact["human_label"] = row["human_label"]

    return compact


def _compact_tables(structured):
    tables = []
    for table in (structured or {}).get("tables", []) or []:
        header = table.get("header")
        tables.append(
            {
                "row_range": table.get("row_range"),
                "header_row": header.get("row") if header else None,
                "header_signals": header.get("signals") if header else None,
                "date_row": table.get("date_row"),
                "column_index": table.get("column_index"),
                "ambiguous": table.get("ambiguous"),
                "ambiguous_reason": table.get("ambiguous_reason"),
            }
        )
    return tables


def build_llm_payload(raw_json):
    sheets_payload = {}
    for sheet_name, sheet in raw_json.get("sheets", {}).items():
        cells = []
        for row in sheet.get("rows", []) or []:
            compact = _compact_cell(row)
            if len(compact) > 1:  # more than just "ref" => actually meaningful
                cells.append(compact)

        sheets_payload[sheet_name] = {
            "total_cells": sheet.get("total_cells"),
            "colored_cells": sheet.get("colored_cells"),
            "blank_cells": sheet.get("blank_cells"),
            "color_inventory": sheet.get("color_inventory"),
            "detected_tables": _compact_tables(sheet.get("structured")),
            "cells": cells,
        }

    return {"workbook": raw_json.get("workbook"), "sheets": sheets_payload}
