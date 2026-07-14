"""
Per-workbook, per-sheet extraction. Sheets are always processed independently
and never merged together. Every statistic reported here is a direct count of
something that was actually read from the file - nothing is estimated.
"""
import datetime
import time

import openpyxl

from .cf_resolver import resolve_conditional_format_color
from .color_resolver import resolve_color_object, resolve_direct_fill
from .structure import build_tables
from .theme_resolver import load_theme_palette

MAX_INVENTORY_EXAMPLES = 5
_BORDER_SIDES = ("top", "bottom", "left", "right")


def _serialize_value(value):
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _resolve_cell_color(cell, cell_value, worksheet, workbook, theme_palette, theme_error):
    """
    Direct fill first; conditional formatting only if there is no direct
    pattern fill defined on the cell.
    """
    direct = resolve_direct_fill(cell, theme_palette, theme_error)
    if direct is not None:
        direct["source"] = "direct_fill"
        return direct

    cf_result = resolve_conditional_format_color(
        cell, cell_value, worksheet, workbook, theme_palette, theme_error
    )
    if cf_result is not None:
        cf_result["source"] = "conditional_formatting"
        return cf_result

    return None  # no fill at all: blank / no-fill cell


def _resolve_font_color(cell, theme_palette, theme_error):
    font = getattr(cell, "font", None)
    color = getattr(font, "color", None) if font is not None else None
    if color is None:
        return None
    return resolve_color_object(color, theme_palette, theme_error)


def _resolve_borders(cell, theme_palette, theme_error):
    """Returns {side: {"style": str, "color": color-fact}} for sides that actually have a border."""
    border = getattr(cell, "border", None)
    if border is None:
        return {}
    result = {}
    for side in _BORDER_SIDES:
        side_border = getattr(border, side, None)
        style = getattr(side_border, "style", None) if side_border is not None else None
        if style is None:
            continue
        result[side] = {
            "style": style,
            "color": resolve_color_object(
                getattr(side_border, "color", None), theme_palette, theme_error
            ),
        }
    return result


def _resolve_comment(cell):
    comment = getattr(cell, "comment", None)
    if comment is None:
        return None
    return {"text": comment.text, "author": comment.author}


def _resolve_formula(formula_cell):
    """formula_cell comes from a parallel data_only=False load; None if not applicable."""
    if formula_cell is None or formula_cell.data_type != "f":
        return None
    return formula_cell.value


def _inventory_key(color_result):
    if color_result is None:
        return None
    if color_result.get("resolved"):
        return f"#{color_result['hex']}"
    if color_result.get("theme_ref"):
        return color_result["theme_ref"]
    return f"unresolved:{color_result.get('reason_unresolved', 'unknown')}"


def _build_merge_maps(worksheet):
    anchor_to_members = {}
    member_to_anchor = {}
    for merged_range in worksheet.merged_cells.ranges:
        cells = list(merged_range.cells)
        if not cells:
            continue
        anchor_row, anchor_col = min(cells)
        anchor_cell = worksheet.cell(row=anchor_row, column=anchor_col)
        anchor_ref = anchor_cell.coordinate
        members = []
        for row, col in cells:
            if (row, col) == (anchor_row, anchor_col):
                continue
            ref = worksheet.cell(row=row, column=col).coordinate
            members.append(ref)
            member_to_anchor[ref] = anchor_ref
        anchor_to_members[anchor_ref] = members
    return anchor_to_members, member_to_anchor


def extract_sheet(workbook, worksheet, theme_palette=None, theme_error=None, formula_worksheet=None):
    anchor_to_members, member_to_anchor = _build_merge_maps(worksheet)

    total_cells = 0
    colored_cells = 0
    blank_cells = 0
    unresolved_cells = 0
    ambiguous_cells = 0

    inventory = {}  # key -> {hex_or_theme_ref, resolved, cell_count, example_refs, reason_unresolved}
    rows = []

    for row in worksheet.iter_rows():
        for cell in row:
            ref = cell.coordinate
            if cell.value is None and ref in member_to_anchor:
                # non-anchor merged cell: still scanned for coverage stats,
                # but not reported as its own row (its value belongs to the anchor).
                total_cells += 1
                raw_value = None
            else:
                total_cells += 1
                raw_value = cell.value

            color_result = _resolve_cell_color(
                cell, raw_value, worksheet, workbook, theme_palette, theme_error
            )

            if color_result is None:
                blank_cells += 1
            elif color_result.get("resolved"):
                colored_cells += 1
            else:
                unresolved_cells += 1
                if color_result.get("ambiguous"):
                    ambiguous_cells += 1

            if color_result is not None:
                key = _inventory_key(color_result)
                entry = inventory.setdefault(
                    key,
                    {
                        "key": key,
                        "hex_or_theme_ref": color_result.get("hex")
                        and f"#{color_result['hex']}"
                        or color_result.get("theme_ref")
                        or key,
                        "resolved": bool(color_result.get("resolved")),
                        "reason_unresolved": color_result.get("reason_unresolved"),
                        "cell_count": 0,
                        "example_refs": [],
                    },
                )
                entry["cell_count"] += 1
                if len(entry["example_refs"]) < MAX_INVENTORY_EXAMPLES:
                    entry["example_refs"].append(ref)

            if ref in member_to_anchor:
                # Skip emitting a separate row for non-anchor merged members.
                continue

            formula_cell = formula_worksheet[ref] if formula_worksheet is not None else None

            rows.append(
                {
                    "cell_ref": ref,
                    "value": _serialize_value(cell.value),
                    "color": {
                        "resolved": bool(color_result and color_result.get("resolved")),
                        "hex": color_result.get("hex") if color_result else None,
                        "theme_ref": color_result.get("theme_ref") if color_result else None,
                        "source": color_result.get("source") if color_result else None,
                        "ambiguous": bool(color_result and color_result.get("ambiguous")),
                        "reason_unresolved": color_result.get("reason_unresolved")
                        if color_result
                        else None,
                    },
                    "font_color": _resolve_font_color(cell, theme_palette, theme_error),
                    "borders": _resolve_borders(cell, theme_palette, theme_error),
                    "comment": _resolve_comment(cell),
                    "formula": _resolve_formula(formula_cell),
                    "merged_with": anchor_to_members.get(ref, []),
                    "human_label": None,
                }
            )

    color_inventory = sorted(inventory.values(), key=lambda e: e["cell_count"], reverse=True)

    raw_cells = {
        "total_cells": total_cells,
        "colored_cells": colored_cells,
        "blank_cells": blank_cells,
        "unresolved_cells": unresolved_cells,
        "ambiguous_cells": ambiguous_cells,
        "color_inventory": color_inventory,
        "rows": rows,
    }

    raw_by_ref = {row["cell_ref"]: row for row in rows}
    structured = {"tables": build_tables(worksheet, raw_by_ref, member_to_anchor)}

    return {
        # Flat fields kept for backward compatibility with existing callers
        # (routers, reports) - never removed or degraded.
        "total_cells": total_cells,
        "colored_cells": colored_cells,
        "blank_cells": blank_cells,
        "unresolved_cells": unresolved_cells,
        "ambiguous_cells": ambiguous_cells,
        "color_inventory": color_inventory,
        "rows": rows,
        # New: explicit raw/structured split, side by side, per the
        # workbook -> sheets -> tables -> rows -> cells hierarchy.
        "raw_cells": raw_cells,
        "structured": structured,
    }


def extract_workbook(file_path):
    """
    Returns (per_sheet_results: dict[sheet_name -> extraction result with
    processing_time_seconds], workbook_total_time_seconds).
    """
    workbook_start = time.perf_counter()
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    # A second, parallel load with formulas intact (data_only=True replaces
    # formula cells with their cached computed value, so formula text can
    # only be read from a workbook opened without that flag).
    formula_workbook = openpyxl.load_workbook(file_path, data_only=False)
    theme_palette, theme_error = load_theme_palette(file_path)

    per_sheet_results = {}
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        formula_worksheet = formula_workbook[sheet_name]
        sheet_start = time.perf_counter()
        result = extract_sheet(workbook, worksheet, theme_palette, theme_error, formula_worksheet)
        result["processing_time_seconds"] = time.perf_counter() - sheet_start
        per_sheet_results[sheet_name] = result

    workbook_total_time = time.perf_counter() - workbook_start
    return per_sheet_results, workbook_total_time


def read_sheet_dimensions(file_path):
    """Used right after upload, before extraction, for immediate metadata."""
    workbook = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
    dimensions = {}
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        dimensions[sheet_name] = {
            "rows": worksheet.max_row or 0,
            "columns": worksheet.max_column or 0,
        }
    workbook.close()
    return workbook.sheetnames, dimensions
