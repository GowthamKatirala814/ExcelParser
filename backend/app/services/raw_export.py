"""
Assembles the "Raw Excel Parsing JSON" from what's already stored in Mongo
(workbook doc + one extraction doc per sheet). Pure reshaping - no
re-parsing of the .xlsx file, so it can be produced or re-downloaded at any
time without reprocessing the workbook. Available at both workbook level
(all sheets together) and single-sheet level (exactly one sheet's data) -
every sheet's raw JSON is independent of every other sheet's.
"""


def _workbook_meta(workbook_doc):
    return {
        "filename": workbook_doc.get("filename"),
        "size_bytes": workbook_doc.get("size"),
        "uploaded_at": workbook_doc.get("uploaded_at"),
        "extracted_at": workbook_doc.get("extracted_at"),
        "sheet_names": workbook_doc.get("sheet_names"),
        "sheet_dimensions": workbook_doc.get("sheet_dimensions"),
    }


def _sheet_payload(extraction):
    return {
        "processing_time_seconds": extraction.get("processing_time_seconds"),
        "total_cells": extraction.get("total_cells"),
        "colored_cells": extraction.get("colored_cells"),
        "blank_cells": extraction.get("blank_cells"),
        "unresolved_cells": extraction.get("unresolved_cells"),
        "ambiguous_cells": extraction.get("ambiguous_cells"),
        "color_inventory": extraction.get("color_inventory"),
        "rows": extraction.get("rows"),
        "structured": extraction.get("structured"),
    }


def build_raw_workbook_json(workbook_doc, extractions):
    return {
        "workbook": _workbook_meta(workbook_doc),
        "sheets": {extraction["sheet_name"]: _sheet_payload(extraction) for extraction in extractions},
    }


def build_raw_sheet_json(workbook_doc, extraction):
    """Same shape as build_raw_workbook_json, but with exactly one sheet - so
    an LLM call or a download scoped to one sheet never carries any other
    sheet's data."""
    return {
        "workbook": _workbook_meta(workbook_doc),
        "sheets": {extraction["sheet_name"]: _sheet_payload(extraction)},
    }
