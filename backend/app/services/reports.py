"""
Report generation. Every generated report bundles the raw extracted data
together with the coverage/accuracy statistics it was computed from - never
data alone, and never statistics without the data that produced them.

CSV and JSON only - no XLSX export.
"""
import csv
import io
import json
import zipfile

ROW_COLUMNS = [
    "cell_ref",
    "value",
    "color_resolved",
    "color_hex",
    "color_theme_ref",
    "color_source",
    "color_ambiguous",
    "color_reason_unresolved",
    "merged_with",
    "human_label",
]


def _row_to_flat(row):
    color = row.get("color", {}) or {}
    return {
        "cell_ref": row.get("cell_ref"),
        "value": row.get("value"),
        "color_resolved": color.get("resolved"),
        "color_hex": color.get("hex"),
        "color_theme_ref": color.get("theme_ref"),
        "color_source": color.get("source"),
        "color_ambiguous": color.get("ambiguous"),
        "color_reason_unresolved": color.get("reason_unresolved"),
        "merged_with": ", ".join(row.get("merged_with") or []),
        "human_label": row.get("human_label"),
    }


def _sheet_summary(extraction):
    total = extraction.get("total_cells", 0)
    return {
        "sheet_name": extraction.get("sheet_name"),
        "processing_time_seconds": extraction.get("processing_time_seconds"),
        "total_cells": total,
        "colored_cells": extraction.get("colored_cells", 0),
        "blank_cells": extraction.get("blank_cells", 0),
        "unresolved_cells": extraction.get("unresolved_cells", 0),
        "ambiguous_cells": extraction.get("ambiguous_cells", 0),
        "color_coverage_ratio": (
            extraction.get("colored_cells", 0) / total if total else 0.0
        ),
    }


def _sheet_warnings(extraction):
    """Advisory, human-review warnings derived from the descriptive-only
    extraction passes (embedded images, low-contrast colors). Never affects
    any data or color value - purely surfaced for a human to double-check."""
    warnings = []

    images = extraction.get("embedded_images") or []
    if images:
        locations = ", ".join(img.get("anchor_cell") or "?" for img in images)
        warnings.append(
            f"{len(images)} image(s) detected on this sheet at {locations} - "
            "their content was not read; check manually if they contain data."
        )

    for pair in extraction.get("low_contrast_pairs") or []:
        warnings.append(
            f"Low-contrast colors {pair.get('color_a')} and {pair.get('color_b')} "
            f"(Delta-E {pair.get('delta_e')}) - {pair.get('note')}"
        )

    return warnings


def _sheet_payload(extraction):
    return {
        "summary": _sheet_summary(extraction),
        "warnings": _sheet_warnings(extraction),
        "color_inventory": extraction.get("color_inventory", []),
        "rows": [_row_to_flat(r) for r in extraction.get("rows", [])],
    }


def sheet_report_json(extraction):
    payload = _sheet_payload(extraction)
    return json.dumps(payload, indent=2, default=str).encode("utf-8")


def sheet_report_csv(extraction):
    payload = _sheet_payload(extraction)
    buffer = io.StringIO()

    buffer.write("# Coverage Summary\n")
    summary_writer = csv.writer(buffer)
    for key, value in payload["summary"].items():
        summary_writer.writerow([key, value])

    if payload["warnings"]:
        buffer.write("\n# Warnings (advisory - check manually)\n")
        warn_writer = csv.writer(buffer)
        for warning in payload["warnings"]:
            warn_writer.writerow([warning])

    buffer.write("\n# Color Inventory\n")
    inv_writer = csv.writer(buffer)
    inv_writer.writerow(["key", "hex_or_theme_ref", "resolved", "cell_count", "example_refs", "reason_unresolved"])
    for entry in payload["color_inventory"]:
        inv_writer.writerow(
            [
                entry.get("key"),
                entry.get("hex_or_theme_ref"),
                entry.get("resolved"),
                entry.get("cell_count"),
                ", ".join(entry.get("example_refs", [])),
                entry.get("reason_unresolved"),
            ]
        )

    buffer.write("\n# Extracted Data\n")
    data_writer = csv.DictWriter(buffer, fieldnames=ROW_COLUMNS)
    data_writer.writeheader()
    for row in payload["rows"]:
        data_writer.writerow(row)

    return buffer.getvalue().encode("utf-8")


def _workbook_summary(workbook_doc, extractions):
    per_sheet = [_sheet_summary(e) for e in extractions]
    total_cells = sum(s["total_cells"] for s in per_sheet)
    colored_cells = sum(s["colored_cells"] for s in per_sheet)
    blank_cells = sum(s["blank_cells"] for s in per_sheet)
    unresolved_cells = sum(s["unresolved_cells"] for s in per_sheet)
    ambiguous_cells = sum(s["ambiguous_cells"] for s in per_sheet)
    total_time = sum(s["processing_time_seconds"] or 0 for s in per_sheet)
    return {
        "workbook_id": str(workbook_doc.get("_id")),
        "filename": workbook_doc.get("filename"),
        "sheet_count": len(per_sheet),
        "total_processing_time_seconds": total_time,
        "total_cells": total_cells,
        "colored_cells": colored_cells,
        "blank_cells": blank_cells,
        "unresolved_cells": unresolved_cells,
        "ambiguous_cells": ambiguous_cells,
        "color_coverage_ratio": colored_cells / total_cells if total_cells else 0.0,
        "per_sheet": per_sheet,
    }


def workbook_report_json(workbook_doc, extractions):
    payload = {
        "workbook_summary": _workbook_summary(workbook_doc, extractions),
        "sheets": {e["sheet_name"]: _sheet_payload(e) for e in extractions},
    }
    return json.dumps(payload, indent=2, default=str).encode("utf-8")


def workbook_report_csv_zip(workbook_doc, extractions):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        summary = _workbook_summary(workbook_doc, extractions)
        summary_csv = io.StringIO()
        writer = csv.writer(summary_csv)
        writer.writerow(["metric", "value"])
        for key, value in summary.items():
            if key == "per_sheet":
                continue
            writer.writerow([key, value])
        writer.writerow([])
        writer.writerow(["sheet_name", "processing_time_seconds", "total_cells", "colored_cells", "blank_cells", "unresolved_cells", "ambiguous_cells", "color_coverage_ratio"])
        for s in summary["per_sheet"]:
            writer.writerow([
                s["sheet_name"], s["processing_time_seconds"], s["total_cells"],
                s["colored_cells"], s["blank_cells"], s["unresolved_cells"],
                s["ambiguous_cells"], s["color_coverage_ratio"],
            ])
        zf.writestr("workbook_summary.csv", summary_csv.getvalue())

        for extraction in extractions:
            name = extraction["sheet_name"]
            zf.writestr(f"{name}.csv", sheet_report_csv(extraction))

    return buffer.getvalue()


def workbook_report_json_zip(workbook_doc, extractions):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "workbook_summary.json",
            json.dumps(_workbook_summary(workbook_doc, extractions), indent=2, default=str),
        )
        for extraction in extractions:
            name = extraction["sheet_name"]
            zf.writestr(f"{name}.json", sheet_report_json(extraction))
    return buffer.getvalue()
