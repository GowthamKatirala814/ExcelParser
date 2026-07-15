import asyncio
import datetime
import io
import json
import logging
import uuid
from pathlib import Path

import bson
from bson import ObjectId
from bson.errors import InvalidId
from fastapi import APIRouter, BackgroundTasks, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from .. import config
from ..db import extractions_collection, structured_collection, workbooks_collection
from ..services import reports
from ..services.extraction_validation import validate_extraction_completeness
from ..services.extractor import extract_workbook, read_sheet_dimensions
from ..services.llm_payload import build_llm_payload
from ..services.lossless_extractor import extract_lossless
from ..services.raw_export import build_raw_sheet_json, build_raw_workbook_json
from ..services.structured_assembly import assemble_sheet_structured_output
from ..services.structured_schema_validation import validate_structured_sheet
from ..services.structured_validation import validate_sheet_structured_output
from ..services.structuring_service import StructuringError, generate_structured_json
from ..services.transformation_engine import transform_sheet, transform_workbook

logger = logging.getLogger("app.structuring")

router = APIRouter(prefix="/workbooks", tags=["workbooks"])

# MongoDB rejects any document over 16 MiB. Very large sheets (tens of
# thousands of cells) can still exceed that even after the compact-row
# slimming, so extraction degrades gracefully rather than 500-ing: the dense
# per-cell `structured` grid (rebuildable on demand from the .xlsx via the
# lossless/intermediate endpoints) is dropped first, then `rows` if needed.
_MONGO_DOC_SAFE_BYTES = 15_500_000


def _fit_extraction_doc(extraction_doc, sheet_name):
    """Ensures the extraction doc fits under MongoDB's 16 MiB limit, shedding
    the largest rebuildable fields first and recording what was dropped."""
    if len(bson.BSON.encode(extraction_doc)) <= _MONGO_DOC_SAFE_BYTES:
        return
    logger.warning("extraction doc for sheet '%s' exceeds safe size; shedding heavy fields", sheet_name)
    extraction_doc["structured"] = {"tables": [], "omitted": True}
    extraction_doc.setdefault("oversize_notes", []).append(
        "structured grid omitted (sheet too large for one document); "
        "use the lossless/intermediate endpoints for full detail"
    )
    if len(bson.BSON.encode(extraction_doc)) <= _MONGO_DOC_SAFE_BYTES:
        return
    kept = len(extraction_doc.get("rows", []))
    extraction_doc["rows"] = []
    extraction_doc["oversize_notes"].append(
        f"{kept} cell rows omitted (sheet too large for one document); "
        "use the raw.json / lossless endpoints for full cell data"
    )


def _object_id(workbook_id: str) -> ObjectId:
    try:
        return ObjectId(workbook_id)
    except InvalidId:
        raise HTTPException(status_code=400, detail="Invalid workbook id")


def _serialize_workbook(doc):
    doc = dict(doc)
    doc["id"] = str(doc.pop("_id"))
    return doc


def _serialize_extraction(doc):
    doc = dict(doc)
    doc["id"] = str(doc.pop("_id"))
    doc["workbook_id"] = str(doc["workbook_id"])
    return doc


@router.post("/upload")
async def upload_workbook(file: UploadFile):
    suffix = Path(file.filename).suffix.lower()
    if suffix not in config.ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail=f"Unsupported file extension '{suffix}'")

    contents = await file.read()
    stored_name = f"{uuid.uuid4().hex}{suffix}"
    stored_path = config.STORAGE_DIR / stored_name
    stored_path.write_bytes(contents)

    try:
        sheet_names, sheet_dimensions = read_sheet_dimensions(stored_path)
    except Exception as exc:
        stored_path.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail=f"Could not read workbook: {exc}")

    doc = {
        "filename": file.filename,
        "stored_path": str(stored_path),
        "size": len(contents),
        "sheet_names": sheet_names,
        "sheet_dimensions": sheet_dimensions,
        "status": "uploaded",
        "uploaded_at": datetime.datetime.utcnow().isoformat(),
    }
    result = await workbooks_collection.insert_one(doc)
    doc["_id"] = result.inserted_id
    return _serialize_workbook(doc)


# --- Per-sheet structuring pipeline -----------------------------------------
#
# Every sheet gets its own raw JSON, its own Gemini request, and its own
# structured JSON, processed one sheet at a time and saved immediately after
# each one completes. A later sheet hitting a token limit, quota error, or
# malformed response can never affect an earlier sheet's already-saved
# result - and the loop always continues to the next sheet regardless of
# whether the current one succeeded.


async def _set_sheet_structuring_state(oid, sheet_name, **fields):
    await structured_collection.update_one(
        {"workbook_id": oid, "sheet_name": sheet_name},
        {"$set": fields},
        upsert=True,
    )


# Max Gemini attempts for a single sheet before accepting the best output
# with its violations recorded. A hard schema-validation failure triggers a
# retry of ONLY this sheet (never the whole workbook).
_MAX_SHEET_ATTEMPTS = 2


def _build_intermediate_sheet_sync(stored_path, sheet_name):
    """Deterministic transformation-engine intermediate for one sheet, built
    from the lossless extraction. Ground truth for structured validation.
    Runs in a threadpool (CPU-bound openpyxl parse)."""
    try:
        lossless = extract_lossless(str(stored_path), only_sheet=sheet_name)
    except Exception:
        logger.exception("could not build intermediate for %s", sheet_name)
        return None
    sheet = lossless.get("sheets", {}).get(sheet_name)
    return transform_sheet(sheet) if sheet is not None else None


async def _run_structuring_for_sheet(oid, sheet_name):
    await _set_sheet_structuring_state(oid, sheet_name, status="processing", error=None, error_kind=None)

    workbook_doc = await workbooks_collection.find_one({"_id": oid})
    extraction = await extractions_collection.find_one({"workbook_id": oid, "sheet_name": sheet_name})
    if workbook_doc is None or extraction is None:
        await _set_sheet_structuring_state(
            oid, sheet_name,
            status="failed",
            error="workbook or sheet extraction is missing",
            error_kind="other",
            generated_at=datetime.datetime.utcnow().isoformat(),
        )
        return

    # This sheet's raw JSON contains only this sheet - never any other
    # sheet's rows, cells, or tables.
    raw_sheet_json = build_raw_sheet_json(workbook_doc, extraction)
    llm_payload = build_llm_payload(raw_sheet_json)
    sheet_raw = raw_sheet_json["sheets"][sheet_name]

    # Deterministic ground truth for validating whatever the LLM returns.
    stored_path = workbook_doc.get("stored_path")
    intermediate_sheet = None
    if stored_path and Path(stored_path).exists():
        loop = asyncio.get_event_loop()
        intermediate_sheet = await loop.run_in_executor(
            None, _build_intermediate_sheet_sync, stored_path, sheet_name
        )

    best = None  # (structured_data, warnings, schema_result)
    for attempt in range(1, _MAX_SHEET_ATTEMPTS + 1):
        try:
            # The LLM classifies this one sheet's meaning; it never computes
            # actual date/status values.
            llm_semantics = await generate_structured_json(llm_payload)
        except StructuringError as exc:
            logger.error(
                "structuring failed for workbook %s sheet '%s' (%s): %s",
                oid, sheet_name, exc.kind, exc,
            )
            await _set_sheet_structuring_state(
                oid, sheet_name,
                status="failed",
                error=str(exc),
                error_kind=exc.kind,
                generated_at=datetime.datetime.utcnow().isoformat(),
            )
            return

        # Every restriction/date range is computed deterministically here from
        # the trusted raw cells - the LLM never transcribes grid data.
        structured_data, assembly_warnings = assemble_sheet_structured_output(sheet_raw, llm_semantics)
        warnings = assembly_warnings + validate_sheet_structured_output(sheet_name, sheet_raw, structured_data)

        if intermediate_sheet is not None:
            schema_result = validate_structured_sheet(intermediate_sheet, structured_data)
        else:
            schema_result = {"valid": True, "hard_violations": [], "warnings": []}

        best = (structured_data, warnings + schema_result["warnings"], schema_result)

        if schema_result["valid"]:
            break
        logger.warning(
            "structuring for %s sheet '%s' failed validation on attempt %d/%d: %s",
            oid, sheet_name, attempt, _MAX_SHEET_ATTEMPTS, schema_result["hard_violations"],
        )
        # loop retries ONLY this sheet

    structured_data, all_warnings, schema_result = best
    validation_status = "valid" if schema_result["valid"] else "invalid"

    if not schema_result["valid"]:
        logger.warning(
            "structuring for %s sheet '%s' accepted with %d unresolved validation issue(s) after %d attempts",
            oid, sheet_name, len(schema_result["hard_violations"]), _MAX_SHEET_ATTEMPTS,
        )
    elif all_warnings:
        logger.warning("structuring for %s sheet '%s' completed with %d warning(s)", oid, sheet_name, len(all_warnings))
    else:
        logger.info("structuring for %s sheet '%s' completed cleanly", oid, sheet_name)

    await _set_sheet_structuring_state(
        oid, sheet_name,
        status="completed",
        error=None,
        error_kind=None,
        data=structured_data,
        validation_warnings=all_warnings,
        validation_status=validation_status,
        schema_violations=schema_result["hard_violations"],
        generated_at=datetime.datetime.utcnow().isoformat(),
    )


async def _run_structuring_for_workbook(oid, sheet_names):
    logger.info("structuring started for workbook %s across %d sheet(s): %s", oid, len(sheet_names), sheet_names)
    for sheet_name in sheet_names:
        try:
            await _run_structuring_for_sheet(oid, sheet_name)
        except Exception:
            # A truly unexpected bug in one sheet's processing must never
            # take down the sheets after it.
            logger.exception("unexpected error structuring workbook %s sheet '%s'", oid, sheet_name)
            await _set_sheet_structuring_state(
                oid, sheet_name,
                status="failed",
                error="unexpected internal error while structuring this sheet",
                error_kind="other",
                generated_at=datetime.datetime.utcnow().isoformat(),
            )
    logger.info("structuring finished for workbook %s", oid)


@router.post("/{workbook_id}/extract")
async def extract(workbook_id: str, background_tasks: BackgroundTasks):
    oid = _object_id(workbook_id)
    workbook_doc = await workbooks_collection.find_one({"_id": oid})
    if workbook_doc is None:
        raise HTTPException(status_code=404, detail="Workbook not found")

    stored_path = Path(workbook_doc["stored_path"])
    if not stored_path.exists():
        raise HTTPException(status_code=410, detail="Stored workbook file is missing")

    per_sheet_results, workbook_total_time = extract_workbook(stored_path)

    await extractions_collection.delete_many({"workbook_id": oid})

    sheet_summaries = {}
    for sheet_name, result in per_sheet_results.items():
        extraction_doc = {
            "workbook_id": oid,
            "sheet_name": sheet_name,
            "processing_time_seconds": result["processing_time_seconds"],
            "total_cells": result["total_cells"],
            "colored_cells": result["colored_cells"],
            "blank_cells": result["blank_cells"],
            "unresolved_cells": result["unresolved_cells"],
            "ambiguous_cells": result["ambiguous_cells"],
            "color_inventory": result["color_inventory"],
            "rows": result["rows"],
            "structured": result["structured"],
            "low_contrast_pairs": result.get("low_contrast_pairs", []),
            "embedded_images": result.get("embedded_images", []),
            "extraction_summary": result.get("extraction_summary"),
        }
        _fit_extraction_doc(extraction_doc, sheet_name)
        await extractions_collection.insert_one(extraction_doc)
        sheet_summaries[sheet_name] = {
            "processing_time_seconds": result["processing_time_seconds"],
            "total_cells": result["total_cells"],
            "colored_cells": result["colored_cells"],
            "blank_cells": result["blank_cells"],
            "unresolved_cells": result["unresolved_cells"],
            "ambiguous_cells": result["ambiguous_cells"],
            "color_inventory": result["color_inventory"],
        }

    await workbooks_collection.update_one(
        {"_id": oid},
        {"$set": {"status": "extracted", "extracted_at": datetime.datetime.utcnow().isoformat()}},
    )

    # Plain extraction is complete and durable at this point, with zero AI /
    # token cost. Structured-JSON generation is NOT started here - it runs
    # strictly on demand via /structured/generate, so the plain extract step
    # stays instant and free. Clear any structured results from a previous
    # extraction of this workbook so nothing stale is served.
    await structured_collection.delete_many({"workbook_id": oid})

    return {
        "workbook_id": str(oid),
        "total_processing_time_seconds": workbook_total_time,
        "sheets": sheet_summaries,
    }


@router.get("/{workbook_id}")
async def get_workbook(workbook_id: str):
    oid = _object_id(workbook_id)
    workbook_doc = await workbooks_collection.find_one({"_id": oid})
    if workbook_doc is None:
        raise HTTPException(status_code=404, detail="Workbook not found")

    result = _serialize_workbook(workbook_doc)

    extractions = await extractions_collection.find({"workbook_id": oid}).to_list(length=None)
    if extractions:
        result["sheets"] = {
            e["sheet_name"]: {
                "processing_time_seconds": e["processing_time_seconds"],
                "total_cells": e["total_cells"],
                "colored_cells": e["colored_cells"],
                "blank_cells": e["blank_cells"],
                "unresolved_cells": e["unresolved_cells"],
                "ambiguous_cells": e["ambiguous_cells"],
                "color_inventory": e["color_inventory"],
            }
            for e in extractions
        }
    return result


async def _get_extraction_or_404(oid, sheet_name):
    extraction = await extractions_collection.find_one({"workbook_id": oid, "sheet_name": sheet_name})
    if extraction is None:
        raise HTTPException(status_code=404, detail="Sheet extraction not found. Run /extract first.")
    return extraction


@router.get("/{workbook_id}/sheets/{sheet_name}")
async def get_sheet(workbook_id: str, sheet_name: str):
    oid = _object_id(workbook_id)
    extraction = await _get_extraction_or_404(oid, sheet_name)
    return _serialize_extraction(extraction)


@router.post("/{workbook_id}/sheets/{sheet_name}/labels")
async def add_label(workbook_id: str, sheet_name: str, payload: dict):
    """
    Human-driven only. payload: {"color_key": "#RRGGBB" or theme_ref, "label": "text"}
    Attaches the label to every row in this sheet whose color matches color_key,
    and to the matching color_inventory entry. Never applied automatically,
    never carried over to other workbooks.
    """
    color_key = payload.get("color_key")
    label = payload.get("label")
    if not color_key or label is None:
        raise HTTPException(status_code=400, detail="color_key and label are required")

    oid = _object_id(workbook_id)
    extraction = await _get_extraction_or_404(oid, sheet_name)

    def _row_matches(row):
        color = row.get("color") or {}
        if color.get("resolved") and color.get("hex"):
            return f"#{color['hex']}" == color_key
        if color.get("theme_ref"):
            return color["theme_ref"] == color_key
        return False

    updated_rows = []
    for row in extraction["rows"]:
        if _row_matches(row):
            row = dict(row)
            row["human_label"] = label
        updated_rows.append(row)

    updated_inventory = []
    matched_inventory = False
    for entry in extraction["color_inventory"]:
        entry = dict(entry)
        if entry.get("key") == color_key:
            entry["human_label"] = label
            matched_inventory = True
        updated_inventory.append(entry)

    if not matched_inventory:
        raise HTTPException(status_code=404, detail=f"No color '{color_key}' found in this sheet's inventory")

    await extractions_collection.update_one(
        {"_id": extraction["_id"]},
        {"$set": {"rows": updated_rows, "color_inventory": updated_inventory}},
    )
    return {"status": "ok", "color_key": color_key, "label": label}


def _content_type_for(fmt):
    return {
        "json": "application/json",
        "csv": "text/csv",
    }.get(fmt)


@router.get("/{workbook_id}/sheets/{sheet_name}/report")
async def sheet_report(workbook_id: str, sheet_name: str, format: str = "json"):
    if format not in ("csv", "json"):
        raise HTTPException(status_code=400, detail="format must be csv or json")

    oid = _object_id(workbook_id)
    extraction = await _get_extraction_or_404(oid, sheet_name)
    extraction = dict(extraction)
    extraction["sheet_name"] = sheet_name

    if format == "json":
        content = reports.sheet_report_json(extraction)
    else:
        content = reports.sheet_report_csv(extraction)

    filename = f"{sheet_name}_report.{format}"
    return StreamingResponse(
        io.BytesIO(content),
        media_type=_content_type_for(format),
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/{workbook_id}/report")
async def workbook_report(workbook_id: str, format: str = "json"):
    if format not in ("csv", "json"):
        raise HTTPException(status_code=400, detail="format must be csv or json")

    oid = _object_id(workbook_id)
    workbook_doc = await workbooks_collection.find_one({"_id": oid})
    if workbook_doc is None:
        raise HTTPException(status_code=404, detail="Workbook not found")

    extractions = await extractions_collection.find({"workbook_id": oid}).to_list(length=None)
    if not extractions:
        raise HTTPException(status_code=404, detail="No extraction found. Run /extract first.")

    if format == "json":
        content = reports.workbook_report_json_zip(workbook_doc, extractions)
        filename = f"{workbook_doc['filename']}_report.zip"
        media_type = "application/zip"
    else:
        content = reports.workbook_report_csv_zip(workbook_doc, extractions)
        filename = f"{workbook_doc['filename']}_report.zip"
        media_type = "application/zip"

    return StreamingResponse(
        io.BytesIO(content),
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# --- Raw JSON downloads (always available immediately after /extract) ------


@router.get("/{workbook_id}/raw.json")
async def download_raw_json(workbook_id: str):
    """
    The full raw extraction for every sheet in the workbook, as one JSON
    file. Rebuilt on demand from already-stored extraction docs - never
    reprocesses the .xlsx file, so it's available any time after /extract.
    """
    oid = _object_id(workbook_id)
    workbook_doc = await workbooks_collection.find_one({"_id": oid})
    if workbook_doc is None:
        raise HTTPException(status_code=404, detail="Workbook not found")

    extractions = await extractions_collection.find({"workbook_id": oid}).to_list(length=None)
    if not extractions:
        raise HTTPException(status_code=404, detail="No extraction found. Run /extract first.")

    raw_json = build_raw_workbook_json(workbook_doc, extractions)
    content = json.dumps(raw_json, indent=2, default=str).encode("utf-8")
    filename = f"{workbook_doc['filename']}_raw.json"
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/json",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/{workbook_id}/sheets/{sheet_name}/raw.json")
async def download_sheet_raw_json(workbook_id: str, sheet_name: str):
    """The raw extraction for exactly this sheet - no other sheet's data."""
    oid = _object_id(workbook_id)
    workbook_doc = await workbooks_collection.find_one({"_id": oid})
    if workbook_doc is None:
        raise HTTPException(status_code=404, detail="Workbook not found")

    extraction = await _get_extraction_or_404(oid, sheet_name)
    raw_json = build_raw_sheet_json(workbook_doc, extraction)
    content = json.dumps(raw_json, indent=2, default=str).encode("utf-8")
    filename = f"{workbook_doc['filename']}_{sheet_name}_raw.json"
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/json",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# --- Lossless raw extraction (full-fidelity, every cell) -------------------
#
# Re-parses the stored .xlsx on demand (the file is the source of truth, so
# the result is always faithful) and runs the CPU-heavy parse in a threadpool
# so the event loop is never blocked. This is additive - it does not touch
# the existing compact extract/structuring pipeline.


async def _stored_path_or_404(oid):
    workbook_doc = await workbooks_collection.find_one({"_id": oid})
    if workbook_doc is None:
        raise HTTPException(status_code=404, detail="Workbook not found")
    stored_path = Path(workbook_doc["stored_path"])
    if not stored_path.exists():
        raise HTTPException(status_code=410, detail="Stored workbook file is missing")
    return workbook_doc, stored_path


async def _run_lossless(stored_path, filename, only_sheet=None):
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(
        None, lambda: extract_lossless(str(stored_path), filename=filename, only_sheet=only_sheet)
    )


def _json_download(payload, filename):
    content = json.dumps(payload, indent=2, default=str).encode("utf-8")
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/json",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/{workbook_id}/lossless/raw.json")
async def download_lossless_workbook(workbook_id: str):
    """Full-fidelity lossless raw JSON for the whole workbook (every cell,
    all colour internals, metadata, merges, images)."""
    oid = _object_id(workbook_id)
    workbook_doc, stored_path = await _stored_path_or_404(oid)
    lossless = await _run_lossless(stored_path, workbook_doc.get("filename"))
    return _json_download(lossless, f"{workbook_doc['filename']}_lossless.json")


@router.get("/{workbook_id}/sheets/{sheet_name}/lossless/raw.json")
async def download_lossless_sheet(workbook_id: str, sheet_name: str):
    """Full-fidelity lossless raw JSON for exactly one sheet - the lighter,
    scalable path for large workbooks."""
    oid = _object_id(workbook_id)
    workbook_doc, stored_path = await _stored_path_or_404(oid)
    if sheet_name not in (workbook_doc.get("sheet_names") or []):
        raise HTTPException(status_code=404, detail=f"Sheet '{sheet_name}' not found in this workbook")
    lossless = await _run_lossless(stored_path, workbook_doc.get("filename"), only_sheet=sheet_name)
    return _json_download(lossless, f"{workbook_doc['filename']}_{sheet_name}_lossless.json")


@router.get("/{workbook_id}/lossless/validation")
async def lossless_validation(workbook_id: str):
    """Deterministic validation report over the lossless extraction: total
    cells processed, colored cells, formulas, comments, hyperlinks, merged
    regions, hidden rows/columns, and images - per sheet and in total."""
    oid = _object_id(workbook_id)
    workbook_doc, stored_path = await _stored_path_or_404(oid)
    lossless = await _run_lossless(stored_path, workbook_doc.get("filename"))
    return {
        "workbook": workbook_doc.get("filename"),
        "sheet_count": lossless["workbook"]["sheet_count"],
        "validation": lossless["validation"],
    }


# --- Deterministic transformation engine (intermediate business-aware JSON) -
#
# Converts the lossless raw JSON into the intermediate representation
# (tables, headers, date columns, colour map, legend candidates, row objects
# with cell-ref traceability) with NO LLM involvement. Additive - the live
# structuring flow is unchanged; these let you inspect/validate the
# deterministic layer across supplier workbooks before it replaces anything.


def _lossless_and_intermediate_workbook(stored_path, filename):
    lossless = extract_lossless(str(stored_path), filename=filename)
    return lossless, transform_workbook(lossless)


def _lossless_and_intermediate_sheet(stored_path, filename, sheet_name):
    lossless = extract_lossless(str(stored_path), filename=filename, only_sheet=sheet_name)
    sheet = lossless.get("sheets", {}).get(sheet_name)
    return transform_sheet(sheet) if sheet is not None else None


@router.get("/{workbook_id}/intermediate.json")
async def download_intermediate_workbook(workbook_id: str):
    """Deterministic intermediate JSON for the whole workbook."""
    oid = _object_id(workbook_id)
    workbook_doc, stored_path = await _stored_path_or_404(oid)
    loop = asyncio.get_event_loop()
    _lossless, intermediate = await loop.run_in_executor(
        None, _lossless_and_intermediate_workbook, str(stored_path), workbook_doc.get("filename")
    )
    return _json_download(intermediate, f"{workbook_doc['filename']}_intermediate.json")


@router.get("/{workbook_id}/sheets/{sheet_name}/intermediate.json")
async def download_intermediate_sheet(workbook_id: str, sheet_name: str):
    """Deterministic intermediate JSON for exactly one sheet."""
    oid = _object_id(workbook_id)
    workbook_doc, stored_path = await _stored_path_or_404(oid)
    if sheet_name not in (workbook_doc.get("sheet_names") or []):
        raise HTTPException(status_code=404, detail=f"Sheet '{sheet_name}' not found in this workbook")
    loop = asyncio.get_event_loop()
    intermediate = await loop.run_in_executor(
        None, _lossless_and_intermediate_sheet, str(stored_path), workbook_doc.get("filename"), sheet_name
    )
    return _json_download(intermediate, f"{workbook_doc['filename']}_{sheet_name}_intermediate.json")


@router.get("/{workbook_id}/extraction-validation")
async def extraction_completeness(workbook_id: str):
    """
    Pre-structuring completeness check: independently recounts the workbook
    from openpyxl and compares against the lossless extraction (cell counts,
    colored cells, merged ranges, comments, hyperlinks, hidden rows/columns,
    sheet metadata), plus a legend-usage audit (matching-cell counts per
    legend colour/code, and colours with no legend meaning).
    """
    oid = _object_id(workbook_id)
    workbook_doc, stored_path = await _stored_path_or_404(oid)

    def _build():
        lossless, intermediate = _lossless_and_intermediate_workbook(
            str(stored_path), workbook_doc.get("filename")
        )
        return validate_extraction_completeness(str(stored_path), lossless, intermediate)

    loop = asyncio.get_event_loop()
    report = await loop.run_in_executor(None, _build)
    return report


# --- Structuring progress + per-sheet / workbook-level downloads -----------


def _progress_message(sheet_names, sheets):
    total = len(sheet_names)
    if total == 0:
        return "No sheets to process."

    processing_name = next((n for n in sheet_names if sheets[n]["status"] == "processing"), None)
    if processing_name:
        return f"Processing '{processing_name}'... waiting for Gemini..."

    completed = [n for n in sheet_names if sheets[n]["status"] == "completed"]
    failed = [n for n in sheet_names if sheets[n]["status"] == "failed"]
    pending = [n for n in sheet_names if sheets[n]["status"] in ("pending", "not_started")]

    if len(completed) + len(failed) == total:
        if not failed:
            return f"Structured JSON generated for all {total} sheet(s)."
        return (
            f"Structured JSON generated for {len(completed)} of {total} sheet(s). "
            f"Failed: {', '.join(failed)}."
        )

    if len(pending) == total:
        return "Waiting to start..."

    return f"{len(completed)} of {total} sheet(s) completed so far."


@router.get("/{workbook_id}/structured/progress")
async def structured_progress(workbook_id: str):
    """
    Per-sheet structuring status plus a ready-to-display overall message -
    e.g. "Processing 'Sheet 4'... waiting for Gemini..." or "Structured JSON
    generated for 6 of 8 sheet(s). Failed: Sheet 3, Sheet 7." Poll this to
    drive a real per-sheet progress UI instead of a single spinner.
    """
    oid = _object_id(workbook_id)
    workbook_doc = await workbooks_collection.find_one({"_id": oid})
    if workbook_doc is None:
        raise HTTPException(status_code=404, detail="Workbook not found")

    sheet_names = workbook_doc.get("sheet_names") or []
    docs = await structured_collection.find({"workbook_id": oid}).to_list(length=None)
    by_name = {d["sheet_name"]: d for d in docs}

    sheets = {}
    counts = {"not_started": 0, "pending": 0, "processing": 0, "completed": 0, "failed": 0}
    for name in sheet_names:
        d = by_name.get(name)
        status = (d or {}).get("status") or "not_started"
        counts[status] = counts.get(status, 0) + 1
        sheets[name] = {
            "status": status,
            "error": (d or {}).get("error"),
            "error_kind": (d or {}).get("error_kind"),
            "validation_warnings": (d or {}).get("validation_warnings") or [],
            "validation_status": (d or {}).get("validation_status"),
            "schema_violations": (d or {}).get("schema_violations") or [],
            "generated_at": (d or {}).get("generated_at"),
        }

    return {
        "sheets": sheets,
        "overall": {
            "total": len(sheet_names),
            **{k: v for k, v in counts.items()},
            "message": _progress_message(sheet_names, sheets),
        },
    }


async def _sheet_names_or_404(oid):
    workbook_doc = await workbooks_collection.find_one({"_id": oid})
    if workbook_doc is None:
        raise HTTPException(status_code=404, detail="Workbook not found")
    return workbook_doc, workbook_doc.get("sheet_names") or []


@router.post("/{workbook_id}/sheets/{sheet_name}/structured/generate")
async def generate_sheet_structuring(
    workbook_id: str, sheet_name: str, background_tasks: BackgroundTasks, force: bool = False
):
    """
    Runs AI structuring for exactly one sheet, on demand only.

    - force=false (default): if this sheet's structured JSON already exists
      (status "completed"), returns immediately WITHOUT calling the AI again -
      the cached result is reused. Otherwise it starts generation.
    - force=true: always regenerates, even if a cached result exists (this is
      the "Regenerate" action, the only thing that re-spends tokens).
    """
    oid = _object_id(workbook_id)
    _, sheet_names = await _sheet_names_or_404(oid)
    if sheet_name not in sheet_names:
        raise HTTPException(status_code=404, detail=f"Sheet '{sheet_name}' not found in this workbook")

    has_extraction = await extractions_collection.count_documents({"workbook_id": oid, "sheet_name": sheet_name})
    if not has_extraction:
        raise HTTPException(status_code=400, detail="No extraction found for this sheet. Run /extract first.")

    existing = await structured_collection.find_one({"workbook_id": oid, "sheet_name": sheet_name})
    if not force and existing and existing.get("status") == "completed":
        return {"status": "completed", "already_generated": True, "sheet_name": sheet_name}

    await _set_sheet_structuring_state(oid, sheet_name, status="pending", error=None, error_kind=None)
    background_tasks.add_task(_run_structuring_for_sheet, oid, sheet_name)
    return {"status": "pending", "already_generated": False, "sheet_name": sheet_name}


@router.post("/{workbook_id}/structured/generate")
async def generate_structuring(workbook_id: str, background_tasks: BackgroundTasks, force: bool = False):
    """
    Runs AI structuring for the whole workbook, on demand only.

    - force=false (default): generates only sheets that don't already have a
      cached "completed" result - already-generated sheets are reused, so no
      tokens are re-spent on them.
    - force=true: regenerates every sheet from scratch ("Regenerate").
    """
    oid = _object_id(workbook_id)
    _, sheet_names = await _sheet_names_or_404(oid)

    has_extraction = await extractions_collection.count_documents({"workbook_id": oid})
    if not has_extraction:
        raise HTTPException(status_code=400, detail="No extraction found. Run /extract first.")

    docs = await structured_collection.find({"workbook_id": oid}).to_list(length=None)
    completed = {d["sheet_name"] for d in docs if d.get("status") == "completed"}
    to_generate = sheet_names if force else [name for name in sheet_names if name not in completed]

    for name in to_generate:
        await _set_sheet_structuring_state(oid, name, status="pending", error=None, error_kind=None)

    if to_generate:
        background_tasks.add_task(_run_structuring_for_workbook, oid, to_generate)
    return {
        "status": "pending" if to_generate else "completed",
        "sheets_to_generate": to_generate,
    }


@router.get("/{workbook_id}/sheets/{sheet_name}/structured")
async def download_sheet_structured_json(workbook_id: str, sheet_name: str):
    """Downloads exactly this sheet's structured JSON. Only succeeds once
    this sheet's structuring has completed."""
    oid = _object_id(workbook_id)
    workbook_doc, sheet_names = await _sheet_names_or_404(oid)
    if sheet_name not in sheet_names:
        raise HTTPException(status_code=404, detail=f"Sheet '{sheet_name}' not found in this workbook")

    doc = await structured_collection.find_one({"workbook_id": oid, "sheet_name": sheet_name})
    if doc is None or doc.get("status") in (None, "not_started"):
        raise HTTPException(status_code=404, detail="Structuring has not been started for this sheet")
    if doc.get("status") in ("pending", "processing"):
        raise HTTPException(status_code=409, detail="Structured JSON is still being generated for this sheet")
    if doc.get("status") == "failed":
        raise HTTPException(status_code=424, detail=doc.get("error") or "Structuring failed for this sheet")

    content = json.dumps(doc["data"], indent=2, default=str).encode("utf-8")
    filename = f"{workbook_doc['filename']}_{sheet_name}_structured.json"
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/json",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/{workbook_id}/structured")
async def download_structured_json(workbook_id: str):
    """
    Downloads the combined workbook-level structured JSON: every sheet that
    has completed structuring gets its real data; any sheet that's still
    pending/processing/failed is included honestly with its status and error
    instead of being silently dropped from the file. Always available - a
    failure on one sheet never blocks downloading the sheets that succeeded.
    """
    oid = _object_id(workbook_id)
    workbook_doc, sheet_names = await _sheet_names_or_404(oid)

    docs = await structured_collection.find({"workbook_id": oid}).to_list(length=None)
    by_name = {d["sheet_name"]: d for d in docs}
    if not docs:
        raise HTTPException(status_code=404, detail="Structuring has not been started for this workbook")

    sheets_out = {}
    for name in sheet_names:
        d = by_name.get(name)
        status = (d or {}).get("status") or "not_started"
        if status == "completed":
            sheets_out[name] = (d or {}).get("data")
        else:
            sheets_out[name] = {
                "_status": status,
                "_error": (d or {}).get("error"),
            }

    content = json.dumps({"sheets": sheets_out}, indent=2, default=str).encode("utf-8")
    filename = f"{workbook_doc['filename']}_structured.json"
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/json",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
