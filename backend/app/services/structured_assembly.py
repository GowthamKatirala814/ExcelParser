"""
Deterministic assembly of one sheet's final structured JSON from the LLM's
semantic classification of that same sheet. The LLM identifies *meaning*
(legend, entity master data, which row belongs to which entity) - genuinely
a judgment call. Every actual date/status value in the output is then
computed here, directly from the already-trusted raw extraction, by simple
date-range collapsing. This is what guarantees restriction ranges match the
source cells exactly: the LLM never transcribes or counts grid cells, so it
cannot introduce an off-by-N-day error or a dropped/merged range - that
arithmetic never touches the LLM at all.

Each sheet is processed independently and only ever needs its own raw data
and its own LLM classification - there is no cross-sheet reconciliation here,
since sheets are structured one at a time.

Entirely generic: nothing here references a specific workbook, sheet name,
or entity vocabulary ("rooms" is just whatever key the LLM chose for this
workbook's entities).
"""
import datetime


def _cell_status(raw_cell, legend, default_status):
    """
    Resolves one cell's status by looking its value or fill color up in the
    LLM-provided legend. Returns None for "no restriction" (blank cell, or a
    status equal to the default), or "unknown" if the cell is populated but
    matches no legend entry - never a guess.
    """
    if raw_cell is None:
        return None

    value = raw_cell.get("value")
    color = raw_cell.get("color") or {}
    hex_key = f"#{color['hex']}" if color.get("resolved") and color.get("hex") else None

    status = None
    if isinstance(value, str) and value in legend:
        status = legend[value]
    elif hex_key is not None and hex_key in legend:
        status = legend[hex_key]
    elif value is None and hex_key is None:
        return None
    else:
        status = "unknown"

    if default_status is not None and status == default_status:
        return None
    return status


def _collapse_ranges(day_status_pairs):
    """
    day_status_pairs: (date_iso, status) pairs, pre-filtered to exclude
    None/default days, in any order. Collapses consecutive *calendar* days
    (not just consecutive list entries) sharing the same status into a
    single {"start_date", "end_date", "status"} range.
    """
    ranges = []
    ordered = sorted(day_status_pairs, key=lambda pair: pair[0])
    if not ordered:
        return ranges

    start_date, current_status = ordered[0]
    prev_date = start_date

    for date_str, status in ordered[1:]:
        prev_dt = datetime.date.fromisoformat(prev_date[:10])
        cur_dt = datetime.date.fromisoformat(date_str[:10])
        if status == current_status and (cur_dt - prev_dt).days <= 1:
            prev_date = date_str
            continue
        ranges.append({"start_date": start_date, "end_date": prev_date, "status": current_status})
        start_date, current_status, prev_date = date_str, status, date_str

    ranges.append({"start_date": start_date, "end_date": prev_date, "status": current_status})
    return ranges


def _assemble_entity_calendar(sheet_raw, sheet_semantics):
    """
    sheet_raw: this sheet's section of the raw_json (rows + structured.tables).
    sheet_semantics: the LLM's {"legend", "default_status", "entities", "row_to_entity"}.
    Returns ({"default_status", "legend", "entities": {...}}, warnings).
    """
    warnings = []
    legend = sheet_semantics.get("legend") or {}
    default_status = sheet_semantics.get("default_status")
    entities = sheet_semantics.get("entities") or {}

    row_to_entity = {}
    for row_key, code in (sheet_semantics.get("row_to_entity") or {}).items():
        try:
            row_to_entity[int(row_key)] = code
        except (TypeError, ValueError):
            continue

    raw_by_ref = {row["cell_ref"]: row for row in sheet_raw.get("rows", []) or []}
    tables = (sheet_raw.get("structured") or {}).get("tables", []) or []

    timelines = {code: [] for code in entities}
    orphan_codes_seen = set()

    for table in tables:
        column_index = table.get("column_index")
        if not column_index:
            continue
        dated_columns = [
            (col, info["date"]) for col, info in column_index.items() if info.get("date")
        ]
        row_range = table.get("row_range") or [0, -1]

        for row_num in range(row_range[0], row_range[1] + 1):
            code = row_to_entity.get(row_num)
            if code is None:
                continue
            if code not in timelines:
                if code not in orphan_codes_seen:
                    orphan_codes_seen.add(code)
                    warnings.append(
                        f"row {row_num} (and possibly others) maps to entity code '{code}' via "
                        "row_to_entity, but that code has no entry in 'entities' - its data was "
                        "skipped rather than guessed"
                    )
                continue
            for col_letter, date_str in dated_columns:
                raw_cell = raw_by_ref.get(f"{col_letter}{row_num}")
                status = _cell_status(raw_cell, legend, default_status)
                if status is not None:
                    timelines[code].append((date_str, status))

    entities_out = {}
    for code, info in entities.items():
        info = dict(info) if isinstance(info, dict) else {}
        info.pop("restrictions", None)  # any LLM-authored restrictions are discarded, not trusted
        info["restrictions"] = _collapse_ranges(timelines.get(code, []))
        entities_out[code] = info
        if code not in {row_to_entity.get(r) for r in row_to_entity}:
            warnings.append(
                f"entity '{code}' has master data but no calendar row was mapped to it via "
                "row_to_entity - it will show no restrictions"
            )

    assembled = {
        "pattern": "entity_calendar",
        "default_status": default_status,
        "legend": legend,
        "entities": entities_out,
    }
    return assembled, warnings


def assemble_sheet_structured_output(sheet_raw, llm_semantics):
    """
    Produces the final structured JSON for exactly one sheet from its own raw
    data and its own LLM classification. Returns (data_dict, warnings).
    """
    if not isinstance(llm_semantics, dict):
        return (
            {
                "pattern": "unclassified",
                "reason": f"LLM response was a {type(llm_semantics).__name__}, not a JSON object",
            },
            [f"could not classify this sheet: response was a {type(llm_semantics).__name__}"],
        )

    if llm_semantics.get("pattern") == "entity_calendar":
        return _assemble_entity_calendar(sheet_raw, llm_semantics)

    return {"pattern": "generic", "tables": llm_semantics.get("tables", [])}, []
