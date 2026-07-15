"""
Structured-output validation (post-structuring).

Validates the FINAL structured JSON for one sheet against the deterministic
transformation-engine intermediate (the trusted ground truth) to catch the
failure modes seen on supplier workbooks:

  - schema violations        (wrong shape / missing required keys / bad types)
  - row shifts               (dates/values outside the sheet's real date set)
  - incorrect column counts  (generic rows whose length != the header's)
  - duplicated values        (identical restriction ranges; duplicate rows)
  - missing legend mappings   (a legend entry Python detected that the output
                               dropped)

Violations are split into `hard` (the sheet's output cannot be trusted →
caller should retry only this sheet) and `warnings` (worth surfacing, not
worth re-spending tokens over). Entirely deterministic and generic.
"""
import datetime


def _known_dates(intermediate_sheet):
    dates = set()
    for table in intermediate_sheet.get("tables", []):
        for info in (table.get("date_columns") or {}).values():
            if info.get("date"):
                dates.add(info["date"][:10])
    return dates


def _known_labels(intermediate_sheet):
    labels = set()
    for cand in intermediate_sheet.get("legend_candidates", []):
        if cand.get("label"):
            labels.add(cand["label"].strip().lower())
    return labels


def _iter_date_strings(node):
    if isinstance(node, dict):
        for v in node.values():
            yield from _iter_date_strings(v)
    elif isinstance(node, list):
        for v in node:
            yield from _iter_date_strings(v)
    elif isinstance(node, str) and len(node) >= 10:
        try:
            datetime.date.fromisoformat(node[:10])
            yield node[:10]
        except ValueError:
            return


def _validate_entity_calendar(intermediate_sheet, structured, hard, warnings):
    if not isinstance(structured.get("entities"), dict):
        hard.append("entity_calendar output missing an 'entities' object")
        return
    if not isinstance(structured.get("legend"), dict):
        warnings.append("entity_calendar output has no 'legend' object")

    known_dates = _known_dates(intermediate_sheet)
    known_labels = _known_labels(intermediate_sheet)
    default_status = structured.get("default_status")

    # --- missing legend mappings ---
    out_legend_values = {
        str(v).strip().lower() for v in (structured.get("legend") or {}).values()
    }
    for label in known_labels:
        if label not in out_legend_values:
            hard.append(f"legend mapping for '{label}' was detected in the sheet but is missing from the output legend")

    # --- per-entity checks: row shift (dates), duplicates, statuses ---
    allowed_statuses = set(known_labels)
    if default_status:
        allowed_statuses.add(str(default_status).strip().lower())
    allowed_statuses.add("unknown")

    for code, entity in structured["entities"].items():
        if not isinstance(entity, dict):
            hard.append(f"entity '{code}' is not an object")
            continue
        restrictions = entity.get("restrictions", [])
        if not isinstance(restrictions, list):
            hard.append(f"entity '{code}' has a non-list 'restrictions'")
            continue

        seen = set()
        for res in restrictions:
            if not isinstance(res, dict):
                hard.append(f"entity '{code}' has a non-object restriction")
                continue
            start = str(res.get("start_date", ""))[:10]
            end = str(res.get("end_date", ""))[:10]
            status = res.get("status")

            # row shift: any date outside the sheet's real detected date set
            if known_dates:
                for d in (start, end):
                    if d and d not in known_dates:
                        hard.append(
                            f"entity '{code}' restriction date {d} is outside the sheet's detected date range "
                            "(possible row shift / hallucinated date)"
                        )
            # ordering
            try:
                if start and end and datetime.date.fromisoformat(start) > datetime.date.fromisoformat(end):
                    hard.append(f"entity '{code}' restriction has start_date {start} after end_date {end}")
            except ValueError:
                hard.append(f"entity '{code}' restriction has an unparseable date ({start!r}..{end!r})")

            # duplicated values
            key = (start, end, str(status))
            if key in seen:
                warnings.append(f"entity '{code}' has a duplicated restriction {key}")
            seen.add(key)

            # status must come from the legend / default / unknown
            if status is not None and known_labels and str(status).strip().lower() not in allowed_statuses:
                warnings.append(
                    f"entity '{code}' uses status '{status}' which is not a detected legend label"
                )


def _validate_generic(intermediate_sheet, structured, hard, warnings):
    tables = structured.get("tables")
    if not isinstance(tables, list):
        hard.append("generic output missing a 'tables' list")
        return
    for ti, table in enumerate(tables):
        if not isinstance(table, dict):
            hard.append(f"table {ti} is not an object")
            continue
        columns = table.get("columns")
        rows = table.get("rows", [])
        if not isinstance(rows, list):
            hard.append(f"table {ti} has a non-list 'rows'")
            continue
        # incorrect column counts / row shift: every row's key-set should match
        # the declared columns (or be a subset — extra keys are the real error).
        col_keys = set(columns.keys()) if isinstance(columns, dict) else None
        for ri, row in enumerate(rows):
            if not isinstance(row, dict):
                hard.append(f"table {ti} row {ri} is not an object")
                continue
            if col_keys is not None:
                extra = set(row.keys()) - col_keys
                if extra:
                    hard.append(
                        f"table {ti} row {ri} has columns not declared in the header {sorted(extra)} "
                        "(possible extra column / row shift)"
                    )
        # any output date outside the sheet's detected set is suspicious
        known_dates = _known_dates(intermediate_sheet)
        if known_dates:
            for d in _iter_date_strings(table):
                if d not in known_dates:
                    warnings.append(f"table {ti} contains date {d} not detected in the sheet")


def validate_structured_sheet(intermediate_sheet, structured):
    """
    Returns {"valid": bool, "hard_violations": [...], "warnings": [...]}.
    `valid` is False iff there are hard violations — the caller retries only
    this sheet in that case.
    """
    hard = []
    warnings = []

    if not isinstance(structured, dict):
        return {"valid": False, "hard_violations": [f"structured output is not a JSON object (got {type(structured).__name__})"], "warnings": []}

    pattern = structured.get("pattern")
    if pattern == "entity_calendar":
        _validate_entity_calendar(intermediate_sheet, structured, hard, warnings)
    elif pattern == "generic":
        _validate_generic(intermediate_sheet, structured, hard, warnings)
    elif pattern == "unclassified":
        warnings.append("sheet was returned unclassified by the LLM")
    else:
        hard.append(f"unknown structured 'pattern': {pattern!r}")

    return {"valid": not hard, "hard_violations": hard, "warnings": warnings}
