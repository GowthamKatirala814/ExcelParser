"""
Post-generation validation of one sheet's LLM-derived structured output
against the raw extraction it was built from. Entirely generic - no
workbook-specific vocabulary. Never blocks the result (the raw extraction
and even the structured JSON stay downloadable); it only surfaces concrete,
checkable inconsistencies instead of silently trusting the LLM's output.

Checks performed, cross-referenced against facts we already trust (computed
by our own deterministic code, never by the LLM):
  - Every date string appearing anywhere in the output is valid and falls
    within the min/max date actually observed in this sheet's detected date
    rows (column_index). A date outside that range, or any date at all when
    the sheet had none, is very likely hallucinated.
  - For "generic"-pattern sheets (small tables the LLM transcribes directly,
    since entity_calendar sheets never have the LLM transcribe data at all):
    every distinct non-blank raw cell value should appear somewhere in the
    output. Values that don't are reported by name (up to a sample) so a
    silent drop is visible instead of invisible.

This is intentionally a narrow, high-confidence set of checks - a broader
"every status word must appear verbatim in the raw cells" check was
considered and dropped for entity_calendar sheets: an LLM legitimately
normalizes/translates a color legend into a status label that won't
literally match the source text, so that check would mostly produce
false-positive noise rather than real signal. Extend this module only with
checks that are similarly verifiable against ground truth we already
computed, not fuzzy text matching.
"""
import datetime
import re

_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}")


def _iter_leaf_values(node):
    if isinstance(node, dict):
        for value in node.values():
            yield from _iter_leaf_values(value)
    elif isinstance(node, list):
        for item in node:
            yield from _iter_leaf_values(item)
    else:
        yield node


def _iter_strings(node):
    for value in _iter_leaf_values(node):
        if isinstance(value, str):
            yield value


def _sheet_date_bounds(sheet_raw):
    dates = []
    for table in (sheet_raw.get("structured") or {}).get("tables", []) or []:
        for info in (table.get("column_index") or {}).values():
            date_str = info.get("date")
            if date_str:
                try:
                    dates.append(datetime.date.fromisoformat(date_str[:10]))
                except ValueError:
                    continue
    if not dates:
        return None
    return min(dates), max(dates)


def _validate_dates(sheet_name, sheet_raw, sheet_structured):
    warnings = []
    bounds = _sheet_date_bounds(sheet_raw)

    out_of_range_dates = set()
    unexpected_dates = set()
    for text in _iter_strings(sheet_structured):
        if not _DATE_RE.match(text):
            continue
        try:
            d = datetime.date.fromisoformat(text[:10])
        except ValueError:
            continue
        if bounds is None:
            unexpected_dates.add(text[:10])
        elif not (bounds[0] <= d <= bounds[1]):
            out_of_range_dates.add(text[:10])

    if unexpected_dates:
        warnings.append(
            f"sheet '{sheet_name}': output contains dates {sorted(unexpected_dates)[:5]} "
            "but no date header was detected anywhere in this sheet's raw extraction"
        )
    if out_of_range_dates:
        lo, hi = bounds
        warnings.append(
            f"sheet '{sheet_name}': output contains dates {sorted(out_of_range_dates)[:5]} "
            f"outside the extracted range {lo.isoformat()}..{hi.isoformat()}"
        )
    return warnings


def _validate_generic_completeness(sheet_name, sheet_raw, sheet_structured):
    """
    generic-pattern sheets are small enough that the LLM transcribes cell
    values directly - so every distinct raw value should reappear somewhere
    in the output. A value that vanished is a silent data-loss bug, not a
    stylistic difference, so it's reported by name.
    """
    raw_values = set()
    for row in sheet_raw.get("rows", []) or []:
        value = row.get("value")
        if value is not None and value != "":
            raw_values.add(str(value))

    if not raw_values:
        return []

    output_values = {str(v) for v in _iter_leaf_values(sheet_structured) if v is not None}
    missing = sorted(raw_values - output_values)

    if not missing:
        return []

    ratio = len(missing) / len(raw_values)
    return [
        f"sheet '{sheet_name}': {len(missing)} of {len(raw_values)} distinct non-blank raw cell "
        f"value(s) ({ratio:.0%}) do not appear anywhere in the structured output - sample: "
        f"{missing[:8]}"
    ]


def validate_sheet_structured_output(sheet_name, sheet_raw, sheet_structured):
    """Returns a list of human-readable warning strings (empty if nothing found)."""
    if not isinstance(sheet_structured, dict):
        return [
            f"sheet '{sheet_name}': structured output is not a JSON object "
            f"(got {type(sheet_structured).__name__})"
        ]

    warnings = _validate_dates(sheet_name, sheet_raw, sheet_structured)

    if sheet_structured.get("pattern") == "generic":
        warnings += _validate_generic_completeness(sheet_name, sheet_raw, sheet_structured)

    return warnings
