"""
Gemini implementation of StructuringProvider. Talks to the Gemini
generateContent REST API directly (no SDK dependency) so the only thing
another provider needs to replace is this one file plus an entry in
factory.py.

Called once per sheet (the caller passes a raw_json containing exactly one
sheet) - this keeps every request small and independent, so one sheet
hitting a token limit or a malformed response can never affect any other
sheet's result.

Key rotation: this provider is handed a shared ApiKeyPool (see key_pool.py)
rather than a single key. On a key-specific failure (quota exhausted, rate
limited, invalid key) it puts that key on cooldown and immediately retries
the same request with the next available key, transparently to the caller -
a sheet only fails once every configured key has been tried. Non-key-specific
failures (timeout, network error, malformed response) are not retried with a
different key, since swapping keys can't fix those.

Every error surfaced to the caller (and from there, to the UI) is a short,
human-readable sentence - no raw HTTP bodies, no stack traces. Full
diagnostic detail (status codes, response bodies, exception reprs) is only
ever written to the server-side logger, never returned to the frontend.
"""
import asyncio
import json
import logging
import re

import httpx

from .base import (
    ERROR_KIND_INVALID_JSON,
    ERROR_KIND_INVALID_KEY,
    ERROR_KIND_MALFORMED_RESPONSE,
    ERROR_KIND_MODEL_UNAVAILABLE,
    ERROR_KIND_NETWORK,
    ERROR_KIND_OTHER,
    ERROR_KIND_QUOTA_EXCEEDED,
    ERROR_KIND_RATE_LIMITED,
    ERROR_KIND_TIMEOUT,
    ERROR_KIND_TOKEN_LIMIT,
    StructuringError,
    StructuringProvider,
)

logger = logging.getLogger("app.structuring.gemini")

_INSTRUCTIONS = """You are given a size-reduced projection of one sheet from a raw Excel \
workbook extraction. You get:
  - `cells`: only the non-blank/non-colorless cells, each with a `ref` (its exact cell \
    reference, e.g. "B12" - column letter(s) + row number, exactly as Excel shows it) plus \
    whichever of value/fill/comment/formula/merged_with/human_label actually apply. A cell \
    reference not listed is blank.
  - `detected_tables`: an already-computed structural layout per table - its row_range,
    header_row (or null), date_row/column_index if a horizontal date header was found
    (column_index maps a column letter to an actual calendar date), and an
    ambiguous/ambiguous_reason flag when detection was uncertain.
  - `color_inventory`, `total_cells`, `colored_cells`, `blank_cells`: sheet-level stats.

IMPORTANT - what your job is NOT: you are NOT asked to collapse a calendar grid into date
ranges, count consecutive days, or transcribe grid values yourself. That arithmetic is done
deterministically by code afterward directly from the trusted cell data, specifically to avoid
the transcription/counting errors an LLM makes on a dense grid of dozens of near-identical
cells. Your job is purely to identify what the sheet *means* - the parts that genuinely
require judgment, not mechanical counting.

Work out from this sheet's actual content (headers, a legend/key section if present, merged
cells, fill colors, row/column shape) whether it matches this pattern - never force it onto a
sheet that doesn't actually have it:

  - A legend/key section mapping colors or short codes to their meaning (e.g. a color swatch
    next to a label like "Closed").
  - An entity master table: rows that each describe one named/coded entity (a code, a name,
    maybe a count/attribute column).
  - A calendar/availability grid: columns are dates (via `column_index`), rows are entities.

If the sheet matches this pattern, respond with:
  {
    "pattern": "entity_calendar",
    "legend": {"<code_or_hex_as_it_appears_in_cells>": "<meaning>", ...},
    "default_status": "<the single most common meaning across the grid - the fallback when
       nothing else applies>",
    "entities": {"<code>": {"name": "...", "attributes": {"<other master-table column
       name>": <value>, ...}}},
    "row_to_entity": {"<absolute row number as a string, e.g. \\"13\\">": "<code>"}
       for every row across every calendar-grid table in this sheet that represents an entity
       (skip header/label rows)
  }
`row_to_entity` is the single most important field: for each row of a calendar grid, identify
which entity code that row belongs to (usually readable from a cell earlier in that same row,
or from the entity master table) - code will use it to pull the actual per-day values straight
from the trusted cell data, so it must cover every entity row in every table on the sheet.

If the sheet does NOT match this pattern (plain data, free-form notes, metadata), respond with:
  {"pattern": "generic", "tables": [{"title": "...", "columns": {...}, "rows": [...]}]}
(For "generic" sheets only, you may transcribe cell values directly since they are small.)
Every non-blank cell listed in `cells` should be represented somewhere in your "generic" tables
output - do not silently drop rows or columns from the transcription.

Rules, all mandatory:

A. `ref` and `column_index` are ground truth for location and dates - never renumber, reindex,
   or invent a reference or date that isn't directly derivable from the input.
B. Never invent, guess, or hallucinate an entity, a legend entry, or an attribute that cannot
   be derived directly from the input.
C. If something is genuinely unknown or ambiguous, use null or the string "unknown" - never a
   plausible-sounding guess.
D. Respond with ONLY the JSON object itself - no markdown code fences, no commentary before or
   after it.
"""

_FENCE_RE = re.compile(r"^```(?:json)?\s*|\s*```$", re.IGNORECASE | re.MULTILINE)

_RETRYABLE_STATUS = {500, 502, 503, 504}

_NON_STOP_FINISH_REASONS = {
    "MAX_TOKENS": (ERROR_KIND_TOKEN_LIMIT, "Gemini's response was too large and got cut off"),
    "SAFETY": (ERROR_KIND_MALFORMED_RESPONSE, "Gemini declined to respond"),
    "RECITATION": (ERROR_KIND_MALFORMED_RESPONSE, "Gemini declined to respond"),
}


class _KeyExhausted(Exception):
    """Signals a key-specific failure (quota/rate-limit/invalid key): the
    caller should put this key on cooldown and try the next one."""

    def __init__(self, structuring_error, reason):
        super().__init__(reason)
        self.structuring_error = structuring_error
        self.reason = reason


def _repair_unclosed_brackets(text):
    """
    Some responses arrive with one or more trailing `}`/`]` missing even
    though the API reports finishReason=STOP (observed with
    gemini-flash-latest on small responses) - the content itself is intact,
    only the final closing punctuation is absent. This appends exactly the
    closers needed to balance whatever braces/brackets are already open,
    tracking string/escape state so a `{`/`[` inside a quoted value is never
    mistaken for real structure. Returns None if the text isn't a simple
    "missing closers" case (nothing to safely repair) - never invents or
    alters any actual content, only appends punctuation.
    """
    stack = []
    in_string = False
    escape = False
    for ch in text:
        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch in "{[":
            stack.append(ch)
        elif ch in "}]":
            if stack:
                stack.pop()

    if not stack or in_string:
        return None

    closers = {"{": "}", "[": "]"}
    return text + "".join(closers[opener] for opener in reversed(stack))


def _strip_code_fences(text: str) -> str:
    """
    Defensive: responseMimeType=application/json should prevent this, but
    strip a leading/trailing ``` or ```json fence if the model added one
    anyway, rather than failing a JSON parse that would otherwise succeed.
    """
    stripped = text.strip()
    if stripped.startswith("```"):
        stripped = _FENCE_RE.sub("", stripped).strip()
    return stripped


def _classify_http_error(exc: httpx.HTTPError) -> StructuringError:
    if isinstance(exc, httpx.TimeoutException):
        return StructuringError("Request timed out", kind=ERROR_KIND_TIMEOUT)
    return StructuringError("Network error", kind=ERROR_KIND_NETWORK)


def _classify_http_status(status_code, body_text):
    """Returns (StructuringError, is_key_specific). Key-specific failures are
    what trigger rotating to the next configured key."""
    lower = body_text.lower()
    if status_code == 429:
        if "quota" in lower or "resource_exhausted" in lower:
            return StructuringError("Gemini quota reached", kind=ERROR_KIND_QUOTA_EXCEEDED), True
        return StructuringError("Rate limit exceeded", kind=ERROR_KIND_RATE_LIMITED), True
    if status_code in (401, 403):
        return StructuringError("Invalid API key", kind=ERROR_KIND_INVALID_KEY), True
    if status_code == 404:
        # Model access is granted per API key/project, not just per model
        # name - the same model can be reachable on one key's project and
        # blocked ("no longer available to new users", or simply not yet
        # enabled) on another's. That makes this a key-specific condition:
        # worth trying the next key rather than failing the whole sheet.
        return StructuringError("Model not available for this API key", kind=ERROR_KIND_MODEL_UNAVAILABLE), True
    if status_code in _RETRYABLE_STATUS:
        return StructuringError("Gemini service error - please try again", kind=ERROR_KIND_OTHER), False
    return StructuringError(f"Gemini request failed (HTTP {status_code})", kind=ERROR_KIND_OTHER), False


class GeminiStructuringProvider(StructuringProvider):
    def __init__(self, key_pool, model, api_base, max_input_bytes, timeout=300.0, max_retries=2):
        self._key_pool = key_pool
        self._model = model
        self._api_base = api_base.rstrip("/")
        self._max_input_bytes = max_input_bytes
        self._timeout = timeout
        self._max_retries = max_retries

    async def _post_once(self, url, request_body, headers):
        async with httpx.AsyncClient(timeout=self._timeout) as client:
            return await client.post(url, json=request_body, headers=headers)

    async def _attempt_with_key(self, url, request_body, headers):
        """
        One key's full attempt, including same-key retries for transient,
        non-key-specific failures (5xx, network hiccups). Raises
        StructuringError directly for a fatal, non-key-specific problem - the
        caller does not rotate keys for these. Raises _KeyExhausted for a
        key-specific problem so the caller can cool this key down and move
        on to the next one.
        """
        response = None
        for attempt in range(1, self._max_retries + 2):
            try:
                response = await self._post_once(url, request_body, headers)
            except httpx.HTTPError as exc:
                logger.warning(
                    "gemini request attempt %d/%d failed: %s: %s",
                    attempt, self._max_retries + 1, type(exc).__name__, exc,
                )
                if attempt <= self._max_retries:
                    await asyncio.sleep(2 ** (attempt - 1))
                    continue
                raise _classify_http_error(exc) from exc

            if response.status_code == 200:
                break

            structuring_error, is_key_specific = _classify_http_status(response.status_code, response.text)
            logger.warning(
                "gemini HTTP %d on attempt %d/%d (key_specific=%s): %s",
                response.status_code, attempt, self._max_retries + 1, is_key_specific, response.text[:800],
            )
            if is_key_specific:
                raise _KeyExhausted(structuring_error, reason=str(structuring_error))
            if response.status_code in _RETRYABLE_STATUS and attempt <= self._max_retries:
                await asyncio.sleep(2 ** (attempt - 1))
                continue
            raise structuring_error

        return self._parse_response(response)

    def _parse_response(self, response):
        try:
            body = response.json()
        except json.JSONDecodeError as exc:
            logger.error("gemini response envelope was not valid JSON: %s", response.text[:800])
            raise StructuringError("Gemini returned an unreadable response", kind=ERROR_KIND_MALFORMED_RESPONSE) from exc

        try:
            candidate = body["candidates"][0]
            text = candidate["content"]["parts"][0]["text"]
        except (KeyError, IndexError, TypeError) as exc:
            finish_reason = None
            if body.get("candidates"):
                finish_reason = body["candidates"][0].get("finishReason")
            if finish_reason in _NON_STOP_FINISH_REASONS:
                kind, message = _NON_STOP_FINISH_REASONS[finish_reason]
                raise StructuringError(message, kind=kind) from exc
            logger.error("gemini response had unexpected shape (finishReason=%s): %s", finish_reason, json.dumps(body)[:1000])
            raise StructuringError("Gemini returned an unreadable response", kind=ERROR_KIND_MALFORMED_RESPONSE) from exc

        finish_reason = candidate.get("finishReason")
        if finish_reason in _NON_STOP_FINISH_REASONS:
            kind, message = _NON_STOP_FINISH_REASONS[finish_reason]
            logger.warning("gemini finishReason=%s", finish_reason)
            raise StructuringError(message, kind=kind)
        if finish_reason not in (None, "STOP"):
            logger.warning("gemini finishReason=%s (response may be affected)", finish_reason)

        cleaned = _strip_code_fences(text)
        try:
            parsed = json.loads(cleaned)
        except json.JSONDecodeError as exc:
            repaired = _repair_unclosed_brackets(cleaned)
            if repaired is not None:
                try:
                    parsed = json.loads(repaired)
                    logger.warning(
                        "gemini output was missing closing bracket(s) despite finishReason=%s; "
                        "auto-repaired (appended closers only, no content invented) and re-parsed OK",
                        finish_reason,
                    )
                except json.JSONDecodeError:
                    logger.error(
                        "gemini output failed JSON parse even after bracket repair (finishReason=%s): %s",
                        finish_reason, cleaned[:2000],
                    )
                    raise StructuringError("Gemini returned an unreadable response", kind=ERROR_KIND_INVALID_JSON) from exc
            else:
                logger.error(
                    "gemini output failed JSON parse (finishReason=%s): %s", finish_reason, cleaned[:2000]
                )
                raise StructuringError("Gemini returned an unreadable response", kind=ERROR_KIND_INVALID_JSON) from exc

        if isinstance(parsed, list) and len(parsed) == 1 and isinstance(parsed[0], dict):
            # Benign, safely-unwrappable deviation: the model wrapped the
            # single expected object in a one-element array.
            logger.warning("gemini wrapped its response in a single-element array; unwrapping")
            parsed = parsed[0]

        if not isinstance(parsed, dict):
            logger.error("gemini returned non-object JSON: %s", json.dumps(parsed, default=str)[:500])
            raise StructuringError("Gemini returned an unreadable response", kind=ERROR_KIND_INVALID_JSON)

        return parsed

    async def structure(self, raw_json: dict) -> dict:
        payload_text = json.dumps(raw_json, default=str)
        payload_bytes = len(payload_text.encode("utf-8"))
        logger.info("gemini structuring: payload=%d bytes, model=%s", payload_bytes, self._model)

        if payload_bytes > self._max_input_bytes:
            logger.error(
                "payload too large: %d bytes > limit %d bytes", payload_bytes, self._max_input_bytes
            )
            raise StructuringError("Sheet is too large to send to Gemini", kind=ERROR_KIND_TOKEN_LIMIT)

        url = f"{self._api_base}/models/{self._model}:generateContent"
        request_body = {
            "contents": [
                {
                    "role": "user",
                    "parts": [{"text": _INSTRUCTIONS}, {"text": payload_text}],
                }
            ],
            "generationConfig": {"responseMimeType": "application/json"},
        }

        available = self._key_pool.available_keys()
        if not available:
            logger.error(
                "no Gemini API keys available (%d configured, all on cooldown)",
                self._key_pool.total_keys(),
            )
            raise StructuringError("Gemini quota reached", kind=ERROR_KIND_QUOTA_EXCEEDED)

        last_error = None
        for index, key in enumerate(available, start=1):
            headers = {"x-goog-api-key": key, "Content-Type": "application/json"}
            try:
                result = await self._attempt_with_key(url, request_body, headers)
            except _KeyExhausted as exhausted:
                self._key_pool.mark_cooldown(key, exhausted.reason)
                last_error = exhausted.structuring_error
                logger.warning(
                    "key #%d/%d exhausted (%s); trying next key", index, len(available), exhausted.reason
                )
                continue

            logger.info("gemini structuring succeeded using key #%d of %d available", index, len(available))
            return result

        # Every available key failed with a key-specific error.
        raise last_error or StructuringError("Gemini quota reached", kind=ERROR_KIND_QUOTA_EXCEEDED)
