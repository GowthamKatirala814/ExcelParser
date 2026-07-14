"""
Conditional-formatting evaluation.

We only ever declare a rule "matches" when it can be evaluated with total
certainty from literal operands in the rule itself (e.g. `cellIs greaterThan 100`,
or `containsText` with a literal search string). Anything that requires
evaluating a formula/relative-reference/aggregate (expression rules, colorScale,
dataBar, iconSet, or a cellIs rule whose operand is itself a formula) is
reported as "not evaluated" rather than assumed true or false.

When multiple rules cover the same cell, we resolve them in the workbook's own
rule priority order (lower `priority` = evaluated first, matching Excel's own
precedence): the first rule that definitely evaluates true wins outright. If a
rule that cannot be evaluated is encountered before any definite winner, we
cannot know whether it (or something after it) would actually win, so the cell
is reported unresolved with the exact reason instead of guessed.
"""
import re

from .color_resolver import resolve_color_object

_NUMERIC_RE = re.compile(r"^-?\d+(\.\d+)?$")
_CELL_REF_RE = re.compile(r"[A-Za-z]+\$?\d")


def _literal_operand(raw):
    """Return (is_literal, python_value) for a formula operand string."""
    if raw is None:
        return False, None
    text = raw.strip()
    if text.startswith('"') and text.endswith('"'):
        return True, text[1:-1]
    if _NUMERIC_RE.match(text):
        return True, float(text)
    if _CELL_REF_RE.search(text):
        return False, None
    if "(" in text:
        return False, None
    return True, text


def _evaluate_cell_is(rule, cell_value):
    """
    Returns True/False if determinable with certainty, or None if the rule
    cannot be safely evaluated (non-literal operand).
    """
    operator = getattr(rule, "operator", None)
    formulas = list(getattr(rule, "formula", None) or [])
    if cell_value is None or not formulas:
        return None

    operands = []
    for f in formulas:
        is_literal, val = _literal_operand(f)
        if not is_literal:
            return None
        operands.append(val)

    try:
        if isinstance(operands[0], (int, float)) and not isinstance(cell_value, bool):
            left = float(cell_value)
        else:
            left = cell_value
    except (TypeError, ValueError):
        return None

    try:
        if operator in ("equal", "notEqual"):
            eq = left == operands[0]
            return eq if operator == "equal" else not eq
        if operator == "greaterThan":
            return left > operands[0]
        if operator == "greaterThanOrEqual":
            return left >= operands[0]
        if operator == "lessThan":
            return left < operands[0]
        if operator == "lessThanOrEqual":
            return left <= operands[0]
        if operator == "between" and len(operands) == 2:
            return operands[0] <= left <= operands[1]
        if operator == "notBetween" and len(operands) == 2:
            return not (operands[0] <= left <= operands[1])
    except TypeError:
        return None

    # Any other operator (beginsWith, timePeriod, etc.) is not implemented;
    # be honest rather than silently mis-evaluating it.
    return None


def _evaluate_contains_text(rule, cell_value):
    """
    containsText rules carry the literal search string on rule.text (parsed
    directly from the XML attribute, not from the formula). A case-insensitive
    substring check against the cell's actual value is unambiguous.
    """
    search_text = getattr(rule, "text", None)
    if not search_text:
        return None
    if cell_value is None:
        return False
    return search_text.lower() in str(cell_value).lower()


def _evaluate_rule(rule, cell_value):
    """Returns True/False if determinable with certainty, else None."""
    rtype = getattr(rule, "type", None)
    if rtype == "cellIs":
        return _evaluate_cell_is(rule, cell_value)
    if rtype == "containsText":
        return _evaluate_contains_text(rule, cell_value)
    return None


def _get_dxf(rule, workbook):
    dxf = getattr(rule, "dxf", None)
    if dxf is not None:
        return dxf
    dxf_id = getattr(rule, "dxfId", None)
    if dxf_id is None:
        return None
    try:
        styles = workbook._differential_styles.styles
        if 0 <= dxf_id < len(styles):
            return styles[dxf_id]
    except AttributeError:
        return None
    return None


def _dxf_fill_color(rule, workbook, theme_palette, theme_error):
    dxf = _get_dxf(rule, workbook)
    if dxf is None or getattr(dxf, "fill", None) is None:
        return {
            "resolved": False,
            "hex": None,
            "theme_ref": None,
            "ambiguous": False,
            "reason_unresolved": "matching conditional formatting rule defines no fill",
        }
    fg = getattr(dxf.fill, "fgColor", None) or getattr(dxf.fill, "bgColor", None)
    result = resolve_color_object(fg, theme_palette, theme_error)
    result["ambiguous"] = False
    return result


def resolve_conditional_format_color(
    cell, cell_value, worksheet, workbook, theme_palette=None, theme_error=None
):
    """
    Returns None if no conditional-formatting range covers this cell, or a
    resolve_color_object()-shaped dict (plus 'ambiguous') describing what was
    found.
    """
    coordinate = cell.coordinate
    covering_rules = []  # (priority, rule)

    for cf_range in worksheet.conditional_formatting:
        try:
            covers = coordinate in cf_range
        except TypeError:
            covers = False
        if not covers:
            continue

        for rule in cf_range.rules:
            priority = getattr(rule, "priority", None)
            covering_rules.append((priority if priority is not None else 0, rule))

    if not covering_rules:
        return None

    # Excel evaluates conditional-formatting rules in ascending priority order
    # (lower number = higher precedence) and the first true rule determines
    # the rendered fill.
    covering_rules.sort(key=lambda pair: pair[0])

    for priority, rule in covering_rules:
        verdict = _evaluate_rule(rule, cell_value)
        if verdict is True:
            return _dxf_fill_color(rule, workbook, theme_palette, theme_error)
        if verdict is False:
            continue

        # verdict is None: this rule precedes (in priority order) any rule
        # not yet examined, so we cannot know whether it - or something after
        # it - actually wins. Report exactly why instead of guessing.
        return {
            "resolved": False,
            "hex": None,
            "theme_ref": None,
            "ambiguous": False,
            "reason_unresolved": (
                f"conditional formatting rule type '{getattr(rule, 'type', 'unknown')}' at "
                f"priority {priority} covers this cell but requires formula/relative-reference/"
                "scale evaluation that was not performed, and it precedes other rules in "
                "priority order so the winning format cannot be determined"
            ),
        }

    # All covering rules were definitely evaluated and none matched: no CF fill applies.
    return None
