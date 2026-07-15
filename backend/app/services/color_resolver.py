"""
Raw color-fact resolution.

Every function here returns exactly what is *provably* present in the file.
Nothing is guessed. When a color cannot be resolved to an exact RGB value
(theme colors without a further theme-XML lookup, "auto" colors, indexed
system colors), the result honestly says so with a reason instead of
substituting a default or a best guess.
"""
from openpyxl.styles.colors import COLOR_INDEX

from .theme_resolver import resolve_theme_color

# Indexed color slots that Excel treats as "system" / "automatic" rather than
# an explicit author-chosen color. These are not resolvable to a meaningful
# fixed RGB without knowing the user's system/theme, so we report them honestly.
INDEXED_AUTO_SLOTS = {64, 65}


def _strip_alpha(argb):
    if not argb or not isinstance(argb, str):
        return None
    hex_str = argb[-6:]
    if len(hex_str) != 6:
        return None
    return hex_str.upper()


def resolve_color_object(color, theme_palette=None, theme_error=None):
    """
    Resolve an openpyxl Color object to a raw fact dict:
      {
        "resolved": bool,
        "hex": "RRGGBB" or None,
        "theme_ref": "theme:<idx>:tint<tint>" or None,
        "reason_unresolved": str or None,
      }

    theme_palette/theme_error come from theme_resolver.load_theme_palette(),
    loaded once per workbook, so "theme" colors can be resolved to the exact
    RGB Excel would render instead of being reported unresolved.
    """
    if color is None:
        return {
            "resolved": False,
            "hex": None,
            "theme_ref": None,
            "reason_unresolved": "no color information present",
        }

    ctype = getattr(color, "type", None)

    if ctype == "rgb":
        rgb = getattr(color, "rgb", None)
        if not rgb or not isinstance(rgb, str):
            return {
                "resolved": False,
                "hex": None,
                "theme_ref": None,
                "reason_unresolved": "rgb color type but no rgb value present",
            }
        alpha = rgb[:2] if len(rgb) == 8 else "FF"
        hex_str = _strip_alpha(rgb)
        if hex_str is None:
            return {
                "resolved": False,
                "hex": None,
                "theme_ref": None,
                "reason_unresolved": f"malformed rgb value '{rgb}'",
            }
        if alpha == "00":
            # Fully transparent explicit color -> not a visible fill.
            return {
                "resolved": False,
                "hex": None,
                "theme_ref": None,
                "reason_unresolved": "explicit color has zero alpha (transparent)",
            }
        return {
            "resolved": True,
            "hex": hex_str,
            "theme_ref": None,
            "reason_unresolved": None,
        }

    if ctype == "theme":
        theme_idx = getattr(color, "theme", None)
        tint = getattr(color, "tint", 0.0) or 0.0
        theme_ref = f"theme:{theme_idx}:tint{tint}"

        if theme_palette is not None:
            hex_str = resolve_theme_color(theme_idx, tint, theme_palette)
            if hex_str is not None:
                return {
                    "resolved": True,
                    "hex": hex_str,
                    "theme_ref": theme_ref,
                    "reason_unresolved": None,
                }
            return {
                "resolved": False,
                "hex": None,
                "theme_ref": theme_ref,
                "reason_unresolved": (
                    f"theme index {theme_idx} has no corresponding entry in the "
                    "workbook's own theme palette"
                ),
            }

        reason = theme_error or "workbook theme could not be loaded"
        return {
            "resolved": False,
            "hex": None,
            "theme_ref": theme_ref,
            "reason_unresolved": (
                f"theme color (theme index {theme_idx}, tint {tint}) unresolved: {reason}"
            ),
        }

    if ctype == "indexed":
        idx = getattr(color, "indexed", None)
        if idx is None:
            return {
                "resolved": False,
                "hex": None,
                "theme_ref": None,
                "reason_unresolved": "indexed color type but no index present",
            }
        if idx in INDEXED_AUTO_SLOTS:
            return {
                "resolved": False,
                "hex": None,
                "theme_ref": f"indexed:{idx}",
                "reason_unresolved": f"indexed color {idx} is a system/auto slot, not an explicit color",
            }
        if 0 <= idx < len(COLOR_INDEX):
            hex_str = _strip_alpha(COLOR_INDEX[idx])
            if hex_str:
                return {
                    "resolved": True,
                    "hex": hex_str,
                    "theme_ref": None,
                    "reason_unresolved": None,
                }
        return {
            "resolved": False,
            "hex": None,
            "theme_ref": f"indexed:{idx}",
            "reason_unresolved": f"indexed color {idx} is out of range of the known legacy palette",
        }

    if ctype == "auto":
        return {
            "resolved": False,
            "hex": None,
            "theme_ref": None,
            "reason_unresolved": "auto color - no explicit RGB specified",
        }

    return {
        "resolved": False,
        "hex": None,
        "theme_ref": None,
        "reason_unresolved": f"unrecognized color type '{ctype}'",
    }


def resolve_direct_fill(cell, theme_palette=None, theme_error=None):
    """
    Returns None if the cell has no direct pattern fill at all (blank/no-fill).
    Otherwise returns a resolve_color_object()-shaped dict for the fill's
    foreground color, since that is the color actually rendered for solid fills.
    """
    fill = getattr(cell, "fill", None)
    if fill is None:
        return None
    pattern_type = getattr(fill, "patternType", None)
    if pattern_type is None:
        return None
    fg = getattr(fill, "fgColor", None)
    result = resolve_color_object(fg, theme_palette, theme_error)
    result["pattern_type"] = pattern_type
    return result


# --- Column / row default fill inheritance --------------------------------
#
# Excel lets a fill be applied at the column level (<col style="N">) or row
# level (<row s="N" customFormat="1">) rather than on each cell. A cell with
# no style of its own then RENDERS with that default fill, even though its own
# `cell.fill` is empty. Without this, such cells look blank to the extractor
# even though a human plainly sees them colored. `dim.style` is an index into
# the workbook's cellXfs table (openpyxl `_cell_styles`), which points at a
# fill in `_fills`. These are the same internals openpyxl uses itself; guarded
# so an unexpected shape degrades to "no default" rather than raising.


def _style_array_fill(workbook, style_id):
    """Given a cellXfs index, return the resolved PatternFill's fgColor,
    or None if that style has no solid fill."""
    try:
        style_array = workbook._cell_styles[int(style_id)]
        fill = workbook._fills[style_array.fillId]
    except (IndexError, AttributeError, TypeError, ValueError):
        return None
    if getattr(fill, "patternType", None) != "solid":
        return None
    return getattr(fill, "fgColor", None)


def resolve_column_default_fill(cell, worksheet, workbook, theme_palette=None, theme_error=None):
    """
    Falls back to the column's <col style="..."> default fill. The caller is
    responsible for only invoking this when the cell has no explicit style of
    its own (style_id == 0) - an explicit per-cell style always wins in Excel,
    even one that resolves to "no fill".
    """
    col_idx = cell.column
    for dim in worksheet.column_dimensions.values():
        if dim.min <= col_idx <= dim.max and dim.style not in (None, 0, "0"):
            fg = _style_array_fill(workbook, dim.style)
            if fg is not None:
                result = resolve_color_object(fg, theme_palette, theme_error)
                result["pattern_type"] = "solid"
                return result
    return None


def resolve_row_default_fill(cell, worksheet, workbook, theme_palette=None, theme_error=None):
    """Same fallback, for a row's <row s="..." customFormat="1"> default."""
    row_dim = worksheet.row_dimensions.get(cell.row)
    if row_dim is None or not getattr(row_dim, "customFormat", False):
        return None
    style_id = row_dim.style
    if style_id in (None, 0, "0"):
        return None
    fg = _style_array_fill(workbook, style_id)
    if fg is None:
        return None
    result = resolve_color_object(fg, theme_palette, theme_error)
    result["pattern_type"] = "solid"
    return result
