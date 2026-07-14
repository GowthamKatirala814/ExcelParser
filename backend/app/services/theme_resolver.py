"""
Theme color resolution.

Parses the workbook's xl/theme/themeN.xml part to get the real RGB palette
(dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) and applies the standard
OOXML tint/shade formula to resolve a theme:index:tint reference into a
final hex color. Only reports failure when the theme file is genuinely
missing or malformed - never a guessed default palette.
"""
import colorsys
import re
import zipfile
from xml.etree import ElementTree as ET

_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"a": _DRAWING_NS}

_SCHEME_ENTRIES = {
    "dk1",
    "lt1",
    "dk2",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
}

# The theme index used on cell styles (color.theme) does NOT match the
# authoring order of <a:clrScheme>. Excel/OOXML swap the first two pairs:
# index 0/1 map to lt1/dk1 (not dk1/lt1) and 2/3 map to lt2/dk2. This is a
# well-documented quirk of how Excel applies themeElements to cell styles.
THEME_INDEX_TO_NAME = {
    0: "lt1",
    1: "dk1",
    2: "lt2",
    3: "dk2",
    4: "accent1",
    5: "accent2",
    6: "accent3",
    7: "accent4",
    8: "accent5",
    9: "accent6",
    10: "hlink",
    11: "folHlink",
}


def load_theme_palette(file_path):
    """
    Returns (palette, error):
      palette: dict[str, str] mapping clrScheme name -> "RRGGBB", or None
      error: human-readable reason the theme could not be loaded, or None
    """
    try:
        with zipfile.ZipFile(file_path) as zf:
            theme_names = sorted(
                n for n in zf.namelist() if re.match(r"xl/theme/theme\d*\.xml$", n)
            )
            if not theme_names:
                return None, "workbook contains no xl/theme/themeN.xml part"
            with zf.open(theme_names[0]) as f:
                tree = ET.parse(f)
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as exc:
        return None, f"theme file missing or malformed: {exc}"

    root = tree.getroot()
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is None:
        return None, "theme file has no clrScheme element"

    palette = {}
    for child in scheme:
        tag = child.tag.split("}")[-1]
        if tag not in _SCHEME_ENTRIES:
            continue
        srgb = child.find("a:srgbClr", NS)
        sysclr = child.find("a:sysClr", NS)
        if srgb is not None and srgb.get("val"):
            palette[tag] = srgb.get("val").upper()
        elif sysclr is not None and sysclr.get("lastClr"):
            palette[tag] = sysclr.get("lastClr").upper()

    missing = _SCHEME_ENTRIES - palette.keys()
    if missing:
        return None, f"theme clrScheme is missing entries: {sorted(missing)}"

    return palette, None


def apply_tint(hex_rgb, tint):
    """Standard OOXML tint/shade formula, applied in HSL space."""
    if not tint:
        return hex_rgb
    r = int(hex_rgb[0:2], 16) / 255.0
    g = int(hex_rgb[2:4], 16) / 255.0
    b = int(hex_rgb[4:6], 16) / 255.0
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint < 0:
        l = l * (1.0 + tint)
    else:
        l = l * (1.0 - tint) + tint
    l = min(1.0, max(0.0, l))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return f"{round(r * 255):02X}{round(g * 255):02X}{round(b * 255):02X}"


def resolve_theme_color(theme_index, tint, palette):
    """Returns "RRGGBB" or None if theme_index has no known mapping/palette entry."""
    name = THEME_INDEX_TO_NAME.get(theme_index)
    if name is None or palette is None or name not in palette:
        return None
    return apply_tint(palette[name], tint or 0.0)
