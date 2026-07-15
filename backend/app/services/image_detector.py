"""
Embedded-image detection by reading the .xlsx package directly.

openpyxl only surfaces `worksheet._images` when Pillow is installed; without
it, image loading is silently skipped and a sheet that contains a stamped
watermark or a pasted screenshot looks identical to an empty one. That false
"no images" is exactly the blind spot this is meant to remove, so detection
here does NOT go through openpyxl/Pillow at all - it parses the OOXML drawing
parts straight out of the zip:

  xl/workbook.xml (+ rels)        -> sheet name  -> xl/worksheets/sheetN.xml
  xl/worksheets/_rels/*.rels      -> the sheet's drawing part
  xl/drawings/drawingM.xml        -> anchors (from-cell + extent)
  xl/drawings/_rels/*.rels        -> the media file each anchor references

Location and size only - never the image's content (OCR/vision is out of
scope). Deterministic and dependency-free.
"""
import posixpath
import struct
import zipfile
from xml.etree import ElementTree as ET

from openpyxl.utils import get_column_letter

_SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PR = "http://schemas.openxmlformats.org/package/2006/relationships"


def _read(zf, name):
    try:
        return zf.read(name)
    except KeyError:
        return None


def _rels(zf, part_path):
    """Return [(id, target, type)] from a part's .rels, or []."""
    base, fname = posixpath.split(part_path)
    rels_path = posixpath.join(base, "_rels", fname + ".rels")
    data = _read(zf, rels_path)
    if data is None:
        return []
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []
    out = []
    for rel in root.findall(f"{{{_PR}}}Relationship"):
        out.append((rel.get("Id"), rel.get("Target"), rel.get("Type")))
    return out


def _resolve(base_dir, target):
    return posixpath.normpath(posixpath.join(base_dir, target))


def _image_dimensions(data):
    """
    Native (intrinsic) pixel dimensions read straight from the image file
    header - no Pillow. Supports PNG, GIF, BMP, and JPEG (the formats Excel
    embeds); returns (None, None) for anything unrecognized rather than
    guessing. This gives the image's true resolution, which is more useful
    and more reliable than the drawing's on-sheet display extent (often
    absent for two-cell anchors).
    """
    if data is None or len(data) < 24:
        return None, None
    # PNG: 8-byte signature, then IHDR width/height as big-endian uint32.
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return struct.unpack(">II", data[16:24])
    # GIF: logical screen width/height as little-endian uint16.
    if data[:6] in (b"GIF87a", b"GIF89a"):
        w, h = struct.unpack("<HH", data[6:10])
        return w, h
    # BMP: DIB header width/height as little-endian int32.
    if data[:2] == b"BM":
        w, h = struct.unpack("<ii", data[18:26])
        return abs(w), abs(h)
    # JPEG: scan segments for a Start-Of-Frame marker carrying dimensions.
    if data[:2] == b"\xff\xd8":
        i = 2
        n = len(data)
        while i + 9 < n:
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                h, w = struct.unpack(">HH", data[i + 5 : i + 9])
                return w, h
            seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
            i += 2 + seg_len
    return None, None


def _anchor_images(zf, drawing_path):
    data = _read(zf, drawing_path)
    if data is None:
        return []
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []

    # rId -> resolved media path inside the package (from the drawing's rels).
    base_dir = posixpath.dirname(drawing_path)
    embed_media = {}
    for rid, target, _type in _rels(zf, drawing_path):
        if target:
            embed_media[rid] = _resolve(base_dir, target)

    dim_cache = {}

    def _media_info(media_path):
        if media_path not in dim_cache:
            fmt = posixpath.splitext(media_path)[1].lstrip(".").lower() or None
            w, h = _image_dimensions(_read(zf, media_path))
            dim_cache[media_path] = (fmt, w, h)
        return dim_cache[media_path]

    images = []
    for anchor in list(root):
        tag = anchor.tag.split("}")[-1]
        if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            continue

        anchor_cell = None
        frm = anchor.find(f"{{{_XDR}}}from")
        if frm is not None:
            col_el = frm.find(f"{{{_XDR}}}col")
            row_el = frm.find(f"{{{_XDR}}}row")
            try:
                col = int(col_el.text)
                row = int(row_el.text)
                anchor_cell = f"{get_column_letter(col + 1)}{row + 1}"
            except (AttributeError, TypeError, ValueError):
                anchor_cell = None

        fmt = width_px = height_px = None
        blip = anchor.find(f".//{{{_A}}}blip")
        if blip is not None:
            media_path = embed_media.get(blip.get(f"{{{_R}}}embed"))
            if media_path:
                fmt, width_px, height_px = _media_info(media_path)

        images.append(
            {
                "anchor_cell": anchor_cell,
                "width_px": width_px,
                "height_px": height_px,
                "format": fmt,
            }
        )
    return images


def detect_images_by_sheet(file_path):
    """
    Returns {sheet_name: [ {anchor_cell, width_px, height_px, format}, ... ]}
    for every sheet, including sheets with no images (empty list).
    """
    result = {}
    try:
        zf = zipfile.ZipFile(file_path)
    except (zipfile.BadZipFile, OSError):
        return result

    with zf:
        wb_xml = _read(zf, "xl/workbook.xml")
        if wb_xml is None:
            return result
        try:
            wb_root = ET.fromstring(wb_xml)
        except ET.ParseError:
            return result

        rid_to_target = {rid: target for rid, target, _t in _rels(zf, "xl/workbook.xml")}

        for sheet_el in wb_root.findall(f"{{{_SML}}}sheets/{{{_SML}}}sheet"):
            name = sheet_el.get("name")
            rid = sheet_el.get(f"{{{_R}}}id")
            result[name] = []
            target = rid_to_target.get(rid)
            if not target:
                continue
            ws_path = _resolve("xl", target)

            drawing_path = None
            for _rid, dtarget, dtype in _rels(zf, ws_path):
                if dtype and dtype.endswith("/drawing"):
                    drawing_path = _resolve(posixpath.dirname(ws_path), dtarget)
                    break
            if drawing_path is None:
                continue

            result[name] = _anchor_images(zf, drawing_path)

    return result
