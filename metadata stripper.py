#!/usr/bin/env python3
"""
MetadataStripper - أداة لحذف الميتاداتا والمعلومات الشخصية (PII) من الملفات.

هذه النسخة لا تعتمد على أي مكتبة خارجية (لا PIL ولا pypdf) وتستخدم فقط
مكتبة بايثون القياسية: argparse, os, sys, struct, re, zipfile, pathlib.
"""

__version__ = "2.0.0"

import argparse
import os
import re
import struct
import sys
import zipfile
from pathlib import Path

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".webp"}
PDF_EXTS = {".pdf"}
OFFICE_EXTS = {".docx", ".xlsx", ".pptx"}

OFFICE_METADATA_PARTS = {
    "docProps/core.xml",
    "docProps/app.xml",
    "docProps/custom.xml",
}

# ---------------------------------------------------------------------------
# TIFF / EXIF parsing helpers (used by JPEG APP1, PNG eXIf, TIFF, WEBP EXIF)
# ---------------------------------------------------------------------------

TIFF_TYPE_SIZES = {1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8}

# IFD0 ("TIFF") tags + Exif SubIFD tags + Interoperability IFD tags, merged
# into one lookup (tag numbers don't collide across these three groups in
# practice). This mirrors the common fields exiftool reports.
TAG_NAMES = {
    # --- IFD0 / TIFF ---
    0x00FE: "NewSubfileType",
    0x0100: "ImageWidth",
    0x0101: "ImageLength",
    0x0102: "BitsPerSample",
    0x0103: "Compression",
    0x0106: "PhotometricInterpretation",
    0x010E: "ImageDescription",
    0x010F: "Make",
    0x0110: "Model",
    0x0111: "StripOffsets",
    0x0112: "Orientation",
    0x0115: "SamplesPerPixel",
    0x0116: "RowsPerStrip",
    0x0117: "StripByteCounts",
    0x011A: "XResolution",
    0x011B: "YResolution",
    0x011C: "PlanarConfiguration",
    0x0128: "ResolutionUnit",
    0x0131: "Software",
    0x0132: "DateTime",
    0x013B: "Artist",
    0x013C: "HostComputer",
    0x0211: "YCbCrCoefficients",
    0x0213: "YCbCrPositioning",
    0x0214: "ReferenceBlackWhite",
    0x8298: "Copyright",
    0x8769: "ExifIFD",
    0x8825: "GPSInfo",
    0x9C9B: "XPTitle",
    0x9C9C: "XPComment",
    0x9C9D: "XPAuthor",
    0x9C9E: "XPKeywords",
    0x9C9F: "XPSubject",
    # Windows Photo "Rating"
    0x4746: "Rating",
    0x4749: "RatingPercent",
    # --- Exif SubIFD ---
    0x829A: "ExposureTime",
    0x829D: "FNumber",
    0x8822: "ExposureProgram",
    0x8824: "SpectralSensitivity",
    0x8827: "ISOSpeedRatings",
    0x8830: "SensitivityType",
    0x8832: "RecommendedExposureIndex",
    0x9000: "ExifVersion",
    0x9003: "DateTimeOriginal",
    0x9004: "DateTimeDigitized",
    0x9101: "ComponentsConfiguration",
    0x9102: "CompressedBitsPerPixel",
    0x9201: "ShutterSpeedValue",
    0x9202: "ApertureValue",
    0x9203: "BrightnessValue",
    0x9204: "ExposureBiasValue",
    0x9205: "MaxApertureValue",
    0x9206: "SubjectDistance",
    0x9207: "MeteringMode",
    0x9208: "LightSource",
    0x9209: "Flash",
    0x920A: "FocalLength",
    0x9214: "SubjectArea",
    0x927C: "MakerNote",
    0x9286: "UserComment",
    0x9290: "SubSecTime",
    0x9291: "SubSecTimeOriginal",
    0x9292: "SubSecTimeDigitized",
    0xA000: "FlashpixVersion",
    0xA001: "ColorSpace",
    0xA002: "PixelXDimension",
    0xA003: "PixelYDimension",
    0xA004: "RelatedSoundFile",
    0xA005: "InteroperabilityIFD",
    0xA20B: "FlashEnergy",
    0xA20E: "FocalPlaneXResolution",
    0xA20F: "FocalPlaneYResolution",
    0xA210: "FocalPlaneResolutionUnit",
    0xA214: "SubjectLocation",
    0xA215: "ExposureIndex",
    0xA217: "SensingMethod",
    0xA300: "FileSource",
    0xA301: "SceneType",
    0xA302: "CFAPattern",
    0xA401: "CustomRendered",
    0xA402: "ExposureMode",
    0xA403: "WhiteBalance",
    0xA404: "DigitalZoomRatio",
    0xA405: "FocalLengthIn35mmFilm",
    0xA406: "SceneCaptureType",
    0xA407: "GainControl",
    0xA408: "Contrast",
    0xA409: "Saturation",
    0xA40A: "Sharpness",
    0xA40C: "SubjectDistanceRange",
    0xA420: "ImageUniqueID",
    0xA430: "BodySerialNumber",
    0xA431: "LensSpecification",
    0xA432: "LensSpecificationRaw",
    0xA433: "LensMake",
    0xA434: "LensModel",
    0xA435: "LensSerialNumber",
    # --- Interoperability IFD ---
    0x0001: "InteropIndex",
    0x0002: "InteropVersion",
}

GPS_TAG_NAMES = {
    0x0000: "GPSVersionID",
    0x0001: "GPSLatitudeRef",
    0x0002: "GPSLatitude",
    0x0003: "GPSLongitudeRef",
    0x0004: "GPSLongitude",
    0x0005: "GPSAltitudeRef",
    0x0006: "GPSAltitude",
    0x0007: "GPSTimeStamp",
    0x0008: "GPSSatellites",
    0x0009: "GPSStatus",
    0x000A: "GPSMeasureMode",
    0x000B: "GPSDOP",
    0x000C: "GPSSpeedRef",
    0x000D: "GPSSpeed",
    0x000E: "GPSTrackRef",
    0x000F: "GPSTrack",
    0x0010: "GPSImgDirectionRef",
    0x0011: "GPSImgDirection",
    0x0012: "GPSMapDatum",
    0x0013: "GPSDestLatitudeRef",
    0x0014: "GPSDestLatitude",
    0x0015: "GPSDestLongitudeRef",
    0x0016: "GPSDestLongitude",
    0x0017: "GPSDestBearingRef",
    0x0018: "GPSDestBearing",
    0x0019: "GPSDestDistanceRef",
    0x001A: "GPSDestDistance",
    0x001B: "GPSProcessingMethod",
    0x001C: "GPSAreaInformation",
    0x001D: "GPSDateStamp",
    0x001E: "GPSDifferential",
}

# Human-readable enum lookups, exiftool-style.
_ORIENTATION_NAMES = {
    1: "Horizontal (normal)", 2: "Mirror horizontal", 3: "Rotate 180",
    4: "Mirror vertical", 5: "Mirror horizontal and rotate 270 CW",
    6: "Rotate 90 CW", 7: "Mirror horizontal and rotate 90 CW", 8: "Rotate 270 CW",
}
_RESOLUTION_UNIT_NAMES = {1: "None", 2: "inches", 3: "cm"}
_EXPOSURE_PROGRAM_NAMES = {
    0: "Not Defined", 1: "Manual", 2: "Program AE", 3: "Aperture-priority AE",
    4: "Shutter speed priority AE", 5: "Creative (Slow speed)",
    6: "Action (High speed)", 7: "Portrait", 8: "Landscape",
}
_METERING_MODE_NAMES = {
    0: "Unknown", 1: "Average", 2: "Center-weighted average", 3: "Spot",
    4: "Multi-spot", 5: "Multi-segment", 6: "Partial", 255: "Other",
}
_LIGHT_SOURCE_NAMES = {
    0: "Unknown", 1: "Daylight", 2: "Fluorescent", 3: "Tungsten", 4: "Flash",
    9: "Fine Weather", 10: "Cloudy", 11: "Shade", 17: "Standard Light A",
    18: "Standard Light B", 19: "Standard Light C", 20: "D55", 21: "D65",
    22: "D75", 23: "D50", 24: "ISO Studio Tungsten", 255: "Other",
}
_WHITE_BALANCE_NAMES = {0: "Auto", 1: "Manual"}
_COLOR_SPACE_NAMES = {1: "sRGB", 2: "Adobe RGB", 0xFFFF: "Uncalibrated"}
_EXPOSURE_MODE_NAMES = {0: "Auto", 1: "Manual", 2: "Auto bracket"}
_SCENE_CAPTURE_NAMES = {0: "Standard", 1: "Landscape", 2: "Portrait", 3: "Night"}


def _flash_description(value):
    try:
        fired = bool(value & 0x1)
    except TypeError:
        return value
    return "Flash fired" if fired else "Flash did not fire"


# Tags that control how pixel data is located/decoded in a standalone TIFF.
# We must never touch these when cleaning, or the image will be corrupted.
TIFF_STRUCTURAL_TAGS = {
    0x0100, 0x0101, 0x0102, 0x0103, 0x0106, 0x0111, 0x0115, 0x0116, 0x0117,
    0x011C, 0x0140, 0x0142, 0x0143, 0x0144, 0x0145, 0x0153,
}

# Tags we consider "metadata" and therefore safe (and worth) blanking out
# inside IFD0 of a standalone TIFF (everything inside ExifIFD/GPSInfo/
# InteroperabilityIFD subtrees is always blanked regardless of this list).
TIFF_METADATA_TAGS = {
    0x010E, 0x010F, 0x0110, 0x0131, 0x0132, 0x013B, 0x013C, 0x8298,
    0x9C9B, 0x9C9C, 0x9C9D, 0x9C9E, 0x9C9F, 0x4746, 0x4749,
}


def _humanize_tag(name, value):
    """Apply exiftool-style formatting to a decoded tag value, where useful."""
    try:
        if name == "Orientation":
            return _ORIENTATION_NAMES.get(value, value)
        if name == "ResolutionUnit" or name == "FocalPlaneResolutionUnit":
            return _RESOLUTION_UNIT_NAMES.get(value, value)
        if name == "ExposureProgram":
            return _EXPOSURE_PROGRAM_NAMES.get(value, value)
        if name == "MeteringMode":
            return _METERING_MODE_NAMES.get(value, value)
        if name == "LightSource":
            return _LIGHT_SOURCE_NAMES.get(value, value)
        if name == "WhiteBalance":
            return _WHITE_BALANCE_NAMES.get(value, value)
        if name == "ColorSpace":
            return _COLOR_SPACE_NAMES.get(value, value)
        if name == "ExposureMode":
            return _EXPOSURE_MODE_NAMES.get(value, value)
        if name == "SceneCaptureType":
            return _SCENE_CAPTURE_NAMES.get(value, value)
        if name == "Flash":
            return _flash_description(value)
        if name == "ExposureTime" and isinstance(value, (int, float)) and value:
            if value < 1:
                denom = round(1 / value)
                return f"1/{denom} s"
            return f"{value:g} s"
        if name == "FNumber" and isinstance(value, (int, float)):
            return f"f/{value:g}"
        if name in ("FocalLength", "FocalLengthIn35mmFilm") and isinstance(value, (int, float)):
            return f"{value:g} mm"
        if name == "ExifVersion" or name == "FlashpixVersion" or name == "InteropVersion":
            if isinstance(value, (bytes, bytearray)):
                return value.decode("ascii", errors="replace")
        if name in ("MakerNote", "UserComment") and isinstance(value, (bytes, bytearray)):
            return f"(binary data, {len(value)} bytes)"
        if name == "LensSpecification" and isinstance(value, tuple) and len(value) == 4:
            lo, hi, apo, apc = value
            if lo == hi:
                return f"{lo:g} mm f/{apo:g}"
            return f"{lo:g}-{hi:g} mm f/{apo:g}-{apc:g}"
    except Exception:
        return value
    return value


def _convert_to_degrees(value):
    d, m, s = float(value[0]), float(value[1]), float(value[2])
    return d + (m / 60.0) + (s / 3600.0)


def _format_dms(value):
    """Format a (deg, min, sec) GPS tuple as exiftool-style '37 deg 25' 19.00"'."""
    try:
        d, m, s = float(value[0]), float(value[1]), float(value[2])
        return f"{d:g} deg {m:g}' {s:.2f}\""
    except Exception:
        return value


def _read_ifd(buf, offset, endian):
    """Parse one IFD. Returns (entries, next_ifd_offset)."""
    if offset + 2 > len(buf):
        return [], 0
    (num_entries,) = struct.unpack_from(endian + "H", buf, offset)
    entries = []
    pos = offset + 2
    for _ in range(num_entries):
        if pos + 12 > len(buf):
            break
        tag, typ, count = struct.unpack_from(endian + "HHI", buf, pos)
        type_size = TIFF_TYPE_SIZES.get(typ, 1)
        value_size = type_size * count
        if value_size <= 4:
            value_offset = pos + 8
            is_inline = True
        else:
            (rel_offset,) = struct.unpack_from(endian + "I", buf, pos + 8)
            value_offset = rel_offset
            is_inline = False
        entries.append({
            "tag": tag, "type": typ, "count": count,
            "value_offset": value_offset, "size": value_size,
            "is_inline": is_inline, "entry_pos": pos,
        })
        pos += 12
    if pos + 4 <= len(buf):
        (next_ifd,) = struct.unpack_from(endian + "I", buf, pos)
    else:
        next_ifd = 0
    return entries, next_ifd


def _decode_value(buf, entry, endian):
    typ, count, off, size = entry["type"], entry["count"], entry["value_offset"], entry["size"]
    if off + size > len(buf) or off < 0:
        return None
    try:
        if typ == 2:  # ASCII
            raw = buf[off: off + size]
            return raw.split(b"\x00", 1)[0].decode("utf-8", errors="replace")
        if typ == 3:  # SHORT
            vals = struct.unpack_from(endian + "H" * count, buf, off)
            return vals if count > 1 else vals[0]
        if typ == 4:  # LONG
            vals = struct.unpack_from(endian + "I" * count, buf, off)
            return vals if count > 1 else vals[0]
        if typ == 5:  # RATIONAL
            vals = []
            for i in range(count):
                num, den = struct.unpack_from(endian + "II", buf, off + i * 8)
                vals.append(num / den if den else 0.0)
            return tuple(vals) if count > 1 else vals[0]
        if typ == 10:  # SRATIONAL
            vals = []
            for i in range(count):
                num, den = struct.unpack_from(endian + "ii", buf, off + i * 8)
                vals.append(num / den if den else 0.0)
            return tuple(vals) if count > 1 else vals[0]
        if typ == 11:  # FLOAT
            vals = struct.unpack_from(endian + "f" * count, buf, off)
            return vals if count > 1 else vals[0]
        if typ == 12:  # DOUBLE
            vals = struct.unpack_from(endian + "d" * count, buf, off)
            return vals if count > 1 else vals[0]
        if typ == 1:  # BYTE
            if count == 1:
                return buf[off]
            return buf[off: off + size]
        if typ == 7:  # UNDEFINED
            return buf[off: off + size]
    except struct.error:
        return None
    return None


def parse_tiff_metadata(buf: bytes) -> dict:
    """Parse a raw TIFF/EXIF byte blob (buf[0:2] must be 'II' or 'MM')."""
    if len(buf) < 8 or buf[:2] not in (b"II", b"MM"):
        return {}
    endian = "<" if buf[:2] == b"II" else ">"
    (ifd0_offset,) = struct.unpack_from(endian + "I", buf, 4)

    metadata = {}
    gps_info = {}

    ifd0_entries, _next = _read_ifd(buf, ifd0_offset, endian)
    exif_ifd_offset = None
    gps_ifd_offset = None
    interop_ifd_offset = None

    for entry in ifd0_entries:
        name = TAG_NAMES.get(entry["tag"])
        if entry["tag"] == 0x8769:
            exif_ifd_offset = _decode_value(buf, entry, endian)
            continue
        if entry["tag"] == 0x8825:
            gps_ifd_offset = _decode_value(buf, entry, endian)
            continue
        if name:
            val = _decode_value(buf, entry, endian)
            if val is not None:
                metadata[name] = val

    if exif_ifd_offset:
        exif_entries, _ = _read_ifd(buf, exif_ifd_offset, endian)
        for entry in exif_entries:
            if entry["tag"] == 0xA005:
                interop_ifd_offset = _decode_value(buf, entry, endian)
                continue
            name = TAG_NAMES.get(entry["tag"])
            if name:
                val = _decode_value(buf, entry, endian)
                if val is not None:
                    metadata[name] = val

    if interop_ifd_offset:
        interop_entries, _ = _read_ifd(buf, interop_ifd_offset, endian)
        for entry in interop_entries:
            name = TAG_NAMES.get(entry["tag"])
            if name:
                val = _decode_value(buf, entry, endian)
                if val is not None:
                    metadata[name] = val

    if gps_ifd_offset:
        gps_entries, _ = _read_ifd(buf, gps_ifd_offset, endian)
        for entry in gps_entries:
            name = GPS_TAG_NAMES.get(entry["tag"], entry["tag"])
            val = _decode_value(buf, entry, endian)
            if val is not None:
                gps_info[name] = val

    if gps_info:
        try:
            lat, lat_ref = gps_info.get("GPSLatitude"), gps_info.get("GPSLatitudeRef")
            lon, lon_ref = gps_info.get("GPSLongitude"), gps_info.get("GPSLongitudeRef")
            if lat and lat_ref and lon and lon_ref:
                lat_deg = _convert_to_degrees(lat)
                if lat_ref != "N":
                    lat_deg = -lat_deg
                lon_deg = _convert_to_degrees(lon)
                if lon_ref != "E":
                    lon_deg = -lon_deg
                metadata["GoogleMapsURL"] = (
                    f"https://www.google.com/maps/search/?api=1&query={lat_deg},{lon_deg}"
                )
                metadata["GPSPosition"] = f"{lat_deg:.6f}, {lon_deg:.6f}"
        except Exception:
            pass

        for k, v in list(gps_info.items()):
            if k in ("GPSLatitude", "GPSLongitude", "GPSDestLatitude", "GPSDestLongitude") and isinstance(v, tuple):
                gps_info[k] = _format_dms(v)
        metadata["GPSInfo"] = gps_info

    # Apply exiftool-style human-readable formatting to the flat tags.
    for key, val in list(metadata.items()):
        if key in ("GPSInfo", "GoogleMapsURL", "GPSPosition"):
            continue
        metadata[key] = _humanize_tag(key, val)

    return metadata


def _blank_tiff_metadata(buf: bytearray, base_offset: int, endian: str) -> None:
    """
    In-place blank (zero out) the *values* of known metadata tags inside a
    TIFF/EXIF blob, without touching structural tags, entry layout, or the
    overall byte length. base_offset is where this TIFF header starts inside
    the parent buffer (0 for standalone TIFF files, or the start of the
    'II'/'MM' bytes for an embedded Exif blob).
    """
    (ifd0_offset,) = struct.unpack_from(endian + "I", buf, base_offset + 4)

    def resolve(rel):
        return base_offset + rel

    def blank_ifd(offset, allowed_tags, skip_structural):
        entries, _ = _read_ifd(bytes(buf), offset, endian)
        exif_off = gps_off = None
        for entry in entries:
            tag = entry["tag"]
            if tag == 0x8769:
                exif_off = _decode_value(bytes(buf), entry, endian)
                continue
            if tag == 0x8825:
                gps_off = _decode_value(bytes(buf), entry, endian)
                continue
            if skip_structural and tag in TIFF_STRUCTURAL_TAGS:
                continue
            if allowed_tags is not None and tag not in allowed_tags:
                continue
            if entry["is_inline"]:
                start = entry["entry_pos"] + 8
            else:
                start = resolve(entry["value_offset"])
            length = entry["size"]
            if 0 <= start and start + length <= len(buf):
                buf[start:start + length] = b"\x00" * length
        return exif_off, gps_off

    exif_off, gps_off = blank_ifd(offset=ifd0_offset, allowed_tags=TIFF_METADATA_TAGS, skip_structural=True)
    if exif_off:
        blank_ifd(offset=resolve(exif_off), allowed_tags=None, skip_structural=True)
    if gps_off:
        blank_ifd(offset=resolve(gps_off), allowed_tags=None, skip_structural=True)


# ---------------------------------------------------------------------------
# JPEG
# ---------------------------------------------------------------------------

# Marker segments that never contain payload after them (no length field).
_JPEG_NO_LENGTH_MARKERS = {0xD8, 0xD9, 0x01} | set(range(0xD0, 0xD8))
_JPEG_METADATA_MARKERS = {0xE1, 0xED, 0xFE}  # APP1 (Exif/XMP), APP13 (IPTC/Photoshop), COM


def _iter_jpeg_segments(data: bytes):
    """Yield (marker, seg_start, seg_end) for each marker segment. seg_end is
    exclusive and includes the 2-byte length field when present."""
    pos = 0
    length = len(data)
    while pos < length - 1:
        if data[pos] != 0xFF:
            pos += 1
            continue
        marker = data[pos + 1]
        if marker == 0xFF:
            pos += 1
            continue
        seg_start = pos
        pos += 2
        if marker in _JPEG_NO_LENGTH_MARKERS:
            yield marker, seg_start, pos
            if marker == 0xD9:  # EOI
                break
            continue
        if marker == 0xDA:  # Start of Scan: rest is entropy-coded data
            yield marker, seg_start, length
            break
        if pos + 2 > length:
            break
        seg_len = struct.unpack_from(">H", data, pos)[0]
        seg_end = pos + seg_len
        yield marker, seg_start, seg_end
        pos = seg_end


def get_jpeg_metadata(data: bytes) -> dict:
    metadata = {}
    for marker, seg_start, seg_end in _iter_jpeg_segments(data):
        if marker == 0xE1:
            payload = data[seg_start + 4: seg_end]
            if payload.startswith(b"Exif\x00\x00"):
                tiff_blob = payload[6:]
                metadata.update(parse_tiff_metadata(tiff_blob))
            elif payload.startswith(b"http://ns.adobe.com/xap/1.0/\x00"):
                metadata["XMP"] = "(XMP metadata present)"
        elif marker == 0xED:
            metadata["Photoshop/IPTC"] = "(IPTC/Photoshop metadata present)"
        elif marker == 0xFE:
            comment = data[seg_start + 4: seg_end]
            metadata["Comment"] = comment.decode("utf-8", errors="replace")
    return metadata


def clean_jpeg(data: bytes) -> bytes:
    out = bytearray()
    pos = 0
    for marker, seg_start, seg_end in _iter_jpeg_segments(data):
        out.extend(data[pos:seg_start])
        if marker not in _JPEG_METADATA_MARKERS:
            out.extend(data[seg_start:seg_end])
        pos = seg_end
    out.extend(data[pos:])
    return bytes(out)


# ---------------------------------------------------------------------------
# PNG
# ---------------------------------------------------------------------------

_PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
_PNG_METADATA_CHUNKS = {b"tEXt", b"zTXt", b"iTXt", b"tIME", b"eXIf"}


def _iter_png_chunks(data: bytes):
    pos = len(_PNG_SIGNATURE)
    length = len(data)
    while pos + 8 <= length:
        (chunk_len,) = struct.unpack_from(">I", data, pos)
        chunk_type = data[pos + 4: pos + 8]
        chunk_start = pos
        chunk_end = pos + 8 + chunk_len + 4  # length + type + data + crc
        yield chunk_type, chunk_start, chunk_end, data[pos + 8: pos + 8 + chunk_len]
        pos = chunk_end
        if chunk_type == b"IEND":
            break


def get_png_metadata(data: bytes) -> dict:
    metadata = {}
    for chunk_type, _s, _e, payload in _iter_png_chunks(data):
        if chunk_type == b"tEXt":
            key, _, value = payload.partition(b"\x00")
            metadata[key.decode("latin-1", errors="replace")] = value.decode("latin-1", errors="replace")
        elif chunk_type == b"iTXt":
            key, _, rest = payload.partition(b"\x00")
            metadata[key.decode("utf-8", errors="replace")] = "(iTXt text present)"
        elif chunk_type == b"zTXt":
            key, _, _ = payload.partition(b"\x00")
            metadata[key.decode("latin-1", errors="replace")] = "(compressed text present)"
        elif chunk_type == b"tIME":
            if len(payload) == 7:
                year, month, day, hour, minute, second = struct.unpack(">HBBBBB", payload)
                metadata["ModifyTime"] = f"{year:04d}-{month:02d}-{day:02d} {hour:02d}:{minute:02d}:{second:02d}"
        elif chunk_type == b"eXIf":
            metadata.update(parse_tiff_metadata(payload))
        elif chunk_type == b"pHYs" and len(payload) == 9:
            ppux, ppuy, unit = struct.unpack(">IIB", payload)
            if unit == 1:  # meters -> dpi
                metadata["XResolution(dpi)"] = round(ppux * 0.0254, 2)
                metadata["YResolution(dpi)"] = round(ppuy * 0.0254, 2)
            else:
                metadata["PixelsPerUnitX"] = ppux
                metadata["PixelsPerUnitY"] = ppuy
        elif chunk_type == b"gAMA" and len(payload) == 4:
            (gamma,) = struct.unpack(">I", payload)
            metadata["Gamma"] = round(100000 / gamma, 5) if gamma else None
        elif chunk_type == b"sRGB" and len(payload) == 1:
            rendering = {0: "Perceptual", 1: "Relative colorimetric", 2: "Saturation", 3: "Absolute colorimetric"}
            metadata["sRGBRenderingIntent"] = rendering.get(payload[0], payload[0])
    return metadata


def clean_png(data: bytes) -> bytes:
    out = bytearray(_PNG_SIGNATURE)
    for chunk_type, chunk_start, chunk_end, _payload in _iter_png_chunks(data):
        if chunk_type not in _PNG_METADATA_CHUNKS:
            out.extend(data[chunk_start:chunk_end])
    return bytes(out)


# ---------------------------------------------------------------------------
# TIFF (standalone) — must preserve absolute byte offsets, so we blank values
# in place instead of removing bytes.
# ---------------------------------------------------------------------------

def get_tiff_metadata(data: bytes) -> dict:
    return parse_tiff_metadata(data)


def clean_tiff(data: bytes) -> bytes:
    if len(data) < 8 or data[:2] not in (b"II", b"MM"):
        return data
    endian = "<" if data[:2] == b"II" else ">"
    buf = bytearray(data)
    _blank_tiff_metadata(buf, 0, endian)
    return bytes(buf)


# ---------------------------------------------------------------------------
# WEBP (RIFF container)
# ---------------------------------------------------------------------------

def _iter_riff_chunks(data: bytes):
    pos = 12  # skip 'RIFF', size, 'WEBP'
    length = len(data)
    while pos + 8 <= length:
        fourcc = data[pos:pos + 4]
        (size,) = struct.unpack_from("<I", data, pos + 4)
        payload_start = pos + 8
        payload_end = payload_start + size
        chunk_end = payload_end + (size % 2)  # chunks are padded to even size
        yield fourcc, pos, chunk_end, data[payload_start:payload_end]
        pos = chunk_end


def get_webp_metadata(data: bytes) -> dict:
    metadata = {}
    if data[:4] != b"RIFF" or data[8:12] != b"WEBP":
        return metadata
    for fourcc, _s, _e, payload in _iter_riff_chunks(data):
        if fourcc == b"EXIF":
            metadata.update(parse_tiff_metadata(payload))
        elif fourcc == b"XMP ":
            metadata["XMP"] = "(XMP metadata present)"
    return metadata


def clean_webp(data: bytes) -> bytes:
    if data[:4] != b"RIFF" or data[8:12] != b"WEBP":
        return data
    out = bytearray(data[:12])
    for fourcc, _s, _e, payload in _iter_riff_chunks(data):
        if fourcc in (b"EXIF", b"XMP "):
            continue
        chunk = fourcc + struct.pack("<I", len(payload)) + payload
        if len(payload) % 2:
            chunk += b"\x00"
        out.extend(chunk)
    new_size = len(out) - 8
    struct.pack_into("<I", out, 4, new_size)
    return bytes(out)


# ---------------------------------------------------------------------------
# Image dispatch
# ---------------------------------------------------------------------------

def get_image_metadata(path: Path) -> dict:
    data = path.read_bytes()
    ext = path.suffix.lower()
    if ext in (".jpg", ".jpeg"):
        return get_jpeg_metadata(data)
    if ext == ".png":
        return get_png_metadata(data)
    if ext in (".tif", ".tiff"):
        return get_tiff_metadata(data)
    if ext == ".webp":
        return get_webp_metadata(data)
    return {}


def clean_image(path: Path, output_path: Path) -> None:
    data = path.read_bytes()
    ext = path.suffix.lower()
    if ext in (".jpg", ".jpeg"):
        cleaned = clean_jpeg(data)
    elif ext == ".png":
        cleaned = clean_png(data)
    elif ext in (".tif", ".tiff"):
        cleaned = clean_tiff(data)
    elif ext == ".webp":
        cleaned = clean_webp(data)
    else:
        cleaned = data
    output_path.write_bytes(cleaned)


# ---------------------------------------------------------------------------
# PDF — no absolute-offset issues need to be avoided for anything except
# preserving overall file length (xref offsets point at byte positions), so
# we blank metadata values in place rather than deleting bytes.
# ---------------------------------------------------------------------------

_PDF_STRING_RE = re.compile(rb"\((?:\\.|[^\\()])*\)")
_PDF_HEXSTR_RE = re.compile(rb"<[0-9A-Fa-f\s]*>")


def _find_info_object(data: bytes):
    """Locate the /Info N G R reference and return (num, gen) or None."""
    matches = list(re.finditer(rb"/Info\s+(\d+)\s+(\d+)\s+R", data))
    if not matches:
        return None
    num, gen = matches[-1].group(1), matches[-1].group(2)
    return int(num), int(gen)


def _find_object_span(data: bytes, num: int, gen: int):
    pattern = re.compile(
        rb"(?<!\d)" + str(num).encode() + rb"\s+" + str(gen).encode() + rb"\s+obj(.*?)endobj",
        re.DOTALL,
    )
    m = pattern.search(data)
    if not m:
        return None
    return m.start(1), m.end(1)


def _parse_pdf_dict_entries(dict_bytes: bytes):
    """Yield (key, value_start, value_end) for simple /Key (str) or /Key <hex> pairs."""
    for m in re.finditer(rb"/([A-Za-z0-9_]+)\s*(\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]*>)", dict_bytes):
        key = m.group(1).decode("latin-1")
        yield key, m.start(2), m.end(2)


def get_pdf_metadata(path: Path) -> dict:
    data = path.read_bytes()
    metadata = {}
    info_ref = _find_info_object(data)
    if info_ref:
        span = _find_object_span(data, *info_ref)
        if span:
            dict_bytes = data[span[0]:span[1]]
            for key, vs, ve in _parse_pdf_dict_entries(dict_bytes):
                raw = dict_bytes[vs:ve]
                if raw.startswith(b"("):
                    text = raw[1:-1].decode("latin-1", errors="replace")
                else:
                    text = raw.decode("latin-1", errors="replace")
                metadata[key] = text
    if re.search(rb"<\?xpacket begin", data):
        metadata["XMP"] = "(XMP metadata present)"
    return metadata


def clean_pdf(path: Path, output_path: Path) -> None:
    data = bytearray(path.read_bytes())

    # 1) Blank the values inside the document Info dictionary, keeping the
    #    dictionary keys/structure and overall file length intact.
    info_ref = _find_info_object(bytes(data))
    if info_ref:
        span = _find_object_span(bytes(data), *info_ref)
        if span:
            start, end = span
            dict_bytes = bytes(data[start:end])
            for _key, vs, ve in _parse_pdf_dict_entries(dict_bytes):
                raw = dict_bytes[vs:ve]
                if raw.startswith(b"("):
                    inner_len = len(raw) - 2
                    blanked = b"(" + b" " * inner_len + b")"
                else:
                    inner_len = len(raw) - 2
                    blanked = b"<" + b"0" * inner_len + b">"
                data[start + vs: start + ve] = blanked

    # 2) Blank any XMP metadata packets in place (keeps /Length correct since
    #    total size is unchanged).
    for m in re.finditer(rb"<\?xpacket begin.*?<\?xpacket end=.*?\?>", bytes(data), re.DOTALL):
        s, e = m.span()
        original = bytes(data[s:e])
        blanked = re.sub(rb"[^\r\n]", b" ", original)
        data[s:e] = blanked

    output_path.write_bytes(bytes(data))


# ---------------------------------------------------------------------------
# Office (docx/xlsx/pptx) — zip-based, already pure standard library
# ---------------------------------------------------------------------------

def get_office_metadata(path: Path) -> dict:
    found = {}
    with zipfile.ZipFile(path, "r") as z:
        for part in OFFICE_METADATA_PARTS:
            if part in z.namelist():
                found[part] = z.read(part).decode("utf-8", errors="ignore")
    return found


def clean_office(path: Path, output_path: Path) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in names:
                data = zin.read(item)
                if item.endswith("docProps/core.xml"):
                    data = _empty_core_xml()
                elif item.endswith("docProps/app.xml"):
                    data = _empty_app_xml()
                elif item.endswith("docProps/custom.xml"):
                    data = _empty_custom_xml()
                zout.writestr(item, data)


def _empty_core_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/'
        b'package/2006/metadata/core-properties" '
        b'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        b'xmlns:dcterms="http://purl.org/dc/terms/" '
        b'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        b"</cp:coreProperties>"
    )


def _empty_app_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Properties xmlns="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/extended-properties" '
        b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/docPropsVTypes"></Properties>'
    )


def _empty_custom_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Properties xmlns="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/custom-properties" '
        b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/docPropsVTypes"></Properties>'
    )


# ---------------------------------------------------------------------------
# Generic dispatch / CLI (unchanged from original)
# ---------------------------------------------------------------------------

def get_metadata(path: Path) -> dict:
    ext = path.suffix.lower()
    if ext in IMAGE_EXTS:
        return get_image_metadata(path)
    if ext in PDF_EXTS:
        return get_pdf_metadata(path)
    if ext in OFFICE_EXTS:
        return get_office_metadata(path)
    return {}


def clean_file(path: Path, output_path: Path) -> bool:
    ext = path.suffix.lower()
    if ext in IMAGE_EXTS:
        clean_image(path, output_path)
    elif ext in PDF_EXTS:
        clean_pdf(path, output_path)
    elif ext in OFFICE_EXTS:
        clean_office(path, output_path)
    else:
        return False
    return True


def default_output_path(path: Path) -> Path:
    return path.with_name(f"{path.stem}_clean{path.suffix}")


def print_report(path: Path, meta: dict) -> None:
    print(f"\n📄 {path}")
    if not meta:
        print("   No notable metadata found, or file format is unsupported.")
        return
    for key, value in meta.items():
        if isinstance(value, dict):
            print(f"   - {key}:")
            for sub_key, sub_val in value.items():
                text = str(sub_val)
                if len(text) > 120:
                    text = text[:120] + "..."
                print(f"       • {sub_key}: {text}")
            continue
        text = str(value)
        if len(text) > 120:
            text = text[:120] + "..."
        print(f"   - {key}: {text}")


def is_supported(path: Path) -> bool:
    return path.suffix.lower() in (IMAGE_EXTS | PDF_EXTS | OFFICE_EXTS)


def collect_files(input_path: Path, recursive: bool) -> list:
    if input_path.is_file():
        return [input_path]
    pattern = "**/*" if recursive else "*"
    return [p for p in input_path.glob(pattern) if p.is_file() and is_supported(p)]


def main():
    parser = argparse.ArgumentParser(
        description="MetadataStripper - A tool to strip metadata and PII from files."
    )
    parser.add_argument("input", nargs="?", help="Path to the file or directory to process")
    parser.add_argument("-o", "--output", help="Path to save the output file (single file only)")
    parser.add_argument(
        "--report-only",
        action="store_true",
        help="Display metadata without deleting it",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Process files in subdirectories recursively if input is a directory",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
        help="Show tool version",
    )
    parser.add_argument(
        "--config",
        help="Path to configuration file (optional)",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Set logging level",
    )

    args = parser.parse_args()

    if not args.input:
        parser.print_help()
        sys.exit(1)

    input_path = Path(args.input)

    try:
        if not input_path.exists():
            print(f"❌ Error: Path does not exist: {input_path}")
            sys.exit(1)
    except PermissionError:
        print(f"❌ Error: Insufficient permissions to access path: {input_path}")
        sys.exit(1)

    files = collect_files(input_path, args.recursive)
    if not files:
        print("⚠️ Warning: No supported files found.")
        sys.exit(0)

    for f in files:
        try:
            meta = get_metadata(f)
            print_report(f, meta)

            if args.report_only:
                continue

            if args.output and len(files) == 1:
                out = Path(args.output)
            else:
                out = default_output_path(f)

            ok = clean_file(f, out)
            if ok:
                print(f"   ✅ Successfully saved without metadata → {out}")

        except PermissionError:
            print(f"   ❌ Error: Insufficient permissions for file {f}")
        except zipfile.BadZipFile:
            print(f"   ❌ Error: Corrupted archive or file: {f}")
        except Exception as e:
            if args.log_level == "DEBUG":
                import traceback
                traceback.print_exc()
            else:
                print(f"   ❌ An error occurred while processing {f}: {e}")


if __name__ == "__main__":
    main()
