__version__ = "1.0.0"

import argparse
import os
import sys
import zipfile
from pathlib import Path

from PIL import Image
from pypdf import PdfReader, PdfWriter

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".webp"}
PDF_EXTS = {".pdf"}
OFFICE_EXTS = {".docx", ".xlsx", ".pptx"}

OFFICE_METADATA_PARTS = {
    "docProps/core.xml",
    "docProps/app.xml",
    "docProps/custom.xml",
}

def get_image_metadata(path: Path) -> dict:
    img = Image.open(path)
    exif = img.getexif()
    if not exif:
        return {}
    from PIL.ExifTags import TAGS
    return {TAGS.get(tag_id, tag_id): value for tag_id, value in exif.items()}

def clean_image(path: Path, output_path: Path) -> None:
    img = Image.open(path)
    data = list(img.getdata())
    clean_img = Image.new(img.mode, img.size)
    clean_img.putdata(data)
    clean_img.save(output_path)

def get_pdf_metadata(path: Path) -> dict:
    reader = PdfReader(str(path))
    meta = reader.metadata
    return dict(meta) if meta else {}

def clean_pdf(path: Path, output_path: Path) -> None:
    reader = PdfReader(str(path))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.add_metadata({"/Producer": "", "/Creator": ""})
    try:
        writer.xmp_metadata = None
    except Exception:
        pass
    with open(output_path, "wb") as f:
        writer.write(f)

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
        text = str(value)
        if len(text) > 80:
            text = text[:80] + "..."
        print(f"   - {key}: {text}")

def is_supported(path: Path) -> bool:
    return path.suffix.lower() in (IMAGE_EXTS | PDF_EXTS | OFFICE_EXTS)

def collect_files(input_path: Path, recursive: bool) -> list[Path]:
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