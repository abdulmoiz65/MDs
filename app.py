import argparse
import os
import posixpath
import re
import sys
import warnings
import zipfile
import xml.etree.ElementTree as ET

# pydub (pulled in by markitdown[audio-transcription]) warns about ffmpeg on
# import even though we use faster-whisper which needs no ffmpeg at all.
warnings.filterwarnings("ignore", category=RuntimeWarning, module="pydub")

from markitdown import MarkItDown

# Maps MIME sub-types / file extensions to output file extensions
_EXT_NORM = {
    "png": ".png",
    "jpg": ".jpg",
    "jpeg": ".jpg",
    "gif": ".gif",
    "bmp": ".bmp",
    "webp": ".webp",
    "tiff": ".tiff",
    "svg": ".svg",
    "emf": ".emf",
    "wmf": ".wmf",
}

# Matches the stub "base64..." placeholders MarkItDown writes for docx/xlsx
# when it can't embed actual image bytes.
#   group 1 = alt text
_IMG_PLACEHOLDER_RE = re.compile(
    r'!\[([^\]]*)\]\(data:image/[^;]+;base64[^)]*\)'
)

# Matches the plain-filename image refs MarkItDown writes for pptx, e.g.:
#   ![](Picture5.jpg)  or  ![some alt](image1.png)
# Only matches bare filenames (no directory separator) with image extensions.
#   group 1 = alt text
#   group 2 = filename
_PPTX_IMG_RE = re.compile(
    r'!\[([^\]]*)\]\(([^/\)]+\.(?:png|jpe?g|gif|bmp|webp|tiff?|svg|emf|wmf))\)',
    re.IGNORECASE,
)

# XML namespaces used in Office Open XML
_NS = {
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "v":   "urn:schemas-microsoft-com:vml",
    "o":   "urn:schemas-microsoft-com:office:office",
}

# ooxml relationship type for images
_IMG_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

# EMF/WMF are Windows-only vector formats invisible in browsers.
# We try to convert them to PNG automatically.
_VECTOR_FORMATS = {".emf", ".wmf"}


def _convert_to_png(img_path: str) -> str:
    """
    Try to convert a vector image (EMF/WMF) to PNG using Pillow (Windows GDI).
    Returns the path to the PNG if successful, or the original path on failure.
    """
    if not img_path.lower().endswith(tuple(_VECTOR_FORMATS)):
        return img_path
    png_path = os.path.splitext(img_path)[0] + ".png"
    # Strategy 1: Pillow (works on Windows via GDI for EMF/WMF)
    try:
        from PIL import Image
        img = Image.open(img_path)
        img.save(png_path, "PNG")
        os.remove(img_path)
        return png_path
    except Exception:
        pass
    # Strategy 2: ImageMagick CLI
    try:
        import subprocess
        result = subprocess.run(
            ["magick", "convert", img_path, png_path],
            capture_output=True, timeout=15
        )
        if result.returncode == 0 and os.path.exists(png_path):
            os.remove(img_path)
            return png_path
    except Exception:
        pass
    # Could not convert — keep original
    return img_path


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

SUPPORTED_EXTENSIONS = {
    "pptx", "ppt",          # PowerPoint
    "docx", "doc",          # Word
    "xlsx",                 # Excel (modern)
    "xls",                  # Excel (legacy)
    "wav", "mp3",           # Audio — handled by faster-whisper (no ffmpeg needed)
    "csv", "json", "xml",   # Data formats
    "epub", "zip",          # Other
}

AUDIO_EXTENSIONS = {"wav", "mp3"}


def get_file_extension(filename: str) -> str:
    """Extract the file extension from a filename."""
    return filename.rsplit(".", 1)[1].lower() if "." in filename else ""


def is_youtube_url(url):
    """Check if a URL is a valid YouTube URL."""
    youtube_regex = r"^(https?://)?(www\.)?(youtube\.com|youtu\.?be)/.+$"
    return bool(re.match(youtube_regex, url))


def extract_youtube_id(url):
    """Extract the YouTube video ID from a URL."""
    patterns = [
        r"(?:v=|\/)([0-9A-Za-z_-]{11}).*",  # Standard YouTube URLs
        r"(?:youtu\.be\/)([0-9A-Za-z_-]{11})",  # Short YouTube URLs
        r"(?:embed\/)([0-9A-Za-z_-]{11})",  # Embedded YouTube URLs
    ]

    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)

    return None


def _build_audio_markdown(name: str, language: str, duration: float, segments) -> str:
    """Shared formatter for both Whisper backends."""
    lines = [
        f"# Transcription: {name}\n",
        f"**Language:** {language}  ",
        f"**Duration:** {duration:.1f}s\n",
        "## Content\n",
    ]
    for seg in segments:
        if isinstance(seg, dict):
            ts = f"[{seg['start']:.1f}s → {seg['end']:.1f}s]"
            text = seg["text"].strip()
        else:
            ts = f"[{seg.start:.1f}s → {seg.end:.1f}s]"
            text = seg.text.strip()
        lines.append(f"{ts} {text}\n")
    return "\n".join(lines)


def convert_audio(input_path: str):
    """Transcribe audio using faster-whisper — no ffmpeg required."""
    try:
        from faster_whisper import WhisperModel
        from pathlib import Path

        print("[INFO] Transcribing with faster-whisper...")
        model = WhisperModel("base", device="cpu", compute_type="int8")
        segments, info = model.transcribe(input_path)
        md = _build_audio_markdown(
            Path(input_path).name, info.language, info.duration, list(segments)
        )
        return md, None
    except ImportError:
        return "", "faster-whisper not installed. Run: pip install faster-whisper"
    except Exception as exc:
        return "", str(exc)


def _rels_to_image_map(zf: zipfile.ZipFile, rels_path: str) -> dict:
    """
    Parse an Office Open XML .rels file and return a dict mapping
    rId → zip-internal path for every image relationship.
    """
    try:
        data = zf.read(rels_path).decode("utf-8")
    except KeyError:
        return {}

    root = ET.fromstring(data)
    mapping = {}
    for rel in root:
        if rel.get("Type") == _IMG_REL_TYPE:
            rid = rel.get("Id")
            target = rel.get("Target", "")
            # Target is relative to the folder containing the .rels file
            # e.g.  rels_path = "ppt/slides/_rels/slide1.xml.rels"
            # so base = "ppt/slides/"
            base = rels_path.replace("_rels/", "").rsplit("/", 1)[0] + "/"
            if not target.startswith("/"):
                # posixpath.normpath resolves ".." segments, e.g.:
                # "ppt/slides/" + "../media/image1.png" → "ppt/media/image1.png"
                full = posixpath.normpath(base + target)
            else:
                full = target.lstrip("/")
            mapping[rid] = full
    return mapping


def _ordered_rids_from_xml(zf: zipfile.ZipFile, xml_path: str) -> list:
    """
    Parse an Office Open XML document part and return the list of r:embed / r:id
    attribute values *in document order* (images only, preserving duplicates).
    """
    try:
        data = zf.read(xml_path).decode("utf-8")
    except KeyError:
        return []
    root = ET.fromstring(data)
    rids = []
    r_ns = _NS["r"]
    for elem in root.iter():
        for attr in ("embed", "id", "href"):
            val = elem.get(f"{{{r_ns}}}{attr}")
            if val:
                rids.append(val)
    return rids


def _extract_images_from_zip(
    input_path: str,
    images_dir: str,
) -> tuple[list, list]:
    """
    Extract every image from an Office Open XML (zip-based) file and save them
    to *images_dir*.

    Returns:
        (body_imgs, orphan_imgs)
        body_imgs   — images from the main document body, in document order.
                      These are matched 1-to-1 with the markdown placeholders.
        orphan_imgs — images from headers, footers, embedded objects etc.
                      MarkItDown generates no placeholder for these; they are
                      appended to the markdown under "Additional Images".
    """
    # Determine which XML part and rels file to use
    ext = input_path.rsplit(".", 1)[-1].lower()
    if ext in {"docx", "doc"}:
        xml_part = "word/document.xml"
        rels_part = "word/_rels/document.xml.rels"
        media_prefix = "word/media/"
    elif ext in {"pptx", "ppt"}:
        # For pptx we scan all slides in order
        xml_part = None
        rels_part = None
        media_prefix = "ppt/media/"
    elif ext in {"xlsx", "xls"}:
        xml_part = "xl/workbook.xml"
        rels_part = "xl/_rels/workbook.xml.rels"
        media_prefix = "xl/media/"
    else:
        media_prefix = "media/"
        xml_part = None
        rels_part = None

    # body_imgs  = images referenced in main document body (matched to placeholders)
    # orphan_imgs = images in headers/footers/other parts (appended at end of MD)
    body_imgs: list = []
    orphan_imgs: list = []

    with zipfile.ZipFile(input_path, "r") as zf:
        all_names = zf.namelist()

        # --- Non-pptx: use rels + document xml to get ordered image list ---
        if xml_part and rels_part:
            rid_to_path = _rels_to_image_map(zf, rels_part)
            ordered_rids = _ordered_rids_from_xml(zf, xml_part)
            # Build ordered list of zip-internal image paths from the body
            seen = set()
            ordered_body = []
            for rid in ordered_rids:
                p = rid_to_path.get(rid)
                if p and p in all_names and p not in seen:
                    ordered_body.append(p)
                    seen.add(p)

            # Collect header/footer rels separately (orphan images)
            orphan_zip_paths = []
            other_rels = [
                n for n in all_names
                if n.endswith(".rels") and n != rels_part and n != "_rels/.rels"
                and "customXml" not in n
            ]
            for rel_f in sorted(other_rels):
                for rid, p in _rels_to_image_map(zf, rel_f).items():
                    if p in all_names and p not in seen:
                        orphan_zip_paths.append(p)
                        seen.add(p)

            # Safety fallback: any remaining media not yet seen
            for name in sorted(all_names):
                if name.startswith(media_prefix) and name not in seen:
                    orphan_zip_paths.append(name)
                    seen.add(name)

        else:
            # pptx: collect slide images in slide number order
            slide_entries = sorted(
                [n for n in all_names if re.match(r"ppt/slides/slide\d+\.xml", n)],
                key=lambda x: int(re.search(r"\d+", x).group()),
            )
            seen = set()
            ordered_body = []
            orphan_zip_paths = []
            for slide_xml in slide_entries:
                slide_num = re.search(r"\d+", slide_xml).group()
                rels_f = f"ppt/slides/_rels/slide{slide_num}.xml.rels"
                rid_map = _rels_to_image_map(zf, rels_f)
                for rid in _ordered_rids_from_xml(zf, slide_xml):
                    p = rid_map.get(rid)
                    if p and p in all_names and p not in seen:
                        ordered_body.append(p)
                        seen.add(p)
            # fallback
            for name in sorted(all_names):
                if name.startswith(media_prefix) and name not in seen:
                    orphan_zip_paths.append(name)
                    seen.add(name)

        os.makedirs(images_dir, exist_ok=True)

        def _save_zip_image(zip_path: str, idx: int) -> str | None:
            raw_ext = zip_path.rsplit(".", 1)[-1].lower() if "." in zip_path else "bin"
            norm_ext = _EXT_NORM.get(raw_ext, f".{raw_ext}")
            out_name = f"image_{idx:03d}{norm_ext}"
            out_path = os.path.join(images_dir, out_name)
            try:
                data = zf.read(zip_path)
                with open(out_path, "wb") as fh:
                    fh.write(data)
                return out_path
            except Exception as exc:
                print(f"[WARN] Could not extract {zip_path}: {exc}", file=sys.stderr)
                return None

        idx = 1
        for zip_path in ordered_body:
            p = _save_zip_image(zip_path, idx)
            if p:
                body_imgs.append(p)
            idx += 1

        for zip_path in orphan_zip_paths:
            p = _save_zip_image(zip_path, idx)
            if p:
                orphan_imgs.append(p)
            idx += 1

    return body_imgs, orphan_imgs


def extract_and_save_images(
    markdown_content: str,
    output_path: str,
    input_path: str | None = None,
) -> str:
    """
    Replace every ``![alt](data:image/...;base64...)`` stub placeholder that
    MarkItDown writes for Office zip formats (docx/pptx/xlsx) with a link to
    a real saved image file extracted directly from the zip archive.
    """
    stem = os.path.splitext(output_path)[0]
    images_dir = stem + "_images"
    md_dir = os.path.dirname(os.path.abspath(output_path))

    # --- Strategy A: zip-based Office document ---
    zip_exts = {"docx", "doc", "pptx", "ppt", "xlsx", "xls"}
    use_zip = (
        input_path is not None
        and "." in input_path
        and input_path.rsplit(".", 1)[-1].lower() in zip_exts
    )

    ext_lower = input_path.rsplit(".", 1)[-1].lower() if use_zip else ""

    # pptx: MarkItDown writes bare filenames  ![](Picture5.jpg)  — use dedicated regex.
    # xlsx: MarkItDown produces NO placeholders at all (pandas reads the sheet).
    # Both need special handling so we don't bail out early.
    is_pptx  = use_zip and ext_lower in {"pptx", "ppt"}
    is_xlsx  = use_zip and ext_lower in {"xlsx", "xls"}

    # Pick the right placeholder regex for this format
    if is_pptx:
        placeholders = list(_PPTX_IMG_RE.finditer(markdown_content))
    else:
        placeholders = list(_IMG_PLACEHOLDER_RE.finditer(markdown_content))

    if not placeholders and not is_xlsx:
        return markdown_content

    if use_zip:
        body_imgs, orphan_imgs = _extract_images_from_zip(input_path, images_dir)
        updated = markdown_content

        # Replace in-body placeholders with real image paths (positional match)
        for m, img_path in zip(placeholders, body_imgs):
            img_path = _convert_to_png(img_path)          # EMF/WMF → PNG
            rel = os.path.relpath(img_path, md_dir).replace("\\", "/")
            rel_encoded = rel.replace(" ", "%20")
            alt = m.group(1)
            updated = updated.replace(m.group(0), f"![{alt}]({rel_encoded})", 1)
            print(f"  [IMG] Saved {rel}")

        # Append header/footer/orphan images at the bottom of the markdown
        if orphan_imgs:
            if is_xlsx:
                header = "\n\n---\n\n## Sheet Images\n\n"
            else:
                header = (
                    "\n\n---\n\n## Additional Images\n\n"
                    "> *These images appear in the document's headers, "
                    "footers, or embedded objects.*\n"
                )
            lines = [header]
            for img_path in orphan_imgs:
                img_path = _convert_to_png(img_path)      # EMF/WMF → PNG
                rel = os.path.relpath(img_path, md_dir).replace("\\", "/")
                rel_encoded = rel.replace(" ", "%20")
                lines.append(f"\n![{os.path.basename(img_path)}]({rel_encoded})\n")
                print(f"  [IMG] Appended (header/footer) {rel}")
            updated += "".join(lines)

        return updated

    # Non-zip format with no placeholders — nothing to do
    return markdown_content


def _clean_xlsx_markdown(content: str) -> str:
    """
    Post-process MarkItDown's xlsx output:
    - Replace NaN cells (empty Excel cells rendered by pandas) with blank.
    - Replace 'Unnamed: N' column headers (pandas auto-names for merged/blank
      header columns) with blank, so the table looks clean.
    """
    content = re.sub(r'(?<=\|)\s*NaN\s*(?=\|)', ' ', content)
    content = re.sub(r'(?<=\|)\s*Unnamed:\s*\d+\s*(?=\|)', ' ', content)
    return content


def convert_path_to_markdown(input_path: str):
    """
    Convert a local file OR a URL (e.g. YouTube) to Markdown.

    Returns:
        Tuple of (markdown_content: str, error: str | None)
    """
    ext = input_path.rsplit(".", 1)[-1].lower() if "." in input_path else ""

    if ext in AUDIO_EXTENSIONS:
        return convert_audio(input_path)

    try:
        md = MarkItDown(enable_plugins=False)
        result = md.convert(input_path)
        content = result.text_content
        # Clean up pandas artefacts in xlsx/xls output
        if ext in {"xlsx", "xls"}:
            content = _clean_xlsx_markdown(content)
        return content, None
    except Exception as exc:
        return "", str(exc)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="office_to_markdown",
        description="Convert Office documents, audio, and YouTube URLs to Markdown.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Supported file types:\n"
            "  Documents   : docx, doc, epub\n"
            "  Spreadsheets: xlsx, xls\n"
            "  Presentations: pptx, ppt\n"
            "  Audio       : wav, mp3  (faster-whisper, no ffmpeg needed)\n"
            "  Data        : csv, json, xml\n"
            "  Other       : epub, zip\n"
            "  YouTube URL : https://www.youtube.com/watch?v=...\n"
            "\nNOTE: For Outlook .msg files use outlook_to_md.py instead.\n"
            "\nExamples:\n"
            "  python app.py report.docx\n"
            "  python app.py slides.pptx --output slides.md\n"
            "  python app.py data.xlsx --output data.md\n"
            "  python app.py recording.mp3\n"
            "  python app.py https://www.youtube.com/watch?v=dQw4w9WgXcQ\n"
        ),
    )
    parser.add_argument(
        "input",
        help="Path to the input file or a YouTube URL.",
    )
    parser.add_argument(
        "-o", "--output",
        metavar="OUTPUT",
        help=(
            "Path for the output Markdown file. "
            "Defaults to <input_stem>.md in the same directory "
            "(or youtube_<id>.md for YouTube URLs)."
        ),
    )
    parser.add_argument(
        "--preview",
        metavar="N",
        type=int,
        default=0,
        help="Print the first N characters of the result to stdout (0 = disabled).",
    )
    return parser


def derive_output_path(input_arg: str) -> str:
    """Build a default output filename from the input path or YouTube URL."""
    if is_youtube_url(input_arg):
        video_id = extract_youtube_id(input_arg)
        stem = f"youtube_{video_id}" if video_id else "youtube_video"
        return f"{stem}.md"
    base = os.path.splitext(os.path.abspath(input_arg))[0]
    return f"{base}.md"


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_arg: str = args.input

    # --- Validate input ---
    is_url = is_youtube_url(input_arg)
    if not is_url:
        if not os.path.isfile(input_arg):
            parser.error(f"File not found: {input_arg}")
        ext = get_file_extension(input_arg)
        if ext not in SUPPORTED_EXTENSIONS:
            parser.error(
                f"Unsupported file extension '.{ext}'.\n"
                f"Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
            )

    # --- Convert ---
    print(f"Converting: {input_arg}")
    markdown_content, error = convert_path_to_markdown(input_arg)

    if error:
        print(f"[ERROR] {error}", file=sys.stderr)
        sys.exit(1)

    # --- Write output ---
    output_path: str = args.output or derive_output_path(input_arg)

    # Extract base64 images → real files, update markdown references
    if not is_url:
        ext = get_file_extension(input_arg)
        if ext not in AUDIO_EXTENSIONS:
            markdown_content = extract_and_save_images(
                markdown_content, output_path, input_path=input_arg
            )

    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write(markdown_content)
    print(f"Saved: {output_path}")

    # --- Optional preview ---
    if args.preview > 0:
        preview = markdown_content[: args.preview]
        if len(markdown_content) > args.preview:
            preview += f"\n\n... (truncated at {args.preview} chars)"
        print("\n--- Preview ---")
        print(preview)


if __name__ == "__main__":
    main()