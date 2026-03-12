#!/usr/bin/env python3
"""PDF to Markdown Converter.

Converts PDF documents to clean, well-structured Markdown with support for:
text formatting, headings, lists, hyperlinks, images, tables, annotations,
bookmarks, metadata, embedded files, and encrypted PDFs.

Dependencies:
    pymupdf (fitz), pymupdf4llm, Pillow, PyYAML
"""

# ── Standard library ────────────────────────────────────────────────
import argparse
import re
import sys
from pathlib import Path

# ── Third-party (optional with fallbacks) ───────────────────────────
try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False

try:
    import fitz  # PyMuPDF
except ImportError:
    print("❌ PyMuPDF (fitz) is required. Install with:")
    print("   pip install pymupdf")
    sys.exit(1)

try:
    from PIL import Image
    HAS_PILLOW = True
except ImportError:
    Image = None
    HAS_PILLOW = False

try:
    import pymupdf4llm
    HAS_PYMUPDF4LLM = True
except ImportError:
    HAS_PYMUPDF4LLM = False

# ── Project paths ───────────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "converted_pdfs"


# ─────────────────────────────────────────────────────────────────────
# EXTRACTORS — each class handles one type of PDF content
# ─────────────────────────────────────────────────────────────────────

class PDFMetadataExtractor:
    """Extract document-level metadata (title, author, dates, etc.)."""

    @staticmethod
    def extract(doc) -> dict:
        """Return a dict of metadata fields with fallback defaults."""
        meta = doc.metadata
        return {
            "title": meta.get("title") or "Untitled",
            "author": meta.get("author") or "Unknown",
            "subject": meta.get("subject") or "",
            "creator": meta.get("creator") or "",
            "created": meta.get("creationDate") or "",
            "modified": meta.get("modDate") or "",
            "producer": meta.get("producer") or "",
            "pages": len(doc),
            "encrypted": doc.is_encrypted,
        }


class URLExtractor:
    """Find and format visible URLs in text content."""

    URL_PATTERN = re.compile(r"https?://[^\s\)]+")

    @staticmethod
    def find_visible_urls(text: str) -> list[str]:
        """Return all URLs found in the given text."""
        return URLExtractor.URL_PATTERN.findall(text)

    @staticmethod
    def markdown_url(text: str, url: str) -> str:
        """Format a display text and URL as a Markdown link."""
        return f"[{text}]({url})"


class ListDetector:
    """Detect and normalize bullet / numbered list items."""

    BULLET_PATTERNS = [r"^[\•\-\*\·▪▸►→✓◆○●◦–—]\s+", r"^\s+[-•◦–]\s+"]
    NUMBERED_PATTERNS = [r"^\d{1,2}[\.\\)]\s+", r"^[a-zA-Z][\.\\)]\s+"]

    @staticmethod
    def is_bullet_item(text: str) -> bool:
        """Return True if *text* starts with a bullet-style marker."""
        return any(re.match(pat, text) for pat in ListDetector.BULLET_PATTERNS)

    @staticmethod
    def is_numbered_item(text: str) -> bool:
        """Return True if *text* starts with a numbered/lettered marker."""
        return any(re.match(pat, text) for pat in ListDetector.NUMBERED_PATTERNS)

    @staticmethod
    def normalize_bullet(text: str) -> str:
        """Replace any bullet variant with a standard Markdown dash '- '."""
        return re.sub(r"^[\•\*\·▪▸►→✓◆○●◦–—]\s+", "- ", text)


class TextFormatter:
    """Apply inline formatting (bold, italic, underline) to span text."""

    @staticmethod
    def apply_formatting(span: dict) -> str:
        """Return Markdown-formatted text from a PyMuPDF text span.

        Applies bold (flag 16), italic (flag 2), and underline (flag 8).
        Strips whitespace and returns empty string for blank spans.
        """
        text = span.get("text", "").strip()
        if not text:
            return ""

        flags = span.get("flags", 0)
        is_bold = bool(flags & 16)
        is_italic = bool(flags & 2)
        is_underline = bool(flags & 8)

        if is_bold and is_italic:
            text = f"***{text}***"
        elif is_bold:
            text = f"**{text}**"
        elif is_italic:
            text = f"*{text}*"

        if is_underline:
            text = f"<u>{text}</u>"

        return text

    @staticmethod
    def get_hex_color(color_tuple) -> str:
        """Convert an RGB float tuple (0.0–1.0 per channel) to a hex string."""
        if not color_tuple or len(color_tuple) < 3:
            return "#000000"
        red, green, blue = color_tuple[:3]
        return f"#{int(red*255):02x}{int(green*255):02x}{int(blue*255):02x}"


class HeadingDetector:
    """Detect headings by comparing font sizes to the page average."""

    @staticmethod
    def get_average_font_size(page) -> float:
        """Calculate the average font size across all spans on *page*."""
        sizes = []
        try:
            text_dict = page.get_text("dict")
            for block in text_dict.get("blocks", []):
                if block["type"] == 0:
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            size = span.get("size", 12)
                            if size > 0:
                                sizes.append(size)
        except Exception as exc:
            print(f"  [warn] Could not calculate font size: {exc}")

        return sum(sizes) / len(sizes) if sizes else 12.0

    @staticmethod
    def detect_heading(font_size: float, avg_size: float, is_bold: bool = False) -> str:
        """Return Markdown heading prefix ('# ', '## ', '### ') or empty string."""
        if font_size == 0:
            return ""

        ratio = font_size / avg_size if avg_size else 1

        if ratio >= 1.8:
            return "# "
        elif ratio >= 1.35:
            return "## "
        elif ratio >= 1.15 or (ratio >= 1.1 and is_bold):
            return "### "
        return ""


class ImageExtractor:
    """Extract and save images from PDF pages."""

    @staticmethod
    def extract_all(page, page_num: int, output_dir: str) -> list[tuple[str, float]]:
        """Extract all images from *page* and save to *output_dir*.

        Returns a list of (markdown_reference, y_position) tuples.
        """
        images = []

        if not output_dir or not Path(output_dir).exists():
            return images

        try:
            image_list = page.get_images(full=True)

            for img_idx, img_ref in enumerate(image_list):
                try:
                    xref = img_ref[0]
                    base_image = page.parent.extract_image(xref)
                    img_bytes = base_image["image"]
                    img_ext = base_image["ext"]
                    img_name = f"page{page_num}_img{img_idx}.{img_ext}"
                    img_path = Path(output_dir) / img_name

                    with open(img_path, "wb") as file:
                        file.write(img_bytes)
                    ImageExtractor.resize_image(str(img_path))

                    img_rects = page.get_image_rects(xref)
                    if not img_rects:
                        continue
                    first_rect = img_rects[0]
                    img_top = first_rect.y0
                    width_inches = first_rect.width / 72.0
                    height_inches = first_rect.height / 72.0
                    rel_path = f"images/{img_name}".replace("\\", "/")
                    md_ref = f'<img src="{rel_path}" style="width:{width_inches:.2f}in;height:{height_inches:.2f}in" />'
                    images.append((md_ref, img_top))

                except Exception as exc:
                    print(f"  [warn] Could not extract image {img_idx}: {exc}")

        except Exception as exc:
            print(f"  [warn] Error extracting images: {exc}")

        return images

    @staticmethod
    def resize_image(img_path: str, max_width: int = 600, max_height: int = 800):
        """Resize an image to fit within *max_width* × *max_height* pixels."""
        if not HAS_PILLOW:
            return
        try:
            with Image.open(img_path) as img:
                img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
                img.save(img_path, optimize=True, quality=85)
        except Exception as exc:
            print(f"  [warn] Could not resize {img_path}: {exc}")


class TableExtractor:
    """Extract tables from PDF pages and convert to Markdown syntax."""

    @staticmethod
    def extract_all(page) -> list[tuple[str, tuple]]:
        """Return a list of (markdown_table, bbox) for each table on *page*."""
        tables = []

        try:
            table_list = page.find_tables()

            for table_idx, table_obj in enumerate(table_list):
                try:
                    data = table_obj.extract()
                    md_table = TableExtractor.to_markdown(data)
                    tables.append((md_table, tuple(table_obj.bbox)))
                except Exception as exc:
                    print(f"  [warn] Could not extract table {table_idx}: {exc}")

        except Exception as exc:
            print(f"  [warn] Error extracting tables: {exc}")

        return tables

    @staticmethod
    def _escape_cell(cell: str) -> str:
        """Escape pipe characters and newlines so Markdown table syntax is preserved."""
        return str(cell).strip().replace("|", "\\|").replace("\n", " ").replace("\r", " ")

    @staticmethod
    def to_markdown(data: list[list]) -> str:
        """Convert a 2D list of cell values into a Markdown table string."""
        if not data:
            return ""

        header = data[0]
        separator = ["---"] * len(header)
        rows = data[1:] if len(data) > 1 else []

        lines = []
        lines.append("| " + " | ".join(TableExtractor._escape_cell(h) for h in header) + " |")
        lines.append("| " + " | ".join(separator) + " |")

        for row in rows:
            while len(row) < len(header):
                row.append("")
            lines.append("| " + " | ".join(TableExtractor._escape_cell(c) for c in row[:len(header)]) + " |")

        return "\n".join(lines)


class DrawingExtractor:
    """Detect full-width horizontal rules from vector graphics."""

    @staticmethod
    def find_horizontal_rules(page) -> list[float]:
        """Return sorted y-positions of full-width horizontal lines on *page*."""
        rules = []
        try:
            try:
                page_width = page.rect.width
            except Exception:
                page_width = 612  # default US Letter width in points
            min_width = max(200, page_width * 0.6)
            for path in page.get_drawings():
                try:
                    rect = path.get("rect")
                    if not rect:
                        continue
                    width = rect[2] - rect[0]
                    height = rect[3] - rect[1]
                    if width >= min_width and 0 < height < 4:
                        rules.append(rect[1])
                except Exception:
                    pass
        except Exception as exc:
            print(f"  [warn] Could not extract drawings: {exc}")
        return sorted(rules)


class AnnotationExtractor:
    """Extract comment and highlight annotations from PDF pages."""

    @staticmethod
    def extract_all(page) -> list[tuple[str, float]]:
        """Return a list of (markdown_annotation, y_position) tuples."""
        annotations = []

        try:
            annots = page.annots()
            if not annots:
                return annotations

            for annot in annots:
                try:
                    annot_type = annot.type[1]
                    content = annot.info.get("content", "")

                    if annot_type == "Text" and content:
                        markdown = f"> 💬 **Note:** {content}"
                        annotations.append((markdown, annot.rect[1]))

                    elif annot_type == "Highlight" and content:
                        markdown = f"==**Highlight:** {content}=="
                        annotations.append((markdown, annot.rect[1]))

                except Exception:
                    pass

        except Exception as exc:
            print(f"  [warn] Could not extract annotations: {exc}")

        return annotations


def _rect_overlap(rect_a: tuple, rect_b: tuple) -> bool:
    """Return True if two rectangles (x0, y0, x1, y1) overlap."""
    try:
        x0_a, y0_a, x1_a, y1_a = rect_a[0], rect_a[1], rect_a[2], rect_a[3]
        x0_b, y0_b, x1_b, y1_b = rect_b[0], rect_b[1], rect_b[2], rect_b[3]
    except (IndexError, TypeError):
        return False
    if x1_a < x0_b or x0_a > x1_b or y1_a < y0_b or y0_a > y1_b:
        return False
    return True


class LinkExtractor:
    """Extract external hyperlinks from PDF pages."""

    @staticmethod
    def extract_all(page) -> list[tuple[str, float]]:
        """Return (markdown_link, y_pos) for standalone link display."""
        out = []
        for _rect, uri, y_pos in LinkExtractor.extract_with_rects(page):
            display_text = uri if len(uri) <= 60 else (uri[:57] + "...")
            out.append((f"[{display_text}]({uri})", y_pos))
        return out

    @staticmethod
    def extract_with_rects(page) -> list[tuple[tuple, str, float]]:
        """Return (rect, uri, y_pos) for each external link on *page*.

        Internal goto links are skipped — those are handled by InternalLinkExtractor.
        """
        links = []
        try:
            for link in page.get_links():
                if link.get("kind") == fitz.LINK_GOTO or ("page" in link and not link.get("uri")):
                    continue
                uri = link.get("uri") or link.get("file")
                if not uri:
                    continue
                rect = link.get("from")
                if not rect:
                    continue
                try:
                    rect_coords = (float(rect[0]), float(rect[1]), float(rect[2]), float(rect[3]))
                except (IndexError, TypeError, ValueError):
                    continue
                links.append((rect_coords, uri, rect_coords[1]))
        except Exception as exc:
            print(f"  [warn] Could not extract links: {exc}")
        return links


class InternalLinkExtractor:
    """Detect and convert internal PDF cross-reference (goto) links."""

    @staticmethod
    def slugify(text: str) -> str:
        """Convert heading text to a Markdown-compatible anchor slug."""
        slug = text.lower().strip()
        slug = re.sub(r"[^\w\s-]", "", slug)
        slug = re.sub(r"[\s_]+", "-", slug)
        slug = slug.strip("-")
        return slug

    @staticmethod
    def extract_with_rects(page) -> list[tuple[tuple, str, float]]:
        """Return (rect, anchor, y_pos) for each internal goto link.

        Anchor format: '#page-N' (1-indexed page number).
        """
        internal_links = []
        try:
            for link in page.get_links():
                kind = link.get("kind", -1)
                if kind != fitz.LINK_GOTO:
                    if not ("page" in link and not link.get("uri")):
                        continue

                target_page = link.get("page")
                if target_page is None:
                    continue
                # PyMuPDF pages are 0-indexed; anchors use 1-indexed
                anchor = f"#page-{target_page + 1}"

                rect = link.get("from")
                if not rect:
                    continue
                try:
                    rect_coords = (float(rect[0]), float(rect[1]), float(rect[2]), float(rect[3]))
                except (IndexError, TypeError, ValueError):
                    continue
                internal_links.append((rect_coords, anchor, rect_coords[1]))
        except Exception as exc:
            print(f"  [warn] Could not extract internal links: {exc}")
        return internal_links


class BookmarkExtractor:
    """Extract PDF bookmarks (outline) as a Markdown table of contents."""

    @staticmethod
    def extract(doc) -> str:
        """Return a Markdown TOC string built from PDF bookmarks, or empty string."""
        try:
            toc = doc.get_toc()
            if not toc:
                return ""

            markdown = "## 📑 Table of Contents\n\n"

            for level, title, page in toc:
                indent = "  " * (level - 1)
                markdown += f"{indent}- [{title}](#page-{page})\n"

            return markdown + "\n"

        except Exception as exc:
            print(f"  [warn] Could not extract bookmarks: {exc}")
            return ""


class EmbeddedFileExtractor:
    """List embedded/attached files in the PDF."""

    @staticmethod
    def extract(doc) -> str:
        """Return a Markdown section listing embedded files, or empty string."""
        try:
            count = doc.embfile_count()
            if count == 0:
                return ""

            markdown = "\n## 📎 Embedded Files\n\n"

            for i in range(count):
                try:
                    info = doc.embfile_info(i)
                    filename = info.get("filename", f"file_{i}")
                    size = info.get("size", 0)
                    size_kb = size / 1024 if size else 0
                    markdown += f"- **{filename}** ({size_kb:.1f} KB)\n"
                except Exception:
                    pass

            return markdown

        except Exception as exc:
            print(f"  [warn] Could not extract embedded files: {exc}")
            return ""


class SecurityHandler:
    """Handle PDF encryption, authentication, and permission reporting."""

    @staticmethod
    def check_encryption(doc) -> tuple[bool, str]:
        """Return (needs_password, message) for the given document."""
        if not doc.is_encrypted:
            return False, ""
        if not doc.needs_pass:
            return False, ""
        return True, "PDF is password-protected"

    @staticmethod
    def authenticate(doc, password: str) -> bool:
        """Authenticate with *password*. Returns True if access is granted."""
        if doc.is_encrypted and doc.needs_pass:
            return doc.authenticate(password)
        return True

    @staticmethod
    def get_permissions_info(doc) -> str:
        """Return a Markdown section describing PDF permissions."""
        try:
            perms = doc.permissions
            can_print = bool(perms & fitz.PDF_PERM_PRINT)
            can_copy = bool(perms & fitz.PDF_PERM_COPY)
            can_modify = bool(perms & fitz.PDF_PERM_MODIFY)

            markdown = "\n## 🔐 Security Info\n\n"
            markdown += f"- Encrypted: {'Yes' if doc.is_encrypted else 'No'}\n"
            markdown += f"- Can Print: {'Yes' if can_print else 'No'}\n"
            markdown += f"- Can Copy: {'Yes' if can_copy else 'No'}\n"
            markdown += f"- Can Modify: {'Yes' if can_modify else 'No'}\n"

            return markdown
        except Exception:
            return ""


# ─────────────────────────────────────────────────────────────────────
# CORE CONVERSION — frontmatter, post-processing, link injection
# ─────────────────────────────────────────────────────────────────────

def create_frontmatter(metadata: dict) -> str:
    """Create YAML frontmatter from document metadata."""
    frontmatter = {
        "title": metadata.get("title", "Untitled"),
        "author": metadata.get("author", "Unknown"),
        "pages": metadata.get("pages", 0),
    }

    if metadata.get("subject"):
        frontmatter["subject"] = metadata["subject"]

    if metadata.get("creator"):
        frontmatter["creator"] = metadata["creator"]

    if HAS_YAML:
        yaml_str = yaml.dump(frontmatter, default_flow_style=False, allow_unicode=True)
    else:
        # Manual fallback when PyYAML is not installed
        lines = []
        for key, value in frontmatter.items():
            if isinstance(value, bool):
                lines.append(f"{key}: {'true' if value else 'false'}")
            elif isinstance(value, str):
                lines.append(f'{key}: "{value}"')
            else:
                lines.append(f"{key}: {value}")
        yaml_str = "\n".join(lines) + "\n"
    return f"---\n{yaml_str}---\n\n"


# ── List marker patterns (pre-compiled) ─────────────────────────────
_BULLET_CHARS = r"•\-\*·▪▸►→✓◆○●◦–—"

_RE_LONE_BULLET = re.compile(
    r"^(?:#{1,6}\s+)?"           # optional heading prefix  (### )
    r"(?:\*{2,3})?"              # optional leading bold/bold-italic
    r"(?:-\s*)?"                 # optional Markdown dash already present
    r"[" + _BULLET_CHARS + r"]"  # the actual bullet character
    r"(?:\*{2,3})?"              # optional trailing bold
    r"\s*$"                      # nothing else on the line
)

_RE_LONE_NUMBER = re.compile(
    r"^(?:#{1,6}\s+)?"           # optional heading prefix
    r"(?:\*{2,3})?"              # optional leading bold
    r"(?:\d{1,3}|[a-zA-Z])"     # the number or letter
    r"[\.)\:]"                   # punctuation after (1. 2) a: etc.)
    r"(?:\*{2,3})?"              # optional trailing bold
    r"\s*$"                      # nothing else
)

# ── Table-cell bullet patterns (pre-compiled) ───────────────────────
_RE_TABLE_BULLET_BR = re.compile(r'([•◦▪‣])\s*<br>\s*(?!<br>)')
_RE_TABLE_BR_BULLET_BR = re.compile(r'<br>\s*([•◦▪‣])\s*<br>\s*(?!<br>)')


def _merge_orphan_list_markers(lines: list[str]) -> list[str]:
    """Merge orphaned bullet / numbered markers with the following non-empty line.

    PDF extraction often puts '•' or '1.' on its own line with the actual
    item text on the next line.  This pass stitches them together so the
    output is valid Markdown list syntax.

    Also handles heading-prefixed markers (### ◦) and bold-wrapped markers
    (**1.**) that the converter may emit.
    """
    merged: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        is_lone_bullet = bool(_RE_LONE_BULLET.match(stripped))
        is_lone_number = bool(_RE_LONE_NUMBER.match(stripped))

        if (is_lone_bullet or is_lone_number) and i + 1 < len(lines):
            # Find the next non-empty line to merge with
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1

            if j < len(lines):
                next_text = lines[j].strip()
                leading_ws = line[: len(line) - len(line.lstrip())]

                if is_lone_bullet:
                    merged.append(f"{leading_ws}- {next_text}")
                else:
                    # Extract the raw numeric/letter marker, stripping
                    # heading hashes and bold asterisks
                    raw = re.sub(r"^#{1,6}\s+", "", stripped)
                    raw = raw.replace("*", "")
                    marker = raw.rstrip()
                    if not marker.endswith(" "):
                        marker += " "
                    merged.append(f"{leading_ws}{marker}{next_text}")

                i = j + 1  # skip the consumed next line
                continue

        merged.append(line)
        i += 1

    return merged


def _unwrap_code_block_bullets(text: str) -> str:
    """Remove code fences that incorrectly wrap list items.

    pymupdf4llm sometimes wraps monospaced bullet / numbered text inside
    ``` fenced blocks.  This detects such blocks and emits the content as
    normal Markdown list items instead.
    """
    lines = text.split("\n")
    out: list[str] = []
    i = 0
    while i < len(lines):
        stripped = lines[i].strip()
        if stripped.startswith("```"):
            block: list[str] = []
            j = i + 1
            while j < len(lines) and not lines[j].strip().startswith("```"):
                block.append(lines[j])
                j += 1
            list_like = any(
                ListDetector.is_bullet_item(block_line.strip()) or ListDetector.is_numbered_item(block_line.strip())
                for block_line in block if block_line.strip()
            )
            if list_like and block:
                for block_line in block:
                    stripped_line = block_line.strip()
                    if not stripped_line:
                        out.append("")
                    elif ListDetector.is_bullet_item(stripped_line):
                        out.append(ListDetector.normalize_bullet(stripped_line))
                    else:
                        out.append(stripped_line)
                i = j + 1 if j < len(lines) else j
                continue
            out.append(lines[i])
            i += 1
            continue
        out.append(lines[i])
        i += 1
    return "\n".join(out)


def _extract_link_display_text(page, rect: tuple) -> str:
    """Extract and clean the visible display text under a link rectangle.

    Returns the cleaned text, or empty string if extraction fails.
    """
    try:
        link_text = page.get_text("text", clip=fitz.Rect(rect)).strip()
    except Exception:
        link_text = ""
    return re.sub(r"\s+", " ", link_text).strip()


def _inject_internal_links(page_chunks: list, pdf_path: str, password: str = None) -> str:
    """Join per-page markdown chunks with page anchors and inject inline links.

    *page_chunks* comes from ``pymupdf4llm.to_markdown(page_chunks=True)``.
    Each element is a dict with at least a ``"text"`` key.

    1. Joins chunks, inserting ``<a id="page-N"></a>`` before each page.
    2. Opens the PDF with fitz and converts plain text that overlaps link
       rects into ``[text](url)`` / ``[text](#page-N)`` Markdown links.

    Does NOT add a TOC or frontmatter.
    Does NOT alter images or their positions.
    """
    # ── 1) Assemble pages with anchors ────────────────────────────────
    parts: list[str] = []
    for idx, chunk in enumerate(page_chunks):
        page_num = idx + 1  # 1-indexed
        parts.append(f'<a id="page-{page_num}"></a>')
        parts.append("")  # blank line after anchor
        text = chunk["text"] if isinstance(chunk, dict) else str(chunk)
        parts.append(text.rstrip())
        parts.append("")  # blank line between pages
    md_text = "\n".join(parts)

    # ── 2) Inline links (external + internal) ─────────────────────────
    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        print(f"  [warn] Could not open PDF for link injection: {exc}")
        return md_text

    if password and doc.is_encrypted and doc.needs_pass:
        doc.authenticate(password)

    link_replacements: list[tuple[str, str]] = []
    for page_idx in range(len(doc)):
        page = doc[page_idx]

        for rect, uri, _ in LinkExtractor.extract_with_rects(page):
            display_text = _extract_link_display_text(page, rect)
            if display_text and len(display_text) >= 2:
                link_replacements.append((display_text, f"[{display_text}]({uri})"))

        for rect, anchor, _ in InternalLinkExtractor.extract_with_rects(page):
            display_text = _extract_link_display_text(page, rect)
            if display_text and len(display_text) >= 2:
                link_replacements.append((display_text, f"[{display_text}]({anchor})"))

    # Longest first to avoid partial matches
    link_replacements.sort(key=lambda x: len(x[0]), reverse=True)
    for plain, md_link in link_replacements:
        escaped = re.escape(plain)
        pattern = r"(?<!\[)" + escaped + r"(?!\]\()"
        md_text = re.sub(pattern, md_link, md_text, count=1)

    doc.close()
    return md_text


def page_to_markdown(page, page_num: int, image_dir: str = None) -> str:
    """Convert a single PDF page to Markdown (fallback pipeline).

    This is used when pymupdf4llm is not available. It processes text blocks,
    images, tables, annotations, and links with positional ordering.
    """
    md_lines = []

    # Page-level anchor for internal link navigation
    md_lines.append(f'<a id="page-{page_num}"></a>')
    md_lines.append("")

    avg_font_size = HeadingDetector.get_average_font_size(page)

    # Extract all page elements
    images = ImageExtractor.extract_all(page, page_num, image_dir) if image_dir else []
    tables = TableExtractor.extract_all(page)
    annotations = AnnotationExtractor.extract_all(page)
    ext_links_with_rects = LinkExtractor.extract_with_rects(page)
    int_links_with_rects = InternalLinkExtractor.extract_with_rects(page)

    text_dict = page.get_text("dict")
    blocks = text_dict.get("blocks", [])

    # Track which links have been inlined to avoid duplicate standalone entries
    inlined_ext_indices = set()
    inlined_int_indices = set()

    # Table bboxes: skip text blocks that overlap tables
    table_bboxes = [table_item[1] for table_item in tables]

    # Collect all elements by y-position for proper ordering
    elements = []  # (y_position, type, content, [font_size, is_bold])

    # ── Build text elements with span-level link matching ─────────────
    for block in blocks:
        if block["type"] != 0:
            continue
        block_bbox = block.get("bbox", [0, 0, 0, 0])
        try:
            block_rect = (float(block_bbox[0]), float(block_bbox[1]),
                          float(block_bbox[2]), float(block_bbox[3]))
        except (IndexError, TypeError, ValueError):
            block_rect = (0, 0, 0, 0)

        if any(_rect_overlap(block_rect, table_bbox) for table_bbox in table_bboxes):
            continue
        block_y = block_rect[1]

        for line in block.get("lines", []):
            line_spans = line.get("spans", [])
            if not line_spans:
                continue

            max_font_size = 0
            is_line_bold = False
            line_parts = []

            for span in line_spans:
                formatted = TextFormatter.apply_formatting(span)
                span_size = span.get("size", 0)
                if span_size > max_font_size:
                    max_font_size = span_size
                if span.get("flags", 0) & 16:
                    is_line_bold = True

                # Build span bounding box for link matching
                span_bbox = span.get("bbox") or span.get("origin")
                if span_bbox is not None:
                    try:
                        if len(span_bbox) >= 4:
                            span_rect = (float(span_bbox[0]), float(span_bbox[1]),
                                         float(span_bbox[2]), float(span_bbox[3]))
                        elif len(span_bbox) >= 2:
                            sx, sy = float(span_bbox[0]), float(span_bbox[1])
                            span_rect = (sx, sy, sx + 1, sy + 1)
                        else:
                            span_rect = (0, 0, 0, 0)
                    except (TypeError, ValueError):
                        span_rect = (0, 0, 0, 0)
                else:
                    span_rect = (0, 0, 0, 0)

                # Check for external link overlap
                matched_ext_idx = None
                for idx, (link_rect, uri, _) in enumerate(ext_links_with_rects):
                    if idx in inlined_ext_indices:
                        continue
                    if _rect_overlap(link_rect, span_rect):
                        matched_ext_idx = idx
                        break

                # Check for internal link overlap
                matched_int_idx = None
                if matched_ext_idx is None:
                    for idx, (link_rect, anchor, _) in enumerate(int_links_with_rects):
                        if idx in inlined_int_indices:
                            continue
                        if _rect_overlap(link_rect, span_rect):
                            matched_int_idx = idx
                            break

                if matched_ext_idx is not None and formatted.strip():
                    line_parts.append(f"[{formatted.strip()}]({ext_links_with_rects[matched_ext_idx][1]})")
                    inlined_ext_indices.add(matched_ext_idx)
                elif matched_int_idx is not None and formatted.strip():
                    line_parts.append(f"[{formatted.strip()}]({int_links_with_rects[matched_int_idx][1]})")
                    inlined_int_indices.add(matched_int_idx)
                else:
                    line_parts.append(formatted)

            line_text = "".join(line_parts).strip()
            if not line_text:
                continue

            elements.append((block_y, "text", line_text, max_font_size, is_line_bold))

    # ── Add non-text elements ─────────────────────────────────────────
    for table_md, table_bbox in tables:
        elements.append((table_bbox[1], "table", table_md))

    for img_md, img_y in images:
        elements.append((img_y, "image", img_md))

    for annot_md, annot_y in annotations:
        elements.append((annot_y, "annotation", annot_md))

    # External links not yet inlined
    inlined_uris = {ext_links_with_rects[i][1] for i in inlined_ext_indices}
    for idx, (_rect, uri, link_y) in enumerate(ext_links_with_rects):
        if idx not in inlined_ext_indices and uri not in inlined_uris:
            display_text = uri if len(uri) <= 60 else (uri[:57] + "...")
            elements.append((link_y, "link", f"[{display_text}]({uri})"))
            inlined_uris.add(uri)

    # Internal links not yet inlined
    inlined_anchors = {int_links_with_rects[i][1] for i in inlined_int_indices}
    for idx, (_rect, anchor, link_y) in enumerate(int_links_with_rects):
        if idx not in inlined_int_indices and anchor not in inlined_anchors:
            page_label = anchor.replace("#page-", "Page ")
            elements.append((link_y, "link", f"[{page_label}]({anchor})"))
            inlined_anchors.add(anchor)

    # Sort by vertical position and render
    elements.sort(key=lambda x: x[0])

    prev_y = None
    for item in elements:
        if len(item) == 5:
            y_pos, elem_type, content, font_size, is_bold = item
        else:
            y_pos, elem_type, content = item
            font_size, is_bold = None, False

        # Add paragraph break if significant vertical gap
        if prev_y is not None and (y_pos - prev_y) > 30:
            md_lines.append("")

        if elem_type == "text":
            heading = HeadingDetector.detect_heading(font_size, avg_font_size, is_bold) if font_size else ""

            if ListDetector.is_bullet_item(content):
                content = ListDetector.normalize_bullet(content)
                md_lines.append(content)
            elif ListDetector.is_numbered_item(content):
                md_lines.append(content)
            elif heading:
                heading_content = content.strip()
                if len(heading_content) > 4 and heading_content.startswith("**") and heading_content.endswith("**") and heading_content.count("**") == 2:
                    heading_content = heading_content[2:-2]
                elif len(heading_content) > 6 and heading_content.startswith("***") and heading_content.endswith("***"):
                    heading_content = heading_content[3:-3]
                md_lines.append(f"{heading}{heading_content}")
            else:
                md_lines.append(content)

        elif elem_type == "table":
            md_lines.append(f"\n{content}\n")

        elif elem_type == "image":
            md_lines.append(content)

        elif elem_type == "annotation":
            md_lines.append(f"\n{content}\n")

        elif elem_type == "link":
            md_lines.append(content)

        prev_y = y_pos

    # ── Post-process: merge orphaned list markers ─────────────────────
    md_lines = _merge_orphan_list_markers(md_lines)

    # Add soft line breaks (two trailing spaces) between consecutive text lines
    for i in range(len(md_lines)):
        line = md_lines[i]
        if not line:
            continue
        if i + 1 >= len(md_lines):
            continue
        next_line = md_lines[i + 1]
        if not next_line:
            continue
        stripped_line = line.strip()
        if stripped_line.startswith("|") or stripped_line.startswith("<") or stripped_line.startswith(">"):
            continue
        md_lines[i] = line.rstrip() + "  "

    return "\n".join(md_lines)


# ─────────────────────────────────────────────────────────────────────
# MAIN CONVERSION ENTRY POINT
# ─────────────────────────────────────────────────────────────────────

def pdf_to_markdown(pdf_path: str, output_path: str = None, image_dir: str = None,
                   password: str = None) -> str:
    """Convert an entire PDF to Markdown.

    When pymupdf4llm is available, uses it for base extraction with automatic
    image writing.  Falls back to the manual page-by-page pipeline otherwise.

    Args:
        pdf_path: Path to the source PDF file.
        output_path: If provided, write the result to this file.
        image_dir: Directory to save extracted images.
        password: Password for encrypted PDFs.

    Returns:
        The full Markdown content as a string.
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    if image_dir:
        Path(image_dir).mkdir(parents=True, exist_ok=True)

    # ── pymupdf4llm path (preferred) ─────────────────────────────────
    if HAS_PYMUPDF4LLM:
        print("📄 Extracting with pymupdf4llm (page_chunks + write_images) ...")
        try:
            # Pre-open and authenticate so pymupdf4llm gets a decrypted doc
            doc = fitz.open(str(pdf_path))
            if doc.is_encrypted:
                if password:
                    if not doc.authenticate(password):
                        raise PermissionError("❌ Invalid password!")
                else:
                    raise PermissionError("❌ PDF is password-protected!")
            chunks = pymupdf4llm.to_markdown(
                doc,
                write_images=True,
                image_path=image_dir or "images",
                image_format="png",
                dpi=150,
                page_chunks=True,
            )
            doc.close()
        except PermissionError:
            raise
        except Exception as exc:
            raise RuntimeError(f"pymupdf4llm extraction failed: {exc}")

        # ── Post-process each chunk ───────────────────────────────────
        # Normalize image paths so they are relative to the output .md file (for correct display)
        if image_dir:
            out_dir = Path(output_path).resolve().parent if output_path else Path(image_dir).resolve().parent
            img_dir_resolved = Path(image_dir).resolve()
            try:
                rel_prefix = img_dir_resolved.relative_to(out_dir)
                md_image_prefix = str(rel_prefix).replace("\\", "/") + "/"
            except ValueError:
                md_image_prefix = "images/"
            # Path variants pymupdf4llm may emit (absolute/relative, forward/back slash)
            abs_posix = img_dir_resolved.as_posix()
            abs_win = str(img_dir_resolved)
            rel_img = str(Path(image_dir)).replace("\\", "/")
            rel_img_win = str(Path(image_dir))
            prefixes_to_replace = [
                abs_posix + "/",
                abs_posix + "\\",
                abs_win + "\\",
                abs_win + "/",
                rel_img + "/",
                rel_img_win + "\\",
                rel_img_win + "/",
            ]

        for chunk in chunks:
            chunk_text = chunk["text"]

            # Fix image paths: use path relative to output .md so images display correctly
            if image_dir:
                for prefix in prefixes_to_replace:
                    chunk_text = chunk_text.replace(prefix, md_image_prefix)

            # Unwrap code-block bullets
            chunk_text = _unwrap_code_block_bullets(chunk_text)

            # Merge orphaned list markers
            chunk_lines = chunk_text.split("\n")
            chunk_lines = _merge_orphan_list_markers(chunk_lines)
            chunk["text"] = "\n".join(chunk_lines)

        # ── Assemble pages with anchors + inject links ────────────────
        result = _inject_internal_links(chunks, str(pdf_path), password)

        # ── Merge orphaned bullets in table cells ─────────────────────
        # pymupdf4llm puts bullet chars (•, ◦, ▪, ‣) on their own <br>
        # segment inside table cells.  Merge them with the following text.
        result = _RE_TABLE_BULLET_BR.sub(r'\1 ', result)
        result = _RE_TABLE_BR_BULLET_BR.sub(r'<br>\1 ', result)

        # Clean up excessive blank lines
        result = re.sub(r"\n{3,}", "\n\n", result).strip()

        if output_path:
            Path(output_path).write_text(result, encoding="utf-8")
            print(f"✅ Saved: {output_path}")

        return result

    # ── Fallback: manual page-by-page pipeline ───────────────────────
    md_content = []

    try:
        doc = fitz.open(pdf_path)

        # Check encryption
        is_encrypted, _enc_msg = SecurityHandler.check_encryption(doc)
        if is_encrypted:
            if password:
                if not SecurityHandler.authenticate(doc, password):
                    raise PermissionError("❌ Invalid password!")
            else:
                raise PermissionError("❌ PDF is password-protected!")

        # Extract metadata and create frontmatter
        metadata = PDFMetadataExtractor.extract(doc)
        md_content.append(create_frontmatter(metadata))

        # Process pages
        print(f"📄 Processing {len(doc)} pages (manual fallback) ...")

        for page_idx in range(len(doc)):
            try:
                page = doc[page_idx]
                page_md = page_to_markdown(page, page_idx + 1, image_dir)
                md_content.append(page_md)
                print(f"  ✓ Page {page_idx + 1}/{len(doc)}")
            except Exception as exc:
                print(f"  ⚠️  Error on page {page_idx + 1}: {exc}")

        # Add embedded files info
        embedded = EmbeddedFileExtractor.extract(doc)
        if embedded:
            md_content.append(embedded)

        doc.close()

    except Exception as exc:
        raise RuntimeError(f"Error processing PDF: {exc}")

    # Combine all content
    result = "\n\n".join(md_content)
    result = re.sub(r"\n{3,}", "\n\n", result).strip()

    if output_path:
        Path(output_path).write_text(result, encoding="utf-8")
        print(f"✅ Saved: {output_path}")

    return result


# ─────────────────────────────────────────────────────────────────────
# CLI ENTRY POINT
# ─────────────────────────────────────────────────────────────────────

def main():
    """Parse CLI arguments and run the PDF to Markdown conversion."""
    parser = argparse.ArgumentParser(
        description="🚀 PDF to Markdown Converter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Output (when -o/--extract-images not set):
  Saved under project folder: converted_pdfs/<pdf_name>/<pdf_name>.md
  Images: converted_pdfs/<pdf_name>/images/

Examples:
  python pdf_to_md_converter.py pdfs/document.pdf
  python pdf_to_md_converter.py pdfs/document.pdf -o output.md --extract-images ./images
  python pdf_to_md_converter.py pdfs/secure.pdf --password mypassword
        """
    )
    parser.add_argument("pdf", help="PDF file to convert")
    parser.add_argument("-o", "--output", help="Output Markdown file (default: converted_pdfs/<name>/<name>.md)")
    parser.add_argument("--extract-images", metavar="DIR", help="Extract images to folder (default: converted_pdfs/<name>/images)")
    parser.add_argument("--password", help="PDF password (if encrypted)")

    args = parser.parse_args()

    try:
        pdf_path = Path(args.pdf)
        stem = pdf_path.stem

        if args.output:
            output = Path(args.output)
        else:
            doc_dir = DEFAULT_OUTPUT_DIR / stem
            doc_dir.mkdir(parents=True, exist_ok=True)
            output = doc_dir / f"{stem}.md"

        if args.extract_images is not None:
            image_dir = args.extract_images
        else:
            image_dir = str(output.parent / "images")

        output.parent.mkdir(parents=True, exist_ok=True)

        print(f"\n🚀 PDF to Markdown Converter")
        print("=" * 60)
        print(f"   Output: {output}")
        if image_dir:
            print(f"   Images: {image_dir}")
        print()

        result = pdf_to_markdown(
            args.pdf,
            str(output),
            image_dir,
            args.password
        )

        print("\n✅ Conversion Complete!")
        print(f"   Total: {len(result)} characters")
        print(f"   Output: {output}")

    except Exception as exc:
        print(f"\n❌ Error: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
