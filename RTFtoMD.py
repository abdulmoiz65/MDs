from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import quote

from striprtf.striprtf import rtf_to_text, destinations as _RTF_DESTINATIONS


UNICODE_SPACES_PATTERN = re.compile(r"[\u00A0\u2000-\u200B\u202F\u205F\u3000]")
MULTI_BLANKS_PATTERN = re.compile(r"\n{3,}")
TEMPLATE_MARKER_PATTERN = re.compile(r"^\s*(<\?.*\?>|/?for-each\b.*|end\s+\w+\s*)$", re.IGNORECASE)
INLINE_TEMPLATE_PATTERN = re.compile(r"<\?.*?\?>")


# ─── RTF pre-processors ───────────────────────────────────────────────────────

def _rtf_find_group_end(text: str, start: int) -> int:
    """Return index after the matching '}' for the '{' at `start`."""
    depth = 1
    i = start + 1
    n = len(text)
    while i < n and depth > 0:
        c = text[i]
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
        i += 1
    return i


def expand_listtext_markers(rtf_text: str) -> str:
    """Replace {\\listtext … N.\\tab} groups with 'N. ' so list numbers
    appear in the striprtf plain-text output."""
    marker = '{\\listtext'
    if marker not in rtf_text:
        return rtf_text
    result: list[str] = []
    i = 0
    n = len(rtf_text)
    while i < n:
        if rtf_text[i] == '{' and rtf_text[i:i + len(marker)] == marker:
            group_end = _rtf_find_group_end(rtf_text, i)
            block = rtf_text[i:group_end]
            m = re.search(r'(\w+)\.\s*\\tab', block)
            if m:
                result.append(m.group(1) + '. ')
            i = group_end
        else:
            result.append(rtf_text[i])
            i += 1
    return ''.join(result)


def _extract_image_hex(group: str) -> tuple[str, str]:
    """Extract hex-encoded image data from a \\pict group string.
    Returns (hex_data, extension) where extension is 'png' or 'jpg'."""
    if '\\pngblip' in group:
        marker = '\\pngblip'
        ext = 'png'
    elif '\\jpegblip' in group:
        marker = '\\jpegblip'
        ext = 'jpg'
    else:
        return '', ''
    idx = group.find(marker)
    i = idx + len(marker)
    n = len(group) - 1  # stop before the final closing '}'
    hex_chars: list[str] = []
    while i < n:
        c = group[i]
        if c == '{':
            i = _rtf_find_group_end(group, i)
        elif c == '\\':
            j = i + 1
            while j < n and group[j].isalpha():
                j += 1
            while j < n and (group[j].isdigit() or group[j] == '-'):
                j += 1
            if j < n and group[j] == ' ':
                j += 1
            i = j
        elif c in '0123456789abcdefABCDEF':
            hex_chars.append(c)
            i += 1
        else:
            i += 1
    return ''.join(hex_chars), ext


_IGNORABLE_KW_RE = re.compile(r'^\{\\(\*\\)?([a-z]+)', re.ASCII | re.IGNORECASE)


def _is_ignorable_group_start(text: str) -> bool:
    """Return True if the RTF group opening at `text` is ignorable by striprtf."""
    m = _IGNORABLE_KW_RE.match(text)
    if not m:
        return False
    star, kw = m.group(1), m.group(2).lower()
    return bool(star) or kw in _RTF_DESTINATIONS


def _find_outermost_ignorable_ancestor(rtf_text: str, inner_pos: int) -> tuple[int, int]:
    """Given the start position of {\\*\\shppict}, find the outermost ancestor
    group that is RTF-ignorable (header/footer/shp/etc.).
    Returns (start, end) for the group to replace.
    Falls back to the (inner_pos, its group end) if no ignorable ancestor found.
    """
    depth = 0
    i = inner_pos - 1
    ignorable_ancestors: list[int] = []

    while i > 0:
        c = rtf_text[i]
        if c == '}':
            depth += 1
        elif c == '{':
            if depth == 0:
                if _is_ignorable_group_start(rtf_text[i:i + 60]):
                    ignorable_ancestors.append(i)
            else:
                depth -= 1
        i -= 1

    if ignorable_ancestors:
        outermost = ignorable_ancestors[-1]  # last = outermost (walked inward→outward)
        return outermost, _rtf_find_group_end(rtf_text, outermost)

    return inner_pos, _rtf_find_group_end(rtf_text, inner_pos)


def extract_images_from_rtf(
    rtf_text: str, output_dir: Path, stem: str
) -> tuple[str, list[Path]]:
    """Extract PNG and JPEG images from RTF {\\*\\shppict} groups AND from
    \\fillBlip shape properties, save them to *output_dir*, and replace the
    groups with Markdown image references.

    Correctly handles images buried inside ignorable RTF groups (headers,
    footers, \\shp drawing objects) by bubbling the replacement up to the
    outermost ignorable ancestor so striprtf doesn't swallow it.

    Returns (modified_rtf_text, list_of_saved_image_paths).
    """
    _SHPPICT = '{\\*\\shppict'
    _NONSHPPICT = '{\\nonshppict'
    _FILLBLIP_SN = '{\\sn fillBlip}'

    if 'pngblip' not in rtf_text and 'jpegblip' not in rtf_text:
        return rtf_text, []

    saved: list[Path] = []
    # replacements: (start, end, replacement_text)
    replacements: list[tuple[int, int, str]] = []
    img_index = 0

    def _save_image(hex_data: str, ext: str) -> str:
        """Save hex-decoded image, return markdown reference or ''."""
        nonlocal img_index
        if len(hex_data) < 16:
            return ''
        try:
            img_bytes = bytes.fromhex(hex_data)
            img_index += 1
            output_dir.mkdir(parents=True, exist_ok=True)
            img_filename = f"image_{img_index}.{ext}"
            (output_dir / img_filename).write_bytes(img_bytes)
            saved.append(output_dir / img_filename)
            rel = quote(output_dir.name) + "/" + img_filename
            return f'\n![Image {img_index}]({rel})\n'
        except Exception:
            img_index -= 1
            return ''

    # ── Pass 1: {\\*\\shppict} groups ─────────────────────────────────────────
    pos = 0
    while True:
        found = rtf_text.find(_SHPPICT, pos)
        if found == -1:
            break
        sg_end = _rtf_find_group_end(rtf_text, found)
        group = rtf_text[found:sg_end]
        pos = sg_end

        img_ref = ''
        if 'pngblip' in group or 'jpegblip' in group:
            hex_data, ext = _extract_image_hex(group)
            img_ref = _save_image(hex_data, ext)

        rep_start, rep_end = _find_outermost_ignorable_ancestor(rtf_text, found)
        replacements.append((rep_start, rep_end, img_ref))

    # ── Pass 2: \\fillBlip inside \\shp drawing shapes ────────────────────────
    # Pattern: {\sp{\sn fillBlip}{\sv {\pict ... \jpegblip/\pngblip ... hex}}}
    pos = 0
    while True:
        found = rtf_text.find(_FILLBLIP_SN, pos)
        if found == -1:
            break
        pos = found + len(_FILLBLIP_SN)

        # Navigate to the sibling {\sv ...} group
        sv_start = rtf_text.find('{\\sv ', found)
        if sv_start == -1 or sv_start - found > 200:
            continue
        sv_end = _rtf_find_group_end(rtf_text, sv_start)
        sv_content = rtf_text[sv_start:sv_end]

        if 'pngblip' not in sv_content and 'jpegblip' not in sv_content:
            continue

        # Find the enclosing {\sp ...} group
        sp_start = rtf_text.rfind('{\\sp', 0, found)
        if sp_start == -1:
            continue
        sp_end = _rtf_find_group_end(rtf_text, sp_start)

        hex_data, ext = _extract_image_hex(sv_content)
        img_ref = _save_image(hex_data, ext)
        if not img_ref:
            continue

        # Replace the outermost ignorable ancestor (the \shp group itself)
        rep_start, rep_end = _find_outermost_ignorable_ancestor(rtf_text, sp_start)
        # Avoid double-covering a range already queued
        covered = any(start <= rep_start < end for start, end, _ in replacements)
        if not covered:
            replacements.append((rep_start, rep_end, img_ref))

    # ── Pass 3: remove bare {\\nonshppict} groups not already covered ─────────
    pos = 0
    while True:
        found = rtf_text.find(_NONSHPPICT, pos)
        if found == -1:
            break
        ns_end = _rtf_find_group_end(rtf_text, found)
        covered = any(start <= found < end for start, end, _ in replacements)
        if not covered:
            replacements.append((found, ns_end, ''))
        pos = ns_end

    if not replacements:
        return rtf_text, saved

    # ── Deduplicate overlapping ranges (keep outermost per cluster) ────────────
    replacements.sort(key=lambda x: x[0])
    merged: list[tuple[int, int, str]] = []
    for start, end, rep in replacements:
        if merged and start < merged[-1][1]:
            prev_s, prev_e, prev_r = merged[-1]
            merged[-1] = (prev_s, max(prev_e, end), prev_r + rep)
        else:
            merged.append((start, end, rep))

    # Apply right-to-left so earlier positions stay valid
    result = rtf_text
    for start, end, rep in sorted(merged, key=lambda x: x[0], reverse=True):
        result = result[:start] + rep + result[end:]
    return result, saved


# ─── RTF plain-text post-processor ───────────────────────────────────────────

def fix_interlaced_table_headers(text: str) -> str:
    """Fix bilingual RTF table headers where \\par inside cells causes the
    header to appear as an interlaced diagonal pattern across multiple lines.

    Detects the pattern:
        non-pipe line          ← English label of column 1
        Arabic1|English2       ← Arabic for col-1, English for col-2
        Arabic2|English3
        …
        ArabicN|               ← Arabic for col-N (trailing pipe)
        | data | … |           ← real multi-pipe data row

    Reconstructs a proper single pipe-separated header row.
    """
    lines = text.split('\n')
    result: list[str] = []
    i = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped and '|' not in stripped:
            j = i + 1
            single_pipe: list[str] = []
            while j < len(lines) and lines[j].count('|') == 1:
                single_pipe.append(lines[j])
                j += 1

            if (
                len(single_pipe) >= 2
                and j < len(lines)
                and lines[j].count('|') >= 3
            ):
                cells: list[str] = []
                english_cur = stripped
                for sp in single_pipe:
                    parts = sp.split('|', 1)
                    arabic = parts[0].strip()
                    english_next = parts[1].strip() if len(parts) > 1 else ''
                    cell_text = ' '.join(filter(None, [english_cur, arabic]))
                    cells.append(cell_text)
                    english_cur = english_next
                if english_cur:
                    cells.append(english_cur)
                result.append('| ' + ' | '.join(cells) + ' |')
                i = j
                continue

        result.append(line)
        i += 1

    return '\n'.join(result)


@dataclass(frozen=True)
class ConversionConfig:
    fallback_encoding: str
    decode_errors: str
    overwrite: bool
    preserve_table_layout: bool
    add_document_title: bool


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="RTFtoMD",
        description="Convert RTF files to AI-friendly Markdown using striprtf.",
    )
    parser.add_argument("input", type=Path, help="Input .rtf file or directory")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output .md file for single input, or output directory for folder input",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Recursively scan directories for .rtf files",
    )
    parser.add_argument(
        "--encoding",
        default="cp1252",
        help="Fallback encoding for striprtf when RTF has no explicit codepage (default: cp1252)",
    )
    parser.add_argument(
        "--decode-errors",
        choices=("strict", "ignore", "replace"),
        default="ignore",
        help="How to handle byte decode issues before parsing (default: ignore)",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing output files",
    )
    parser.add_argument(
        "--no-table-layout",
        action="store_true",
        help="Disable markdown table normalization and keep raw text layout",
    )
    parser.add_argument(
        "--title",
        action="store_true",
        help="Add top-level Markdown title from file name",
    )
    return parser.parse_args()


def discover_inputs(input_path: Path, recursive: bool) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".rtf":
            raise ValueError(f"Input file must be .rtf: {input_path}")
        return [input_path]

    if not input_path.is_dir():
        raise ValueError(f"Input path does not exist: {input_path}")

    pattern = "**/*.rtf" if recursive else "*.rtf"
    files = sorted(input_path.glob(pattern))
    if not files:
        raise ValueError(f"No .rtf files found in {input_path}")
    return files


def detect_text(raw_bytes: bytes, decode_errors: str) -> str:
    for encoding in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("latin-1", errors=decode_errors)


def normalize_whitespace(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = UNICODE_SPACES_PATTERN.sub(" ", text)
    text = INLINE_TEMPLATE_PATTERN.sub("", text)
    cleaned_lines: list[str] = []
    for raw_line in text.split("\n"):
        line = raw_line.rstrip()
        if TEMPLATE_MARKER_PATTERN.match(line):
            continue
        if line.strip() == "|":
            continue
        if line.count("|") <= 2 and line.rstrip().endswith("||"):
            line = line.rstrip("|").rstrip()
        if line.count("|") <= 1:
            line = line.replace("|", " ")
            line = re.sub(r"\s{2,}", " ", line).strip()
        if line.strip().lower() == "page of":
            continue
        cleaned_lines.append(line)
    text = "\n".join(cleaned_lines)
    text = MULTI_BLANKS_PATTERN.sub("\n\n", text)
    return text.strip() + "\n"


def looks_like_table_line(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if TEMPLATE_MARKER_PATTERN.match(stripped):
        return False
    if stripped == "|":
        return False
    return "|" in stripped


def has_meaningful_text(line: str) -> bool:
    return bool(line.strip()) and not TEMPLATE_MARKER_PATTERN.match(line)


def merge_wrapped_table_lines(lines: list[str]) -> list[str]:
    merged: list[str] = []
    index = 0
    while index < len(lines):
        current = lines[index]
        if (
            has_meaningful_text(current)
            and "|" not in current
            and ":" not in current
            and index + 1 < len(lines)
            and looks_like_table_line(lines[index + 1])
            and "<?" not in lines[index + 1]
            and len(current.strip().split()) == 1
        ):
            merged.append(f"{current.strip()} {lines[index + 1].lstrip()}")
            index += 2
            continue
        merged.append(current)
        index += 1
    return merged


def parse_table_line(raw: str) -> list[str]:
    parts = [part.strip() for part in raw.split("|")]
    while parts and parts[0] == "":
        parts.pop(0)
    while parts and parts[-1] == "":
        parts.pop()
    return parts


def normalize_table_block(lines: list[str]) -> list[str]:
    rows: list[list[str]] = []
    for raw in lines:
        parts = parse_table_line(raw)
        if parts:
            rows.append(parts)

    if len(rows) < 2:
        return lines

    column_counts = [len(row) for row in rows]
    narrow_rows = sum(1 for count in column_counts if count <= 2)
    if (
        len(rows) >= 6
        and max(column_counts) - min(column_counts) >= 4
        and narrow_rows >= max(3, int(len(column_counts) * 0.5))
    ):
        return lines

    width = max(2, max(len(row) for row in rows))
    padded_rows = [row + [""] * (width - len(row)) for row in rows]

    header = "| " + " | ".join(padded_rows[0]) + " |"
    separator = "| " + " | ".join(["---"] * width) + " |"
    body = ["| " + " | ".join(row) + " |" for row in padded_rows[1:]]

    return [header, separator, *body]


def normalize_tables(text: str) -> str:
    lines = merge_wrapped_table_lines(text.split("\n"))
    output: list[str] = []
    block: list[str] = []

    def flush_block() -> None:
        nonlocal block
        if not block:
            return
        output.extend(normalize_table_block(block))
        block = []

    for line in lines:
        if looks_like_table_line(line):
            block.append(line)
        else:
            flush_block()
            output.append(line)

    flush_block()
    return "\n".join(output)


def render_markdown(text: str, source_file: Path, config: ConversionConfig) -> str:
    normalized = normalize_whitespace(text)
    if config.preserve_table_layout:
        normalized = normalize_tables(normalized)
    if config.add_document_title:
        title = source_file.stem.replace("_", " ").strip()
        normalized = f"# {title}\n\n{normalized}"
    return normalized


def resolve_output_path(
    source: Path,
    input_root: Path,
    output_arg: Path | None,
    multiple_inputs: bool,
) -> Path:
    if source.is_absolute():
        source = source.resolve()

    if output_arg is None:
        # Default: place output in converted_rtfs/ next to this script
        script_dir = Path(__file__).resolve().parent
        return script_dir / "converted_rtfs" / source.with_suffix(".md").name

    if not multiple_inputs and output_arg.suffix.lower() == ".md":
        return output_arg

    relative = source.relative_to(input_root)
    return output_arg / relative.with_suffix(".md")


def convert_file(source_file: Path, destination_file: Path, config: ConversionConfig) -> None:
    if destination_file.exists() and not config.overwrite:
        raise FileExistsError(f"Output exists, use --overwrite: {destination_file}")

    raw_bytes = source_file.read_bytes()
    decoded_rtf = detect_text(raw_bytes, config.decode_errors)

    # Pre-process RTF: extract embedded images and restore list numbering
    image_dir = destination_file.parent / f"{destination_file.stem}_images"
    decoded_rtf, _saved_images = extract_images_from_rtf(
        decoded_rtf, image_dir, destination_file.stem
    )
    decoded_rtf = expand_listtext_markers(decoded_rtf)

    plain_text = rtf_to_text(
        decoded_rtf,
        encoding=config.fallback_encoding,
        errors=config.decode_errors,
    )

    # Post-process: fix interlaced bilingual table headers
    plain_text = fix_interlaced_table_headers(plain_text)

    markdown = render_markdown(plain_text, source_file, config)

    destination_file.parent.mkdir(parents=True, exist_ok=True)
    destination_file.write_text(markdown, encoding="utf-8")


def convert_all(
    sources: Iterable[Path],
    input_root: Path,
    output_arg: Path | None,
    config: ConversionConfig,
) -> tuple[int, int]:
    success = 0
    failed = 0
    source_list = list(sources)
    multiple_inputs = len(source_list) > 1

    for source in source_list:
        destination = resolve_output_path(source, input_root, output_arg, multiple_inputs)
        try:
            convert_file(source, destination, config)
            print(f"OK   {source} -> {destination}")
            success += 1
        except Exception as exc:
            print(f"FAIL {source}: {exc}", file=sys.stderr)
            failed += 1

    return success, failed


def main() -> int:
    args = parse_args()

    config = ConversionConfig(
        fallback_encoding=args.encoding,
        decode_errors=args.decode_errors,
        overwrite=args.overwrite,
        preserve_table_layout=not args.no_table_layout,
        add_document_title=args.title,
    )

    input_path = args.input.resolve()
    output_arg = args.output.resolve() if args.output else None

    try:
        sources = discover_inputs(input_path, args.recursive)
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    input_root = input_path if input_path.is_dir() else input_path.parent
    success, failed = convert_all(sources, input_root, output_arg, config)

    print(f"\nSummary: {success} converted, {failed} failed")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
