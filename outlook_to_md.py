#!/usr/bin/env python3
"""
outlook_to_md.py
----------------
Convert Outlook .msg email files to clean Markdown.

Features
--------
- Metadata header  : Subject, From, To, CC, BCC, Date
- Body             : HTML → Markdown (html2text); fallback to plain text
- Inline images    : extracted from cid: references and linked in the document
- Attachments      : listed; optionally saved to <stem>_attachments/
- Nested .msg      : recursively converted and quoted at the bottom
- Batch mode       : convert every .msg inside a folder tree

Dependencies
------------
    pip install extract-msg html2text

Usage
-----
    python outlook_to_md.py email.msg
    python outlook_to_md.py email.msg -o report.md
    python outlook_to_md.py email.msg --save-attachments
    python outlook_to_md.py emails_folder/ --batch
    python outlook_to_md.py email.msg --preview 500
"""

import argparse
import os
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# Lazy imports — clear error if a dependency is missing
# ---------------------------------------------------------------------------

def _require(module: str, package: str):
    import importlib
    try:
        return importlib.import_module(module)
    except ImportError:
        print(f"[ERROR] Missing dependency — run:  pip install {package}", file=sys.stderr)
        sys.exit(1)


# ---------------------------------------------------------------------------
# HTML → Markdown
# ---------------------------------------------------------------------------

def _html_to_md(html: str) -> str:
    """Convert an HTML string to Markdown via html2text."""
    h2t = _require("html2text", "html2text")
    converter = h2t.HTML2Text()
    converter.ignore_links = False
    converter.ignore_images = False
    converter.body_width = 0        # no forced line wrapping
    converter.unicode_snob = True
    converter.skip_internal_links = True
    return converter.handle(html)


# ---------------------------------------------------------------------------
# Inline CID image extraction
# ---------------------------------------------------------------------------

def _save_inline_images(
    html_bytes: bytes,
    attachments: list,
    images_dir: str,
    md_dir: str,
) -> tuple[str, list]:
    """
    Detect inline-image attachments (those with a Content-ID / cid).
    Save each one to *images_dir* and replace the ``cid:xxx`` reference in the
    HTML with a workspace-relative path so the markdown renderer can display it.

    Returns
    -------
    (modified_html_str, list_of_handled_attachment_names)
    """
    os.makedirs(images_dir, exist_ok=True)
    handled_names: list[str] = []

    try:
        html_str = html_bytes.decode("utf-8", errors="replace")
    except Exception:
        html_str = html_bytes.decode("latin-1", errors="replace")

    for att in attachments:
        cid = getattr(att, "cid", None)
        if not cid:
            continue

        name = (
            getattr(att, "longFilename", None)
            or getattr(att, "shortFilename", None)
            or cid
        )
        data = getattr(att, "data", None)
        if not isinstance(data, (bytes, bytearray)) or not data:
            continue

        # Save image to disk
        out_path = os.path.join(images_dir, name)
        with open(out_path, "wb") as fh:
            fh.write(data)

        # Relative path for the markdown image link — spaces must be %20
        # so the markdown renderer treats the whole token as one URL.
        rel = os.path.relpath(out_path, md_dir).replace("\\", "/")
        rel_encoded = rel.replace(" ", "%20")

        # Replace cid: reference in HTML (with and without angle brackets)
        cid_clean = cid.strip("<>")
        html_str = html_str.replace(f"cid:{cid_clean}", rel_encoded)
        html_str = html_str.replace(f"cid:<{cid_clean}>", rel_encoded)

        handled_names.append(name)
        print(f"  [IMG] Saved inline image → {rel}")

    return html_str, handled_names


# ---------------------------------------------------------------------------
# Core conversion
# ---------------------------------------------------------------------------

def _msg_to_markdown(msg, output_path: str, save_attachments: bool, depth: int) -> str:
    """
    Build the Markdown string from an already-opened extract_msg Message.
    Handles body, inline images, regular attachments, and nested .msg files.
    """
    stem = os.path.splitext(output_path)[0]
    md_dir = os.path.dirname(os.path.abspath(output_path))
    images_dir = stem + "_images"
    att_dir = stem + "_attachments"

    lines: list[str] = []

    # ------------------------------------------------------------------
    # Metadata header
    # ------------------------------------------------------------------
    subject = (msg.subject or "").strip()
    lines.append(f"# {subject or '(No Subject)'}\n")
    lines.append("")

    def _meta(label: str, value):
        v = str(value).strip() if value else None
        if v:
            lines.append(f"**{label}:** {v}  ")

    _meta("From", msg.sender)
    _meta("To",   msg.to)
    _meta("CC",   msg.cc)
    _meta("BCC",  getattr(msg, "bcc", None))
    _meta("Date", msg.date)
    lines.append("")

    # ------------------------------------------------------------------
    # Body
    # ------------------------------------------------------------------
    inline_handled: list[str] = []

    if msg.htmlBody:
        html_bytes = msg.htmlBody
        if isinstance(html_bytes, str):
            html_bytes = html_bytes.encode("utf-8")

        # Replace cid: inline images with saved paths before converting
        html_str, inline_handled = _save_inline_images(
            html_bytes, msg.attachments, images_dir, md_dir
        )
        body_md = _html_to_md(html_str)
    else:
        body_md = (msg.body or "").strip()

    lines.append("## Body\n")
    lines.append(body_md.strip())
    lines.append("")

    # ------------------------------------------------------------------
    # Attachments
    # ------------------------------------------------------------------
    regular_atts: list[tuple[str, object]] = []
    nested_msgs:  list[tuple[str, object]] = []

    for att in msg.attachments:
        name = (
            getattr(att, "longFilename", None)
            or getattr(att, "shortFilename", None)
            or "attachment"
        )
        cid  = getattr(att, "cid", None)
        data = getattr(att, "data", None)

        # Skip inline images already embedded in the body
        if cid and name in inline_handled:
            continue

        # Detect nested .msg (data is a parsed Message object, not raw bytes)
        if data is not None and hasattr(data, "subject"):
            nested_msgs.append((name, data))
        else:
            regular_atts.append((name, data))

    if regular_atts:
        lines.append("## Attachments\n")
        for name, data in regular_atts:
            if save_attachments and isinstance(data, (bytes, bytearray)) and data:
                os.makedirs(att_dir, exist_ok=True)
                out_path = os.path.join(att_dir, name)
                with open(out_path, "wb") as fh:
                    fh.write(data)
                rel = os.path.relpath(out_path, md_dir).replace("\\", "/")
                lines.append(f"- [{name}]({rel})")
                print(f"  [ATT] Saved attachment → {rel}")
            else:
                lines.append(f"- `{name}`")
        lines.append("")

    # ------------------------------------------------------------------
    # Nested .msg attachments (recursive, rendered as block-quotes)
    # ------------------------------------------------------------------
    for name, nested_msg in nested_msgs:
        lines.append(f"---\n\n## Attached Email: {name}\n")
        try:
            nested_md = _msg_to_markdown(nested_msg, output_path, save_attachments, depth + 1)
            # Indent the nested email as a block-quote
            quoted = "\n".join(f"> {ln}" for ln in nested_md.splitlines())
            lines.append(quoted)
        except Exception as exc:
            lines.append(f"> *(Could not convert nested message: {exc})*")
        lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def convert_msg(
    msg_path: str,
    output_path: str,
    save_attachments: bool = False,
) -> str:
    """
    Convert a single .msg file and return the Markdown string.

    Parameters
    ----------
    msg_path         Absolute or relative path to the .msg file.
    output_path      Destination .md path — used to derive image/attachment dirs.
    save_attachments If True, non-inline attachments are saved to disk.
    """
    extract_msg = _require("extract_msg", "extract-msg")
    with extract_msg.openMsg(msg_path) as msg:
        return _msg_to_markdown(msg, output_path, save_attachments, depth=0)


# ---------------------------------------------------------------------------
# Batch mode
# ---------------------------------------------------------------------------

def convert_directory(folder: str, save_attachments: bool) -> None:
    """Recursively convert every .msg file found inside *folder*."""
    msg_files = sorted(Path(folder).glob("**/*.msg"))
    if not msg_files:
        print(f"[WARN] No .msg files found in: {folder}", file=sys.stderr)
        return

    ok = fail = 0
    for msg_path in msg_files:
        out_path = msg_path.with_suffix(".md")
        print(f"Converting: {msg_path.name}  →  {out_path.name}")
        try:
            md = convert_msg(str(msg_path), str(out_path), save_attachments)
            out_path.write_text(md, encoding="utf-8")
            print(f"  Saved: {out_path}")
            ok += 1
        except Exception as exc:
            print(f"  [ERROR] {exc}", file=sys.stderr)
            fail += 1

    print(f"\nDone — {ok} converted, {fail} failed.")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="outlook_to_md",
        description="Convert Outlook .msg email files to Markdown.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python outlook_to_md.py email.msg\n"
            "  python outlook_to_md.py email.msg -o report.md\n"
            "  python outlook_to_md.py email.msg --save-attachments\n"
            "  python outlook_to_md.py emails/ --batch\n"
            "  python outlook_to_md.py email.msg --preview 500\n"
        ),
    )
    parser.add_argument(
        "input",
        help="Path to a .msg file, or a directory when --batch is used.",
    )
    parser.add_argument(
        "-o", "--output",
        metavar="OUTPUT",
        help=(
            "Output .md file path. "
            "Defaults to <input_stem>.md next to the source file."
        ),
    )
    parser.add_argument(
        "--save-attachments",
        action="store_true",
        help="Save non-inline attachments to <stem>_attachments/ next to the .md file.",
    )
    parser.add_argument(
        "--batch",
        action="store_true",
        help="Convert all .msg files found inside a directory (recursive).",
    )
    parser.add_argument(
        "--preview",
        metavar="N",
        type=int,
        default=0,
        help="Print the first N characters of the result to stdout (0 = off).",
    )
    return parser


def main() -> None:
    parser = _build_parser()
    args   = parser.parse_args()
    input_arg: str = args.input

    # ---- Batch mode ----
    if args.batch:
        if not os.path.isdir(input_arg):
            parser.error(f"--batch requires a directory path, got: {input_arg}")
        convert_directory(input_arg, args.save_attachments)
        return

    # ---- Single file ----
    if not os.path.isfile(input_arg):
        parser.error(f"File not found: {input_arg}")
    if not input_arg.lower().endswith(".msg"):
        parser.error("Input must be a .msg file (or pass a folder with --batch).")

    output_path: str = args.output or str(Path(input_arg).with_suffix(".md"))

    print(f"Converting: {input_arg}")
    try:
        md = convert_msg(input_arg, output_path, args.save_attachments)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        sys.exit(1)

    Path(output_path).write_text(md, encoding="utf-8")
    print(f"Saved: {output_path}")

    if args.preview > 0:
        snippet = md[: args.preview]
        if len(md) > args.preview:
            snippet += f"\n\n... (truncated at {args.preview} chars)"
        print("\n--- Preview ---")
        print(snippet)


if __name__ == "__main__":
    main()
