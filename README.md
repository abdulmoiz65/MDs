# Markdown Conversion Toolkit

A collection of Python scripts that convert various document formats to clean, well-structured Markdown. Each script is purpose-built for its format to give the best possible output quality.

---

## Scripts Overview

| Script | Input Formats | Key Capability |
|---|---|---|
| `app.py` | docx, pptx, xlsx, xls, doc, ppt, epub, zip, csv, json, xml, wav, mp3, YouTube | Office documents + audio + YouTube |
| `outlook_to_md.py` | `.msg` | Outlook emails with inline images & attachments |
| `pdf_to_md_converter.py` | `.pdf` | PDFs with images, tables, bookmarks, metadata |
| `RTFtoMD.py` | `.rtf` | RTF files and template documents |

---

## Installation

Install all dependencies at once:

```
pip install -r requirements.txt
```

---

## Usage

### `app.py` — Office, Audio & YouTube

Converts Word, Excel, PowerPoint, and other formats to Markdown. Extracts and saves embedded images as real files with correct paths.

```
python app.py <input_file> [options]
python app.py <youtube_url> [options]
```

**Options:**
```
-o, --output OUTPUT    Output .md file path (default: <input_stem>.md)
--preview N            Print first N characters to stdout
```

**Examples:**
```
python app.py report.docx
python app.py slides.pptx --output slides.md
python app.py data.xlsx
python app.py recording.mp3
python app.py https://www.youtube.com/watch?v=dQw4w9WgXcQ
```

**Supported extensions:** `docx`, `doc`, `pptx`, `ppt`, `xlsx`, `xls`, `epub`, `zip`, `csv`, `json`, `xml`, `wav`, `mp3`

**Image handling:**
- `docx` / `pptx` / `xlsx` — images extracted directly from the zip archive and linked inline at their original position
- EMF/WMF vector images auto-converted to PNG via Pillow or ImageMagick

---

### `outlook_to_md.py` — Outlook `.msg` Files

Converts Outlook email messages to Markdown. Handles HTML bodies, inline CID images, attachments, and nested `.msg` attachments recursively.

```
python outlook_to_md.py <file.msg> [options]
python outlook_to_md.py <folder/> --batch [options]
```

**Options:**
```
-o, --output OUTPUT      Output .md file path
--save-attachments       Save non-inline attachments to <stem>_attachments/
--batch                  Convert all .msg files in a directory (recursive)
--preview N              Print first N characters to stdout
```

**Examples:**
```
python outlook_to_md.py email.msg
python outlook_to_md.py email.msg --save-attachments
python outlook_to_md.py email.msg -o report.md
python outlook_to_md.py emails/ --batch
```

**Output includes:** Subject, From, To, CC, BCC, Date, body (HTML → Markdown), inline images saved to `<stem>_images/`, attachment list, nested emails as block-quotes.

---

### `pdf_to_md_converter.py` — PDF Files

Converts PDF documents to Markdown using PyMuPDF, with rich content extraction.

```
python pdf_to_md_converter.py <file.pdf> [options]
```

**Options:**
```
-o, --output OUTPUT        Output .md file (default: converted_pdfs/<name>/<name>.md)
--extract-images DIR       Extract images to folder (default: converted_pdfs/<name>/images)
--password PASSWORD        Password for encrypted PDFs
```

**Examples:**
```
python pdf_to_md_converter.py document.pdf
python pdf_to_md_converter.py document.pdf -o output.md
python pdf_to_md_converter.py secure.pdf --password mypassword
python pdf_to_md_converter.py document.pdf --extract-images ./images
```

**Supports:** text formatting (bold/italic), headings, lists, hyperlinks, images, tables, annotations, bookmarks, metadata, embedded files, encrypted PDFs.

---

### `RTFtoMD.py` — RTF Files

Converts RTF files to Markdown using `striprtf`. Supports single files and batch directory processing.

```
python RTFtoMD.py <file.rtf> [options]
python RTFtoMD.py <folder/> --recursive [options]
```

**Options:**
```
-o, --output OUTPUT        Output .md file or directory (default: converted_rtfs/<name>.md)
-r, --recursive            Recursively scan directories for .rtf files
--encoding ENCODING        Fallback encoding (default: cp1252)
--decode-errors MODE       strict / ignore / replace (default: ignore)
--overwrite                Overwrite existing output files
--no-table-layout          Disable markdown table normalization
--title                    Add top-level title from filename
```

**Examples:**
```
python RTFtoMD.py document.rtf
python RTFtoMD.py document.rtf -o output.md --title
python RTFtoMD.py templates/ --recursive --overwrite
```

---

## Output Structure

```
# app.py / outlook_to_md.py
assets/
  MyDocument.md
  MyDocument_images/
    image_001.png
    ...

# pdf_to_md_converter.py
converted_pdfs/
  MyDocument/
    MyDocument.md
    images/
      image_001.png
      ...

# RTFtoMD.py
converted_rtfs/
  MyDocument.md
  MyDocument_images/
    image_001.png
    ...
```

Image paths inside the `.md` file use `%20`-encoded relative paths so they render correctly in all Markdown viewers.

---

## Dependencies

| Package | Used by |
|---|---|
| `markitdown[pptx,docx,xlsx,xls]` | `app.py` |
| `markitdown[audio-transcription]` | `app.py` (audio) |
| `markitdown[youtube-transcription]` | `app.py` (YouTube) |
| `faster-whisper` | `app.py` (audio transcription, no ffmpeg needed) |
| `extract-msg` | `outlook_to_md.py` |
| `html2text` | `outlook_to_md.py` |
| `striprtf` | `RTFtoMD.py` |
| `pymupdf` | `pdf_to_md_converter.py` |
| `pymupdf4llm` | `pdf_to_md_converter.py` |
| `Pillow` | `app.py` (EMF→PNG), `pdf_to_md_converter.py` |
| `PyYAML` | `pdf_to_md_converter.py` |
