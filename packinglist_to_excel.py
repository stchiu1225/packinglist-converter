#!/usr/bin/env python3
"""Diptyque packing list PDF -> Excel converter.

- Supports batch conversion of all PDFs in a folder.
- Removes headers/summary/pallet-colis metadata.
- Keeps one output row per original item line (no merge).
- Output columns are fixed.
"""

from __future__ import annotations

import argparse
import re
import zipfile
import zlib
from pathlib import Path
from typing import Iterable, List
from xml.sax.saxutils import escape

COLUMNS = [
    "Reference",
    "Item Name",
    "Batch Number",
    "ELD / DLP",
    "Quantity",
    "Net Weight",
    "Alcohol Vol",
]

ORDER_REF_RE = re.compile(r"\b(?:SO|PO)\d{6,}-X\d+\b", re.IGNORECASE)
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
BARCODE_RE = re.compile(r"^\d{8,14}$")
QTY_RE = re.compile(r"^\d+$")
DECIMAL_RE = re.compile(r"^\d+(?:[\.,]\d+)?$")
DECIMAL_WITH_SEPARATOR_RE = re.compile(r"^\d+[\.,]\d+$")
BATCH_RE = re.compile(r"^(?=.*[A-Z])(?=.*\d)[A-Z0-9-]{2,20}$", re.IGNORECASE)
REF_RE = re.compile(r"^[A-Z0-9][A-Z0-9-]{2,}$")

SKIP_PREFIXES = (
    "Palette SSCC:",
    "Colis SSCC:",
    "Emballage",
    "Dimension:",
    "Poids brut:",
    "Poids net:",
    "Page :",
)


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _decode_pdf_literal(raw: bytes) -> str:
    """Decode bytes from a PDF literal string `( ... )`."""
    out = bytearray()
    i = 0
    while i < len(raw):
        ch = raw[i]
        if ch != 0x5C:  # backslash
            out.append(ch)
            i += 1
            continue

        i += 1
        if i >= len(raw):
            break
        esc = raw[i]

        mapping = {
            ord("n"): ord("\n"),
            ord("r"): ord("\r"),
            ord("t"): ord("\t"),
            ord("b"): 8,
            ord("f"): 12,
            ord("("): ord("("),
            ord(")"): ord(")"),
            ord("\\"): ord("\\"),
        }
        if esc in mapping:
            out.append(mapping[esc])
            i += 1
            continue

        # Octal escape: \ddd
        if 48 <= esc <= 55:
            oct_digits = [esc]
            i += 1
            for _ in range(2):
                if i < len(raw) and 48 <= raw[i] <= 55:
                    oct_digits.append(raw[i])
                    i += 1
                else:
                    break
            out.append(int(bytes(oct_digits), 8))
            continue

        out.append(esc)
        i += 1

    return out.decode("latin1", errors="replace")


def extract_pdf_strings(pdf_path: Path) -> List[str]:
    """Extract printable strings from compressed PDF content streams.

    This parser targets iText-style packing list PDFs with `(text)Tj` operators.
    """
    blob = pdf_path.read_bytes()
    tokens: List[str] = []

    for stream in re.finditer(rb"stream\r?\n", blob):
        start = stream.end()
        end = blob.find(b"endstream", start)
        if end < 0:
            continue
        data = blob[start:end]
        if data.endswith(b"\r\n"):
            data = data[:-2]
        elif data.endswith(b"\n"):
            data = data[:-1]

        try:
            decoded = zlib.decompress(data)
        except Exception:
            continue

        for match in re.finditer(rb"\((?:\\.|[^\\)])*\)\s*Tj", decoded):
            literal = match.group(0).rsplit(b")", 1)[0][1:]
            text = clean_text(_decode_pdf_literal(literal))
            if text:
                tokens.append(text)

    return tokens


def looks_like_reference(token: str) -> bool:
    first = token.split(" ", 1)[0]
    return bool(REF_RE.match(first)) and not QTY_RE.match(first)


def looks_like_batch(token: str) -> bool:
    return bool(BATCH_RE.match(token)) and not QTY_RE.match(token) and not DATE_RE.match(token)


def should_skip(token: str) -> bool:
    return any(token.startswith(prefix) for prefix in SKIP_PREFIXES)


def parse_rows(tokens: List[str]) -> List[List[str]]:
    rows: List[List[str]] = []

    # Start parsing from first package block if present.
    start_idx = 0
    for idx, tok in enumerate(tokens):
        if tok.startswith("Colis SSCC:") or tok.startswith("Palette SSCC:"):
            start_idx = idx
            break

    i = start_idx
    while i < len(tokens):
        token = tokens[i]

        if should_skip(token):
            i += 1
            continue

        if not looks_like_reference(token):
            i += 1
            continue

        ref, rest = (token.split(" ", 1) + [""])[:2]
        if not ref:
            i += 1
            continue

        name_parts = [rest.strip()] if rest.strip() else []
        j = i + 1

        # Continue item name until a likely batch/qty field starts.
        while j < len(tokens):
            nxt = tokens[j]
            if should_skip(nxt):
                break
            if QTY_RE.match(nxt) or looks_like_batch(nxt):
                break
            name_parts.append(nxt)
            j += 1

        batch = ""
        qty = ""
        net_weight = ""
        alcohol = ""
        eld = ""

        if j < len(tokens) and looks_like_batch(tokens[j]) and (j + 1) < len(tokens) and QTY_RE.match(tokens[j + 1]):
            batch = tokens[j]
            j += 1

        if j < len(tokens) and QTY_RE.match(tokens[j]):
            qty = tokens[j]
            j += 1

        if j < len(tokens) and DECIMAL_RE.match(tokens[j]):
            net_weight = tokens[j]
            j += 1

        if j < len(tokens) and DECIMAL_WITH_SEPARATOR_RE.match(tokens[j]):
            alcohol = tokens[j]
            j += 1

        if j < len(tokens) and DATE_RE.match(tokens[j]):
            eld = tokens[j]
            j += 1

        # Explicitly ignore barcode.
        if j < len(tokens) and BARCODE_RE.match(tokens[j]):
            j += 1

        if qty:
            rows.append([
                ref,
                clean_text(" ".join(name_parts)),
                batch,
                eld,
                qty,
                net_weight,
                alcohol,
            ])

        i = max(j, i + 1)

    return rows


def find_order_ref(tokens: Iterable[str], fallback: str) -> str:
    for token in tokens:
        match = ORDER_REF_RE.search(token)
        if match:
            return match.group(0).upper()
    return fallback


def _cell_xml(cell_ref: str, value: str) -> str:
    text = escape(value or "")
    if text[:1].isspace() or text[-1:].isspace():
        return f'<c r="{cell_ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
    return f'<c r="{cell_ref}" t="inlineStr"><is><t>{text}</t></is></c>'


def write_xlsx(rows: List[List[str]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    all_rows = [COLUMNS] + rows
    sheet_rows = []
    for r_idx, row in enumerate(all_rows, start=1):
        cells = []
        for c_idx, val in enumerate(row, start=1):
            col = chr(ord("A") + c_idx - 1)
            cells.append(_cell_xml(f"{col}{r_idx}", str(val)))
        sheet_rows.append(f"<row r=\"{r_idx}\">{''.join(cells)}</row>")

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>' + "".join(sheet_rows) + "</sheetData></worksheet>"
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Packing List" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )

    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '</Relationships>'
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        '</Relationships>'
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
        '<cellXfs count="1"><xf xfId="0"/></cellXfs>'
        '</styleSheet>'
    )

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/styles.xml", styles_xml)


def convert_pdf(pdf_path: Path, output_dir: Path) -> Path:
    tokens = extract_pdf_strings(pdf_path)
    rows = parse_rows(tokens)
    order_ref = find_order_ref(tokens, pdf_path.stem)
    output_path = output_dir / f"{order_ref}.xlsx"
    write_xlsx(rows, output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert Diptyque packing list PDFs to Excel")
    parser.add_argument("input", nargs="?", default="sample_pdfs", help="Input PDF file or folder (default: sample_pdfs)")
    parser.add_argument("-o", "--output", default="output", help="Output folder (default: output)")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output)

    if input_path.is_file() and input_path.suffix.lower() == ".pdf":
        pdf_files = [input_path]
    else:
        pdf_files = sorted(input_path.glob("*.pdf"))

    if not pdf_files:
        raise SystemExit(f"No PDF files found in: {input_path}")

    for pdf in pdf_files:
        out = convert_pdf(pdf, output_dir)
        print(f"Converted: {pdf.name} -> {out}")


if __name__ == "__main__":
    main()
