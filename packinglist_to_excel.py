#!/usr/bin/env python3
# Packing List PDF -> Excel converter

import re
from pathlib import Path
import pdfplumber
import pandas as pd

COLUMNS = [
    "Reference",
    "Item Name",
    "Batch Number",
    "ELD / DLP",
    "Quantity",
    "Net Weight",
    "Alcohol Vol",
]

def clean_text(s):
    return re.sub(r"\s+", " ", (s or "")).strip()

def find_order_ref(text, fallback):
    m = re.search(r"(SO\d{6,}-X\d+)", text)
    return m.group(1) if m else fallback

def extract_rows(pdf_path):
    rows = []
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            full_text += text + "\n"
            lines = text.split("\n")
            for i, line in enumerate(lines):
                line = clean_text(line)
                if re.match(r"^[A-Z0-9]{3,}", line):
                    parts = line.split(" ",1)
                    if len(parts) < 2:
                        continue
                    ref = parts[0]
                    name = parts[1]
                    batch=""
                    qty=""
                    net=""
                    alcohol=""
                    expiry=""
                    if i+1 < len(lines):
                        next_line = clean_text(lines[i+1])
                        tokens = next_line.split()
                        if len(tokens) >= 3:
                            batch=tokens[0]
                            qty=tokens[1]
                            net=tokens[2]
                            if len(tokens) >=4:
                                alcohol=tokens[3]
                    if i+2 < len(lines):
                        if re.match(r"\d{2}/\d{2}/\d{4}", clean_text(lines[i+2])):
                            expiry=clean_text(lines[i+2])
                    rows.append([ref,name,batch,expiry,qty,net,alcohol])
    order_ref = find_order_ref(full_text, pdf_path.stem)
    df = pd.DataFrame(rows, columns=COLUMNS)
    return order_ref, df

def main():
    input_dir = Path("PackingList")
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)
    for pdf in input_dir.glob("*.pdf"):
        order, df = extract_rows(pdf)
        df.to_excel(output_dir / f"{order}.xlsx", index=False)

if __name__ == "__main__":
    main()
