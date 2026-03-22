#!/usr/bin/env python3
"""
Make printable mailing labels from a CSV.

Each label looks like:

    Mantrakshata Prasadam

    Name
    Address line 1
    Address line 2   (optional)
    City, State ZIP

Features
--------
- Prompts for CSV path if not provided
- Prompts for paper size if not provided
- Outputs a PDF you can print on plain paper
- Adds space between title and recipient name
- Draws light cut lines around each label
- Handles common CSV column names
- Wraps long address lines automatically

Install once:
    pip install reportlab

Examples
--------
Interactive:
    python make_labels.py

Specify CSV and paper:
    python make_labels.py --csv addresses.csv --paper letter

More control:
    python make_labels.py --csv addresses.csv --paper a4 --output labels.pdf --title-gap 10 --cols 2 --margin 0.5

Supported paper sizes:
    letter, a4, legal
"""

from __future__ import annotations

import argparse
import csv
import math
import os
import sys
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple

from reportlab.lib.colors import HexColor, black, lightgrey
from reportlab.lib.pagesizes import letter, A4, legal
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


# ----------------------------
# Configuration / defaults
# ----------------------------

TITLE_TEXT = "Mantrakshata Prasadam"

PAPER_SIZES = {
    "letter": letter,
    "a4": A4,
    "legal": legal,
}

COMMON_FIELD_ALIASES = {
    "name": ["name", "full_name", "fullname", "recipient", "person"],
    "address1": ["address1", "address_1", "street", "street1", "line1", "address"],
    "address2": ["address2", "address_2", "street2", "line2", "apt", "suite", "unit"],
    "city": ["city", "town"],
    "state": ["state", "province", "region"],
    "zip": ["zip", "zipcode", "zip_code", "postal", "postal_code", "postcode"],
}


@dataclass
class LabelRecord:
    name: str
    address1: str
    address2: str
    city: str
    state: str
    zip_code: str


@dataclass
class Layout:
    page_width: float
    page_height: float
    margin: float
    cols: int
    rows: int
    h_gap: float
    v_gap: float
    label_width: float
    label_height: float


# ----------------------------
# Helpers
# ----------------------------

def clean_text(value: Optional[str]) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\n", " ").split()).strip()


def prompt_if_missing(prompt_text: str, default: Optional[str] = None) -> str:
    if default:
        full = f"{prompt_text} [{default}]: "
    else:
        full = f"{prompt_text}: "
    value = input(full).strip()
    return value or (default or "")


def normalize_header(h: str) -> str:
    return clean_text(h).lower().replace("-", "_").replace(" ", "_")


def find_column(fieldnames: List[str], logical_name: str) -> Optional[str]:
    normalized = {normalize_header(f): f for f in fieldnames}
    for alias in COMMON_FIELD_ALIASES[logical_name]:
        if alias in normalized:
            return normalized[alias]
    return None


def wrap_text_to_width(text: str, font_name: str, font_size: float, max_width: float) -> List[str]:
    """
    Simple word-wrap using ReportLab stringWidth.
    """
    text = clean_text(text)
    if not text:
        return []

    words = text.split()
    lines: List[str] = []
    current = words[0]

    for word in words[1:]:
        candidate = current + " " + word
        if stringWidth(candidate, font_name, font_size) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word

    lines.append(current)
    return lines


def fit_font_size_for_text(text: str, font_name: str, max_width: float, start_size: float, min_size: float) -> float:
    """
    Shrinks font size until the text fits, or hits min size.
    """
    size = start_size
    while size > min_size and stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    return max(size, min_size)


def paper_size_from_name(name: str) -> Tuple[float, float]:
    key = clean_text(name).lower()
    if key not in PAPER_SIZES:
        valid = ", ".join(PAPER_SIZES.keys())
        raise ValueError(f"Unsupported paper size '{name}'. Supported: {valid}")
    return PAPER_SIZES[key]


def compute_layout(
    page_width: float,
    page_height: float,
    margin_inch: float,
    cols: int,
    rows: Optional[int],
    h_gap_inch: float,
    v_gap_inch: float,
) -> Layout:
    margin = margin_inch * inch
    h_gap = h_gap_inch * inch
    v_gap = v_gap_inch * inch

    usable_width = page_width - 2 * margin - (cols - 1) * h_gap
    if usable_width <= 0:
        raise ValueError("Margins/gaps too large for selected paper width.")

    label_width = usable_width / cols

    if rows is None:
        # Choose a practical default based on paper height and reasonable label height
        target_label_height = 2.0 * inch
        rows_guess = int((page_height - 2 * margin + v_gap) // (target_label_height + v_gap))
        rows = max(1, rows_guess)

    usable_height = page_height - 2 * margin - (rows - 1) * v_gap
    if usable_height <= 0:
        raise ValueError("Margins/gaps too large for selected paper height.")

    label_height = usable_height / rows

    return Layout(
        page_width=page_width,
        page_height=page_height,
        margin=margin,
        cols=cols,
        rows=rows,
        h_gap=h_gap,
        v_gap=v_gap,
        label_width=label_width,
        label_height=label_height,
    )


# ----------------------------
# CSV parsing
# ----------------------------

def read_csv_records(csv_path: str) -> List[LabelRecord]:
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("CSV appears to have no header row.")

        fieldnames = reader.fieldnames

        name_col = find_column(fieldnames, "name")
        address1_col = find_column(fieldnames, "address1")
        address2_col = find_column(fieldnames, "address2")
        city_col = find_column(fieldnames, "city")
        state_col = find_column(fieldnames, "state")
        zip_col = find_column(fieldnames, "zip")

        missing = []
        if not name_col:
            missing.append("name")
        if not address1_col:
            missing.append("address1")
        if not city_col:
            missing.append("city")
        if not state_col:
            missing.append("state")
        if not zip_col:
            missing.append("zip")

        if missing:
            raise ValueError(
                "CSV is missing required columns: "
                + ", ".join(missing)
                + "\nAccepted variants include common names like name/address1/city/state/zip."
            )

        records: List[LabelRecord] = []
        for row in reader:
            name = clean_text(row.get(name_col))
            address1 = clean_text(row.get(address1_col))
            address2 = clean_text(row.get(address2_col)) if address2_col else ""
            city = clean_text(row.get(city_col))
            state = clean_text(row.get(state_col))
            zip_code = clean_text(row.get(zip_col))

            if not any([name, address1, city, state, zip_code]):
                continue

            records.append(
                LabelRecord(
                    name=name,
                    address1=address1,
                    address2=address2,
                    city=city,
                    state=state,
                    zip_code=zip_code,
                )
            )

        if not records:
            raise ValueError("CSV has headers, but no usable address rows were found.")

        return records


# ----------------------------
# Drawing
# ----------------------------

def draw_label(
    c: canvas.Canvas,
    rec: LabelRecord,
    x: float,
    y: float,
    w: float,
    h: float,
    title_gap: float,
    show_cut_lines: bool = True,
    title_color: str = "#4b2323",
    text_color: str = "#222222",
):
    """
    Draw one label in rectangle with lower-left at (x, y).
    """
    pad_x = 0.18 * inch
    pad_y = 0.15 * inch

    if show_cut_lines:
        c.setStrokeColor(lightgrey)
        c.setLineWidth(0.5)
        c.rect(x, y, w, h, stroke=1, fill=0)

    inner_x = x + pad_x
    inner_y_top = y + h - pad_y
    inner_w = w - 2 * pad_x

    # Decorative top accent line (subtle)
    c.setStrokeColor(HexColor("#bfa851"))
    c.setLineWidth(1.0)
    c.line(inner_x, inner_y_top, inner_x + min(1.1 * inch, inner_w * 0.35), inner_y_top)

    cursor_y = inner_y_top - 0.18 * inch

    # Title
    title_font = "Helvetica-Bold"
    title_size = fit_font_size_for_text(TITLE_TEXT, title_font, inner_w, 12, 9)
    c.setFillColor(HexColor(title_color))
    c.setFont(title_font, title_size)
    c.drawString(inner_x, cursor_y, TITLE_TEXT)

    cursor_y -= title_size + title_gap

    # Name
    name_font = "Helvetica-Bold"
    name_size = fit_font_size_for_text(rec.name or "", name_font, inner_w, 11.5, 9)
    c.setFillColor(HexColor(text_color))
    c.setFont(name_font, name_size)
    c.drawString(inner_x, cursor_y, rec.name)
    cursor_y -= name_size + 3

    # Address lines
    addr_font = "Helvetica"
    addr_size = 10.5
    line_step = addr_size + 2

    lines: List[str] = []

    lines.extend(wrap_text_to_width(rec.address1, addr_font, addr_size, inner_w))
    if rec.address2:
        lines.extend(wrap_text_to_width(rec.address2, addr_font, addr_size, inner_w))

    city_state_zip = ", ".join(filter(None, [rec.city, rec.state]))
    if rec.zip_code:
        city_state_zip = f"{city_state_zip} {rec.zip_code}".strip()

    lines.extend(wrap_text_to_width(city_state_zip, addr_font, addr_size, inner_w))

    c.setFont(addr_font, addr_size)
    for line in lines:
        if cursor_y < y + pad_y:
            break
        c.drawString(inner_x, cursor_y, line)
        cursor_y -= line_step


def generate_pdf(
    records: List[LabelRecord],
    output_path: str,
    layout: Layout,
    title_gap_points: float,
    show_cut_lines: bool,
):
    c = canvas.Canvas(output_path, pagesize=(layout.page_width, layout.page_height))
    c.setTitle("Mantrakshata Prasadam Labels")

    labels_per_page = layout.cols * layout.rows

    for idx, rec in enumerate(records):
        page_index = idx // labels_per_page
        position = idx % labels_per_page

        if idx > 0 and position == 0:
            c.showPage()

        row = position // layout.cols
        col = position % layout.cols

        x = layout.margin + col * (layout.label_width + layout.h_gap)
        y = (
            layout.page_height
            - layout.margin
            - (row + 1) * layout.label_height
            - row * layout.v_gap
        )

        draw_label(
            c,
            rec,
            x,
            y,
            layout.label_width,
            layout.label_height,
            title_gap=title_gap_points,
            show_cut_lines=show_cut_lines,
        )

    c.save()


# ----------------------------
# Main
# ----------------------------

def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Create printable mailing label PDFs from a CSV.")
    p.add_argument("--csv", help="Path to input CSV")
    p.add_argument("--paper", choices=sorted(PAPER_SIZES.keys()), help="Paper size: letter, a4, legal")
    p.add_argument("--output", help="Output PDF path")
    p.add_argument("--cols", type=int, default=2, help="Number of label columns (default: 2)")
    p.add_argument("--rows", type=int, default=None, help="Number of label rows (default: auto)")
    p.add_argument("--margin", type=float, default=0.5, help="Page margin in inches (default: 0.5)")
    p.add_argument("--hgap", type=float, default=0.18, help="Horizontal gap between labels in inches (default: 0.18)")
    p.add_argument("--vgap", type=float, default=0.18, help="Vertical gap between labels in inches (default: 0.18)")
    p.add_argument("--title-gap", type=float, default=8.0, help="Gap between title and name in points (default: 8)")
    p.add_argument("--no-cut-lines", action="store_true", help="Do not draw light cut lines")
    return p


def main():
    parser = build_arg_parser()
    args = parser.parse_args()

    csv_path = args.csv or prompt_if_missing("Enter path to CSV file")
    while not os.path.isfile(csv_path):
        print(f"File not found: {csv_path}")
        csv_path = prompt_if_missing("Enter a valid path to CSV file")

    paper_name = args.paper or prompt_if_missing("Paper size (letter / a4 / legal)", "letter").lower()
    while paper_name not in PAPER_SIZES:
        print("Unsupported paper size.")
        paper_name = prompt_if_missing("Paper size (letter / a4 / legal)", "letter").lower()

    output_path = args.output
    if not output_path:
        base = os.path.splitext(os.path.basename(csv_path))[0]
        csv_dir = os.path.dirname(os.path.abspath(csv_path))
        output_path = os.path.join(csv_dir, f"{base}-labels-{paper_name}.pdf")

    try:
        records = read_csv_records(csv_path)
    except Exception as e:
        print(f"\nError reading CSV:\n{e}")
        sys.exit(1)

    try:
        page_width, page_height = paper_size_from_name(paper_name)
        layout = compute_layout(
            page_width=page_width,
            page_height=page_height,
            margin_inch=args.margin,
            cols=max(1, args.cols),
            rows=args.rows,
            h_gap_inch=args.hgap,
            v_gap_inch=args.vgap,
        )
    except Exception as e:
        print(f"\nLayout error:\n{e}")
        sys.exit(1)

    try:
        generate_pdf(
            records=records,
            output_path=output_path,
            layout=layout,
            title_gap_points=max(0, args.title_gap),
            show_cut_lines=not args.no_cut_lines,
        )
    except Exception as e:
        print(f"\nError generating PDF:\n{e}")
        sys.exit(1)

    labels_per_page = layout.cols * layout.rows
    page_count = math.ceil(len(records) / labels_per_page)

    print("\nDone.")
    print(f"CSV rows used: {len(records)}")
    print(f"Output PDF: {output_path}")
    print(f"Paper size: {paper_name}")
    print(f"Layout: {layout.cols} columns x {layout.rows} rows = {labels_per_page} labels/page")
    print(f"Pages: {page_count}")
    print("\nPrint settings:")
    print("- Print at 100% / Actual Size")
    print("- Do not use 'Fit to page'")
    print("- Print on plain paper first and test one sheet")


if __name__ == "__main__":
    main()
