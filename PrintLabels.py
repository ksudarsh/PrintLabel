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
    python PrintLabels.py

Specify CSV and paper:
    python PrintLabels.py --csv addresses.csv --paper letter

More control:
    python PrintLabels.py --csv addresses.csv --paper a4 --output labels.pdf --title-gap 10 --cols 2 --margin 0.5

With sender labels too:
    python PrintLabels.py --csv recipients.csv --sender-csv sender.csv --paper letter

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
DEFAULT_TITLE_GAP = 14.0
LABEL_PAD_X = 0.18 * inch
LABEL_PAD_Y = 0.15 * inch
TITLE_TOP_OFFSET = 0.18 * inch
TITLE_UNDERLINE_GAP = 2
TITLE_UNDERLINE_AFTER = 4
NAME_GAP_AFTER = 3
ADDR_FONT_SIZE = 10.5
ADDR_LINE_STEP = ADDR_FONT_SIZE + 2
SENDER_TOP_OFFSET = 0.38 * inch
SENDER_MAX_TEXT_WIDTH = 2.35 * inch

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


def compute_label_width(page_width: float, margin_inch: float, cols: int, h_gap_inch: float) -> float:
    margin = margin_inch * inch
    h_gap = h_gap_inch * inch
    usable_width = page_width - 2 * margin - (cols - 1) * h_gap
    if usable_width <= 0:
        raise ValueError("Margins/gaps too large for selected paper width.")
    return usable_width / cols


def label_address_lines(rec: LabelRecord, inner_w: float) -> List[str]:
    addr_font = "Helvetica"
    lines: List[str] = []

    lines.extend(wrap_text_to_width(rec.address1, addr_font, ADDR_FONT_SIZE, inner_w))
    if rec.address2:
        lines.extend(wrap_text_to_width(rec.address2, addr_font, ADDR_FONT_SIZE, inner_w))

    city_state_zip = ", ".join(filter(None, [rec.city, rec.state]))
    if rec.zip_code:
        city_state_zip = f"{city_state_zip} {rec.zip_code}".strip()

    lines.extend(wrap_text_to_width(city_state_zip, addr_font, ADDR_FONT_SIZE, inner_w))
    return lines


def inner_text_width(label_width: float, show_title: bool) -> float:
    available = label_width - 2 * LABEL_PAD_X
    if available <= 0:
        raise ValueError("Label width is too small for the configured padding.")
    if show_title:
        return available
    return min(available, SENDER_MAX_TEXT_WIDTH)


def measure_label_height(rec: LabelRecord, label_width: float, title_gap: float, show_title: bool = True) -> float:
    inner_w = inner_text_width(label_width, show_title)

    name_font = "Helvetica-Bold"
    name_size = fit_font_size_for_text(rec.name or "", name_font, inner_w, 11.5, 9)
    lines = label_address_lines(rec, inner_w)

    title_height = 0.0
    if show_title:
        title_font = "Helvetica-Bold"
        title_size = fit_font_size_for_text(TITLE_TEXT, title_font, inner_w, 12, 9)
        title_height = (
            TITLE_TOP_OFFSET
            + title_size
            + TITLE_UNDERLINE_GAP
            + TITLE_UNDERLINE_AFTER
            + title_gap
        )
    else:
        title_height = SENDER_TOP_OFFSET

    return LABEL_PAD_Y + title_height + name_size + NAME_GAP_AFTER + len(lines) * ADDR_LINE_STEP + LABEL_PAD_Y


def compute_layout(
    page_width: float,
    page_height: float,
    margin_inch: float,
    cols: int,
    rows: Optional[int],
    h_gap_inch: float,
    v_gap_inch: float,
    min_label_height: Optional[float] = None,
) -> Layout:
    margin = margin_inch * inch
    h_gap = h_gap_inch * inch
    v_gap = v_gap_inch * inch

    label_width = compute_label_width(page_width, margin_inch, cols, h_gap_inch)

    if rows is None:
        target_label_height = min_label_height or (2.0 * inch)
        rows_guess = int((page_height - 2 * margin + v_gap) // (target_label_height + v_gap))
        rows = max(1, rows_guess)

    usable_height = page_height - 2 * margin - (rows - 1) * v_gap
    if usable_height <= 0:
        raise ValueError("Margins/gaps too large for selected paper height.")

    label_height = usable_height / rows
    if min_label_height is not None and label_height < min_label_height:
        raise ValueError(
            "Label content does not fit with the current page/layout settings. "
            "Reduce columns, margins, or title gap, or specify fewer rows."
        )

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


def repeated_record(rec: LabelRecord, count: int) -> List[LabelRecord]:
    return [
        LabelRecord(
            name=rec.name,
            address1=rec.address1,
            address2=rec.address2,
            city=rec.city,
            state=rec.state,
            zip_code=rec.zip_code,
        )
        for _ in range(count)
    ]


def default_output_path(csv_path: str, paper_name: str, suffix: str = "labels") -> str:
    base = os.path.splitext(os.path.basename(csv_path))[0]
    csv_dir = os.path.dirname(os.path.abspath(csv_path))
    return os.path.join(csv_dir, f"{base}-{suffix}-{paper_name}.pdf")


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
    show_title: bool = True,
    show_cut_lines: bool = True,
    title_color: str = "#4b2323",
    text_color: str = "#222222",
):
    """
    Draw one label in rectangle with lower-left at (x, y).
    """
    if show_cut_lines:
        c.setStrokeColor(lightgrey)
        c.setLineWidth(0.5)
        c.rect(x, y, w, h, stroke=1, fill=0)

    inner_w = inner_text_width(w, show_title)
    inner_x = x + (w - inner_w) / 2
    inner_y_top = y + h - LABEL_PAD_Y

    cursor_y = inner_y_top

    if show_title:
        cursor_y -= TITLE_TOP_OFFSET

        title_font = "Helvetica-Bold"
        title_size = fit_font_size_for_text(TITLE_TEXT, title_font, inner_w, 12, 9)
        c.setFillColor(HexColor(title_color))
        c.setFont(title_font, title_size)
        c.drawString(inner_x, cursor_y, TITLE_TEXT)

        underline_y = cursor_y - TITLE_UNDERLINE_GAP
        title_width = stringWidth(TITLE_TEXT, title_font, title_size)
        c.setStrokeColor(HexColor("#bfa851"))
        c.setLineWidth(1.0)
        c.line(inner_x, underline_y, inner_x + min(title_width, inner_w), underline_y)

        cursor_y -= title_size + TITLE_UNDERLINE_GAP + TITLE_UNDERLINE_AFTER + title_gap
    else:
        cursor_y -= SENDER_TOP_OFFSET

    # Name
    name_font = "Helvetica-Bold"
    name_size = fit_font_size_for_text(rec.name or "", name_font, inner_w, 11.5, 9)
    c.setFillColor(HexColor(text_color))
    c.setFont(name_font, name_size)
    c.drawString(inner_x, cursor_y, rec.name)
    cursor_y -= name_size + NAME_GAP_AFTER

    # Address lines
    addr_font = "Helvetica"
    addr_size = ADDR_FONT_SIZE
    line_step = ADDR_LINE_STEP
    lines = label_address_lines(rec, inner_w)

    c.setFont(addr_font, addr_size)
    for line in lines:
        if cursor_y < y + LABEL_PAD_Y:
            break
        c.drawString(inner_x, cursor_y, line)
        cursor_y -= line_step


def generate_pdf(
    records: List[LabelRecord],
    output_path: str,
    layout: Layout,
    title_gap_points: float,
    show_cut_lines: bool,
    title_flags: Optional[List[bool]] = None,
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

        show_title = True if title_flags is None else title_flags[idx]
        draw_label(
            c,
            rec,
            x,
            y,
            layout.label_width,
            layout.label_height,
            title_gap=title_gap_points,
            show_title=show_title,
            show_cut_lines=show_cut_lines,
        )

    c.save()


# ----------------------------
# Main
# ----------------------------

def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Create printable mailing label PDFs from a CSV.")
    p.add_argument("--csv", help="Path to input CSV")
    p.add_argument("--sender-csv", help="Path to sender CSV containing exactly one address row")
    p.add_argument("--paper", choices=sorted(PAPER_SIZES.keys()), help="Paper size: letter, a4, legal")
    p.add_argument("--output", help="Output PDF path")
    p.add_argument("--sender-output", help="Output PDF path for sender labels")
    p.add_argument("--cols", type=int, default=2, help="Number of label columns (default: 2)")
    p.add_argument("--rows", type=int, default=None, help="Number of label rows (default: auto)")
    p.add_argument("--margin", type=float, default=0.5, help="Page margin in inches (default: 0.5)")
    p.add_argument("--hgap", type=float, default=0.0, help="Horizontal gap between labels in inches (default: 0)")
    p.add_argument("--vgap", type=float, default=0.0, help="Vertical gap between labels in inches (default: 0)")
    p.add_argument("--title-gap", type=float, default=DEFAULT_TITLE_GAP, help=f"Gap between title and name in points (default: {DEFAULT_TITLE_GAP:g})")
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
        output_path = default_output_path(csv_path, paper_name)

    sender_csv_path = args.sender_csv
    if sender_csv_path:
        while not os.path.isfile(sender_csv_path):
            print(f"Sender CSV file not found: {sender_csv_path}")
            sender_csv_path = prompt_if_missing("Enter a valid path to sender CSV file")

    try:
        records = read_csv_records(csv_path)
    except Exception as e:
        print(f"\nError reading CSV:\n{e}")
        sys.exit(1)

    sender_records: Optional[List[LabelRecord]] = None
    sender_output_path: Optional[str] = None
    if sender_csv_path:
        try:
            raw_sender_records = read_csv_records(sender_csv_path)
            if len(raw_sender_records) != 1:
                raise ValueError("Sender CSV must contain exactly one usable address row.")
            sender_records = repeated_record(raw_sender_records[0], len(records))
            if args.sender_output:
                sender_output_path = args.sender_output
        except Exception as e:
            print(f"\nError reading sender CSV:\n{e}")
            sys.exit(1)

    try:
        page_width, page_height = paper_size_from_name(paper_name)
        label_width = compute_label_width(page_width, args.margin, max(1, args.cols), args.hgap)
        min_label_height = max([
            *(measure_label_height(rec, label_width, max(0, args.title_gap), show_title=True) for rec in records),
            *(measure_label_height(rec, label_width, max(0, args.title_gap), show_title=False) for rec in sender_records or []),
        ])
        layout = compute_layout(
            page_width=page_width,
            page_height=page_height,
            margin_inch=args.margin,
            cols=max(1, args.cols),
            rows=args.rows,
            h_gap_inch=args.hgap,
            v_gap_inch=args.vgap,
            min_label_height=min_label_height,
        )
    except Exception as e:
        print(f"\nLayout error:\n{e}")
        sys.exit(1)

    try:
        output_records = records + sender_records if sender_records and not sender_output_path else records
        output_title_flags = ([True] * len(records)) + ([False] * len(sender_records)) if sender_records and not sender_output_path else None
        generate_pdf(
            records=output_records,
            output_path=output_path,
            layout=layout,
            title_gap_points=max(0, args.title_gap),
            show_cut_lines=not args.no_cut_lines,
            title_flags=output_title_flags,
        )
        if sender_records and sender_output_path:
            generate_pdf(
                records=sender_records,
                output_path=sender_output_path,
                layout=layout,
                title_gap_points=max(0, args.title_gap),
                show_cut_lines=not args.no_cut_lines,
                title_flags=[False] * len(sender_records),
            )
    except Exception as e:
        print(f"\nError generating PDF:\n{e}")
        sys.exit(1)

    labels_per_page = layout.cols * layout.rows
    page_count = math.ceil(len(records) / labels_per_page)

    print("\nDone.")
    print(f"CSV rows used: {len(records)}")
    print(f"Output PDF: {output_path}")
    if sender_records and not sender_output_path:
        print(f"Sender labels appended in same PDF: {len(sender_records)}")
    if sender_output_path:
        print(f"Sender output PDF: {sender_output_path}")
    print(f"Paper size: {paper_name}")
    print(f"Layout: {layout.cols} columns x {layout.rows} rows = {labels_per_page} labels/page")
    print(f"Pages: {page_count}")
    print("\nPrint settings:")
    print("- Print at 100% / Actual Size")
    print("- Do not use 'Fit to page'")
    print("- Print on plain paper first and test one sheet")


if __name__ == "__main__":
    main()
