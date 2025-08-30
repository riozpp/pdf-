import os
from typing import List, Tuple

from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
import fitz  # PyMuPDF
from PIL import Image


def parse_page_ranges(ranges: str, max_page: int) -> List[int]:
    """Parse a page range string like "1-3,5,7-9" into a 0-based page index list.
    max_page is the total number of pages in the PDF (1-based upper bound).
    """
    pages: List[int] = []
    if not ranges:
        return pages
    parts = [p.strip() for p in ranges.split(',') if p.strip()]
    for part in parts:
        if '-' in part:
            start_s, end_s = [x.strip() for x in part.split('-', 1)]
            if not start_s.isdigit() or not end_s.isdigit():
                raise ValueError(f"Invalid range: {part}")
            start, end = int(start_s), int(end_s)
            if start < 1 or end < 1 or start > end or end > max_page:
                raise ValueError(f"Range out of bounds: {part}")
            pages.extend(list(range(start - 1, end)))
        else:
            if not part.isdigit():
                raise ValueError(f"Invalid page: {part}")
            page = int(part)
            if page < 1 or page > max_page:
                raise ValueError(f"Page out of bounds: {part}")
            pages.append(page - 1)
    # de-duplicate while preserving order
    seen = set()
    unique_pages: List[int] = []
    for p in pages:
        if p not in seen:
            seen.add(p)
            unique_pages.append(p)
    return unique_pages


def split_pdf(input_path: str, ranges: str, output_dir: str) -> str:
    """Split a PDF by page ranges, producing a new PDF in output_dir.
    Returns the output file path.
    """
    if not os.path.isfile(input_path):
        raise FileNotFoundError(input_path)
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    page_indexes = parse_page_ranges(ranges, total_pages)
    if not page_indexes:
        raise ValueError("No pages to export. Provide ranges like '1-3,5'.")

    writer = PdfWriter()
    for idx in page_indexes:
        writer.add_page(reader.pages[idx])

    base = os.path.splitext(os.path.basename(input_path))[0]
    out_path = os.path.join(output_dir, f"{base}_split.pdf")
    with open(out_path, "wb") as f:
        writer.write(f)
    return out_path


def merge_pdfs(input_paths: List[str], output_path: str) -> str:
    """Merge multiple PDFs into one output_path. Returns output_path."""
    writer = PdfWriter()
    for p in input_paths:
        if not os.path.isfile(p):
            raise FileNotFoundError(p)
        reader = PdfReader(p)
        for page in reader.pages:
            writer.add_page(page)
    out_dir = os.path.dirname(output_path) or os.getcwd()
    os.makedirs(out_dir, exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


def pdf_to_word(input_path: str, output_docx_path: str) -> str:
    """Convert PDF to DOCX using pdf2docx. Returns output path."""
    if not os.path.isfile(input_path):
        raise FileNotFoundError(input_path)
    out_dir = os.path.dirname(output_docx_path) or os.getcwd()
    os.makedirs(out_dir, exist_ok=True)
    cv = Converter(input_path)
    try:
        cv.convert(output_docx_path, start=0, end=None)
    finally:
        cv.close()
    return output_docx_path


def pdf_to_images(input_path: str, output_dir: str, dpi: int = 200, image_format: str = "png") -> List[str]:
    """Convert PDF pages to images. Returns list of image file paths."""
    if not os.path.isfile(input_path):
        raise FileNotFoundError(input_path)
    os.makedirs(output_dir, exist_ok=True)

    image_format = image_format.lower()
    if image_format not in {"png", "jpg", "jpeg"}:
        raise ValueError("image_format must be png/jpg/jpeg")

    doc = fitz.open(input_path)
    output_files: List[str] = []

    # scale by DPI using a zoom matrix
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        mode = "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        ext = "jpg" if image_format in ("jpg", "jpeg") else "png"
        out_path = os.path.join(output_dir, f"page_{page_num + 1}.{ext}")
        img.save(out_path, quality=95)
        output_files.append(out_path)

    return output_files



