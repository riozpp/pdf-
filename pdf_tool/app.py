import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List

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


def pdf_to_images(input_path: str, output_dir: str, dpi: int = 300, image_format: str = "tiff") -> List[str]:
    """Convert PDF pages to images. Returns list of image file paths."""
    if not os.path.isfile(input_path):
        raise FileNotFoundError(input_path)
    os.makedirs(output_dir, exist_ok=True)

    image_format = image_format.lower()
    if image_format not in {"png", "jpg", "jpeg", "tiff", "tif"}:
        raise ValueError("image_format must be png/jpg/jpeg/tiff/tif")

    doc = fitz.open(input_path)
    output_files: List[str] = []

    # scale by DPI using a zoom matrix
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        # 使用更高的颜色深度和抗锯齿
        pix = page.get_pixmap(matrix=mat, alpha=False, colorspace="rgb")
        mode = "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        
        # 根据格式选择保存参数
        if image_format in ("tiff", "tif"):
            ext = "tiff"
            out_path = os.path.join(output_dir, f"page_{page_num + 1}.{ext}")
            # TIFF无损压缩，LZW压缩
            img.save(out_path, format="TIFF", compression="tiff_lzw", dpi=(dpi, dpi))
        elif image_format in ("jpg", "jpeg"):
            ext = "jpg"
            out_path = os.path.join(output_dir, f"page_{page_num + 1}.{ext}")
            # JPEG高质量
            img.save(out_path, format="JPEG", quality=100, optimize=True, dpi=(dpi, dpi))
        else:  # png
            ext = "png"
            out_path = os.path.join(output_dir, f"page_{page_num + 1}.{ext}")
            # PNG无损压缩
            img.save(out_path, format="PNG", optimize=True, dpi=(dpi, dpi))
        
        output_files.append(out_path)

    return output_files


class PDFToolApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("橙子PDF")
        self.geometry("760x580")
        self.resizable(False, False)

        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True)

        self.split_tab = self._build_split_tab(notebook)
        self.merge_tab = self._build_merge_tab(notebook)
        self.word_tab = self._build_word_tab(notebook)
        self.image_tab = self._build_image_tab(notebook)

        notebook.add(self.split_tab, text="拆分PDF")
        notebook.add(self.merge_tab, text="合并PDF")
        notebook.add(self.word_tab, text="转Word")
        notebook.add(self.image_tab, text="转图片")

    # --- Split Tab ---
    def _build_split_tab(self, parent: ttk.Notebook):
        frame = ttk.Frame(parent, padding=12)

        ttk.Label(frame, text="选择 PDF：").grid(row=0, column=0, sticky=tk.W)
        self.split_pdf_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.split_pdf_path, width=70).grid(row=0, column=1, padx=6)
        ttk.Button(frame, text="浏览", command=self._choose_split_pdf).grid(row=0, column=2)

        ttk.Label(frame, text="页码范围（如 1-3,5）：").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        self.split_ranges = tk.StringVar()
        ttk.Entry(frame, textvariable=self.split_ranges, width=70).grid(row=1, column=1, padx=6, pady=(8, 0))

        ttk.Label(frame, text="输出文件夹：").grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        self.split_out_dir = tk.StringVar()
        ttk.Entry(frame, textvariable=self.split_out_dir, width=70).grid(row=2, column=1, padx=6, pady=(8, 0))
        ttk.Button(frame, text="选择", command=self._choose_split_outdir).grid(row=2, column=2, pady=(8, 0))

        ttk.Button(frame, text="开始拆分", command=self._do_split).grid(row=3, column=1, pady=16)

        self.split_status = tk.StringVar(value="就绪")
        ttk.Label(frame, textvariable=self.split_status, foreground="#666").grid(row=4, column=0, columnspan=3, sticky=tk.W)

        for i in range(3):
            frame.columnconfigure(i, weight=0)
        return frame

    def _choose_split_pdf(self):
        path = filedialog.askopenfilename(title="选择 PDF", filetypes=[["PDF", "*.pdf"]])
        if path:
            self.split_pdf_path.set(path)

    def _choose_split_outdir(self):
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.split_out_dir.set(path)

    def _do_split(self):
        try:
            out = split_pdf(self.split_pdf_path.get(), self.split_ranges.get(), self.split_out_dir.get())
            self.split_status.set(f"完成：{out}")
            messagebox.showinfo("完成", f"已输出：\n{out}")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    # --- Merge Tab ---
    def _build_merge_tab(self, parent: ttk.Notebook):
        frame = ttk.Frame(parent, padding=12)

        ttk.Label(frame, text="选择多个 PDF（按顺序）：").grid(row=0, column=0, sticky=tk.W)
        self.merge_listbox = tk.Listbox(frame, width=80, height=10)
        self.merge_listbox.grid(row=1, column=0, columnspan=3, pady=(6, 0))

        ttk.Button(frame, text="添加 PDF", command=self._add_merge_files).grid(row=2, column=0, pady=8, sticky=tk.W)
        ttk.Button(frame, text="上移", command=lambda: self._move_item(-1)).grid(row=2, column=1, pady=8)
        ttk.Button(frame, text="下移", command=lambda: self._move_item(1)).grid(row=2, column=2, pady=8)

        ttk.Label(frame, text="输出文件：").grid(row=3, column=0, sticky=tk.W)
        self.merge_out_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.merge_out_path, width=70).grid(row=3, column=1, padx=6)
        ttk.Button(frame, text="选择", command=self._choose_merge_outfile).grid(row=3, column=2)

        ttk.Button(frame, text="开始合并", command=self._do_merge).grid(row=4, column=1, pady=16)

        self.merge_status = tk.StringVar(value="就绪")
        ttk.Label(frame, textvariable=self.merge_status, foreground="#666").grid(row=5, column=0, columnspan=3, sticky=tk.W)

        return frame

    def _add_merge_files(self):
        paths = filedialog.askopenfilenames(title="选择 PDF", filetypes=[["PDF", "*.pdf"]])
        for p in paths:
            self.merge_listbox.insert(tk.END, p)

    def _move_item(self, direction: int):
        sel = list(self.merge_listbox.curselection())
        if not sel:
            return
        idx = sel[0]
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= self.merge_listbox.size():
            return
        value = self.merge_listbox.get(idx)
        self.merge_listbox.delete(idx)
        self.merge_listbox.insert(new_idx, value)
        self.merge_listbox.selection_set(new_idx)

    def _choose_merge_outfile(self):
        path = filedialog.asksaveasfilename(title="保存为", defaultextension=".pdf", filetypes=[["PDF", "*.pdf"]])
        if path:
            self.merge_out_path.set(path)

    def _do_merge(self):
        try:
            files = list(self.merge_listbox.get(0, tk.END))
            out = merge_pdfs(files, self.merge_out_path.get())
            self.merge_status.set(f"完成：{out}")
            messagebox.showinfo("完成", f"已输出：\n{out}")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    # --- Word Tab ---
    def _build_word_tab(self, parent: ttk.Notebook):
        frame = ttk.Frame(parent, padding=12)

        ttk.Label(frame, text="选择 PDF：").grid(row=0, column=0, sticky=tk.W)
        self.word_pdf_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.word_pdf_path, width=70).grid(row=0, column=1, padx=6)
        ttk.Button(frame, text="浏览", command=self._choose_word_pdf).grid(row=0, column=2)

        ttk.Label(frame, text="输出 Word：").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        self.word_out_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.word_out_path, width=70).grid(row=1, column=1, padx=6, pady=(8, 0))
        ttk.Button(frame, text="选择", command=self._choose_word_outfile).grid(row=1, column=2, pady=(8, 0))

        ttk.Button(frame, text="开始转换", command=self._do_word).grid(row=2, column=1, pady=16)

        self.word_status = tk.StringVar(value="就绪")
        ttk.Label(frame, textvariable=self.word_status, foreground="#666").grid(row=3, column=0, columnspan=3, sticky=tk.W)

        return frame

    def _choose_word_pdf(self):
        path = filedialog.askopenfilename(title="选择 PDF", filetypes=[["PDF", "*.pdf"]])
        if path:
            self.word_pdf_path.set(path)

    def _choose_word_outfile(self):
        path = filedialog.asksaveasfilename(title="保存为", defaultextension=".docx", filetypes=[["Word", "*.docx"]])
        if path:
            self.word_out_path.set(path)

    def _do_word(self):
        try:
            out = pdf_to_word(self.word_pdf_path.get(), self.word_out_path.get())
            self.word_status.set(f"完成：{out}")
            messagebox.showinfo("完成", f"已输出：\n{out}")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    # --- Image Tab ---
    def _build_image_tab(self, parent: ttk.Notebook):
        frame = ttk.Frame(parent, padding=12)

        ttk.Label(frame, text="选择 PDF：").grid(row=0, column=0, sticky=tk.W)
        self.img_pdf_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.img_pdf_path, width=70).grid(row=0, column=1, padx=6)
        ttk.Button(frame, text="浏览", command=self._choose_img_pdf).grid(row=0, column=2)

        ttk.Label(frame, text="DPI：").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        self.img_dpi = tk.IntVar(value=300)
        ttk.Entry(frame, textvariable=self.img_dpi, width=10).grid(row=1, column=1, sticky=tk.W, pady=(8, 0))

        ttk.Label(frame, text="图片格式：").grid(row=1, column=1, sticky=tk.E, pady=(8, 0))
        self.img_fmt = tk.StringVar(value="tiff")
        ttk.Combobox(frame, textvariable=self.img_fmt, values=["tiff", "png", "jpg"], width=8, state="readonly").grid(row=1, column=2, sticky=tk.W, pady=(8, 0))

        ttk.Label(frame, text="输出文件夹：").grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        self.img_out_dir = tk.StringVar()
        ttk.Entry(frame, textvariable=self.img_out_dir, width=70).grid(row=2, column=1, padx=6, pady=(8, 0))
        ttk.Button(frame, text="选择", command=self._choose_img_outdir).grid(row=2, column=2, pady=(8, 0))

        ttk.Button(frame, text="开始转换", command=self._do_images).grid(row=3, column=1, pady=16)

        self.img_status = tk.StringVar(value="就绪")
        ttk.Label(frame, textvariable=self.img_status, foreground="#666").grid(row=4, column=0, columnspan=3, sticky=tk.W)

        # 添加使用建议
        ttk.Separator(frame, orient='horizontal').grid(row=5, column=0, columnspan=3, sticky='ew', pady=10)
        
        tips_frame = ttk.LabelFrame(frame, text="使用建议", padding=8)
        tips_frame.grid(row=6, column=0, columnspan=3, sticky='ew', pady=(0, 10))
        
        tips_text = """专业印刷：使用TIFF格式，DPI设置为300-600
高质量存档：使用TIFF或PNG格式，DPI设置为300
网页使用：使用PNG格式，DPI设置为200-300
一般用途：使用JPEG格式，DPI设置为200-300"""
        
        tips_label = ttk.Label(tips_frame, text=tips_text, justify=tk.LEFT, foreground="#666")
        tips_label.grid(row=0, column=0, sticky='w')

        return frame

    def _choose_img_pdf(self):
        path = filedialog.askopenfilename(title="选择 PDF", filetypes=[["PDF", "*.pdf"]])
        if path:
            self.img_pdf_path.set(path)

    def _choose_img_outdir(self):
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.img_out_dir.set(path)

    def _do_images(self):
        try:
            files = pdf_to_images(self.img_pdf_path.get(), self.img_out_dir.get(), dpi=int(self.img_dpi.get()), image_format=self.img_fmt.get())
            self.img_status.set(f"完成：{len(files)} 张图片")
            messagebox.showinfo("完成", f"共输出 {len(files)} 张图片。\n示例：\n{files[0] if files else ''}")
        except Exception as e:
            messagebox.showerror("错误", str(e))


if __name__ == "__main__":
    app = PDFToolApp()
    app.mainloop()



