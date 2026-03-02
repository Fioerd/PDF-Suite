"""PDF to CSV Converter Tool.

Two-panel layout: left settings panel, right page preview + report.
Bottom: scrollable thumbnail strip (same pattern as split_tool.py).
"""

from __future__ import annotations

import csv
import datetime
import io
import os
import re
import subprocess
import sys
import tkinter as tk
import unicodedata
from tkinter import filedialog, messagebox, simpledialog
from typing import Optional

import customtkinter as ctk

# ---------------------------------------------------------------------------
# Optional dependency check
# ---------------------------------------------------------------------------
try:
    import fitz
    from PIL import Image, ImageTk
    _HAS_FITZ = True
    _HAS_PIL = True
except ImportError:
    _HAS_FITZ = False
    _HAS_PIL = False

try:
    import pdfplumber
    _HAS_PLUMBER = True
except ImportError:
    _HAS_PLUMBER = False

# ---------------------------------------------------------------------------
# Colors (shared with main.py)
# ---------------------------------------------------------------------------
BG          = "#EEF2F7"
WHITE       = "#FFFFFF"
G100        = "#F3F4F6"
G200        = "#E5E7EB"
G300        = "#D1D5DB"
G400        = "#9CA3AF"
G500        = "#6B7280"
G700        = "#374151"
G900        = "#111827"
BLUE        = "#3B82F6"
BLUE_HOVER  = "#2563EB"
GREEN       = "#16A34A"
RED         = "#EF4444"
THUMB_BG    = "#F0F2F5"

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
ENCODING_MAP = {
    "UTF-8":           "utf-8",
    "UTF-8 with BOM":  "utf-8-sig",
    "UTF-16":          "utf-16",
    "ASCII":           "ascii",
    "Windows-1252":    "cp1252",
    "ISO-8859-1":      "iso-8859-1",
}

DELIMITER_MAP = {
    "Comma (,)":     ",",
    "Semicolon (;)": ";",
    "Tab":           "\t",
    "Pipe (|)":      "|",
}

LINE_ENDING_MAP = {
    "System default": os.linesep,
    "Unix (LF)":      "\n",
    "Windows (CRLF)": "\r\n",
}

# ---------------------------------------------------------------------------
# Helpers shared across tools
# ---------------------------------------------------------------------------

def _pix_to_tk(pixmap):
    if _HAS_PIL:
        img = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        return ImageTk.PhotoImage(img)
    return tk.PhotoImage(data=pixmap.tobytes("ppm"))


def _render_page(doc, idx: int, max_w: int):
    page = doc[idx]
    scale = max_w / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    return _pix_to_tk(pix), scale


def _render_thumb(doc, idx: int, thumb_w: int):
    page = doc[idx]
    scale = thumb_w / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    return _pix_to_tk(pix)


# ---------------------------------------------------------------------------
# Main tool class
# ---------------------------------------------------------------------------

class PDFtoCSVTool:
    LEFT_W  = 460
    THUMB_W = 80

    def __init__(self, parent: ctk.CTkFrame):
        self.parent = parent

        # State
        self.pdf_path: str = ""
        self.output_dir: str = ""
        self._password: str = ""
        self._doc   = None          # fitz.Document
        self._pldoc = None          # pdfplumber document
        self._total_pages = 0
        self._current_page = 0
        self._thumb_imgs: list = []
        self._thumb_render_next = 0
        self._highlighted_thumb_cv = None
        self._page_img = None       # keep PhotoImage alive
        self._page_scale = 1.0
        self._page_ox = 0           # canvas offset x/y for page image
        self._page_oy = 0
        self._table_bboxes: list = []  # [(x0,y0,x1,y1) in PDF pts] for current page
        self._report_frame = None   # report panel (shown post-extraction)

        # ── Core Settings StringVars ──────────────────────────────────────
        self._sv_detection  = tk.StringVar(value="Auto")
        self._sv_row_tol    = tk.StringVar(value="3")
        self._sv_col_tol    = tk.StringVar(value="3")
        self._sv_header     = tk.StringVar(value="Auto-detect")
        self._sv_linebreak  = tk.StringVar(value="Replace with space")
        self._sv_custom_lb  = tk.StringVar(value=" | ")   # custom line-break replacement
        self._sv_merged     = tk.StringVar(value="First column only")
        self._sv_vert_merge = tk.StringVar(value="First row only")
        self._sv_empty_marker = tk.StringVar(value="")    # what to put in empty merged cells
        self._sv_delimiter  = tk.StringVar(value="Comma (,)")
        self._sv_encoding   = tk.StringVar(value="UTF-8 with BOM")
        self._sv_multi      = tk.StringVar(value="Separate file per table")
        self._sv_range      = tk.StringVar(value="all")

        # ── New Tier-2 Settings StringVars ────────────────────────────────
        self._sv_overwrite    = tk.StringVar(value="Rename with suffix")
        self._sv_image_only   = tk.StringVar(value="Skip with warning")
        self._sv_min_rows     = tk.StringVar(value="1")
        self._sv_min_cols     = tk.StringVar(value="1")
        self._sv_source_meta  = tk.StringVar(value="None")
        self._sv_strip_ws     = tk.StringVar(value="Enabled")
        self._sv_line_ending  = tk.StringVar(value="System default")
        self._sv_unicode_norm = tk.StringVar(value="NFC (recommended)")
        self._sv_type_detect  = tk.StringVar(value="Disabled")

        # Widget references needed for dynamic show/hide
        self._custom_lb_row: Optional[ctk.CTkFrame] = None

        if not _HAS_FITZ or not _HAS_PLUMBER:
            self._build_missing_deps()
        else:
            self._build()

    # ======================================================================
    # DEPENDENCY ERROR SCREEN
    # ======================================================================

    def _build_missing_deps(self):
        f = ctk.CTkFrame(self.parent, fg_color=WHITE)
        f.pack(fill="both", expand=True)
        msg = "Missing dependencies:\n"
        if not _HAS_FITZ:
            msg += "  • pymupdf  (pip install pymupdf)\n"
        if not _HAS_PLUMBER:
            msg += "  • pdfplumber  (pip install pdfplumber)\n"
        ctk.CTkLabel(f, text=msg, font=ctk.CTkFont(size=14),
                     text_color=RED, justify="left").pack(expand=True)

    # ======================================================================
    # BUILD UI
    # ======================================================================

    def _build(self):
        # Main two-column layout using pack (enforces left panel width)
        main = ctk.CTkFrame(self.parent, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=0, pady=0)

        # Left settings panel – pack with fixed width
        left_outer = ctk.CTkFrame(main, fg_color=G100, corner_radius=0,
                                   width=self.LEFT_W)
        left_outer.pack(side="left", fill="y")
        left_outer.pack_propagate(False)

        left = ctk.CTkScrollableFrame(left_outer, fg_color="transparent",
                                       corner_radius=0,
                                       scrollbar_button_color=G300,
                                       scrollbar_button_hover_color=G400)
        left.pack(fill="both", expand=True, padx=0, pady=0)

        self._build_left(left)

        # Right panel
        right = ctk.CTkFrame(main, fg_color=WHITE, corner_radius=0)
        right.pack(side="left", fill="both", expand=True)
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        self._right = right
        self._build_right(right)

        # Bottom thumbnail strip
        self._build_thumb_strip()

    # ---- LEFT PANEL -------------------------------------------------------

    def _build_left(self, parent):
        # Title
        ctk.CTkLabel(parent, text="PDF to CSV",
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=G900).pack(anchor="w", padx=18, pady=(18, 14))

        # ── File Input ──────────────────────────────────────────────────
        self._section(parent, "Input File")

        file_row = ctk.CTkFrame(parent, fg_color="transparent")
        file_row.pack(fill="x", padx=18, pady=(0, 8))
        file_row.columnconfigure(0, weight=1)

        self._file_entry = ctk.CTkEntry(
            file_row, state="readonly", fg_color=WHITE,
            border_color=G200, text_color=G700, height=34,
            placeholder_text="No file selected…",
        )
        self._file_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        ctk.CTkButton(
            file_row, text="Browse", width=72, height=34,
            fg_color=BLUE, hover_color=BLUE_HOVER,
            command=self._browse_file,
        ).grid(row=0, column=1)

        # Page range
        range_row = ctk.CTkFrame(parent, fg_color="transparent")
        range_row.pack(fill="x", padx=18, pady=(0, 14))
        ctk.CTkLabel(range_row, text="Page range:",
                     font=ctk.CTkFont(size=12), text_color=G500).pack(side="left")
        ctk.CTkEntry(range_row, textvariable=self._sv_range,
                     width=110, height=30, fg_color=WHITE,
                     border_color=G200, text_color=G700,
                     font=ctk.CTkFont(size=12)).pack(side="left", padx=(8, 0))
        ctk.CTkLabel(range_row, text='e.g. all  1  3-7  1,3,5-7',
                     font=ctk.CTkFont(size=10), text_color=G400).pack(
                     side="left", padx=(8, 0))

        # ── Detection Settings ──────────────────────────────────────────
        self._section(parent, "Table Detection")

        self._dropdown(parent, "Detection method", self._sv_detection,
                       ["Auto", "Lattice", "Stream", "Hybrid"])
        self._labeled_entry(parent, "Row tolerance (pt)", self._sv_row_tol, width=60)
        self._labeled_entry(parent, "Column tolerance (pt)", self._sv_col_tol, width=60)

        # Min table size filter
        self._section(parent, "Table Filters")

        self._labeled_entry(parent, "Min rows (skip smaller)", self._sv_min_rows, width=60)
        self._labeled_entry(parent, "Min columns (skip smaller)", self._sv_min_cols, width=60)

        # Image-only page handling
        self._dropdown(parent, "Image-only pages", self._sv_image_only,
                       ["Skip with warning", "Fail entirely"])

        # ── Extraction Settings ─────────────────────────────────────────
        self._section(parent, "Extraction Settings")

        self._dropdown(parent, "Header row", self._sv_header,
                       ["Auto-detect", "First row is header", "No headers"])

        # Line breaks + conditional custom field
        self._dropdown(parent, "Line breaks in cells", self._sv_linebreak,
                       ["Replace with space", "Replace with custom",
                        "Preserve (\\n in cell)", "Remove entirely"],
                       command=self._on_linebreak_change)

        self._custom_lb_row = ctk.CTkFrame(parent, fg_color="transparent")
        ctk.CTkLabel(self._custom_lb_row, text="Custom replacement:",
                     font=ctk.CTkFont(size=12), text_color=G500,
                     width=160, anchor="w").pack(side="left", padx=(18, 0))
        ctk.CTkEntry(self._custom_lb_row, textvariable=self._sv_custom_lb,
                     width=120, height=28, fg_color=WHITE,
                     border_color=G200, text_color=G700,
                     font=ctk.CTkFont(size=12)).pack(side="left", fill="x", expand=True, padx=(0, 18))
        # hidden by default
        self._custom_lb_row.pack_forget()

        # Horizontal merged cells
        self._dropdown(parent, "Horizontal merged cells", self._sv_merged,
                       ["First column only", "Duplicate across columns", "Leave empty"])

        # Vertical merged cells
        self._dropdown(parent, "Vertical merged cells", self._sv_vert_merge,
                       ["First row only", "Duplicate down rows", "Leave empty"])

        # Empty cell marker
        self._labeled_entry(parent, "Empty cell marker", self._sv_empty_marker, width=80)
        ctk.CTkLabel(parent, text="  (text placed in empty merged cells; blank = leave empty)",
                     font=ctk.CTkFont(size=10), text_color=G400,
                     anchor="w").pack(fill="x", padx=18, pady=(0, 6))

        # Strip whitespace toggle
        self._dropdown(parent, "Strip cell whitespace", self._sv_strip_ws,
                       ["Enabled", "Disabled"])

        # Unicode normalization
        self._dropdown(parent, "Unicode normalization", self._sv_unicode_norm,
                       ["NFC (recommended)", "NFKC (compatibility)", "None"])
        ctk.CTkLabel(parent, text="  (NFC fixes invisible combining chars from PDF fonts)",
                     font=ctk.CTkFont(size=10), text_color=G400,
                     anchor="w").pack(fill="x", padx=18, pady=(0, 6))

        # Type detection
        self._dropdown(parent, "Type detection", self._sv_type_detect,
                       ["Disabled", "Numbers only", "Dates only", "Numbers + Dates"])
        ctk.CTkLabel(parent,
                     text="  ⚠ May alter leading zeros, zip codes, phone numbers",
                     font=ctk.CTkFont(size=10), text_color="#D97706",
                     anchor="w").pack(fill="x", padx=18, pady=(0, 6))

        # ── Output Settings ─────────────────────────────────────────────
        self._section(parent, "Output Settings")

        self._dropdown(parent, "Delimiter", self._sv_delimiter,
                       list(DELIMITER_MAP.keys()))
        self._dropdown(parent, "Encoding", self._sv_encoding,
                       list(ENCODING_MAP.keys()))
        self._dropdown(parent, "Line endings", self._sv_line_ending,
                       list(LINE_ENDING_MAP.keys()))
        self._dropdown(parent, "Multiple tables", self._sv_multi,
                       ["Separate file per table", "Single file (concatenate)"])
        self._dropdown(parent, "Source metadata column", self._sv_source_meta,
                       ["None", "Page number", "Table number", "Page + Table"])

        ctk.CTkLabel(parent, text="  (adds source column(s) in concatenated output)",
                     font=ctk.CTkFont(size=10), text_color=G400,
                     anchor="w").pack(fill="x", padx=18, pady=(0, 4))

        # Overwrite protection
        self._dropdown(parent, "If file exists", self._sv_overwrite,
                       ["Rename with suffix", "Overwrite", "Skip"])

        # Output folder
        folder_row = ctk.CTkFrame(parent, fg_color="transparent")
        folder_row.pack(fill="x", padx=18, pady=(4, 14))
        folder_row.columnconfigure(0, weight=1)

        self._folder_entry = ctk.CTkEntry(
            folder_row, state="readonly", fg_color=WHITE,
            border_color=G200, text_color=G700, height=34,
            placeholder_text="Output folder…",
        )
        self._folder_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        ctk.CTkButton(
            folder_row, text="Browse", width=72, height=34,
            fg_color="transparent", hover_color=G200,
            text_color=G700, border_color=G300, border_width=1,
            command=self._browse_folder,
        ).grid(row=0, column=1)

        # ── Action ──────────────────────────────────────────────────────
        sep = ctk.CTkFrame(parent, fg_color=G200, height=1)
        sep.pack(fill="x", padx=18, pady=(6, 14))

        self._extract_btn = ctk.CTkButton(
            parent, text="Extract to CSV",
            height=40, fg_color=BLUE, hover_color=BLUE_HOVER,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._run_extraction,
        )
        self._extract_btn.pack(fill="x", padx=18, pady=(0, 10))

        self._progress = ctk.CTkProgressBar(parent, progress_color=GREEN,
                                             fg_color=G200, height=8)
        self._progress.set(0)
        self._progress.pack(fill="x", padx=18, pady=(0, 6))

        self._status_lbl = ctk.CTkLabel(
            parent, text="", font=ctk.CTkFont(size=11),
            text_color=G500, anchor="w", wraplength=400,
        )
        self._status_lbl.pack(fill="x", padx=18, pady=(0, 18))

    # ---- RIGHT PANEL -------------------------------------------------------

    def _build_right(self, parent):
        # Canvas for page preview
        canvas_frame = ctk.CTkFrame(parent, fg_color=THUMB_BG, corner_radius=0)
        canvas_frame.grid(row=0, column=0, sticky="nsew")
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self._canvas = tk.Canvas(canvas_frame, bg=THUMB_BG, highlightthickness=0)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        self._canvas.bind("<Configure>", self._on_canvas_resize)

        # Page navigation bar
        nav = ctk.CTkFrame(parent, fg_color=G100, corner_radius=0, height=44)
        nav.grid(row=1, column=0, sticky="ew")
        nav.pack_propagate(False)
        nav.columnconfigure(2, weight=1)

        ctk.CTkButton(nav, text="←", width=34, height=30,
                      fg_color="transparent", hover_color=G200,
                      text_color=G700,
                      command=self._prev_page).pack(side="left", padx=(10, 2), pady=7)
        ctk.CTkButton(nav, text="→", width=34, height=30,
                      fg_color="transparent", hover_color=G200,
                      text_color=G700,
                      command=self._next_page).pack(side="left", padx=(2, 10), pady=7)

        self._page_lbl = ctk.CTkLabel(nav, text="No file loaded",
                                       font=ctk.CTkFont(size=12),
                                       text_color=G500)
        self._page_lbl.pack(side="left")

        # Placeholder text on canvas
        self._show_placeholder()

    # ---- THUMBNAIL STRIP ---------------------------------------------------

    def _build_thumb_strip(self):
        strip_frame = ctk.CTkFrame(self.parent, fg_color=G100,
                                    corner_radius=0, height=155)
        strip_frame.pack(fill="x", side="bottom")
        strip_frame.pack_propagate(False)

        # Scrollable canvas inside strip
        self._strip_canvas = tk.Canvas(strip_frame, bg=G100,
                                        highlightthickness=0, height=150)
        scrollbar = tk.Scrollbar(strip_frame, orient="horizontal",
                                  command=self._strip_canvas.xview)
        self._strip_canvas.configure(xscrollcommand=scrollbar.set)
        scrollbar.pack(side="bottom", fill="x")
        self._strip_canvas.pack(fill="both", expand=True)

        self._strip_inner = tk.Frame(self._strip_canvas, bg=G100)
        self._strip_window = self._strip_canvas.create_window(
            (0, 0), window=self._strip_inner, anchor="nw"
        )
        self._strip_inner.bind(
            "<Configure>",
            lambda e: self._strip_canvas.configure(
                scrollregion=self._strip_canvas.bbox("all")
            ),
        )

    # ======================================================================
    # UI HELPERS
    # ======================================================================

    def _section(self, parent, title: str):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=(12, 6))
        ctk.CTkLabel(row, text=title,
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=G700).pack(side="left")
        ctk.CTkFrame(row, fg_color=G300, height=1).pack(
            side="left", fill="x", expand=True, padx=(8, 0))

    def _dropdown(self, parent, label: str, var: tk.StringVar, values: list,
                  command=None):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=(0, 6))
        lbl = ctk.CTkLabel(row, text=label + ":",
                            font=ctk.CTkFont(size=12), text_color=G500,
                            width=160, anchor="w")
        lbl.pack(side="left")
        om = ctk.CTkOptionMenu(row, variable=var, values=values,
                               height=28,
                               fg_color=WHITE, button_color=G300,
                               button_hover_color=G400,
                               text_color=G700,
                               dropdown_fg_color=WHITE,
                               dropdown_text_color=G700,
                               dropdown_hover_color=G100,
                               font=ctk.CTkFont(size=12),
                               command=command)
        om.pack(side="left", fill="x", expand=True)

    def _labeled_entry(self, parent, label: str, var: tk.StringVar, width=80):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=(0, 6))
        ctk.CTkLabel(row, text=label + ":",
                     font=ctk.CTkFont(size=12), text_color=G500,
                     width=160, anchor="w").pack(side="left")
        ctk.CTkEntry(row, textvariable=var, width=width, height=28,
                     fg_color=WHITE, border_color=G200,
                     text_color=G700, font=ctk.CTkFont(size=12)).pack(side="left", fill="x", expand=True)

    def _show_placeholder(self):
        self._canvas.delete("all")
        self._canvas.create_text(
            self._canvas.winfo_reqwidth() // 2 or 300,
            150,
            text="Open a PDF to preview it here",
            font=("Segoe UI", 13),
            fill=G400,
        )

    def _on_linebreak_change(self, value: str):
        """Show/hide the custom replacement field based on selection."""
        if value == "Replace with custom":
            self._custom_lb_row.pack(fill="x", pady=(0, 6))
        else:
            self._custom_lb_row.pack_forget()

    # ======================================================================
    # FILE / FOLDER BROWSE
    # ======================================================================

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Open PDF",
            filetypes=[("PDF files", "*.pdf")],
        )
        if path:
            self._open_pdf(path)

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_dir = path
            self._folder_entry.configure(state="normal")
            self._folder_entry.delete(0, "end")
            self._folder_entry.insert(0, path)
            self._folder_entry.configure(state="readonly")

    # ======================================================================
    # PDF OPEN / VALIDATION
    # ======================================================================

    def _open_pdf(self, path: str):
        # Close any previously open docs
        if self._doc:
            try:
                self._doc.close()
            except Exception:
                pass
        if self._pldoc:
            try:
                self._pldoc.close()
            except Exception:
                pass
        self._doc = None
        self._pldoc = None
        self._password = ""

        try:
            doc = fitz.open(path)
        except Exception as e:
            messagebox.showerror("Cannot Open PDF",
                                 f"Failed to open file:\n{e}")
            return

        # Encrypted?
        if doc.needs_pass:
            pw = simpledialog.askstring(
                "Password Required",
                "This PDF is password-protected.\nEnter password:",
                show="*",
            )
            if pw is None:
                doc.close()
                return
            if not doc.authenticate(pw):
                messagebox.showerror("Wrong Password",
                                     "Incorrect password. Cannot open PDF.")
                doc.close()
                return
            self._password = pw

        if doc.page_count == 0:
            messagebox.showerror("Empty PDF", "This PDF contains no pages.")
            doc.close()
            return

        # Check for text layer
        if not self._detect_text_layer(doc):
            messagebox.showwarning(
                "No Text Layer Detected",
                "This PDF appears to contain scanned images without an "
                "extractable text layer.\n\n"
                "OCR support will be added in a future release.\n"
                "Extraction may produce empty tables.",
            )

        # Open pdfplumber handle
        try:
            if self._password:
                self._pldoc = pdfplumber.open(path, password=self._password)
            else:
                self._pldoc = pdfplumber.open(path)
        except Exception as e:
            messagebox.showerror("pdfplumber Error",
                                 f"Could not open PDF with pdfplumber:\n{e}")
            doc.close()
            return

        self._doc = doc
        self.pdf_path = path
        self._total_pages = doc.page_count
        self._current_page = 0

        # Set default output dir to same folder as PDF
        if not self.output_dir:
            self.output_dir = os.path.dirname(path)
            self._folder_entry.configure(state="normal")
            self._folder_entry.delete(0, "end")
            self._folder_entry.insert(0, self.output_dir)
            self._folder_entry.configure(state="readonly")

        # Update file entry
        self._file_entry.configure(state="normal")
        self._file_entry.delete(0, "end")
        self._file_entry.insert(0, os.path.basename(path))
        self._file_entry.configure(state="readonly")

        # Build thumbnails
        self._build_thumbnails()
        self._render_page_canvas()
        self._status_lbl.configure(
            text=f"{doc.page_count} page{'s' if doc.page_count != 1 else ''} loaded.")

    def _detect_text_layer(self, doc) -> bool:
        sample_pages = min(3, doc.page_count)
        for i in range(sample_pages):
            text = doc[i].get_text().strip()
            if len(text) > 10:
                return True
        return False

    def _page_is_image_only(self, page_idx: int) -> bool:
        """Return True if the fitz page has no extractable text."""
        if not self._doc:
            return False
        text = self._doc[page_idx].get_text().strip()
        return len(text) < 5

    # ======================================================================
    # THUMBNAILS
    # ======================================================================

    def _build_thumbnails(self):
        # Clear strip
        for w in self._strip_inner.winfo_children():
            w.destroy()
        self._thumb_imgs.clear()
        self._highlighted_thumb_cv = None
        self._thumb_render_next = 0

        for i in range(self._total_pages):
            frame = tk.Frame(self._strip_inner, bg=G100)
            frame.pack(side="left", padx=4, pady=8)

            cv = tk.Canvas(frame, width=self.THUMB_W, height=110,
                            bg=THUMB_BG, highlightthickness=1,
                            highlightbackground=G300)
            cv.pack()
            lbl = tk.Label(frame, text=str(i + 1),
                           font=("Segoe UI", 9), bg=G100, fg=G500)
            lbl.pack()

            cv.bind("<Button-1>", lambda e, idx=i: self._go_to_page(idx))
            lbl.bind("<Button-1>", lambda e, idx=i: self._go_to_page(idx))
            self._thumb_imgs.append((None, frame, cv))

        self._render_thumb_batch()

    def _render_thumb_batch(self, batch=8):
        if not self._doc:
            return
        start = self._thumb_render_next
        end   = min(start + batch, self._total_pages)
        for i in range(start, end):
            img_old, frame, cv = self._thumb_imgs[i]
            if img_old is not None:
                continue
            try:
                img = _render_thumb(self._doc, i, self.THUMB_W)
            except Exception:
                continue
            self._thumb_imgs[i] = (img, frame, cv)
            th = img.height() if hasattr(img, "height") else 110
            cv.configure(width=self.THUMB_W, height=th)
            cv.create_image(0, 0, anchor="nw", image=img)
        self._thumb_render_next = end
        if end < self._total_pages:
            self._strip_inner.after(0, self._render_thumb_batch)
        else:
            self._highlight_thumb(self._current_page)

    def _highlight_thumb(self, idx: int):
        if self._highlighted_thumb_cv is not None:
            try:
                self._highlighted_thumb_cv.configure(
                    highlightbackground=G300, highlightthickness=1)
            except Exception:
                pass
        if 0 <= idx < len(self._thumb_imgs):
            _, _, cv = self._thumb_imgs[idx]
            cv.configure(highlightbackground=BLUE, highlightthickness=2)
            self._highlighted_thumb_cv = cv
        else:
            self._highlighted_thumb_cv = None

    # ======================================================================
    # PAGE RENDERING
    # ======================================================================

    def _render_page_canvas(self):
        if not self._doc:
            return
        self._canvas.delete("all")
        cw = max(self._canvas.winfo_width(), 100)
        ch = max(self._canvas.winfo_height(), 100)

        try:
            img, scale = _render_page(self._doc, self._current_page, cw - 20)
        except Exception as e:
            self._canvas.create_text(cw // 2, ch // 2,
                                     text=f"Render error: {e}",
                                     font=("Segoe UI", 11), fill=RED)
            return

        self._page_img = img
        self._page_scale = scale
        iw = img.width() if hasattr(img, "width") else cw
        ih = img.height() if hasattr(img, "height") else ch
        self._page_ox = (cw - iw) // 2
        self._page_oy = 10
        self._canvas.create_image(self._page_ox, self._page_oy,
                                   anchor="nw", image=img)

        # Update nav label
        self._page_lbl.configure(
            text=f"Page {self._current_page + 1} / {self._total_pages}")

        # Overlay table detection outlines
        self._draw_table_outlines()
        self._highlight_thumb(self._current_page)

    def _on_canvas_resize(self, event=None):
        if self._doc:
            self._render_page_canvas()

    def _draw_table_outlines(self):
        """Draw dashed blue rectangles on the canvas for each detected table."""
        if not self._pldoc:
            return
        try:
            pl_page = self._pldoc.pages[self._current_page]
            tables = pl_page.find_tables(self._build_table_settings())
        except Exception:
            return

        self._table_bboxes = []

        for table in tables:
            # pdfplumber bbox: (x0, top, x1, bottom) in pts from top-left
            x0, top, x1, bottom = table.bbox
            self._table_bboxes.append((x0, top, x1, bottom))

            # Map to canvas coordinates
            cx0 = self._page_ox + x0 * self._page_scale
            cy0 = self._page_oy + top * self._page_scale
            cx1 = self._page_ox + x1 * self._page_scale
            cy1 = self._page_oy + bottom * self._page_scale

            self._canvas.create_rectangle(
                cx0, cy0, cx1, cy1,
                outline=BLUE, width=2, dash=(6, 3),
            )

    # ======================================================================
    # NAVIGATION
    # ======================================================================

    def _go_to_page(self, idx: int):
        if not self._doc:
            return
        idx = max(0, min(idx, self._total_pages - 1))
        self._current_page = idx
        self._render_page_canvas()

    def _prev_page(self):
        self._go_to_page(self._current_page - 1)

    def _next_page(self):
        self._go_to_page(self._current_page + 1)

    # ======================================================================
    # TABLE SETTINGS BUILDER
    # ======================================================================

    def _build_table_settings(self, method: Optional[str] = None) -> dict:
        if method is None:
            method = self._sv_detection.get()
        try:
            row_tol = max(1, int(self._sv_row_tol.get()))
        except ValueError:
            row_tol = 3
        try:
            col_tol = max(1, int(self._sv_col_tol.get()))
        except ValueError:
            col_tol = 3

        if method == "Lattice":
            v_strat = "lines"
            h_strat = "lines"
        elif method == "Stream":
            v_strat = "text"
            h_strat = "text"
        elif method == "Hybrid":
            v_strat = "lines_strict"
            h_strat = "lines_strict"
        else:  # Auto — start with lines, fall back handled in extraction
            v_strat = "lines"
            h_strat = "lines"

        return {
            "vertical_strategy":        v_strat,
            "horizontal_strategy":      h_strat,
            "intersection_y_tolerance": row_tol,
            "intersection_x_tolerance": col_tol,
            "snap_y_tolerance":         row_tol,
            "snap_x_tolerance":         col_tol,
            "edge_min_length":          3,
            "min_words_vertical":       1,
            "min_words_horizontal":     1,
            "keep_blank_chars":         False,
            "text_tolerance":           3,
            "text_x_tolerance":         3,
            "text_y_tolerance":         3,
            "explicit_vertical_lines":  [],
            "explicit_horizontal_lines": [],
        }

    # ======================================================================
    # PAGE RANGE PARSER
    # ======================================================================

    @staticmethod
    def _parse_page_range(spec: str, total: int) -> list[int]:
        spec = spec.strip().lower()
        if spec in ("", "all"):
            return list(range(total))
        pages: set[int] = set()
        for part in spec.split(","):
            part = part.strip()
            if "-" in part:
                lo, _, hi = part.partition("-")
                lo_i = int(lo.strip()) - 1
                hi_i = int(hi.strip()) - 1
                if lo_i < 0 or hi_i >= total or lo_i > hi_i:
                    raise ValueError(f"Page range '{part}' is out of bounds "
                                     f"(document has {total} pages).")
                pages.update(range(lo_i, hi_i + 1))
            else:
                idx = int(part) - 1
                if idx < 0 or idx >= total:
                    raise ValueError(f"Page number {part} is out of bounds "
                                     f"(document has {total} pages).")
                pages.add(idx)
        return sorted(pages)

    # ======================================================================
    # OUTPUT PATH HELPERS (overwrite protection)
    # ======================================================================

    def _resolve_output_path(self, fpath: str) -> Optional[str]:
        """
        Apply the overwrite-protection policy to *fpath*.
        Returns the final path to write to, or None if the file should be
        skipped (policy = "Skip").
        """
        if not os.path.exists(fpath):
            return fpath

        policy = self._sv_overwrite.get()

        if policy == "Overwrite":
            return fpath

        if policy == "Skip":
            return None  # caller interprets None as "skipped"

        # "Rename with suffix" — append _1, _2, … until free
        base, ext = os.path.splitext(fpath)
        counter = 1
        while True:
            candidate = f"{base}_{counter}{ext}"
            if not os.path.exists(candidate):
                return candidate
            counter += 1

    # ======================================================================
    # TABLE PROCESSING
    # ======================================================================

    def _process_table(self, raw: list) -> list[list[str]]:
        """Clean a raw pdfplumber table (list of rows, each row a list of cells)."""
        lb_mode      = self._sv_linebreak.get()
        custom_lb    = self._sv_custom_lb.get()
        merge_h_mode = self._sv_merged.get()
        merge_v_mode = self._sv_vert_merge.get()
        empty_marker = self._sv_empty_marker.get()
        strip_ws     = self._sv_strip_ws.get() == "Enabled"
        uni_norm     = self._sv_unicode_norm.get()
        type_detect  = self._sv_type_detect.get()

        rows: list[list[str]] = []

        # First pass: basic per-cell cleaning
        for raw_row in raw:
            if raw_row is None:
                continue
            row: list[str] = []
            for cell in raw_row:
                if cell is None:
                    # Horizontal merged cell placeholder from pdfplumber
                    if merge_h_mode == "Duplicate across columns" and row:
                        row.append(row[-1])
                    elif empty_marker:
                        row.append(empty_marker)
                    else:
                        row.append("")
                    continue

                text = str(cell)

                # Line break handling
                if lb_mode == "Replace with space":
                    text = text.replace("\n", " ").replace("\r", " ")
                elif lb_mode == "Replace with custom":
                    text = text.replace("\n", custom_lb).replace("\r", "")
                elif lb_mode == "Remove entirely":
                    text = text.replace("\n", "").replace("\r", "")
                # else "Preserve": leave as-is

                # Strip whitespace (optional)
                if strip_ws:
                    text = text.strip()

                # Smart quotes → straight quotes
                text = (text.replace("\u2018", "'").replace("\u2019", "'")
                            .replace("\u201c", '"').replace("\u201d", '"'))

                # Ligature expansion
                text = (text.replace("\ufb01", "fi").replace("\ufb02", "fl")
                            .replace("\ufb00", "ff").replace("\ufb03", "ffi")
                            .replace("\ufb04", "ffl"))

                # Remove control characters (except standard whitespace)
                text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)

                # Collapse multiple spaces
                text = re.sub(r"  +", " ", text)

                # Unicode normalization — fixes invisible combining characters
                # that come from certain PDF font encodings and silently break
                # spreadsheet formulas / string comparisons.
                if uni_norm == "NFC (recommended)":
                    text = unicodedata.normalize("NFC", text)
                elif uni_norm == "NFKC (compatibility)":
                    text = unicodedata.normalize("NFKC", text)
                # else "None": leave as-is

                row.append(text)
            rows.append(row)

        # Second pass: vertical merge handling
        # pdfplumber returns None for cells that are vertically spanned —
        # those were already converted to "" or empty_marker above.
        # "Duplicate down rows" propagates the last non-empty value in each column.
        if merge_v_mode == "Duplicate down rows" and rows:
            n_cols = max(len(r) for r in rows)
            last_vals = [""] * n_cols
            for row in rows:
                for ci in range(len(row)):
                    if row[ci] == "" or row[ci] == empty_marker:
                        if last_vals[ci]:
                            row[ci] = last_vals[ci]
                    else:
                        last_vals[ci] = row[ci]

        # Third pass: type detection (opt-in)
        if type_detect != "Disabled":
            do_numbers = "Numbers" in type_detect
            do_dates   = "Dates"   in type_detect
            rows = [
                [self._convert_cell_type(c, do_numbers, do_dates) for c in row]
                for row in rows
            ]

        return rows

    # ── Type conversion helpers ────────────────────────────────────────────

    # Date formats attempted in order (most specific first)
    _DATE_FORMATS = [
        "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d",
        "%d-%m-%Y", "%m-%d-%Y",
        "%d.%m.%Y", "%m.%d.%Y",
        "%d %b %Y", "%d %B %Y",
        "%b %d, %Y", "%B %d, %Y",
        "%Y/%m/%d",
    ]

    @staticmethod
    def _try_parse_date(text: str) -> Optional[str]:
        """Try to parse *text* as a date; return ISO-8601 string or None."""
        t = text.strip()
        for fmt in PDFtoCSVTool._DATE_FORMATS:
            try:
                return datetime.datetime.strptime(t, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return None

    @staticmethod
    def _try_parse_number(text: str) -> Optional[str]:
        """
        Try to parse *text* as a number.
        Returns a canonical numeric string (e.g. "1234.56") or None.
        Guards against strings that look numeric but are identifiers:
        leading zeros (zip codes, IDs), pure integers ≤ 4 digits that could
        be year values in a date column, phone-number-like strings.
        """
        t = text.strip()
        if not t:
            return None

        # Reject strings with leading zeros (zip codes, account numbers, etc.)
        # Allow "-0.5", "0.5", but not "007" or "01234"
        if re.match(r"^0\d", t):
            return None

        # Strip common currency / percent symbols and thousands separators
        cleaned = re.sub(r"[£€$¥₹,%]", "", t)
        # Handle thousands separators: 1,234,567 or 1.234.567
        # Determine which separator is the decimal point heuristically:
        # if there's exactly one comma/dot and it's followed by ≤ 3 digits at end → decimal
        cleaned = cleaned.replace(" ", "")  # non-breaking/normal spaces
        # Remove thousands commas (e.g. 1,234,567 → 1234567; 1,234.56 → 1234.56)
        if re.match(r"^-?\d{1,3}(,\d{3})+(\.\d+)?$", cleaned):
            cleaned = cleaned.replace(",", "")
        # European style: 1.234,56 → 1234.56
        elif re.match(r"^-?\d{1,3}(\.\d{3})+(,\d+)?$", cleaned):
            cleaned = cleaned.replace(".", "").replace(",", ".")

        try:
            val = float(cleaned)
        except ValueError:
            return None

        # Return integer representation if no fractional part
        if val == int(val) and "." not in cleaned:
            return str(int(val))
        return str(val)

    @classmethod
    def _convert_cell_type(cls, text: str,
                            do_numbers: bool, do_dates: bool) -> str:
        """Convert a cell string to its canonical type representation if possible."""
        if not text.strip():
            return text
        if do_dates:
            parsed = cls._try_parse_date(text)
            if parsed is not None:
                return parsed
        if do_numbers:
            parsed = cls._try_parse_number(text)
            if parsed is not None:
                return parsed
        return text

    def _detect_header(self, rows: list[list[str]]) -> tuple[bool, list[str], list[list[str]]]:
        """Return (has_header, header_row, data_rows)."""
        mode = self._sv_header.get()
        if not rows:
            return False, [], []

        if mode == "No headers":
            return False, [], rows

        if mode == "First row is header":
            return True, rows[0], rows[1:]

        # Auto-detect
        first = rows[0]
        if not first:
            return False, [], rows

        score = 0
        if all(c != "" for c in first):
            score += 1
        if not any(re.match(r"^[\d.,%-]+$", c) for c in first if c):
            score += 1
        if len(set(c for c in first if c)) == len([c for c in first if c]):
            score += 1  # unique values

        if score >= 2:
            return True, first, rows[1:]
        return False, [], rows

    # ======================================================================
    # MINIMUM SIZE FILTER
    # ======================================================================

    def _passes_size_filter(self, rows: list[list[str]]) -> bool:
        """Return True if the table meets the minimum row/column thresholds."""
        try:
            min_rows = max(1, int(self._sv_min_rows.get()))
        except ValueError:
            min_rows = 1
        try:
            min_cols = max(1, int(self._sv_min_cols.get()))
        except ValueError:
            min_cols = 1

        if len(rows) < min_rows:
            return False
        if rows and max(len(r) for r in rows) < min_cols:
            return False
        return True

    # ======================================================================
    # CSV WRITER
    # ======================================================================

    def _write_csv(self, rows: list[list[str]], path: str) -> None:
        enc_name    = self._sv_encoding.get()
        encoding    = ENCODING_MAP.get(enc_name, "utf-8-sig")
        delim       = DELIMITER_MAP.get(self._sv_delimiter.get(), ",")
        line_ending = LINE_ENDING_MAP.get(self._sv_line_ending.get(), os.linesep)

        # Use io.open so we can control the line terminator precisely.
        # csv.writer's newline="" suppresses its own line endings; we add ours.
        with open(path, "w", newline="", encoding=encoding, errors="replace") as f:
            writer = csv.writer(f, delimiter=delim, quoting=csv.QUOTE_MINIMAL,
                                lineterminator=line_ending)
            writer.writerows(rows)

    # ======================================================================
    # SOURCE METADATA INJECTION
    # ======================================================================

    def _add_source_metadata(self, rows: list[list[str]],
                              page_num: int, table_num: int,
                              is_header: bool) -> list[list[str]]:
        """
        Prepend source column(s) to every row.
        page_num and table_num are 1-based for display.
        is_header=True means the first row is a header — give it a label col.
        """
        meta_mode = self._sv_source_meta.get()
        if meta_mode == "None":
            return rows

        result = []
        for i, row in enumerate(rows):
            is_hdr_row = (is_header and i == 0)
            if meta_mode == "Page number":
                prefix = ["Source page"] if is_hdr_row else [str(page_num)]
            elif meta_mode == "Table number":
                prefix = ["Source table"] if is_hdr_row else [str(table_num)]
            else:  # "Page + Table"
                prefix = (["Source page", "Source table"] if is_hdr_row
                           else [str(page_num), str(table_num)])
            result.append(prefix + list(row))
        return result

    # ======================================================================
    # COLUMN CONSISTENCY CHECK
    # ======================================================================

    @staticmethod
    def _check_column_consistency(rows: list[list[str]]) -> Optional[str]:
        """
        Return a warning string if the table has varying column counts,
        or None if all rows are consistent.
        """
        if not rows:
            return None
        col_counts = [len(r) for r in rows]
        unique_counts = set(col_counts)
        if len(unique_counts) <= 1:
            return None
        min_c = min(unique_counts)
        max_c = max(unique_counts)
        return (f"Inconsistent column count: rows range from {min_c} to {max_c} columns. "
                f"Some cells may be misaligned.")

    # ======================================================================
    # EXTRACTION PIPELINE
    # ======================================================================

    def _run_extraction(self):
        if not self._doc or not self._pldoc:
            messagebox.showwarning("No File", "Please open a PDF file first.")
            return
        if not self.output_dir:
            messagebox.showwarning("No Output Folder",
                                   "Please select an output folder.")
            return

        # Parse page range
        try:
            pages = self._parse_page_range(self._sv_range.get(),
                                           self._total_pages)
        except ValueError as e:
            messagebox.showerror("Invalid Page Range", str(e))
            return

        if not pages:
            messagebox.showwarning("Empty Selection", "No pages to process.")
            return

        base_name  = os.path.splitext(os.path.basename(self.pdf_path))[0]
        multi_mode = self._sv_multi.get()
        method     = self._sv_detection.get()
        settings   = self._build_table_settings(method)
        meta_mode  = self._sv_source_meta.get()
        image_only_policy = self._sv_image_only.get()

        # Disable button during extraction
        self._extract_btn.configure(state="disabled", text="Extracting…")
        self._progress.set(0)

        report_lines: list[str] = []
        report_lines.append("=== Extraction Complete ===\n")
        report_lines.append(f"Input:  {os.path.basename(self.pdf_path)}")
        report_lines.append(f"Output: {self.output_dir}")
        report_lines.append(f"Pages processed: {len(pages)}"
                             f"  (pages {pages[0]+1}–{pages[-1]+1})\n")

        all_table_rows: list[list[str]] = []    # for single-file mode
        total_tables   = 0
        total_rows     = 0
        skipped_files  = 0
        warnings: list[str] = []
        output_files: list[str] = []

        # Collect all tables first to compute progress
        page_table_data: list[tuple[int, int, list[list[str]], bool]] = []
        # Each entry: (page_idx, table_num_on_page, final_rows, has_header)

        for pg_idx in pages:
            self._status_lbl.configure(
                text=f"Detecting tables on page {pg_idx + 1}…")
            self.parent.update_idletasks()

            # Image-only page check
            if self._page_is_image_only(pg_idx):
                if image_only_policy == "Fail entirely":
                    messagebox.showerror(
                        "Image-Only Page",
                        f"Page {pg_idx+1} contains only scanned images and "
                        "cannot be extracted.\n\nChange 'Image-only pages' to "
                        "'Skip with warning' to continue past such pages."
                    )
                    self._extract_btn.configure(state="normal", text="Extract to CSV")
                    return
                else:
                    warnings.append(
                        f"Page {pg_idx+1}: image-only page — no text layer, skipped.")
                    continue

            try:
                pl_page = self._pldoc.pages[pg_idx]
                raw_tables = pl_page.extract_tables(settings)
            except Exception as e:
                warnings.append(f"Page {pg_idx+1}: extraction error — {e}")
                continue

            # Auto fallback: if lines strategy found nothing, try text
            if method == "Auto" and not raw_tables:
                try:
                    fallback = self._build_table_settings("Stream")
                    raw_tables = pl_page.extract_tables(fallback)
                    if raw_tables:
                        warnings.append(
                            f"Page {pg_idx+1}: lattice detection found no tables, "
                            "fell back to stream mode.")
                except Exception:
                    pass

            if not raw_tables:
                warnings.append(f"Page {pg_idx+1}: no tables detected (skipped).")
                continue

            for tbl_idx, raw in enumerate(raw_tables):
                rows = self._process_table(raw)
                has_hdr, hdr, data_rows = self._detect_header(rows)

                # Minimum size filter
                if not self._passes_size_filter(rows):
                    warnings.append(
                        f"Page {pg_idx+1} table {tbl_idx+1}: "
                        f"too small ({len(rows)} rows × "
                        f"{max(len(r) for r in rows) if rows else 0} cols), skipped.")
                    continue

                if has_hdr:
                    final_rows = [hdr] + data_rows
                else:
                    final_rows = rows

                # Add source metadata columns (for concatenated mode)
                if multi_mode == "Single file (concatenate)" and meta_mode != "None":
                    final_rows = self._add_source_metadata(
                        final_rows, pg_idx + 1, tbl_idx + 1, has_hdr)

                page_table_data.append((pg_idx, tbl_idx + 1, final_rows, has_hdr))

        n_total = len(page_table_data)

        for i, (pg_idx, tbl_num, final_rows, has_hdr) in enumerate(page_table_data):
            total_tables += 1
            n_rows = len(final_rows)
            total_rows += n_rows
            n_cols = max(len(r) for r in final_rows) if final_rows else 0

            self._progress.set((i + 1) / max(n_total, 1))
            self._status_lbl.configure(
                text=f"Writing table {i+1}/{n_total} "
                     f"(page {pg_idx+1}, table {tbl_num})…")
            self.parent.update_idletasks()

            # Column consistency check
            col_warn = self._check_column_consistency(final_rows)
            if col_warn:
                warnings.append(f"Page {pg_idx+1} table {tbl_num}: {col_warn}")

            if multi_mode == "Separate file per table":
                fname = f"{base_name}_page{pg_idx+1}_table{tbl_num}.csv"
                fpath_raw = os.path.join(self.output_dir, fname)
                fpath = self._resolve_output_path(fpath_raw)

                if fpath is None:
                    # Skipped due to overwrite policy
                    skipped_files += 1
                    warnings.append(
                        f"Page {pg_idx+1} table {tbl_num}: "
                        f"'{fname}' already exists — skipped.")
                    continue

                fname_actual = os.path.basename(fpath)
                try:
                    self._write_csv(final_rows, fpath)
                    output_files.append(fname_actual)
                except Exception as e:
                    warnings.append(
                        f"Page {pg_idx+1} table {tbl_num}: write error — {e}")
                    continue

                report_lines.append(f"Table {total_tables} — page {pg_idx+1}")
                report_lines.append(
                    f"  Dimensions: {n_rows} rows × {n_cols} columns")
                if fname_actual != fname:
                    report_lines.append(
                        f"  Output: {fname_actual}  (renamed — original existed)")
                else:
                    report_lines.append(f"  Output: {fname_actual}\n")

            else:  # single file — collect rows
                if all_table_rows and final_rows:
                    all_table_rows.append([])  # blank separator row
                all_table_rows.extend(final_rows)

                report_lines.append(f"Table {total_tables} — page {pg_idx+1}")
                report_lines.append(
                    f"  Dimensions: {n_rows} rows × {n_cols} columns\n")

        # Write combined file if single-file mode
        if multi_mode == "Single file (concatenate)" and all_table_rows:
            fname = f"{base_name}_all_tables.csv"
            fpath_raw = os.path.join(self.output_dir, fname)
            fpath = self._resolve_output_path(fpath_raw)

            if fpath is None:
                skipped_files += 1
                warnings.append(f"'{fname}' already exists — skipped (overwrite policy).")
            else:
                fname_actual = os.path.basename(fpath)
                try:
                    self._write_csv(all_table_rows, fpath)
                    output_files.append(fname_actual)
                    if fname_actual != fname:
                        report_lines.append(
                            f"\nOutput: {fname_actual}  (renamed — original existed)")
                    else:
                        report_lines.append(f"\nOutput: {fname_actual}")
                except Exception as e:
                    warnings.append(f"Write error for combined file: {e}")

        # Summary
        report_lines.append(f"\n── Summary ──────────────────────")
        report_lines.append(f"Tables found:    {total_tables}")
        report_lines.append(f"Total rows:      {total_rows}")
        report_lines.append(f"Output files:    {len(output_files)}")
        if skipped_files:
            report_lines.append(f"Files skipped:   {skipped_files} (already existed)")

        if warnings:
            report_lines.append(f"\n── Warnings ─────────────────────")
            for w in warnings:
                report_lines.append(f"  • {w}")

        if not output_files:
            report_lines.append("\n⚠ No CSV files were created.")
            report_lines.append("  The PDF may not contain extractable tables.")
            report_lines.append("  Try changing the Detection Method (Stream or Hybrid).")

        self._extract_btn.configure(state="normal", text="Extract to CSV")
        self._progress.set(1)
        self._status_lbl.configure(
            text=f"Done. {total_tables} table{'s' if total_tables != 1 else ''} "
                 f"extracted to {len(output_files)} file{'s' if len(output_files) != 1 else ''}.")

        self._show_report("\n".join(report_lines), self.output_dir)

    # ======================================================================
    # REPORT PANEL
    # ======================================================================

    def _show_report(self, text: str, output_dir: str):
        """Replace the page preview canvas with a scrollable report."""
        # Destroy old report frame if any
        if self._report_frame and self._report_frame.winfo_exists():
            self._report_frame.destroy()

        # Overlay a new frame on top of the right panel grid slot 0,0
        report = ctk.CTkFrame(self._right, fg_color=WHITE, corner_radius=0)
        report.grid(row=0, column=0, sticky="nsew")
        self._report_frame = report

        inner = ctk.CTkScrollableFrame(report, fg_color=WHITE,
                                        corner_radius=0,
                                        scrollbar_button_color=G300,
                                        scrollbar_button_hover_color=G400)
        inner.pack(fill="both", expand=True, padx=20, pady=16)

        ctk.CTkLabel(inner, text=text,
                     font=ctk.CTkFont(family="Courier New", size=12),
                     text_color=G700, anchor="nw", justify="left",
                     wraplength=700).pack(anchor="nw", fill="x")

        # Open folder button
        btn_row = ctk.CTkFrame(report, fg_color=G100, corner_radius=0, height=50)
        btn_row.pack(fill="x", side="bottom")
        btn_row.pack_propagate(False)

        ctk.CTkButton(
            btn_row, text="Open Output Folder", width=180, height=34,
            fg_color=BLUE, hover_color=BLUE_HOVER,
            command=lambda: self._open_folder(output_dir),
        ).pack(side="left", padx=16, pady=8)

        ctk.CTkButton(
            btn_row, text="← Back to Preview", width=160, height=34,
            fg_color="transparent", hover_color=G200,
            text_color=G700, border_color=G300, border_width=1,
            command=self._back_to_preview,
        ).pack(side="left", padx=(0, 8), pady=8)

    def _back_to_preview(self):
        if self._report_frame and self._report_frame.winfo_exists():
            self._report_frame.destroy()
            self._report_frame = None
        if self._doc:
            self._render_page_canvas()

    def _open_folder(self, path: str):
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", path])
            elif sys.platform == "win32":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{e}")
