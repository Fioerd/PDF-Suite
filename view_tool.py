"""View Tool – Full-featured PDF viewer with annotations, search, forms.

This module is loaded by main.py when the user clicks "View PDF".

Enhanced with:
  - Undo/Redo system (Ctrl+Z / Ctrl+Y)
  - Keyboard shortcuts (V=View, S=Select, H=Highlight, etc.)
  - Shift-key constraints (aspect ratio lock, H/V snap)
  - Smooth Bezier signature paths (Catmull-Rom interpolation)
  - Enhanced sidebar visibility (higher contrast, borders)
  - Vibrant semi-transparent text selection with negative padding
"""

import io
import math
import os
import subprocess
import tempfile
import tkinter as tk
from enum import Enum
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog

import customtkinter as ctk

try:
    import fitz  # pymupdf
except ImportError:
    fitz = None

try:
    from PIL import Image, ImageTk, ImageDraw
except ImportError:
    Image = ImageTk = ImageDraw = None

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------
BLUE       = "#3B82F6"
BLUE_HOVER = "#2563EB"
GREEN      = "#16A34A"
ORANGE     = "#F97316"
YELLOW_HL  = "#FDE047"
RED        = "#EF4444"
G50        = "#F9FAFB"
G100       = "#F3F4F6"
G200       = "#E5E7EB"
G300       = "#D1D5DB"
G400       = "#9CA3AF"
G500       = "#6B7280"
G600       = "#4B5563"
G700       = "#374151"
G800       = "#1F2937"
G900       = "#111827"
WHITE      = "#FFFFFF"
THUMB_BG   = "#F0F2F5"
TOOL_ACTIVE = "#DBEAFE"
TOOL_BORDER = "#93C5FD"

# Enhanced sidebar colors
SIDEBAR_BG     = "#E2E6EC"
SIDEBAR_BORDER = "#C8CDD5"
TAB_BG         = "#EDF0F4"

# Selection highlight (vibrant royal blue)
SEL_BLUE_R, SEL_BLUE_G, SEL_BLUE_B = 59, 100, 246
SEL_ALPHA = 80  # ~31% opacity

# Annotation preset colors  (display_name, hex, fitz_rgb)
ANNOT_COLORS = [
    ("Yellow",  "#FBBF24", (1.0, 0.75, 0.14)),
    ("Red",     "#EF4444", (0.94, 0.27, 0.27)),
    ("Blue",    "#3B82F6", (0.23, 0.51, 0.96)),
    ("Green",   "#22C55E", (0.13, 0.77, 0.37)),
    ("Orange",  "#F97316", (0.98, 0.45, 0.09)),
    ("Black",   "#111827", (0.07, 0.09, 0.15)),
]


# ---------------------------------------------------------------------------
# Tool Enum
# ---------------------------------------------------------------------------
class Tool(Enum):
    VIEW          = "view"
    SELECT        = "select"
    HIGHLIGHT     = "highlight"
    UNDERLINE     = "underline"
    STRIKETHROUGH = "strikethrough"
    FREEHAND      = "freehand"
    TEXT_BOX      = "textbox"
    STICKY_NOTE   = "note"
    RECT          = "rect"
    CIRCLE        = "circle"
    LINE          = "line"
    ARROW         = "arrow"
    SIGN          = "sign"


# (tool, icon, label, shortcut_key)
TOOL_DEFS = [
    (Tool.SELECT,        "\u270f",  "Select",    "S"),
    (Tool.HIGHLIGHT,     "\U0001f58d", "Highlight", "H"),
    (Tool.UNDERLINE,     "_",  "Underline", "U"),
    (Tool.STRIKETHROUGH, "\u2014",  "Strikeout", "K"),
    (Tool.TEXT_BOX,      "T",  "Text Box",  "X"),
    (Tool.STICKY_NOTE,   "\U0001f4ac", "Note",      "N"),
    (Tool.FREEHAND,      "\u270d",  "Freehand",  "D"),
    (Tool.RECT,          "\u25a1",  "Rectangle", "R"),
    (Tool.CIRCLE,        "\u25cb",  "Circle",    "O"),
    (Tool.LINE,          "/",  "Line",      "L"),
    (Tool.ARROW,         "\u2192",  "Arrow",     "A"),
    (Tool.SIGN,          "\u2712",  "Sign",      ""),
]

FIT_PAGE  = -1.0
FIT_WIDTH = -2.0
MAX_UNDO  = 30


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _pix_to_tk(pixmap):
    if Image and ImageTk:
        img = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        return ImageTk.PhotoImage(img)
    return tk.PhotoImage(data=pixmap.tobytes("ppm"))


def _render_thumb(doc, idx, max_w):
    page = doc[idx]
    s = max_w / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
    return _pix_to_tk(pix)


def _catmull_rom_segment(p0, p1, p2, p3, num_pts=6):
    """Catmull-Rom spline interpolation between p1 and p2."""
    result = []
    for i in range(num_pts):
        t = i / num_pts
        t2 = t * t
        t3 = t2 * t
        x = 0.5 * ((2 * p1[0]) +
            (-p0[0] + p2[0]) * t +
            (2 * p0[0] - 5 * p1[0] + 4 * p2[0] - p3[0]) * t2 +
            (-p0[0] + 3 * p1[0] - 3 * p2[0] + p3[0]) * t3)
        y = 0.5 * ((2 * p1[1]) +
            (-p0[1] + p2[1]) * t +
            (2 * p0[1] - 5 * p1[1] + 4 * p2[1] - p3[1]) * t2 +
            (-p0[1] + 3 * p1[1] - 3 * p2[1] + p3[1]) * t3)
        result.append((x, y))
    return result


def _smooth_stroke(points, num_interp=6):
    """Smooth a list of (x, y) points using Catmull-Rom spline."""
    if len(points) < 3:
        return list(points)
    result = []
    pts = list(points)
    for i in range(len(pts) - 1):
        p0 = pts[max(i - 1, 0)]
        p1 = pts[i]
        p2 = pts[min(i + 1, len(pts) - 1)]
        p3 = pts[min(i + 2, len(pts) - 1)]
        result.extend(_catmull_rom_segment(p0, p1, p2, p3, num_interp))
    result.append(pts[-1])
    return result


def _make_sel_overlay(width, height):
    """Create a semi-transparent royal blue PIL image for text selection."""
    if not Image or not ImageTk:
        return None
    w = max(1, int(width))
    h = max(1, int(height))
    img = Image.new("RGBA", (w, h), (SEL_BLUE_R, SEL_BLUE_G, SEL_BLUE_B, SEL_ALPHA))
    return ImageTk.PhotoImage(img)


# ===========================================================================
# View Tool
# ===========================================================================

class ViewTool:
    THUMB_W    = 80
    ZOOM_STEP  = 0.25
    ZOOM_MIN   = 0.25
    ZOOM_MAX   = 5.0
    SIDEBAR_W  = 190

    def __init__(self, parent: ctk.CTkFrame):
        self.parent = parent

        if fitz is None:
            ctk.CTkLabel(parent,
                text="\u26a0  Missing dependencies.\n\npip install pymupdf Pillow",
                font=ctk.CTkFont(size=16), text_color=G500).pack(expand=True)
            return

        # -- Document state --
        self.pdf_path     = ""
        self.doc          = None
        self.total_pages  = 0
        self.current_page = 0
        self._modified    = False

        # -- View state --
        self.zoom         = FIT_PAGE
        self._rotation    = 0
        self._preview_img = None
        self._thumb_imgs: list = []
        self._highlighted_thumb_cv = None
        self._thumb_render_next = 0

        # -- Coordinate mapping (set during render) --
        self._page_ox     = 0.0
        self._page_oy     = 0.0
        self._render_mat  = fitz.Matrix(1, 1)
        self._inv_mat     = fitz.Matrix(1, 1)

        # -- Tool state --
        self._tool        = Tool.VIEW
        self._annot_color_idx = 0
        self._stroke_width = 2
        self._drag_start  = None
        self._freehand_pts: list[tuple] = []
        self._selected_words: list = []
        self._selection_text = ""
        self._shift_held  = False
        self._custom_color = None   # (name, hex, fitz_rgb) or None

        # -- Selection overlay images (keep references so GC doesn't eat them) --
        self._sel_images: list = []

        # -- Search state --
        self._search_results: list[tuple] = []
        self._search_flat: list[tuple]    = []
        self._search_idx   = -1
        self._search_visible = False

        # -- Form widget windows --
        self._form_windows: list = []

        # -- Undo / Redo --
        self._undo_stack: list[tuple] = []   # [(doc_bytes, page_idx), ...]
        self._redo_stack: list[tuple] = []

        self._build_ui()

    # ==================================================================
    # BUILD UI
    # ==================================================================

    def _build_ui(self):
        top = self.parent.winfo_toplevel()

        # -- Toolbar -------------------------------------------------------
        tb = ctk.CTkFrame(self.parent, fg_color=G100, corner_radius=0, height=52)
        tb.pack(fill="x")
        tb.pack_propagate(False)

        ctk.CTkButton(tb, text="\U0001f4c1 Open", width=90, height=34,
            fg_color=BLUE, hover_color=BLUE_HOVER,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._pick_pdf).pack(side="left", padx=(12, 4), pady=8)
        ctk.CTkButton(tb, text="\U0001f4be Save", width=80, height=34,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            command=self._save_pdf).pack(side="left", padx=4, pady=8)
        ctk.CTkButton(tb, text="\U0001f5a8 Print", width=80, height=34,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            command=self._print_pdf).pack(side="left", padx=4, pady=8)
        ctk.CTkButton(tb, text="+ Add PDF", width=90, height=34,
            fg_color=GREEN, hover_color="#15803D",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._add_pdf).pack(side="left", padx=4, pady=8)

        # Undo / Redo buttons
        ctk.CTkButton(tb, text="\u21a9", width=34, height=34,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            font=ctk.CTkFont(size=16),
            command=self._undo).pack(side="left", padx=(12, 2), pady=8)
        ctk.CTkButton(tb, text="\u21aa", width=34, height=34,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            font=ctk.CTkFont(size=16),
            command=self._redo).pack(side="left", padx=2, pady=8)

        self.file_lbl = ctk.CTkLabel(tb, text="No file loaded",
            font=ctk.CTkFont(size=12), text_color=G500, anchor="w")
        self.file_lbl.pack(side="left", padx=(12, 0))

        # Right side: zoom + rotate
        zf = ctk.CTkFrame(tb, fg_color="transparent")
        zf.pack(side="right", padx=12, pady=8)

        ctk.CTkButton(zf, text="\u21bb", width=34, height=32,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            font=ctk.CTkFont(size=16), command=self._rotate_view
        ).pack(side="left", padx=(0, 10))

        self.btn_fit = ctk.CTkButton(zf, text="Fit", width=42, height=32,
            fg_color=BLUE, hover_color=BLUE_HOVER,
            font=ctk.CTkFont(size=11), command=self._zoom_fit)
        self.btn_fit.pack(side="left", padx=(0, 2))

        self.btn_fitw = ctk.CTkButton(zf, text="W", width=34, height=32,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            font=ctk.CTkFont(size=11, weight="bold"), command=self._zoom_fit_width)
        self.btn_fitw.pack(side="left", padx=(0, 8))

        ctk.CTkButton(zf, text="\u2212", width=30, height=32,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._zoom_out).pack(side="left", padx=(0, 2))

        self.zoom_lbl = ctk.CTkLabel(zf, text="Fit", width=50,
            font=ctk.CTkFont(size=11), text_color=G700)
        self.zoom_lbl.pack(side="left", padx=2)

        ctk.CTkButton(zf, text="+", width=30, height=32,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._zoom_in).pack(side="left")

        # -- Bottom area: search bar + nav ---------------------------------
        bottom = ctk.CTkFrame(self.parent, fg_color="transparent", corner_radius=0)
        bottom.pack(side="bottom", fill="x")

        # Search bar (hidden initially)
        self._search_frame = ctk.CTkFrame(bottom, fg_color=G100,
            corner_radius=0, height=42)
        self._search_frame.pack_propagate(False)

        sf_inner = ctk.CTkFrame(self._search_frame, fg_color="transparent")
        sf_inner.pack(expand=True)

        ctk.CTkLabel(sf_inner, text="\U0001f50d", font=ctk.CTkFont(size=14),
            text_color=G500).pack(side="left", padx=(0, 4))

        self._search_entry = ctk.CTkEntry(sf_inner, width=250, height=30,
            placeholder_text="Search\u2026", fg_color=WHITE, border_color=G300,
            text_color=G900, font=ctk.CTkFont(size=12))
        self._search_entry.pack(side="left", padx=4)
        self._search_entry.bind("<Return>", lambda e: self._do_search())

        ctk.CTkButton(sf_inner, text="\u25c0", width=30, height=28,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            command=self._search_prev).pack(side="left", padx=2)
        ctk.CTkButton(sf_inner, text="\u25b6", width=30, height=28,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            command=self._search_next).pack(side="left", padx=2)

        self._search_count_lbl = ctk.CTkLabel(sf_inner, text="",
            font=ctk.CTkFont(size=11), text_color=G500, width=80)
        self._search_count_lbl.pack(side="left", padx=4)

        ctk.CTkButton(sf_inner, text="\u2715", width=28, height=28,
            fg_color="transparent", hover_color=G200, text_color=G500,
            command=self._hide_search).pack(side="left", padx=2)

        # Navigation bar
        nav = ctk.CTkFrame(bottom, fg_color=G100, corner_radius=0, height=44)
        nav.pack(fill="x")
        nav.pack_propagate(False)

        nav_inner = ctk.CTkFrame(nav, fg_color="transparent")
        nav_inner.pack(expand=True)

        self.btn_first = ctk.CTkButton(nav_inner, text="\u23ee", width=34, height=30,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            command=self._first_page, state="disabled")
        self.btn_first.pack(side="left", padx=2)

        self.btn_prev = ctk.CTkButton(nav_inner, text="\u2190", width=34, height=30,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            command=self._prev_page, state="disabled")
        self.btn_prev.pack(side="left", padx=2)

        self.page_entry = ctk.CTkEntry(nav_inner, width=50, height=30,
            justify="center", fg_color=WHITE, border_color=G300,
            text_color=G900, font=ctk.CTkFont(size=12))
        self.page_entry.pack(side="left", padx=4)
        self.page_entry.insert(0, "\u2013")
        self.page_entry.bind("<Return>", self._goto_page)

        self.total_lbl = ctk.CTkLabel(nav_inner, text="/ \u2013",
            font=ctk.CTkFont(size=12), text_color=G500)
        self.total_lbl.pack(side="left", padx=(0, 4))

        self.btn_next = ctk.CTkButton(nav_inner, text="\u2192", width=34, height=30,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            command=self._next_page, state="disabled")
        self.btn_next.pack(side="left", padx=2)

        self.btn_last = ctk.CTkButton(nav_inner, text="\u23ed", width=34, height=30,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            command=self._last_page, state="disabled")
        self.btn_last.pack(side="left", padx=2)

        # -- Body: sidebar + canvas ----------------------------------------
        body = ctk.CTkFrame(self.parent, fg_color="transparent")
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # ENHANCED SIDEBAR (higher contrast, clear border)
        sidebar = ctk.CTkFrame(body, fg_color=SIDEBAR_BG, width=self.SIDEBAR_W,
            corner_radius=0, border_width=2, border_color=SIDEBAR_BORDER)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        self._tabs = ctk.CTkTabview(sidebar, width=self.SIDEBAR_W - 12,
            height=100, fg_color=TAB_BG,
            segmented_button_fg_color=SIDEBAR_BG,
            segmented_button_selected_color=BLUE,
            segmented_button_selected_hover_color=BLUE_HOVER,
            segmented_button_unselected_color=G300,
            segmented_button_unselected_hover_color=G400,
            corner_radius=6)
        self._tabs.pack(fill="both", expand=True, padx=5, pady=5)

        tab_pages = self._tabs.add("Pages")
        tab_toc   = self._tabs.add("TOC")
        tab_tools = self._tabs.add("Tools")

        # -- Pages tab
        self.thumb_scroll = ctk.CTkScrollableFrame(tab_pages,
            fg_color=WHITE, scrollbar_button_color=G300,
            scrollbar_button_hover_color=G400, corner_radius=4)
        self.thumb_scroll.pack(fill="both", expand=True)

        # -- TOC tab
        self.toc_scroll = ctk.CTkScrollableFrame(tab_toc,
            fg_color=WHITE, scrollbar_button_color=G300,
            scrollbar_button_hover_color=G400, corner_radius=4)
        self.toc_scroll.pack(fill="both", expand=True)
        ctk.CTkLabel(self.toc_scroll, text="Open a PDF to\nsee table of contents",
            text_color=G400, font=ctk.CTkFont(size=11)).pack(pady=20)

        # -- Tools tab
        self._build_tools_tab(tab_tools)

        # Canvas area
        pf = ctk.CTkFrame(body, fg_color=WHITE, corner_radius=0)
        pf.grid(row=0, column=1, sticky="nsew")
        pf.rowconfigure(0, weight=1)
        pf.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(pf, bg=G50, highlightthickness=0, relief="flat")
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self._vbar = tk.Scrollbar(pf, orient="vertical", command=self.canvas.yview)
        self._vbar.grid(row=0, column=1, sticky="ns")
        self._hbar = tk.Scrollbar(pf, orient="horizontal", command=self.canvas.xview)
        self._hbar.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(xscrollcommand=self._hbar.set,
                               yscrollcommand=self._vbar.set)

        self.canvas.create_text(300, 250, text="Open a PDF to start viewing",
            font=("", 18), fill=G400, tags="ph")

        # -- Bindings ------------------------------------------------------
        self.canvas.bind("<ButtonPress-1>",   self._on_mouse_down)
        self.canvas.bind("<B1-Motion>",       self._on_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self._on_mouse_up)
        self.canvas.bind("<Double-Button-1>", self._on_double_click)
        self.canvas.bind("<MouseWheel>",      self._on_scroll)
        self.canvas.bind("<Configure>",       self._on_canvas_resize)

        # Shift key tracking
        top.bind("<KeyPress-Shift_L>",   lambda e: self._set_shift(True))
        top.bind("<KeyPress-Shift_R>",   lambda e: self._set_shift(True))
        top.bind("<KeyRelease-Shift_L>", lambda e: self._set_shift(False))
        top.bind("<KeyRelease-Shift_R>", lambda e: self._set_shift(False))

        # Navigation & zoom
        top.bind("<Left>",          lambda e: self._prev_page())
        top.bind("<Right>",         lambda e: self._next_page())
        top.bind("<Home>",          lambda e: self._first_page())
        top.bind("<End>",           lambda e: self._last_page())
        top.bind("<plus>",          lambda e: self._zoom_in())
        top.bind("<minus>",         lambda e: self._zoom_out())
        top.bind("<KP_Add>",        lambda e: self._zoom_in())
        top.bind("<KP_Subtract>",   lambda e: self._zoom_out())

        # Ctrl shortcuts
        top.bind("<Control-f>",     lambda e: self._toggle_search())
        top.bind("<Control-F>",     lambda e: self._toggle_search())
        top.bind("<Control-c>",     lambda e: self._copy_selection())
        top.bind("<Control-C>",     lambda e: self._copy_selection())
        top.bind("<Control-s>",     lambda e: self._save_pdf())
        top.bind("<Control-S>",     lambda e: self._save_pdf())
        top.bind("<Control-z>",     lambda e: self._undo())
        top.bind("<Control-Z>",     lambda e: self._undo())
        top.bind("<Control-y>",     lambda e: self._redo())
        top.bind("<Control-Y>",     lambda e: self._redo())
        top.bind("<Escape>",        lambda e: self._escape())

        # Single-key tool shortcuts
        top.bind("<Key>", self._on_key_shortcut)

    # -- Shift state -------------------------------------------------------
    def _set_shift(self, held: bool):
        self._shift_held = held

    # -- Single-key shortcuts ----------------------------------------------
    def _on_key_shortcut(self, event):
        """Handle single-key shortcuts for tool switching."""
        # Don't trigger when typing in entries
        focus = self.parent.winfo_toplevel().focus_get()
        if isinstance(focus, (tk.Entry, ctk.CTkEntry)):
            return
        if event.state & 0x4:  # Ctrl held, skip
            return

        key = event.char.upper() if event.char else ""
        shortcut_map = {
            "V": Tool.VIEW,
            "S": Tool.SELECT,
            "H": Tool.HIGHLIGHT,
            "U": Tool.UNDERLINE,
            "K": Tool.STRIKETHROUGH,
            "X": Tool.TEXT_BOX,
            "N": Tool.STICKY_NOTE,
            "D": Tool.FREEHAND,
            "R": Tool.RECT,
            "O": Tool.CIRCLE,
            "L": Tool.LINE,
            "A": Tool.ARROW,
        }
        if key in shortcut_map:
            self._set_tool(shortcut_map[key])
        elif key == "T":
            self._tabs.set("Tools")
        elif key == "P":
            self._tabs.set("Pages")

    # -- Tools tab (enhanced visibility) -----------------------------------
    def _build_tools_tab(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color=WHITE,
            scrollbar_button_color=G300, scrollbar_button_hover_color=G400,
            corner_radius=4)
        scroll.pack(fill="both", expand=True)

        # Section header
        hdr = ctk.CTkFrame(scroll, fg_color=G100, corner_radius=6)
        hdr.pack(fill="x", padx=2, pady=(4, 8))
        ctk.CTkLabel(hdr, text="\U0001f527 Tools",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=G700).pack(anchor="w", padx=8, pady=4)

        self._tool_buttons: dict[Tool, ctk.CTkButton] = {}

        # View tool (always first)
        btn = ctk.CTkButton(scroll, text="\U0001f446  View  (V)", height=32, anchor="w",
            fg_color=TOOL_ACTIVE, hover_color=G200, text_color=G700,
            border_color=TOOL_BORDER, border_width=1,
            font=ctk.CTkFont(size=11), corner_radius=6,
            command=lambda: self._set_tool(Tool.VIEW))
        btn.pack(fill="x", padx=2, pady=1)
        self._tool_buttons[Tool.VIEW] = btn

        for tool, icon, label, shortcut in TOOL_DEFS:
            hint = f"  ({shortcut})" if shortcut else ""
            btn = ctk.CTkButton(scroll, text=f"{icon}  {label}{hint}",
                height=32, anchor="w",
                fg_color="transparent", hover_color=G200, text_color=G700,
                font=ctk.CTkFont(size=11), corner_radius=6,
                command=lambda t=tool: self._set_tool(t))
            btn.pack(fill="x", padx=2, pady=1)
            self._tool_buttons[tool] = btn

        # Separator
        ctk.CTkFrame(scroll, fg_color=G300, height=1).pack(fill="x", pady=8, padx=4)

        # Color header
        chdr = ctk.CTkFrame(scroll, fg_color=G100, corner_radius=6)
        chdr.pack(fill="x", padx=2, pady=(0, 4))
        ctk.CTkLabel(chdr, text="\U0001f3a8 Color",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=G700).pack(anchor="w", padx=8, pady=4)

        color_row = ctk.CTkFrame(scroll, fg_color="transparent")
        color_row.pack(fill="x", padx=4, pady=(0, 8))

        self._color_indicators: list[tk.Canvas] = []
        for i, (name, hex_c, _) in enumerate(ANNOT_COLORS):
            cv = tk.Canvas(color_row, width=28, height=28,
                bg=WHITE, highlightthickness=2,
                highlightbackground=BLUE if i == 0 else G200,
                cursor="hand2")
            cv.pack(side="left", padx=2)
            cv.create_oval(4, 4, 24, 24, fill=hex_c, outline="")
            cv.bind("<Button-1>", lambda e, idx=i: self._set_color(idx))
            self._color_indicators.append(cv)

        # Custom hex color entry
        hex_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        hex_frame.pack(fill="x", padx=4, pady=(4, 8))
        ctk.CTkLabel(hex_frame, text="#",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=G600, width=14).pack(side="left")
        self._hex_entry = ctk.CTkEntry(hex_frame, width=80, height=28,
            placeholder_text="FF5733", fg_color=WHITE,
            border_color=G300, text_color=G900,
            font=ctk.CTkFont(size=11, family="Consolas"))
        self._hex_entry.pack(side="left", padx=2)
        self._hex_entry.bind("<Return>", lambda e: self._apply_hex_color())
        self._hex_preview = tk.Canvas(hex_frame, width=28, height=28,
            bg=WHITE, highlightthickness=2,
            highlightbackground=G200)
        self._hex_preview.pack(side="left", padx=4)
        self._hex_preview.create_oval(4, 4, 24, 24, fill=G300,
            outline="", tags="prev")
        ctk.CTkButton(hex_frame, text="Set", width=36, height=28,
            fg_color=BLUE, hover_color=BLUE_HOVER,
            font=ctk.CTkFont(size=10),
            command=self._apply_hex_color).pack(side="left", padx=2)

        # Width header
        whdr = ctk.CTkFrame(scroll, fg_color=G100, corner_radius=6)
        whdr.pack(fill="x", padx=2, pady=(0, 4))
        ctk.CTkLabel(whdr, text="\U0001f4cf Width",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=G700).pack(anchor="w", padx=8, pady=4)

        self._width_slider = ctk.CTkSlider(scroll, from_=1, to=10,
            number_of_steps=9, width=140, height=18,
            fg_color=G200, progress_color=BLUE, button_color=BLUE,
            button_hover_color=BLUE_HOVER)
        self._width_slider.set(2)
        self._width_slider.pack(padx=6, pady=(0, 4), anchor="w")
        self._width_slider.configure(command=self._on_width_change)

        self._width_lbl = ctk.CTkLabel(scroll, text="2 px",
            font=ctk.CTkFont(size=10), text_color=G500)
        self._width_lbl.pack(anchor="w", padx=6)

    def _set_tool(self, tool: Tool):
        self._tool = tool
        for t, btn in self._tool_buttons.items():
            if t == tool:
                btn.configure(fg_color=TOOL_ACTIVE, border_color=TOOL_BORDER,
                              border_width=1)
            else:
                btn.configure(fg_color="transparent", border_color=G200,
                              border_width=0)
        cursors = {
            Tool.VIEW: "", Tool.SELECT: "xterm",
            Tool.HIGHLIGHT: "xterm", Tool.UNDERLINE: "xterm",
            Tool.STRIKETHROUGH: "xterm", Tool.FREEHAND: "pencil",
            Tool.TEXT_BOX: "crosshair", Tool.STICKY_NOTE: "crosshair",
            Tool.RECT: "crosshair", Tool.CIRCLE: "crosshair",
            Tool.LINE: "crosshair", Tool.ARROW: "crosshair",
            Tool.SIGN: "crosshair",
        }
        self.canvas.configure(cursor=cursors.get(tool, ""))

    def _set_color(self, idx):
        self._annot_color_idx = idx
        self._custom_color = None
        for i, cv in enumerate(self._color_indicators):
            cv.configure(highlightbackground=BLUE if i == idx else G200)
        # Reset hex preview
        if hasattr(self, '_hex_preview'):
            self._hex_preview.delete("prev")
            self._hex_preview.create_oval(4, 4, 24, 24, fill=G300,
                outline="", tags="prev")

    def _apply_hex_color(self):
        """Apply a custom hex color from the entry field."""
        raw = self._hex_entry.get().strip().lstrip('#')
        if len(raw) not in (3, 6):
            return
        if len(raw) == 3:
            raw = raw[0]*2 + raw[1]*2 + raw[2]*2
        try:
            r = int(raw[0:2], 16)
            g = int(raw[2:4], 16)
            b = int(raw[4:6], 16)
        except ValueError:
            return
        hex_str = f"#{raw.upper()}"
        fitz_rgb = (r / 255.0, g / 255.0, b / 255.0)
        self._custom_color = ("Custom", hex_str, fitz_rgb)
        self._annot_color_idx = -1
        # Deselect all preset indicators
        for cv in self._color_indicators:
            cv.configure(highlightbackground=G200)
        # Update preview circle
        self._hex_preview.delete("prev")
        self._hex_preview.create_oval(4, 4, 24, 24, fill=hex_str,
            outline="", tags="prev")
        self._hex_preview.configure(highlightbackground=BLUE)

    def _on_width_change(self, val):
        w = int(float(val))
        self._stroke_width = w
        self._width_lbl.configure(text=f"{w} px")

    @property
    def _annot_color(self):
        if self._custom_color:
            return self._custom_color
        return ANNOT_COLORS[self._annot_color_idx]

    # ==================================================================
    # UNDO / REDO
    # ==================================================================

    def _push_undo(self):
        """Save current document state before a modification."""
        if not self.doc:
            return
        try:
            buf = self.doc.tobytes()
            self._undo_stack.append((buf, self.current_page))
            if len(self._undo_stack) > MAX_UNDO:
                self._undo_stack.pop(0)
            self._redo_stack.clear()
        except Exception:
            pass

    def _undo(self):
        """Restore the previous document state."""
        if not self._undo_stack or not self.doc:
            return
        try:
            cur_buf = self.doc.tobytes()
            self._redo_stack.append((cur_buf, self.current_page))
            buf, page_idx = self._undo_stack.pop()
            self.doc.close()
            self.doc = fitz.open(stream=buf, filetype="pdf")
            self.total_pages = len(self.doc)
            self._modified = bool(self._undo_stack)
            self._show_page(min(page_idx, self.total_pages - 1))
        except Exception as e:
            messagebox.showerror("Undo Error", str(e))

    def _redo(self):
        """Re-apply the last undone action."""
        if not self._redo_stack or not self.doc:
            return
        try:
            cur_buf = self.doc.tobytes()
            self._undo_stack.append((cur_buf, self.current_page))
            buf, page_idx = self._redo_stack.pop()
            self.doc.close()
            self.doc = fitz.open(stream=buf, filetype="pdf")
            self.total_pages = len(self.doc)
            self._modified = True
            self._show_page(min(page_idx, self.total_pages - 1))
        except Exception as e:
            messagebox.showerror("Redo Error", str(e))

    # ==================================================================
    # FILE LOADING
    # ==================================================================

    def _pick_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not p:
            return
        self.pdf_path = p
        self._load_pdf()

    def _load_pdf(self):
        try:
            if self.doc:
                self.doc.close()
            self.doc = fitz.open(self.pdf_path)
            self.total_pages = len(self.doc)
            self.current_page = 0
            self.zoom = FIT_PAGE
            self._rotation = 0
            self._modified = False
            self._search_results.clear()
            self._search_flat.clear()
            self._search_idx = -1
            self._undo_stack.clear()
            self._redo_stack.clear()

            name = Path(self.pdf_path).name
            self.file_lbl.configure(text=name)
            self.total_lbl.configure(text=f"/ {self.total_pages}")
            self._update_zoom_label()

            self._render_thumbnails()
            self._build_toc()
            self._show_page(0)
        except Exception as e:
            self.total_pages = 0
            messagebox.showerror("Error", f"Could not load PDF:\n{e}")

    def _add_pdf(self):
        """Append another PDF's pages after the current document."""
        if not self.doc:
            # No document loaded yet – just do a normal open
            self._pick_pdf()
            return
        p = filedialog.askopenfilename(
            title="Add PDF",
            filetypes=[("PDF", "*.pdf")])
        if not p:
            return
        try:
            self._push_undo()
            src = fitz.open(p)
            self.doc.insert_pdf(src)
            src.close()
            self.total_pages = len(self.doc)
            self._modified = True

            name = Path(self.pdf_path).name
            self.file_lbl.configure(
                text=f"{name} (+{Path(p).name})")
            self.total_lbl.configure(text=f"/ {self.total_pages}")

            self._render_thumbnails()
            self._build_toc()
            self._show_page(self.current_page)
        except Exception as e:
            messagebox.showerror("Error", f"Could not add PDF:\n{e}")

    # ==================================================================
    # THUMBNAILS
    # ==================================================================

    def _render_thumbnails(self):
        for w in self.thumb_scroll.winfo_children():
            w.destroy()
        self._thumb_imgs.clear()
        self._highlighted_thumb_cv = None
        if not self.doc:
            return
        # Create placeholder frames immediately (instant), then fill lazily
        for i in range(self.total_pages):
            frame = ctk.CTkFrame(self.thumb_scroll, fg_color="transparent",
                cursor="hand2")
            frame.pack(fill="x", pady=2, padx=2)
            # Placeholder canvas at fixed size while image loads
            placeholder_h = int(self.THUMB_W * 1.4)
            cv = tk.Canvas(frame, width=self.THUMB_W, height=placeholder_h,
                bg=G200, highlightthickness=2, highlightbackground=G300)
            cv.pack(padx=4, pady=(2, 0))
            lbl = ctk.CTkLabel(frame, text=str(i + 1),
                font=ctk.CTkFont(size=10), text_color=G500)
            lbl.pack(pady=(1, 2))
            for w in (frame, cv, lbl):
                w.bind("<Button-1>", lambda e, idx=i: self._show_page(idx))
            # Store (None, frame, cv) — image filled in lazily
            self._thumb_imgs.append((None, frame, cv))
        # Kick off lazy rendering in small batches
        self._thumb_render_next = 0
        self._render_thumb_batch()

    def _render_thumb_batch(self, batch=6):
        """Render a batch of thumbnails, then schedule the next batch."""
        if not self.doc:
            return
        start = self._thumb_render_next
        end   = min(start + batch, self.total_pages)
        for i in range(start, end):
            img, frame, cv = self._thumb_imgs[i]
            if img is not None:
                continue  # already rendered (e.g. after undo/redo)
            new_img = _render_thumb(self.doc, i, self.THUMB_W)
            self._thumb_imgs[i] = (new_img, frame, cv)
            cv.configure(width=new_img.width(), height=new_img.height(),
                          bg=WHITE)
            cv.create_image(0, 0, anchor="nw", image=new_img)
        self._thumb_render_next = end
        if end < self.total_pages:
            self.parent.after(0, self._render_thumb_batch)

    def _highlight_thumb(self, idx):
        # Clear previous highlight in O(1)
        if self._highlighted_thumb_cv is not None:
            try:
                self._highlighted_thumb_cv.configure(
                    highlightbackground=G300, highlightthickness=2)
            except Exception:
                pass
        # Set new highlight
        if 0 <= idx < len(self._thumb_imgs):
            _, frame, cv = self._thumb_imgs[idx]
            cv.configure(highlightbackground=BLUE, highlightthickness=3)
            self._highlighted_thumb_cv = cv
        else:
            self._highlighted_thumb_cv = None

    # ==================================================================
    # TOC (Table of Contents)
    # ==================================================================

    def _build_toc(self):
        for w in self.toc_scroll.winfo_children():
            w.destroy()
        if not self.doc:
            return
        toc = self.doc.get_toc(simple=True)
        if not toc:
            ctk.CTkLabel(self.toc_scroll, text="No table of contents",
                text_color=G400, font=ctk.CTkFont(size=11)).pack(pady=20)
            return
        for level, title, page_num in toc:
            indent = (level - 1) * 12
            btn = ctk.CTkButton(self.toc_scroll,
                text=f"{title}",
                height=26, anchor="w",
                fg_color="transparent", hover_color=G200,
                text_color=G700,
                font=ctk.CTkFont(size=11,
                    weight="bold" if level == 1 else "normal"),
                command=lambda p=page_num: self._show_page(p - 1))
            btn.pack(fill="x", padx=(indent + 4, 4), pady=1)

    # ==================================================================
    # COORDINATE MAPPING
    # ==================================================================

    def _canvas_to_pdf(self, ex, ey):
        """Convert canvas event coords to PDF page coords."""
        cx = self.canvas.canvasx(ex)
        cy = self.canvas.canvasy(ey)
        pt = fitz.Point(cx - self._page_ox, cy - self._page_oy) * self._inv_mat
        return pt.x, pt.y

    def _pdf_to_canvas(self, px, py):
        """Convert PDF page coords to canvas scroll-region coords."""
        pt = fitz.Point(px, py) * self._render_mat
        return pt.x + self._page_ox, pt.y + self._page_oy

    def _pdf_rect_to_canvas(self, r):
        """Convert fitz.Rect to canvas coords (x0, y0, x1, y1)."""
        x0, y0 = self._pdf_to_canvas(r.x0, r.y0)
        x1, y1 = self._pdf_to_canvas(r.x1, r.y1)
        return min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)

    # ==================================================================
    # PAGE RENDERING
    # ==================================================================

    def _show_page(self, idx):
        if not self.doc or idx < 0 or idx >= self.total_pages:
            return
        self.current_page = idx
        self._clear_form_widgets()
        self._render_page()

        # Update nav
        self.page_entry.delete(0, "end")
        self.page_entry.insert(0, str(idx + 1))
        has_prev = idx > 0
        has_next = idx < self.total_pages - 1
        self.btn_first.configure(state="normal" if has_prev else "disabled")
        self.btn_prev.configure(state="normal"  if has_prev else "disabled")
        self.btn_next.configure(state="normal"  if has_next else "disabled")
        self.btn_last.configure(state="normal"  if has_next else "disabled")
        self._highlight_thumb(idx)

    def _render_page(self):
        """Render current page at current zoom/rotation, then draw overlays."""
        cw = max(self.canvas.winfo_width(), 300)
        ch = max(self.canvas.winfo_height(), 300)
        page = self.doc[self.current_page]
        pw, ph = page.rect.width, page.rect.height

        if self._rotation in (90, 270):
            pw, ph = ph, pw

        if self.zoom == FIT_PAGE:
            fw, fh = cw - 40, ch - 40
            scale = min(fw / pw, fh / ph)
            scale = max(scale, 0.05)
        elif self.zoom == FIT_WIDTH:
            scale = (cw - 40) / pw
            scale = max(scale, 0.05)
        else:
            scale = self.zoom

        mat = fitz.Matrix(scale, scale).prerotate(self._rotation)
        self._render_mat = mat
        self._inv_mat = ~mat

        pix = page.get_pixmap(matrix=mat, alpha=False)
        self._preview_img = _pix_to_tk(pix)
        iw, ih = self._preview_img.width(), self._preview_img.height()

        self.canvas.delete("all")

        if self.zoom == FIT_PAGE:
            self.canvas.configure(scrollregion=(0, 0, cw, ch))
            ox = (cw - iw) / 2
            oy = (ch - ih) / 2
        else:
            pad = 20
            total_w = max(iw + pad * 2, cw)
            total_h = max(ih + pad * 2, ch)
            self.canvas.configure(scrollregion=(0, 0, total_w, total_h))
            ox = (total_w - iw) / 2
            oy = pad

        self._page_ox = ox
        self._page_oy = oy

        self.canvas.create_image(ox, oy, anchor="nw",
            image=self._preview_img, tags="page_img")
        self.canvas.create_rectangle(ox - 1, oy - 1, ox + iw + 1, oy + ih + 1,
            outline=G300, width=1, tags="page_border")

        self._draw_overlays()

    def _draw_overlays(self):
        """Draw search highlights, selection, form widgets on top of page."""
        self.canvas.delete("search_hl", "sel_hl", "active_draw")
        self._sel_images.clear()
        self._draw_search_highlights()
        self._draw_text_selection()
        self._draw_form_widgets()

    def _refresh_overlays(self):
        """Redraw only overlays without re-rendering the page image."""
        self.canvas.delete("search_hl", "sel_hl", "active_draw")
        self._sel_images.clear()
        self._draw_search_highlights()
        self._draw_text_selection()

    # ==================================================================
    # NAVIGATION
    # ==================================================================

    def _first_page(self):
        self._show_page(0)

    def _prev_page(self):
        if self.current_page > 0:
            self._show_page(self.current_page - 1)

    def _next_page(self):
        if self.current_page < self.total_pages - 1:
            self._show_page(self.current_page + 1)

    def _last_page(self):
        if self.total_pages > 0:
            self._show_page(self.total_pages - 1)

    def _goto_page(self, _event=None):
        try:
            num = int(self.page_entry.get())
            if 1 <= num <= self.total_pages:
                self._show_page(num - 1)
        except ValueError:
            pass

    # ==================================================================
    # ZOOM & ROTATION
    # ==================================================================

    def _zoom_fit(self):
        self.zoom = FIT_PAGE
        self._update_zoom_label()
        if self.doc:
            self._show_page(self.current_page)

    def _zoom_fit_width(self):
        self.zoom = FIT_WIDTH
        self._update_zoom_label()
        if self.doc:
            self._show_page(self.current_page)

    def _zoom_in(self):
        if self.zoom in (FIT_PAGE, FIT_WIDTH):
            self.zoom = self._effective_zoom()
        self.zoom = min(self.zoom + self.ZOOM_STEP, self.ZOOM_MAX)
        self._update_zoom_label()
        if self.doc:
            self._show_page(self.current_page)

    def _zoom_out(self):
        if self.zoom in (FIT_PAGE, FIT_WIDTH):
            self.zoom = self._effective_zoom()
        self.zoom = max(self.zoom - self.ZOOM_STEP, self.ZOOM_MIN)
        self._update_zoom_label()
        if self.doc:
            self._show_page(self.current_page)

    def _effective_zoom(self):
        if not self.doc:
            return 1.0
        self.canvas.update_idletasks()
        cw = max(self.canvas.winfo_width(), 300)
        ch = max(self.canvas.winfo_height(), 300)
        page = self.doc[self.current_page]
        pw, ph = page.rect.width, page.rect.height
        if self._rotation in (90, 270):
            pw, ph = ph, pw
        if self.zoom == FIT_WIDTH:
            return (cw - 40) / pw
        fw, fh = cw - 40, ch - 40
        return min(fw / pw, fh / ph)

    def _update_zoom_label(self):
        if self.zoom == FIT_PAGE:
            self.zoom_lbl.configure(text="Fit")
            self.btn_fit.configure(fg_color=BLUE)
            self.btn_fitw.configure(fg_color=WHITE)
        elif self.zoom == FIT_WIDTH:
            self.zoom_lbl.configure(text="Width")
            self.btn_fit.configure(fg_color=WHITE)
            self.btn_fitw.configure(fg_color=BLUE)
        else:
            self.zoom_lbl.configure(text=f"{int(self.zoom * 100)}%")
            self.btn_fit.configure(fg_color=WHITE)
            self.btn_fitw.configure(fg_color=WHITE)

    def _rotate_view(self):
        self._rotation = (self._rotation + 90) % 360
        if self.doc:
            self._show_page(self.current_page)

    # ==================================================================
    # SCROLL & RESIZE
    # ==================================================================

    def _on_scroll(self, event):
        if event.state & 0x4:  # Ctrl
            if event.delta > 0:
                self._zoom_in()
            else:
                self._zoom_out()
        else:
            if self.zoom == FIT_PAGE:
                if event.delta > 0:
                    self._prev_page()
                else:
                    self._next_page()
            else:
                self.canvas.yview_scroll(-1 * (event.delta // 120), "units")

    def _on_canvas_resize(self, _event=None):
        if self.doc and self.zoom in (FIT_PAGE, FIT_WIDTH):
            self._render_page()

    def _escape(self):
        if self._search_visible:
            self._hide_search()
        else:
            self._set_tool(Tool.VIEW)
            self._selected_words.clear()
            self._selection_text = ""
            self._refresh_overlays()

    # ==================================================================
    # SEARCH  (Ctrl+F)
    # ==================================================================

    def _toggle_search(self):
        if self._search_visible:
            self._hide_search()
        else:
            self._show_search()

    def _show_search(self):
        if not self._search_visible:
            self._search_frame.pack(fill="x",
                before=self._search_frame.master.winfo_children()[-1])
            self._search_visible = True
        self._search_entry.focus_set()

    def _hide_search(self):
        self._search_frame.pack_forget()
        self._search_visible = False
        self._search_results.clear()
        self._search_flat.clear()
        self._search_idx = -1
        self._search_count_lbl.configure(text="")
        if self.doc:
            self._refresh_overlays()

    def _do_search(self):
        query = self._search_entry.get().strip()
        if not query or not self.doc:
            return
        self._search_results.clear()
        self._search_flat.clear()
        for i in range(self.total_pages):
            page = self.doc[i]
            rects = page.search_for(query)
            if rects:
                self._search_results.append((i, rects))
                for r in rects:
                    self._search_flat.append((i, r))
        total = len(self._search_flat)
        if total == 0:
            self._search_count_lbl.configure(text="0 results")
            self._search_idx = -1
        else:
            self._search_idx = 0
            for j, (pg, _) in enumerate(self._search_flat):
                if pg >= self.current_page:
                    self._search_idx = j
                    break
            self._goto_search_result()

    def _search_next(self):
        if not self._search_flat:
            return
        self._search_idx = (self._search_idx + 1) % len(self._search_flat)
        self._goto_search_result()

    def _search_prev(self):
        if not self._search_flat:
            return
        self._search_idx = (self._search_idx - 1) % len(self._search_flat)
        self._goto_search_result()

    def _goto_search_result(self):
        if self._search_idx < 0 or not self._search_flat:
            return
        pg, rect = self._search_flat[self._search_idx]
        total = len(self._search_flat)
        self._search_count_lbl.configure(text=f"{self._search_idx + 1}/{total}")
        if pg != self.current_page:
            self._show_page(pg)
        else:
            self._refresh_overlays()

    def _draw_search_highlights(self):
        if not self._search_flat or not self.doc:
            return
        for idx, (pg, rect) in enumerate(self._search_flat):
            if pg != self.current_page:
                continue
            x0, y0, x1, y1 = self._pdf_rect_to_canvas(rect)
            is_current = (idx == self._search_idx)
            fill = ORANGE if is_current else YELLOW_HL
            self.canvas.create_rectangle(x0, y0, x1, y1,
                fill=fill, stipple="gray50", outline="", tags="search_hl")

    # ==================================================================
    # TEXT SELECTION & COPY  (vibrant semi-transparent royal blue)
    # ==================================================================

    def _get_words(self):
        if not self.doc:
            return []
        page = self.doc[self.current_page]
        return page.get_text("words")

    def _words_in_rect(self, pdf_x0, pdf_y0, pdf_x1, pdf_y1):
        x0, y0 = min(pdf_x0, pdf_x1), min(pdf_y0, pdf_y1)
        x1, y1 = max(pdf_x0, pdf_x1), max(pdf_y0, pdf_y1)
        words = self._get_words()
        result = []
        for w in words:
            wx0, wy0, wx1, wy1 = w[0], w[1], w[2], w[3]
            if wx1 >= x0 and wx0 <= x1 and wy1 >= y0 and wy0 <= y1:
                result.append(w)
        result.sort(key=lambda w: (w[1], w[0]))
        return result

    def _select_words_flow(self, sx, sy, ex, ey):
        """Select words in reading-flow order between two PDF points.

        Unlike _words_in_rect, this selects all text between start and
        end in natural reading order — like dragging to select in a
        text editor.  pymupdf word tuples are:
        (x0, y0, x1, y1, word, block_no, line_no, word_no)
        """
        words = self._get_words()
        if not words:
            return []
        # Sort in reading order: block, line, then word index
        words.sort(key=lambda w: (w[5], w[6], w[7]))
        si = self._nearest_word_index(words, sx, sy)
        ei = self._nearest_word_index(words, ex, ey)
        if si is None or ei is None:
            return []
        lo, hi = min(si, ei), max(si, ei)
        return words[lo:hi + 1]

    @staticmethod
    def _nearest_word_index(words, px, py):
        """Return the index of the word whose centre is nearest to (px, py)."""
        best_idx = None
        best_dist = float('inf')
        for i, w in enumerate(words):
            cx = (w[0] + w[2]) * 0.5
            cy = (w[1] + w[3]) * 0.5
            d = (px - cx) ** 2 + (py - cy) ** 2
            if d < best_dist:
                best_dist = d
                best_idx = i
        return best_idx

    def _draw_text_selection(self):
        """Draw vibrant semi-transparent royal blue selection with
        negative horizontal padding (inset slightly from word edges)."""
        for w in self._selected_words:
            r = fitz.Rect(w[0], w[1], w[2], w[3])
            x0, y0, x1, y1 = self._pdf_rect_to_canvas(r)
            box_w = x1 - x0
            box_h = y1 - y0
            # Negative horizontal padding: inset ~4% from each side
            pad_x = max(1, box_w * 0.04)
            x0 += pad_x
            x1 -= pad_x

            # Try PIL semi-transparent overlay (true alpha blending)
            overlay = _make_sel_overlay(x1 - x0, box_h)
            if overlay:
                self._sel_images.append(overlay)
                self.canvas.create_image(x0, y0, anchor="nw",
                    image=overlay, tags="sel_hl")
            else:
                # Fallback: stipple-based highlight
                self.canvas.create_rectangle(x0, y0, x1, y1,
                    fill=BLUE, stipple="gray50", outline="", tags="sel_hl")

    def _copy_selection(self):
        if self._selection_text:
            top = self.parent.winfo_toplevel()
            top.clipboard_clear()
            top.clipboard_append(self._selection_text)

    # ==================================================================
    # MOUSE EVENTS (dispatch by tool)
    # ==================================================================

    def _on_mouse_down(self, event):
        if not self.doc:
            return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        self._drag_start = (cx, cy)
        px, py = self._canvas_to_pdf(event.x, event.y)

        if self._tool == Tool.FREEHAND:
            self._freehand_pts = [(px, py)]
        elif self._tool == Tool.SIGN:
            self._open_sign_dialog(event.x, event.y)
        elif self._tool == Tool.TEXT_BOX:
            self._open_textbox_dialog(px, py)
        elif self._tool == Tool.STICKY_NOTE:
            self._open_sticky_dialog(px, py)

    def _on_mouse_move(self, event):
        if not self.doc or self._drag_start is None:
            return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        sx, sy = self._drag_start
        px, py = self._canvas_to_pdf(event.x, event.y)

        if self._tool == Tool.VIEW:
            return

        elif self._tool in (Tool.SELECT, Tool.HIGHLIGHT,
                            Tool.UNDERLINE, Tool.STRIKETHROUGH):
            start_pt = fitz.Point(sx - self._page_ox,
                                  sy - self._page_oy) * self._inv_mat
            self._selected_words = self._select_words_flow(
                start_pt.x, start_pt.y, px, py)
            self._selection_text = " ".join(w[4] for w in self._selected_words)
            self._refresh_overlays()

        elif self._tool == Tool.FREEHAND:
            self._freehand_pts.append((px, py))
            if len(self._freehand_pts) >= 2:
                p0 = self._freehand_pts[-2]
                p1 = self._freehand_pts[-1]
                c0 = self._pdf_to_canvas(p0[0], p0[1])
                c1 = self._pdf_to_canvas(p1[0], p1[1])
                self.canvas.create_line(c0[0], c0[1], c1[0], c1[1],
                    fill=self._annot_color[1], width=self._stroke_width,
                    tags="active_draw")

        elif self._tool in (Tool.RECT, Tool.CIRCLE, Tool.LINE, Tool.ARROW):
            self.canvas.delete("active_draw")
            color = self._annot_color[1]
            w = self._stroke_width
            dx = cx - sx
            dy = cy - sy

            # Shift-key constraints
            if self._shift_held:
                if self._tool in (Tool.RECT, Tool.CIRCLE):
                    # Force square / circle (aspect ratio 1:1)
                    side = max(abs(dx), abs(dy))
                    dx = side if dx >= 0 else -side
                    dy = side if dy >= 0 else -side
                elif self._tool in (Tool.LINE, Tool.ARROW):
                    # Snap to nearest 45-degree angle
                    angle = math.atan2(dy, dx)
                    snap = round(angle / (math.pi / 4)) * (math.pi / 4)
                    length = math.hypot(dx, dy)
                    dx = length * math.cos(snap)
                    dy = length * math.sin(snap)

            ex = sx + dx
            ey = sy + dy

            if self._tool == Tool.RECT:
                self.canvas.create_rectangle(sx, sy, ex, ey,
                    outline=color, width=w, tags="active_draw")
            elif self._tool == Tool.CIRCLE:
                self.canvas.create_oval(sx, sy, ex, ey,
                    outline=color, width=w, tags="active_draw")
            elif self._tool in (Tool.LINE, Tool.ARROW):
                self.canvas.create_line(sx, sy, ex, ey,
                    fill=color, width=w,
                    arrow="last" if self._tool == Tool.ARROW else "none",
                    tags="active_draw")

    def _on_mouse_up(self, event):
        if not self.doc or self._drag_start is None:
            return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        sx, sy = self._drag_start
        px, py = self._canvas_to_pdf(event.x, event.y)
        start_pt = fitz.Point(sx - self._page_ox,
                              sy - self._page_oy) * self._inv_mat
        self._drag_start = None

        page = self.doc[self.current_page]
        _, _, fitz_rgb = self._annot_color

        if self._tool == Tool.SELECT:
            pass  # selection already done in _on_mouse_move

        elif self._tool in (Tool.HIGHLIGHT, Tool.UNDERLINE, Tool.STRIKETHROUGH):
            if self._selected_words:
                self._push_undo()
                quads = []
                for w in self._selected_words:
                    quads.append(fitz.Rect(w[0], w[1], w[2], w[3]).quad)
                if quads:
                    if self._tool == Tool.HIGHLIGHT:
                        annot = page.add_highlight_annot(quads)
                    elif self._tool == Tool.UNDERLINE:
                        annot = page.add_underline_annot(quads)
                    else:
                        annot = page.add_strikeout_annot(quads)
                    annot.set_colors(stroke=fitz_rgb)
                    annot.update()
                    self._modified = True
                    self._selected_words.clear()
                    self._selection_text = ""
                    self._show_page(self.current_page)

        elif self._tool == Tool.FREEHAND:
            if len(self._freehand_pts) >= 2:
                self._push_undo()
                points = [(float(x), float(y)) for x, y in self._freehand_pts]
                annot = page.add_ink_annot([points])
                annot.set_colors(stroke=fitz_rgb)
                annot.set_border(width=self._stroke_width)
                annot.update()
                self._modified = True
                self._freehand_pts.clear()
                self._show_page(self.current_page)

        elif self._tool == Tool.RECT:
            dx = cx - sx
            dy = cy - sy
            if self._shift_held:
                side = max(abs(dx), abs(dy))
                dx = side if dx >= 0 else -side
                dy = side if dy >= 0 else -side
            end_canvas_x = sx + dx
            end_canvas_y = sy + dy
            end_pt = fitz.Point(end_canvas_x - self._page_ox,
                                end_canvas_y - self._page_oy) * self._inv_mat
            r = fitz.Rect(start_pt, end_pt)
            r.normalize()
            if r.width > 2 and r.height > 2:
                self._push_undo()
                annot = page.add_rect_annot(r)
                annot.set_colors(stroke=fitz_rgb)
                annot.set_border(width=self._stroke_width)
                annot.update()
                self._modified = True
                self._show_page(self.current_page)

        elif self._tool == Tool.CIRCLE:
            dx = cx - sx
            dy = cy - sy
            if self._shift_held:
                side = max(abs(dx), abs(dy))
                dx = side if dx >= 0 else -side
                dy = side if dy >= 0 else -side
            end_canvas_x = sx + dx
            end_canvas_y = sy + dy
            end_pt = fitz.Point(end_canvas_x - self._page_ox,
                                end_canvas_y - self._page_oy) * self._inv_mat
            r = fitz.Rect(start_pt, end_pt)
            r.normalize()
            if r.width > 2 and r.height > 2:
                self._push_undo()
                annot = page.add_circle_annot(r)
                annot.set_colors(stroke=fitz_rgb)
                annot.set_border(width=self._stroke_width)
                annot.update()
                self._modified = True
                self._show_page(self.current_page)

        elif self._tool in (Tool.LINE, Tool.ARROW):
            dx = cx - sx
            dy = cy - sy
            if self._shift_held:
                angle = math.atan2(dy, dx)
                snap = round(angle / (math.pi / 4)) * (math.pi / 4)
                length = math.hypot(dx, dy)
                dx = length * math.cos(snap)
                dy = length * math.sin(snap)
            end_canvas_x = sx + dx
            end_canvas_y = sy + dy
            end_pt = fitz.Point(end_canvas_x - self._page_ox,
                                end_canvas_y - self._page_oy) * self._inv_mat
            p1 = fitz.Point(start_pt.x, start_pt.y)
            p2 = fitz.Point(end_pt.x, end_pt.y)
            if abs(p1.x - p2.x) > 2 or abs(p1.y - p2.y) > 2:
                self._push_undo()
                annot = page.add_line_annot(p1, p2)
                annot.set_colors(stroke=fitz_rgb)
                annot.set_border(width=self._stroke_width)
                if self._tool == Tool.ARROW:
                    annot.set_line_ends(
                        fitz.PDF_ANNOT_LE_NONE,
                        fitz.PDF_ANNOT_LE_CLOSED_ARROW)
                annot.update()
                self._modified = True
                self._show_page(self.current_page)

    # ==================================================================
    # DOUBLE-CLICK TO EDIT EXISTING TEXT ANNOTATIONS
    # ==================================================================

    def _on_double_click(self, event):
        """Double-click on a FreeText or Text (sticky) annotation to edit it."""
        if not self.doc:
            return
        px, py = self._canvas_to_pdf(event.x, event.y)
        click_pt = fitz.Point(px, py)
        page = self.doc[self.current_page]

        for annot in page.annots():
            if annot.rect.contains(click_pt):
                atype = annot.type[0]
                if atype == fitz.PDF_ANNOT_FREE_TEXT:
                    # Edit an existing FreeText annotation
                    self._open_textbox_dialog(px, py,
                                              existing_annot=annot)
                    return
                elif atype == fitz.PDF_ANNOT_TEXT:
                    # Edit a sticky note
                    self._edit_sticky(annot)
                    return

    def _edit_sticky(self, annot):
        """Re-edit an existing sticky note annotation."""
        old_text = annot.info.get("content", "")
        dialog = ctk.CTkInputDialog(
            text="Edit note text:", title="Edit Sticky Note")
        try:
            dialog._entry.insert(0, old_text)
        except Exception:
            pass
        text = dialog.get_input()
        if text is None:
            return
        self._push_undo()
        page = self.doc[self.current_page]
        if text:
            annot.set_info(content=text)
            annot.update()
        else:
            page.delete_annot(annot)
        self._modified = True
        self._show_page(self.current_page)

    # ==================================================================
    # TEXT BOX & STICKY NOTE DIALOGS
    # ==================================================================

    def _open_textbox_dialog(self, pdf_x, pdf_y, existing_annot=None):
        """Open dialog for creating or editing a FreeText annotation.

        If *existing_annot* is given the dialog is pre-filled and the
        old annotation is replaced with the updated one.
        """
        old_text = ""
        if existing_annot:
            old_text = existing_annot.info.get("content", "")
        dialog = ctk.CTkInputDialog(
            text="Enter text for the text box:" if not existing_annot
                 else "Edit text box:",
            title="Add Text Box" if not existing_annot else "Edit Text Box")
        # Pre-fill with existing text (CTkInputDialog doesn't natively
        # support default text, so we do it after construction)
        if old_text:
            try:
                entry = dialog._entry
                entry.insert(0, old_text)
            except Exception:
                pass
        text = dialog.get_input()
        if text is None:
            return  # cancelled
        if text == "" and not existing_annot:
            return

        self._push_undo()
        page = self.doc[self.current_page]
        _, hex_c, fitz_rgb = self._annot_color
        fontsize = max(8, self._stroke_width * 3)

        if existing_annot:
            # Replace the old annotation
            old_rect = existing_annot.rect
            page.delete_annot(existing_annot)
            if text:
                width = max(old_rect.width,
                            len(text) * fontsize * 0.6)
                height = max(old_rect.height, fontsize * 2)
                rect = fitz.Rect(old_rect.x0, old_rect.y0,
                                 old_rect.x0 + width,
                                 old_rect.y0 + height)
                annot = page.add_freetext_annot(
                    rect, text, fontsize=fontsize,
                    text_color=fitz_rgb, fontname="helv",
                    fill_color=(1, 1, 1))
                annot.update()
        else:
            width = max(100, len(text) * fontsize * 0.6)
            height = fontsize * 2.5
            rect = fitz.Rect(pdf_x, pdf_y,
                             pdf_x + width, pdf_y + height)
            annot = page.add_freetext_annot(
                rect, text, fontsize=fontsize,
                text_color=fitz_rgb, fontname="helv",
                fill_color=(1, 1, 1))
            annot.update()

        self._modified = True
        self._show_page(self.current_page)

    def _open_sticky_dialog(self, pdf_x, pdf_y):
        dialog = ctk.CTkInputDialog(
            text="Enter note text:",
            title="Add Sticky Note")
        text = dialog.get_input()
        if text:
            self._push_undo()
            page = self.doc[self.current_page]
            point = fitz.Point(pdf_x, pdf_y)
            annot = page.add_text_annot(point, text, icon="Comment")
            _, _, fitz_rgb = self._annot_color
            annot.set_colors(stroke=fitz_rgb)
            annot.update()
            self._modified = True
            self._show_page(self.current_page)

    # ==================================================================
    # SIGNATURE (smooth Bezier path interpolation)
    # ==================================================================

    def _open_sign_dialog(self, event_x, event_y):
        if not Image:
            messagebox.showerror("Error", "Pillow is required for signatures.")
            return

        pdf_x, pdf_y = self._canvas_to_pdf(event_x, event_y)

        win = ctk.CTkToplevel(self.parent.winfo_toplevel())
        win.title("Draw Signature")
        win.geometry("420x270")
        win.resizable(False, False)
        win.grab_set()

        ctk.CTkLabel(win, text="Draw your signature below:",
            font=ctk.CTkFont(size=13)).pack(pady=(8, 4))

        sig_canvas = tk.Canvas(win, width=400, height=150, bg="white",
            highlightthickness=1, highlightbackground=G300)
        sig_canvas.pack(padx=10)

        # Raw points per stroke; None separates strokes
        raw_points: list = []

        def _draw(e):
            if raw_points and raw_points[-1] is not None:
                sig_canvas.create_line(raw_points[-1][0], raw_points[-1][1],
                    e.x, e.y, fill="black", width=2, smooth=True,
                    capstyle="round", joinstyle="round")
            raw_points.append((e.x, e.y))

        def _reset(e):
            raw_points.append(None)  # stroke separator

        sig_canvas.bind("<B1-Motion>", _draw)
        sig_canvas.bind("<ButtonRelease-1>", _reset)

        def _apply():
            actual = [p for p in raw_points if p is not None]
            if len(actual) < 3:
                win.destroy()
                return

            self._push_undo()

            # Split into strokes and smooth each with Catmull-Rom
            img = Image.new("RGBA", (400, 150), (255, 255, 255, 0))
            draw = ImageDraw.Draw(img)
            stroke: list = []
            for pt in raw_points:
                if pt is None:
                    if len(stroke) >= 2:
                        smoothed = _smooth_stroke(stroke, num_interp=6)
                        flat = []
                        for s in smoothed:
                            flat.extend(s)
                        if len(flat) >= 4:
                            draw.line(flat, fill="black", width=2,
                                      joint="curve")
                    stroke = []
                else:
                    stroke.append(pt)
            if len(stroke) >= 2:
                smoothed = _smooth_stroke(stroke, num_interp=6)
                flat = []
                for s in smoothed:
                    flat.extend(s)
                if len(flat) >= 4:
                    draw.line(flat, fill="black", width=2, joint="curve")

            bbox = img.getbbox()
            if bbox:
                img = img.crop(bbox)

            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)

            page = self.doc[self.current_page]
            sig_w = img.width * 0.5
            sig_h = img.height * 0.5
            rect = fitz.Rect(pdf_x, pdf_y, pdf_x + sig_w, pdf_y + sig_h)
            page.insert_image(rect, stream=buf.read())
            self._modified = True
            win.destroy()
            self._show_page(self.current_page)

        def _clear():
            sig_canvas.delete("all")
            raw_points.clear()

        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(pady=8)
        ctk.CTkButton(btn_row, text="Clear", width=80, height=32,
            fg_color=G400, hover_color=G500,
            command=_clear).pack(side="left", padx=8)
        ctk.CTkButton(btn_row, text="Apply", width=80, height=32,
            fg_color=BLUE, hover_color=BLUE_HOVER,
            command=_apply).pack(side="left", padx=8)
        ctk.CTkButton(btn_row, text="Cancel", width=80, height=32,
            fg_color=WHITE, hover_color=G200, text_color=G700,
            border_color=G300, border_width=1,
            command=win.destroy).pack(side="left", padx=8)

    # ==================================================================
    # FORM FILLING
    # ==================================================================

    def _clear_form_widgets(self):
        for wid in self._form_windows:
            try:
                self.canvas.delete(wid)
            except Exception:
                pass
        self._form_windows.clear()

    def _draw_form_widgets(self):
        if not self.doc:
            return
        page = self.doc[self.current_page]
        widget_iter = page.widgets()
        if widget_iter is None:
            return

        for widget in widget_iter:
            rect = widget.rect
            x0, y0, x1, y1 = self._pdf_rect_to_canvas(rect)
            w = max(int(x1 - x0), 20)
            h = max(int(y1 - y0), 18)

            if widget.field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                entry = ctk.CTkEntry(self.canvas, width=w, height=h,
                    fg_color=WHITE, border_color=BLUE, border_width=1,
                    text_color=G900, font=ctk.CTkFont(size=max(9, h - 8)))
                if widget.field_value:
                    entry.insert(0, widget.field_value)
                entry._pdf_widget = widget
                entry.bind("<FocusOut>",
                    lambda e, ent=entry: self._update_form_field(ent))
                entry.bind("<Return>",
                    lambda e, ent=entry: self._update_form_field(ent))
                wid = self.canvas.create_window(x0, y0, anchor="nw",
                    window=entry, width=w, height=h, tags="form_w")
                self._form_windows.append(wid)

            elif widget.field_type in (fitz.PDF_WIDGET_TYPE_CHECKBOX,):
                var = tk.BooleanVar(value=bool(widget.field_value))
                cb = ctk.CTkCheckBox(self.canvas, text="", width=h,
                    height=h, variable=var,
                    command=lambda w=widget, v=var: self._update_checkbox(w, v))
                wid = self.canvas.create_window(x0, y0, anchor="nw",
                    window=cb, width=h, height=h, tags="form_w")
                self._form_windows.append(wid)

            elif widget.field_type == fitz.PDF_WIDGET_TYPE_COMBOBOX:
                choices = widget.choice_values or []
                if choices:
                    combo = ctk.CTkComboBox(self.canvas, values=choices,
                        width=w, height=h, fg_color=WHITE,
                        border_color=BLUE, text_color=G900)
                    if widget.field_value:
                        combo.set(widget.field_value)
                    combo._pdf_widget = widget
                    combo.configure(
                        command=lambda val, c=combo: self._update_combo(c, val))
                    wid = self.canvas.create_window(x0, y0, anchor="nw",
                        window=combo, width=w, height=h, tags="form_w")
                    self._form_windows.append(wid)

    def _update_form_field(self, entry):
        self._push_undo()
        widget = entry._pdf_widget
        widget.field_value = entry.get()
        widget.update()
        self._modified = True

    def _update_checkbox(self, widget, var):
        self._push_undo()
        widget.field_value = var.get()
        widget.update()
        self._modified = True

    def _update_combo(self, combo, val):
        self._push_undo()
        widget = combo._pdf_widget
        widget.field_value = val
        widget.update()
        self._modified = True

    # ==================================================================
    # SAVE & PRINT
    # ==================================================================

    def _save_pdf(self):
        if not self.doc:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=Path(self.pdf_path).name if self.pdf_path else "output.pdf")
        if not path:
            return
        try:
            if path == self.pdf_path:
                self.doc.save(path, incremental=True,
                              encryption=fitz.PDF_ENCRYPT_KEEP)
            else:
                self.doc.save(path)
            self._modified = False
            messagebox.showinfo("Saved", f"PDF saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save:\n{e}")

    def _print_pdf(self):
        if not self.doc:
            return
        try:
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp_path = tmp.name
            tmp.close()
            self.doc.save(tmp_path)
            if os.name == "nt":
                os.startfile(tmp_path, "print")
            elif os.name == "posix":
                subprocess.run(["lpr", tmp_path], check=True)
            else:
                messagebox.showinfo("Print",
                    f"PDF saved to {tmp_path}\nPlease print it manually.")
        except Exception as e:
            messagebox.showerror("Error", f"Print failed:\n{e}")

    # ==================================================================
    # CLEANUP
    # ==================================================================

    def cleanup(self):
        if self.doc:
            self.doc.close()
