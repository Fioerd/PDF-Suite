"""Extractor – PDF Toolbox Application.

Run this single file to start the app:
    python main.py
"""

import tkinter as tk
import customtkinter as ctk

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------
BG        = "#EEF2F7"
WHITE     = "#FFFFFF"
G100      = "#F3F4F6"
G200      = "#E5E7EB"
G300      = "#D1D5DB"
G400      = "#9CA3AF"
G500      = "#6B7280"
G600      = "#4B5563"
G700      = "#374151"
G900      = "#111827"

TEAL      = "#22B8A0"   # Organize
BLUE      = "#3B82F6"   # Convert
GREEN     = "#16A34A"   # Sign & Security
CORAL     = "#F87171"   # View & Edit
RED       = "#EF4444"   # Advanced

SOON_TXT  = "#B0B8C4"
SOON_BG   = "#E8ECF0"

QS_BG         = "#E8F1FB"   # Quick Start drag-drop zone background
BLUE_ACCENT   = "#2563EB"   # Active tab underline + browse button

# ---------------------------------------------------------------------------
# Implemented tools  (add tool IDs here as you build them)
# ---------------------------------------------------------------------------
IMPLEMENTED = {"split", "view", "excerpt", "pdf_to_csv"}

# ---------------------------------------------------------------------------
# Tool definitions  –  id · display name · icon char
# ---------------------------------------------------------------------------
CATEGORIES = [
    {
        "title": "Organize",
        "color": TEAL,
        "tools": [
            ("adjust_size",       "Adjust page size/scale", "⇲"),
            ("crop",              "Crop PDF",               "⬚"),
            ("extract_pages",     "Extract page(s)",        "⤓"),
            ("merge",             "Merge",                  "+"),
            ("multi_layout",      "Multi-Page Layout",      "⊞"),
            ("organize",          "Organize",               "☰"),
            ("multi_tool",        "PDF Multi Tool",         "✦"),
            ("remove",            "Remove",                 "🗑"),
            ("rotate",            "Rotate",                 "↻"),
            ("single_large_page", "Single Large Page",      "▯"),
            ("split",             "Split",                  "✂"),
            ("excerpt",           "Excerpt Tool",           "📋"),
        ],
    },
    {
        "title": "Convert to PDF",
        "color": BLUE,
        "tools": [
            ("img_to_pdf", "Image to PDF", "🖼"),
        ],
    },
    {
        "title": "Convert from PDF",
        "color": BLUE,
        "tools": [
            ("pdf_to_csv",  "PDF to CSV",        "📊"),
            ("pdf_to_img",  "PDF to Image",       "🖼"),
            ("pdf_to_rtf",  "PDF to RTF (Text)",  "T"),
        ],
    },
    {
        "title": "Sign & Security",
        "color": GREEN,
        "tools": [
            ("add_password",       "Add Password",              "🔒"),
            ("add_stamp",          "Add Stamp to PDF",          "📤"),
            ("add_watermark",      "Add Watermark",             "💧"),
            ("auto_redact",        "Auto Redact",               "▮"),
            ("change_permissions",  "Change Permissions",       "🛡"),
            ("manual_redaction",   "Manual Redaction",          "▬"),
            ("remove_cert_sign",   "Remove Certificate Sign",   "📜"),
            ("remove_password",    "Remove Password",           "🔓"),
            ("sanitize",           "Sanitize",                  "🧹"),
            ("sign",               "Sign",                      "✍"),
            ("sign_cert",          "Sign with Certificate",     "📝"),
            ("validate_sig",       "Validate PDF Signature",    "✅"),
        ],
    },
    {
        "title": "View & Edit",
        "color": CORAL,
        "tools": [
            ("view",               "View PDF",                 "👁"),
            ("add_image",          "Add image",                "Tᵢ"),
            ("add_page_numbers",   "Add Page Numbers",         "1²³"),
            ("change_metadata",    "Change Metadata",          "✏"),
            ("compare",            "Compare",                  "🔍"),
            ("extract_images",     "Extract Images",           "🖼"),
            ("flatten",            "Flatten",                  "▱"),
            ("get_info",           "Get ALL Info on PDF",      "ℹ"),
            ("remove_annotations", "Remove Annotations",       "🗑"),
            ("remove_blank",       "Remove Blank pages",       "📄"),
            ("remove_image",       "Remove image",             "✖"),
            ("replace_color",      "Replace and Invert Color", "🎨"),
            ("unlock_forms",       "Unlock PDF Forms",         "🔓"),
            ("view_edit",          "View/Edit PDF",            "📝"),
        ],
    },
    {
        "title": "Advanced",
        "color": RED,
        "tools": [
            ("adjust_colors",      "Adjust Colors/Contrast",    "🎨"),
            ("auto_rename",        "Auto Rename PDF File",      "✏"),
            ("auto_split_size",    "Auto Split by Size/Count",  "📏"),
            ("auto_split_pages",   "Auto Split Pages",          "✂"),
            ("compress",           "Compress",                  "📦"),
            ("overlay",            "Overlay PDFs",              "🗂"),
            ("pipeline",           "Pipeline",                  "⛓"),
            ("show_js",            "Show Javascript",           "Js"),
            ("split_chapters",     "Split PDF by Chapters",     "📖"),
            ("split_sections",     "Split PDF by Sections",     "⊞"),
        ],
    },
]

# ---------------------------------------------------------------------------
# Short descriptions for every tool (shown on cards)
# ---------------------------------------------------------------------------
TOOL_DESCRIPTIONS = {
    "adjust_size":        "Resize or rescale PDF pages.",
    "crop":               "Trim pages to a custom region.",
    "extract_pages":      "Pull out specific pages.",
    "merge":              "Combine multiple PDFs into one.",
    "multi_layout":       "Arrange pages in multi-up layouts.",
    "organize":           "Reorder, delete, or rotate pages.",
    "multi_tool":         "Batch operations on PDF pages.",
    "remove":             "Delete selected pages from PDF.",
    "rotate":             "Rotate pages to any angle.",
    "single_large_page":  "Combine pages into one large page.",
    "split":              "Separate a PDF into individual pages.",
    "excerpt":            "Capture regions from multiple PDFs.",
    "img_to_pdf":         "Convert images into a PDF file.",
    "pdf_to_csv":         "Extract tables to spreadsheet format.",
    "pdf_to_img":         "Export pages as image files.",
    "pdf_to_rtf":         "Convert PDF content to editable text.",
    "add_password":       "Encrypt PDF with a password.",
    "add_stamp":          "Stamp a mark on each page.",
    "add_watermark":      "Overlay custom text or image.",
    "auto_redact":        "Automatically hide sensitive content.",
    "change_permissions": "Set printing and copying rights.",
    "manual_redaction":   "Manually black out private content.",
    "remove_cert_sign":   "Strip digital certificate signatures.",
    "remove_password":    "Remove security restrictions.",
    "sanitize":           "Clean hidden data from PDF.",
    "sign":               "Add a handwritten digital signature.",
    "sign_cert":          "Sign with a certificate authority.",
    "validate_sig":       "Verify digital signature validity.",
    "view":               "Open and read any PDF file.",
    "add_image":          "Insert images into PDF pages.",
    "add_page_numbers":   "Number pages automatically.",
    "change_metadata":    "Edit title, author, and keywords.",
    "compare":            "Spot differences between two PDFs.",
    "extract_images":     "Pull images out of a PDF.",
    "flatten":            "Flatten annotations and form fields.",
    "get_info":           "View all metadata and properties.",
    "remove_annotations": "Strip comments and markup.",
    "remove_blank":       "Delete empty pages automatically.",
    "remove_image":       "Remove embedded images from PDF.",
    "replace_color":      "Invert or swap page colors.",
    "unlock_forms":       "Make locked form fields editable.",
    "view_edit":          "Read and annotate PDF files.",
    "adjust_colors":      "Fine-tune brightness and contrast.",
    "auto_rename":        "Rename file based on PDF content.",
    "auto_split_size":    "Split by file size or page count.",
    "auto_split_pages":   "Split pages by content structure.",
    "compress":           "Reduce file size, keep quality.",
    "overlay":            "Layer two PDFs on top of each other.",
    "pipeline":           "Chain multiple PDF operations.",
    "show_js":            "Inspect embedded JavaScript code.",
    "split_chapters":     "Split PDF by bookmark chapters.",
    "split_sections":     "Split PDF by logical sections.",
}

# ---------------------------------------------------------------------------
# Tab filter mapping  (None = show all)
# ---------------------------------------------------------------------------
TAB_CATEGORIES = {
    "All Tools": None,
    "Convert":   {"Convert to PDF", "Convert from PDF"},
    "Edit":      {"Organize", "View & Edit", "Advanced"},
    "Protect":   {"Sign & Security"},
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _lighten(hex_color: str, factor: float = 0.55) -> str:
    """Return a lighter version of a hex color."""
    r = int(hex_color[1:3], 16)
    g = int(hex_color[3:5], 16)
    b = int(hex_color[5:7], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"#{r:02x}{g:02x}{b:02x}"


def _draw_rounded_rect(cv, x0, y0, x1, y1, radius=8, fill="#000", outline=""):
    """Draw a filled rounded rectangle on a tk.Canvas using a smooth polygon."""
    points = [
        x0 + radius, y0,
        x1 - radius, y0,
        x1, y0,
        x1, y0 + radius,
        x1, y1 - radius,
        x1, y1,
        x1 - radius, y1,
        x0 + radius, y1,
        x0, y1,
        x0, y1 - radius,
        x0, y0 + radius,
        x0, y0,
    ]
    cv.create_polygon(points, fill=fill, outline=outline, smooth=True)


def _draw_pdf_icon(cv, x, y, w, h, color=BLUE_ACCENT):
    """Draw a simple line-art PDF page icon on a canvas.
    (x, y) is top-left corner; w×h is the bounding box.
    """
    fold = w * 0.30   # dog-ear size
    # Page outline with dog-ear
    cv.create_polygon(
        x, y,
        x + w - fold, y,
        x + w, y + fold,
        x + w, y + h,
        x, y + h,
        fill=WHITE, outline=color, width=1.5,
    )
    # Dog-ear fold
    cv.create_line(x + w - fold, y, x + w - fold, y + fold,
                   x + w, y + fold, fill=color, width=1.5)
    # Text lines
    lx0, lx1 = x + w * 0.18, x + w * 0.78
    for ly in [y + h * 0.44, y + h * 0.57, y + h * 0.70]:
        cv.create_line(lx0, ly, lx1, ly, fill=color, width=1.5)


def _draw_pdf_download_icon(cv, cx, cy, w, h, color=BLUE_ACCENT):
    """Draw a PDF-with-down-arrow icon centred at (cx, cy)."""
    x0, y0 = cx - w // 2, cy - h // 2
    _draw_pdf_icon(cv, x0, y0, w, h, color=color)
    # Down-arrow below page
    ax, ay = cx, y0 + h + 6
    cv.create_line(ax, ay, ax, ay + 12, fill=color, width=2)
    cv.create_line(ax - 6, ay + 6, ax, ay + 12,
                   ax + 6, ay + 6, fill=color, width=2)


# ═══════════════════════════════════════════════════════════════════════════
# Main Application
# ═══════════════════════════════════════════════════════════════════════════

class ExtractorApp:

    def __init__(self):
        ctk.set_appearance_mode("light")
        self.root = ctk.CTk()
        self.root.title("Extractor")
        self.root.geometry("1420x880")
        self.root.minsize(1100, 700)
        self.root.configure(fg_color=BG)

        # Content container – swapped between views
        self._content = ctk.CTkFrame(self.root, fg_color="transparent")
        self._content.pack(fill="both", expand=True)

        # Search / filter state
        self._tool_widgets: list[tuple[str, tk.Widget]] = []
        self._tool_visible: list[bool] = []
        self._search_after_id = None
        self._active_tab = "All Tools"
        self._tab_buttons: dict = {}
        self._all_tool_data: list = []
        self._qs_canvas = None
        self._grid_frame = None
        self._search_entry = None

        self.show_home()

    # ==================================================================
    # NAVIGATION
    # ==================================================================

    def _clear(self):
        if self._search_after_id is not None:
            self.root.after_cancel(self._search_after_id)
            self._search_after_id = None
        for w in self._content.winfo_children():
            w.destroy()
        self._tool_widgets.clear()
        self._tool_visible.clear()
        self._tab_buttons.clear()
        self._all_tool_data.clear()
        self._active_tab = "All Tools"
        self._qs_canvas = None
        self._grid_frame = None
        self._search_entry = None

    def show_home(self):
        self._clear()
        self._build_home()

    def show_tool(self, tool_id: str):
        if tool_id not in IMPLEMENTED:
            from tkinter import messagebox
            messagebox.showinfo(
                "Coming Soon",
                "This tool is not yet implemented.\nStay tuned for future updates!",
            )
            return

        self._clear()
        self._build_tool_view(tool_id)

    # ==================================================================
    # HOME SCREEN
    # ==================================================================

    def _build_home(self):
        scroll = ctk.CTkScrollableFrame(
            self._content, fg_color=BG, corner_radius=0,
            scrollbar_button_color="#C4CCD8",
            scrollbar_button_hover_color="#A0AAB6",
        )
        scroll.pack(fill="both", expand=True)

        inner = ctk.CTkFrame(scroll, fg_color="transparent")
        inner.pack(fill="x", expand=True, padx=40, pady=(16, 30))

        self._build_header(inner)
        self._build_quickstart(inner)
        self._build_recently_used(inner)
        self._build_filter_row(inner)
        self._build_tool_grid(inner)
        self._build_footer(inner)

    # ---- header -------------------------------------------------------
    def _build_header(self, parent):
        bar = ctk.CTkFrame(parent, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 4))

        # PDF page icon drawn with canvas (no external assets)
        cv = tk.Canvas(bar, width=28, height=34, bg=BG, highlightthickness=0)
        cv.pack(side="left", padx=(0, 8))
        _draw_pdf_icon(cv, 2, 1, 24, 30, color=BLUE_ACCENT)

        ctk.CTkLabel(
            bar, text="Extractor",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=G900,
        ).pack(side="left")

    # ---- quick start --------------------------------------------------
    def _build_quickstart(self, parent):
        ctk.CTkLabel(
            parent, text="Quick Start",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=G900, anchor="w",
        ).pack(fill="x", pady=(20, 8))

        zone = tk.Canvas(parent, height=170, bg=QS_BG, highlightthickness=0)
        zone.pack(fill="x", pady=(0, 24))
        zone.bind("<Configure>", self._draw_quickstart_zone)
        zone.bind("<Button-1>", self._on_quickstart_click)
        self._qs_canvas = zone

    def _draw_quickstart_zone(self, event=None):
        cv = self._qs_canvas
        if cv is None:
            return
        cv.delete("all")
        w = cv.winfo_width()
        h = cv.winfo_height()
        if w < 2 or h < 2:
            return

        # Dashed rounded border
        r = 12
        pad = 6
        cv.create_arc(pad, pad, pad + r*2, pad + r*2,
                      start=90, extent=90, style="arc",
                      outline=BLUE_ACCENT, width=1, dash=(6, 4))
        cv.create_arc(w - pad - r*2, pad, w - pad, pad + r*2,
                      start=0, extent=90, style="arc",
                      outline=BLUE_ACCENT, width=1, dash=(6, 4))
        cv.create_arc(pad, h - pad - r*2, pad + r*2, h - pad,
                      start=180, extent=90, style="arc",
                      outline=BLUE_ACCENT, width=1, dash=(6, 4))
        cv.create_arc(w - pad - r*2, h - pad - r*2, w - pad, h - pad,
                      start=270, extent=90, style="arc",
                      outline=BLUE_ACCENT, width=1, dash=(6, 4))
        cv.create_line(pad + r, pad, w - pad - r, pad,
                       fill=BLUE_ACCENT, width=1, dash=(6, 4))
        cv.create_line(pad + r, h - pad, w - pad - r, h - pad,
                       fill=BLUE_ACCENT, width=1, dash=(6, 4))
        cv.create_line(pad, pad + r, pad, h - pad - r,
                       fill=BLUE_ACCENT, width=1, dash=(6, 4))
        cv.create_line(w - pad, pad + r, w - pad, h - pad - r,
                       fill=BLUE_ACCENT, width=1, dash=(6, 4))

        cx = w // 2

        # PDF with download arrow
        iw, ih = 38, 46
        _draw_pdf_download_icon(cv, cx, 58, iw, ih, color=BLUE_ACCENT)

        # Main text
        cv.create_text(cx, 116,
                       text="Drag & Drop your PDF here",
                       font=("Segoe UI", 14, "bold"),
                       fill=G700)

        # Browse button (drawn as rounded pill)
        bw, bh = 160, 30
        bx0, by0 = cx - bw // 2, 138
        bx1, by1 = cx + bw // 2, by0 + bh
        self._browse_btn_bounds = (bx0, by0, bx1, by1)
        _draw_rounded_rect(cv, bx0, by0, bx1, by1, radius=15,
                           fill=WHITE, outline=G300)
        cv.create_text(cx, by0 + bh // 2,
                       text="Or click to browse",
                       font=("Segoe UI", 11),
                       fill=G600)

    def _on_quickstart_click(self, event):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Open PDF",
            filetypes=[("PDF files", "*.pdf")],
        )
        if path:
            self.show_tool("view")

    # ---- recently used ------------------------------------------------
    def _build_recently_used(self, parent):
        implemented_tools = []
        for cat in CATEGORIES:
            for tid, tname, ticon in cat["tools"]:
                if tid in IMPLEMENTED:
                    implemented_tools.append((tid, tname, ticon, cat["color"]))

        if not implemented_tools:
            return

        lbl = ctk.CTkLabel(parent, text="Recently Used",
                            font=ctk.CTkFont(size=14, weight="bold"),
                            text_color=G500, anchor="w")
        lbl.pack(fill="x", pady=(0, 10))

        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(0, 24))

        for tid, tname, ticon, color in implemented_tools:
            card = ctk.CTkFrame(row, fg_color=WHITE, corner_radius=12,
                                 border_width=1, border_color=G200,
                                 cursor="hand2")
            card.pack(side="left", padx=(0, 12))

            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="both", expand=True, padx=12, pady=12)

            # Colored rounded-square icon
            cv = tk.Canvas(inner, width=46, height=46,
                            bg=WHITE, highlightthickness=0)
            cv.pack(side="left", padx=(0, 12))
            _draw_rounded_rect(cv, 0, 0, 46, 46, radius=8, fill=color)
            cv.create_text(23, 23, text=ticon, fill="white",
                            font=("Segoe UI Emoji", 16, "bold"))

            text_frame = ctk.CTkFrame(inner, fg_color="transparent")
            text_frame.pack(side="left", fill="y", expand=False)

            name_lbl = ctk.CTkLabel(
                text_frame, text=tname,
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=G700, anchor="w", cursor="hand2",
            )
            name_lbl.pack(anchor="w")

            desc = TOOL_DESCRIPTIONS.get(tid, "")
            if desc:
                ctk.CTkLabel(
                    text_frame, text=desc,
                    font=ctk.CTkFont(size=11),
                    text_color=G500, anchor="w",
                ).pack(anchor="w", pady=(2, 0))

            # Hover + click on every widget in the card
            def _enter(e, c=card): c.configure(border_color=G300)
            def _leave(e, c=card): c.configure(border_color=G200)
            for widget in (card, inner, cv, text_frame, name_lbl):
                widget.bind("<Enter>", _enter)
                widget.bind("<Leave>", _leave)
                widget.bind("<Button-1>",
                            lambda e, t=tid: self.show_tool(t))

    # ---- filter row (tabs + search) -----------------------------------
    def _build_filter_row(self, parent):
        sep = ctk.CTkFrame(parent, fg_color=G200, height=1)
        sep.pack(fill="x", pady=(0, 16))

        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(0, 16))

        # Tab pills (left side)
        tabs_frame = ctk.CTkFrame(row, fg_color="transparent")
        tabs_frame.pack(side="left")

        for tab_name in TAB_CATEGORIES:
            # Container for button + underline
            tab_wrap = ctk.CTkFrame(tabs_frame, fg_color="transparent")
            tab_wrap.pack(side="left", padx=(0, 4))

            btn = ctk.CTkButton(
                tab_wrap,
                text=tab_name,
                font=ctk.CTkFont(size=13),
                height=32,
                corner_radius=6,
                fg_color="transparent",
                text_color=G500,
                hover_color=G100,
                command=lambda t=tab_name: self._on_tab_click(t),
            )
            btn.pack()

            # Underline bar (hidden by default)
            underline = ctk.CTkFrame(tab_wrap, fg_color="transparent",
                                      height=2, corner_radius=1)
            underline.pack(fill="x", padx=4)

            self._tab_buttons[tab_name] = (btn, underline)

        self._update_tab_styles()

        # Search bar (right side)
        search_frame = ctk.CTkFrame(row, fg_color=WHITE, corner_radius=20,
                                     border_width=1, border_color=G200,
                                     height=38)
        search_frame.pack(side="right")
        search_frame.pack_propagate(False)

        icon = ctk.CTkLabel(search_frame, text="🔍",
                             font=ctk.CTkFont(size=14),
                             text_color=G400, width=26)
        icon.pack(side="left", padx=(12, 0))

        self._search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="Search tools...",
            font=ctk.CTkFont(size=13),
            fg_color="transparent", border_width=0,
            text_color=G700, placeholder_text_color=G400,
            width=220, height=36,
        )
        self._search_entry.pack(side="left", padx=(2, 12))
        self._search_entry.bind("<KeyRelease>", self._on_search)
        self._search_entry.bind("<FocusIn>", self._on_search)

    def _update_tab_styles(self):
        for tab_name, (btn, underline) in self._tab_buttons.items():
            if tab_name == self._active_tab:
                btn.configure(text_color=BLUE_ACCENT, fg_color="transparent")
                underline.configure(fg_color=BLUE_ACCENT)
            else:
                btn.configure(text_color=G500, fg_color="transparent")
                underline.configure(fg_color="transparent")

    def _on_tab_click(self, tab_name):
        self._active_tab = tab_name
        self._update_tab_styles()
        self._render_tool_grid()

    # ---- tool grid  (4 columns, card layout) -------------------------
    def _build_tool_grid(self, parent):
        self._grid_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._grid_frame.pack(fill="x", pady=(0, 30))
        for i in range(4):
            self._grid_frame.columnconfigure(i, weight=1)

        # Flatten all tools with their category metadata
        for cat in CATEGORIES:
            for tid, tname, ticon in cat["tools"]:
                self._all_tool_data.append(
                    (tid, tname, ticon, cat["color"], cat["title"])
                )

        self._render_tool_grid()

    def _render_tool_grid(self):
        """Clear and repopulate the grid based on active tab + search query."""
        self._search_after_id = None
        if self._grid_frame is None:
            return

        for w in self._grid_frame.winfo_children():
            w.destroy()
        self._tool_widgets.clear()
        self._tool_visible.clear()

        q = ""
        if self._search_entry is not None:
            try:
                q = self._search_entry.get().strip().lower()
            except Exception:
                pass

        tab_filter = TAB_CATEGORIES.get(self._active_tab)

        col, row_idx = 0, 0
        for (tid, tname, ticon, color, cat_title) in self._all_tool_data:
            if tab_filter is not None and cat_title not in tab_filter:
                continue
            if q and q not in tname.lower():
                continue
            is_impl = tid in IMPLEMENTED
            card = self._make_tool_card(
                self._grid_frame, col, row_idx,
                tid, tname, ticon, color, is_impl,
            )
            self._tool_widgets.append((tname.lower(), card))
            self._tool_visible.append(True)
            col += 1
            if col == 4:
                col = 0
                row_idx += 1

    def _make_tool_card(self, parent, col, row_idx,
                         tool_id, name, icon_char, color, implemented):
        card = ctk.CTkFrame(
            parent, fg_color=WHITE, corner_radius=12,
            border_width=1, border_color=G200,
            cursor="hand2" if implemented else "arrow",
        )
        card.grid(row=row_idx, column=col, padx=6, pady=6, sticky="nsew")

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=12, pady=12)

        # Colored rounded-square icon
        icon_color = color if implemented else _lighten(color, 0.55)
        cv = tk.Canvas(inner, width=46, height=46,
                        bg=WHITE, highlightthickness=0)
        cv.pack(side="left", padx=(0, 12))
        _draw_rounded_rect(cv, 0, 0, 46, 46, radius=8, fill=icon_color)
        cv.create_text(23, 23, text=icon_char, fill="white",
                        font=("Segoe UI Emoji", 16, "bold"))

        # Text section
        text_frame = ctk.CTkFrame(inner, fg_color="transparent")
        text_frame.pack(side="left", fill="both", expand=True)

        name_color = G700 if implemented else G400
        ctk.CTkLabel(
            text_frame, text=name,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=name_color, anchor="w",
            wraplength=160,
        ).pack(fill="x", anchor="w")

        desc = TOOL_DESCRIPTIONS.get(tool_id, "")
        desc_text = desc if implemented else "Coming Soon"
        desc_color = G500 if implemented else SOON_TXT
        ctk.CTkLabel(
            text_frame, text=desc_text,
            font=ctk.CTkFont(size=11),
            text_color=desc_color, anchor="w",
            wraplength=160,
        ).pack(fill="x", anchor="w", pady=(2, 0))

        if implemented:
            def _enter(e, c=card): c.configure(border_color=G300)
            def _leave(e, c=card): c.configure(border_color=G200)
            for w in (card, inner, cv, text_frame):
                w.bind("<Enter>", _enter)
                w.bind("<Leave>", _leave)
                w.bind("<Button-1>",
                       lambda e, t=tool_id: self.show_tool(t))

        return card

    # ---- footer -------------------------------------------------------
    def _build_footer(self, parent):
        sep = ctk.CTkFrame(parent, fg_color=G300, height=1)
        sep.pack(fill="x", pady=(10, 14))

        ctk.CTkLabel(
            parent,
            text="Licenses  ·  Releases  ·  Privacy Policy  ·  Terms and Conditions",
            font=ctk.CTkFont(size=11), text_color=G400,
        ).pack()

        ctk.CTkLabel(
            parent,
            text="Powered by Extractor",
            font=ctk.CTkFont(size=11), text_color=G400,
        ).pack(pady=(4, 0))

    # ---- search -------------------------------------------------------
    def _on_search(self, _event=None):
        if self._search_after_id is not None:
            self.root.after_cancel(self._search_after_id)
        self._search_after_id = self.root.after(80, self._render_tool_grid)

    # ==================================================================
    # TOOL VIEW  (loads tool screen with back button)
    # ==================================================================

    def _build_tool_view(self, tool_id):
        topbar = ctk.CTkFrame(self._content, fg_color=WHITE,
                               corner_radius=0, height=52)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)

        back = ctk.CTkButton(
            topbar, text="←  Back to Home", width=160, height=36,
            font=ctk.CTkFont(size=13), fg_color="transparent",
            text_color=G700, hover_color=G100,
            command=self.show_home,
        )
        back.pack(side="left", padx=16, pady=8)

        tool_area = ctk.CTkFrame(self._content, fg_color=WHITE, corner_radius=0)
        tool_area.pack(fill="both", expand=True)

        if tool_id == "split":
            from split_tool import SplitTool
            SplitTool(tool_area)
        elif tool_id == "view":
            from view_tool import ViewTool
            ViewTool(tool_area)
        elif tool_id == "excerpt":
            from excerpt_tool import ExcerptTool
            ExcerptTool(tool_area)
        elif tool_id == "pdf_to_csv":
            from pdf_to_csv_tool import PDFtoCSVTool
            PDFtoCSVTool(tool_area)

    # ==================================================================
    # RUN
    # ==================================================================

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = ExtractorApp()
    app.run()
