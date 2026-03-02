"""Excerpt Tool – Multi-document rubber-band region capture to a new PDF.

Load multiple PDFs, drag to select rectangular regions on any page,
and collect them into a growing excerpt document. Save at any time.
Uses native PDF crop (show_pdf_page with clip) to preserve text and links.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from dataclasses import dataclass, field

import customtkinter as ctk

try:
    import fitz  # pymupdf
except ImportError:
    fitz = None

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = ImageTk = None

# ---------------------------------------------------------------------------
# Colors  (matching existing tools)
# ---------------------------------------------------------------------------
BLUE        = "#3B82F6"
BLUE_HOVER  = "#2563EB"
GREEN       = "#16A34A"
GREEN_HOVER = "#15803D"
GREEN_TXT   = "#16A34A"
RED         = "#EF4444"
G100        = "#F3F4F6"
G200        = "#E5E7EB"
G300        = "#D1D5DB"
G400        = "#9CA3AF"
G500        = "#6B7280"
G700        = "#374151"
G900        = "#111827"
WHITE       = "#FFFFFF"

SEL_BLUE    = "#3B82F6"   # rubber-band outline colour


# ---------------------------------------------------------------------------
# Snippet data structure
# ---------------------------------------------------------------------------

@dataclass
class Snippet:
    source_path: str          # absolute path to source PDF
    page_index:  int          # 0-based page index within that PDF
    crop_rect:   object       # fitz.Rect in PDF coordinate space
    label:       str  = ""
    thumbnail:   object = field(default=None, repr=False)  # PhotoImage ref


# ---------------------------------------------------------------------------
# Shared rendering helpers
# ---------------------------------------------------------------------------

def _pix_to_tk(pixmap):
    if Image and ImageTk:
        img = Image.frombytes("RGB", (pixmap.width, pixmap.height),
                               pixmap.samples)
        return ImageTk.PhotoImage(img)
    return tk.PhotoImage(data=pixmap.tobytes("ppm"))


def _render(doc, idx, max_w):
    page = doc[idx]
    s = max_w / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
    return _pix_to_tk(pix)


# ===========================================================================
# ExcerptTool
# ===========================================================================

class ExcerptTool:
    THUMB_W = 80    # page thumbnail width in bottom strip
    SNIP_W  = 60    # snippet thumbnail width in left panel list
    LEFT_W  = 370   # fixed left panel width in pixels

    def __init__(self, parent: ctk.CTkFrame):
        self.parent = parent

        if fitz is None:
            ctk.CTkLabel(
                parent,
                text="Missing dependencies.\n\npip install pymupdf Pillow",
                font=ctk.CTkFont(size=16), text_color=G500,
            ).pack(expand=True)
            return

        # ---- Multi-document state ----
        # Each entry: {"path": str, "doc": fitz.Document, "name": str}
        self._pdf_list: list  = []
        self._active_idx: int = -1

        # ---- Per-active-doc view state ----
        self._current_page: int  = 0
        self._page_ox: float     = 0.0
        self._page_oy: float     = 0.0
        self._page_iw: float     = 0.0
        self._page_ih: float     = 0.0
        self._render_mat         = fitz.Matrix(1, 1)
        self._inv_mat            = fitz.Matrix(1, 1)
        self._preview_img        = None
        self._thumb_imgs: list   = []   # list of PhotoImage (or None while pending)
        self._thumb_positions: list = []  # list of (x, iw, ih) per page
        self._thumb_render_next: int = 0
        self._highlighted_thumb: int = -1  # currently highlighted page index

        # ---- Rubber-band state ----
        self._rb_start              = None
        self._rb_rect_id            = None
        self._selecting: bool       = False

        # ---- Snippets & output doc ----
        self._snippets: list        = []
        self._out_doc               = fitz.open()   # empty in-memory PDF

        self._build_ui()

    # ==================================================================
    # BUILD UI
    # ==================================================================

    def _build_ui(self):
        main = ctk.CTkFrame(self.parent, fg_color="transparent")
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        # Left panel
        left = ctk.CTkFrame(main, width=self.LEFT_W, fg_color=G100,
                             corner_radius=0)
        left.grid(row=0, column=0, sticky="nsew")
        left.grid_propagate(False)
        self._build_left_panel(left)

        # Right panel
        right = ctk.CTkFrame(main, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        self._build_right_panel(right)

        # Bottom thumbnail strip
        self._build_thumb_strip()

    # ------------------------------------------------------------------
    # Left panel
    # ------------------------------------------------------------------

    def _build_left_panel(self, left):
        left.rowconfigure(3, weight=1)   # snippet frame expands
        left.columnconfigure(0, weight=1)

        # Title
        ctk.CTkLabel(left, text="Excerpt Tool",
                      font=ctk.CTkFont(size=22, weight="bold"),
                      text_color=G900).grid(
            row=0, column=0, sticky="w", padx=16, pady=(14, 6))

        # ---- Loaded PDFs section ----
        pdf_sec = ctk.CTkFrame(left, fg_color="transparent")
        pdf_sec.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 6))
        pdf_sec.columnconfigure(0, weight=1)

        ctk.CTkLabel(pdf_sec, text="Loaded PDFs",
                      font=ctk.CTkFont(size=13, weight="bold"),
                      text_color=G700).grid(row=0, column=0, sticky="w",
                                             pady=(0, 4))

        btn_row = ctk.CTkFrame(pdf_sec, fg_color="transparent")
        btn_row.grid(row=1, column=0, sticky="ew", pady=(0, 6))

        ctk.CTkButton(btn_row, text="+ Add PDF",
                       width=110, height=32,
                       fg_color=BLUE, hover_color=BLUE_HOVER,
                       font=ctk.CTkFont(size=12, weight="bold"),
                       command=self._add_pdf).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_row, text="Remove",
                       width=80, height=32,
                       fg_color="transparent", border_color=G300,
                       border_width=1, text_color=G700, hover_color=G200,
                       command=self._remove_active_pdf).pack(side="left")

        self._pdf_listbox = ctk.CTkScrollableFrame(
            pdf_sec, fg_color=WHITE, height=130,
            scrollbar_button_color=G300,
            scrollbar_button_hover_color=G400,
            corner_radius=8,
        )
        self._pdf_listbox.grid(row=2, column=0, sticky="ew")

        # ---- Separator ----
        ctk.CTkFrame(left, fg_color=G300, height=1).grid(
            row=2, column=0, sticky="ew", padx=12, pady=(4, 6))

        # ---- Snippets section ----
        snip_sec = ctk.CTkFrame(left, fg_color="transparent")
        snip_sec.grid(row=3, column=0, sticky="nsew", padx=16, pady=(0, 0))
        snip_sec.rowconfigure(1, weight=1)
        snip_sec.columnconfigure(0, weight=1)

        snip_hdr = ctk.CTkFrame(snip_sec, fg_color="transparent")
        snip_hdr.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ctk.CTkLabel(snip_hdr, text="Captured Snippets",
                      font=ctk.CTkFont(size=13, weight="bold"),
                      text_color=G700).pack(side="left")
        self._snip_count_lbl = ctk.CTkLabel(snip_hdr, text="(0)",
                                              font=ctk.CTkFont(size=11),
                                              text_color=G400)
        self._snip_count_lbl.pack(side="left", padx=4)

        self._snip_frame = ctk.CTkScrollableFrame(
            snip_sec, fg_color=WHITE,
            scrollbar_button_color=G300,
            scrollbar_button_hover_color=G400,
            corner_radius=8,
        )
        self._snip_frame.grid(row=1, column=0, sticky="nsew")

        # ---- Save section ----
        save_sec = ctk.CTkFrame(left, fg_color="transparent")
        save_sec.grid(row=4, column=0, sticky="ew", padx=16,
                       pady=(8, 14))
        save_sec.columnconfigure(0, weight=1)

        ctk.CTkFrame(save_sec, fg_color=G300, height=1).grid(
            row=0, column=0, sticky="ew", pady=(0, 8))

        self._save_btn = ctk.CTkButton(
            save_sec, text="Save Excerpt PDF",
            height=42, font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=GREEN, hover_color=GREEN_HOVER,
            command=self._save_excerpt,
        )
        self._save_btn.grid(row=1, column=0, sticky="ew")

        self._status_lbl = ctk.CTkLabel(
            save_sec, text="",
            font=ctk.CTkFont(size=11), text_color=GREEN_TXT, anchor="w")
        self._status_lbl.grid(row=2, column=0, sticky="w", pady=(4, 0))

    # ------------------------------------------------------------------
    # Right panel (canvas + nav)
    # ------------------------------------------------------------------

    def _build_right_panel(self, right):
        self.canvas = tk.Canvas(right, bg=WHITE, highlightthickness=0,
                                 relief="flat")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.create_text(
            300, 250,
            text="Add a PDF, then drag to select a region",
            font=("", 16), fill=G400, tags="ph",
        )

        # Mouse bindings
        self.canvas.bind("<ButtonPress-1>",   self._on_rb_start)
        self.canvas.bind("<B1-Motion>",       self._on_rb_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_rb_release)
        self.canvas.bind("<Configure>",       self._on_canvas_resize)

        # Navigation bar
        nav = ctk.CTkFrame(right, fg_color="transparent")
        nav.grid(row=1, column=0, pady=(6, 0))

        self.btn_prev = ctk.CTkButton(
            nav, text="←", width=40, height=34,
            fg_color=WHITE, border_color=G300, border_width=1,
            text_color=G700, hover_color=G100, state="disabled",
            command=self._prev_page)
        self.btn_prev.pack(side="left", padx=4)

        self.page_entry = ctk.CTkEntry(
            nav, width=52, height=34, justify="center",
            fg_color=WHITE, border_color=G300, text_color=G900,
            font=ctk.CTkFont(size=12))
        self.page_entry.pack(side="left", padx=2)
        self.page_entry.insert(0, "–")
        self.page_entry.bind("<Return>", self._goto_page)

        self.total_lbl = ctk.CTkLabel(
            nav, text="/ –",
            font=ctk.CTkFont(size=12), text_color=G500)
        self.total_lbl.pack(side="left", padx=(0, 4))

        self.btn_next = ctk.CTkButton(
            nav, text="→", width=40, height=34,
            fg_color=WHITE, border_color=G300, border_width=1,
            text_color=G700, hover_color=G100, state="disabled",
            command=self._next_page)
        self.btn_next.pack(side="left", padx=4)

        # Hint
        ctk.CTkLabel(
            right,
            text="Drag on the page to capture a region",
            font=ctk.CTkFont(size=11), text_color=G400,
        ).grid(row=2, column=0, pady=(2, 4))

    # ------------------------------------------------------------------
    # Bottom thumbnail strip
    # ------------------------------------------------------------------

    def _build_thumb_strip(self):
        bot = ctk.CTkFrame(self.parent, fg_color="transparent", height=155)
        bot.pack(fill="x", side="bottom", padx=0, pady=(4, 0))
        bot.pack_propagate(False)

        self.btn_tl = ctk.CTkButton(
            bot, text="‹", width=28, height=100,
            font=ctk.CTkFont(size=22),
            fg_color=G100, hover_color=G200, text_color=G500,
            corner_radius=8,
            command=lambda: self.thumb_cv.xview_scroll(-3, "units"))
        self.btn_tl.pack(side="left", padx=(0, 4))

        self.thumb_cv = tk.Canvas(bot, height=140, bg=G100,
                                   highlightthickness=0, relief="flat")
        self.thumb_cv.pack(side="left", fill="both", expand=True)

        self.btn_tr = ctk.CTkButton(
            bot, text="›", width=28, height=100,
            font=ctk.CTkFont(size=22),
            fg_color=G100, hover_color=G200, text_color=G500,
            corner_radius=8,
            command=lambda: self.thumb_cv.xview_scroll(3, "units"))
        self.btn_tr.pack(side="right", padx=(4, 0))

        self.thumb_cv.bind(
            "<MouseWheel>",
            lambda e: self.thumb_cv.xview_scroll(
                -1 * (e.delta // 120), "units"))

    # ==================================================================
    # PDF LIST MANAGEMENT
    # ==================================================================

    def _add_pdf(self):
        paths = filedialog.askopenfilenames(
            title="Add PDF(s)",
            filetypes=[("PDF", "*.pdf")])
        if not paths:
            return
        added = 0
        for p in paths:
            if any(d["path"] == p for d in self._pdf_list):
                continue
            try:
                doc = fitz.open(p)
                self._pdf_list.append({
                    "path": p,
                    "doc":  doc,
                    "name": Path(p).name,
                })
                added += 1
            except Exception as e:
                messagebox.showerror("Error",
                                      f"Could not open {Path(p).name}:\n{e}")
        if added and self._active_idx == -1:
            self._set_active_pdf(0)
        else:
            self._rebuild_pdf_list()

    def _remove_active_pdf(self):
        if self._active_idx < 0:
            return
        entry = self._pdf_list.pop(self._active_idx)
        try:
            entry["doc"].close()
        except Exception:
            pass
        self._active_idx = -1
        if self._pdf_list:
            self._set_active_pdf(0)
        else:
            self._thumb_imgs.clear()
            self.thumb_cv.delete("all")
            self.canvas.delete("all")
            self.canvas.create_text(
                300, 250,
                text="Add a PDF, then drag to select a region",
                font=("", 16), fill=G400, tags="ph")
            self.page_entry.delete(0, "end")
            self.page_entry.insert(0, "–")
            self.total_lbl.configure(text="/ –")
            self.btn_prev.configure(state="disabled")
            self.btn_next.configure(state="disabled")
            self._rebuild_pdf_list()

    def _set_active_pdf(self, idx: int):
        if idx < 0 or idx >= len(self._pdf_list):
            return
        self._active_idx = idx
        self._current_page = 0
        self._rebuild_pdf_list()
        self._render_thumbs()
        self._show_page(0)

    def _rebuild_pdf_list(self):
        for w in self._pdf_listbox.winfo_children():
            w.destroy()
        for i, entry in enumerate(self._pdf_list):
            is_active = (i == self._active_idx)
            bg     = "#DBEAFE" if is_active else WHITE
            border = BLUE if is_active else G200

            card = ctk.CTkFrame(self._pdf_listbox, fg_color=bg,
                                 border_color=border, border_width=2,
                                 corner_radius=8, height=36)
            card.pack(fill="x", pady=(0, 4))
            card.pack_propagate(False)

            name_lbl = ctk.CTkLabel(
                card, text=entry["name"],
                font=ctk.CTkFont(size=12),
                text_color=G900 if is_active else G700,
                anchor="w")
            name_lbl.pack(side="left", padx=10, fill="x", expand=True)

            pg_lbl = ctk.CTkLabel(
                card, text=f"{len(entry['doc'])}p",
                font=ctk.CTkFont(size=10), text_color=G400)
            pg_lbl.pack(side="right", padx=8)

            for widget in [card, name_lbl, pg_lbl]:
                widget.bind("<Button-1>",
                            lambda e, ii=i: self._set_active_pdf(ii))

    # ==================================================================
    # PAGE RENDERING
    # ==================================================================

    @property
    def _active_doc(self):
        if self._active_idx < 0 or self._active_idx >= len(self._pdf_list):
            return None
        return self._pdf_list[self._active_idx]["doc"]

    def _show_page(self, idx: int):
        doc = self._active_doc
        if doc is None or idx < 0 or idx >= len(doc):
            return
        self._current_page = idx
        self._render_page()
        total = len(doc)
        self.page_entry.delete(0, "end")
        self.page_entry.insert(0, str(idx + 1))
        self.total_lbl.configure(text=f"/ {total}")
        self.btn_prev.configure(
            state="normal" if idx > 0 else "disabled")
        self.btn_next.configure(
            state="normal" if idx < total - 1 else "disabled")
        self._hl_thumb(idx)

    def _render_page(self):
        doc = self._active_doc
        if doc is None:
            return
        self.canvas.update_idletasks()
        cw = max(self.canvas.winfo_width(),  300)
        ch = max(self.canvas.winfo_height(), 300)

        page  = doc[self._current_page]
        pw    = page.rect.width
        scale = max((cw - 40) / pw, 0.05)

        mat              = fitz.Matrix(scale, scale)
        self._render_mat = mat
        self._inv_mat    = ~mat

        pix = page.get_pixmap(matrix=mat, alpha=False)
        self._preview_img = _pix_to_tk(pix)
        iw = self._preview_img.width()
        ih = self._preview_img.height()

        ox = (cw - iw) / 2
        oy = 20.0
        self._page_ox = ox
        self._page_oy = oy
        self._page_iw = float(iw)
        self._page_ih = float(ih)

        self.canvas.delete("all")
        # Drop shadow
        self.canvas.create_rectangle(
            ox + 3, oy + 3, ox + iw + 3, oy + ih + 3,
            fill=G300, outline="")
        self.canvas.create_image(ox, oy, anchor="nw",
                                  image=self._preview_img)
        self.canvas.create_rectangle(
            ox - 1, oy - 1, ox + iw + 1, oy + ih + 1,
            outline=G300, width=1)
        self.canvas.configure(
            scrollregion=(0, 0, cw, max(ih + 60, ch)))

    def _on_canvas_resize(self, _event=None):
        if self._active_doc:
            self._render_page()

    def _prev_page(self):
        if self._current_page > 0:
            self._show_page(self._current_page - 1)

    def _next_page(self):
        doc = self._active_doc
        if doc and self._current_page < len(doc) - 1:
            self._show_page(self._current_page + 1)

    def _goto_page(self, _event=None):
        doc = self._active_doc
        if not doc:
            return
        try:
            n = int(self.page_entry.get())
            if 1 <= n <= len(doc):
                self._show_page(n - 1)
        except ValueError:
            pass

    # ==================================================================
    # THUMBNAIL STRIP
    # ==================================================================

    def _render_thumbs(self):
        self.thumb_cv.delete("all")
        self._thumb_imgs.clear()
        self._thumb_positions.clear()
        self._thumb_render_next = 0
        self._highlighted_thumb = -1
        doc = self._active_doc
        if doc is None:
            return
        # Use a fixed placeholder size for all slots; images fill in lazily
        ph_w = self.THUMB_W
        ph_h = int(self.THUMB_W * 1.4)
        x, sp = 14, 10
        total = len(doc)
        for i in range(total):
            self._thumb_imgs.append(None)
            # Gray placeholder rectangle
            self.thumb_cv.create_rectangle(
                x, 6, x + ph_w, 6 + ph_h,
                fill=G200, outline=G300, width=1, tags=f"r_{i}")
            self.thumb_cv.create_text(
                x + ph_w // 2, 6 + ph_h + 6,
                text=str(i + 1), font=("", 9), fill=G500, tags=f"l_{i}")
            for tag in (f"r_{i}", f"l_{i}"):
                self.thumb_cv.tag_bind(
                    tag, "<Button-1>",
                    lambda e, ii=i: self._show_page(ii))
            self._thumb_positions.append((x, ph_w, ph_h))
            x += ph_w + sp
        # Set scrollregion immediately so strip is usable right away
        total_w = x
        self.thumb_cv.configure(
            scrollregion=(0, 0, total_w, ph_h + 30))
        # Kick off lazy image loading
        self._render_thumb_batch()

    def _render_thumb_batch(self, batch: int = 8):
        """Render a small batch of thumbnails, then schedule the next batch."""
        doc = self._active_doc
        if doc is None:
            return
        start = self._thumb_render_next
        end   = min(start + batch, len(doc))
        for i in range(start, end):
            if self._thumb_imgs[i] is not None:
                continue
            img = _render(doc, i, self.THUMB_W)
            self._thumb_imgs[i] = img
            x, _, _ = self._thumb_positions[i]
            iw, ih = img.width(), img.height()
            # Replace placeholder with real image
            self.thumb_cv.delete(f"r_{i}")
            self.thumb_cv.create_rectangle(
                x - 3, 4, x + iw + 3, ih + 8,
                outline=G300, width=1, tags=f"r_{i}")
            self.thumb_cv.create_image(
                x, 6, anchor="nw", image=img, tags=f"t_{i}")
            self.thumb_cv.tag_bind(
                f"t_{i}", "<Button-1>",
                lambda e, ii=i: self._show_page(ii))
            # Re-apply highlight if this is the active page
            if i == self._highlighted_thumb:
                self.thumb_cv.itemconfig(f"r_{i}",
                                          outline=BLUE, width=3)
        self._thumb_render_next = end
        if end < len(doc):
            self.parent.after(0, self._render_thumb_batch)
        else:
            # Final scrollregion update with real image sizes
            bb = self.thumb_cv.bbox("all")
            if bb:
                self.thumb_cv.configure(scrollregion=bb)

    def _hl_thumb(self, idx: int):
        # Clear old highlight in O(1)
        old = self._highlighted_thumb
        if old >= 0:
            self.thumb_cv.itemconfig(f"r_{old}", outline=G300, width=1)
        # Set new highlight in O(1)
        self._highlighted_thumb = idx
        if idx >= 0:
            self.thumb_cv.itemconfig(f"r_{idx}", outline=BLUE, width=3)

    # ==================================================================
    # RUBBER-BAND SELECTION
    # ==================================================================

    def _point_on_page(self, cx: float, cy: float) -> bool:
        return (self._page_ox <= cx <= self._page_ox + self._page_iw and
                self._page_oy <= cy <= self._page_oy + self._page_ih)

    def _canvas_to_pdf_rect(self, x0, y0, x1, y1):
        """Convert clamped canvas pixel coords to a normalised fitz.Rect
        in PDF page coordinate space."""
        p0 = fitz.Point(x0 - self._page_ox,
                         y0 - self._page_oy) * self._inv_mat
        p1 = fitz.Point(x1 - self._page_ox,
                         y1 - self._page_oy) * self._inv_mat
        r = fitz.Rect(p0, p1)
        r.normalize()
        return r

    def _on_rb_start(self, event):
        if self._active_doc is None:
            return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        if not self._point_on_page(cx, cy):
            return
        self._rb_start = (cx, cy)
        self._selecting = True
        if self._rb_rect_id is not None:
            self.canvas.delete(self._rb_rect_id)
            self._rb_rect_id = None

    def _on_rb_drag(self, event):
        if not self._selecting or self._rb_start is None:
            return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        sx, sy = self._rb_start
        if self._rb_rect_id is not None:
            self.canvas.delete(self._rb_rect_id)
        self._rb_rect_id = self.canvas.create_rectangle(
            sx, sy, cx, cy,
            outline=SEL_BLUE, width=2, dash=(6, 3),
            tags="rubber_band")

    def _on_rb_release(self, event):
        if not self._selecting or self._rb_start is None:
            return
        self._selecting = False
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        sx, sy = self._rb_start
        self._rb_start = None

        if self._rb_rect_id is not None:
            self.canvas.delete(self._rb_rect_id)
            self._rb_rect_id = None

        # Reject tiny drags (accidental clicks)
        if abs(cx - sx) < 8 or abs(cy - sy) < 8:
            return

        # Clamp to the rendered page image boundaries
        x0 = max(min(sx, cx), self._page_ox)
        y0 = max(min(sy, cy), self._page_oy)
        x1 = min(max(sx, cx), self._page_ox + self._page_iw)
        y1 = min(max(sy, cy), self._page_oy + self._page_ih)
        if x1 <= x0 or y1 <= y0:
            return

        crop_rect = self._canvas_to_pdf_rect(x0, y0, x1, y1)
        if crop_rect.is_empty or crop_rect.width < 2 or crop_rect.height < 2:
            return

        self._do_capture(crop_rect)
        self._flash_feedback(x0, y0, x1, y1)

    def _flash_feedback(self, x0, y0, x1, y1,
                         flashes: int = 3, interval_ms: int = 80):
        """Brief green overlay on the captured region."""
        fid = self.canvas.create_rectangle(
            x0, y0, x1, y1,
            outline=GREEN, fill=GREEN,
            stipple="gray25", width=2, tags="flash")
        self.canvas.after(interval_ms * flashes,
                           lambda: self.canvas.delete(fid))

    # ==================================================================
    # CAPTURE PIPELINE
    # ==================================================================

    def _make_snippet_thumbnail(self, snip: Snippet):
        """Render just the crop region as a small PhotoImage."""
        try:
            src_doc  = fitz.open(snip.source_path)
            src_page = src_doc[snip.page_index]
            clip     = snip.crop_rect
            scale    = self.SNIP_W / max(clip.width, 1)
            mat      = fitz.Matrix(scale, scale)
            pix      = src_page.get_pixmap(matrix=mat, clip=clip, alpha=False)
            src_doc.close()
            return _pix_to_tk(pix)
        except Exception:
            return None

    def _append_to_output(self, snip: Snippet):
        """Append snippet crop as a new page in self._out_doc.
        Uses show_pdf_page with clip= to preserve the text layer."""
        try:
            src_doc  = fitz.open(snip.source_path)
            clip     = snip.crop_rect
            new_page = self._out_doc.new_page(
                width=clip.width, height=clip.height)
            new_page.show_pdf_page(
                new_page.rect,
                src_doc,
                snip.page_index,
                clip=clip,
            )
            src_doc.close()
        except Exception as e:
            messagebox.showerror("Capture Error",
                                  f"Failed to capture region:\n{e}")

    def _do_capture(self, crop_rect):
        doc_entry = self._pdf_list[self._active_idx]
        label = (f"{doc_entry['name']}  p.{self._current_page + 1}  "
                 f"({crop_rect.width:.0f}\u00d7{crop_rect.height:.0f} pt)")
        snip = Snippet(
            source_path=doc_entry["path"],
            page_index=self._current_page,
            crop_rect=crop_rect,
            label=label,
        )
        snip.thumbnail = self._make_snippet_thumbnail(snip)
        self._append_to_output(snip)
        self._snippets.append(snip)
        self._rebuild_snippet_list()
        self._status_lbl.configure(
            text=f"{len(self._snippets)} snippet(s) captured")

    # ==================================================================
    # SNIPPET LIST UI
    # ==================================================================

    def _rebuild_snippet_list(self):
        for w in self._snip_frame.winfo_children():
            w.destroy()
        self._snip_count_lbl.configure(text=f"({len(self._snippets)})")
        for i, snip in enumerate(self._snippets):
            self._make_snippet_card(i, snip)

    def _make_snippet_card(self, idx: int, snip: Snippet):
        card = ctk.CTkFrame(self._snip_frame, fg_color=WHITE,
                             border_color=G200, border_width=1,
                             corner_radius=8)
        card.pack(fill="x", pady=(0, 5))

        # Thumbnail
        if snip.thumbnail:
            tw = snip.thumbnail.width()
            th = snip.thumbnail.height()
            thumb_cv = tk.Canvas(card, width=tw, height=th,
                                  bg=WHITE, highlightthickness=0)
            thumb_cv.pack(side="left", padx=(6, 4), pady=6)
            thumb_cv.create_image(0, 0, anchor="nw", image=snip.thumbnail)

        # Label
        ctk.CTkLabel(card, text=snip.label,
                      font=ctk.CTkFont(size=10),
                      text_color=G700, anchor="w",
                      wraplength=170).pack(
            side="left", fill="x", expand=True, padx=(0, 2))

        # Reorder / delete buttons
        btn_col = ctk.CTkFrame(card, fg_color="transparent")
        btn_col.pack(side="right", padx=4, pady=4)

        ctk.CTkButton(
            btn_col, text="↑", width=26, height=22,
            fg_color="transparent", hover_color=G100,
            text_color=G500, font=ctk.CTkFont(size=13),
            command=lambda ii=idx: self._move_snippet(ii, -1),
        ).pack(pady=(0, 2))
        ctk.CTkButton(
            btn_col, text="↓", width=26, height=22,
            fg_color="transparent", hover_color=G100,
            text_color=G500, font=ctk.CTkFont(size=13),
            command=lambda ii=idx: self._move_snippet(ii, +1),
        ).pack(pady=(0, 2))
        ctk.CTkButton(
            btn_col, text="🗑", width=26, height=22,
            fg_color="transparent", hover_color="#FEE2E2",
            text_color=RED, font=ctk.CTkFont(size=13),
            command=lambda ii=idx: self._delete_snippet(ii),
        ).pack()

    def _delete_snippet(self, idx: int):
        if 0 <= idx < len(self._snippets):
            self._snippets.pop(idx)
            self._rebuild_output_doc()
            self._rebuild_snippet_list()
            self._status_lbl.configure(
                text=f"{len(self._snippets)} snippet(s) captured")

    def _move_snippet(self, idx: int, direction: int):
        new_idx = idx + direction
        if 0 <= new_idx < len(self._snippets):
            self._snippets[idx], self._snippets[new_idx] = (
                self._snippets[new_idx], self._snippets[idx])
            self._rebuild_output_doc()
            self._rebuild_snippet_list()

    # ==================================================================
    # OUTPUT DOCUMENT
    # ==================================================================

    def _rebuild_output_doc(self):
        """Reconstruct _out_doc from _snippets in current order."""
        try:
            self._out_doc.close()
        except Exception:
            pass
        self._out_doc = fitz.open()
        for snip in self._snippets:
            self._append_to_output(snip)

    def _save_excerpt(self):
        if self._out_doc.page_count == 0:
            messagebox.showwarning("Empty",
                                    "No snippets captured yet.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="excerpt.pdf",
        )
        if not path:
            return
        try:
            self._out_doc.save(path)
            messagebox.showinfo(
                "Saved",
                f"Excerpt PDF saved:\n{path}\n"
                f"({self._out_doc.page_count} pages)")
            self._status_lbl.configure(text="Saved.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save:\n{e}")

    # ==================================================================
    # CLEANUP
    # ==================================================================

    def cleanup(self):
        for entry in self._pdf_list:
            try:
                entry["doc"].close()
            except Exception:
                pass
        try:
            self._out_doc.close()
        except Exception:
            pass
