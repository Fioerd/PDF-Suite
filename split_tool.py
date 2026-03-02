"""Split Tool – PDF splitting with page preview and cut lines.

This module is loaded by main.py when the user clicks "Split".
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

import customtkinter as ctk

try:
    import fitz  # pymupdf
except ImportError:
    fitz = None

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import RectangleObject
except ImportError:
    PdfReader = PdfWriter = RectangleObject = None

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = ImageTk = None

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------
BLUE       = "#3B82F6"
BLUE_HOVER = "#2563EB"
GREEN      = "#16A34A"
GREEN_HOVER= "#15803D"
GREEN_TXT  = "#16A34A"
RED        = "#EF4444"
G100       = "#F3F4F6"
G200       = "#E5E7EB"
G300       = "#D1D5DB"
G400       = "#9CA3AF"
G500       = "#6B7280"
G700       = "#374151"
G900       = "#111827"
WHITE      = "#FFFFFF"
CUT_ACTIVE = "#DC2626"
CUT_INACTIVE = "#CBD5E1"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _pix_to_tk(pixmap):
    if Image and ImageTk:
        img = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        return ImageTk.PhotoImage(img)
    return tk.PhotoImage(data=pixmap.tobytes("ppm"))


def _render(doc, idx, max_w):
    page = doc[idx]
    s = max_w / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
    return _pix_to_tk(pix)


# ═══════════════════════════════════════════════════════════════════════════
# Split Tool  (builds UI inside a parent frame)
# ═══════════════════════════════════════════════════════════════════════════

class SplitTool:
    THUMB_W = 80

    def __init__(self, parent: ctk.CTkFrame):
        self.parent = parent

        # Check deps
        if fitz is None or PdfReader is None:
            ctk.CTkLabel(
                parent,
                text="⚠  Missing dependencies.\n\n"
                     "Install them with:\n"
                     "  pip install pymupdf pypdf Pillow",
                font=ctk.CTkFont(size=16), text_color=G500,
            ).pack(expand=True)
            return

        # State
        self.pdf_path = ""
        self.output_dir = ""
        self.total_pages = 0
        self.current_page = 0
        self.doc = None
        self.ranges: list[tuple[int, int]] = []
        self.page_cuts: dict[int, float] = {}
        self._cut_ratio = 1.0
        self._dragging = False
        self._img_t = self._img_b = self._img_l = self._img_r = 0.0
        self._preview_img = None
        self._thumb_imgs: list = []

        self._build_ui()

    # ==================================================================
    # BUILD UI
    # ==================================================================

    def _build_ui(self):
        main = ctk.CTkFrame(self.parent, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=24, pady=(12, 0))
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        # ======== LEFT PANEL ========
        left = ctk.CTkFrame(main, width=440, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 24))
        left.grid_propagate(False)

        ctk.CTkLabel(left, text="Split PDF", font=("", 26, "bold"),
                      text_color=G900).pack(anchor="w", pady=(0, 18))

        # File
        ctk.CTkLabel(left, text="File", font=("", 15, "bold"),
                      text_color=G900).pack(anchor="w")
        file_row = ctk.CTkFrame(left, fg_color="transparent")
        file_row.pack(fill="x", pady=(4, 14))
        self.file_entry = ctk.CTkEntry(file_row, state="readonly",
                                        fg_color=G100, border_color=G200,
                                        text_color=G700, height=38)
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(file_row, text="📁  Browse", width=110, height=38,
                       fg_color=BLUE, hover_color=BLUE_HOVER,
                       command=self._pick_pdf).pack(side="right")

        # Output Folder
        ctk.CTkLabel(left, text="Output Folder", font=("", 15, "bold"),
                      text_color=G900).pack(anchor="w")
        out_row = ctk.CTkFrame(left, fg_color="transparent")
        out_row.pack(fill="x", pady=(4, 14))
        self.out_entry = ctk.CTkEntry(out_row, state="readonly",
                                       fg_color=G100, border_color=G200,
                                       text_color=G700, height=38)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(out_row, text="📁  Browse", width=110, height=38,
                       fg_color=BLUE, hover_color=BLUE_HOVER,
                       command=self._pick_output).pack(side="right")

        # Page Ranges
        ctk.CTkLabel(left, text="Page Ranges (Parts)", font=("", 15, "bold"),
                      text_color=G900).pack(anchor="w", pady=(0, 4))

        self.ranges_frame = ctk.CTkScrollableFrame(
            left, fg_color="transparent", height=180,
            scrollbar_button_color=G300,
            scrollbar_button_hover_color=G400,
        )
        self.ranges_frame.pack(fill="x", pady=(0, 10))

        # From / To / buttons
        input_grid = ctk.CTkFrame(left, fg_color="transparent")
        input_grid.pack(fill="x", pady=(0, 14))
        input_grid.columnconfigure(0, weight=1)
        input_grid.columnconfigure(1, weight=1)

        ctk.CTkLabel(input_grid, text="From:", text_color=G700,
                      font=("", 13)).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(input_grid, text="To:", text_color=G700,
                      font=("", 13)).grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.from_entry = ctk.CTkEntry(input_grid, height=38,
                                        fg_color=G100, border_color=G200,
                                        text_color=G900)
        self.from_entry.grid(row=1, column=0, sticky="ew", pady=(2, 0))
        self.from_entry.insert(0, "1")

        self.to_entry = ctk.CTkEntry(input_grid, height=38,
                                      fg_color=G100, border_color=G200,
                                      text_color=G900)
        self.to_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(2, 0))
        self.to_entry.insert(0, "1")

        btn_col = ctk.CTkFrame(input_grid, fg_color="transparent")
        btn_col.grid(row=0, column=2, rowspan=2, sticky="ns", padx=(12, 0))

        ctk.CTkButton(btn_col, text="Current Page", width=120, height=34,
                       fg_color="transparent", border_color=G300,
                       border_width=1, text_color=G700,
                       hover_color=G100,
                       command=self._page_to_bis).pack(pady=(0, 4))
        ctk.CTkButton(btn_col, text="Add Range", width=120, height=34,
                       fg_color=GREEN, hover_color=GREEN_HOVER,
                       command=self._add_range).pack()

        # Split button
        self.split_btn = ctk.CTkButton(
            left, text="Split PDF", height=44,
            font=("", 15, "bold"), fg_color=BLUE, hover_color=BLUE_HOVER,
            command=self._split_pdf,
        )
        self.split_btn.pack(fill="x", pady=(4, 8))

        # Progress
        self.progress = ctk.CTkProgressBar(left, progress_color=GREEN,
                                            fg_color=G200, height=8)
        self.progress.pack(fill="x")
        self.progress.set(0)

        # Status
        self.status_lbl = ctk.CTkLabel(left, text="", text_color=GREEN_TXT,
                                        font=("", 13), anchor="w")
        self.status_lbl.pack(anchor="w", pady=(6, 0))

        # ======== RIGHT PANEL (Preview) ========
        right = ctk.CTkFrame(main, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        self.preview = tk.Canvas(right, bg=WHITE, highlightthickness=0,
                                  relief="flat")
        self.preview.grid(row=0, column=0, sticky="nsew")
        self.preview.create_text(
            200, 200,
            text="Load a PDF to see\npage preview here",
            font=("", 16), fill=G400, tags="ph",
        )
        self.preview.bind("<B1-Motion>", self._do_drag)
        self.preview.bind("<ButtonRelease-1>", self._end_drag)

        # Navigation
        nav = ctk.CTkFrame(right, fg_color="transparent")
        nav.grid(row=1, column=0, pady=(8, 0))

        self.btn_prev = ctk.CTkButton(
            nav, text="←  Back", width=100, height=36,
            fg_color=WHITE, border_color=G300, border_width=1,
            text_color=G700, hover_color=G100, state="disabled",
            command=self._prev,
        )
        self.btn_prev.pack(side="left", padx=6)

        self.page_lbl = ctk.CTkLabel(nav, text="Page – / –", width=130,
                                      font=("", 14), text_color=G700)
        self.page_lbl.pack(side="left", padx=6)

        self.btn_next = ctk.CTkButton(
            nav, text="Next  →", width=100, height=36,
            fg_color=WHITE, border_color=G300, border_width=1,
            text_color=G700, hover_color=G100, state="disabled",
            command=self._next,
        )
        self.btn_next.pack(side="left", padx=6)

        # Cut info
        self.cut_lbl = ctk.CTkLabel(right, text="", text_color=G500,
                                     font=("", 11))
        self.cut_lbl.grid(row=2, column=0, pady=(4, 0))

        # ======== BOTTOM – Thumbnails ========
        bot = ctk.CTkFrame(self.parent, fg_color="transparent", height=155)
        bot.pack(fill="x", side="bottom", padx=24, pady=(4, 14))
        bot.pack_propagate(False)

        self.btn_tl = ctk.CTkButton(
            bot, text="‹", width=28, height=100, font=("", 22),
            fg_color=G100, hover_color=G200, text_color=G500, corner_radius=8,
            command=lambda: self.thumb_cv.xview_scroll(-3, "units"),
        )
        self.btn_tl.pack(side="left", padx=(0, 6))

        self.thumb_cv = tk.Canvas(bot, height=140, bg=G100,
                                   highlightthickness=0, relief="flat")
        self.thumb_cv.pack(side="left", fill="both", expand=True)

        self.btn_tr = ctk.CTkButton(
            bot, text="›", width=28, height=100, font=("", 22),
            fg_color=G100, hover_color=G200, text_color=G500, corner_radius=8,
            command=lambda: self.thumb_cv.xview_scroll(3, "units"),
        )
        self.btn_tr.pack(side="right", padx=(6, 0))

        self.thumb_cv.bind(
            "<MouseWheel>",
            lambda e: self.thumb_cv.xview_scroll(-1 * (e.delta // 120), "units"),
        )

        # Keyboard nav
        self.parent.winfo_toplevel().bind("<Left>",  lambda e: self._prev())
        self.parent.winfo_toplevel().bind("<Right>", lambda e: self._next())

    # ==================================================================
    # FILE LOADING
    # ==================================================================

    def _pick_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not p:
            return
        self.pdf_path = p
        self.file_entry.configure(state="normal")
        self.file_entry.delete(0, "end")
        self.file_entry.insert(0, p)
        self.file_entry.configure(state="readonly")
        self._load_pdf()

    def _pick_output(self):
        p = filedialog.askdirectory()
        if not p:
            return
        self.output_dir = p
        self.out_entry.configure(state="normal")
        self.out_entry.delete(0, "end")
        self.out_entry.insert(0, p)
        self.out_entry.configure(state="readonly")

    def _load_pdf(self):
        try:
            if self.doc:
                self.doc.close()
            self.doc = fitz.open(self.pdf_path)
            self.total_pages = len(self.doc)
            self.current_page = 0
            self.ranges.clear()
            self.page_cuts.clear()
            self._cut_ratio = 1.0

            self.from_entry.delete(0, "end")
            self.from_entry.insert(0, "1")
            self.to_entry.delete(0, "end")
            self.to_entry.insert(0, str(self.total_pages))

            self.status_lbl.configure(text=f"{self.total_pages} pages loaded.")
            self._rebuild_cards()
            self._render_thumbs()
            self._show(0)
        except Exception as e:
            self.total_pages = 0
            messagebox.showerror("Error", f"Could not load PDF:\n{e}")

    # ==================================================================
    # PREVIEW
    # ==================================================================

    def _show(self, idx):
        if not self.doc or idx < 0 or idx >= self.total_pages:
            return
        self.current_page = idx

        self.preview.update_idletasks()
        cw = max(self.preview.winfo_width(), 300)
        ch = max(self.preview.winfo_height(), 300)

        page = self.doc[idx]
        ratio = page.rect.height / page.rect.width
        fw, fh = cw - 30, ch - 30
        rw = fw if fw * ratio <= fh else int(fh / ratio)
        rw = max(rw, 100)

        self._preview_img = _render(self.doc, idx, rw)
        iw, ih = self._preview_img.width(), self._preview_img.height()
        cx, cy = cw // 2, ch // 2
        self._img_l = cx - iw / 2
        self._img_r = cx + iw / 2
        self._img_t = cy - ih / 2
        self._img_b = cy + ih / 2

        self.preview.delete("all")
        self.preview.create_image(cx, cy, anchor="center",
                                   image=self._preview_img)

        self._cut_ratio = self.page_cuts.get(idx, 1.0)
        self._draw_cut()

        self.page_lbl.configure(text=f"Page {idx + 1} / {self.total_pages}")
        self.btn_prev.configure(
            state="normal" if idx > 0 else "disabled")
        self.btn_next.configure(
            state="normal" if idx < self.total_pages - 1 else "disabled")
        self._hl_thumb(idx)

    def _prev(self):
        if self.current_page > 0:
            self._show(self.current_page - 1)

    def _next(self):
        if self.current_page < self.total_pages - 1:
            self._show(self.current_page + 1)

    # ==================================================================
    # CUT LINE
    # ==================================================================

    def _draw_cut(self):
        if self._img_b <= self._img_t:
            return
        has = self._cut_ratio < 0.98
        col = CUT_ACTIVE if has else CUT_INACTIVE
        w = 3 if has else 2
        dash = () if has else (6, 4)
        ih = self._img_b - self._img_t
        y = self._img_t + self._cut_ratio * ih

        self.preview.create_rectangle(
            self._img_l, y - 10, self._img_r, y + 10,
            fill="", outline="", tags="cut_hit")
        self.preview.create_line(
            self._img_l, y, self._img_r, y,
            fill=col, width=w, dash=dash, tags="cut_vis")

        hs = 9
        for hx, d in [(self._img_l, 1), (self._img_r, -1)]:
            self.preview.create_polygon(
                hx, y - hs, hx + d * hs, y, hx, y + hs,
                fill=col, outline="", tags="cut_h")

        txt = f"✂ {int(self._cut_ratio * 100)}%" if has else "▼ full page"
        self.preview.create_text(
            self._img_r + 10, y, text=txt, anchor="w",
            font=("", 10), fill=col, tags="cut_lbl")

        if has:
            bx = self._img_l - 6
            self.preview.create_rectangle(
                bx, self._img_t, bx + 4, y,
                fill=BLUE, outline="", tags="cut_z")
            self.preview.create_rectangle(
                bx, y, bx + 4, self._img_b,
                fill="#F97316", outline="", tags="cut_z")

        self.cut_lbl.configure(
            text="✂ Above → current part | Below → next part" if has
            else "Drag the line up to cut this page")

        for t in ("cut_hit", "cut_vis", "cut_h", "cut_lbl"):
            self.preview.tag_bind(t, "<Button-1>", self._start_drag)
        for t in ("cut_hit", "cut_vis", "cut_h"):
            self.preview.tag_bind(
                t, "<Enter>",
                lambda e: self.preview.config(cursor="sb_v_double_arrow"))
            self.preview.tag_bind(
                t, "<Leave>", lambda e: self.preview.config(cursor=""))

    def _start_drag(self, e):
        self._dragging = True

    def _do_drag(self, e):
        if not self._dragging or self._img_b <= self._img_t:
            return
        y = max(self._img_t + 5, min(e.y, self._img_b))
        self._cut_ratio = (y - self._img_t) / (self._img_b - self._img_t)
        if self._cut_ratio >= 0.98:
            self._cut_ratio = 1.0
        if self._cut_ratio < 0.98:
            self.page_cuts[self.current_page] = self._cut_ratio
        else:
            self.page_cuts.pop(self.current_page, None)
        for t in ("cut_hit", "cut_vis", "cut_h", "cut_lbl", "cut_z"):
            self.preview.delete(t)
        self._draw_cut()

    def _end_drag(self, e):
        if not self._dragging:
            return
        self._dragging = False
        if self._cut_ratio < 0.98:
            self.page_cuts[self.current_page] = self._cut_ratio
        else:
            self.page_cuts.pop(self.current_page, None)
        self._rebuild_cards()

    # ==================================================================
    # THUMBNAILS
    # ==================================================================

    def _render_thumbs(self):
        self.thumb_cv.delete("all")
        self._thumb_imgs.clear()
        if not self.doc:
            return
        x = 14
        sp = 10
        for i in range(self.total_pages):
            img = _render(self.doc, i, self.THUMB_W)
            self._thumb_imgs.append(img)
            self.thumb_cv.create_rectangle(
                x - 3, 4, x + img.width() + 3, 6 + img.height(),
                outline=G300, width=1, tags=f"r_{i}")
            self.thumb_cv.create_image(
                x, 6, anchor="nw", image=img, tags=f"t_{i}")
            self.thumb_cv.create_text(
                x + img.width() // 2, 10 + img.height(),
                text=str(i + 1), font=("", 9), fill=G500, tags=f"l_{i}")
            for tag in (f"t_{i}", f"r_{i}", f"l_{i}"):
                self.thumb_cv.tag_bind(
                    tag, "<Button-1>",
                    lambda e, ii=i: self._show(ii))
            x += img.width() + sp
            if i % 10 == 0:
                self.parent.update_idletasks()
        self.thumb_cv.config(scrollregion=self.thumb_cv.bbox("all"))

    def _hl_thumb(self, idx):
        for i in range(self.total_pages):
            c = BLUE if i == idx else G300
            w = 3 if i == idx else 1
            self.thumb_cv.itemconfig(f"r_{i}", outline=c, width=w)

    # ==================================================================
    # RANGE MANAGEMENT
    # ==================================================================

    def _add_range(self):
        try:
            s, e = int(self.from_entry.get()), int(self.to_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter valid page numbers.")
            return
        if s <= 0 or e <= 0:
            messagebox.showerror("Error", "Page numbers must be > 0.")
            return
        if s > e:
            messagebox.showerror("Error", "Start page must be ≤ end page.")
            return
        if e > self.total_pages:
            messagebox.showerror(
                "Error",
                f"Page {e} doesn't exist. PDF has {self.total_pages} pages.")
            return

        self.ranges.append((s, e))
        self._rebuild_cards()

        self.from_entry.delete(0, "end")
        self.from_entry.insert(0, str(e))
        self.to_entry.delete(0, "end")
        self.to_entry.insert(0, str(self.total_pages))

    def _delete_range(self, idx):
        if 0 <= idx < len(self.ranges):
            self.ranges.pop(idx)
            self._rebuild_cards()

    def _rebuild_cards(self):
        for w in self.ranges_frame.winfo_children():
            w.destroy()
        for i, (s, e) in enumerate(self.ranges):
            txt = (f"Part {i + 1}: Pages {s}-{e}" if s != e
                   else f"Part {i + 1}: Page {s}")

            cuts = []
            if (e - 1) in self.page_cuts and s != e:
                p = int(self.page_cuts[e - 1] * 100)
                cuts.append(f"p.{e} ✂{p}%↑")
            if (s - 1) in self.page_cuts and s != e:
                p = int(self.page_cuts[s - 1] * 100)
                cuts.append(f"p.{s} ✂{p}%↓")
            if cuts:
                txt += f"  ({', '.join(cuts)})"

            card = ctk.CTkFrame(self.ranges_frame, fg_color=WHITE,
                                 border_color=G300, border_width=1,
                                 corner_radius=10, height=44)
            card.pack(fill="x", pady=(0, 6))
            card.pack_propagate(False)

            ctk.CTkLabel(card, text=txt, text_color=G900,
                          font=("", 13), anchor="w").pack(
                side="left", padx=14, fill="x", expand=True)

            ctk.CTkButton(
                card, text="🗑", width=34, height=34,
                fg_color="transparent", hover_color=G100,
                text_color=G400, font=("", 16),
                command=lambda ii=i: self._delete_range(ii),
            ).pack(side="right", padx=6)

    def _page_to_bis(self):
        if self.total_pages == 0:
            return
        self.to_entry.delete(0, "end")
        self.to_entry.insert(0, str(self.current_page + 1))

    # ==================================================================
    # SPLIT
    # ==================================================================

    def _split_pdf(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return
        if not self.output_dir:
            messagebox.showerror("Error", "Please select an output folder.")
            return
        if not self.ranges:
            messagebox.showerror("Error", "Please add at least one page range.")
            return

        self.split_btn.configure(state="disabled")
        try:
            self._do_split()
        except Exception as ex:
            messagebox.showerror("Error", f"Split failed:\n{ex}")
            self.status_lbl.configure(text="Error during split.")
        finally:
            self.split_btn.configure(state="normal")

    def _add_page(self, writer, page, crop="full", ratio=0.5):
        writer.add_page(page)
        if crop == "full":
            return
        a = writer.pages[-1]
        mb = a.mediabox
        pt, pb = float(mb.top), float(mb.bottom)
        pl, pr = float(mb.left), float(mb.right)
        cy = pt - ratio * (pt - pb)
        if crop == "top":
            a.cropbox = RectangleObject([pl, cy, pr, pt])
        elif crop == "bottom":
            a.cropbox = RectangleObject([pl, pb, pr, cy])

    def _do_split(self):
        reader = PdfReader(self.pdf_path)
        base = Path(self.pdf_path).stem
        n = len(self.ranges)
        self.progress.set(0)

        for i, (s, e) in enumerate(self.ranges):
            writer = PdfWriter()
            for pn in range(s - 1, e):
                pg = reader.pages[pn]
                if pn in self.page_cuts and s != e:
                    cr = self.page_cuts[pn]
                    first = pn == s - 1
                    last = pn == e - 1
                    if first and last:
                        self._add_page(writer, pg)
                    elif last:
                        self._add_page(writer, pg, "top", cr)
                    elif first:
                        self._add_page(writer, pg, "bottom", cr)
                    else:
                        self._add_page(writer, pg)
                else:
                    self._add_page(writer, pg)

            out = Path(self.output_dir) / f"{base}_part{i + 1}.pdf"
            with open(out, "wb") as f:
                writer.write(f)

            self.progress.set((i + 1) / n)
            self.status_lbl.configure(
                text=f"Part {i + 1} of {n} created (p. {s}–{e})...")
            self.parent.update_idletasks()

        self.status_lbl.configure(text=f"Done! {n} parts created.")

    # ==================================================================
    # CLEANUP
    # ==================================================================

    def cleanup(self):
        if self.doc:
            self.doc.close()
