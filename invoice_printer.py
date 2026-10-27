# -*- coding: utf-8 -*-
"""发票打印助手：选择 PDF，按每页 4 张或 2 张排版并打印。"""

from __future__ import annotations

import queue
import shutil
import sys
import tempfile
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path

import pymupdf
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

import win32con
import win32gui
import win32print
import win32ui

try:
    from PIL import Image, ImageWin
except ImportError as exc:  # pragma: no cover
    raise SystemExit("缺少 Pillow 依赖，请先执行：pip install -r requirements.txt") from exc

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "缺少 tkinterdnd2 依赖，请先执行：pip install -r requirements.txt"
    ) from exc


APP_TITLE = "发票打印助手"
LANDSCAPE_WIDTH = 841.89
LANDSCAPE_HEIGHT = 595.276
PAGE_MARGIN = 14.0
CELL_GAP = 10.0
MAX_PRINT_DPI = 300

LAYOUTS = {
    "4_landscape": {"label": "4 张（横）", "per_sheet": 4, "landscape": True},
    "2_landscape": {"label": "2 张（横）", "per_sheet": 2, "landscape": True},
    "2_portrait": {"label": "2 张（竖）", "per_sheet": 2, "landscape": False},
}
DEFAULT_LAYOUT = "4_landscape"

BG = "#f3f5f9"
CARD = "#ffffff"
TEXT = "#1f2937"
MUTED = "#6b7280"
BORDER = "#e3e8f0"
ACCENT = "#2563eb"
ACCENT_HOVER = "#1d4ed8"
ACCENT_DISABLED = "#b9caf0"
ROW_ALT = "#fafbfe"
SELECTED = "#e8f0ff"


@dataclass(frozen=True)
class InvoiceFile:
    path: Path
    pages: int


def read_invoice_file(path: Path) -> InvoiceFile:
    if path.suffix.lower() != ".pdf":
        raise ValueError("不是 PDF 文件")
    try:
        doc = pymupdf.open(str(path))
    except Exception as exc:
        raise ValueError(f"无法打开文件（{exc}）") from exc
    try:
        if doc.needs_pass:
            raise ValueError("PDF 受密码保护")
        if doc.page_count == 0:
            raise ValueError("PDF 中没有页面")
        return InvoiceFile(path=path, pages=doc.page_count)
    finally:
        doc.close()


def cell_rect(
    index: int, per_sheet: int, sheet_width: float, sheet_height: float
) -> pymupdf.Rect:
    if per_sheet == 2:
        cell_w = sheet_width - 2 * PAGE_MARGIN
        cell_h = (sheet_height - 2 * PAGE_MARGIN) / 2
        x0 = PAGE_MARGIN
        y0 = PAGE_MARGIN + index * cell_h
        return pymupdf.Rect(x0, y0, x0 + cell_w, y0 + cell_h)

    cell_w = (sheet_width - 2 * PAGE_MARGIN - CELL_GAP) / 2
    cell_h = (sheet_height - 2 * PAGE_MARGIN - CELL_GAP) / 2
    col = index % 2
    row = index // 2
    x0 = PAGE_MARGIN + col * (cell_w + CELL_GAP)
    y0 = PAGE_MARGIN + row * (cell_h + CELL_GAP)
    return pymupdf.Rect(x0, y0, x0 + cell_w, y0 + cell_h)


def compose_sheets(
    invoices: list[InvoiceFile],
    output_path: Path,
    progress,
    per_sheet: int = 4,
    landscape: bool = True,
) -> int:
    """把每页 PDF 按固定 2x2 或 1x2 位置排到 A4 页面，返回生成的 A4 页数。"""
    source_docs: list[pymupdf.Document] = []
    records: list[tuple[pymupdf.Document, int, Path]] = []
    out_doc = pymupdf.open()
    try:
        for item in invoices:
            src_doc = pymupdf.open(str(item.path))
            source_docs.append(src_doc)
            if src_doc.needs_pass:
                raise RuntimeError(f"“{item.path.name}”受密码保护，无法打印")
            for page_no in range(src_doc.page_count):
                page = src_doc[page_no]
                if page.rotation:
                    page.remove_rotation()
                records.append((src_doc, page_no, item.path))

        if not records:
            raise RuntimeError("所选 PDF 中没有可打印的页面")

        sheet_width, sheet_height = (
            (LANDSCAPE_WIDTH, LANDSCAPE_HEIGHT)
            if landscape
            else (LANDSCAPE_HEIGHT, LANDSCAPE_WIDTH)
        )
        cells = [
            cell_rect(index, per_sheet, sheet_width, sheet_height)
            for index in range(per_sheet)
        ]
        sheet = None
        total = len(records)
        for index, (src_doc, page_no, path) in enumerate(records):
            if index % per_sheet == 0:
                sheet = out_doc.new_page(width=sheet_width, height=sheet_height)
            try:
                sheet.show_pdf_page(cells[index % per_sheet], src_doc, page_no)
            except Exception as exc:
                raise RuntimeError(
                    f"排版“{path.name}”第 {page_no + 1} 页失败：{exc}"
                ) from exc
            progress(index + 1, total)

        out_doc.save(str(output_path), garbage=4, deflate=True)
        return out_doc.page_count
    finally:
        out_doc.close()
        for src_doc in source_docs:
            try:
                src_doc.close()
            except Exception:
                pass


def _create_printer_dc(printer_name: str, landscape: bool):
    hdc = win32gui.CreateDC("WINSPOOL", printer_name, None)
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            devmode = win32print.GetPrinter(hprinter, 2).get("pDevMode")
            if devmode:
                devmode.PaperSize = win32con.DMPAPER_A4
                devmode.Orientation = (
                    win32con.DMORIENT_LANDSCAPE
                    if landscape
                    else win32con.DMORIENT_PORTRAIT
                )
                devmode.Fields = (
                    devmode.Fields | win32con.DM_ORIENTATION | win32con.DM_PAPERSIZE
                )
                reset_hdc = win32gui.ResetDC(hdc, devmode)
                if reset_hdc:
                    hdc = reset_hdc
        finally:
            win32print.ClosePrinter(hprinter)
    except Exception:
        pass
    return win32ui.CreateDCFromHandle(hdc)


def send_to_printer(
    pdf_path: Path, printer_name: str, landscape: bool, progress
) -> int:
    """按打印机分辨率把 A4 页面绘制到打印设备，返回已发送页数。"""
    doc = pymupdf.open(str(pdf_path))
    dc = _create_printer_dc(printer_name, landscape)
    try:
        printable_w = int(dc.GetDeviceCaps(win32con.HORZRES))
        printable_h = int(dc.GetDeviceCaps(win32con.VERTRES))
        if printable_w <= 0 or printable_h <= 0:
            raise RuntimeError("无法获取打印机纸张尺寸")

        total = doc.page_count
        dc.StartDoc("发票打印")
        try:
            for index in range(total):
                page = doc[index]
                dc.StartPage()
                try:
                    ratio = page.rect.width / page.rect.height
                    if printable_w / printable_h > ratio:
                        dest_h = printable_h
                        dest_w = max(1, round(printable_h * ratio))
                    else:
                        dest_w = printable_w
                        dest_h = max(1, round(printable_w / ratio))
                    x0 = (printable_w - dest_w) // 2
                    y0 = (printable_h - dest_h) // 2

                    scale = min(
                        dest_w / page.rect.width,
                        dest_h / page.rect.height,
                        MAX_PRINT_DPI / 72.0,
                    )
                    pix = page.get_pixmap(
                        matrix=pymupdf.Matrix(scale, scale),
                        colorspace=pymupdf.csRGB,
                        alpha=False,
                    )
                    image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    dib = ImageWin.Dib(image)
                    dib.draw(dc.GetSafeHdc(), (x0, y0, x0 + dest_w, y0 + dest_h))
                finally:
                    dc.EndPage()
                progress(index + 1, total)
            dc.EndDoc()
        except Exception:
            try:
                dc.AbortDoc()
            except Exception:
                pass
            raise
        return total
    finally:
        doc.close()
        try:
            dc.DeleteDC()
        except Exception:
            pass


def available_font(root: tk.Misc) -> str:
    families = set(tkfont.families(root))
    for name in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI"):
        if name in families:
            return name
    return "TkDefaultFont"


class InvoicePrinterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font = available_font(root)
        self.files: list[InvoiceFile] = []
        self.busy = False
        self.closed = False
        self.layout_mode = DEFAULT_LAYOUT
        mode = LAYOUTS[self.layout_mode]
        self.per_sheet = mode["per_sheet"]
        self.landscape = mode["landscape"]
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()

        root.title(APP_TITLE)
        root.configure(bg=BG)
        root.minsize(700, 540)
        root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
        self._setup_dnd()
        self._refresh()
        self._place_window()
        self.root.after(80, self._drain_events)

    def _place_window(self) -> None:
        self.root.update_idletasks()
        width = 820
        height = 600
        x = max(0, (self.root.winfo_screenwidth() - width) // 2)
        y = max(0, (self.root.winfo_screenheight() - height) // 3)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Files.Treeview",
            background=CARD,
            fieldbackground=CARD,
            foreground=TEXT,
            borderwidth=0,
            rowheight=34,
            font=(self.font, 10),
        )
        style.configure(
            "Files.Treeview.Heading",
            background="#f7f9fc",
            foreground=MUTED,
            relief="flat",
            borderwidth=0,
            padding=(10, 9),
            font=(self.font, 9, "bold"),
        )
        style.map(
            "Files.Treeview",
            background=[("selected", SELECTED)],
            foreground=[("selected", TEXT)],
        )
        style.map("Files.Treeview.Heading", background=[("active", "#eef3fb")])
        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor="#e9edf5",
            background=ACCENT,
            bordercolor="#e9edf5",
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            thickness=6,
        )
        style.configure(
            "Vertical.TScrollbar",
            background="#d6dce8",
            troughcolor=CARD,
            bordercolor=CARD,
            arrowcolor="#8792a6",
            relief="flat",
        )

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True, padx=26, pady=22)

        header = tk.Frame(outer, bg=BG)
        header.pack(fill="x")
        tk.Label(
            header, text=APP_TITLE, bg=BG, fg=TEXT, font=(self.font, 20, "bold")
        ).pack(anchor="w")
        tk.Label(
            header,
            text="每个 PDF 页面作为一张 · 4 张 / 2 张、横版 / 竖版可切换 · 使用系统默认打印机",
            bg=BG,
            fg=MUTED,
            font=(self.font, 10),
        ).pack(anchor="w", pady=(6, 0))

        card = tk.Frame(outer, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True, pady=(18, 0))
        self.card = card

        list_wrap = tk.Frame(card, bg=CARD)
        list_wrap.pack(fill="both", expand=True, padx=12, pady=(12, 0))

        columns = ("name", "pages", "folder")
        self.tree = ttk.Treeview(
            list_wrap,
            columns=columns,
            show="headings",
            style="Files.Treeview",
            selectmode="browse",
            height=6,
        )
        self.tree.heading("name", text="文件名", anchor="w")
        self.tree.heading("pages", text="页数", anchor="center")
        self.tree.heading("folder", text="所在文件夹", anchor="w")
        self.tree.column("name", width=260, minwidth=160, anchor="w")
        self.tree.column("pages", width=64, minwidth=54, anchor="center", stretch=False)
        self.tree.column("folder", width=330, minwidth=120, anchor="w")
        self.tree.tag_configure("even", background=CARD)
        self.tree.tag_configure("odd", background=ROW_ALT)
        self.tree.bind("<Delete>", self._remove_selected)

        scrollbar = ttk.Scrollbar(list_wrap, orient="vertical", command=self.tree.yview)
        self.scrollbar = scrollbar
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.empty_label = tk.Label(
            list_wrap,
            text="还没有选择发票\n\n点击下方「选择 PDF 发票」添加文件\n也可以把 PDF 拖到这里（仅支持 PDF）",
            bg=CARD,
            fg="#9aa4b5",
            font=(self.font, 11),
            justify="center",
        )
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

        summary = tk.Frame(card, bg=CARD)
        summary.pack(fill="x", padx=14, pady=(8, 12))
        self.summary_label = tk.Label(
            summary, text="", bg=CARD, fg=MUTED, font=(self.font, 9)
        )
        self.summary_label.pack(side="left")
        self.hint_label = tk.Label(
            summary,
            text="可拖入 PDF · 选中后按 Delete 移除",
            bg=CARD,
            fg="#9aa4b5",
            font=(self.font, 9),
        )
        self.hint_label.pack(side="right")

        actions = tk.Frame(outer, bg=BG)
        actions.pack(fill="x", pady=(14, 0))
        self.select_btn = tk.Button(
            actions,
            text="选择 PDF 发票",
            command=self.choose_files,
            bg=CARD,
            fg=TEXT,
            activebackground="#eef3fb",
            activeforeground=TEXT,
            disabledforeground="#9aa4b5",
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=BORDER,
            font=(self.font, 11),
            padx=22,
            pady=9,
            cursor="hand2",
        )
        self.select_btn.pack(side="left")

        mode_box = tk.Frame(actions, bg=BG)
        mode_box.pack(side="left", padx=(18, 0))
        tk.Label(
            mode_box, text="每页：", bg=BG, fg=MUTED, font=(self.font, 9)
        ).pack(side="left")
        self.mode_buttons: dict[str, tk.Button] = {}
        for key, mode in LAYOUTS.items():
            button = tk.Button(
                mode_box,
                text=mode["label"],
                command=lambda value=key: self.set_layout_mode(value),
                bg=CARD,
                fg=TEXT,
                activebackground="#eef3fb",
                activeforeground=TEXT,
                disabledforeground="#9aa4b5",
                relief="flat",
                bd=0,
                highlightthickness=1,
                highlightbackground=BORDER,
                highlightcolor=BORDER,
                font=(self.font, 9),
                padx=9,
                pady=6,
                cursor="hand2",
            )
            button.pack(side="left", padx=(6, 0))
            self.mode_buttons[key] = button

        self.print_btn = tk.Button(
            actions,
            text="开始打印",
            command=self.start_print,
            bg=ACCENT,
            fg="#ffffff",
            activebackground=ACCENT_HOVER,
            activeforeground="#ffffff",
            disabledforeground="#f7faff",
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=(self.font, 11, "bold"),
            padx=38,
            pady=9,
            cursor="hand2",
        )
        self.print_btn.pack(side="right")

        self._bind_hover(self.select_btn, CARD, "#eef3fb")
        self._bind_hover(self.print_btn, ACCENT, ACCENT_HOVER)

        footer = tk.Frame(outer, bg=BG)
        footer.pack(fill="x", pady=(14, 0))
        try:
            printer_name = win32print.GetDefaultPrinter()
        except Exception:
            printer_name = None
        self.printer_label = tk.Label(
            footer,
            text=f"默认打印机：{printer_name}" if printer_name else "默认打印机：未检测到打印机",
            bg=BG,
            fg=MUTED,
            font=(self.font, 9),
            anchor="w",
        )
        self.printer_label.pack(side="left")
        self.status_label = tk.Label(
            footer, text="就绪", bg=BG, fg=MUTED, font=(self.font, 9), anchor="e"
        )
        self.status_label.pack(side="right")

        self.progress = ttk.Progressbar(
            outer, style="Accent.Horizontal.TProgressbar", mode="determinate", maximum=100
        )
        self.progress.pack(fill="x", pady=(8, 0))

    @staticmethod
    def _bind_hover(button: tk.Button, normal: str, hover: str) -> None:
        def on_enter(_event):
            if str(button["state"]) != "disabled":
                button.configure(bg=hover)

        def on_leave(_event):
            if str(button["state"]) != "disabled":
                button.configure(bg=normal)

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def _refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, item in enumerate(self.files):
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(item.path.name, item.pages, str(item.path.parent)),
                tags=("odd" if index % 2 else "even",),
            )

        total_pages = sum(item.pages for item in self.files)
        sheets = (
            (total_pages + self.per_sheet - 1) // self.per_sheet if total_pages else 0
        )
        if total_pages:
            direction = "横" if self.landscape else "竖"
            self.summary_label.configure(
                text=(
                    f"共 {len(self.files)} 个文件 · {total_pages} 张 · "
                    f"每页 {self.per_sheet} 张（{direction}）· 占用 {sheets} 张 A4"
                )
            )
            self.empty_label.place_forget()
            self.hint_label.pack(side="right")
        else:
            self.summary_label.configure(text="尚未选择发票")
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
            self.empty_label.lift()
            self.hint_label.pack_forget()

        self._refresh_buttons()
        self._update_scrollbar()

    def _update_scrollbar(self) -> None:
        self.tree.update_idletasks()
        first, last = self.tree.yview()
        if first <= 0.0 and last >= 1.0:
            self.scrollbar.pack_forget()
        elif not self.scrollbar.winfo_ismapped():
            self.scrollbar.pack(side="right", fill="y")

    def _refresh_buttons(self) -> None:
        can_print = bool(self.files) and not self.busy
        self.print_btn.configure(
            state="normal" if can_print else "disabled",
            bg=ACCENT if can_print else ACCENT_DISABLED,
            cursor="hand2" if can_print else "arrow",
        )
        self.select_btn.configure(
            state="disabled" if self.busy else "normal",
            bg=CARD if not self.busy else "#f1f3f7",
            cursor="arrow" if self.busy else "hand2",
        )
        self._refresh_mode_buttons()

    def _refresh_mode_buttons(self) -> None:
        for key, button in self.mode_buttons.items():
            selected = key == self.layout_mode
            if self.busy:
                bg = ACCENT_DISABLED if selected else "#f1f3f7"
                fg = "#f7faff" if selected else "#9aa4b5"
            else:
                bg = ACCENT if selected else CARD
                fg = "#ffffff" if selected else TEXT
            button.configure(
                state="disabled" if self.busy else "normal",
                bg=bg,
                fg=fg,
                disabledforeground=fg,
                activebackground=ACCENT_HOVER if selected else "#eef3fb",
                activeforeground="#ffffff" if selected else TEXT,
                highlightbackground=ACCENT if selected else BORDER,
                cursor="arrow" if self.busy else "hand2",
            )

    def set_layout_mode(self, key: str) -> None:
        if self.busy or key == self.layout_mode or key not in LAYOUTS:
            return
        self.layout_mode = key
        mode = LAYOUTS[key]
        self.per_sheet = mode["per_sheet"]
        self.landscape = mode["landscape"]
        self._refresh()
        self.status_label.configure(text=f"已切换：{mode['label']}")

    def _setup_dnd(self) -> None:
        registered = 0
        for widget in (self.tree, self.empty_label, self.card, self.root):
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)
                registered += 1
            except Exception:
                pass
        if not registered:
            self.status_label.configure(text="拖拽不可用，请使用“选择 PDF 发票”按钮")

    def _on_drop(self, event) -> None:
        if self.busy:
            self.status_label.configure(text="正在打印，暂不能添加文件")
            return
        raw = event.data
        if isinstance(raw, (tuple, list)):
            items = [str(item) for item in raw]
        else:
            try:
                items = [str(item) for item in self.root.tk.splitlist(raw)]
            except Exception:
                items = [str(raw)]

        paths: list[Path] = []
        ignored = 0
        for item in items:
            path = Path(item)
            if path.suffix.lower() == ".pdf" and path.is_file():
                paths.append(path)
            else:
                ignored += 1

        if not paths:
            self.status_label.configure(text="只能拖入 PDF 文件")
            try:
                self.root.bell()
            except Exception:
                pass
            return

        self._add_paths(paths, ignored)

    def _add_paths(self, paths: list[Path], ignored: int = 0) -> None:
        existing = {str(item.path) for item in self.files}
        errors: list[str] = []
        added = 0
        duplicates = 0
        for path in paths:
            key = str(path)
            if key in existing:
                duplicates += 1
                continue
            try:
                invoice = read_invoice_file(path)
            except ValueError as exc:
                errors.append(f"{path.name}：{exc}")
                continue
            self.files.append(invoice)
            existing.add(key)
            added += 1

        self._refresh()

        messages: list[str] = []
        if added:
            messages.append(f"已添加 {added} 个文件")
        if duplicates:
            messages.append(f"跳过 {duplicates} 个重复文件")
        if ignored:
            messages.append(f"忽略 {ignored} 个非 PDF 项目")
        if not messages:
            messages.append("未添加文件")
        self.status_label.configure(text="，".join(messages))

        if errors:
            messagebox.showwarning(
                "部分文件未添加",
                "以下文件无法读取：\n\n" + "\n".join(errors[:8]),
                parent=self.root,
            )

    def choose_files(self) -> None:
        if self.busy:
            return
        selected = filedialog.askopenfilenames(
            parent=self.root,
            title="选择 PDF 发票",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        )
        if not selected:
            return
        self._add_paths([Path(raw_path) for raw_path in selected])

    def _remove_selected(self, _event=None) -> None:
        if self.busy:
            return
        selected = self.tree.selection()
        if not selected:
            return
        indexes = sorted((int(item) for item in selected), reverse=True)
        for index in indexes:
            if 0 <= index < len(self.files):
                del self.files[index]
        self._refresh()
        self.status_label.configure(text="已移除所选文件")

    def start_print(self) -> None:
        if self.busy or not self.files:
            return
        try:
            printer_name = win32print.GetDefaultPrinter()
        except Exception:
            messagebox.showerror(
                "未找到打印机",
                "系统中没有可用的默认打印机，请先在 Windows 设置中添加或启用打印机。",
                parent=self.root,
            )
            return

        self.busy = True
        self.progress["value"] = 0
        self.status_label.configure(text="正在准备…")
        self._refresh_buttons()
        files = list(self.files)
        per_sheet = self.per_sheet
        landscape = self.landscape
        threading.Thread(
            target=self._print_worker,
            args=(files, printer_name, per_sheet, landscape),
            daemon=True,
        ).start()

    def _print_worker(
        self,
        files: list[InvoiceFile],
        printer_name: str,
        per_sheet: int,
        landscape: bool,
    ) -> None:
        temp_dir = Path(tempfile.mkdtemp(prefix="invoice_printer_"))
        try:
            total_pages = sum(item.pages for item in files)
            merged_pdf = temp_dir / "merged.pdf"

            direction = "横版" if landscape else "竖版"
            self._post(
                "status", f"正在排版（每页 {per_sheet} 张，{direction}）…"
            )

            def compose_progress(done: int, total: int) -> None:
                percent = 2 + int(done / max(total, 1) * 13)
                self._post("progress", percent)
                self._post("status", f"正在排版：第 {done}/{total} 张")

            sheets = compose_sheets(
                files,
                merged_pdf,
                compose_progress,
                per_sheet=per_sheet,
                landscape=landscape,
            )
            self._post("status", f"正在发送到打印机：{printer_name}")

            def print_progress(done: int, total: int) -> None:
                percent = 15 + int(done / max(total, 1) * 85)
                self._post("progress", percent)
                self._post("status", f"正在打印：第 {done}/{total} 页 A4")

            send_to_printer(merged_pdf, printer_name, landscape, print_progress)
            self._post(
                "done",
                {
                    "pages": sheets,
                    "invoices": total_pages,
                    "printer": printer_name,
                    "per_sheet": per_sheet,
                    "landscape": landscape,
                },
            )
        except Exception as exc:
            traceback.print_exc()
            self._post("error", str(exc) or exc.__class__.__name__)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _post(self, kind: str, payload: object = None) -> None:
        self.events.put((kind, payload))

    def _drain_events(self) -> None:
        if self.closed:
            return
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "status":
                    self.status_label.configure(text=str(payload))
                elif kind == "progress":
                    self.progress["value"] = int(payload)
                elif kind == "done":
                    info = payload if isinstance(payload, dict) else {}
                    self.progress["value"] = 100
                    self.status_label.configure(text="打印任务已发送")
                    self._finish_print()
                    direction = "横向" if info.get("landscape") else "竖向"
                    messagebox.showinfo(
                        "打印任务已发送",
                        f"已将 {info.get('invoices', 0)} 张排成 "
                        f"{info.get('pages', 0)} 页 A4"
                        f"（{direction}，每页 {info.get('per_sheet', 4)} 张），发送到：\n"
                        f"{info.get('printer', '')}",
                        parent=self.root,
                    )
                elif kind == "error":
                    self.progress["value"] = 0
                    self.status_label.configure(text="打印失败")
                    self._finish_print()
                    messagebox.showerror(
                        "打印失败",
                        f"未能完成打印：\n\n{payload}",
                        parent=self.root,
                    )
        except queue.Empty:
            pass
        self.root.after(80, self._drain_events)

    def _finish_print(self) -> None:
        self.busy = False
        self._refresh_buttons()

    def _on_close(self) -> None:
        if self.busy and not messagebox.askyesno(
            "正在打印",
            "打印任务尚未完成，确定要退出吗？",
            parent=self.root,
        ):
            return
        self.closed = True
        self.root.destroy()


def main() -> None:
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = TkinterDnD.Tk()
    InvoicePrinterApp(root)
    root.mainloop()


if __name__ == "__main__":
    sys.exit(main())
