#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
MarkItDown Handy

A simple macOS-friendly Tk GUI for Microsoft MarkItDown.

Design goals:
- Main window is simple: add/drop files, choose output, convert.
- Existing output files are never overwritten by default.
- Advanced settings are hidden in a popup.
- Details/logs are hidden in a popup but stream live during conversion.
- No forced dark/light custom colors: use Tk/macOS default system theme as much as possible.
- Same source supports:
    1) conda-wrapper app: uses existing conda env named "markitdown"
    2) portable app: uses embedded runtime at Contents/Resources/env
"""

import os
import re
import sys
import shlex
import shutil
import tempfile
import threading
import subprocess
import time
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    TkinterDnD = None
    DND_FILES = None
    DND_AVAILABLE = False


APP_TITLE = "MarkItDown Handy"
APP_VERSION = "0.1.0"
CONDA_ENV_NAME = "markitdown"
CANCEL_CODE = -999


def desktop_dir():
    p = Path.home() / "Desktop"
    return p if p.exists() else Path.home()


def default_output_dir():
    return desktop_dir() / "MarkItDown Output"


def bundled_env_dir():
    value = os.environ.get("MARKITDOWN_BUNDLED_ENV", "").strip()
    if value and Path(value).exists():
        return Path(value)
    return None


def resource_dir():
    value = os.environ.get("MARKITDOWN_RESOURCE_DIR", "").strip()
    if value and Path(value).exists():
        return Path(value)
    return Path(__file__).resolve().parent


def find_conda_executable():
    candidates = [
        "/opt/anaconda3/bin/conda",
        str(Path.home() / "anaconda3/bin/conda"),
        str(Path.home() / "miniconda3/bin/conda"),
        "/opt/homebrew/Caskroom/miniforge/base/bin/conda",
        "/opt/homebrew/bin/conda",
        "/usr/local/bin/conda",
    ]
    for item in candidates:
        if Path(item).exists():
            return item
    return shutil.which("conda") or "conda"


def find_ocrmypdf_executable():
    env_dir = bundled_env_dir()
    if env_dir is not None:
        bundled_cli = env_dir / "bin" / "ocrmypdf"
        bundled_python = env_dir / "bin" / "python"
        if bundled_cli.exists():
            return shlex.quote(str(bundled_cli))
        if bundled_python.exists():
            # Robust fallback for conda-pack builds where console scripts are missing.
            return f"{shlex.quote(str(bundled_python))} -m ocrmypdf"

    candidates = [
        "/opt/homebrew/bin/ocrmypdf",
        "/usr/local/bin/ocrmypdf",
    ]
    for item in candidates:
        if Path(item).exists():
            return item
    return shutil.which("ocrmypdf") or "ocrmypdf"


def default_markitdown_command():
    env_dir = bundled_env_dir()
    if env_dir is not None:
        bundled_cli = env_dir / "bin" / "markitdown"
        bundled_python = env_dir / "bin" / "python"
        if bundled_cli.exists():
            return shlex.quote(str(bundled_cli))
        if bundled_python.exists():
            # Robust fallback for conda-pack builds where console scripts are missing.
            return f"{shlex.quote(str(bundled_python))} -m markitdown"
    return f"{shlex.quote(find_conda_executable())} run -n {CONDA_ENV_NAME} markitdown"


def add_gui_paths(env):
    env = env.copy()
    extra_paths = []

    env_dir = bundled_env_dir()
    if env_dir is not None:
        extra_paths.append(str(env_dir / "bin"))
        old_dyld = env.get("DYLD_LIBRARY_PATH", "")
        env["DYLD_LIBRARY_PATH"] = ":".join([str(env_dir / "lib")] + ([old_dyld] if old_dyld else []))

        tessdata = resource_dir() / "tessdata"
        if tessdata.exists():
            env["TESSDATA_PREFIX"] = str(tessdata)

    extra_paths.extend([
        "/opt/anaconda3/bin",
        str(Path.home() / "anaconda3/bin"),
        str(Path.home() / "miniconda3/bin"),
        "/opt/homebrew/bin",
        "/usr/local/bin",
        str(Path(sys.executable).parent),
    ])
    env["PATH"] = ":".join(extra_paths + [env.get("PATH", "")])
    return env


def safe_stem(path):
    name = Path(path).stem or "converted"
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._-")
    return name or "converted"


def unique_path(path):
    """Return a non-existing path by appending _1, _2, ... when needed."""
    path = Path(path)
    if not path.exists():
        return path

    parent = path.parent
    stem = path.stem
    suffix = path.suffix

    index = 1
    while True:
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def read_text_safe(path):
    try:
        return Path(path).read_text(encoding="utf-8", errors="replace")
    except Exception:
        return ""


def quality_score(text):
    text = text or ""
    stripped = text.strip()
    stripped_len = len(stripped)

    if stripped_len == 0:
        return {
            "score": 0,
            "status": "blank",
            "text_len": 0,
            "replacement_count": 0,
            "reason": "empty output",
        }

    replacement_count = text.count("\ufffd")
    visible_chars = re.findall(r"[A-Za-z0-9\u4e00-\u9fff]", text)
    visible_ratio = len(visible_chars) / max(stripped_len, 1)

    markdown_signal = 0
    markdown_signal += len(re.findall(r"^#{1,6}\s+", text, flags=re.MULTILINE)) * 4
    markdown_signal += len(re.findall(r"^\s*[-*+]\s+", text, flags=re.MULTILINE)) * 2
    markdown_signal += len(re.findall(r"\|.*\|", text)) * 2

    score = 0
    if stripped_len >= 80:
        score += 35
    elif stripped_len >= 30:
        score += 20
    elif stripped_len >= 10:
        score += 8

    if visible_ratio >= 0.35:
        score += 25
    elif visible_ratio >= 0.18:
        score += 12

    if replacement_count == 0:
        score += 15
    elif replacement_count <= 3:
        score += 6
    else:
        score -= min(25, replacement_count)

    score += min(25, markdown_signal)
    score = max(0, min(100, score))

    if stripped_len < 20:
        status = "poor"
        reason = "too little text"
    elif replacement_count > max(5, stripped_len * 0.02):
        status = "poor"
        reason = "many replacement characters"
    elif score >= 55:
        status = "good"
        reason = "good text signal"
    elif score >= 35:
        status = "usable"
        reason = "usable text signal"
    else:
        status = "poor"
        reason = "weak text signal"

    return {
        "score": score,
        "status": status,
        "text_len": stripped_len,
        "replacement_count": replacement_count,
        "visible_ratio": visible_ratio,
        "reason": reason,
    }


BaseTk = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk


class DetailsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Conversion Details")
        self.geometry("960x560")

        container = ttk.Frame(self, padding=12)
        container.pack(fill="both", expand=True)

        self.text = tk.Text(container, wrap="word")
        scroll = ttk.Scrollbar(container, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        self.text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def append(self, msg):
        self.text.insert(tk.END, msg + "\n")
        self.text.see(tk.END)


class AdvancedWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Advanced Settings")
        self.geometry("980x700")
        self.minsize(900, 620)

        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text="These settings are optional. The app normally chooses direct conversion, OCR, and fallbacks automatically.",
            wraplength=930,
        ).pack(anchor="w", pady=(0, 10))

        nb = ttk.Notebook(outer)
        nb.pack(fill="both", expand=True)

        # Commands tab
        tab_cmd = ttk.Frame(nb, padding=16)
        nb.add(tab_cmd, text="Commands")
        tab_cmd.columnconfigure(1, weight=1)
        ttk.Label(tab_cmd, text="MarkItDown command").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(tab_cmd, textvariable=parent.markitdown_cmd_var).grid(row=0, column=1, sticky="ew", padx=(10,0), pady=6)
        ttk.Label(tab_cmd, text="Used for the main conversion command. Conda build uses conda run; portable build uses the embedded runtime.", wraplength=760).grid(row=1, column=1, sticky="w")
        ttk.Label(tab_cmd, text="OCRmyPDF command").grid(row=2, column=0, sticky="w", pady=12)
        ttk.Entry(tab_cmd, textvariable=parent.ocrmypdf_cmd_var).grid(row=2, column=1, sticky="ew", padx=(10,0), pady=12)
        ttk.Label(tab_cmd, text="Used only for scanned PDFs or when OCR fallback is needed.", wraplength=760).grid(row=3, column=1, sticky="w")
        ttk.Label(tab_cmd, text="Extra MarkItDown args").grid(row=4, column=0, sticky="w", pady=12)
        ttk.Entry(tab_cmd, textvariable=parent.extra_markitdown_args_var).grid(row=4, column=1, sticky="ew", padx=(10,0), pady=12)
        ttk.Label(tab_cmd, text="Manual CLI override. Usually leave blank unless you know the exact arguments you want.", wraplength=760).grid(row=5, column=1, sticky="w")

        # OCR tab
        tab_ocr = ttk.Frame(nb, padding=16)
        nb.add(tab_ocr, text="OCR")
        tab_ocr.columnconfigure(1, weight=1)
        ttk.Label(tab_ocr, text="Primary OCR language").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(tab_ocr, textvariable=parent.primary_ocr_lang_var, width=20).grid(row=0, column=1, sticky="w", pady=6)
        ttk.Label(tab_ocr, text="Example: eng").grid(row=0, column=2, sticky="w", padx=(12,0), pady=6)
        ttk.Label(tab_ocr, text="Fallback OCR language").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(tab_ocr, textvariable=parent.fallback_ocr_lang_var, width=20).grid(row=1, column=1, sticky="w", pady=6)
        ttk.Label(tab_ocr, text="Example: eng+chi_sim").grid(row=1, column=2, sticky="w", padx=(12,0), pady=6)

        ocr_box = ttk.LabelFrame(tab_ocr, text="OCR options", padding=12)
        ocr_box.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(16,0))
        ttk.Checkbutton(ocr_box, text="Deskew pages before OCR", variable=parent.ocr_deskew_var).grid(row=0, column=0, sticky="w", pady=4)
        ttk.Checkbutton(ocr_box, text="Rotate pages automatically", variable=parent.ocr_rotate_var).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Checkbutton(ocr_box, text="Force OCR if output still looks poor", variable=parent.force_ocr_on_poor_var).grid(row=2, column=0, sticky="w", pady=4)

        # Fallback tab
        tab_fb = ttk.Frame(nb, padding=16)
        nb.add(tab_fb, text="Fallback & Quality")
        tab_fb.columnconfigure(1, weight=1)
        ttk.Label(tab_fb, text="Good score threshold").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Spinbox(tab_fb, from_=10, to=95, textvariable=parent.min_good_score_var, width=8).grid(row=0, column=1, sticky="w", pady=8)
        ttk.Label(tab_fb, text="When a result reaches this score, fallback usually stops.").grid(row=0, column=2, sticky="w", padx=(12,0), pady=8)
        ttk.Label(tab_fb, text="Usable score threshold").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Spinbox(tab_fb, from_=5, to=90, textvariable=parent.min_usable_score_var, width=8).grid(row=1, column=1, sticky="w", pady=8)
        ttk.Label(tab_fb, text="Below this score, output is treated as poor.").grid(row=1, column=2, sticky="w", padx=(12,0), pady=8)

        more_box = ttk.LabelFrame(tab_fb, text="Fallback helpers", padding=12)
        more_box.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(16,0))
        ttk.Checkbutton(more_box, text="Try charset fallback for text / CSV / HTML", variable=parent.try_charset_fallback_var).grid(row=0, column=0, sticky="w", pady=4)
        ttk.Checkbutton(more_box, text="Use MarkItDown plugins", variable=parent.use_plugins_var).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Checkbutton(more_box, text="Keep data URIs", variable=parent.keep_data_uris_var).grid(row=2, column=0, sticky="w", pady=4)

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(10,0))
        ttk.Button(footer, text="Close", command=self.destroy).pack(side="right")


class MarkItDownHandyApp(BaseTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x820")
        self.minsize(1040, 760)

        self.files = []
        self.file_items = {}
        self.is_running = False
        self.cancel_requested = False
        self.current_process = None
        self.run_start_time = None
        self.current_step_text = "Idle"
        self.success_count = 0
        self.fail_count = 0

        self.details_window = None
        self.advanced_window = None
        self.log_lines = []

        self.output_mode_var = tk.StringVar(value="default")
        self.output_dir_var = tk.StringVar(value=str(default_output_dir()))
        self.auto_ocr_var = tk.BooleanVar(value=True)
        self.keep_ocr_pdf_var = tk.BooleanVar(value=True)
        self.open_after_done_var = tk.BooleanVar(value=False)

        self.markitdown_cmd_var = tk.StringVar(value=default_markitdown_command())
        self.ocrmypdf_cmd_var = tk.StringVar(value=find_ocrmypdf_executable())
        self.primary_ocr_lang_var = tk.StringVar(value="eng")
        self.fallback_ocr_lang_var = tk.StringVar(value="eng+chi_sim")
        self.ocr_deskew_var = tk.BooleanVar(value=True)
        self.ocr_rotate_var = tk.BooleanVar(value=True)
        self.min_good_score_var = tk.IntVar(value=55)
        self.min_usable_score_var = tk.IntVar(value=35)
        self.force_ocr_on_poor_var = tk.BooleanVar(value=True)
        self.try_charset_fallback_var = tk.BooleanVar(value=True)
        self.extra_markitdown_args_var = tk.StringVar(value="")
        self.keep_data_uris_var = tk.BooleanVar(value=False)
        self.use_plugins_var = tk.BooleanVar(value=False)

        self.output_mode_var.trace_add("write", lambda *_: self._update_output_controls() if hasattr(self, "workflow_summary") else None)
        self.output_dir_var.trace_add("write", lambda *_: self._update_workflow_summary() if hasattr(self, "workflow_summary") else None)
        self.auto_ocr_var.trace_add("write", lambda *_: self._update_workflow_summary() if hasattr(self, "workflow_summary") else None)
        self.keep_ocr_pdf_var.trace_add("write", lambda *_: self._update_workflow_summary() if hasattr(self, "workflow_summary") else None)

        self._configure_style()
        self._build_ui()
        self._refresh_status()

    def _configure_style(self):
        style = ttk.Style(self)
        available = set(style.theme_names())
        if "aqua" in available:
            try:
                style.theme_use("aqua")
            except Exception:
                pass

        style.configure("Title.TLabel", font=("SF Pro Display", 24, "bold"))
        style.configure("Subtitle.TLabel", font=("SF Pro Text", 12))
        style.configure("Section.TLabelframe.Label", font=("SF Pro Text", 12, "bold"))
        style.configure("StatusKey.TLabel", font=("SF Pro Text", 11, "bold"))
        style.configure("Primary.TButton", font=("SF Pro Text", 12, "bold"), padding=(18, 10))
        style.configure("Secondary.TButton", padding=(12, 7))
        style.configure("Subtle.TLabel", font=("SF Pro Text", 11))


    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        outer = ttk.Frame(self, padding=18)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)
        outer.rowconfigure(2, weight=0)

        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Batch convert files with MarkItDown. The app automatically tries direct extraction first, then OCR for scanned PDFs when needed.",
            style="Subtitle.TLabel",
            wraplength=1200,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        supported_text = (
            "Common inputs with MarkItDown extras: PDF, DOCX, PPTX, XLSX, HTML, TXT, CSV/TSV, "
            "JSON, XML, images, audio, ZIP, and EPUB. Actual support depends on installed extras/plugins."
        )
        ttk.Label(
            header,
            text=supported_text,
            style="Subtle.TLabel",
            wraplength=1200,
            justify="left",
        ).grid(row=2, column=0, sticky="w", pady=(4, 0))

        main = ttk.Frame(outer)
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=5)
        main.columnconfigure(1, weight=3)
        main.rowconfigure(0, weight=1)

        # Left side: queue-centric workflow
        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)

        queue_box = ttk.LabelFrame(left, text="Files", style="Section.TLabelframe")
        queue_box.grid(row=0, column=0, sticky="nsew")
        queue_box.columnconfigure(0, weight=1)
        queue_box.rowconfigure(2, weight=1)

        queue_top = ttk.Frame(queue_box)
        queue_top.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 8))
        queue_top.columnconfigure(0, weight=1)
        self.drop_label = tk.Label(
            queue_top,
            text=self._drop_label_text(),
            relief="groove",
            justify="center",
            cursor="hand2",
            font=("SF Pro Text", 14),
            padx=20,
            pady=22,
        )
        self.drop_label.grid(row=0, column=0, sticky="ew")
        self.drop_label.bind("<Button-1>", lambda _e: self.add_files())
        if DND_AVAILABLE:
            try:
                self.drop_label.drop_target_register(DND_FILES)
                self.drop_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        actions = ttk.Frame(queue_box)
        actions.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))
        ttk.Button(actions, text="Add Files", style="Secondary.TButton", command=self.add_files).pack(side="left")
        ttk.Button(actions, text="Add Folder", style="Secondary.TButton", command=self.add_folder).pack(side="left", padx=(8,0))
        ttk.Button(actions, text="Remove Selected", style="Secondary.TButton", command=self.remove_selected).pack(side="left", padx=(8,0))
        ttk.Button(actions, text="Clear", style="Secondary.TButton", command=self.clear_files).pack(side="left", padx=(8,0))
        self.queue_meta_label = ttk.Label(actions, text="No files yet", style="Subtle.TLabel")
        self.queue_meta_label.pack(side="right")

        tree_area = ttk.Frame(queue_box)
        tree_area.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 10))
        tree_area.columnconfigure(0, weight=1)
        tree_area.rowconfigure(0, weight=1)
        self.file_tree = ttk.Treeview(
            tree_area,
            columns=("name", "folder", "type", "status"),
            show="headings",
            selectmode="extended",
        )
        self.file_tree.heading("name", text="Name")
        self.file_tree.heading("folder", text="Folder")
        self.file_tree.heading("type", text="Type")
        self.file_tree.heading("status", text="Status")
        self.file_tree.column("name", width=270, anchor="w")
        self.file_tree.column("folder", width=280, anchor="w")
        self.file_tree.column("type", width=70, anchor="center")
        self.file_tree.column("status", width=90, anchor="center")
        tree_v = ttk.Scrollbar(tree_area, orient="vertical", command=self.file_tree.yview)
        tree_h = ttk.Scrollbar(tree_area, orient="horizontal", command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=tree_v.set, xscrollcommand=tree_h.set)
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        tree_v.grid(row=0, column=1, sticky="ns")
        tree_h.grid(row=1, column=0, sticky="ew")
        self.file_tree.bind("<Double-1>", self._open_selected_in_finder)
        self.bind("<Delete>", lambda _e: self.remove_selected())

        help_row = ttk.Label(
            queue_box,
            text="Tip: double-click a queued file to reveal it in Finder.",
            style="Subtle.TLabel",
            wraplength=720,
            justify="left",
        )
        help_row.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 12))

        # Right side: decisions + run actions
        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)

        output_box = ttk.LabelFrame(right, text="Save to", style="Section.TLabelframe")
        output_box.grid(row=0, column=0, sticky="ew")
        output_box.columnconfigure(0, weight=1)
        ttk.Radiobutton(output_box, text="Desktop / MarkItDown Output", value="default", variable=self.output_mode_var).grid(row=0, column=0, sticky="w", padx=14, pady=(12,6))
        ttk.Radiobutton(output_box, text="Same folder as each source file", value="same", variable=self.output_mode_var).grid(row=1, column=0, sticky="w", padx=14, pady=6)
        ttk.Radiobutton(output_box, text="Custom folder", value="custom", variable=self.output_mode_var).grid(row=2, column=0, sticky="w", padx=14, pady=6)
        custom_row = ttk.Frame(output_box)
        custom_row.grid(row=3, column=0, sticky="ew", padx=34, pady=(0, 12))
        custom_row.columnconfigure(0, weight=1)
        self.output_entry = ttk.Entry(custom_row, textvariable=self.output_dir_var)
        self.output_entry.grid(row=0, column=0, sticky="ew")
        self.output_choose_btn = ttk.Button(custom_row, text="Choose…", style="Secondary.TButton", command=self.choose_output_dir)
        self.output_choose_btn.grid(row=0, column=1, padx=(8,0))

        options_box = ttk.LabelFrame(right, text="Quick options", style="Section.TLabelframe")
        options_box.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        ttk.Checkbutton(options_box, text="Automatically OCR scanned PDFs", variable=self.auto_ocr_var).pack(anchor="w", padx=14, pady=(12, 6))
        ttk.Checkbutton(options_box, text="Keep OCR-generated PDF copies", variable=self.keep_ocr_pdf_var).pack(anchor="w", padx=14, pady=6)
        ttk.Checkbutton(options_box, text="Open output folder when finished", variable=self.open_after_done_var).pack(anchor="w", padx=14, pady=(6, 12))

        summary_box = ttk.LabelFrame(right, text="Plan", style="Section.TLabelframe")
        summary_box.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        self.workflow_summary = ttk.Label(summary_box, justify="left", wraplength=360)
        self.workflow_summary.pack(anchor="w", fill="x", padx=14, pady=12)

        run_box = ttk.LabelFrame(right, text="Run", style="Section.TLabelframe")
        run_box.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        run_box.columnconfigure(0, weight=1)
        self.convert_btn = ttk.Button(run_box, text="Convert Files", style="Primary.TButton", command=self.start_conversion)
        self.convert_btn.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 8))
        self.cancel_btn = ttk.Button(run_box, text="Cancel Current Run", style="Secondary.TButton", command=self.cancel_conversion, state="disabled")
        self.cancel_btn.grid(row=1, column=0, sticky="ew", padx=14)
        utility = ttk.Frame(run_box)
        utility.grid(row=2, column=0, sticky="ew", padx=14, pady=(10, 14))
        utility.columnconfigure(0, weight=1)
        utility.columnconfigure(1, weight=1)
        ttk.Button(utility, text="Open Output Folder", style="Secondary.TButton", command=self.open_output_folder).grid(row=0, column=0, sticky="ew")
        ttk.Button(utility, text="Advanced Settings", style="Secondary.TButton", command=self.open_advanced).grid(row=0, column=1, sticky="ew", padx=(8,0))
        ttk.Button(utility, text="Full Logs", style="Secondary.TButton", command=self.open_details).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8,0))

        # Bottom: progress and recent activity
        progress_box = ttk.LabelFrame(outer, text="Progress", style="Section.TLabelframe")
        progress_box.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        progress_box.columnconfigure(0, weight=1)

        status_grid = ttk.Frame(progress_box)
        status_grid.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 6))
        status_grid.columnconfigure(1, weight=1)
        ttk.Label(status_grid, text="Current file", style="StatusKey.TLabel").grid(row=0, column=0, sticky="w")
        self.current_file_value = ttk.Label(status_grid, text="-")
        self.current_file_value.grid(row=0, column=1, sticky="w")
        self.queue_summary_label = ttk.Label(status_grid, text="0 files in queue", style="Subtle.TLabel")
        self.queue_summary_label.grid(row=0, column=2, sticky="e")
        ttk.Label(status_grid, text="Current step", style="StatusKey.TLabel").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.current_step_value = ttk.Label(status_grid, text="idle")
        self.current_step_value.grid(row=1, column=1, sticky="w", pady=(6,0))

        self.activity = ttk.Progressbar(progress_box, orient="horizontal", mode="indeterminate")
        self.activity.grid(row=1, column=0, sticky="ew", padx=14, pady=(4, 8))
        self.overall_progress = ttk.Progressbar(progress_box, orient="horizontal", mode="determinate")
        self.overall_progress.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 8))
        self.status_label = ttk.Label(progress_box, text="", wraplength=1200, justify="left")
        self.status_label.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 8))

        recent_box = ttk.Frame(progress_box)
        recent_box.grid(row=4, column=0, sticky="ew", padx=14, pady=(0, 12))
        recent_box.columnconfigure(0, weight=1)
        ttk.Label(recent_box, text="Recent activity", style="StatusKey.TLabel").grid(row=0, column=0, sticky="w")
        self.activity_preview = tk.Text(recent_box, height=7, wrap="word")
        self.activity_preview.grid(row=1, column=0, sticky="ew", pady=(6,0))
        self.activity_preview.configure(state="disabled")

        self._update_output_controls()
        self._update_workflow_summary()
        self._update_queue_summary()
        self._fit_initial_window()

    def _fit_initial_window(self):
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        req_w = max(self.winfo_reqwidth() + 24, 1240)
        req_h = max(self.winfo_reqheight() + 24, 860)
        w = min(req_w, max(1160, screen_w - 110))
        h = min(req_h, max(820, screen_h - 120))
        x = max((screen_w - w) // 2, 24)
        y = max((screen_h - h) // 2, 24)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _drop_label_text(self):
        if DND_AVAILABLE:
            return "Drop files here\nor click to choose files"
        return "Click here to choose files\nInstall tkinterdnd2 to enable drag-and-drop"

    def _update_output_controls(self):
        mode = self.output_mode_var.get()
        custom = mode == "custom"
        self.output_choose_btn.configure(state="normal" if custom else "disabled")
        try:
            self.output_entry.configure(state="normal" if custom else "disabled")
        except Exception:
            pass
        self._update_workflow_summary()

    def _update_workflow_summary(self):
        count = len(self.files)
        queue_text = f"{count} file{'s' if count != 1 else ''} queued"
        destination = {
            "default": "Desktop/MarkItDown Output",
            "same": "each source file's folder",
            "custom": self.output_dir_var.get().strip() or "your custom folder",
        }.get(self.output_mode_var.get(), "Desktop/MarkItDown Output")
        ocr_mode = "automatic OCR for weak/scanned PDFs" if self.auto_ocr_var.get() else "direct conversion only"
        keep_ocr = "keeps OCR PDF copies" if self.keep_ocr_pdf_var.get() else "removes temporary OCR PDFs"
        self.workflow_summary.configure(
            text=f"{queue_text}. Output goes to {destination}. The app uses {ocr_mode} and {keep_ocr}. Existing files are not overwritten; the app adds _1, _2, etc."
        )
        if hasattr(self, 'queue_meta_label'):
            self.queue_meta_label.configure(text=queue_text if count else "No files yet")

    def _update_queue_summary(self):
        count = len(self.files)
        queued_text = f"{count} file{'s' if count != 1 else ''} in queue"
        if hasattr(self, 'queue_summary_label'):
            self.queue_summary_label.configure(text=queued_text)
        if hasattr(self, 'queue_meta_label'):
            self.queue_meta_label.configure(text=queued_text if count else "No files yet")
        self._update_workflow_summary()

    def _set_file_status(self, path, status):
        item = self.file_items.get(str(Path(path).expanduser()))
        if item and self.file_tree.exists(item):
            self.file_tree.set(item, 'status', status)

    def _append_preview_log(self, msg):
        if not hasattr(self, 'activity_preview'):
            return
        self.activity_preview.configure(state='normal')
        self.activity_preview.insert('end', msg + "\n")
        lines = self.activity_preview.get('1.0', 'end-1c').splitlines()
        if len(lines) > 120:
            self.activity_preview.delete('1.0', f"{len(lines)-120}.0")
        self.activity_preview.see('end')
        self.activity_preview.configure(state='disabled')

    def _add_folder_contents(self, folder_path, quiet=False):
        folder_path = Path(folder_path)
        added = 0
        for item in sorted(folder_path.iterdir()):
            if item.is_file() and not item.name.startswith('.'):
                if self._add_file(str(item), quiet=True):
                    added += 1
        if not quiet:
            self.log(f"Added {added} files from {folder_path}")
        return added

    def _open_selected_in_finder(self, _event=None):
        selected = self.file_tree.selection()
        if not selected:
            return
        item = selected[0]
        path = None
        for file_path, item_id in self.file_items.items():
            if item_id == item:
                path = file_path
                break
        if path:
            subprocess.Popen(["open", "-R", path])

    def open_advanced(self):
        if self.advanced_window and self.advanced_window.winfo_exists():
            self.advanced_window.lift()
            return
        self.advanced_window = AdvancedWindow(self)

    def open_details(self):
        if self.details_window and self.details_window.winfo_exists():
            self.details_window.lift()
            return
        self.details_window = DetailsWindow(self)
        for line in self.log_lines:
            self.details_window.append(line)

    def _refresh_status(self):
        if not self.is_running:
            env = add_gui_paths(os.environ)
            md_ok = self._command_available(self.markitdown_cmd_var.get(), env)
            ocr_ok = self._command_available(self.ocrmypdf_cmd_var.get(), env)
            dnd_status = "enabled" if DND_AVAILABLE else "not installed"
            mode = "portable runtime" if bundled_env_dir() is not None else "conda-wrapper mode"
            self.status_label.configure(
                text=f"Ready. v{APP_VERSION} | Mode: {mode} | MarkItDown: {'found' if md_ok else 'not found'} | OCRmyPDF: {'found' if ocr_ok else 'not found'} | Drag/drop: {dnd_status}"
            )
        self.after(5000, self._refresh_status)

    def _tick_elapsed(self):
        if not self.is_running or self.run_start_time is None:
            return
        elapsed = int(time.time() - self.run_start_time)
        self.current_step_value.configure(text=f"{self.current_step_text} | elapsed {elapsed}s")
        self.after(1000, self._tick_elapsed)

    def _command_available(self, cmd_string, env=None):
        try:
            parts = shlex.split(cmd_string)
        except Exception:
            return False
        if not parts:
            return False
        path = None if env is None else env.get("PATH")
        first_ok = Path(parts[0]).exists() or shutil.which(parts[0], path=path) is not None
        if not first_ok:
            return False

        # Portable fallback commands look like: env/bin/python -m markitdown.
        # If the Python exists, command is considered available; real failures will be shown in Details.
        if len(parts) >= 3 and parts[1] == "-m":
            return True

        return True

    def _on_drop(self, event):
        try:
            paths = self.tk.splitlist(event.data)
        except Exception:
            paths = event.data.split()
        for raw in paths:
            path = str(Path(raw).expanduser())
            p = Path(path)
            if p.is_dir():
                self._add_folder_contents(p, quiet=False)
            else:
                self._add_file(path)

    def add_files(self):
        paths = filedialog.askopenfilenames(title="Choose files to convert", filetypes=[("All files", "*.*")])
        for path in paths:
            self._add_file(path)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Choose a folder")
        if not folder:
            return
        self._add_folder_contents(folder)

    def _add_file(self, path, quiet=False):
        if not path:
            return False
        path = str(Path(path).expanduser())
        if path in self.files:
            return False
        self.files.append(path)
        p = Path(path)
        ext = p.suffix.lower().lstrip('.') or 'file'
        item_id = self.file_tree.insert('', 'end', values=(p.name, str(p.parent), ext.upper(), 'Queued'))
        self.file_items[path] = item_id
        self._update_queue_summary()
        if not quiet:
            self.log(f"Added: {path}")
        return True

    def remove_selected(self):
        selected_items = list(self.file_tree.selection())
        if not selected_items:
            return
        all_items = list(self.file_tree.get_children())
        selected_indexes = sorted((all_items.index(item) for item in selected_items), reverse=True)
        for idx in selected_indexes:
            item_id = all_items[idx]
            path = self.files[idx]
            self.file_tree.delete(item_id)
            self.file_items.pop(path, None)
            del self.files[idx]
        self._update_queue_summary()

    def clear_files(self):
        self.files.clear()
        self.file_items.clear()
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        self.overall_progress.configure(value=0)
        self.current_file_value.configure(text="-")
        self.current_step_value.configure(text="idle")
        self._update_queue_summary()

    def choose_output_dir(self):
        folder = filedialog.askdirectory(title="Choose output folder")
        if folder:
            self.output_dir_var.set(folder)
            self._update_workflow_summary()

    def get_output_dir_for_file(self, file_path):
        path = Path(file_path)
        mode = self.output_mode_var.get()
        if mode == "same":
            return path.parent
        if mode == "custom":
            out_dir = Path(self.output_dir_var.get()).expanduser()
        else:
            out_dir = default_output_dir()
        out_dir.mkdir(parents=True, exist_ok=True)
        return out_dir

    def open_output_folder(self):
        if self.output_mode_var.get() == "same" and self.files:
            folder = Path(self.files[0]).expanduser().parent
        elif self.output_mode_var.get() == "custom":
            folder = Path(self.output_dir_var.get()).expanduser()
        else:
            folder = default_output_dir()
        folder.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(["open", str(folder)])

    def log(self, msg=""):
        self.log_lines.append(msg)
        self._append_preview_log(msg)
        if self.details_window and self.details_window.winfo_exists():
            self.details_window.append(msg)

    def safe_log(self, msg=""):
        self.after(0, self.log, msg)

    def set_status(self, msg):
        self.after(0, lambda: self.status_label.configure(text=msg))

    def set_current_file(self, msg):
        self.after(0, lambda: self.current_file_value.configure(text=msg))

    def set_current_step(self, msg):
        self.current_step_text = msg
        self.after(0, lambda: self.current_step_value.configure(text=msg))

    def start_activity(self):
        self.after(0, lambda: self.activity.start(12))

    def stop_activity(self):
        self.after(0, self.activity.stop)

    def start_conversion(self):
        if self.is_running:
            return
        if not self.files:
            messagebox.showwarning(APP_TITLE, "Add at least one file first.")
            return

        self.is_running = True
        self.cancel_requested = False
        self.success_count = 0
        self.fail_count = 0
        self.run_start_time = time.time()

        self.convert_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.overall_progress.configure(maximum=len(self.files), value=0)

        self.log("===== Conversion started =====")
        self._tick_elapsed()
        threading.Thread(target=self.convert_all, daemon=True).start()

    def cancel_conversion(self):
        if not self.is_running:
            return
        self.cancel_requested = True
        self.set_status("Cancelling current process...")
        self.safe_log("Cancel requested by user.")
        proc = self.current_process
        if proc is not None and proc.poll() is None:
            try:
                proc.terminate()
            except Exception:
                pass

    def convert_all(self):
        try:
            env = add_gui_paths(os.environ)
            total = len(self.files)
            for idx, file_path in enumerate(list(self.files), start=1):
                if self.cancel_requested:
                    break
                ok = self.convert_one(file_path, idx, total, env)
                if ok:
                    self.success_count += 1
                    self.after(0, self._set_file_status, file_path, "Done")
                else:
                    self.fail_count += 1
                    if not self.cancel_requested:
                        self.after(0, self._set_file_status, file_path, "Failed")
                self.after(0, lambda v=idx: self.overall_progress.configure(value=v))

            if self.cancel_requested:
                self.safe_log("===== Cancelled =====")
                self.set_status("Cancelled.")
            else:
                self.safe_log("===== All done =====")
                self.set_status("Done.")

            if self.open_after_done_var.get() and not self.cancel_requested:
                self.after(0, self.open_output_folder)

            self.after(0, self.show_completion_message)

        except Exception as exc:
            self.safe_log(f"ERROR: {exc}")
            self.set_status(f"Error: {exc}")
            self.after(0, lambda: messagebox.showerror(APP_TITLE, f"Conversion stopped due to an error:\n\n{exc}"))
        finally:
            self.after(0, self.finish_running)

    def show_completion_message(self):
        if self.cancel_requested:
            messagebox.showwarning(APP_TITLE, f"Conversion cancelled.\n\nSuccess: {self.success_count}\nFailed: {self.fail_count}")
            return
        messagebox.showinfo(
            APP_TITLE,
            f"Conversion finished.\n\nSuccess: {self.success_count}\nFailed: {self.fail_count}\n\nOpen Details to see automatic fallback decisions.",
        )

    def finish_running(self):
        self.is_running = False
        self.current_process = None
        self.run_start_time = None
        self.convert_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.stop_activity()
        self.current_step_value.configure(text="idle")

    def run_cmd(self, cmd, env, step_name):
        self.set_current_step(step_name)
        self.start_activity()
        self.safe_log("$ " + " ".join(shlex.quote(str(x)) for x in cmd))

        output_lines = []
        code = 1

        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                env=env,
            )
            self.current_process = proc

            while True:
                if self.cancel_requested and proc.poll() is None:
                    try:
                        proc.terminate()
                    except Exception:
                        pass

                line = proc.stdout.readline() if proc.stdout else ""
                if line:
                    line = line.rstrip("\n")
                    output_lines.append(line)
                    self.safe_log("  " + line)

                if proc.poll() is not None:
                    rest = proc.stdout.read() if proc.stdout else ""
                    if rest:
                        for rest_line in rest.rstrip().splitlines():
                            output_lines.append(rest_line)
                            self.safe_log("  " + rest_line)
                    break

                time.sleep(0.08)

            code = proc.returncode
            if self.cancel_requested:
                code = CANCEL_CODE

        except FileNotFoundError as exc:
            output_lines.append(str(exc))
            self.safe_log(f"  {exc}")
            code = 127
        finally:
            self.current_process = None
            self.stop_activity()

        return code, "\n".join(output_lines)

    def markitdown_base_cmd(self):
        parts = shlex.split(self.markitdown_cmd_var.get())
        if not parts:
            raise RuntimeError("MarkItDown command is empty.")
        return parts

    def ocrmypdf_base_cmd(self):
        parts = shlex.split(self.ocrmypdf_cmd_var.get())
        if not parts:
            raise RuntimeError("OCRmyPDF command is empty.")
        return parts

    def build_markitdown_cmd(self, source, out_md, extra_args=None):
        cmd = self.markitdown_base_cmd()
        cmd.append(str(source))
        cmd.extend(["-o", str(out_md)])
        if self.use_plugins_var.get():
            cmd.append("--use-plugins")
        if self.keep_data_uris_var.get():
            cmd.append("--keep-data-uris")
        user_extra = self.extra_markitdown_args_var.get().strip()
        if user_extra:
            cmd.extend(shlex.split(user_extra))
        if extra_args:
            cmd.extend(extra_args)
        return cmd

    def run_markitdown_attempt(self, source, out_md, env, label, extra_args=None):
        self.safe_log(f"Attempt: {label}")
        tmp_out = Path(tempfile.mkstemp(suffix=".md")[1])
        cmd = self.build_markitdown_cmd(source, tmp_out, extra_args=extra_args)
        code, output = self.run_cmd(cmd, env, f"MarkItDown: {label}")
        text = read_text_safe(tmp_out)
        q = quality_score(text)
        self.safe_log(f"Quality: {q['status']} | score={q['score']} | chars={q['text_len']} | reason={q['reason']}")

        accepted = q["score"] >= int(self.min_good_score_var.get()) or q["status"] in ["good", "usable"]

        if code == 0 and accepted:
            shutil.copyfile(tmp_out, out_md)

        tmp_out.unlink(missing_ok=True)

        return {
            "ok": code == 0,
            "cancelled": code == CANCEL_CODE,
            "accepted": code == 0 and accepted,
            "quality": q,
            "text": text,
            "label": label,
            "error": output,
        }

    def run_ocr(self, input_pdf, out_dir, env, lang, mode, suffix):
        if self.keep_ocr_pdf_var.get():
            preferred_ocr_pdf = out_dir / f"{Path(input_pdf).stem}_{suffix}.pdf"
            ocr_pdf = unique_path(preferred_ocr_pdf)
            if ocr_pdf != preferred_ocr_pdf:
                self.safe_log(f"OCR PDF exists; using non-overwrite name: {ocr_pdf.name}")
        else:
            ocr_pdf = Path(tempfile.mkstemp(suffix=".pdf")[1])

        cmd = self.ocrmypdf_base_cmd()
        if self.ocr_deskew_var.get():
            cmd.append("--deskew")
        if self.ocr_rotate_var.get():
            cmd.append("--rotate-pages")
        if mode == "auto":
            cmd.append("--skip-text")
        elif mode == "force":
            cmd.append("--force-ocr")
        cmd.extend(["-l", lang, str(input_pdf), str(ocr_pdf)])

        self.safe_log(f"OCR attempt: lang={lang}, mode={mode}")
        code, output = self.run_cmd(cmd, env, f"OCRmyPDF: lang={lang}, mode={mode}")

        if code == CANCEL_CODE:
            raise RuntimeError("Cancelled")
        if code != 0:
            if not self.keep_ocr_pdf_var.get():
                ocr_pdf.unlink(missing_ok=True)
            raise RuntimeError(f"OCRmyPDF failed: {output.strip()}")
        return ocr_pdf

    def convert_one(self, file_path, idx, total, env):
        src = Path(file_path).expanduser()
        self.safe_log("")
        self.safe_log(f"[{idx}/{total}] {src}")
        self.set_current_file(f"{idx}/{total} — {src.name}")
        self.set_status(f"Converting {idx}/{total}: {src.name}")
        self.after(0, self._set_file_status, str(src), "Running")

        if not src.exists():
            self.safe_log(f"SKIP: file not found: {src}")
            self.after(0, self._set_file_status, str(src), "Missing")
            return False

        out_dir = self.get_output_dir_for_file(src)
        preferred_out_md = out_dir / f"{safe_stem(src)}.md"
        out_md = unique_path(preferred_out_md)
        if out_md != preferred_out_md:
            self.safe_log(f"Output exists; using non-overwrite name: {out_md.name}")

        try:
            if src.suffix.lower() == ".pdf":
                return self.convert_pdf_auto(src, out_dir, out_md, env)
            return self.convert_non_pdf_auto(src, out_md, env)
        except RuntimeError as exc:
            if str(exc) == "Cancelled":
                self.safe_log("Cancelled current file.")
                self.after(0, self._set_file_status, str(src), "Cancelled")
                return False
            self.safe_log(f"FAILED: {exc}")
            self.after(0, self._set_file_status, str(src), "Failed")
            return False
        except Exception as exc:
            self.safe_log(f"FAILED: {exc}")
            self.after(0, self._set_file_status, str(src), "Failed")
            return False

    def convert_non_pdf_auto(self, src, out_md, env):
        direct = self.run_markitdown_attempt(src, out_md, env, "direct conversion")
        if direct["cancelled"]:
            raise RuntimeError("Cancelled")
        if direct["accepted"]:
            self.safe_log(f"Accepted direct result -> {out_md}")
            return True

        if self.try_charset_fallback_var.get() and src.suffix.lower() in [".txt", ".csv", ".html", ".htm", ".xml"]:
            for charset in ["utf-8", "gbk", "latin-1"]:
                attempt = self.run_markitdown_attempt(
                    src,
                    out_md,
                    env,
                    f"charset fallback: {charset}",
                    extra_args=["--charset", charset],
                )
                if attempt["cancelled"]:
                    raise RuntimeError("Cancelled")
                if attempt["accepted"]:
                    self.safe_log(f"Accepted charset fallback {charset} -> {out_md}")
                    return True

        if direct["ok"]:
            Path(out_md).write_text(direct["text"], encoding="utf-8")
            self.safe_log(f"Saved best available direct result -> {out_md}")
            return True

        self.safe_log("FAILED: MarkItDown could not convert this file.")
        if direct.get("error"):
            self.safe_log(direct["error"].strip())
        return False

    def convert_pdf_auto(self, src, out_dir, out_md, env):
        direct = self.run_markitdown_attempt(src, out_md, env, "direct PDF conversion")
        if direct["cancelled"]:
            raise RuntimeError("Cancelled")

        if direct["accepted"] and direct["quality"]["score"] >= int(self.min_good_score_var.get()):
            self.safe_log(f"Accepted direct PDF result -> {out_md}")
            return True

        if not self.auto_ocr_var.get():
            if direct["ok"]:
                Path(out_md).write_text(direct["text"], encoding="utf-8")
                self.safe_log(f"Automatic OCR disabled; saved direct result -> {out_md}")
                return True
            self.safe_log("FAILED: direct PDF conversion failed and OCR is disabled.")
            return False

        best = direct if direct["ok"] else None
        temp_ocr_files = []

        try:
            primary_lang = self.primary_ocr_lang_var.get().strip() or "eng"
            fallback_lang = self.fallback_ocr_lang_var.get().strip() or "eng+chi_sim"

            attempts = [
                (primary_lang, "auto", "ocr"),
                (fallback_lang, "auto", "ocr_fallback_lang"),
            ]

            for lang, mode, suffix in attempts:
                if self.cancel_requested:
                    raise RuntimeError("Cancelled")

                ocr_pdf = self.run_ocr(src, out_dir, env, lang=lang, mode=mode, suffix=suffix)
                if not self.keep_ocr_pdf_var.get():
                    temp_ocr_files.append(ocr_pdf)

                attempt = self.run_markitdown_attempt(ocr_pdf, out_md, env, f"after OCR: lang={lang}, mode={mode}")
                if attempt["cancelled"]:
                    raise RuntimeError("Cancelled")

                if best is None or attempt["quality"]["score"] > best["quality"]["score"]:
                    best = attempt

                if attempt["accepted"] and attempt["quality"]["score"] >= int(self.min_good_score_var.get()):
                    self.safe_log(f"Accepted OCR result -> {out_md}")
                    return True

            if self.force_ocr_on_poor_var.get():
                if self.cancel_requested:
                    raise RuntimeError("Cancelled")

                ocr_pdf = self.run_ocr(src, out_dir, env, lang=primary_lang, mode="force", suffix="force_ocr")
                if not self.keep_ocr_pdf_var.get():
                    temp_ocr_files.append(ocr_pdf)

                attempt = self.run_markitdown_attempt(ocr_pdf, out_md, env, f"after force OCR: lang={primary_lang}")
                if attempt["cancelled"]:
                    raise RuntimeError("Cancelled")

                if best is None or attempt["quality"]["score"] > best["quality"]["score"]:
                    best = attempt

                if attempt["accepted"]:
                    self.safe_log(f"Accepted force OCR result -> {out_md}")
                    return True

            if best and best["ok"]:
                Path(out_md).write_text(best["text"], encoding="utf-8")
                self.safe_log(f"Saved best available result from: {best['label']} -> {out_md}")
                return True

            self.safe_log("FAILED: all PDF conversion attempts failed.")
            return False

        finally:
            for temp_path in temp_ocr_files:
                try:
                    temp_path.unlink(missing_ok=True)
                except Exception:
                    pass


if __name__ == "__main__":
    app = MarkItDownHandyApp()
    app.mainloop()
