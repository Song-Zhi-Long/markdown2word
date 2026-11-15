from __future__ import annotations

import json
import shutil
import subprocess
import sys
from dataclasses import asdict
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from converter import AppConfig, MarkdownToDocxConverter


class MarkdownToWordApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Markdown -> Word 一键转换")
        self.root.geometry("980x680")

        self.settings_path = self._resolve_settings_path()
        self.config = self._load_settings()
        self.converter = MarkdownToDocxConverter()

        self.output_var = tk.StringVar(value=self.config.output_dir)
        self.asset_var = tk.StringVar(value=self.config.asset_root)
        self.file_name_var = tk.StringVar()
        self.body_indent_var = tk.BooleanVar(value=self.config.body_first_line_indent)
        self.body_flush_var = tk.BooleanVar(value=not self.config.body_first_line_indent)
        self.markdown_wrap_var = tk.BooleanVar(value=self.config.markdown_word_wrap)
        self.status_var = tk.StringVar(value="就绪")
        self.find_var = tk.StringVar()
        self.find_status_var = tk.StringVar(value="0 / 0")
        self._find_visible = False
        self._find_matches: list[tuple[str, str]] = []
        self._find_index = -1
        self._find_refresh_job: str | None = None

        self._build_ui()
        self.root.bind_all("<Control-s>", self._run_conversion_shortcut)
        self.root.bind_all("<Control-S>", self._run_conversion_shortcut)
        for key in ("<Control-f>", "<Control-F>"):
            self.root.bind(key, self._show_find_bar)
        self.root.bind("<Escape>", self._hide_find_bar)

    def _build_ui(self) -> None:
        action_frame = tk.Frame(self.root)
        action_frame.pack(fill="x", padx=8, pady=(6, 0))

        filename_frame = tk.Frame(action_frame)
        filename_frame.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        tk.Label(filename_frame, text="保存文件名:").grid(row=0, column=0, sticky="w")
        tk.Entry(filename_frame, textvariable=self.file_name_var, width=1).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(8, 0),
        )
        filename_frame.columnconfigure(1, weight=1)

        status_frame = tk.Frame(action_frame)
        status_frame.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        tk.Button(
            status_frame,
            text="运行并生成 Word",
            command=self._run_conversion,
            height=1,
            padx=6,
            pady=1,
        ).grid(row=0, column=0, sticky="w")
        tk.Label(status_frame, textvariable=self.status_var, anchor="w", width=1).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(12, 0),
        )
        status_frame.columnconfigure(1, weight=1)
        action_frame.columnconfigure(0, weight=1, uniform="top_action")
        action_frame.columnconfigure(1, weight=1, uniform="top_action")

        settings_frame = tk.LabelFrame(self.root, text="默认设置", padx=6, pady=4)
        settings_frame.pack(fill="x", padx=8, pady=(4, 4))

        tk.Label(settings_frame, text="默认输出目录:").grid(row=0, column=0, sticky="w")
        output_entry = tk.Entry(settings_frame, textvariable=self.output_var)
        output_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        tk.Button(
            settings_frame,
            text="浏览",
            command=self._choose_output_dir,
            width=6,
            height=1,
            padx=4,
            pady=1,
        ).grid(row=0, column=2)

        tk.Label(settings_frame, text="资源根目录:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        asset_entry = tk.Entry(settings_frame, textvariable=self.asset_var)
        asset_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(4, 0))
        tk.Button(
            settings_frame,
            text="浏览",
            command=self._choose_asset_dir,
            width=6,
            height=1,
            padx=4,
            pady=1,
        ).grid(row=1, column=2, pady=(4, 0))

        tk.Label(settings_frame, text="正文格式:").grid(row=2, column=0, sticky="w", pady=(4, 0))
        body_format_frame = tk.Frame(settings_frame)
        body_format_frame.grid(row=2, column=1, columnspan=2, sticky="w", padx=(8, 0), pady=(4, 0))
        tk.Checkbutton(
            body_format_frame,
            text="缩进",
            variable=self.body_indent_var,
            command=self._select_body_indent,
        ).pack(side="left")
        tk.Checkbutton(
            body_format_frame,
            text="顶格",
            variable=self.body_flush_var,
            command=self._select_body_flush,
        ).pack(side="left", padx=(12, 0))
        tk.Checkbutton(
            body_format_frame,
            text="自动换行",
            variable=self.markdown_wrap_var,
            command=self._sync_markdown_wrap,
        ).pack(side="left", padx=(24, 0))

        settings_frame.columnconfigure(1, weight=1)

        editor_frame = tk.LabelFrame(self.root, text="Markdown 输入", padx=4, pady=4)
        editor_frame.pack(fill="both", expand=True, padx=8, pady=(4, 6))
        editor_frame.rowconfigure(1, weight=1)
        editor_frame.columnconfigure(0, weight=1)

        self.find_bar = tk.Frame(editor_frame)
        self.find_bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        self.find_bar.columnconfigure(1, weight=1)
        tk.Label(self.find_bar, text="查找:").grid(row=0, column=0)
        self.find_entry = tk.Entry(self.find_bar, textvariable=self.find_var)
        self.find_entry.grid(row=0, column=1, sticky="ew", padx=(6, 8))
        tk.Label(self.find_bar, textvariable=self.find_status_var, width=12).grid(row=0, column=2)
        tk.Button(self.find_bar, text="上一个", command=lambda: self._navigate_find(-1)).grid(row=0, column=3)
        tk.Button(self.find_bar, text="下一个", command=lambda: self._navigate_find(1)).grid(row=0, column=4, padx=4)
        tk.Button(self.find_bar, text="关闭", command=self._hide_find_bar).grid(row=0, column=5)
        self.find_entry.bind("<Return>", lambda event: self._navigate_find(1))
        self.find_entry.bind("<Shift-Return>", lambda event: self._navigate_find(-1))
        self.find_bar.grid_remove()

        self.horizontal_scrollbar = tk.Scrollbar(editor_frame, orient="horizontal")
        self.text_box = ScrolledText(
            editor_frame,
            wrap=self._current_markdown_wrap(),
            font=("Consolas", 11),
            xscrollcommand=self.horizontal_scrollbar.set,
            undo=True,
            autoseparators=True,
            maxundo=-1,
        )
        self.horizontal_scrollbar.config(command=self.text_box.xview)
        self.text_box.grid(row=1, column=0, sticky="nsew")
        self.horizontal_scrollbar.grid(row=2, column=0, sticky="ew")
        self.text_box.tag_configure("find_match", background="#FFF2A8", foreground="#000000")
        self.text_box.tag_configure("find_current", background="#F7B955", foreground="#000000")
        for key in ("<Control-z>", "<Control-Z>"):
            self.text_box.bind(key, self._undo_edit)
        for key in ("<Control-y>", "<Control-Y>"):
            self.text_box.bind(key, self._redo_edit)
        for key in ("<Control-f>", "<Control-F>"):
            self.text_box.bind(key, self._show_find_bar)
        self.text_box.bind("<<Modified>>", self._on_editor_modified)
        self.find_var.trace_add("write", self._schedule_find_refresh)
        self._sync_markdown_wrap()

    def _undo_edit(self, _event: tk.Event | None = None) -> str:
        try:
            self.text_box.edit_undo()
        except tk.TclError:
            pass  # An empty undo history is harmless.
        return "break"

    def _redo_edit(self, _event: tk.Event | None = None) -> str:
        try:
            self.text_box.edit_redo()
        except tk.TclError:
            pass
        return "break"

    def _show_find_bar(self, _event: tk.Event | None = None) -> str:
        self._find_visible = True
        self.find_bar.grid()
        try:
            selection = self.text_box.get("sel.first", "sel.last")
            if "\n" not in selection:
                self.find_var.set(selection)
        except tk.TclError:
            pass
        self._refresh_find()
        self.find_entry.focus_set()
        self.find_entry.selection_range(0, "end")
        return "break"

    def _hide_find_bar(self, _event: tk.Event | None = None) -> str | None:
        if not self._find_visible:
            return None
        self._find_visible = False
        self._cancel_find_refresh()
        self.find_bar.grid_remove()
        self.text_box.tag_remove("find_match", "1.0", "end")
        self.text_box.tag_remove("find_current", "1.0", "end")
        self._find_matches = []
        self._find_index = -1
        self.text_box.focus_set()
        return "break"

    def _on_editor_modified(self, _event: tk.Event) -> None:
        if self.text_box.edit_modified():
            self.text_box.edit_modified(False)
            self._schedule_find_refresh()

    def _cancel_find_refresh(self) -> None:
        if self._find_refresh_job is not None:
            self.root.after_cancel(self._find_refresh_job)
            self._find_refresh_job = None

    def _schedule_find_refresh(self, *_args) -> None:
        self._cancel_find_refresh()
        if self._find_visible:
            self._find_refresh_job = self.root.after(80, self._refresh_find)

    def _refresh_find(self) -> None:
        self._cancel_find_refresh()
        if not self._find_visible:
            return
        anchor = self.text_box.index("insert")
        self.text_box.tag_remove("find_match", "1.0", "end")
        self.text_box.tag_remove("find_current", "1.0", "end")
        self._find_matches = []
        self._find_index = -1
        query = self.find_var.get()
        if query:
            position = "1.0"
            while match := self.text_box.search(query, position, stopindex="end-1c", nocase=True, exact=True):
                # Tk search counts UTF-16 units on Windows; index offsets count
                # characters, so use the literal query's length (also for emoji).
                end = self.text_box.index(f"{match}+{len(query)}c")
                self._find_matches.append((match, end))
                self.text_box.tag_add("find_match", match, end)
                position = end
        if self._find_matches:
            index = next(
                (i for i, (start, _) in enumerate(self._find_matches) if self.text_box.compare(start, ">=", anchor)),
                0,
            )
            self._activate_find_match(index)
        else:
            self.find_status_var.set("0 / 0")

    def _activate_find_match(self, index: int) -> None:
        self._find_index = index % len(self._find_matches)
        start, end = self._find_matches[self._find_index]
        self.text_box.tag_remove("find_current", "1.0", "end")
        self.text_box.tag_add("find_current", start, end)
        self.text_box.tag_raise("find_current")
        self.text_box.mark_set("insert", start)
        self.text_box.see(start)
        self.find_status_var.set(f"{self._find_index + 1} / {len(self._find_matches)}")

    def _navigate_find(self, direction: int) -> str:
        if self._find_refresh_job is not None:
            self._refresh_find()
        if self._find_matches:
            self._activate_find_match(self._find_index + direction)
        return "break"

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(initialdir=self.output_var.get() or str(Path.cwd()))
        if path:
            self.output_var.set(path)

    def _choose_asset_dir(self) -> None:
        path = filedialog.askdirectory(initialdir=self.asset_var.get() or str(Path.cwd()))
        if path:
            self.asset_var.set(path)

    def _select_body_indent(self) -> None:
        if self.body_indent_var.get():
            self.body_flush_var.set(False)
        else:
            self.body_indent_var.set(True)

    def _select_body_flush(self) -> None:
        if self.body_flush_var.get():
            self.body_indent_var.set(False)
        else:
            self.body_flush_var.set(True)

    def _current_markdown_wrap(self) -> str:
        return "word" if self.markdown_wrap_var.get() else "none"

    def _sync_markdown_wrap(self) -> None:
        if not hasattr(self, "text_box"):
            return

        self.text_box.configure(wrap=self._current_markdown_wrap())
        if self.markdown_wrap_var.get():
            self.horizontal_scrollbar.grid_remove()
        else:
            self.horizontal_scrollbar.grid()

    def _run_conversion_shortcut(self, _event: tk.Event) -> str:
        self._run_conversion()
        return "break"

    def _run_conversion(self) -> None:
        markdown_text = self.text_box.get("1.0", "end-1c")
        if not markdown_text.strip():
            messagebox.showwarning("提示", "请输入 Markdown 内容后再运行。")
            return

        default_output = self._detect_downloads_dir()
        config = AppConfig(
            output_dir=self.output_var.get().strip() or str(default_output),
            asset_root=self.asset_var.get().strip() or str(Path.cwd()),
            title_chars=self.config.title_chars,
            body_first_line_indent=self.body_indent_var.get(),
            markdown_word_wrap=self.markdown_wrap_var.get(),
        )

        self.status_var.set("正在转换，请稍候...")
        self.root.update_idletasks()

        try:
            output_path = self.converter.convert(markdown_text, config, self.file_name_var.get())
            self._save_settings(config)

            warning_count = len(self.converter.last_warnings)
            if warning_count:
                self._show_warning_details(output_path, self.converter.last_warnings)

            opened = self._open_output_file(output_path)
            if warning_count:
                suffix = "，已自动打开" if opened else "，未能自动打开"
                self.status_var.set(f"完成: {output_path}（{warning_count} 条警告{suffix}）")
            else:
                suffix = "（已自动打开）" if opened else "（未能自动打开）"
                self.status_var.set(f"完成: {output_path}{suffix}")
        except Exception as exc:  # noqa: BLE001
            detail = str(exc).strip() or exc.__class__.__name__
            self.status_var.set(f"失败: {detail}")
            messagebox.showerror("转换失败", detail)

    def _load_settings(self) -> AppConfig:
        default_output = self._detect_downloads_dir()
        default_asset = Path.cwd()

        default = AppConfig(
            output_dir=str(default_output),
            asset_root=str(default_asset),
            title_chars=12,
            body_first_line_indent=True,
            markdown_word_wrap=True,
        )

        if not self.settings_path.exists():
            return default

        try:
            payload = json.loads(self.settings_path.read_text(encoding="utf-8"))
            output_dir = str(payload.get("output_dir", default.output_dir)).strip()
            asset_root = str(payload.get("asset_root", default.asset_root)).strip()
            legacy_output = str((Path.cwd() / "output").resolve())

            if not output_dir or not Path(output_dir).exists() or str(Path(output_dir).resolve()) == legacy_output:
                output_dir = default.output_dir
            if not asset_root or not Path(asset_root).exists():
                asset_root = default.asset_root

            return AppConfig(
                output_dir=output_dir,
                asset_root=asset_root,
                title_chars=int(payload.get("title_chars", default.title_chars)),
                body_first_line_indent=bool(payload.get("body_first_line_indent", default.body_first_line_indent)),
                markdown_word_wrap=bool(payload.get("markdown_word_wrap", default.markdown_word_wrap)),
            )
        except Exception:
            return default

    def _resolve_settings_path(self) -> Path:
        runtime_dir = self._resolve_runtime_dir()
        return runtime_dir / "settings.json"

    def _resolve_runtime_dir(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent

    def _detect_downloads_dir(self) -> Path:
        home = Path.home()
        candidates = [home / "Downloads", Path.home() / "下载", home]
        for path in candidates:
            if path.exists() and path.is_dir():
                return path
        return home

    def _save_settings(self, config: AppConfig) -> None:
        self.config = config
        data = asdict(config)
        self.settings_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _show_warning_details(self, output_path: str, warnings: list[str]) -> None:
        max_visible = 20
        visible_warnings = warnings[:max_visible]
        warning_lines = "\n".join(f"{idx}. {warning}" for idx, warning in enumerate(visible_warnings, start=1))
        if len(warnings) > max_visible:
            warning_lines += f"\n... 还有 {len(warnings) - max_visible} 条警告未显示"

        messagebox.showwarning(
            "转换完成，有警告",
            f"Word 已生成:\n{output_path}\n\n警告详情:\n{warning_lines}",
        )

    def _open_output_file(self, output_path: str) -> bool:
        try:
            target = str(Path(output_path).resolve())
            if sys.platform.startswith("win"):
                import os

                os.startfile(target)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", target])
            else:
                opener = shutil.which("xdg-open")
                if opener is None:
                    return False
                subprocess.Popen([opener, target])
            return True
        except Exception:
            return False

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    app = MarkdownToWordApp()
    app.run()
