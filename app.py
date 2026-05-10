from __future__ import annotations

import json
import os
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
        self.root.geometry("980x640")

        self.settings_path = Path(__file__).resolve().parent / "settings.json"
        self.config = self._load_settings()
        self.converter = MarkdownToDocxConverter()

        self.output_var = tk.StringVar(value=self.config.output_dir)
        self.asset_var = tk.StringVar(value=self.config.asset_root)
        self.status_var = tk.StringVar(value="就绪")

        self._build_ui()

    def _build_ui(self) -> None:
        settings_frame = tk.LabelFrame(self.root, text="默认设置", padx=8, pady=8)
        settings_frame.pack(fill="x", padx=10, pady=(10, 8))

        tk.Label(settings_frame, text="默认输出目录:").grid(row=0, column=0, sticky="w")
        output_entry = tk.Entry(settings_frame, textvariable=self.output_var)
        output_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        tk.Button(settings_frame, text="浏览", command=self._choose_output_dir, width=8).grid(row=0, column=2)

        tk.Label(settings_frame, text="资源根目录:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        asset_entry = tk.Entry(settings_frame, textvariable=self.asset_var)
        asset_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(8, 0))
        tk.Button(settings_frame, text="浏览", command=self._choose_asset_dir, width=8).grid(row=1, column=2, pady=(8, 0))

        settings_frame.columnconfigure(1, weight=1)

        editor_frame = tk.LabelFrame(self.root, text="Markdown 输入", padx=8, pady=8)
        editor_frame.pack(fill="both", expand=True, padx=10, pady=8)

        self.text_box = ScrolledText(editor_frame, wrap="word", font=("Consolas", 11))
        self.text_box.pack(fill="both", expand=True)

        action_frame = tk.Frame(self.root)
        action_frame.pack(fill="x", padx=10, pady=(0, 10))

        tk.Button(action_frame, text="运行", width=12, height=2, command=self._run_conversion).pack(side="left")
        tk.Label(action_frame, textvariable=self.status_var, anchor="w").pack(side="left", padx=(12, 0), fill="x", expand=True)

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(initialdir=self.output_var.get() or str(Path.cwd()))
        if path:
            self.output_var.set(path)

    def _choose_asset_dir(self) -> None:
        path = filedialog.askdirectory(initialdir=self.asset_var.get() or str(Path.cwd()))
        if path:
            self.asset_var.set(path)

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
            auto_timestamp=True,
        )

        self.status_var.set("正在转换，请稍候...")
        self.root.update_idletasks()

        try:
            output_path = self.converter.convert(markdown_text, config)
            self._save_settings(config)

            warning_count = len(self.converter.last_warnings)
            if warning_count:
                preview = "\n".join(self.converter.last_warnings[:8])
                messagebox.showwarning(
                    "转换完成（含警告）",
                    f"文档已生成:\n{output_path}\n\n警告数量: {warning_count}\n\n{preview}",
                )
                self.status_var.set(f"完成: {output_path}（{warning_count} 条警告）")
            else:
                messagebox.showinfo("完成", f"文档已生成:\n{output_path}")
                self.status_var.set(f"完成: {output_path}")

            self._open_output_file(output_path)
        except Exception as exc:  # noqa: BLE001
            self.status_var.set(f"失败: {exc}")
            messagebox.showerror("转换失败", str(exc))

    def _load_settings(self) -> AppConfig:
        default_output = self._detect_downloads_dir()
        default_asset = Path.cwd()

        default = AppConfig(
            output_dir=str(default_output),
            asset_root=str(default_asset),
            title_chars=12,
            auto_timestamp=True,
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
                auto_timestamp=bool(payload.get("auto_timestamp", True)),
            )
        except Exception:
            return default

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

    def _open_output_file(self, output_path: str) -> None:
        try:
            os.startfile(output_path)
        except Exception:
            self.status_var.set(f"完成: {output_path}（已生成，但未能自动打开）")

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    app = MarkdownToWordApp()
    app.run()
