#!/usr/bin/env python
"""Tkinter interface for the PDF-to-PPTX background converter."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import fitz

from pdf_to_pptx_background import convert_pdf


class ConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PDF to PPTX Background")
        self.root.geometry("760x440")
        self.root.minsize(680, 400)

        self.pdf_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.resolution_var = tk.StringVar(value="300")
        self.first_page_var = tk.StringVar(value="1")
        self.page_count_var = tk.StringVar()
        self.overwrite_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="请选择一个 PDF 文件。")
        self.progress_var = tk.DoubleVar(value=0)

        self._events: queue.Queue[tuple] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._last_output: Path | None = None
        self._last_auto_output = ""

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll_events)

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=20)
        container.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        title = ttk.Label(
            container,
            text="PDF 转 PPTX 背景",
            font=("Segoe UI", 17, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

        subtitle = ttk.Label(
            container,
            text="每一页 PDF 会成为不可选中的幻灯片背景图片。",
        )
        subtitle.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 18))

        ttk.Label(container, text="PDF 文件").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=6
        )
        self.pdf_entry = ttk.Entry(container, textvariable=self.pdf_var)
        self.pdf_entry.grid(row=2, column=1, sticky="ew", pady=6)
        self.pdf_button = ttk.Button(
            container, text="浏览...", command=self._choose_pdf
        )
        self.pdf_button.grid(row=2, column=2, padx=(10, 0), pady=6)

        ttk.Label(container, text="输出 PPTX").grid(
            row=3, column=0, sticky="w", padx=(0, 10), pady=6
        )
        self.output_entry = ttk.Entry(container, textvariable=self.output_var)
        self.output_entry.grid(row=3, column=1, sticky="ew", pady=6)
        self.output_button = ttk.Button(
            container, text="浏览...", command=self._choose_output
        )
        self.output_button.grid(row=3, column=2, padx=(10, 0), pady=6)

        options = ttk.LabelFrame(container, text="转换选项", padding=12)
        options.grid(
            row=4, column=0, columnspan=3, sticky="ew", pady=(14, 14)
        )
        for column in (1, 3, 5):
            options.columnconfigure(column, weight=1)

        ttk.Label(options, text="分辨率（DPI）").grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        self.resolution_box = ttk.Combobox(
            options,
            textvariable=self.resolution_var,
            values=("150", "200", "300", "450", "600"),
            width=8,
        )
        self.resolution_box.grid(row=0, column=1, sticky="ew", padx=(0, 18))

        ttk.Label(options, text="起始页").grid(
            row=0, column=2, sticky="w", padx=(0, 8)
        )
        self.first_page_spin = ttk.Spinbox(
            options,
            from_=1,
            to=999999,
            textvariable=self.first_page_var,
            width=8,
        )
        self.first_page_spin.grid(row=0, column=3, sticky="ew", padx=(0, 18))

        ttk.Label(options, text="页数（留空=全部）").grid(
            row=0, column=4, sticky="w", padx=(0, 8)
        )
        self.page_count_entry = ttk.Entry(
            options, textvariable=self.page_count_var, width=10
        )
        self.page_count_entry.grid(row=0, column=5, sticky="ew")

        self.overwrite_check = ttk.Checkbutton(
            options,
            text="允许覆盖已存在的输出文件",
            variable=self.overwrite_var,
        )
        self.overwrite_check.grid(
            row=1, column=0, columnspan=6, sticky="w", pady=(12, 0)
        )

        self.progress = ttk.Progressbar(
            container,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        )
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(2, 8))

        self.status_label = ttk.Label(
            container,
            textvariable=self.status_var,
            wraplength=700,
        )
        self.status_label.grid(row=6, column=0, columnspan=3, sticky="w")

        actions = ttk.Frame(container)
        actions.grid(row=7, column=0, columnspan=3, sticky="e", pady=(20, 0))

        self.open_folder_button = ttk.Button(
            actions,
            text="打开输出目录",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_folder_button.grid(row=0, column=0, padx=(0, 10))

        self.convert_button = ttk.Button(
            actions,
            text="开始转换",
            command=self._start_conversion,
        )
        self.convert_button.grid(row=0, column=1)

    def _choose_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")),
        )
        if not selected:
            return

        self.pdf_var.set(selected)
        suggested = str(
            Path(selected).with_name(f"{Path(selected).stem}-background.pptx")
        )
        if not self.output_var.get().strip() or (
            self.output_var.get().strip() == self._last_auto_output
        ):
            self.output_var.set(suggested)
            self._last_auto_output = suggested

        try:
            with fitz.open(selected) as document:
                self.status_var.set(f"已选择 PDF，共 {document.page_count} 页。")
        except Exception as exc:
            self.status_var.set(f"无法读取 PDF：{exc}")

    def _choose_output(self) -> None:
        initial = self.output_var.get().strip()
        selected = filedialog.asksaveasfilename(
            title="选择输出 PPTX",
            defaultextension=".pptx",
            filetypes=(("PowerPoint presentation", "*.pptx"),),
            initialdir=str(Path(initial).parent) if initial else None,
            initialfile=Path(initial).name if initial else None,
        )
        if selected:
            self.output_var.set(selected)

    def _validate_inputs(self) -> tuple[Path, Path, int, int, int | None] | None:
        pdf_path = Path(self.pdf_var.get().strip()).expanduser()
        output_path = Path(self.output_var.get().strip()).expanduser()

        if not pdf_path.is_file():
            messagebox.showerror("输入错误", "请选择一个有效的 PDF 文件。")
            return None
        if not self.output_var.get().strip():
            messagebox.showerror("输入错误", "请选择输出 PPTX 路径。")
            return None
        if output_path.suffix.lower() != ".pptx":
            output_path = output_path.with_suffix(".pptx")
            self.output_var.set(str(output_path))

        try:
            resolution = int(self.resolution_var.get())
            first_page = int(self.first_page_var.get())
            count_text = self.page_count_var.get().strip()
            page_count = int(count_text) if count_text else None
        except ValueError:
            messagebox.showerror("输入错误", "DPI、起始页和页数必须是整数。")
            return None

        if resolution < 72:
            messagebox.showerror("输入错误", "DPI 不能小于 72。")
            return None
        if first_page < 1:
            messagebox.showerror("输入错误", "起始页不能小于 1。")
            return None
        if page_count is not None and page_count < 1:
            messagebox.showerror("输入错误", "页数必须大于 0，或者留空。")
            return None
        if pdf_path.resolve() == output_path.resolve():
            messagebox.showerror("输入错误", "输入和输出文件不能相同。")
            return None
        if output_path.exists() and not self.overwrite_var.get():
            messagebox.showerror(
                "文件已存在",
                "输出文件已存在。请更改文件名，或勾选允许覆盖。",
            )
            return None

        return (
            pdf_path.resolve(),
            output_path.resolve(),
            resolution,
            first_page,
            page_count,
        )

    def _start_conversion(self) -> None:
        values = self._validate_inputs()
        if values is None:
            return

        self._set_running(True)
        self.progress_var.set(0)
        self.status_var.set("正在准备转换...")
        self._last_output = None

        self._worker = threading.Thread(
            target=self._convert_worker,
            args=values,
            daemon=True,
        )
        self._worker.start()

    def _convert_worker(
        self,
        pdf_path: Path,
        output_path: Path,
        resolution: int,
        first_page: int,
        page_count: int | None,
    ) -> None:
        def report_progress(current: int, total: int, pdf_page: int) -> None:
            self._events.put(("progress", current, total, pdf_page))

        try:
            slide_count = convert_pdf(
                pdf_path=pdf_path,
                output_path=output_path,
                resolution=resolution,
                first_page=first_page,
                page_count=page_count,
                quiet=True,
                progress_callback=report_progress,
            )
        except Exception as exc:
            self._events.put(("error", str(exc)))
        else:
            self._events.put(("done", output_path, slide_count))

    def _poll_events(self) -> None:
        try:
            while True:
                event = self._events.get_nowait()
                kind = event[0]

                if kind == "progress":
                    _, current, total, pdf_page = event
                    self.progress_var.set(current * 100 / total)
                    self.status_var.set(
                        f"正在转换 PDF 第 {pdf_page} 页 "
                        f"（{current}/{total}）..."
                    )
                elif kind == "done":
                    _, output_path, slide_count = event
                    self._last_output = output_path
                    self.progress_var.set(100)
                    self.status_var.set(
                        f"转换完成：{output_path}（{slide_count} 张幻灯片）"
                    )
                    self._set_running(False)
                    self.open_folder_button.configure(state="normal")
                    messagebox.showinfo(
                        "转换完成",
                        f"已生成 {slide_count} 张幻灯片：\n{output_path}",
                    )
                elif kind == "error":
                    _, error_message = event
                    self.status_var.set(f"转换失败：{error_message}")
                    self._set_running(False)
                    messagebox.showerror("转换失败", error_message)
        except queue.Empty:
            pass

        self.root.after(100, self._poll_events)

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.convert_button.configure(state=state)
        self.pdf_button.configure(state=state)
        self.output_button.configure(state=state)
        self.pdf_entry.configure(state=state)
        self.output_entry.configure(state=state)
        self.resolution_box.configure(state=state)
        self.first_page_spin.configure(state=state)
        self.page_count_entry.configure(state=state)
        self.overwrite_check.configure(state=state)
        if running:
            self.open_folder_button.configure(state="disabled")

    def _open_output_folder(self) -> None:
        if self._last_output is None:
            return

        folder = self._last_output.parent
        if sys.platform == "win32":
            os.startfile(folder)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])

    def _on_close(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            if not messagebox.askyesno(
                "转换仍在进行",
                "转换尚未完成。确定要关闭窗口吗？",
            ):
                return
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    try:
        ttk.Style(root).theme_use("vista")
    except tk.TclError:
        pass
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

