#!/usr/bin/env python
"""
Excel 批量工具箱 —— 图形界面版
===============================
双击 run_gui.bat，或执行：python gui.py

三个页签：合并多个文件 / 按列拆分 / 批量转 CSV
底层复用 main.py 的核心函数，命令行版本同样可用。
"""
from __future__ import annotations

import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from main import merge_files, read_header, split_by_column, to_csv_files

# Windows 高 DPI 适配（避免文字发虚）
try:
    from ctypes import windll

    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

FONT = "Microsoft YaHei UI"
MONO = "Consolas"


class BaseTab(ttk.Frame):
    """页签基类：统一的两列表单布局。"""

    TITLE = "页签"

    def __init__(self, parent, app: "App"):
        super().__init__(parent, padding=16)
        self.app = app
        self.columnconfigure(1, weight=1)
        self.action_btn: ttk.Button | None = None

    def _path_row(self, row: int, label: str, var: tk.StringVar,
                  browse_cmd, btn_text: str = "浏览") -> None:
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=7)
        wrap = ttk.Frame(self)
        wrap.grid(row=row, column=1, sticky="ew", pady=7)
        wrap.columnconfigure(0, weight=1)
        ttk.Entry(wrap, textvariable=var).grid(row=0, column=0, sticky="ew")
        ttk.Button(wrap, text=btn_text, command=browse_cmd, width=8).grid(row=0, column=1, padx=(6, 0))

    def _hint(self, row: int, text: str) -> None:
        ttk.Label(self, text=text, foreground="#6b6b6b", wraplength=720,
                  justify="left").grid(row=row, column=0, columnspan=2, sticky="w", pady=(10, 0))

    def _action_row(self, row: int, text: str, cmd, extra: tk.Widget | None = None) -> None:
        bar = ttk.Frame(self)
        bar.grid(row=row, column=0, columnspan=2, sticky="w", pady=(14, 4))
        if extra is not None:
            extra.pack(in_=bar, side="left")
        self.action_btn = ttk.Button(bar, text=text, style="Go.TButton", command=cmd)
        self.action_btn.pack(in_=bar, side="left", padx=(18 if extra is not None else 0, 0))

    def buttons(self) -> list[ttk.Button]:
        return [self.action_btn] if self.action_btn else []


class MergeTab(BaseTab):
    TITLE = "合并多个文件"

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.folder = tk.StringVar()
        self.outfile = tk.StringVar()
        self.recursive = tk.BooleanVar(value=False)

        self._path_row(0, "Excel 所在目录", self.folder, self._pick_folder)
        self._path_row(1, "合并后保存为", self.outfile, self._pick_save, "另存为")

        chk = ttk.Checkbutton(self, text="包含子文件夹", variable=self.recursive)
        self._action_row(2, "开始合并", self.start, extra=chk)
        self._hint(3, "把目录下所有 Excel 的所有工作表合并成一张总表，表头以第一个文件为准。")

    def _pick_folder(self):
        p = filedialog.askdirectory(title="选择包含 Excel 文件的目录")
        if p:
            self.folder.set(p)
            if not self.outfile.get().strip():
                self.outfile.set(str(Path(p) / "合并结果.xlsx"))

    def _pick_save(self):
        p = filedialog.asksaveasfilename(
            title="合并结果保存为", defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="合并结果.xlsx",
        )
        if p:
            self.outfile.set(p)

    def start(self):
        folder = self.folder.get().strip()
        if not folder:
            messagebox.showwarning("缺少输入", "请先选择包含 Excel 文件的目录")
            return
        out = self.outfile.get().strip() or str(Path(folder) / "合并结果.xlsx")
        self.app.log(f"开始合并：{folder}", "info")

        def work():
            res = merge_files(folder, out, self.recursive.get())
            return "\n".join(res.lines(out))

        self.app.run_job(work)


class SplitTab(BaseTab):
    TITLE = "按列拆分"

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.file = tk.StringVar()
        self.column = tk.StringVar()
        self.outdir = tk.StringVar()

        self._path_row(0, "要拆分的 Excel", self.file, self._pick_file)

        ttk.Label(self, text="按哪一列拆分").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=7)
        wrap = ttk.Frame(self)
        wrap.grid(row=1, column=1, sticky="ew", pady=7)
        self.combo = ttk.Combobox(wrap, textvariable=self.column, state="normal")
        self.combo.grid(row=0, column=0, sticky="ew")
        wrap.columnconfigure(0, weight=1)
        ttk.Label(wrap, text="选择文件后自动读取列名", foreground="#6b6b6b").grid(
            row=0, column=1, padx=(8, 0))

        self._path_row(2, "拆分结果保存到", self.outdir, self._pick_outdir)
        self._action_row(3, "开始拆分", self.start)
        self._hint(4, "按某一列的取值把一张表拆成多个文件，例如按「部门」「班级」「门店」拆分。")

    def _pick_file(self):
        p = filedialog.askopenfilename(title="选择要拆分的 Excel", filetypes=[("Excel 文件", "*.xlsx")])
        if not p:
            return
        self.file.set(p)
        if not self.outdir.get().strip():
            self.outdir.set(str(Path(p).parent / "拆分结果"))
        try:
            header = read_header(p)
            values = [str(h) for h in header if str(h).strip()]
            self.combo["values"] = values
            if values and not self.column.get().strip():
                self.column.set(values[0])
            self.app.log(f"读取表头成功：{values}", "info")
        except Exception as e:
            self.app.log(f"读取表头失败：{type(e).__name__}: {e}", "err")

    def _pick_outdir(self):
        p = filedialog.askdirectory(title="选择拆分结果保存目录")
        if p:
            self.outdir.set(p)

    def start(self):
        src = self.file.get().strip()
        by = self.column.get().strip()
        outdir = self.outdir.get().strip()
        if not src:
            messagebox.showwarning("缺少输入", "请先选择要拆分的 Excel 文件")
            return
        if not by:
            messagebox.showwarning("缺少输入", "请选择或输入要按哪一列拆分")
            return
        if not outdir:
            outdir = str(Path(src).parent / "拆分结果")
        self.app.log(f"开始拆分：{Path(src).name}  按「{by}」", "info")

        def work():
            result = split_by_column(src, by, outdir)
            lines = [f"拆分完成：{len(result)} 个文件 → {outdir}"]
            for key, count in result.items():
                lines.append(f"    {key}.xlsx  ({count} 行)")
            return "\n".join(lines)

        self.app.run_job(work)


class CsvTab(BaseTab):
    TITLE = "批量转 CSV"

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.folder = tk.StringVar()
        self.outdir = tk.StringVar()
        self.recursive = tk.BooleanVar(value=False)

        self._path_row(0, "Excel 所在目录", self.folder, self._pick_folder)
        self._path_row(1, "CSV 保存到", self.outdir, self._pick_outdir)

        chk = ttk.Checkbutton(self, text="包含子文件夹", variable=self.recursive)
        self._action_row(2, "开始转换", self.start, extra=chk)
        self._hint(3, "把 Excel 批量转成 CSV（UTF-8 带 BOM，用 Excel 打开不会乱码）。")

    def _pick_folder(self):
        p = filedialog.askdirectory(title="选择包含 Excel 文件的目录")
        if p:
            self.folder.set(p)
            if not self.outdir.get().strip():
                self.outdir.set(str(Path(p) / "csv结果"))

    def _pick_outdir(self):
        p = filedialog.askdirectory(title="选择 CSV 保存目录")
        if p:
            self.outdir.set(p)

    def start(self):
        folder = self.folder.get().strip()
        if not folder:
            messagebox.showwarning("缺少输入", "请先选择包含 Excel 文件的目录")
            return
        outdir = self.outdir.get().strip() or str(Path(folder) / "csv结果")
        self.app.log(f"开始转换：{folder}", "info")

        def work():
            n = to_csv_files(folder, outdir, self.recursive.get())
            return f"转换完成：写出 {n} 个 CSV 文件 → {outdir}"

        self.app.run_job(work)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 批量工具箱 v1.0")
        self.minsize(780, 560)
        self._q: queue.Queue = queue.Queue()
        self._busy = False
        self._buttons: list[ttk.Button] = []

        self._setup_style()
        self._center(900, 680)

        nb = ttk.Notebook(self)
        nb.pack(fill="x", expand=False, padx=12, pady=(12, 8))
        for cls in (MergeTab, SplitTab, CsvTab):
            tab = cls(nb, self)
            nb.add(tab, text="  " + tab.TITLE + "  ")
            self._buttons.extend(tab.buttons())

        log_frame = ttk.LabelFrame(self, text="运行日志", padding=6)
        log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        self.log_text = ScrolledText(log_frame, height=11, wrap="word",
                                     font=(MONO, 9), state="disabled")
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_config("ok", foreground="#1a7f37")
        self.log_text.tag_config("err", foreground="#c62828")
        self.log_text.tag_config("info", foreground="#555555")

        self.status = tk.StringVar(value="就绪")
        ttk.Label(self, textvariable=self.status, anchor="w",
                  padding=(14, 3)).pack(fill="x", side="bottom")

        self.log("Excel 批量工具箱已启动。选择目录或文件后点击开始。", "info")
        self.after(120, self._poll)

    # ---------- 外观 ----------

    def _setup_style(self):
        style = ttk.Style(self)
        for theme in ("vista", "winnative", "clam"):
            try:
                style.theme_use(theme)
                break
            except tk.TclError:
                continue
        style.configure(".", font=(FONT, 10))
        style.configure("Go.TButton", font=(FONT, 10, "bold"))
        style.configure("TNotebook.Tab", font=(FONT, 10), padding=(14, 7))
        style.configure("TLabelframe.Label", font=(FONT, 10, "bold"))

    def _center(self, w: int, h: int):
        """按屏幕大小自适应，避免小屏笔记本上窗口超出屏幕。"""
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = max(700, min(w, sw - 60))
        h = max(520, min(h, sh - 120))
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ---------- 日志与任务 ----------

    def log(self, text: str, tag: str = "info"):
        self._q.put(("log", text, tag))

    def run_job(self, work):
        if self._busy:
            messagebox.showinfo("请稍候", "上一个任务还在处理中")
            return
        self._busy = True
        self._set_buttons(False)
        self.status.set("处理中…")

        def worker():
            try:
                summary = work()
                self._q.put(("done", str(summary)))
            except PermissionError as e:
                self._q.put(("error", f"写入被拒绝，目标文件可能正被 Excel 打开，关掉再试：{e}"))
            except Exception as e:
                self._q.put(("error", f"{type(e).__name__}: {e}"))

        threading.Thread(target=worker, daemon=True).start()

    def _poll(self):
        try:
            while True:
                item = self._q.get_nowait()
                kind = item[0]
                if kind == "log":
                    self._append(item[1], item[2])
                elif kind == "done":
                    for line in item[1].splitlines():
                        self._append(line, "ok")
                    self._finish(False)
                elif kind == "error":
                    self._append("[失败] " + item[1], "err")
                    self._finish(True)
        except queue.Empty:
            pass
        self.after(120, self._poll)

    def _append(self, text: str, tag: str = "info"):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{stamp}] {text}\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_buttons(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for b in self._buttons:
            b.configure(state=state)

    def _finish(self, error: bool):
        self._busy = False
        self._set_buttons(True)
        self.status.set("出错，请查看日志" if error else "完成")


if __name__ == "__main__":
    App().mainloop()
