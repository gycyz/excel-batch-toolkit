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

    def apply_drop(self, paths: list[str]) -> bool:
        """子类覆盖：把拖进来的路径填到本页签。返回是否接受。"""
        return False

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


    def apply_drop(self, paths: list[str]) -> bool:
        p = Path(paths[0])
        folder = p if p.is_dir() else p.parent
        self.folder.set(str(folder))
        if not self.outfile.get().strip():
            self.outfile.set(str(folder / "合并结果.xlsx"))
        self.app.log(f"已接收：{folder}", "ok")
        return True


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
        if p:
            self._load_file(p)

    def apply_drop(self, paths: list[str]) -> bool:
        p = Path(paths[0])
        if p.is_dir():
            self.outdir.set(str(p))
            self.app.log(f"拆分结果将保存到：{p}", "ok")
            return True
        if p.suffix.lower() != ".xlsx":
            self.app.log(f"只支持 .xlsx，拖进来的是「{p.suffix or '无后缀'}」", "err")
            return False
        self._load_file(str(p))
        return True

    def _load_file(self, p: str):
        """设置待拆分文件并读出表头（选择按钮与拖入共用）。"""
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


    def apply_drop(self, paths: list[str]) -> bool:
        p = Path(paths[0])
        folder = p if p.is_dir() else p.parent
        self.folder.set(str(folder))
        if not self.outdir.get().strip():
            self.outdir.set(str(folder / "csv结果"))
        self.app.log(f"已接收：{folder}", "ok")
        return True


def friendly_error(exc: BaseException) -> str:
    """把异常翻译成甲方能看懂的话，末尾保留原始信息方便排查。"""
    name, msg = type(exc).__name__, str(exc).strip()
    low = msg.lower()
    tail = f"（原始信息：{name}: {msg}）"

    if isinstance(exc, PermissionError) or "permission" in low:
        return ("写不进去：目标文件可能正被 Excel 打开，或者这个文件夹你没有写入权限。\n"
                "先把 Excel 关掉再试一次，或者换一个输出文件夹。\n" + tail)
    if isinstance(exc, FileNotFoundError):
        return ("找不到文件或文件夹：选中的路径可能已经被移走或改名了。\n"
                "重新选一次目录再试。\n" + tail)
    if isinstance(exc, (ValueError, KeyError)):
        return ("这些表格读不进来，常见三种原因：\n"
                "① 是 .xls 老格式（本工具目前只支持 .xlsx）\n"
                "② 表格在子文件夹里，但没勾「包含子文件夹」\n"
                "③ 文件是空的，或者表头不在第一行\n" + tail)
    if isinstance(exc, ImportError) or "no module" in low:
        return ("缺少运行组件，程序不完整。\n"
                "如果是从压缩包拿到的，请把整个文件夹一起解压，不要只把 exe 拎出来。\n" + tail)
    return ("程序碰到了没预料到的问题，不是你的操作出错。\n"
            "把下面这段整段截图发给我就能定位：\n" + tail)


try:  # 拖入支持：装了 tkinterdnd2 就启用，没装就退回"浏览"按钮
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND = True
except Exception:  # noqa: BLE001
    DND_FILES = None
    TkinterDnD = None
    _DND = False


_BaseTk = TkinterDnD.Tk if _DND else tk.Tk


class App(_BaseTk):
    def __init__(self):
        super().__init__()
        self.title("Excel 批量工具箱 v1.1")
        self.minsize(780, 560)
        self._q: queue.Queue = queue.Queue()
        self._busy = False
        self._buttons: list[ttk.Button] = []

        self._setup_style()
        self._center(900, 680)

        self.nb = nb = ttk.Notebook(self)
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

        # A3：不确定态进度条，只在任务运行时出现，避免「看着像没动」
        self.progress = ttk.Progressbar(self, mode="indeterminate")

        self.status = tk.StringVar(value="就绪")
        ttk.Label(self, textvariable=self.status, anchor="w",
                  padding=(14, 3)).pack(fill="x", side="bottom")

        self.log("Excel 批量工具箱已启动。选择目录或文件后点击开始。", "info")
        self._install_dnd()
        self._poll_id = self.after(120, self._poll)

    def destroy(self):
        """关窗口时先取消轮询定时器，否则 Tcl 会报 invalid command name。"""
        try:
            if getattr(self, "_poll_id", None):
                self.after_cancel(self._poll_id)
                self._poll_id = None
        except Exception:  # noqa: BLE001
            pass
        super().destroy()

    def report_callback_exception(self, exc, val, tb):
        """Tk 回调里的异常别只写 stderr（窗口版看不见），直接进日志窗。"""
        self._append(f"[界面异常] {exc.__name__}: {val}", "err")

    # ---------- 拖入 ----------

    def _install_dnd(self):
        if not _DND:
            return
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
            self.log("提示：文件夹或 Excel 文件可以直接拖进这个窗口。", "info")
        except Exception as e:  # noqa: BLE001
            self.log(f"拖入功能未启用（{type(e).__name__}），用「浏览」按钮一样可以。", "info")

    def _on_drop(self, event):
        try:
            paths = [p for p in (s.strip() for s in self.tk.splitlist(event.data)) if p]
        except Exception:  # noqa: BLE001
            paths = [str(event.data).strip()]
        if paths:
            try:
                tab = self.nb.nametowidget(self.nb.select())
                apply = getattr(tab, "apply_drop", None)
                if apply is None:
                    self.log("这个页签不支持拖入，请用「浏览」按钮。", "err")
                else:
                    apply(paths)
            except Exception as e:  # noqa: BLE001
                self.log(f"拖入处理失败：{type(e).__name__}: {e}", "err")
        return "break"

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
        self.progress.pack(fill="x", padx=14, pady=(0, 2), side="bottom")
        self.progress.start(12)

        def worker():
            try:
                summary = work()
                self._q.put(("done", str(summary)))
            except Exception as e:  # noqa: BLE001
                self._q.put(("error", friendly_error(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _poll(self):
        try:
            while True:
                try:
                    item = self._q.get_nowait()
                except queue.Empty:
                    break
                try:
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
                except Exception as e:  # noqa: BLE001
                    # 单条消息出错不能打死轮询，否则按钮永远恢复不了
                    self._append(f"[内部错误] {type(e).__name__}: {e}", "err")
                    self._finish(True)
        finally:
            self._poll_id = self.after(120, self._poll)

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
        self.progress.stop()
        self.progress.pack_forget()
        self._set_buttons(True)
        self.status.set("出错，请查看日志" if error else "完成")


if __name__ == "__main__":
    App().mainloop()
