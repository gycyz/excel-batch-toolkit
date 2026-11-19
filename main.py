#!/usr/bin/env python
"""
Excel 批量工具箱 —— 核心逻辑 + 命令行入口
==========================================
核心函数供 gui.py 复用，命令行用法见下方 argparse。

    python main.py merge  <输入目录> [--out 合并结果.xlsx] [--recursive]
    python main.py split  <文件.xlsx> --by <列名> [--outdir 拆分结果]
    python main.py to-csv <输入目录> [--outdir csv结果] [--recursive]

依赖：openpyxl

2026-11-19 修订：
1. 合并改为**按表头名对齐**（原来按列位置追加，列顺序不同就会错列）。
2. 合并自动排除输出文件本身（原来第二个月跑会把上次的结果再合并一遍）。
3. 拆分文件名去重 + 分组键规范化（1 与 1.0 合并，A/B 与 A_B 不再互相覆盖）。
4. 转 CSV 导出所有工作表（原来只导活动表），同名自动加序号。
5. 结构不同的工作表（说明页/封面页）跳过而不是灌进总表。
"""
from __future__ import annotations

import argparse
import csv
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TMP_PREFIX = "~$"          # Excel 打开文件时产生的锁文件前缀
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
HEADER_FONT = Font(bold=True)


# ============ 小工具 ============

def xlsx_files(folder: str | Path, recursive: bool = False,
               exclude: tuple | list = ()) -> list[Path]:
    """列出目录下的 .xlsx 文件，跳过 ~$ 锁文件与 exclude 里的路径。"""
    folder = Path(folder)
    if not folder.is_dir():
        raise ValueError(f"不是有效目录：{folder}")
    blocked = {Path(p).expanduser().resolve() for p in exclude}
    pattern = "**/*.xlsx" if recursive else "*.xlsx"
    found = []
    for p in folder.glob(pattern):
        if not p.is_file() or p.name.startswith(TMP_PREFIX):
            continue
        if p.resolve() in blocked:
            continue
        found.append(p)
    return sorted(found)


def is_blank(row) -> bool:
    return not any(v not in (None, "") for v in row)


def safe_name(text: str, limit: int = 40) -> str:
    keep = "".join(c if (c.isalnum() or c in "-_()（）【】") else "_" for c in text)
    return (keep[:limit] or "空值")


def cell_text(v) -> str:
    """分组键的规范文本：1.0 → 1，日期统一成字符串，空 → 空值。"""
    if v is None:
        return "空值"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M") if (v.hour or v.minute) else v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str):
        return v.strip() or "空值"
    return str(v).strip() or "空值"


def sheet_header_keys(first_row) -> list[str]:
    """表头一行 → 对齐用的列键；空表头用位置兜底，重名加序号。"""
    keys, seen = [], {}
    for i, v in enumerate(first_row):
        name = f"__列{i + 1}" if v is None else (str(v).strip() or f"__列{i + 1}")
        if name in seen:
            seen[name] += 1
            name = f"{name} #{seen[name]}"
        else:
            seen[name] = 1
        keys.append(name)
    return keys


def _style_header(ws, ncols: int) -> None:
    for j in range(1, ncols + 1):
        c = ws.cell(row=1, column=j)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"


def _autofit(ws, max_width: int = 40, sample_rows: int = 200) -> None:
    """按内容宽度设置列宽（中文按 2 个字符宽算）。"""
    widths: dict[int, int] = {}
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, sample_rows)):
        for c in row:
            if c.value is None:
                continue
            w = sum(2 if ord(ch) > 127 else 1 for ch in str(c.value))
            if w > widths.get(c.column, 0):
                widths[c.column] = w
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(w + 2, max_width)


@dataclass
class MergeResult:
    files: int
    rows: int
    skipped_sheets: list[str] = field(default_factory=list)
    added_columns: list[str] = field(default_factory=list)

    def lines(self, out_path) -> list[str]:
        out = [f"合并完成：{self.files} 个文件，{self.rows} 行数据 → {out_path}"]
        if self.added_columns:
            out.append(f"    新增列：{'、'.join(self.added_columns)}")
        if self.skipped_sheets:
            shown = "、".join(self.skipped_sheets[:5])
            more = "" if len(self.skipped_sheets) <= 5 else f" 等 {len(self.skipped_sheets)} 个"
            out.append(f"    跳过非数据工作表：{shown}{more}")
        return out


# ============ 核心函数（GUI 与 CLI 共用） ============

def read_header(src: str | Path) -> list:
    """读取某个 xlsx 第一个工作表的表头（第一行非空行）。"""
    src = Path(src)
    if not src.is_file():
        raise ValueError(f"文件不存在：{src}")
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    try:
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            if is_blank(row):
                continue
            return ["" if v is None else v for v in row]
        return []
    finally:
        wb.close()


def merge_files(folder: str | Path, out_path: str | Path,
                recursive: bool = False) -> MergeResult:
    """合并目录下所有 xlsx 的所有工作表，**按表头名对齐列**。"""
    out_path = Path(out_path)
    files = xlsx_files(folder, recursive, exclude=(out_path,))
    if not files:
        raise ValueError(f"{folder} 下没有找到可合并的 .xlsx 文件"
                         f"（输出文件已自动排除，避免把自己再合并一遍）")

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "合并结果"

    header: list[str] = []
    col_of: dict[str, int] = {}
    started = False
    skipped: list[str] = []
    added: list[str] = []
    total_rows = 0

    for f in files:
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        try:
            for ws in wb.worksheets:
                rows = ws.iter_rows(values_only=True)
                first = next(rows, None)
                if first is None or is_blank(first):
                    continue
                keys = sheet_header_keys(first)

                if not started:
                    for name in keys:
                        if name not in col_of:
                            col_of[name] = len(header)
                            header.append(name)
                    for j, name in enumerate(header, start=1):
                        ws_out.cell(row=1, column=j, value=name)
                    started = True
                else:
                    # 与已有列毫无交集 → 说明页/封面页之类的结构表，整张跳过
                    if not any(name in col_of for name in keys):
                        skipped.append(f"{f.name}[{ws.title}]")
                        continue
                    fresh = [name for name in keys if name not in col_of]
                    if fresh:
                        for name in fresh:
                            col_of[name] = len(header)
                            header.append(name)
                            added.append(name)
                        for j, name in enumerate(header, start=1):
                            ws_out.cell(row=1, column=j, value=name)

                positions = [col_of[name] for name in keys]
                for row in rows:
                    if is_blank(row):
                        continue
                    vals = list(row)
                    out = [None] * len(header)
                    for src_j, out_j in enumerate(positions):
                        if src_j < len(vals):
                            out[out_j] = vals[src_j]
                    ws_out.append(out)
                    total_rows += 1
        finally:
            wb.close()

    if not started:
        raise ValueError("所有文件都是空的，没有可合并的内容")

    _style_header(ws_out, len(header))
    _autofit(ws_out)
    if header:
        ws_out.auto_filter.ref = f"A1:{get_column_letter(len(header))}{max(1, total_rows + 1)}"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(out_path)
    return MergeResult(len(files), total_rows, skipped, added)


def split_by_column(src: str | Path, by: str,
                    outdir: str | Path) -> dict[str, int]:
    """按某列取值拆分，返回 {分组名: 行数}。文件名自动去重。"""
    src = Path(src)
    if not src.is_file():
        raise ValueError(f"文件不存在：{src}")

    wb = openpyxl.load_workbook(src, data_only=True)
    try:
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    if not rows:
        raise ValueError("文件是空的")
    header = list(rows[0])
    if by not in header:
        raise ValueError(f"找不到列「{by}」\n现有列：{header}")

    idx = header.index(by)
    groups: dict[str, list] = {}
    for row in rows[1:]:
        if is_blank(row):
            continue
        vals = list(row)
        key = cell_text(vals[idx] if idx < len(vals) else None)
        groups.setdefault(key, []).append(vals)

    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    result: dict[str, int] = {}
    used: set[str] = set()
    for key, group in groups.items():
        base = safe_name(key)
        name = base
        n = 2
        while name.lower() in used:      # A/B 与 A_B 都叫 A_B，不能互相覆盖
            name = f"{base}_{n}"
            n += 1
        used.add(name.lower())

        wb2 = Workbook()
        ws2 = wb2.active
        ws2.title = "数据"
        ws2.append(header)
        for row in group:
            ws2.append(row)
        _style_header(ws2, len(header))
        _autofit(ws2)
        wb2.save(outdir / f"{name}.xlsx")
        result[key] = len(group)
    return result


def to_csv_files(folder: str | Path, outdir: str | Path,
                 recursive: bool = False) -> int:
    """把 xlsx 的**所有工作表**转成 csv（UTF-8 BOM）。返回写出的 csv 个数。"""
    files = xlsx_files(folder, recursive)
    if not files:
        raise ValueError(f"{folder} 下没有找到 .xlsx 文件")

    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    written, used = 0, set()
    for f in files:
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        try:
            for i, ws in enumerate(wb.worksheets):
                suffix = "" if i == 0 else f"__{safe_name(ws.title, 30)}"
                base = f"{f.stem}{suffix}"
                name = base
                n = 2
                while f"{name}.csv".lower() in used:   # 递归时不同目录同名文件
                    name = f"{base}__{n}"
                    n += 1

                it = ws.iter_rows(values_only=True)
                first = next(it, None)
                if first is None:
                    continue                           # 空表不产出空 csv

                used.add(f"{name}.csv".lower())
                with open(outdir / f"{name}.csv", "w", newline="",
                          encoding="utf-8-sig") as fh:
                    writer = csv.writer(fh)
                    writer.writerow(["" if v is None else v for v in first])
                    for row in it:
                        writer.writerow(["" if v is None else v for v in row])
                written += 1
        finally:
            wb.close()
    return written


# ============ 命令行 ============

def cmd_merge(args) -> int:
    res = merge_files(args.folder, args.out, args.recursive)
    lines = res.lines(args.out)
    print(f"[完成] {lines[0]}")
    for line in lines[1:]:
        print(line)
    return 0


def cmd_split(args) -> int:
    result = split_by_column(args.file, args.by, args.outdir)
    print(f"[完成] 按「{args.by}」拆成 {len(result)} 个文件 -> {args.outdir}")
    for key, count in result.items():
        print(f"    {safe_name(key)}.xlsx  ({count} 行)")
    return 0


def cmd_to_csv(args) -> int:
    n = to_csv_files(args.folder, args.outdir, args.recursive)
    print(f"[完成] 写出 {n} 个 CSV 文件 -> {args.outdir}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="excel-tool",
        description="Excel 批量工具箱：合并 / 拆分 / 转 CSV",
    )
    sub = p.add_subparsers(dest="command", required=True)

    m = sub.add_parser("merge", help="合并目录下所有 xlsx（按表头名对齐）")
    m.add_argument("folder")
    m.add_argument("--out", default="合并结果.xlsx")
    m.add_argument("--recursive", action="store_true", help="含子目录")
    m.set_defaults(func=cmd_merge)

    s = sub.add_parser("split", help="按列拆分一个 xlsx")
    s.add_argument("file")
    s.add_argument("--by", required=True, help="按哪一列拆分（写列名）")
    s.add_argument("--outdir", default="拆分结果")
    s.set_defaults(func=cmd_split)

    c = sub.add_parser("to-csv", help="批量 xlsx 转 csv（所有工作表）")
    c.add_argument("folder")
    c.add_argument("--outdir", default="csv结果")
    c.add_argument("--recursive", action="store_true")
    c.set_defaults(func=cmd_to_csv)

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except (ValueError, FileNotFoundError) as e:
        print(f"[失败] {e}", file=sys.stderr)
        return 2
    except PermissionError as e:
        print(f"[失败] 写入被拒绝，目标文件可能正被 Excel 打开：{e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
