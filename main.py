#!/usr/bin/env python
"""
Excel 批量工具箱 —— 核心逻辑 + 命令行入口
==========================================
核心函数供 gui.py 复用，命令行用法见下方 argparse。

    python main.py merge  <输入目录> [--out 合并结果.xlsx] [--recursive]
    python main.py split  <文件.xlsx> --by <列名> [--outdir 拆分结果]
    python main.py to-csv <输入目录> [--outdir csv结果] [--recursive]

依赖：openpyxl
"""
from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

import openpyxl
from openpyxl import Workbook


# ============ 核心函数（GUI 与 CLI 共用） ============

def xlsx_files(folder: str | Path, recursive: bool = False) -> list[Path]:
    """列出目录下的 .xlsx 文件（跳过 Excel 打开时产生的 ~$ 临时文件）。"""
    folder = Path(folder)
    if not folder.is_dir():
        raise ValueError(f"不是有效目录：{folder}")
    pattern = "**/*.xlsx" if recursive else "*.xlsx"
    return sorted(
        p for p in folder.glob(pattern)
        if p.is_file() and not p.name.startswith("~$")
    )


def is_blank(row) -> bool:
    return not any(v not in (None, "") for v in row)


def safe_name(text: str, limit: int = 40) -> str:
    keep = "".join(c if (c.isalnum() or c in "-_()（）【】") else "_" for c in text)
    return (keep[:limit] or "空值")


def read_header(src: str | Path) -> list:
    """读取某个 xlsx 第一个工作表的表头（第一行）。"""
    src = Path(src)
    if not src.is_file():
        raise ValueError(f"文件不存在：{src}")
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    try:
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            return ["" if v is None else v for v in row]
        return []
    finally:
        wb.close()


def merge_files(folder: str | Path, out_path: str | Path,
                recursive: bool = False) -> tuple[int, int]:
    """合并目录下所有 xlsx 的所有工作表。返回 (文件数, 数据行数)。"""
    files = xlsx_files(folder, recursive)
    if not files:
        raise ValueError(f"{folder} 下没有找到 .xlsx 文件")

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "合并结果"

    header: list | None = None
    total_rows = 0

    for f in files:
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        try:
            for ws in wb.worksheets:
                for i, row in enumerate(ws.iter_rows(values_only=True)):
                    if i == 0:
                        if header is None:
                            header = list(row)
                            ws_out.append(header)
                        continue
                    if is_blank(row):
                        continue
                    ws_out.append(list(row))
                    total_rows += 1
        finally:
            wb.close()

    out_path = Path(out_path)
    if out_path.parent and not out_path.parent.exists():
        out_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(out_path)
    return len(files), total_rows


def split_by_column(src: str | Path, by: str,
                    outdir: str | Path) -> dict[str, int]:
    """按某列取值拆分，返回 {分组名: 行数}。"""
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
        key = str(row[idx]) if row[idx] is not None else "空值"
        groups.setdefault(key, []).append(row)

    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    result: dict[str, int] = {}
    for key, group in groups.items():
        wb2 = Workbook()
        ws2 = wb2.active
        ws2.title = "数据"
        ws2.append(header)
        for row in group:
            ws2.append(list(row))
        wb2.save(outdir / f"{safe_name(key)}.xlsx")
        result[key] = len(group)
    return result


def to_csv_files(folder: str | Path, outdir: str | Path,
                 recursive: bool = False) -> int:
    """批量把 xlsx 转成 csv（UTF-8 BOM，Excel 打开不乱码）。返回文件数。"""
    files = xlsx_files(folder, recursive)
    if not files:
        raise ValueError(f"{folder} 下没有找到 .xlsx 文件")

    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    for f in files:
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        try:
            ws = wb.active
            target = outdir / f"{f.stem}.csv"
            with open(target, "w", newline="", encoding="utf-8-sig") as fh:
                writer = csv.writer(fh)
                for row in ws.iter_rows(values_only=True):
                    writer.writerow(["" if v is None else v for v in row])
        finally:
            wb.close()
    return len(files)


# ============ 命令行 ============

def cmd_merge(args) -> int:
    n_files, n_rows = merge_files(args.folder, args.out, args.recursive)
    print(f"[完成] 合并 {n_files} 个文件，{n_rows} 行数据 -> {args.out}")
    return 0


def cmd_split(args) -> int:
    result = split_by_column(args.file, args.by, args.outdir)
    print(f"[完成] 按「{args.by}」拆成 {len(result)} 个文件 -> {args.outdir}")
    for key, count in result.items():
        print(f"    {safe_name(key)}.xlsx  ({count} 行)")
    return 0


def cmd_to_csv(args) -> int:
    n = to_csv_files(args.folder, args.outdir, args.recursive)
    print(f"[完成] {n} 个文件转 CSV -> {args.outdir}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="excel-tool",
        description="Excel 批量工具箱：合并 / 拆分 / 转 CSV",
    )
    sub = p.add_subparsers(dest="command", required=True)

    m = sub.add_parser("merge", help="合并目录下所有 xlsx")
    m.add_argument("folder")
    m.add_argument("--out", default="合并结果.xlsx")
    m.add_argument("--recursive", action="store_true", help="含子目录")
    m.set_defaults(func=cmd_merge)

    s = sub.add_parser("split", help="按列拆分一个 xlsx")
    s.add_argument("file")
    s.add_argument("--by", required=True, help="按哪一列拆分（写列名）")
    s.add_argument("--outdir", default="拆分结果")
    s.set_defaults(func=cmd_split)

    c = sub.add_parser("to-csv", help="批量 xlsx 转 csv")
    c.add_argument("folder")
    c.add_argument("--outdir", default="csv结果")
    c.add_argument("--recursive", action="store_true")
    c.set_defaults(func=cmd_to_csv)

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
