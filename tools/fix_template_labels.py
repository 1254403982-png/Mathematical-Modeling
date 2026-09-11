#!/usr/bin/env python
"""按模型时间索引重写官方结果模板的 10 分钟时间标签（右端点版本）。

官方模板的 144 个 10 分钟列/行标签整体右移了 10 分钟：
首列写成 `0:10-0:20`（应为 `0:00-0:10`），末列越界成 `0:00+1-0:10+1`（即次日 24:00-24:10）。

本脚本**只改标签文字，不动任何结构**（合并单元格、列宽、数据格式均保持原样）。
约定：模型时段 t=1..144 覆盖 [(t-1)*10min, t*10min)，标签取右端点 t*10min。
    宽表：Excel 第 2..145 列 <-> t=1..144
    行式明细：Excel 第 2..145 行 <-> t=1..144
4 小时块表（充放电量/紧急购电量）的标签本就正确，不改。

用法：
    python tools/fix_template_labels.py --src <官方模板目录> --dst <输出目录>
"""
from __future__ import annotations

import argparse
import os

import openpyxl

FILES = ["result1.xlsx", "result2.xlsx", "result3.xlsx", "result4-2.xlsx", "result4-3.xlsx"]


def slot_label(idx: int) -> str:
    """模型时段 idx（1..144）的正确标签，右端点版本。"""
    start, end = (idx - 1) * 10, idx * 10

    def hm(m: int) -> str:
        return f"{m // 60}:{m % 60:02d}"

    return f"{hm(start)}-{hm(end)}"


def fix_workbook(path: str, out_path: str) -> None:
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        hdr = [c.value for c in ws[1]]
        tcols = [
            i + 1
            for i, h in enumerate(hdr)
            if h is not None and "-" in str(h) and ":" in str(h)
        ]
        if len(tcols) >= 140:  # 宽表：第 1 行的 144 个时间列
            for k in tcols:
                ws.cell(row=1, column=k).value = slot_label(k - 1)
        elif ws.max_row >= 144:  # 行式明细（result1「计划购电量」）
            col1 = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
            if all(isinstance(v, str) and "-" in v and ":" in v for v in col1):
                for r in range(2, ws.max_row + 1):
                    ws.cell(row=r, column=1).value = slot_label(r - 1)
    wb.save(out_path)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--src", required=True, help="官方模板所在目录")
    ap.add_argument("--dst", required=True, help="输出目录（会创建）")
    args = ap.parse_args()

    os.makedirs(args.dst, exist_ok=True)
    for fn in FILES:
        src = os.path.join(args.src, fn)
        if not os.path.exists(src):
            print(f"skip (not found): {fn}")
            continue
        fix_workbook(src, os.path.join(args.dst, fn))
        print(f"saved: {fn}")


if __name__ == "__main__":
    main()
