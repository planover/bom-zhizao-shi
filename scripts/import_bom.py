#!/usr/bin/env python3
"""
import_bom.py - BOM智造师 配套逆向导入脚本

读取由 generate_bom.py 生成的 BOM 表 Excel (.xlsx)，反向解析为与正向
输入 JSON Schema 完全一致的结构化数据，便于「重新编辑已有 BOM」或闭环回写。

依赖 openpyxl；若运行环境缺失则自动安装（复用 ensure_openpyxl 逻辑）。

用法:
    python3 import_bom.py --in BOM_2026-07-07.xlsx [--out data.json]

输出 JSON 结构（与 generate_bom.py 的 --data 输入一致）:
{
  "version": "V1.0",
  "date": "2026-07-07",
  "materials": [{"name","unit","usage"(float),"yield_rate"(float)}],
  "processes": [{"step_no","name","desc","work_hours"(float 或 ''),"note"}]
}
"""
import argparse
import json
import sys
import subprocess


def ensure_openpyxl():
    """确保 openpyxl 可用，缺失时自动安装。"""
    try:
        import openpyxl  # noqa: F401
        return
    except ImportError:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet", "openpyxl"]
        )


def _find_cell_value_with_substring(ws, substring):
    """在工作表中查找包含指定子串的单元格值（按行、按列顺序，取首个命中）。

    合并单元格的值存储在其左上角（top-left）单元格，因此可直接扫描所有单元格。
    """
    for row in ws.iter_rows():
        for cell in row:
            value = cell.value
            if value is not None and substring in str(value):
                return str(value)
    return None


def _extract_after_colon(text):
    """从「键：值」文本中提取冒号后的内容；找不到冒号则返回 None。

    同时兼容中文全角冒号「：」与 ASCII 半角冒号「:」。
    """
    if text is None:
        return None
    for sep in ("：", ":"):
        if sep in text:
            return text.split(sep, 1)[1].strip()
    return None


def _to_float(value):
    """尽量转为 float；None 或空串返回空字符串，无法解析则保留原值。"""
    if value is None:
        return ""
    if isinstance(value, str) and value.strip() == "":
        return ""
    try:
        return float(value)
    except (TypeError, ValueError):
        return value


def _str_or_empty(value):
    """转为去空格字符串；None 返回空字符串。"""
    if value is None:
        return ""
    return str(value).strip()


def _row_is_blank(ws, row_index):
    """判断一行（A–E）是否完全为空。"""
    for col in range(1, 6):
        if ws.cell(row_index, col).value not in (None, ""):
            return False
    return True


def _find_marker_row(ws, marker):
    """定位首列包含指定标记文本的行号（1-based）；未找到返回 None。"""
    for idx, row in enumerate(ws.iter_rows(min_col=1, max_col=1), 1):
        value = row[0].value
        if value is not None and marker in str(value):
            return idx
    return None


def parse_bom(path):
    """解析 BOM Excel 文件，返回与正向输入 JSON Schema 一致的 dict。

    若缺少必要标记（一、物料信息 / 二、工艺工序），打印 PARSE_ERROR 并以
    退出码 2 结束进程。
    """
    from openpyxl import load_workbook

    try:
        wb = load_workbook(path, data_only=True)
    except Exception as e:
        print("FILE_ERROR: 无法读取 Excel 文件：%s" % e, file=sys.stderr)
        sys.exit(2)
    ws = wb["BOM表"] if "BOM表" in wb.sheetnames else wb.active

    # 标题容错：非预期标题仅告警，不中断。
    title_value = ws["A1"].value
    if title_value != "BOM表":
        print(
            "WARNING: 标题预期为 'BOM表'，实际为 '%s'，仍继续解析。"
            % title_value,
            file=sys.stderr,
        )

    # 版本号 / 生成日期：扫描包含关键字的单元格并提取冒号后内容。
    version_text = _find_cell_value_with_substring(ws, "版本号")
    date_text = _find_cell_value_with_substring(ws, "生成日期")
    version = _extract_after_colon(version_text) or "V1.0"
    date = _extract_after_colon(date_text) or ""

    # 定位区标记行。
    material_row = _find_marker_row(ws, "一、物料信息")
    if material_row is None:
        print("PARSE_ERROR: 未找到物料区标记『一、物料信息』，无法解析。")
        sys.exit(2)

    process_row = _find_marker_row(ws, "二、工艺工序")
    if process_row is None:
        print("PARSE_ERROR: 未找到工序区标记『二、工艺工序』，无法解析。")
        sys.exit(2)

    # 物料区：标记行 +1 为表头（忽略），+2 起为数据，直到工序标记或空名称。
    materials = []
    r = material_row + 2
    while r < process_row:
        col1 = ws.cell(r, 1).value
        if col1 is None or str(col1).strip() == "":
            if _row_is_blank(ws, r):
                r += 1
                continue
            break
        materials.append(
            {
                "name": str(col1).strip(),
                "unit": _str_or_empty(ws.cell(r, 2).value),
                "usage": _to_float(ws.cell(r, 3).value),
                "yield_rate": _to_float(ws.cell(r, 4).value),
            }
        )
        r += 1

    # 工序区：标记行 +1 为表头（忽略），+2 起为数据，直到空编号。
    processes = []
    r = process_row + 2
    while r <= ws.max_row:
        col1 = ws.cell(r, 1).value
        if col1 is None or str(col1).strip() == "":
            if _row_is_blank(ws, r):
                r += 1
                continue
            break
        processes.append(
            {
                "step_no": str(col1).strip(),
                "name": _str_or_empty(ws.cell(r, 2).value),
                "desc": _str_or_empty(ws.cell(r, 3).value),
                "work_hours": _to_float(ws.cell(r, 4).value),
                "note": _str_or_empty(ws.cell(r, 5).value),
            }
        )
        r += 1

    return {
        "version": version,
        "date": date,
        "materials": materials,
        "processes": processes,
    }


def main():
    ensure_openpyxl()
    parser = argparse.ArgumentParser(description="从 BOM 表 Excel 逆向导入 JSON")
    parser.add_argument("--in", dest="in_path", required=True, help="BOM 表 xlsx 文件路径")
    parser.add_argument("--out", default=None, help="可选，导出 JSON 文件路径")
    args = parser.parse_args()

    data = parse_bom(args.in_path)
    json_str = json.dumps(data, ensure_ascii=False, indent=2)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(json_str)
        print("OK:" + args.out)
    else:
        print("OK:" + json_str)


if __name__ == "__main__":
    main()
