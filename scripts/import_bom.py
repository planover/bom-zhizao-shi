#!/usr/bin/env python3
"""
import_bom.py - BOM智造师 配套逆向导入脚本（V2）

读取由 generate_bom.py 生成的 BOM 表 Excel (.xlsx)，反向解析为与正向
输入 JSON Schema 完全一致的结构化数据，便于「重新编辑已有 BOM」或闭环回写。

V2 增强：
- 改为「按列头文本映射列号」（不再硬编码 A–E），对旧版 5 列 Excel 天然兼容
- 解析新增字段：category / output_rate / material_type / process / output
- 跳过物料区的「【分组】」子标题行
- 不解析、不回写「三、配料表」区块（派生数据，重新生成时按 category 重新派生）
- 缺列时取默认值（material_type=其他, process=空, output=空, category=其他, output_rate=空）

依赖 openpyxl；若运行环境缺失则自动安装（复用 ensure_openpyxl 逻辑）。

用法:
    python3 import_bom.py --in BOM_2026-07-07.xlsx [--out data.json]

输出 JSON 结构（与 generate_bom.py 的 --data 输入一致）:
{
  "product_name": "...", "category": "...", "output_rate": 130.0,
  "version": "V1.0", "date": "2026-07-07",
  "materials": [{"name","unit","usage"(float),"yield_rate"(float),
                 "erp_code","material_type","process"}],
  "processes": [{"step_no","name","desc","work_hours"(float 或 ''),"note","output"}]
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


def _row_is_blank(ws, row_index, max_col=7):
    """判断一行（A–max_col）是否完全为空。"""
    for col in range(1, max_col + 1):
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


def _map_header(ws, header_row, max_col=7):
    """读取表头行，返回 {表头文本(去空格): 列号(1-based)} 映射。"""
    mapping = {}
    for col in range(1, max_col + 1):
        value = ws.cell(header_row, col).value
        if value is not None and str(value).strip() != "":
            mapping[str(value).strip()] = col
    return mapping


def _col_of(mapping, candidates):
    """在映射中按候选表头文本查找列号；返回首个命中，未命中返回 None。

    candidates 形如 ["单位", "计量单位"]（优先级从高到低）。
    """
    for cand in candidates:
        if cand in mapping:
            return mapping[cand]
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

    # 产品级字段：扫描包含关键字的单元格并提取冒号后内容。
    version_text = _find_cell_value_with_substring(ws, "版本号")
    date_text = _find_cell_value_with_substring(ws, "生成日期")
    product_name_text = _find_cell_value_with_substring(ws, "产品名称")
    category_text = _find_cell_value_with_substring(ws, "产品类别")
    output_rate_text = _find_cell_value_with_substring(ws, "全产品出品率")

    version = _extract_after_colon(version_text) or "V1.0"
    date = _extract_after_colon(date_text) or ""
    product_name = _extract_after_colon(product_name_text) or ""
    category = _extract_after_colon(category_text) or "其他"

    raw_rate = _extract_after_colon(output_rate_text)
    if raw_rate is not None:
        cleaned = raw_rate.replace("%", "").strip()
        try:
            output_rate = float(cleaned)
        except (TypeError, ValueError):
            output_rate = ""
    else:
        output_rate = ""

    # 定位区标记行。
    material_row = _find_marker_row(ws, "一、物料信息")
    if material_row is None:
        print("PARSE_ERROR: 未找到物料区标记『一、物料信息』，无法解析。")
        sys.exit(2)

    process_row = _find_marker_row(ws, "二、工艺工序")
    if process_row is None:
        print("PARSE_ERROR: 未找到工序区标记『二、工艺工序』，无法解析。")
        sys.exit(2)

    # 物料区表头行（标记行 +1），建立列头 → 列号映射。
    material_header_row = material_row + 1
    mat_map = _map_header(ws, material_header_row)
    mat_name_col = mat_map.get("物料名称")
    mat_unit_col = _col_of(mat_map, ["单位", "计量单位"])
    mat_usage_col = mat_map.get("用量")
    mat_yield_col = mat_map.get("出品率(%)")
    mat_erp_col = mat_map.get("ERP物料代码")
    mat_type_col = mat_map.get("物料类型")
    mat_proc_col = mat_map.get("所属工序")

    # 物料数据行：从表头行下一行起，到工序标记行止；
    # 首列以「【」开头的分组子标题行跳过；完全空行跳过。
    materials = []
    r = material_header_row + 1
    while r < process_row:
        col1 = ws.cell(r, 1).value
        if col1 is None or str(col1).strip() == "":
            if _row_is_blank(ws, r):
                r += 1
                continue
            break
        if str(col1).strip().startswith("【"):
            r += 1
            continue
        materials.append(
            {
                "name": str(col1).strip(),
                "unit": _str_or_empty(ws.cell(r, mat_unit_col).value)
                if mat_unit_col
                else "",
                "usage": _to_float(ws.cell(r, mat_usage_col).value)
                if mat_usage_col
                else "",
                "yield_rate": _to_float(ws.cell(r, mat_yield_col).value)
                if mat_yield_col
                else "",
                "erp_code": _str_or_empty(ws.cell(r, mat_erp_col).value)
                if mat_erp_col
                else "",
                "material_type": _str_or_empty(ws.cell(r, mat_type_col).value)
                if mat_type_col
                else "其他",
                "process": _str_or_empty(ws.cell(r, mat_proc_col).value)
                if mat_proc_col
                else "",
            }
        )
        r += 1

    # 工序区表头行（标记行 +1），建立列头 → 列号映射。
    process_header_row = process_row + 1
    proc_map = _map_header(ws, process_header_row)
    proc_no_col = proc_map.get("工序编号")
    proc_name_col = proc_map.get("工序名称")
    proc_desc_col = proc_map.get("工序说明")
    proc_wh_col = proc_map.get("工时")
    proc_note_col = proc_map.get("备注")
    proc_out_col = proc_map.get("产物")

    # 工序数据行：从表头行下一行起，到「三、配料表」标记或空编号止；
    # 定位三、配料表标记，若存在则作为停止边界（不解析、不回写）。
    ingredient_row = _find_marker_row(ws, "三、配料表")

    processes = []
    r = process_header_row + 1
    while r <= ws.max_row:
        if ingredient_row is not None and r >= ingredient_row:
            break
        col1 = ws.cell(r, 1).value
        if col1 is None or str(col1).strip() == "":
            if _row_is_blank(ws, r):
                r += 1
                continue
            break
        processes.append(
            {
                "step_no": str(col1).strip(),
                "name": _str_or_empty(ws.cell(r, proc_name_col).value)
                if proc_name_col
                else "",
                "desc": _str_or_empty(ws.cell(r, proc_desc_col).value)
                if proc_desc_col
                else "",
                "work_hours": _to_float(ws.cell(r, proc_wh_col).value)
                if proc_wh_col
                else "",
                "note": _str_or_empty(ws.cell(r, proc_note_col).value)
                if proc_note_col
                else "",
                "output": _str_or_empty(ws.cell(r, proc_out_col).value)
                if proc_out_col
                else "",
            }
        )
        r += 1

    return {
        "product_name": product_name,
        "category": category,
        "output_rate": output_rate,
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
