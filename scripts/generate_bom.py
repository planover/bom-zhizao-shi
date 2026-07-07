#!/usr/bin/env python3
"""
generate_bom.py - BOM智造师 配套脚本

读取包含物料与工艺工序信息的 JSON，生成格式化的 BOM 表 Excel (.xlsx)。
依赖 openpyxl；若运行环境缺失则自动安装。

用法:
    python3 generate_bom.py --data bom.json --out BOM_2026-07-07.xlsx

输入 JSON 结构见 references/bom-spec.md。
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


def load_data(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def validate(data):
    """返回错误列表；为空表示校验通过。"""
    errors = []
    materials = data.get("materials", [])
    if not isinstance(materials, list) or len(materials) == 0:
        errors.append("至少需要一条物料记录")

    for i, m in enumerate(materials, 1):
        name = str(m.get("name") or "").strip()
        unit = str(m.get("unit") or "").strip()
        usage = m.get("usage")
        yield_rate = m.get("yield_rate")

        if not name:
            errors.append(f"物料#{i} 物料名称为必填")
        if not unit:
            errors.append(f"物料#{i} 计量单位为必填")
        try:
            u = float(usage)
            if u <= 0:
                errors.append(f"物料#{i} 用量必须为正数")
        except (TypeError, ValueError):
            errors.append(f"物料#{i} 用量必须为数值")
        try:
            y = float(yield_rate)
            if y <= 0 or y > 100:
                errors.append(f"物料#{i} 出品率须为 0-100 的正数")
        except (TypeError, ValueError):
            errors.append(f"物料#{i} 出品率为必填且须为数值")

    seen = set()
    for i, p in enumerate(data.get("processes", []), 1):
        step_no = str(p.get("step_no") or "").strip()
        name = str(p.get("name") or "").strip()

        if not step_no:
            errors.append(f"工序#{i} 工序编号为必填")
        elif step_no in seen:
            errors.append(f"工序编号 {step_no} 重复")
        else:
            seen.add(step_no)
        if not name:
            errors.append(f"工序#{i} 工序名称为必填")

        wh = p.get("work_hours")
        if wh not in (None, ""):
            try:
                if float(wh) < 0:
                    errors.append(f"工序#{i} 工时须 >= 0")
            except (TypeError, ValueError):
                errors.append(f"工序#{i} 工时格式异常，须为数值或留空")

    return errors


def build_workbook(data):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "BOM表"

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title_font = Font(name="微软雅黑", size=16, bold=True, color="1F3864")
    head_fill = PatternFill("solid", fgColor="D9E1F2")
    head_font = Font(name="微软雅黑", size=11, bold=True, color="1F3864")
    label_font = Font(name="微软雅黑", size=10, bold=True)
    cell_font = Font(name="微软雅黑", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    version = data.get("version") or "V1.0"
    date = data.get("date") or ""

    # 标题
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = "BOM表"
    c.font = title_font
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # 版本号 / 生成日期
    ws["A2"] = f"版本号：{version}"
    ws["A2"].font = label_font
    ws["D2"] = f"生成日期：{date}"
    ws["D2"].font = label_font
    ws.merge_cells("A2:C2")
    ws.merge_cells("D2:E2")

    r = 4
    # 一、物料信息
    ws.cell(r, 1, "一、物料信息").font = label_font
    r += 1
    for col, h in enumerate(["物料名称", "计量单位", "用量", "出品率(%)", ""], 1):
        cell = ws.cell(r, col, h)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = center
        cell.border = border
    r += 1
    for m in data.get("materials", []):
        for col, val in enumerate(
            [m.get("name", ""), m.get("unit", ""), m.get("usage", ""), m.get("yield_rate", ""), ""], 1
        ):
            cell = ws.cell(r, col, val)
            cell.font = cell_font
            cell.border = border
            cell.alignment = left if col in (1, 2) else center
            if col == 4:
                cell.number_format = '0.0"%"'
        r += 1

    r += 1
    # 二、工艺工序
    ws.cell(r, 1, "二、工艺工序").font = label_font
    r += 1
    for col, h in enumerate(["工序编号", "工序名称", "工序说明", "工时", "备注"], 1):
        cell = ws.cell(r, col, h)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = center
        cell.border = border
    r += 1
    for p in data.get("processes", []):
        for col, val in enumerate(
            [p.get("step_no", ""), p.get("name", ""), p.get("desc", ""), p.get("work_hours", ""), p.get("note", "")], 1
        ):
            cell = ws.cell(r, col, val)
            cell.font = cell_font
            cell.border = border
            cell.alignment = left if col in (3, 5) else center
        r += 1

    for i, w in enumerate([16, 12, 12, 14, 30], 1):
        ws.column_dimensions[chr(64 + i)].width = w

    return wb


def main():
    ensure_openpyxl()
    parser = argparse.ArgumentParser(description="生成 BOM 表 Excel")
    parser.add_argument("--data", required=True, help="物料/工序 JSON 文件路径")
    parser.add_argument("--out", required=True, help="输出 xlsx 路径")
    args = parser.parse_args()

    data = load_data(args.data)
    errors = validate(data)
    if errors:
        print("VALIDATION_FAILED")
        for e in errors:
            print(" - " + e)
        sys.exit(2)

    wb = build_workbook(data)
    wb.save(args.out)
    print("OK:" + args.out)


if __name__ == "__main__":
    main()
