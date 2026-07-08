#!/usr/bin/env python3
"""
generate_bom.py - BOM智造师 配套脚本（V2）

读取包含产品信息、物料与工艺工序信息的 JSON，生成格式化的 BOM 表 Excel (.xlsx)。
依赖 openpyxl；若运行环境缺失则自动安装。

V2 新增能力：
- 产品级元数据：product_name（必填）、category（必填，5 类枚举）、output_rate（必填，>0 可>100）
- 物料级新字段：material_type（配料表过滤）、process（工序归属）
- 工序级新字段：output（工序产物，用于流转链 R3 校验）
- 工序流转链校验（R3，阻断级）
- 食品类自动派生「三、配料表」区块（R4）
- Excel 扩至 7 列（A–G），物料区分工序分组呈现

用法:
    python3 generate_bom.py --data bom.json --out BOM_2026-07-07.xlsx

输入 JSON 结构见 references/bom-spec.md。
"""
import argparse
import datetime
import json
import sys
import subprocess


# 产品类别枚举（R1 / R4）
CATEGORIES = {"食品", "工业品", "日化化妆品", "医药", "其他"}
# 可食用物料类型（R4 配料表过滤）
EDIBLE = {"原料", "添加剂", "香精香料"}


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
    """返回错误列表；为空表示校验通过。

    校验项（V1–V7）：
    - V1(R1): product_name 必填非空
    - V2(R1): category 必填且枚举
    - V3(R2): output_rate 必填且 >0（允许 >100）
    - V4(R2): 物料 yield_rate 须 0<值≤100
    - V5(R3): 工序流转链完整性（仅 len(processes)>=2 触发，阻断级）
    - V6(沿用): 工序编号唯一 / 名称非空 / 工时≥0
    - V7(沿用): 物料 name/unit 非空、usage>0
    """
    errors = []
    if not isinstance(data, dict):
        return ["输入数据必须是 JSON 对象"]

    # V1(R1): product_name 必填非空
    product_name = str(data.get("product_name") or "").strip()
    if not product_name:
        errors.append("产品名称为必填，且不可为空")

    # V2(R1): category 必填且枚举
    category = str(data.get("category") or "").strip()
    if not category or category not in CATEGORIES:
        errors.append("产品类别为必填，且须为：食品/工业品/日化化妆品/医药/其他")

    # V3(R2): output_rate 必填且 >0（允许 >100）
    raw_rate = data.get("output_rate")
    try:
        rate = float(raw_rate)
        if rate <= 0:
            errors.append(
                "全产品出品率(output_rate)为必填，且须为正数（可大于100，如干香菇泡发增重）"
            )
    except (TypeError, ValueError):
        errors.append(
            "全产品出品率(output_rate)为必填，且须为正数（可大于100，如干香菇泡发增重）"
        )

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

    # V5(R3): 工序流转链完整性（阻断级）
    procs = data.get("processes", [])
    if isinstance(procs, list) and len(procs) >= 2:
        for i in range(1, len(procs)):
            prev = procs[i - 1]
            cur = procs[i]
            prev_out = str(prev.get("output") or "").strip()
            if not prev_out:
                errors.append(f"工序 {prev.get('step_no')} 的产物(output)为必填")
                continue
            cur_materials = [
                m
                for m in materials
                if str(m.get("process") or "").strip()
                == str(cur.get("step_no") or "").strip()
            ]
            names = {str(m.get("name") or "").strip() for m in cur_materials}
            if prev_out not in names:
                errors.append(
                    f"工序 {cur.get('step_no')} 的物料清单未包含上一工序 "
                    f"{prev.get('step_no')} 的产物『{prev_out}』，流转链不完整"
                )

    return errors


def derive_ingredients(data):
    """派生配料表（仅食品类）。

    返回 (ingredients, excluded)：
    - ingredients: 可食用物料（material_type ∈ EDIBLE），若 category != 食品 则为空列表
    - excluded: 非食用物料（包材/其他/未分类/未填），始终保留用于 WARNING 统计

    ingredients 按 usage 降序排列（食品标签惯例）。
    """
    ingredients, excluded = [], []
    for m in data.get("materials", []):
        mt = str(m.get("material_type") or "其他").strip() or "其他"
        if mt in EDIBLE:
            ingredients.append(m)
        else:
            excluded.append(m)

    if str(data.get("category") or "其他") != "食品":
        return [], excluded

    ingredients.sort(key=lambda x: float(x.get("usage") or 0), reverse=True)
    return ingredients, excluded


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
    group_fill = PatternFill("solid", fgColor="EAF1FB")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    version = data.get("version") or "V1.0"
    date = data.get("date") or datetime.date.today().isoformat()
    product_name = str(data.get("product_name") or "").strip()
    category = str(data.get("category") or "其他").strip()
    output_rate = data.get("output_rate")

    # 行1：标题
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = "BOM表"
    c.font = title_font
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # 行2：版本号（左） / 生成日期（右）
    ws["A2"] = f"版本号：{version}"
    ws["A2"].font = label_font
    ws["D2"] = f"生成日期：{date}"
    ws["D2"].font = label_font
    ws.merge_cells("A2:C2")
    ws.merge_cells("D2:G2")

    # 行3：产品名称（整行）
    ws.merge_cells("A3:G3")
    c = ws.cell(3, 1, f"产品名称：{product_name}")
    c.font = label_font

    # 行4：产品类别（左） / 全产品出品率（右）
    ws.merge_cells("A4:C4")
    c = ws.cell(4, 1, f"产品类别：{category}")
    c.font = label_font
    ws.merge_cells("D4:G4")
    c = ws.cell(4, 4, f"全产品出品率：{float(output_rate):.1f}%")
    c.font = label_font

    # 行5：空行间隔
    # 行6：一、物料信息
    ws.merge_cells("A6:G6")
    ws.cell(6, 1, "一、物料信息").font = label_font

    # 行7：物料表头（7 列）
    material_headers = [
        "物料名称",
        "单位",
        "用量",
        "出品率(%)",
        "ERP物料代码",
        "物料类型",
        "所属工序",
    ]
    for col, h in enumerate(material_headers, 1):
        cell = ws.cell(7, col, h)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = center
        cell.border = border

    # 物料数据（按需分工序分组）
    materials = data.get("materials", [])
    processes = data.get("processes", [])
    valid_step_nos = {str(p.get("step_no") or "").strip() for p in processes}

    def write_material_row(m, r):
        """写入一行物料数据，返回下一行号。"""
        row_vals = [
            m.get("name", ""),
            m.get("unit", ""),
            m.get("usage", ""),
            m.get("yield_rate", ""),
            m.get("erp_code", ""),
            m.get("material_type", ""),
            m.get("process", ""),
        ]
        for col, val in enumerate(row_vals, 1):
            cell = ws.cell(r, col, val)
            cell.font = cell_font
            cell.border = border
            cell.alignment = left if col in (1, 2, 5, 6, 7) else center
            if col == 4:
                cell.number_format = '0.0"%"'
        return r + 1

    def write_group_subtitle(text, r):
        """写入分组子标题（A–G 合并，浅色填充），返回下一行号。"""
        ws.merge_cells(f"A{r}:G{r}")
        c = ws.cell(r, 1, text)
        c.font = label_font
        c.fill = group_fill
        c.alignment = left
        return r + 1

    r = 8
    group_enabled = bool(processes) and any(
        str(m.get("process") or "").strip() in valid_step_nos for m in materials
    )

    if group_enabled:
        attributed = {str(p.get("step_no") or "").strip(): [] for p in processes}
        unattributed = []
        for m in materials:
            sn = str(m.get("process") or "").strip()
            if sn in attributed:
                attributed[sn].append(m)
            else:
                unattributed.append(m)
        for p in processes:
            sn = str(p.get("step_no") or "").strip()
            group = attributed.get(sn, [])
            if group:
                r = write_group_subtitle(f"【工序 {sn} {p.get('name', '')}】", r)
                for m in group:
                    r = write_material_row(m, r)
        if unattributed:
            r = write_group_subtitle("【未归属工序】", r)
            for m in unattributed:
                r = write_material_row(m, r)
    else:
        for m in materials:
            r = write_material_row(m, r)

    # 物料区后空行
    r += 1

    # 二、工艺工序
    ws.merge_cells(f"A{r}:G{r}")
    ws.cell(r, 1, "二、工艺工序").font = label_font
    r += 1
    process_headers = ["工序编号", "工序名称", "工序说明", "工时", "备注", "产物"]
    for col, h in enumerate(process_headers, 1):
        cell = ws.cell(r, col, h)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = center
        cell.border = border
    r += 1
    for p in processes:
        row_vals = [
            p.get("step_no", ""),
            p.get("name", ""),
            p.get("desc", ""),
            p.get("work_hours", ""),
            p.get("note", ""),
            p.get("output", ""),
        ]
        for col, val in enumerate(row_vals, 1):
            cell = ws.cell(r, col, val)
            cell.font = cell_font
            cell.border = border
            cell.alignment = left if col in (3, 5) else center
        r += 1

    # 工序区后空行；仅食品类追加「三、配料表」
    if category == "食品":
        r += 1
        ws.merge_cells(f"A{r}:G{r}")
        ws.cell(r, 1, "三、配料表").font = label_font
        r += 1
        ingredient_headers = [
            "物料名称",
            "物料类型",
            "计量单位",
            "用量",
            "出品率(%)",
        ]
        for col, h in enumerate(ingredient_headers, 1):
            cell = ws.cell(r, col, h)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        r += 1
        ingredients, _ = derive_ingredients(data)
        for m in ingredients:
            row_vals = [
                m.get("name", ""),
                m.get("material_type", ""),
                m.get("unit", ""),
                m.get("usage", ""),
                m.get("yield_rate", ""),
            ]
            for col, val in enumerate(row_vals, 1):
                cell = ws.cell(r, col, val)
                cell.font = cell_font
                cell.border = border
                cell.alignment = left if col in (1, 2, 3) else center
                if col == 5:
                    cell.number_format = '0.0"%"'
            r += 1

    # 列宽（A18/B10/C10/D12/E16/F14/G12）
    for i, w in enumerate([18, 10, 10, 12, 16, 14, 12], 1):
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

    # R4 配料表派生提示（非阻断 WARNING，仅食品类）
    category = str(data.get("category") or "其他")
    if category == "食品":
        _, excluded = derive_ingredients(data)
        if excluded:
            names = [str(m.get("name") or "") for m in excluded]
            print(
                "WARNING: 配料表已排除 %d 条非食用物料（包材/其他/未分类）：%s"
                % (len(excluded), "、".join(names))
            )

    wb = build_workbook(data)
    wb.save(args.out)
    print("OK:" + args.out)


if __name__ == "__main__":
    main()
