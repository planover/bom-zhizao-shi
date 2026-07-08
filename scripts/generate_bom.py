#!/usr/bin/env python3
"""
generate_bom.py - BOM智造师 配套脚本（V3 / V2.1）

读取包含产品信息、物料与工艺工序信息的 JSON，生成格式化的 BOM 表 Excel (.xlsx)。
依赖 openpyxl；若运行环境缺失则自动安装。

V3（V2.1）新增能力（相对 V2）：
- Excel 物料区由 7 列扩展为 **8 列（A–H）**：首列新增「序号」（按输入顺序全局连续、跨工序分组不重置，纯展示不进 JSON）。
- BOM 级可选字段 `approver` / `effective_date` / `standard`：Excel 行 5 单格合并 A5:H5，拼接非空字段（审批人 / 生效日期 / 执行标准），三者皆空则整行留空。
- 物料区末新增「合计用量」行：A 列写「合计」，D 列写全部物料 usage 求和（含包材/其他），纯展示、逆向跳过。
- 食品类「三、配料表」新增「用量占比%」（最大余数法保证列和恰为 100.0%）与「过敏原」两列（仅食品展示，按物料名取 `allergen`）。
- 新增 `ingredient_pct()`（占比% 计算）与 `check_allergen_soft()`（W1 标签合法性 + H1 名称关键词启发式软告警，均非阻断）。
- 分组子标题美化（浅蓝底 + 左侧加粗色条 + 加粗）。
- 全产品出品率仍显示 `130.0%`（1 位小数，沿用 V2 修正）。

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
# 过敏原八大类 + 其他（GB 7718-2025），用于 W1 软校验
ALLERGEN_SET = {
    "含麸质谷物",
    "甲壳类",
    "蛋类",
    "鱼类",
    "花生",
    "大豆",
    "乳",
    "坚果",
    "其他",
}

# 过敏原关键词启发式（H1 软告警，非阻断）：(名称子串, 对应八大类)
# 用于在「名称疑似含致敏物但 allergen 未标注」时给出提示性 WARNING。
ALLERGEN_HINTS = [
    ("牛奶", "乳"),
    ("奶", "乳"),
    ("奶酪", "乳"),
    ("黄油", "乳"),
    ("蛋黄", "蛋类"),
    ("蛋清", "蛋类"),
    ("蛋白", "蛋类"),
    ("蛋", "蛋类"),
    ("花生", "花生"),
    ("大豆", "大豆"),
    ("黄豆", "大豆"),
    ("豆浆", "大豆"),
    ("豆腐", "大豆"),
    ("麸质", "含麸质谷物"),
    ("面筋", "含麸质谷物"),
    ("小麦", "含麸质谷物"),
    ("面粉", "含麸质谷物"),
    ("坚果", "坚果"),
    ("杏仁", "坚果"),
    ("腰果", "坚果"),
    ("核桃", "坚果"),
    ("花生酱", "花生"),
    ("虾", "甲壳类"),
    ("蟹", "甲壳类"),
    ("龙虾", "甲壳类"),
    ("鱼", "鱼类"),
    ("三文鱼", "鱼类"),
    ("鳕鱼", "鱼类"),
]


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

    说明：V3 新增的 `approver`/`effective_date`/`standard`/`allergen` 均为可选字符串，
    不纳入阻断级校验；仅 `allergen` 由 check_allergen_soft() 做非阻断软告警。
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


def ingredient_pct(items):
    """计算可食用物料的用量占比%（1 位小数），保证列和恰为 100.0%。

    分母 = 各物料 usage 求和；每项 round(usage / 合计 × 100, 1)。
    使用最大余数法对末位补差（0.1 粒度），使整列合计严格等于 100.0。

    Args:
        items: 可食用物料列表，每项含 "usage" 键。

    Returns:
        (pct_list, total)：pct_list 与输入 items 顺序一致；total 为用量合计。
    """
    total = sum(float(it.get("usage", 0)) for it in items) or 1.0
    raw = [float(it.get("usage", 0)) / total * 100 for it in items]
    pct = [round(x, 1) for x in raw]
    s = round(sum(pct), 1)
    diff = round(100.0 - s, 1)
    if diff != 0:
        # 按最大小数余数分配补差（0.1 粒度）
        order = sorted(range(len(raw)), key=lambda i: (raw[i] - pct[i]), reverse=True)
        step = 0.1 if diff > 0 else -0.1
        d = abs(diff)
        for i in order:
            if round(d, 1) <= 0:
                break
            pct[i] = round(pct[i] + step, 1)
            d = round(d - 0.1, 1)
    return pct, total


def check_allergen_soft(data):
    """校验 allergen：标签合法性(W1) + 名称关键词启发式(H1)，均非阻断。

    返回 WARNING 文本列表（不进入 errors，不改变退出码）：
    - W1：标签非空且任一标签 ∉ ALLERGEN_SET → 告警「过敏原标签不在八大类集合」。
    - H1：名称含致敏物关键词但其 allergen 未涵盖对应类别 → 告警
          「名称疑似含致敏物『{class}』但未在过敏原中标注，请确认」。
    """
    warnings = []
    for i, m in enumerate(data.get("materials", []), 1):
        name = str(m.get("name") or f"#{i}").strip()
        raw = str(m.get("allergen") or "").strip()
        # allergen 可能为空串，按逗号拆分得到标签集合（去空）
        tags = [t.strip() for t in raw.split(",") if t.strip()] if raw else []
        tag_set = set(tags)

        # W1：标签合法性校验（仅当标签非空且存在非法标签）
        bad = [t for t in tags if t not in ALLERGEN_SET]
        if bad:
            warnings.append(
                "WARNING: 物料『%s』的过敏原标签『%s』不在八大类集合"
                "（含麸质谷物/甲壳类/蛋类/鱼类/花生/大豆/乳/坚果/其他）内，请确认"
                % (name, "、".join(bad))
            )

        # H1：名称关键词启发式（非阻断，提示性 WARNING）
        # 名称含致敏物关键词，但对应八大类未在该物料 allergen 中标注 → 告警。
        # 同一物料对同一类别仅告警一次；已正确标注的不告警；无关键词的不告警。
        if name:
            warned = set()
            for kw, cls in ALLERGEN_HINTS:
                if kw in name and cls not in tag_set and cls not in warned:
                    warned.add(cls)
                    warnings.append(
                        "WARNING: 物料『%s』名称疑似含致敏物『%s』但未在过敏原中标注，请确认"
                        % (name, cls)
                    )
    return warnings


def _build_meta_line(approver, effective_date, standard):
    """拼接行 5 表头区文本：仅拼接非空字段，段间用 4 个空格分隔。

    三者皆空返回空串（等价于 V2 空行，不挤占行 6）。
    """
    parts = []
    if approver:
        parts.append("审批人：" + str(approver))
    if effective_date:
        parts.append("生效日期：" + str(effective_date))
    if standard:
        parts.append("执行标准：" + str(standard))
    return "    ".join(parts)


def build_workbook(data):
    """根据数据构建 BOM 表 Workbook（8 列 A–H，固定行号）。"""
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
    bar_side = Side(style="medium", color="1F3864")
    group_border = Border(left=bar_side, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    pct_fmt = '0.0"%"'

    version = data.get("version") or "V1.0"
    date = data.get("date") or datetime.date.today().isoformat()
    product_name = str(data.get("product_name") or "").strip()
    category = str(data.get("category") or "其他").strip()
    output_rate = data.get("output_rate")
    approver = str(data.get("approver") or "").strip()
    effective_date = str(data.get("effective_date") or "").strip()
    standard = str(data.get("standard") or "").strip()

    # 行1：标题
    ws.merge_cells("A1:H1")
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
    ws.merge_cells("D2:H2")

    # 行3：产品名称（整行）
    ws.merge_cells("A3:H3")
    c = ws.cell(3, 1, f"产品名称：{product_name}")
    c.font = label_font

    # 行4：产品类别（左） / 全产品出品率（右，0.0"%" 格式）
    ws.merge_cells("A4:C4")
    c = ws.cell(4, 1, f"产品类别：{category}")
    c.font = label_font
    ws.merge_cells("D4:H4")
    c = ws.cell(4, 4, f"全产品出品率：{float(output_rate):.1f}%")
    c.font = label_font

    # 行5：审批人 / 生效日期 / 执行标准（可选，单格合并 A5:H5；皆空则留空）
    meta_line = _build_meta_line(approver, effective_date, standard)
    ws.merge_cells("A5:H5")
    c = ws.cell(5, 1, meta_line)
    c.font = label_font
    c.alignment = left

    # 行6：一、物料信息
    ws.merge_cells("A6:H6")
    ws.cell(6, 1, "一、物料信息").font = label_font

    # 行7：物料表头（8 列：序号|物料名称|单位|用量|出品率(%)|ERP物料代码|物料类型|所属工序）
    material_headers = [
        "序号",
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

    def write_material_row(m, r, seq):
        """写入一行物料数据（含首列序号），返回下一行号。"""
        row_vals = [
            seq,
            m.get("name", ""),
            m.get("unit", ""),
            m.get("usage", ""),
            m.get("yield_rate", ""),
            m.get("erp_code", ""),
            m.get("material_type", ""),
            m.get("process", ""),
        ]
        # 列对齐：序号/单位/用量/出品率/所属工序居中；物料名称/ERP/物料类型左对齐
        aligns = [center, left, center, center, center, left, left, center]
        for col, val in enumerate(row_vals, 1):
            cell = ws.cell(r, col, val)
            cell.font = cell_font
            cell.border = border
            cell.alignment = aligns[col - 1]
            if col == 5:  # 出品率(%)
                cell.number_format = pct_fmt
        return r + 1

    def write_group_subtitle(text, r):
        """写入分组子标题（A–H 合并，浅蓝底 + 左侧加粗色条 + 加粗），返回下一行号。"""
        ws.merge_cells(f"A{r}:H{r}")
        c = ws.cell(r, 1, text)
        c.font = label_font
        c.fill = group_fill
        c.alignment = left
        c.border = group_border
        return r + 1

    r = 8
    group_enabled = bool(processes) and any(
        str(m.get("process") or "").strip() in valid_step_nos for m in materials
    )
    seq = 0
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
                    seq += 1
                    r = write_material_row(m, r, seq)
        if unattributed:
            r = write_group_subtitle("【未归属工序】", r)
            for m in unattributed:
                seq += 1
                r = write_material_row(m, r, seq)
    else:
        for m in materials:
            seq += 1
            r = write_material_row(m, r, seq)

    # 物料区末「合计用量」行：A=合计，D=全部物料 usage 求和（含包材/其他）
    total_usage = sum(float(m.get("usage") or 0) for m in materials)
    ws.cell(r, 1, "合计").font = label_font
    ws.cell(r, 1).alignment = left
    ws.cell(r, 1).border = border
    tc = ws.cell(r, 4, round(total_usage, 1))
    tc.font = label_font
    tc.alignment = center
    tc.border = border
    for col in (2, 3, 5, 6, 7, 8):
        ws.cell(r, col).border = border
    r += 1

    # 物料区后空行
    r += 1

    # 二、工艺工序
    ws.merge_cells(f"A{r}:H{r}")
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
        ws.merge_cells(f"A{r}:H{r}")
        ws.cell(r, 1, "三、配料表").font = label_font
        r += 1
        ingredient_headers = [
            "物料名称",
            "物料类型",
            "计量单位",
            "用量",
            "出品率(%)",
            "用量占比%",
            "过敏原",
        ]
        for col, h in enumerate(ingredient_headers, 1):
            cell = ws.cell(r, col, h)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        r += 1
        ingredients, _ = derive_ingredients(data)
        pct_list, _ = ingredient_pct(ingredients)
        for idx, m in enumerate(ingredients):
            row_vals = [
                m.get("name", ""),
                m.get("material_type", ""),
                m.get("unit", ""),
                m.get("usage", ""),
                m.get("yield_rate", ""),
                pct_list[idx],
                m.get("allergen", ""),
            ]
            # 对齐：物料名称/物料类型/过敏原左对齐；其余居中
            aligns = [left, left, center, center, center, center, left]
            for col, val in enumerate(row_vals, 1):
                cell = ws.cell(r, col, val)
                cell.font = cell_font
                cell.border = border
                cell.alignment = aligns[col - 1]
                if col in (5, 6):  # 出品率(%) 与 用量占比%
                    cell.number_format = pct_fmt
            r += 1

    # 列宽（8 列 A–H）：序号/物料名称/单位/用量/出品率/ERP代码/物料类型/所属工序
    for i, w in enumerate([6, 18, 10, 10, 13, 16, 13, 12], 1):
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

    # W1 过敏原软校验（非阻断 WARNING，不进 errors）
    for w in check_allergen_soft(data):
        print(w)

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
