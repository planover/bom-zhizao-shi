#!/usr/bin/env python3
"""
generate_bom.py - BOM智造师 配套脚本（V4 / V3 / V2.1）

读取包含产品信息、物料与工艺工序信息的 JSON，生成格式化的 BOM 表 Excel (.xlsx)。
依赖 openpyxl；若运行环境缺失则自动安装。

V4 新增能力（相对 V3）：
- BOM 级可选字段 `industry`（8 值枚举，选填，默认按 `category` 推断）。
- 行业专属派生视图：电子→「三、元件清单」（8 列），化工→「三、配方表」（8 列）。
- 物料级专属字段：电子 designator/footprint/part_number/rohs；化工 cas_number/concentration/ghs_hazard。
- 配料表触发条件从 `category == "食品"` 改为 `industry == "食品"`（含推断，行为不变）。
- 软校验：V8（industry 枚举）、W2（RoHS 未标）、W3（CAS/GHS 未填）、含量(%) 列和校验。
- 共享常量迁入 `bom_constants.py`（EDIBLE / ALLERGEN_SET / ALLERGEN_HINTS / V4 新常量）。

V5 增量（相对 V4）：
- 行业专属派生视图扩充：纺织→「三、面料辅料清单」（8 列），家具→「三、家具物料清单」（8 列）。
- 电子「三、元件清单」扩列至 14 列（A–N）：新增 manufacturer/tolerance/rated_power/
  rated_voltage/alternate/reflow_temp 6 个工程/合规字段。
- 化工「三、配方表」扩列至 13 列（A–M）：新增 purity/physical_state/flash_point/
  storage_condition/hazard_class 5 个 SDS 关键字段。
- 跨行业「成本明细」视图（有行业视图时为「四、成本明细」，否则「三、成本明细」，
  8 列含单价/币种/总价派生 + 成本合计行）；物料级新增 unit_price/currency。
- 共享常量迁入 V5 新增：TEXTILE_TYPES/TEXTILE_EXCLUDE/FURNITURE_TYPES/FURNITURE_EXCLUDE/
  EDIBLE_LIST/INDUSTRY_STANDARD(增纺织/家具)/INDUSTRY_TEMPLATES。
- 不新增任何阻断/软校验（保持最小变更、最低回归风险）。

V3（V2.1）能力（沿用）：
- Excel 物料区 8 列（A–H）：首列「序号」全局连续、跨工序分组不重置。
- BOM 级可选字段 `approver` / `effective_date` / `standard`（行 5 单格合并）。
- 物料区末「合计用量」行。
- 食品类「三、配料表」含「用量占比%」（列和恰为 100.0%）与「过敏原」两列。
- `ingredient_pct()` + `check_allergen_soft()`（W1/H1 软告警，非阻断）。
- 全产品出品率显示 `130.0%`。

用法:
    python3 generate_bom.py --data bom.json --out BOM_2026-07-07.xlsx

输入 JSON 结构见 references/bom-spec.md。
"""
import argparse
import datetime
import json
import sys
import subprocess

from bom_constants import (
    INDUSTRIES,
    CATEGORY_TO_INDUSTRY,
    COMPONENT_EXCLUDE,
    FORMULA_EXCLUDE,
    EDIBLE,
    ALLERGEN_SET,
    ALLERGEN_HINTS,
    TEXTILE_TYPES,
    TEXTILE_EXCLUDE,
    FURNITURE_TYPES,
    FURNITURE_EXCLUDE,
    EDIBLE_LIST,
    INDUSTRY_TEMPLATES,
)


# 产品类别枚举（R1 / R4）
CATEGORIES = {"食品", "工业品", "日化化妆品", "医药", "其他"}

# V5 成本币种默认值（currency 选填，缺省为人民币(CNY)）
DEFAULT_CURRENCY = "人民币(CNY)"
# 存在「三、」行业派生视图的行业集合（用于成本视图双编号判定）
INDUSTRY_VIEW_SET = {"食品", "电子", "化工", "纺织", "家具"}


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


def infer_industry(data):
    """推断 industry：显式 > category 推断。返回 (industry, warnings)。

    - industry 已显式设置且合法 → 使用该值，warnings=[]
    - industry 已显式设置但非法 → V8 WARNING，回退为推断值
    - industry 未设置 → 按 category 推断（食品→食品，日化/医药→化工，其他→通用）

    Args:
        data: BOM JSON dict。

    Returns:
        (industry_str, warnings_list)
    """
    industry = str(data.get("industry") or "").strip()
    category = str(data.get("category") or "其他").strip()

    if industry:
        if industry in INDUSTRIES:
            return industry, []
        # V8: 非法值 WARNING（非阻断），回退推断
        warning = (
            "WARNING: industry 值『%s』不在枚举内"
            "（食品/电子/化工/机械/纺织/家具/包装/通用），已回退为推断值"
            % industry
        )
        return CATEGORY_TO_INDUSTRY.get(category, "通用"), [warning]

    # 未填 → 按 category 推断
    return CATEGORY_TO_INDUSTRY.get(category, "通用"), []


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

    说明：V3 的 `approver`/`effective_date`/`standard`/`allergen` 与 V4 的
    `industry`/专属物料字段均为可选，不纳入阻断级校验。
    V8（industry 枚举软校验）由 infer_industry() 返回，非阻断。
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


def derive_ingredients(data, industry=None):
    """派生配料表（仅食品行业）。

    返回 (ingredients, excluded)：
    - ingredients: 可食用物料（material_type ∈ EDIBLE），若 industry != 食品 则为空列表
    - excluded: 非食用物料（包材/其他/未分类/未填），始终保留用于 WARNING 统计

    ingredients 按 usage 降序排列（食品标签惯例）。

    V4 变更：触发条件从 category=="食品" 改为 industry=="食品"（含推断）。
    核心逻辑（过滤 EDIBLE / 排序 / 返回 excluded）完全不变。
    industry=None 时内部调用 infer_industry 推断（向后兼容旧调用）。
    """
    ingredients, excluded = [], []
    for m in data.get("materials", []):
        mt = str(m.get("material_type") or "其他").strip() or "其他"
        if mt in EDIBLE:
            ingredients.append(m)
        else:
            excluded.append(m)

    if industry is None:
        industry, _ = infer_industry(data)
    if industry != "食品":
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


def derive_components(data):
    """派生元件清单（仅电子行业）。

    返回 (components, excluded)：
    - components: 电子元件物料（排除 material_type ∈ COMPONENT_EXCLUDE，即"其他"类）
    - excluded: 被排除的物料（散热片/外壳/包装等）

    排序：按物料类型分组 → 同类型内按位号(designator)字母数字排序。
    空位号排同类型末尾（排序键用 "\\uffff" 哨兵）。
    """
    components, excluded = [], []
    for m in data.get("materials", []):
        mt = str(m.get("material_type") or "其他").strip() or "其他"
        if mt in COMPONENT_EXCLUDE:
            excluded.append(m)
        else:
            components.append(m)

    # 排序：物料类型（升序）→ 位号（字母数字升序，空排末尾）
    components.sort(key=lambda x: (
        str(x.get("material_type") or ""),
        str(x.get("designator") or "\uffff"),
    ))
    return components, excluded


def derive_formula(data):
    """派生配方表（仅化工行业）。

    返回 (formula, excluded)：
    - formula: 配方原料（排除 material_type ∈ FORMULA_EXCLUDE，即"包材"类）
    - excluded: 被排除的物料（瓶子/标签等包材）

    排序：按含量(%) 降序（配方表惯例，主成分在前）。含量为空的排末尾。
    """
    formula, excluded = [], []
    for m in data.get("materials", []):
        mt = str(m.get("material_type") or "其他").strip() or "其他"
        if mt in FORMULA_EXCLUDE:
            excluded.append(m)
        else:
            formula.append(m)

    # 排序：含量(%) 降序，空值排末尾
    formula.sort(key=lambda x: (
        -(float(x.get("concentration") or 0)),
    ))
    return formula, excluded


def derive_textile(data):
    """派生面料辅料清单（仅纺织行业）。

    返回 (items, excluded)：
    - items: 纺织物料（排除 material_type ∈ TEXTILE_EXCLUDE，即"其他"类）
    - excluded: 被排除的物料（包装/其他类）

    排序：物料类型升序 → 物料名称升序（与 V4 电子/化工同构，升序稳定）。
    """
    items, excluded = [], []
    for m in data.get("materials", []):
        mt = str(m.get("material_type") or "其他").strip() or "其他"
        if mt in TEXTILE_EXCLUDE:
            excluded.append(m)
        else:
            items.append(m)

    # 排序：物料类型（升序）→ 物料名称（升序，空排末尾）
    items.sort(key=lambda x: (
        str(x.get("material_type") or ""),
        str(x.get("name") or ""),
    ))
    return items, excluded


def derive_furniture(data):
    """派生家具物料清单（仅家具行业）。

    返回 (items, excluded)：
    - items: 家具物料（排除 material_type ∈ FURNITURE_EXCLUDE，即"其他"类）
    - excluded: 被排除的物料

    排序：物料类型升序 → 物料名称升序。
    """
    items, excluded = [], []
    for m in data.get("materials", []):
        mt = str(m.get("material_type") or "其他").strip() or "其他"
        if mt in FURNITURE_EXCLUDE:
            excluded.append(m)
        else:
            items.append(m)

    # 排序：物料类型（升序）→ 物料名称（升序，空排末尾）
    items.sort(key=lambda x: (
        str(x.get("material_type") or ""),
        str(x.get("name") or ""),
    ))
    return items, excluded


def derive_cost(data):
    """派生成本明细（跨行业，任一物料 unit_price 非空即纳入）。

    返回 (cost_items, has_cost)：
    - cost_items: unit_price 非空（≠"" 且可转 float）的物料列表（含被行业视图
      排除的"其他"/"包材"类，成本核算面向全物料）
    - has_cost: 列表非空即为 True（触发生成成本视图）

    注意：total_price 不存 JSON，渲染时按 usage × unit_price 实时计算（派生展示）。
    """
    cost_items = []
    for m in data.get("materials", []):
        up = m.get("unit_price", "")
        if up not in ("", None):
            try:
                if float(up) >= 0:
                    cost_items.append(m)
            except (TypeError, ValueError):
                pass
    return cost_items, bool(cost_items)


def check_industry_soft(data, industry):
    """行业专属软校验（W2/W3，非阻断 WARNING）。

    - W2：industry=="电子" 且物料（未排除的）未标 rohs → WARNING
    - W3：industry=="化工" 且配方原料未填 cas_number 或 ghs_hazard → WARNING
    - 含量(%) 列和校验：所有配方原料均填了 concentration → 校验列和 ≈ 100%（±5%）→ 不达标 WARNING

    Args:
        data: BOM JSON dict。
        industry: 已推断的行业字符串。

    Returns:
        warnings 字符串列表（非阻断）。
    """
    warnings = []
    if industry == "电子":
        components, _ = derive_components(data)
        for m in components:
            rohs = str(m.get("rohs") or "").strip()
            if not rohs:
                warnings.append(
                    "WARNING: 物料『%s』未标注 RoHS 合规状态，请确认"
                    % m.get("name", "")
                )
    elif industry == "化工":
        formula, _ = derive_formula(data)
        for m in formula:
            cas = str(m.get("cas_number") or "").strip()
            ghs = str(m.get("ghs_hazard") or "").strip()
            if not cas:
                warnings.append(
                    "WARNING: 物料『%s』未填写 CAS 号，请确认"
                    % m.get("name", "")
                )
            if not ghs:
                warnings.append(
                    "WARNING: 物料『%s』未填写 GHS 危险标识，请确认"
                    % m.get("name", "")
                )
        # 含量(%) 列和校验：仅当所有配方原料均填了 concentration（非空且 > 0）时才校验
        concs = [float(m.get("concentration") or 0) for m in formula]
        if concs and all(c > 0 for c in concs):
            total = sum(concs)
            if abs(total - 100.0) > 5.0:
                warnings.append(
                    "WARNING: 配方表含量(%%) 列和为 %.1f%%，偏离 100%% 超过 ±5%%，请确认"
                    % total
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


def build_workbook(data, industry=None):
    """根据数据构建 BOM 表 Workbook（8 列 A–H，固定行号）。

    V4：按 industry 分支在工序区后追加专属派生视图：
    - 食品 → 「三、配料表」（7 列，沿用 V3 逻辑）
    - 电子 → 「三、元件清单」（8 列，含 RoHS 红黄字标记）
    - 化工 → 「三、配方表」（8 列，含含量(%) 0.0"%" 格式）
    - 其他 → 不生成专属视图

    Args:
        data: BOM JSON dict。
        industry: 已推断的行业字符串；None 时内部调 infer_industry。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    if industry is None:
        industry, _ = infer_industry(data)

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
    # V4: RoHS 合规标记字体（红色=不合规，黄色=待确认）
    rohs_font_red = Font(name="微软雅黑", size=10, color="FF0000")
    rohs_font_yellow = Font(name="微软雅黑", size=10, color="BF8F00")
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

    # ===== V4: 工序区后按 industry 分支追加专属派生视图 =====

    if industry == "食品":
        # 三、配料表（沿用 V3 逻辑，触发改 industry）
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
        ingredients, _ = derive_ingredients(data, industry)
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

    elif industry == "电子":
        # 三、元件清单（★V4 新增，★V5 扩列至 14 列 A–N）
        r += 1
        ws.merge_cells(f"A{r}:N{r}")
        ws.cell(r, 1, "三、元件清单").font = label_font
        r += 1
        component_headers = [
            "序号",
            "位号(Designator)",
            "型号(Part#)",
            "封装(Footprint)",
            "物料名称",
            "数量",
            "物料类型",
            "RoHS",
            "制造商",
            "容差",
            "额定功率",
            "额定电压",
            "替代料",
            "封装温度",
        ]
        for col, h in enumerate(component_headers, 1):
            cell = ws.cell(r, col, h)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        r += 1
        components, _ = derive_components(data)
        for idx, m in enumerate(components, 1):
            rohs_val = str(m.get("rohs") or "").strip()
            # RoHS 着色规则：否→红字，未知/空→黄字，是→默认
            if rohs_val == "否":
                rohs_font = rohs_font_red
            elif rohs_val == "是":
                rohs_font = cell_font
            else:
                rohs_font = rohs_font_yellow

            row_vals = [
                idx,
                m.get("designator", ""),
                m.get("part_number", ""),
                m.get("footprint", ""),
                m.get("name", ""),
                m.get("usage", ""),
                m.get("material_type", ""),
                rohs_val,
                m.get("manufacturer", ""),
                m.get("tolerance", ""),
                m.get("rated_power", ""),
                m.get("rated_voltage", ""),
                m.get("alternate", ""),
                m.get("reflow_temp", ""),
            ]
            # 对齐：序号/位号/封装/数量/物料类型/RoHS/容差/额定功率/额定电压/替代料/封装温度 居中；
            # 型号/物料名称/制造商 左对齐
            aligns = [
                center, center, left, center, left, center, center, center,
                left, center, center, center, left, center,
            ]
            for col, val in enumerate(row_vals, 1):
                cell = ws.cell(r, col, val)
                cell.border = border
                cell.alignment = aligns[col - 1]
                if col == 8:  # RoHS 列使用特定字体
                    cell.font = rohs_font
                else:
                    cell.font = cell_font
            r += 1

    elif industry == "化工":
        # 三、配方表（★V4 新增，★V5 扩列至 13 列 A–M）
        r += 1
        ws.merge_cells(f"A{r}:M{r}")
        ws.cell(r, 1, "三、配方表").font = label_font
        r += 1
        formula_headers = [
            "序号",
            "物料名称",
            "CAS号",
            "含量(%)",
            "GHS标识",
            "物料类型",
            "计量单位",
            "用量",
            "纯度",
            "物态",
            "闪点",
            "存储条件",
            "危险等级",
        ]
        for col, h in enumerate(formula_headers, 1):
            cell = ws.cell(r, col, h)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        r += 1
        formula, _ = derive_formula(data)
        for idx, m in enumerate(formula, 1):
            conc_raw = m.get("concentration", "")
            # 含量(%)：数值则写入 float + 0.0"%" 格式；空则留空
            if conc_raw != "" and conc_raw is not None:
                try:
                    conc_val = float(conc_raw)
                except (TypeError, ValueError):
                    conc_val = str(conc_raw)
            else:
                conc_val = ""

            row_vals = [
                idx,
                m.get("name", ""),
                m.get("cas_number", ""),
                conc_val,
                m.get("ghs_hazard", ""),
                m.get("material_type", ""),
                m.get("unit", ""),
                m.get("usage", ""),
                m.get("purity", ""),
                m.get("physical_state", ""),
                m.get("flash_point", ""),
                m.get("storage_condition", ""),
                m.get("hazard_class", ""),
            ]
            # 对齐：序号/CAS号/含量/GHS/物料类型/计量单位/用量/纯度/物态/闪点/危险等级 居中；
            # 物料名称/存储条件 左对齐
            aligns = [
                center, left, center, center, left, center, center, center,
                center, center, center, left, center,
            ]
            for col, val in enumerate(row_vals, 1):
                cell = ws.cell(r, col, val)
                cell.font = cell_font
                cell.border = border
                cell.alignment = aligns[col - 1]
                if col == 4 and conc_val != "":  # 含量(%) 数字格式
                    cell.number_format = pct_fmt
            r += 1

    elif industry == "纺织":
        # 三、面料辅料清单（★V5 新增，8 列 A–H）
        r += 1
        ws.merge_cells(f"A{r}:H{r}")
        ws.cell(r, 1, "三、面料辅料清单").font = label_font
        r += 1
        textile_headers = [
            "序号", "物料名称", "物料类型", "成分比例", "纱支",
            "克重(g/m²)", "幅宽", "色号",
        ]
        for col, h in enumerate(textile_headers, 1):
            cell = ws.cell(r, col, h)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        r += 1
        textile_items, _ = derive_textile(data)
        for idx, m in enumerate(textile_items, 1):
            row_vals = [
                idx,
                m.get("name", ""),
                m.get("material_type", ""),
                m.get("composition", ""),
                m.get("yarn_count", ""),
                m.get("fabric_weight", ""),
                m.get("width", ""),
                m.get("color_no", ""),
            ]
            # 对齐：序号/物料类型/纱支/克重/幅宽/色号 居中；物料名称/成分比例 左对齐
            aligns = [center, left, center, left, center, center, center, center]
            for col, val in enumerate(row_vals, 1):
                cell = ws.cell(r, col, val)
                cell.font = cell_font
                cell.border = border
                cell.alignment = aligns[col - 1]
            r += 1

    elif industry == "家具":
        # 三、家具物料清单（★V5 新增，8 列 A–H）
        r += 1
        ws.merge_cells(f"A{r}:H{r}")
        ws.cell(r, 1, "三、家具物料清单").font = label_font
        r += 1
        furniture_headers = [
            "序号", "物料名称", "物料类型", "材质等级", "尺寸规格",
            "表面处理", "用量", "色号/花色",
        ]
        for col, h in enumerate(furniture_headers, 1):
            cell = ws.cell(r, col, h)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        r += 1
        furniture_items, _ = derive_furniture(data)
        for idx, m in enumerate(furniture_items, 1):
            row_vals = [
                idx,
                m.get("name", ""),
                m.get("material_type", ""),
                m.get("material_grade", ""),
                m.get("spec_size", ""),
                m.get("surface_treatment", ""),
                m.get("usage", ""),
                m.get("color_no", ""),
            ]
            # 对齐：序号/物料类型/材质等级/用量/色号 居中；物料名称/尺寸规格/表面处理 左对齐
            aligns = [center, left, center, center, left, left, center, center]
            for col, val in enumerate(row_vals, 1):
                cell = ws.cell(r, col, val)
                cell.font = cell_font
                cell.border = border
                cell.alignment = aligns[col - 1]
            r += 1

    # 其他行业（通用/机械/包装）不生成「三、」行业专属视图

    # ===== V5: 跨行业成本明细视图（双编号） =====
    cost_items, has_cost = derive_cost(data)
    if has_cost:
        # 有行业视图（食品/电子/化工/纺织/家具）→ 四、成本明细；否则 三、成本明细
        cost_label = "四、成本明细" if industry in INDUSTRY_VIEW_SET else "三、成本明细"
        r += 1
        ws.merge_cells(f"A{r}:H{r}")
        ws.cell(r, 1, cost_label).font = label_font
        r += 1
        cost_headers = [
            "序号", "物料名称", "物料类型", "用量", "单位", "单价", "币种", "总价",
        ]
        for col, h in enumerate(cost_headers, 1):
            cell = ws.cell(r, col, h)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        r += 1
        cost_total = 0.0
        for idx, m in enumerate(cost_items, 1):
            up_raw = m.get("unit_price", "")
            usage_raw = m.get("usage", "")
            try:
                up_f = float(up_raw) if up_raw not in ("", None) else 0.0
                usage_f = float(usage_raw) if usage_raw not in ("", None) else 0.0
                total = round(up_f * usage_f, 2)
            except (TypeError, ValueError):
                up_f = 0.0
                total = ""
            if isinstance(total, (int, float)):
                cost_total += total
            currency_val = str(m.get("currency") or "").strip() or DEFAULT_CURRENCY
            row_vals = [
                idx,
                m.get("name", ""),
                m.get("material_type", ""),
                m.get("usage", ""),
                m.get("unit", ""),
                up_f if isinstance(up_f, (int, float)) else "",
                currency_val,
                total,
            ]
            # 对齐：序号/物料类型/用量/单位/单价/总价 居中；物料名称 左对齐；币种 左对齐
            aligns = [center, left, center, center, center, center, left, center]
            for col, val in enumerate(row_vals, 1):
                cell = ws.cell(r, col, val)
                cell.font = cell_font
                cell.border = border
                cell.alignment = aligns[col - 1]
                if col in (6, 8) and isinstance(val, (int, float)):
                    cell.number_format = "0.00"  # 单价/总价 数值格式
            r += 1
        # 成本合计行：A="成本合计"，H=Σ总价（纯展示，逆向跳过）
        ws.cell(r, 1, "成本合计").font = label_font
        ws.cell(r, 1).alignment = left
        ws.cell(r, 1).border = border
        tc = ws.cell(r, 8, round(cost_total, 2))
        tc.font = label_font
        tc.alignment = center
        tc.border = border
        for col in (2, 3, 4, 5, 6, 7):
            ws.cell(r, col).border = border
        r += 1

    # 列宽（依据 industry 与扩列扩展）
    # 电子扩列到 14 列（A–N），化工扩列到 13 列（A–M），其余 8 列（A–H）
    if industry == "电子":
        widths = [6, 18, 18, 12, 18, 10, 13, 10, 14, 12, 14, 14, 16, 14]
    elif industry == "化工":
        widths = [6, 18, 12, 13, 14, 13, 12, 10, 10, 12, 12, 18, 14]
    else:
        widths = [6, 18, 10, 10, 13, 16, 13, 12]
    for i, w in enumerate(widths, 1):
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

    # V4: 推断 industry
    industry, v8_warnings = infer_industry(data)
    for w in v8_warnings:
        print(w)

    # 行业专属软校验 + 排除提示（非阻断 WARNING）
    if industry == "食品":
        # W1/H1 过敏原软校验
        for w in check_allergen_soft(data):
            print(w)
        # 配料表排除提示
        _, excluded = derive_ingredients(data, industry)
        if excluded:
            names = [str(m.get("name") or "") for m in excluded]
            print(
                "WARNING: 配料表已排除 %d 条非食用物料（包材/其他/未分类）：%s"
                % (len(excluded), "、".join(names))
            )
    elif industry == "电子":
        # W2 RoHS 软校验
        for w in check_industry_soft(data, industry):
            print(w)
        # 元件清单排除提示
        _, excluded = derive_components(data)
        if excluded:
            names = [str(m.get("name") or "") for m in excluded]
            print(
                "WARNING: 元件清单已排除 %d 条非元件物料（其他类）：%s"
                % (len(excluded), "、".join(names))
            )
    elif industry == "化工":
        # W3 CAS/GHS 软校验 + 含量和校验
        for w in check_industry_soft(data, industry):
            print(w)
        # 配方表排除提示
        _, excluded = derive_formula(data)
        if excluded:
            names = [str(m.get("name") or "") for m in excluded]
            print(
                "WARNING: 配方表已排除 %d 条包材物料：%s"
                % (len(excluded), "、".join(names))
            )
    elif industry == "纺织":
        # V5: 面料辅料清单排除提示（沿用 V4 排除提示模式，文案同构）
        _, excluded = derive_textile(data)
        if excluded:
            names = [str(m.get("name") or "") for m in excluded]
            print(
                "WARNING: 面料辅料清单已排除 %d 条其他类物料：%s"
                % (len(excluded), "、".join(names))
            )
    elif industry == "家具":
        # V5: 家具物料清单排除提示（沿用 V4 排除提示模式，文案同构）
        _, excluded = derive_furniture(data)
        if excluded:
            names = [str(m.get("name") or "") for m in excluded]
            print(
                "WARNING: 家具物料清单已排除 %d 条其他类物料：%s"
                % (len(excluded), "、".join(names))
            )

    wb = build_workbook(data, industry)
    wb.save(args.out)
    print("OK:" + args.out)


if __name__ == "__main__":
    main()
