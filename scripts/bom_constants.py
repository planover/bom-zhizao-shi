#!/usr/bin/env python3
"""
bom_constants.py - BOM智造师 共享常量模块（V5 / V4）

定义行业枚举、推断映射、元件/配方/纺织/家具过滤集、执行标准建议、
行业模板预设（仅交互引导），以及从 generate_bom.py 迁入的食品配料表
过滤集（EDIBLE）与过敏原集合/关键词（ALLERGEN_SET / ALLERGEN_HINTS）。
generate_bom.py 与 import_bom.py 均从此模块导入，保持单一真相源，避免
两文件各自复制漂移。

V5 增量：新增纺织/家具过滤集与建议类型、导出 EDIBLE_LIST、INDUSTRY_STANDARD
增 纺织=FZ/T 80004 与 家具=QB/T 1951.1、新增 INDUSTRY_TEMPLATES 行业模板预设
（仅交互引导，不写入新 JSON 结构字段）。
V6 增量：新增机械/包装过滤集与建议类型、INDUSTRY_STANDARD 增 机械=GB/T 1804-2000
与 包装=GB/T 6543-2008、INDUSTRY_TEMPLATES 机械/包装由空模板填实（仅交互引导）。

纯 Python 标准库模块，无第三方依赖。
"""

# === V4 行业枚举 ===
INDUSTRIES = {"食品", "电子", "化工", "机械", "纺织", "家具", "包装", "通用"}

# category → industry 推断映射（向后兼容核心）
# 现有食品 JSON 无 industry → 推断"食品" → 配料表照常（与 V3 一致）；
# 现有非食品 JSON → 推断"通用" → 不生成专属视图（与 V3 一致）。
CATEGORY_TO_INDUSTRY = {
    "食品": "食品",
    "日化化妆品": "化工",
    "医药": "化工",
    "工业品": "通用",
    "其他": "通用",
}

# === 电子行业 ===
# 元件物料类型建议值（交互引导，非强制枚举）
COMPONENT_TYPES = ["电阻", "电容", "IC", "连接器", "二极管", "三极管", "晶振", "其他"]
# 元件清单过滤排除集（material_type 在此集合内的物料不进元件清单）
COMPONENT_EXCLUDE = {"其他"}

# === 化工行业 ===
# 配方原料物料类型建议值
FORMULA_TYPES = ["主料", "溶剂", "催化剂", "添加剂", "包材", "其他"]
# 配方表过滤排除集
FORMULA_EXCLUDE = {"包材"}

# === 执行标准行业建议（P1 / V5 增量 / V6 增量） ===
INDUSTRY_STANDARD = {
    "食品": "GB 7718-2025",
    "电子": "GB/T 39560",
    "化工": "GB/T 16483-2008",
    "纺织": "FZ/T 80004",      # V5 新增
    "家具": "QB/T 1951.1",     # V5 新增
    "机械": "GB/T 1804-2000",  # V6 新增（一般公差，默认可改）
    "包装": "GB/T 6543-2008",  # V6 新增（运输包装用瓦楞纸箱，默认可改）
}

# === 食品配料表过滤集（从 generate_bom.py 迁入，单一真相源） ===
# 可食用物料类型（R4 配料表过滤）
EDIBLE = {"原料", "添加剂", "香精香料"}
# V5：EDIBLE 集合转列表，供 INDUSTRY_TEMPLATES 食品模板 special_fields 引用
EDIBLE_LIST = ["原料", "添加剂", "香精香料"]

# === V5 纺织行业 ===
# 面料辅料物料类型建议值（交互引导，非强制枚举）
TEXTILE_TYPES = ["面料", "辅料", "纱线", "印染", "五金", "其他"]
# 面料辅料清单过滤排除集（material_type 在此集合内的物料不进面料辅料清单）
TEXTILE_EXCLUDE = {"其他"}

# === V5 家具行业 ===
# 家具物料类型建议值（交互引导，非强制枚举）
FURNITURE_TYPES = ["主材", "板材", "辅材", "五金", "面料", "其他"]
# 家具物料清单过滤排除集（material_type 在此集合内的物料不进家具物料清单）
FURNITURE_EXCLUDE = {"其他"}

# === V6 机械行业 ===
# 机械物料类型建议值（交互引导，非强制枚举）
MECHANICAL_TYPES = ["零部件", "标准件", "型材", "铸件", "焊接件", "其他"]
# 机械物料清单过滤排除集（material_type 在此集合内的物料不进机械物料清单）
MECHANICAL_EXCLUDE = {"其他"}

# === V6 包装行业 ===
# 包装物料类型建议值（交互引导，非强制枚举）
PACKAGING_TYPES = ["纸箱", "缓冲", "标签", "胶带", "薄膜", "其他"]
# 包装物料清单过滤排除集（material_type 在此集合内的物料不进包装物料清单）
PACKAGING_EXCLUDE = {"其他"}

# === 过敏原八大类 + 其他（GB 7718-2025），用于 W1 软校验 ===
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

# === V5 行业模板预设（仅交互引导，不写入新 JSON 结构字段） ===
# 仅被 SKILL.md 交互层读取，用于：
#   1. 阶段一物料模板按 special_fields 动态追加专属字段行；
#   2. material_type 下拉用 material_types；
#   3. standard 自动预填（可改）；
#   4. 可选「一键载入工序模板」（preset_processes）。
# JSON Schema 不新增任何结构字段，向后兼容旧交互与旧 JSON。
INDUSTRY_TEMPLATES = {
    "电子": {
        "material_types": COMPONENT_TYPES,
        "standard": "GB/T 39560",
        "special_fields": [
            "designator", "footprint", "part_number", "rohs",
            "manufacturer", "tolerance", "rated_power",
            "rated_voltage", "alternate", "reflow_temp",
        ],
        "preset_processes": [
            {"step_no": "S01", "name": "SMT贴片", "desc": "锡膏印刷+贴片",
             "output": "贴片完成板"},
            {"step_no": "S02", "name": "回流焊", "desc": "回流焊接",
             "output": "焊接板"},
            {"step_no": "S03", "name": "检测", "desc": "AOI/功能测试",
             "output": "成品板"},
        ],
    },
    "化工": {
        "material_types": FORMULA_TYPES,
        "standard": "GB/T 16483-2008",
        "special_fields": [
            "cas_number", "concentration", "ghs_hazard",
            "purity", "physical_state", "flash_point",
            "storage_condition", "hazard_class",
        ],
        "preset_processes": [
            {"step_no": "S01", "name": "投料混合", "desc": "按比例投料并搅拌",
             "output": "混合液"},
            {"step_no": "S02", "name": "灌装", "desc": "灌装入容器",
             "output": "成品"},
        ],
    },
    "纺织": {
        "material_types": TEXTILE_TYPES,
        "standard": "FZ/T 80004",
        "special_fields": [
            "composition", "yarn_count", "fabric_weight", "width", "color_no",
        ],
        "preset_processes": [
            {"step_no": "S01", "name": "裁剪", "desc": "按版型裁剪",
             "output": "裁片"},
            {"step_no": "S02", "name": "缝制", "desc": "缝纫组合",
             "output": "半成品"},
            {"step_no": "S03", "name": "整烫检验", "desc": "整烫+质检",
             "output": "成品"},
        ],
    },
    "家具": {
        "material_types": FURNITURE_TYPES,
        "standard": "QB/T 1951.1",
        "special_fields": [
            "material_grade", "spec_size", "surface_treatment", "color_no",
        ],
        "preset_processes": [
            {"step_no": "S01", "name": "开料", "desc": "板材锯切",
             "output": "板件"},
            {"step_no": "S02", "name": "封边", "desc": "边部封边",
             "output": "封边板件"},
            {"step_no": "S03", "name": "组装", "desc": "五金组装",
             "output": "成品"},
        ],
    },
    "食品": {
        "material_types": EDIBLE_LIST,
        "standard": "GB 7718-2025",
        "special_fields": ["allergen"],
        "preset_processes": [],
    },
    # 通用：空模板（保持通用兜底，不预置专属字段与工序）
    "通用": {"material_types": [], "standard": "", "special_fields": [],
             "preset_processes": []},
    # V6 机械：填实模板（仅交互引导，不写入 JSON 结构；special_fields 含与家具同名的 surface_treatment）
    "机械": {
        "material_types": MECHANICAL_TYPES,
        "standard": "GB/T 1804-2000",
        "special_fields": [
            "drawing_no", "material", "heat_treatment",
            "surface_treatment", "weight", "unit_weight",
        ],
        "preset_processes": [
            {"step_no": "S01", "name": "下料", "desc": "切割/锯切下料",
             "output": "坯料"},
            {"step_no": "S02", "name": "机加工", "desc": "车铣钻加工",
             "output": "加工件"},
            {"step_no": "S03", "name": "热处理", "desc": "淬火/回火等",
             "output": "热处理件"},
            {"step_no": "S04", "name": "表面处理", "desc": "镀锌/喷塑/阳极氧化",
             "output": "表面处理件"},
            {"step_no": "S05", "name": "装配", "desc": "零部件组装",
             "output": "成品"},
        ],
    },
    # V6 包装：填实模板（仅交互引导，不写入 JSON 结构）
    "包装": {
        "material_types": PACKAGING_TYPES,
        "standard": "GB/T 6543-2008",
        "special_fields": [
            "material", "basis_weight", "size",
            "print_process", "eco_label",
        ],
        "preset_processes": [
            {"step_no": "S01", "name": "设计制版", "desc": "版面设计与制版",
             "output": "印版"},
            {"step_no": "S02", "name": "印刷", "desc": "胶印/柔印/数码印刷",
             "output": "印刷品"},
            {"step_no": "S03", "name": "模切成型", "desc": "模切与成型",
             "output": "成型包装件"},
            {"step_no": "S04", "name": "检验", "desc": "外观与性能检验",
             "output": "成品"},
        ],
    },
}
