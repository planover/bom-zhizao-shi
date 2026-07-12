# BOM 表结构与输入规范（bom-zhizao-shi · V6 / V5 / V4 / V3 / V2.1）

本文件供 `scripts/generate_bom.py` 与汇总序列化阶段参考，并作为 `scripts/import_bom.py` 逆向解析的契约基线。

> V5 变更摘要：行业专属派生视图扩充——纺织「三、面料辅料清单」（8 列 A–H）、家具「三、家具物料清单」（8 列 A–H）；电子「三、元件清单」扩列至 **14 列（A–N）**（新增 manufacturer/tolerance/rated_power/rated_voltage/alternate/reflow_temp 6 个工程/合规字段，共 10 专属字段）；化工「三、配方表」扩列至 **13 列（A–M）**（新增 purity/physical_state/flash_point/storage_condition/hazard_class 5 个 SDS 字段，共 8 专属字段）；跨行业「成本明细」视图（有行业视图时「四、成本明细」，否则「三、成本明细」，8 列含单价/币种/总价派生 + 成本合计行），物料级新增 `unit_price`/`currency`（currency 默认 人民币(CNY)）；新增行业模板预设 `INDUSTRY_TEMPLATES`（仅交互引导，不写新 JSON 结构）。**V5 不新增任何阻断/软校验**（沿用 V4 同构排除提示）。机械/包装本期仅出评估草案（见 `mechanical-packaging-draft-v5.md`），维持通用兜底。
>
> V6 变更摘要（机械 / 包装行业正式落地，最小变更、100% 向后兼容，不新增任何阻断/软校验）：行业专属派生视图扩充——机械「三、机械物料清单」（8 列 A–H，无「物料类型」展示列，仅用于过滤/排序）、包装「三、包装物料清单」（8 列 A–H，保留「物料类型」展示列）；机械 6 专属字段 `drawing_no`/`material`/`heat_treatment`/`surface_treatment`/`weight`/`unit_weight`（weight 与 unit_weight 双字段保留不合并，Q1 拍板）、包装 5 专属字段 `material`/`basis_weight`/`size`/`print_process`/`eco_label`（`material` 机械/包装同名 key；`surface_treatment` 家具 V5 已存在；`eco_label` 自由文本无受限枚举，Q5 拍板）；跨行业「成本明细」双编号集合 `INDUSTRY_VIEW_SET` 由 V5 的 5 行业扩为 **7 行业**（新增机械/包装），机械/包装带 `unit_price` 时成本视图编号为「四、成本明细」；机械 `standard` 默认 `GB/T 1804-2000`、包装 `standard` 默认 `GB/T 6543-2008`；机械/包装类型枚举与排除集 `MECHANICAL_TYPES`/`MECHANICAL_EXCLUDE={"其他"}`/`PACKAGING_TYPES`/`PACKAGING_EXCLUDE={"其他"}`；逆向 `_SPECIAL_FIELDS` 由 28 唯一键扩至 **37 唯一键**（净增 9：`drawing_no`/`material`(同名复用)/`heat_treatment`/`surface_treatment`(家具已存在)/`weight`/`unit_weight`/`basis_weight`/`size`/`print_process`/`eco_label` 去重后净增 9），`float_fields` 增 `weight`/`unit_weight`/`basis_weight` 使逆向解析为 float；新增机械 6 / 包装 5 字段回收分支与 `_infer_industry_from_blocks` 的「三、机械物料清单」→机械、「三、包装物料清单」→包装 识别。P2-1/P2-2 本期不实现。机械/包装专属字段仅存 JSON、仅专属视图展示，物料区 8 列（A–H）永不变。
>
> V4 变更摘要：新增 BOM 级可选字段 `industry`（8 值枚举，选填，默认按 `category` 推断）；新增物料级专属字段——电子行业 `designator` / `footprint` / `part_number` / `rohs`，化工行业 `cas_number` / `concentration` / `ghs_hazard`；新增行业专属派生视图——电子「三、元件清单」（8 列，含 RoHS 红黄字标记），化工「三、配方表」（8 列，含含量(%) 数字格式）；配料表触发条件从 `category=="食品"` 改为 `industry=="食品"`（含推断，行为不变）；新增软校验 V8（industry 枚举）/ W2（RoHS 未标）/ W3（CAS/GHS 未填）/ 含量(%) 列和校验（±5%）；共享常量迁入 `scripts/bom_constants.py`（EDIBLE / ALLERGEN_SET / ALLERGEN_HINTS / V4 新常量）。
>
> V3（V2.1）变更摘要：Excel 物料区由 7 列扩展为 **8 列（A–H）**，首列新增「序号」（按输入顺序全局连续、跨工序分组不重置）；新增 BOM 级可选字段 `approver` / `effective_date` / `standard`（Excel 行 5 单格合并拼接）；物料区末新增「合计用量」行；食品类「三、配料表」新增「用量占比%」（最大余数法保证列和恰为 100.0%）与「过敏原」两列；新增 `materials[].allergen`（仅食品配料表展示，逆向按物料名回收）；逆向解析改为按列头文本定位并兼容旧版 7/5 列。

---

## 输入 JSON Schema（正向 `--data` 与逆向 `--out` 完全一致）

`generate_bom.py --data <file.json>` 读取、且 `import_bom.py --out <file.json>` 输出的 JSON 结构：

```json
{
  "product_name": "芒果果味糖浆",
  "category": "食品",
  "industry": "食品",
  "output_rate": 130,
  "version": "V1.0",
  "date": "2026-07-07",
  "approver": "张三",
  "effective_date": "2026-07-10",
  "standard": "GB 7718-2025",
  "materials": [
    {
      "name": "芒果原浆",
      "unit": "kg",
      "usage": 46.3,
      "yield_rate": 55,
      "erp_code": "RM-001",
      "material_type": "原料",
      "process": "S01",
      "allergen": "大豆,乳",
      "designator": "",
      "footprint": "",
      "part_number": "",
      "rohs": "",
      "cas_number": "",
      "concentration": "",
      "ghs_hazard": ""
    }
  ],
  "processes": [
    {
      "step_no": "S01",
      "name": "调配",
      "desc": "混合搅拌",
      "work_hours": 30,
      "note": "常温",
      "output": "芒果果味糖浆基料"
    }
  ]
}
```

> 约定：正向 `--data` 输入与逆向 `--out` 输出 **Schema 完全一致**（闭环可回写）。
> **V5 新增**：物料级专属字段按行业分别为——电子 `manufacturer`/`tolerance`/`rated_power`/`rated_voltage`/`alternate`/`reflow_temp`（在 V4 4 个基础上扩至 10）；化工 `purity`/`physical_state`/`flash_point`/`storage_condition`/`hazard_class`（在 V4 3 个基础上扩至 8）；纺织 `composition`/`yarn_count`/`fabric_weight`/`width`/`color_no`；家具 `material_grade`/`spec_size`/`surface_treatment`/`color_no`；跨行业 `unit_price`/`currency`。`total_price` **不入库**（渲染派生）；`currency` 默认 `人民币(CNY)`。上述所有专属字段均默认空串，逆向导入时补全为空串（未回收到的）。
> **V4 新增**：BOM 级可选字段 `industry`；物料级 7 个专属字段（`designator`/`footprint`/`part_number`/`rohs`/`cas_number`/`concentration`/`ghs_hazard`）。八者均默认空串；逆向导入时全部补全为空串（未回收到的）。
> **V3 新增**：`approver`、`effective_date`、`standard`（BOM 级）、`materials[].allergen`（物料级）。四者均默认空串；「序号」「用量占比%」「合计用量」为派生展示，**不进 JSON**。

### V5 行业物料字段示例（节选）

```json
{
  "industry": "电子",
  "materials": [
    {
      "name": "主控芯片", "unit": "个", "usage": 1, "yield_rate": 100,
      "material_type": "IC",
      "designator": "U1", "part_number": "STM32F103C8T6", "footprint": "LQFP48",
      "rohs": "是",
      "manufacturer": "ST", "tolerance": "", "rated_power": "0.36W",
      "rated_voltage": "3.3V", "alternate": "GD32F103C8T6", "reflow_temp": "260℃",
      "unit_price": 12.50, "currency": "人民币(CNY)"
    }
  ]
}
```
```json
{
  "industry": "纺织",
  "materials": [
    {"name": "全棉针织布", "unit": "米", "usage": 2.5, "yield_rate": 100,
     "material_type": "面料",
     "composition": "棉100%", "yarn_count": "40S", "fabric_weight": "180",
     "width": "150cm", "color_no": "藏青 01"}
  ]
}
```
```json
{
  "industry": "通用",
  "materials": [
    {"name": "支架半成品", "unit": "个", "usage": 10, "yield_rate": 100,
     "material_type": "其他",
     "unit_price": 3.50, "currency": "人民币(CNY)"}
  ]
}
```
```json
{
  "industry": "机械",
  "materials": [
    {"name": "传动轴", "unit": "件", "usage": 2, "yield_rate": 100,
     "material_type": "零部件",
     "drawing_no": "DW-2026-011", "material": "45#钢",
     "heat_treatment": "淬火+回火", "surface_treatment": "镀锌",
     "weight": "1.20", "unit_weight": "0.60"},
    {"name": "密封圈", "unit": "件", "usage": 4, "yield_rate": 100,
     "material_type": "标准件",
     "drawing_no": "", "material": "丁腈橡胶",
     "heat_treatment": "", "surface_treatment": "",
     "weight": "0.01", "unit_weight": "0.01"}
  ]
}
```
```json
{
  "industry": "包装",
  "materials": [
    {"name": "外箱", "unit": "个", "usage": 1, "yield_rate": 100,
     "material_type": "纸箱",
     "material": "瓦楞纸", "basis_weight": "200",
     "size": "400×300×200mm", "print_process": "胶印", "eco_label": "FSC"},
    {"name": "缓冲泡棉", "unit": "张", "usage": 2, "yield_rate": 100,
     "material_type": "缓冲",
     "material": "EPE", "basis_weight": "30",
     "size": "380×280mm", "print_process": "", "eco_label": "可回收"}
  ]
}
```

### 字段约束总表

| JSON 路径 | 类型 | 必填 | 约束 / 默认值 | 说明 |
|-----------|------|------|--------------|------|
| `product_name` | string | **必填** | 非空（R1）；默认 `""` → 空则 VALIDATION_FAILED | 沿用 V2 |
| `category` | string(enum) | **必填** | ∈ {食品,工业品,日化化妆品,医药,其他}；默认 `其他` | 沿用 V2（R1/R4） |
| `industry` | string(enum) | **选填（V4 新增）** | ∈ {食品,电子,化工,机械,纺织,家具,包装,通用}；默认按 `category` 推断；非法值 V8 WARNING 回退推断 | 行业标识，决定专属派生视图（食品→配料表，电子→元件清单，化工→配方表，纺织→面料辅料清单，家具→家具物料清单）；有任一行业视图且物料含 `unit_price` → 成本视图编号为「四、成本明细」，否则「三、成本明细」；推断映射：食品→食品，日化/医药→化工，工业品/其他→通用 |
| `output_rate` | number | **必填** | `> 0`；允许 `> 100`；默认 `""` 待补 | 沿用 V2（R2） |
| `version` | string | 选填 | 默认 `V1.0` | 未变 |
| `date` | string | 选填 | 默认当天 `YYYY-MM-DD` | 未变 |
| `approver` | string | **选填（V3 新增）** | 默认 `""`；BOM 级审批人，表头行 5 显示 | 空则行 5 不显示该段 |
| `effective_date` | string | **选填（V3 新增）** | 默认 `""`；BOM 级生效日期，表头行 5 显示 | 空则行 5 不显示该段 |
| `standard` | string | **选填（V3 新增）** | 默认 `""`；BOM 级执行标准，表头行 5 显示；食品建议填 `GB 7718-2025`；电子建议填 `GB/T 39560`；化工建议填 `GB/T 16483-2008` | 空则行 5 不显示该段 |
| `materials[]` | array | **必填** | 至少 1 条 | 未变 |
| `materials[].name` | string | 必填 | 非空 | 未变 |
| `materials[].unit` | string | 必填 | 非空 | 未变 |
| `materials[].usage` | number | 必填 | `> 0` | 未变 |
| `materials[].yield_rate` | number | 必填 | `0 < 值 ≤ 100` | 未变（R2） |
| `materials[].erp_code` | string | 选填 | 默认 `""` | 未变 |
| `materials[].material_type` | string(enum) | 选填 | ∈ {原料,添加剂,香精香料,包材,其他}（食品）；{电阻,电容,IC,连接器,二极管,三极管,晶振,其他}（电子）；{主料,溶剂,催化剂,添加剂,包材,其他}（化工）；{面料,辅料,纱线,印染,五金,其他}（纺织）；{主材,板材,辅材,五金,面料,其他}（家具）；{零部件,标准件,型材,铸件,焊接件,其他}（机械）；{纸箱,缓冲,标签,胶带,薄膜,其他}（包装）；默认 `其他` | 沿用 V2（R4 过滤用）；V5/V6 专属视图过滤依据（电子/化工/纺织/家具/机械/包装均排除"其他"类；化工额外排除"包材"类；机械与包装 `MECHANICAL_EXCLUDE`/`PACKAGING_EXCLUDE` 均为 `{"其他"}`） |
| `materials[].process` | string | 选填 | 引用有效 `step_no`；默认 `""` | 沿用 V2 |
| `materials[].allergen` | string | **选填（V3 新增）** | 逗号分隔标签，∈ 八大类+其他；默认 `""` | 仅食品配料表展示；软校验 W1/H1 |
| `materials[].designator` | string | **选填（V4 新增）** | 默认 `""`；电子元件位号（如 R1/C3/U5） | 仅电子元件清单展示 |
| `materials[].footprint` | string | **选填（V4 新增）** | 默认 `""`；电子元件封装（如 0805/SOIC-8） | 仅电子元件清单展示 |
| `materials[].part_number` | string | **选填（V4 新增）** | 默认 `""`；电子元件型号（如 STM32F103C8T6） | 仅电子元件清单展示 |
| `materials[].rohs` | string(enum) | **选填（V4 新增）** | ∈ {是,否,未知}；默认 `""`（等价未知） | 仅电子元件清单展示；RoHS 着色：否→红字，未知/空→黄字，是→默认；软校验 W2 |
| `materials[].cas_number` | string | **选填（V4 新增）** | 默认 `""`；化学品 CAS 登记号（如 7732-18-5） | 仅化工配方表展示；软校验 W3 |
| `materials[].concentration` | number | **选填（V4 新增）** | 默认 `""`；配方含量(%)（如 65.0）；含量和校验 ±5% | 仅化工配方表展示；空则留空（不显示 0.0%） |
| `materials[].ghs_hazard` | string | **选填（V4 新增）** | 默认 `""`；GHS 危险标识（如 GHS07） | 仅化工配方表展示；软校验 W3 |
| `materials[].manufacturer` | string | **选填（V5 新增）** | 默认 `""`；元件制造商（如 ST、TI） | 仅电子元件清单（14 列）展示 |
| `materials[].tolerance` | string | **选填（V5 新增）** | 默认 `""`；标称容差（如 ±1%） | 仅电子元件清单展示 |
| `materials[].rated_power` | string | **选填（V5 新增）** | 默认 `""`；额定功率（如 0.36W） | 仅电子元件清单展示 |
| `materials[].rated_voltage` | string | **选填（V5 新增）** | 默认 `""`；额定电压（如 3.3V） | 仅电子元件清单展示 |
| `materials[].alternate` | string | **选填（V5 新增）** | 默认 `""`；替代料型号（pin-to-pin） | 仅电子元件清单展示 |
| `materials[].reflow_temp` | string | **选填（V5 新增）** | 默认 `""`；封装/回流焊峰值温度（如 260℃） | 仅电子元件清单展示 |
| `materials[].purity` | string | **选填（V5 新增）** | 默认 `""`；纯度（如 ≥95%） | 仅化工配方表（13 列）展示 |
| `materials[].physical_state` | string | **选填（V5 新增）** | 默认 `""`；物态（固/液/气） | 仅化工配方表展示 |
| `materials[].flash_point` | string | **选填（V5 新增）** | 默认 `""`；闪点（如 13℃） | 仅化工配方表展示 |
| `materials[].storage_condition` | string | **选填（V5 新增）** | 默认 `""`；存储条件（如 阴凉干燥） | 仅化工配方表展示 |
| `materials[].hazard_class` | string | **选填（V5 新增）** | 默认 `""`；危险等级（如 易燃液体 类别2） | 仅化工配方表展示 |
| `materials[].composition` | string | **选填（V5 新增）** | 默认 `""`；成分比例（如 棉60%/涤40%） | 仅纺织面料辅料清单展示 |
| `materials[].yarn_count` | string | **选填（V5 新增）** | 默认 `""`；纱支（如 40S） | 仅纺织面料辅料清单展示 |
| `materials[].fabric_weight` | string/number | **选填（V5 新增）** | 默认 `""`；克重(g/m²)（如 180） | 仅纺织面料辅料清单展示 |
| `materials[].width` | string | **选填（V5 新增）** | 默认 `""`；幅宽（如 150cm） | 仅纺织面料辅料清单展示 |
| `materials[].color_no` | string | **选填（V5 新增）** | 默认 `""`；色号/花色（纺织/家具共用） | 纺织面料辅料清单 / 家具物料清单展示 |
| `materials[].material_grade` | string | **选填（V5 新增）** | 默认 `""`；材质等级（如实木/板木） | 仅家具物料清单展示 |
| `materials[].spec_size` | string | **选填（V5 新增）** | 默认 `""`；尺寸规格（如 1200×600×750mm） | 仅家具物料清单展示 |
| `materials[].surface_treatment` | string | **选填（V5 新增）** | 默认 `""`；表面处理（如 烤漆） | 仅家具物料清单展示 |
| `materials[].unit_price` | number | **选填（V5 新增）** | 默认 `""`（不填则不进成本视图）；数值 | 仅成本明细展示；`total_price` 不入库、按 `用量×单价` 实时派生 |
| `materials[].currency` | string | **选填（V5 新增）** | 默认 `人民币(CNY)`；缺省按人民币(CNY)处理 | 仅成本明细展示 |
| `materials[].drawing_no` | string | **选填（V6 新增）** | 默认 `""`；零部件图纸编号（如 DW-2026-011） | 仅机械物料清单展示 |
| `materials[].material` | string | **选填（V6 新增）** | 默认 `""`；材质（机械如 45#钢/铝合金6061，包装如 瓦楞纸/PET）；机械与包装同名 key（JSON 仅一份） | 仅机械/包装物料清单展示 |
| `materials[].heat_treatment` | string | **选填（V6 新增）** | 默认 `""`；热处理（如 淬火+回火/退火/无） | 仅机械物料清单展示 |
| `materials[].surface_treatment` | string | **选填（V5 家具/V6 机械 复用）** | 默认 `""`；表面处理（如 镀锌/喷塑/阳极氧化）；家具 V5 已存在，机械 V6 同 key 复用 | 家具物料清单 / 机械物料清单展示 |
| `materials[].weight` | number | **选填（V6 新增）** | 默认 `""`；单件总重量（kg）；数值 | 仅机械物料清单展示；与 `unit_weight` 双字段保留不合并（Q1 拍板） |
| `materials[].unit_weight` | number | **选填（V6 新增）** | 默认 `""`；单位重量（kg/件）；数值 | 仅机械物料清单展示 |
| `materials[].basis_weight` | number | **选填（V6 新增）** | 默认 `""`；纸张/薄膜克重（g/m²）；数值 | 仅包装物料清单展示 |
| `materials[].size` | string | **选填（V6 新增）** | 默认 `""`；尺寸（如 400×300×200mm） | 仅包装物料清单展示 |
| `materials[].print_process` | string | **选填（V6 新增）** | 默认 `""`；印刷工艺（如 胶印/柔印/数码印刷） | 仅包装物料清单展示 |
| `materials[].eco_label` | string | **选填（V6 新增）** | 默认 `""`；环保标识（如 FSC/可回收/可降解）；自由文本，无受限枚举（Q5 拍板） | 仅包装物料清单展示 |
| `processes[]` | array | 选填 | 可 0 条 | 未变 |
| `processes[].step_no` | string | 必填 | 唯一不重复 | 未变 |
| `processes[].name` | string | 必填 | 非空 | 未变 |
| `processes[].desc` | string | 选填 | — | 未变 |
| `processes[].work_hours` | number/string | 选填 | 数值须 `≥ 0` | 未变 |
| `processes[].note` | string | 选填 | — | 未变 |
| `processes[].output` | string | **必填** | 非空；该工序产物（R3 流转链源头） | 沿用 V2 |

> **不进 JSON 的派生展示项（正向生成、逆向不输出）**：`序号`（物料区首列）、`用量占比%`（配料表列）、`合计用量`行。

### industry 推断逻辑（V4 核心）

`infer_industry(data)` 返回 `(industry, warnings)`：

| 场景 | industry 来源 | warnings |
|------|--------------|----------|
| `industry` 已显式设置且合法（∈ 8 值枚举） | 使用该值 | `[]` |
| `industry` 已显式设置但非法 | 回退为 `CATEGORY_TO_INDUSTRY.get(category, "通用")` | V8 WARNING（非阻断） |
| `industry` 未设置 | 按 `category` 推断 | `[]` |

**category → industry 推断映射**（`CATEGORY_TO_INDUSTRY`，定义于 `scripts/bom_constants.py`）：

| category | → industry | 专属视图 |
|----------|-----------|---------|
| 食品 | 食品 | 配料表 |
| 日化化妆品 | 化工 | 配方表 |
| 医药 | 化工 | 配方表 |
| 工业品 | 通用 | 无 |
| 其他 | 通用 | 无 |

> 向后兼容核心：现有食品 JSON 无 `industry` → 推断「食品」→ 配料表照常（与 V3 一致）；现有非食品 JSON → 推断「通用」→ 不生成专属视图（与 V3 一致）。

### 向后兼容默认（R5，旧数据 / 旧文件）

- 旧 JSON 缺 `industry` → 按 `category` 推断（行为零变化）。
- 旧 JSON 缺物料级专属字段（`designator`/`footprint`/`part_number`/`rohs`/`cas_number`/`concentration`/`ghs_hazard` → V4；V5 电子 10/化工 8/纺织 5/家具 4/成本 2；V6 机械 6：`drawing_no`/`material`/`heat_treatment`/`surface_treatment`/`weight`/`unit_weight`，包装 5：`material`/`basis_weight`/`size`/`print_process`/`eco_label`）→ 默认空串 `""`，专属视图对应列留空。
- 旧 JSON 缺 `unit_price`/`currency` → `""`/默认人民币(CNY)，不生成成本视图（行为零变化）。
- 旧 JSON 缺 `approver` / `effective_date` / `standard` → `""`（正向行 5 留空）。
- 旧 JSON 缺 `materials[].allergen` → `""`（配料表过敏原列留空）。
- 其余沿用 V2：`category="其他"`、`output_rate=""`、`material_type="其他"`、`process=""`、`output=""`。
- 旧版 7 列 / 5 列 Excel 逆向：缺「序号/占比%/过敏原/审批人/生效日期/执行标准」列 → 取默认（`""` / 不读），完全兼容。
- 旧 Excel 无「三、元件清单」/「三、配方表」区块 → 无区块标记 → 按 category 推断 industry → 完全兼容；旧 8 列电子/化工 Excel 逆向仍可识别（按列头文本定位，max_col 向后兼容）。

---

## 行业模板预设（INDUSTRY_TEMPLATES · V5，仅交互引导，不写新 JSON 结构）

`scripts/bom_constants.py` 新增 `INDUSTRY_TEMPLATES` 字典，**仅被交互层（SKILL.md）读取**，用于降低采集负担。它**不向输入 JSON Schema 写入任何新字段**（向后兼容旧 JSON 与旧交互）。键为 industry，值含四项：

| 键 | 用途 | 说明 |
|----|------|------|
| `material_types` | 阶段一物料类型下拉建议值 | 非强制枚举；如电子 `["电阻","电容","IC","连接器","二极管","三极管","晶振","其他"]` |
| `standard` | 执行标准自动预填 | 可改；如食品 `GB 7718-2025`、电子 `GB/T 39560`、化工 `GB/T 16483-2008`、纺织 `FZ/T 80004`、家具 `QB/T 1951.1`、机械 `GB/T 1804-2000`、包装 `GB/T 6543-2008` |
| `special_fields` | 阶段一物料模板动态追加的专属字段行 | 电子 10 / 化工 8 / 纺织 5 / 家具 4 / 机械 6 / 包装 5 / 食品 1(过敏原) 个；通用为空 |
| `preset_processes` | 可选「一键载入工序模板」 | 阶段二工序预填（仅引导，不写新结构）；机械 5 步 / 包装 4 步 / 通用/食品/电子/化工/纺织/家具为空 |

> 交互约束（最小变更）：阶段一物料模板、物料类型下拉、执行标准预填、一键载入工序，**全部由交互层按 `INDUSTRY_TEMPLATES[industry]` 动态生成**；生成脚本仍按既有 `industry` + 专属字段渲染，JSON Schema 零新增。V6 起机械/包装已由空模板升级为正式预置（`MECHANICAL_TYPES`/`PACKAGING_TYPES` 物料类型下拉、`GB/T 1804-2000`/`GB/T 6543-2008` 执行标准、`drawing_no`/`material`/`heat_treatment`/`surface_treatment`/`weight`/`unit_weight`(机械 6) 与 `material`/`basis_weight`/`size`/`print_process`/`eco_label`(包装 5) 专属字段行、机械 5 步/包装 4 步工序模板），见 `mechanical-packaging-draft-v5.md`（V6 已实现）。

---

## Excel 输出结构（8 列 A–H，固定行号）

> 因 `product_name` 必填非空 + 行 5 可选不挤占行 6，**新格式 Excel 行号固定**（见下表）。旧版 Excel 无行 5 内容、无序号列、无审批人/过敏原，逆向解析**不依赖固定行号**（按标记+列头文本定位），完全兼容。

### 行业专属派生视图总览（V5）

| industry | 专属视图 | 区块标题 | 列数 | 触发条件 |
|----------|---------|---------|------|---------|
| 食品 | 配料表 | `三、配料表` | 7 列（A–G） | `industry == "食品"`（含推断） |
| 电子 | 元件清单 | `三、元件清单` | **14 列（A–N，V5 扩列）** | `industry == "电子"` |
| 化工 | 配方表 | `三、配方表` | **13 列（A–M，V5 扩列）** | `industry == "化工"` |
| 纺织 | 面料辅料清单 | `三、面料辅料清单` | 8 列（A–H，V5 新增） | `industry == "纺织"` |
| 家具 | 家具物料清单 | `三、家具物料清单` | 8 列（A–H，V5 新增） | `industry == "家具"` |
| 机械 | 机械物料清单 | `三、机械物料清单` | 8 列（A–H，V6 新增，无「物料类型」展示列） | `industry == "机械"` |
| 包装 | 包装物料清单 | `三、包装物料清单` | 8 列（A–H，V6 新增，保留「物料类型」展示列） | `industry == "包装"` |
| 任意（含 通用） | 成本明细 | `四、成本明细`（有行业视图时，含食品/电子/化工/纺织/家具/机械/包装）/ `三、成本明细`（无行业视图时，仅通用） | 8 列（A–H） | 任一物料 `unit_price` 非空 |

> 配料表触发条件从 V3 的 `category=="食品"` 改为 V4 的 `industry=="食品"`（含推断），行为不变。
> V6 成本视图为跨行业视图（`INDUSTRY_VIEW_SET = {食品,电子,化工,纺织,家具,机械,包装}` 共 7 行业）：有行业专属视图（食品/电子/化工/纺织/家具/机械/包装）→ 编号为「四、成本明细」（位于行业视图之后）；否则（仅通用）→「三、成本明细」（位于工序区之后）。V5 的 `INDUSTRY_VIEW_SET` 为 5 行业，V6 扩至 7 行业（新增机械/包装）；旧 V5 机械/包装带 `unit_price` 的 Excel 成本块编号由「三」变为「四」为预期行为变更（数据一致，逆向按关键字 `成本明细` 兼容前缀识别，无需迁移）。

### ASCII 框图（食品类 — 配料表，沿用 V3）

```
行1  ┌────────────────────────────────────────────────────────────── A1:H1 合并 ┐
     │                          BOM表（标题 16pt 深蓝加粗居中）                  │
行2  ├───────────────────────────┬────────────────────────────────────────────┤
     │ 版本号：V1.0 (A2:C2 合并)  │ 生成日期：2026-07-07 (D2:H2 合并)          │
行3  ├───────────────────────────┴────────────────────────────────────────────┤
     │ 产品名称：芒果果味糖浆 (A3:H3 合并, label_font 加粗)                     │
行4  ├───────────────────────────┬────────────────────────────────────────────┤
     │ 产品类别：食品 (A4:C4 合并) │ 全产品出品率：130.0% (D4:H4 合并)         │
行5  ├────────────────────────────────────────────────────────────────────────┤  ← 审批人/生效日期/执行标准(可选, 单格A5:H5)
     │ 审批人：张三    生效日期：2026-07-10    执行标准：GB 7718-2025          │  ← 仅拼接非空段, 段间4空格; 皆空则整行留空
行6  │ 一、物料信息 (A6:H6 合并, label_font)                                   │
行7  ├────┬──────────┬──────┬──────┬──────────┬──────────┬────────┬────────┤
     │ 序号│ 物料名称 │ 单位 │ 用量 │ 出品率(%) │ ERP物料代码│ 物料类型│ 所属工序│  ← 表头(蓝底,8列)
行8  ├────┼──────────┼──────┼──────┼──────────┼──────────┼────────┼────────┤
     │【工序 S01 调配】(A8:H8 合并, 分组子标题浅蓝底+左侧色条+加粗)            │
行9  │ 1  │ 芒果原浆  │ kg  │ 46.3 │ 55.0%    │ RM-001   │ 原料   │ S01    │  allergen=大豆,乳
行10 │ 2  │ 白砂糖    │ kg  │ 30.0 │ 100.0%   │ RM-002   │ 原料   │ S01    │
行11 ├────┼──────────┼──────┼──────┼──────────┼──────────┼────────┼────────┤
     │【工序 S02 灌装】(分组子标题)                                          │
行12 │ 3  │ 芒果果味糖浆基料│ kg │ 70.0 │ 98.0%  │ RM-100   │ 原料   │ S02    │
行13 │ 4  │ 柠檬酸    │ kg  │ 0.5  │ 100.0%   │ RM-003   │ 添加剂 │ S02    │
行14 ├────┼──────────┼──────┼──────┼──────────┼──────────┼────────┼────────┤
     │【未归属工序】(仅当存在 process 为空/无效的物料时出现)                  │
行15 │ 5  │ PE 瓶     │ 个  │ 100  │ 100.0%   │ PK-001   │ 包材   │        │
行16 ├────┼──────────┼──────┼──────┼──────────┼──────────┼────────┼────────┤
     │ 合计│         │     │246.8 │          │          │        │        │  ← 合计行(A=合计, D=全部物料usage求和)
行17 ├─────────────────────────────────────────────────────────────────────┤  ← 空行间隔
行18 │ 二、工艺工序 (A18:H18 合并, label_font)                               │
行19 ├──────────┬──────┬──────┬──────┬────────┬──────────┬────────┬────────┤
     │ 工序编号 │ 工序名称│工序说明│ 工时 │ 备注  │ 产物     │(G空)   │(H空)  │  ← 表头(蓝底, 仅A–F)
行20 │ S01     │ 调配  │混合搅拌│ 30  │ 常温  │ 芒果果味糖浆基料│       │        │
行21 │ S02     │ 灌装  │无菌灌装│ 20  │       │ 芒果果味糖浆  │       │        │
行22 ├─────────────────────────────────────────────────────────────────────┤  ← 空行间隔
行23 │ 三、配料表 (A23:H23 合并, label_font, 仅 industry==食品 出现)          │
行24 ├──────────┬────────┬────────┬──────┬──────────┬──────────┬────────┤
     │ 物料名称 │ 物料类型│计量单位│ 用量 │ 出品率(%) │ 用量占比% │ 过敏原  │  ← 表头(蓝底, 7列 A–G)
行25 │ 芒果果味糖浆基料│ 原料 │ kg  │ 70.0 │ 98.0%    │ 47.7%    │        │
行26 │ 芒果原浆 │ 原料   │ kg    │ 46.3 │ 55.0%    │ 31.5%    │ 大豆,乳│
行27 │ 白砂糖   │ 原料   │ kg    │ 30.0 │ 100.0%   │ 20.4%    │        │
行28 │ 柠檬酸   │ 添加剂 │ kg    │ 0.5  │ 100.0%   │ 0.4%     │        │
     └──────────┴────────┴────────┴──────┴──────────┴──────────┴────────┘
```

> 注：示例占比% 经最大余数法补差后列和恰为 100.0%（柠檬酸由 0.3% 补为 0.4%）：47.7+31.5+20.4+0.4=100.0。

### ASCII 框图（电子类 — 元件清单，★V4 新增）

```
行1~行17  (表头区 + 物料区 + 合计行 + 空行，同上，略)
行18 │ 二、工艺工序 (A18:H18 合并, label_font)                               │
行19 │ 工序编号│工序名称│工序说明│工时│备注│产物│(G空)│(H空)│                           │
行20 │ S01     │贴片   │SMT贴片 │ 2  │    │贴片完成板│    │     │                         │
行21 │ S02     │检测   │功能测试│ 1  │    │STM32最小系统板│  │     │                     │
行22 │ (空行)                                                                   │
行23 │ 三、元件清单 (A23:H23 合并, label_font, 仅 industry==电子 出现)        │
行24 ├────┬─────────────┬──────────────┬──────────────┬──────────┬──────┬────────┬──────┤
     │序号│位号(Designator)│型号(Part#)│封装(Footprint)│ 物料名称 │数量│物料类型│RoHS  │ ← 表头(蓝底,8列)
行25 │ 1  │ C1           │ CL10A105KP8 │ 0805         │ 贴片电容 │ 4  │ 电容   │ 是   │
行26 │ 2  │ C2           │ CL21B104KCF │ 0805         │ 贴片电容 │ 2  │ 电容   │ 是   │
行27 │ 3  │ D1           │ 1N4148W     │ SOD-123      │ 开关二极管│ 1  │ 二极管 │ 否(红)│ ← RoHS=否→红字
行28 │ 4  │ R1           │ RC0805FR-07 │ 0805         │ 贴片电阻 │ 2  │ 电阻   │ 是   │
行29 │ 5  │ U1           │ STM32F103C8T6│ LQFP-48     │ MCU      │ 1  │ IC     │ 是   │
行30 │ 6  │ Y1           │ 8MHz        │ HC-49S       │ 晶振     │ 1  │ 晶振   │ (黄) │ ← RoHS=空→黄字(待确认)
     └────┴─────────────┴──────────────┴──────────────┴──────────┴──────┴────────┴──────┘
```

> 元件清单排序：按物料类型升序 → 同类型内按位号字母数字升序；空位号排末尾（`\uffff` 哨兵）。
> RoHS 着色：`"否"`→红色 `FF0000`，`"未知"`/空→黄色 `BF8F00`，`"是"`→默认黑色。
> 过滤排除：`material_type ∈ {"其他"}` 的物料（如裸PCB板、散热片）不进元件清单。

### ASCII 框图（化工类 — 配方表，★V4 新增）

```
行1~行17  (表头区 + 物料区 + 合计行 + 空行，同上，略)
行18 │ 二、工艺工序 (A18:H18 合并, label_font)                               │
行19 │ 工序编号│工序名称│工序说明│工时│备注│产物│(G空)│(H空)│                           │
行20 │ S01     │混合   │按比例混合│ 1  │    │混合液│    │     │                             │
行21 │ S02     │灌装   │灌装入瓶  │ 0.5│    │消毒酒精喷雾│  │     │                         │
行22 │ (空行)                                                                   │
行23 │ 三、配方表 (A23:H23 合并, label_font, 仅 industry==化工 出现)          │
行24 ├────┬──────────┬──────────────┬─────────┬──────────┬────────┬────────┬──────┤
     │序号│ 物料名称 │ CAS号        │ 含量(%) │ GHS标识  │物料类型│计量单位│用量  │ ← 表头(蓝底,8列)
行25 │ 1  │ 乙醇     │ 64-17-5      │ 65.0%   │ GHS02    │ 主料   │ kg     │ 6.5  │
行26 │ 2  │ 去离子水 │ 7732-18-5    │ 30.0%   │          │ 溶剂   │ kg     │ 3.0  │
行27 │ 3  │ 甘油     │ 56-81-5      │ 3.0%    │          │ 添加剂 │ kg     │ 0.3  │
行28 │ 4  │ 薄荷香精 │              │ 2.0%    │ GHS07    │ 添加剂 │ kg     │ 0.2  │
     └────┴──────────┴──────────────┴─────────┴──────────┴────────┴────────┴──────┘
```

> 配方表排序：按 `concentration` 降序（主成分在前）；空含量排末尾。
> 含量(%) 数字格式：`0.0"%"`；空值留空（不显示 `0.0%`）。
> 含量和校验：所有配方原料均填了 concentration 时校验列和 ≈ 100%（±5%），偏差超阈值打印 WARNING。
> 过滤排除：`material_type ∈ {"包材"}` 的物料（如喷雾瓶、标签）不进配方表。

### 区块规则

- **表头区（行 1–5 固定）**：
  - 行 1 标题（A1:H1 合并）；行 2 版本号(A2:C2)+生成日期(D2:H2)；行 3 产品名称(A3:H3 合并)；行 4 产品类别(A4:C4)+全产品出品率(D4:H4，`0.0"%"`)。
  - **行 5（审批人/生效日期/执行标准，可选）**：单格合并 **A5:H5**，内容由非空字段拼接：`审批人：{approver}`、`生效日期：{effective_date}`、`执行标准：{standard}`，段间用 4 个空格分隔；三者皆空则整行留空。执行标准行业建议：食品→`GB 7718-2025`，电子→`GB/T 39560`，化工→`GB/T 16483-2008`。
- **「一、物料信息」（行 6 起）**：
  - 表头 **8 列**（A–H）：`序号|物料名称|单位|用量|出品率(%)|ERP物料代码|物料类型|所属工序`。
  - **「序号」列（A 列）**：全局连续 `1..N`，按 `materials[]` **输入顺序**赋值，**跨工序分组不重置**。纯展示，不进 JSON。
  - **分组子标题**：有工序且存在 `process` 归属时按 `step_no` 升序插入 `【工序 Sxx 名称】`（A–H 合并，浅蓝底+左侧加粗色条+加粗），未归属物料归 `【未归属工序】`；无工序/全空则平铺。
  - 物料行其余列取值同 V2，出品率数字格式 `0.0"%"`。
  - **物料区末「合计用量」行**：A 列写「合计」，D 列（用量）写**所有物料 `usage` 求和**（含包材/其他）。纯展示，**逆向跳过**。
- **「二、工艺工序」（物料区后空行起）**：表头 6 列（A–F）`工序编号|工序名称|工序说明|工时|备注|产物`（G/H 留空）；`产物` 取 `output`。区块标题合并 A–H。工序区停止边界为任一「三、」区块标记行。
- **「三、配料表」（仅 `industry=="食品"`，工序区后空行起）**：派生区块。表头 **7 列**（A–G）`物料名称|物料类型|计量单位|用量|出品率(%)|用量占比%|过敏原`；H 留空。内容为 `derive_ingredients()` 结果，**按 `usage` 降序**。
  - **「用量占比%」列（F）**：`round(usage / 可食用物料用量合计 × 100, 1)`，经最大余数法补差使列和恰为 `100.0`；数字格式 `0.0"%"`；纯派生，逆向不读。
  - **「过敏原」列（G）**：取该物料 `allergen`（空则留空）；仅食品类出现。
  - 非食品类**不输出此区块**。逆向仅回收过敏原列（按物料名匹配）。
- **「三、元件清单」（仅 `industry=="电子"`，工序区后空行起，★V4 新增，★V5 扩列至 14 列 A–N）**：派生区块。表头 **14 列**（A–N）`序号|位号(Designator)|型号(Part#)|封装(Footprint)|物料名称|数量|物料类型|RoHS|制造商|容差|额定功率|额定电压|替代料|封装温度`。
  - 内容为 `derive_components()` 结果：排除 `material_type ∈ {"其他"}` 的物料（散热片/外壳/包装等），按 `(material_type, designator)` 升序排序（空位号用 `\uffff` 哨兵排末尾）。
  - **RoHS 列（H）**：取该物料 `rohs`。着色规则：`"否"`→红色字 `FF0000`，`"未知"`/空→黄色字 `BF8F00`，`"是"`→默认黑色字。
  - 非电子类**不输出此区块**。逆向按物料名回收 `designator`/`footprint`/`part_number`/`rohs` + V5 扩列字段 `manufacturer`/`tolerance`/`rated_power`/`rated_voltage`/`alternate`/`reflow_temp`（按 14 列识别）。
- **「三、配方表」（仅 `industry=="化工"`，工序区后空行起，★V4 新增，★V5 扩列至 13 列 A–M）**：派生区块。表头 **13 列**（A–M）`序号|物料名称|CAS号|含量(%)|GHS标识|物料类型|计量单位|用量|纯度|物态|闪点|存储条件|危险等级`。
  - 内容为 `derive_formula()` 结果：排除 `material_type ∈ {"包材"}` 的物料（瓶子/标签等），按 `concentration` 降序排序（空排末尾）。
  - **含量(%) 列（D）**：取该物料 `concentration`（数值则写入 float + `0.0"%"` 格式；空则留空，不显示 `0.0%`）。
  - 非化工类**不输出此区块**。逆向按物料名回收 `cas_number`/`concentration`/`ghs_hazard` + V5 扩列字段 `purity`/`physical_state`/`flash_point`/`storage_condition`/`hazard_class`（按 13 列识别）。
- **「三、面料辅料清单」（仅 `industry=="纺织"`，工序区后空行起，★V5 新增）**：派生区块。表头 **8 列**（A–H）`序号|物料名称|物料类型|成分比例|纱支|克重(g/m²)|幅宽|色号`。
  - 内容为 `derive_textile()` 结果：排除 `material_type ∈ {"其他"}` 的物料，按 `(material_type, name)` 升序排序。
  - 非纺织类**不输出此区块**。逆向按物料名回收 `composition`/`yarn_count`/`fabric_weight`/`width`/`color_no`。
- **「三、家具物料清单」（仅 `industry=="家具"`，工序区后空行起，★V5 新增）**：派生区块。表头 **8 列**（A–H）`序号|物料名称|物料类型|材质等级|尺寸规格|表面处理|用量|色号/花色`。
  - 内容为 `derive_furniture()` 结果：排除 `material_type ∈ {"其他"}` 的物料，按 `(material_type, name)` 升序排序。
  - 非家具类**不输出此区块**。逆向按物料名回收 `material_grade`/`spec_size`/`surface_treatment`/`color_no`。
- **「三、机械物料清单」（仅 `industry=="机械"`，工序区后空行起，★V6 新增）**：派生区块。表头 **8 列**（A–H）`序号|物料名称|图号|材质|热处理|表面处理|重量(kg)|单重(kg/件)`。**不单列「物料类型」展示列**（物料类型仅用于过滤/排序，主理人拍板 Q2）。
  - 内容为 `derive_mechanical()` 结果：排除 `material_type ∈ {"其他"}` 的物料（毛坯/备品等），按 `(material_type, name)` 升序排序。
  - `重量(kg)` 取 `weight`、`单重(kg/件)` 取 `unit_weight`（均为 float，数字格式 `0.00`；两字段保留不合并，主理人拍板 Q1）。
  - 非机械类**不输出此区块**。逆向按物料名回收 `drawing_no`/`material`/`heat_treatment`/`surface_treatment`/`weight`/`unit_weight`（机械视图无「物料类型」展示列，故不回收 `material_type` 列；`surface_treatment` 与家具同 key）。
- **「三、包装物料清单」（仅 `industry=="包装"`，工序区后空行起，★V6 新增）**：派生区块。表头 **8 列**（A–H）`序号|物料名称|物料类型|材质|克重(g/m²)|尺寸|印刷工艺|环保标识`。**保留「物料类型」展示列**（主理人拍板 Q2）。
  - 内容为 `derive_packaging()` 结果：排除 `material_type ∈ {"其他"}` 的物料，按 `(material_type, name)` 升序排序。
  - `克重(g/m²)` 取 `basis_weight`（float，数字格式 `0.0`）；`材质` 取 `material`（与机械同名 key）；`环保标识` 取 `eco_label`（自由文本，无受限枚举，主理人拍板 Q5）。
  - 非包装类**不输出此区块**。逆向按物料名回收 `material`/`basis_weight`/`size`/`print_process`/`eco_label`（回收 `material_type` 列 → 回写 `materials[].material_type`，因包装视图保留该展示列）。
- **「成本明细」（跨行业，任一物料 `unit_price` 非空即生成，★V5 新增，★V6 扩集）**：派生区块。有行业专属视图（食品/电子/化工/纺织/家具/机械/包装）→ 标题 `四、成本明细`（A–H 合并，位于行业视图之后）；否则（仅通用）→ 标题 `三、成本明细`（位于工序区之后）。表头 **8 列**（A–H）`序号|物料名称|物料类型|用量|单位|单价|币种|总价`。
  - 内容为 `derive_cost()` 结果：纳入**全物料**中 `unit_price` 非空者（含行业视图排除的"其他"/"包材"类，成本核算面向全物料）。
  - **总价（H）**：`round(usage × unit_price, 2)` 实时派生（**不入库**，逆向不回收）；数字格式 `0.00`。
  - **币种（G）**：取 `currency`，缺省 `人民币(CNY)`；数字格式同单价/总价列。
  - 末行「成本合计」：A 列写「成本合计」，H 列写 Σ总价（纯展示，**逆向跳过**）。
  - 逆向按物料名回收 `unit_price`/`currency`（关键字 `成本明细` 识别）。
- **其他行业（通用）**：不生成「三、」行业专属视图；仅当物料含 `unit_price` 时生成「三、成本明细」（位于工序区之后）。机械/包装已落地专属视图（见上），不再走通用兜底。
- **样式**：标题 16pt 深蓝加粗居中；区标题 10pt 加粗；表头蓝底加粗居中带边框；数据单元格带边框、文本列左对齐、数值列居中；**数字格式统一**：出品率 / 全产品出品率 / 用量占比% / 含量(%) 均 `0.0"%"`。

### 列宽建议

依据 industry 与扩列扩展（全表共享）：

| industry | 列数 | A 序号 | B 物料名称 | C/D/E/F…（单位/用量/出品率/ERP…） | 扩列附加列 |
|----------|------|--------|-----------|----------------------------------|-----------|
| 通用/食品/纺织/家具 | 8 列（A–H） | 6 | 18 | C10/D10/E13/F16/G13/H12 | — |
| 机械 | 8 列（A–H） | 6 | 18 | C16/D12/E13/F13/G12/H12 | — |
| 包装 | 8 列（A–H） | 6 | 18 | C10/D12/E12/F18/G13/H13 | — |
| 电子 | **14 列（A–N）** | 6 | 18 | C18/D12/E18/F10/G13/H10 | I14/J12/K14/L14/M16/N14（制造商/容差/额定功率/额定电压/替代料/封装温度） |
| 化工 | **13 列（A–M）** | 6 | 18 | C12/D13/E14/F13/G12/H10 | I10/J12/K12/L18/M14（纯度/物态/闪点/存储条件/危险等级） |

> 工序区仅用 A–F（G/H 留空）；配料表复用 A–G；元件清单/配方表/面料辅料清单/家具物料清单/机械物料清单/包装物料清单/成本明细使用 A–H；电子扩至 A–N、化工扩至 A–M。机械/包装/成本明细列宽同 8 列（A–H）。机械列序 `序号/物料名称/图号/材质/热处理/表面处理/重量(kg)/单重(kg/件)`，包装列序 `序号/物料名称/物料类型/材质/克重(g/m²)/尺寸/印刷工艺/环保标识`。

---

## 逆向导入（import_bom.py · V4）

`scripts/import_bom.py` 读取由本规范生成的 BOM 表 Excel，反向解析为与上方「输入 JSON Schema」**完全一致**的结构化 JSON（含 `category` / `industry` / `output_rate` / `material_type` / `process` / `output` / `approver` / `effective_date` / `standard` / `materials[].allergen` / 7 个专属字段），便于「重新编辑已有 BOM」或闭环回写。

### CLI

```
python3 import_bom.py --in <BOM.xlsx> [--out <data.json>]
```

- `--in`：必填，待解析的 BOM 表 Excel 路径。
- `--out`：可选，指定后将 JSON 写入该路径（utf-8、ensure_ascii=False、indent=2），stdout 打印 `OK:<json路径>`；不指定则 stdout 直接打印 `OK:<json字符串>`。

### 解析策略（关键：按列头文本定位，不硬编码列号）

1. **区块定位**：按首列标记 `一、物料信息` / `二、工艺工序` / `三、配料表` / `三、元件清单` / `三、配方表` / `三、面料辅料清单` / `三、家具物料清单` / `成本明细`（有行业视图时为 `四、成本明细`，逆向按关键字 `成本明细` 统一识别）找区块起始行（兼容旧版缺 `三、`）。
2. **列定位**：读取各区块表头行，用**表头文本 → 列号**映射（而非固定 A–E），缺列则视为「旧版，取默认」。电子区块按 14 列识别（`max_col=14`），化工区块按 13 列识别（`max_col=13`），其余按 8 列（默认 `max_col=8`）；旧 8 列电子/化工 Excel 仍兼容。
3. **industry 推断**（V4 新增，V5 扩充）：
   - 有「三、元件清单」→ `industry = "电子"`
   - 有「三、配方表」→ `industry = "化工"`
   - 有「三、配料表」→ `industry = "食品"`
   - 有「三、面料辅料清单」→ `industry = "纺织"`
   - 有「三、家具物料清单」→ `industry = "家具"`
   - 有「三、机械物料清单」→ `industry = "机械"`
   - 有「三、包装物料清单」→ `industry = "包装"`
   - 无「三、」区块 → 按 `category` 推断（`CATEGORY_TO_INDUSTRY.get(category, "通用")`）

### 列头映射表

| 区块 | 表头文本 | 映射键 | 旧版缺列时默认 / 处理 |
|------|----------|--------|------------------------|
| 物料 | **序号** | — | **忽略**（V3 新增，不读） |
| 物料 | 物料名称 | name（取映射列，**非 A 列**） | —（必存在） |
| 物料 | 单位 / 计量单位 | unit | — |
| 物料 | 用量 | usage(float) | — |
| 物料 | 出品率(%) | yield_rate(float) | — |
| 物料 | ERP物料代码 | erp_code | `""` |
| 物料 | 物料类型 | material_type | `其他` |
| 物料 | 所属工序 | process | `""` |
| 配料表 | 物料名称 | （仅用于匹配） | — |
| 配料表 | **用量占比%** | — | **忽略**（不读） |
| 配料表 | **过敏原** | → `material.allergen`（按物料名称匹配回写） | `""` |
| **元件清单** | 物料名称 | （仅用于匹配） | — |
| **元件清单** | 位号(Designator) / 位号 | → `material.designator`（按物料名回收） | `""` |
| **元件清单** | 型号(Part#) / 型号 | → `material.part_number` | `""` |
| **元件清单** | 封装(Footprint) / 封装 | → `material.footprint` | `""` |
| **元件清单** | RoHS | → `material.rohs` | `""` |
| **元件清单** | 制造商 | → `material.manufacturer`（V5 扩列） | `""` |
| **元件清单** | 容差 | → `material.tolerance`（V5 扩列） | `""` |
| **元件清单** | 额定功率 | → `material.rated_power`（V5 扩列） | `""` |
| **元件清单** | 额定电压 | → `material.rated_voltage`（V5 扩列） | `""` |
| **元件清单** | 替代料 | → `material.alternate`（V5 扩列） | `""` |
| **元件清单** | 封装温度 | → `material.reflow_temp`（V5 扩列） | `""` |
| **配方表** | 物料名称 | （仅用于匹配） | — |
| **配方表** | CAS号 / CAS | → `material.cas_number` | `""` |
| **配方表** | 含量(%) / 含量 | → `material.concentration`（float） | `""` |
| **配方表** | GHS标识 / GHS | → `material.ghs_hazard` | `""` |
| **配方表** | 纯度 | → `material.purity`（V5 扩列） | `""` |
| **配方表** | 物态 | → `material.physical_state`（V5 扩列） | `""` |
| **配方表** | 闪点 | → `material.flash_point`（V5 扩列） | `""` |
| **配方表** | 存储条件 | → `material.storage_condition`（V5 扩列） | `""` |
| **配方表** | 危险等级 | → `material.hazard_class`（V5 扩列） | `""` |
| **面料辅料清单** | 物料名称 | （仅用于匹配） | — |
| **面料辅料清单** | 成分比例 | → `material.composition`（V5） | `""` |
| **面料辅料清单** | 纱支 | → `material.yarn_count`（V5） | `""` |
| **面料辅料清单** | 克重(g/m²) | → `material.fabric_weight`（V5） | `""` |
| **面料辅料清单** | 幅宽 | → `material.width`（V5） | `""` |
| **面料辅料清单** | 色号 | → `material.color_no`（V5） | `""` |
| **家具物料清单** | 物料名称 | （仅用于匹配） | — |
| **家具物料清单** | 材质等级 | → `material.material_grade`（V5） | `""` |
| **家具物料清单** | 尺寸规格 | → `material.spec_size`（V5） | `""` |
| **家具物料清单** | 表面处理 | → `material.surface_treatment`（V5） | `""` |
| **家具物料清单** | 色号/花色 / 色号 | → `material.color_no`（V5） | `""` |
| **机械物料清单** | 物料名称 | （仅用于匹配） | — |
| **机械物料清单** | 图号 | → `material.drawing_no`（V6） | `""` |
| **机械物料清单** | 材质 | → `material.material`（V6） | `""` |
| **机械物料清单** | 热处理 | → `material.heat_treatment`（V6） | `""` |
| **机械物料清单** | 表面处理 | → `material.surface_treatment`（V6，与家具同 key） | `""` |
| **机械物料清单** | 重量(kg) / 重量 | → `material.weight`（V6，float） | `""` |
| **机械物料清单** | 单重(kg/件) / 单重 | → `material.unit_weight`（V6，float） | `""` |
| **包装物料清单** | 物料名称 | （仅用于匹配） | — |
| **包装物料清单** | 物料类型 | → `material.material_type`（V6，回收展示列） | `其他` |
| **包装物料清单** | 材质 | → `material.material`（V6，与机械同名 key） | `""` |
| **包装物料清单** | 克重(g/m²) / 克重 | → `material.basis_weight`（V6，float） | `""` |
| **包装物料清单** | 尺寸 | → `material.size`（V6） | `""` |
| **包装物料清单** | 印刷工艺 | → `material.print_process`（V6） | `""` |
| **包装物料清单** | 环保标识 | → `material.eco_label`（V6，自由文本无枚举） | `""` |
| **成本明细** | 物料名称 | （仅用于匹配） | — |
| **成本明细** | 单价 | → `material.unit_price`（V5，float） | `""` |
| **成本明细** | 币种 | → `material.currency`（V5，缺省 人民币(CNY)） | `""` |
| 工序 | 工序编号 | step_no | — |
| 工序 | 工序名称 | name | — |
| 工序 | 工序说明 | desc | `""` |
| 工序 | 工时 | work_hours(float/原值) | `""` |
| 工序 | 备注 | note | `""` |
| 工序 | 产物 | output | `""` |

> 逆向 `_map_header` 时，元件清单/配方表的候选列名包含中英文括号变体（如 `位号(Designator)` / `位号`、`型号(Part#)` / `型号`），以兼容用户手动编辑后的表头。

### 产品级字段解析（表头区）

| 字段 | 定位方式 | 默认 |
|------|----------|------|
| `product_name` | 扫描含 `产品名称` 的单元格，提取冒号后内容 | `""` |
| `category` | 扫描含 `产品类别` 的单元格，提取冒号后内容 | `其他` |
| `industry` | 从「三、」区块标记推断（有元件清单→电子，有配方表→化工，有配料表→食品，有面料辅料清单→纺织，有家具物料清单→家具，有机械物料清单→机械，有包装物料清单→包装，无→按 category 推断） | 按 category 推断 |
| `output_rate` | 扫描含 `全产品出品率` 的单元格，提取数字（去 `%`）转 float | `""` |
| `version` | 含 `版本号` | `V1.0` |
| `date` | 含 `生成日期` | `""` |
| `approver` | 扫描含 `审批人` 的单元格（行 5 可能单格拼接多字段），按 `审批人：` 截取 | `""` |
| `effective_date` | 扫描含 `生效日期` 的单元格，按 `生效日期：` 截取 | `""` |
| `standard` | 扫描含 `执行标准` 的单元格，按 `执行标准：` 截取 | `""` |

> 行 5 在 V3 为单格合并 A5:H5，内容为 `审批人：…    生效日期：…    执行标准：…`（非空段拼接）。逆向按 key 截取对应值，互不干扰；旧版无行 5 内容 → 三者均默认 `""`。
> V4 `industry` **不写入 Excel 表头区**，逆向从「三、」区块标记的存在性隐式推断。

### 分组子标题、合计行与专属区块的逆向处理

- **分组子标题**（`【工序 ...】` / `【未归属工序】`）：首列以 `【` 开头 → 跳过。
- **合计用量行**：首列含「合计」→ 跳过（不读、不回写）。
- **物料名称来源**：取 `_map_header` 映射的「物料名称」列（V3+ 为 B 列；旧版为 A 列）。
- **「三、配料表」过敏原回收（V3 既有，仅食品）**：定位到区块后，读表头映射「物料名称」与「过敏原」列，逐行按物料名匹配已解析 `materials` 中同名项，回写 `allergen`。不重建配料表其他派生列。
- **「三、元件清单」电子专属字段回收（V4 新增，仅电子；V5 扩列至 14 列）**：定位到区块后，读表头映射「物料名称」及 `位号(Designator)`/`型号(Part#)`/`封装(Footprint)`/`RoHS`/`制造商`/`容差`/`额定功率`/`额定电压`/`替代料`/`封装温度` 列，逐行按物料名匹配回写 `designator`/`part_number`/`footprint`/`rohs` + V5 `manufacturer`/`tolerance`/`rated_power`/`rated_voltage`/`alternate`/`reflow_temp`。
- **「三、配方表」化工专属字段回收（V4 新增，仅化工；V5 扩列至 13 列）**：定位到区块后，读表头映射「物料名称」及 `CAS号`/`含量(%)`/`GHS标识`/`纯度`/`物态`/`闪点`/`存储条件`/`危险等级` 列，逐行按物料名匹配回写 `cas_number`/`concentration`/`ghs_hazard` + V5 `purity`/`physical_state`/`flash_point`/`storage_condition`/`hazard_class`。
- **「三、面料辅料清单」纺织专属字段回收（V5 新增，仅纺织）**：定位到区块后，读表头映射「物料名称」及 `成分比例`/`纱支`/`克重(g/m²)`/`幅宽`/`色号` 列，逐行按物料名匹配回写 `composition`/`yarn_count`/`fabric_weight`/`width`/`color_no`。
   - **「三、家具物料清单」家具专属字段回收（V5 新增，仅家具）**：定位到区块后，读表头映射「物料名称」及 `材质等级`/`尺寸规格`/`表面处理`/`色号/花色` 列，逐行按物料名匹配回写 `material_grade`/`spec_size`/`surface_treatment`/`color_no`。
   - **「三、机械物料清单」机械专属字段回收（V6 新增，仅机械）**：定位到区块后，读表头映射「物料名称」及 `图号`/`材质`/`热处理`/`表面处理`/`重量(kg)`/`单重(kg/件)` 列，逐行按物料名匹配回写 `drawing_no`/`material`/`heat_treatment`/`surface_treatment`/`weight`/`unit_weight`（`weight`/`unit_weight` 经 `float_fields` 解析为 float；机械视图无「物料类型」展示列，故不回收 `material_type`）。
   - **「三、包装物料清单」包装专属字段回收（V6 新增，仅包装）**：定位到区块后，读表头映射「物料名称」及 `物料类型`/`材质`/`克重(g/m²)`/`尺寸`/`印刷工艺`/`环保标识` 列，逐行按物料名匹配回写 `material_type`/`material`/`basis_weight`/`size`/`print_process`/`eco_label`（`basis_weight` 经 `float_fields` 解析为 float）。
   - **「成本明细」成本字段回收（V5 新增，跨行业）**：定位到「成本明细」区块（关键字匹配，兼容「三、成本明细」/「四、成本明细」）后，读表头映射「物料名称」及 `单价`/`币种` 列，逐行按物料名匹配回写 `unit_price`/`currency`；`total_price` 不回收（派生展示）。
- **物料对象补全**：逆向输出时，每条物料对象补全全量专属字段（电子 10：`designator`/`footprint`/`part_number`/`rohs`/`manufacturer`/`tolerance`/`rated_power`/`rated_voltage`/`alternate`/`reflow_temp`；化工 8：`cas_number`/`concentration`/`ghs_hazard`/`purity`/`physical_state`/`flash_point`/`storage_condition`/`hazard_class`；纺织 5：`composition`/`yarn_count`/`fabric_weight`/`width`/`color_no`；家具 4：`material_grade`/`spec_size`/`surface_treatment`/`color_no`；成本 2：`unit_price`/`currency`），未回收到的默认空串 `""`。

### 旧版 Excel 兼容（向后兼容 R5）

- 旧版 7 列 / 5 列 Excel 无「序号/占比%/过敏原/审批人/生效日期/执行标准」列 → 列头映射缺失或忽略 → 取默认。
- 旧版无「三、元件清单」/「三、配方表」区块 → 无区块标记 → 按 category 推断 industry → 完全兼容；旧 8 列电子/化工 Excel 逆向仍可正确解析（按列头文本定位，max_col 向后兼容）。
- 旧版物料无专属字段列 → 回收时未匹配到 → 专属字段默认空串。
- 缺 `一、物料信息`/`二、工艺工序` 标记 → `PARSE_ERROR`（退出码 2）；文件不可读 → `FILE_ERROR`（退出码 2）；标题非「BOM表」仅 `WARNING`。

### 闭环示例

```
# 正向：JSON -> Excel
python3 scripts/generate_bom.py --data bom.json --out BOM_2026-07-07.xlsx

# 逆向：Excel -> JSON（导出的 json 可再次喂回正向流程；专属字段按名回收）
python3 scripts/import_bom.py --in BOM_2026-07-07.xlsx --out bom_back.json

# 重新生成：把逆向得到的 JSON 再传给 generate_bom.py
python3 scripts/generate_bom.py --data bom_back.json --out BOM_v2.xlsx
```

逆向导入得到的 JSON 与正向输入 Schema 一致（含 `industry` + 全量专属字段 `_SPECIAL_FIELDS`=37 唯一键：电子 10/化工 8/纺织 5/家具 4/机械 6/包装 5/成本 2），因此可直接作为 `generate_bom.py --data` 的输入，实现「导入 → 编辑/查看 → 重新生成」的完整闭环（专属字段经回收后保留）。
