# Changelog · BOM智造师（bom-zhizao-shi）

## [V5.0] - 2026-07-11

> P1 全部实现：成本视图 + 行业模板预设；纺织/家具专属派生视图；电子/化工扩列并入单区块。100% 向后兼容，不新增任何阻断/软校验。机械/包装本期仅出评估草案（不实现）。

### 行业专属派生视图扩充（正向）
- **纺织（industry=="纺织"）**：新增「三、面料辅料清单」（8 列 A–H），排除 `material_type ∈ {"其他"}`，按 (物料类型, 名称) 升序；表头 `序号/物料名称/物料类型/成分比例/纱支/克重(g/m²)/幅宽/色号`；`derive_textile(data)` 与 V4 电子/化工同构。
- **家具（industry=="家具"）**：新增「三、家具物料清单」（8 列 A–H），排除 `material_type ∈ {"其他"}`，按 (物料类型, 名称) 升序；表头 `序号|物料名称|物料类型|材质等级|尺寸规格|表面处理|用量|色号/花色`（8 列 A–H，末列「色号/花色」为单表头）；`derive_furniture(data)` 同构。
- **电子「三、元件清单」扩列 14 列（A–N）**：在 V4 8 列基础上新增 制造商/容差/额定功率/额定电压/替代料/封装温度 6 个工程/合规字段（共 10 专属字段）；列宽扩展至 14 列。
- **化工「三、配方表」扩列 13 列（A–M）**：在 V4 8 列基础上新增 纯度/物态/闪点/存储条件/危险等级 5 个 SDS 关键字段（共 8 专属字段）；列宽扩展至 13 列。

### 成本明细视图（跨行业，双编号）
- **触发条件**：任一物料 `unit_price` 非空（≠"" 且可转 float）即纳入；`derive_cost(data)` 面向全物料（含行业视图排除的"其他"/"包材"类）。
- **双编号**：有行业专属视图（食品/电子/化工/纺织/家具）→「四、成本明细」（位于行业视图之后）；否则 →「三、成本明细」。
- **列定义（8 列 A–H）**：`序号/物料名称/物料类型/用量/单位/单价/币种/总价`；单价/总价数字格式 `0.00`。
- **派生规则**：`total_price` **不入库**，渲染时按 `usage × unit_price` 实时计算；末行「成本合计」= Σ总价（纯展示，逆向跳过）。
- **`unit_price`（单价，选填）**：数值，默认 `""`（不填则不进成本视图）；`currency`（币种，选填）默认 `人民币(CNY)`，缺省按人民币(CNY)处理。

### 行业模板预设（INDUSTRY_TEMPLATES，仅交互引导）
- `bom_constants.py` 新增 `INDUSTRY_TEMPLATES`：按 industry 给出 `material_types`（物料类型下拉建议）/ `standard`（执行标准自动预填，可改）/ `special_fields`（阶段一模板动态追加的专属字段行）/ `preset_processes`（可选「一键载入工序模板」）。
- 纺织/家具/电子/化工/食品模板填满，通用/机械/包装为空模板（保持通用兜底，不预置专属字段与工序）。
- 交互层只读此模板，**不写入任何新 JSON 结构字段**，向后兼容旧 JSON（最小变更原则）。

### 共享常量模块（bom_constants.py）增量
- 新增 `TEXTILE_TYPES`/`TEXTILE_EXCLUDE`（纺织过滤集）、`FURNITURE_TYPES`/`FURNITURE_EXCLUDE`（家具过滤集）。
- `INDUSTRY_STANDARD` 增 纺织=FZ/T 80004、家具=QB/T 1951.1。
- 导出 `EDIBLE_LIST`（供食品模板 special_fields 引用）。

### 逆向导入增强（import_bom.py）
- 区块定位新增 面料辅料清单/家具物料清单/成本明细 marker；`_infer_industry_from_blocks` 增 纺织/家具 推断（有面料辅料清单→纺织，有家具物料清单→家具）。
- 电子区块按 14 列识别（`_map_header(ws, row, max_col=14)`），化工区块按 13 列识别（`max_col=13`），恢复 V5 扩列字段（制造商/容差/额定功率/额定电压/替代料/封装温度、纯度/物态/闪点/存储条件/危险等级），旧 8 列 Excel 仍按默认 max_col=8 兼容。
- 成本明细区块回收 `unit_price`/`currency`（逆向关键字 `成本明细` 识别），`total_price` 不回收（派生）。
- 物料对象补全全量专属字段默认空串（电子 10/化工 8/纺织 5/家具 4/成本 2 + V4 7 个）。

### 软校验（V5 不新增任何阻断/软校验）
- 沿用 V4 既有 W2（电子 RoHS 未标）、W3（化工 CAS/GHS 未填）、含量和校验（±5%）。
- 纺织/家具/成本/电子扩列/化工扩列所有新增字段均不设软校验（最低回归风险）。

### 文档与示例
- `references/bom-spec.md`：更新输入 JSON Schema（纺织 5/家具 4/电子 10/化工 8/成本 2 字段）、Excel 视图与列宽（纺织/家具/电子14/化工13/成本明细）、逆向章节（区块定位/列头映射/industry 推断扩展/成本回收）、INDUSTRY_TEMPLATES 章节。
- `SKILL.md`：行业专属字段表格扩充至 V5（电子 10/化工 8/纺织 5/家具 4/成本 2 + INDUSTRY_TEMPLATES 引导）；阶段一模板按 `special_fields` 动态渲染；汇总确认展示纺织/家具/扩列/成本预览；执行标准增 纺织/家具；说明 V5 不新增软校验。
- `README.md`：字段校验表增纺织/家具/扩列/成本；Excel 结构增纺织/家具视图、电子/化工扩列、成本明细双编号；行业模板预设说明；已知限制更新（机械/包装仅评估）。
- `references/bom-demo.svg`：替换 V4 电子/化工区块为 V5 14/13 列版本，追加纺织/家具/成本明细示意图。
- `references/mechanical-packaging-draft-v5.md`：P2 机械/包装评估草案（本期不实现，仅评估）。
- `examples/`：新增 `sample_bom_v5_textile`/`sample_bom_v5_furniture`/`sample_bom_v5_electronic`/`sample_bom_v5_chemical`/`sample_bom_v5_cost` 共 5 组 `.json`/`.xlsx`（展示纺织/家具视图、电子14列/化工13列扩列、通用行业「三、成本明细」）。

---

## [V4.0] - 2026-07-09

> 行业标识 + 专属派生视图 + 软校验增强；物料区 8 列不变，新增 `bom_constants.py` 共享常量模块；向后完全兼容旧 JSON/Excel。

### 新增字段（正向 / 逆向 Schema 一致）
- **`industry`**（行业标识，选填）：BOM 级可选字段，8 值枚举 {食品,电子,化工,机械,纺织,家具,包装,通用}；默认按 `category` 推断（食品→食品，日化/医药→化工，工业品/其他→通用）；非法值 V8 WARNING 回退推断（非阻断）；决定专属派生视图（食品→配料表，电子→元件清单，化工→配方表）。
- **`materials[].designator`**（位号，选填）：电子元件位号（如 R1、C3、U5），默认 `""`；仅电子元件清单展示。
- **`materials[].footprint`**（封装，选填）：电子元件封装（如 0805、SOIC-8），默认 `""`。
- **`materials[].part_number`**（型号，选填）：电子元件型号（如 STM32F103C8T6），默认 `""`。
- **`materials[].rohs`**（RoHS 合规状态，选填）：∈ {是,否,未知}，默认 `""`（等价未知）；仅电子元件清单展示；着色规则：否→红字 FF0000，未知/空→黄字 BF8F00，是→默认。
- **`materials[].cas_number`**（CAS 号，选填）：化学品 CAS 登记号（如 7732-18-5），默认 `""`；仅化工配方表展示。
- **`materials[].concentration`**（含量(%)，选填）：配方中该原料的含量百分比，默认 `""`；仅化工配方表展示；空则留空不显示 0.0%。
- **`materials[].ghs_hazard`**（GHS 危险标识，选填）：如 GHS07，默认 `""`；仅化工配方表展示。

### 共享常量模块（bom_constants.py）
- 新建 `scripts/bom_constants.py`（纯 Python 标准库，无第三方依赖），定义 `INDUSTRIES`、`CATEGORY_TO_INDUSTRY`、`COMPONENT_TYPES`、`COMPONENT_EXCLUDE`、`FORMULA_TYPES`、`FORMULA_EXCLUDE`、`INDUSTRY_STANDARD`。
- 从 `generate_bom.py` 迁入 `EDIBLE`、`ALLERGEN_SET`、`ALLERGEN_HINTS`（单一真相源，避免两脚本各自复制漂移）。
- `generate_bom.py` 与 `import_bom.py` 均 `from bom_constants import ...`。

### 业务规则与派生视图
- **行业推断**：`infer_industry(data)` → 显式优先 → category 推断 → V8 WARNING 回退。
- **元件清单（电子）**：`derive_components(data)` 排除 `material_type ∈ {"其他"}`，按 `(material_type, designator)` 升序排序（空位号用 `\uffff` 哨兵排末尾）；Excel 「三、元件清单」8 列 A–H：`序号|位号(Designator)|型号(Part#)|封装(Footprint)|物料名称|数量|物料类型|RoHS`；RoHS 红黄字标记。
- **配方表（化工）**：`derive_formula(data)` 排除 `material_type ∈ {"包材"}`，按 `concentration` 降序排序（空排末尾）；Excel 「三、配方表」8 列 A–H：`序号|物料名称|CAS号|含量(%)|GHS标识|物料类型|计量单位|用量`；含量(%) 数字格式 `0.0"%"`。
- **配料表触发变更**：从 `category=="食品"` 改为 `industry=="食品"`（含推断，行为不变）。
- **执行标准行业建议**：食品→GB 7718-2025，电子→GB/T 39560，化工→GB/T 16483-2008。

### 软校验增强（非阻断，退出码 0）
- **V8**：`industry` 非空但不在枚举内 → `WARNING: industry 值『{value}』不在枚举内（…），已回退为推断值`。
- **W2**：电子行业，元件清单内物料未标 `rohs` → `WARNING: 物料『{name}』未标注 RoHS 合规状态，请确认`。
- **W3a**：化工行业，配方表内物料未填 `cas_number` → `WARNING: 物料『{name}』未填写 CAS 号，请确认`。
- **W3b**：化工行业，配方表内物料未填 `ghs_hazard` → `WARNING: 物料『{name}』未填写 GHS 危险标识，请确认`。
- **含量和校验**：化工行业，所有配方原料均填 concentration 时校验列和 ≈ 100%（±5%）→ `WARNING: 配方表含量(%) 列和为 {total:.1f}%，偏离 100% 超过 ±5%，请确认`。
- **排除提示（电子）**：`WARNING: 元件清单已排除 {N} 条非元件物料（其他类）：{names}`。
- **排除提示（化工）**：`WARNING: 配方表已排除 {N} 条包材物料：{names}`。
- W1/H1 过敏原软校验沿用 V3（不变）。

### 逆向导入增强（import_bom.py）
- 新增 `_infer_industry_from_blocks(ws, category)`：从「三、」区块标记推断 industry（有元件清单→电子，有配方表→化工，有配料表→食品，无→按 category 推断）。
- 新增 `_recover_block_fields(ws, marker_row, field_col_map, materials)`：通用区块字段回收函数，按物料名匹配回写专属字段。
- 从「三、元件清单」回收 `designator`/`part_number`/`footprint`/`rohs`；从「三、配方表」回收 `cas_number`/`concentration`/`ghs_hazard`。
- 物料对象补全 7 个专属字段默认空串（未回收到的）。
- 输出 JSON 增 `industry` 字段。
- 工序区停止边界扩展为所有「三、」区块标记。
- 旧 Excel（无 industry / 无专属区块）完全兼容。

### 向后兼容
- 旧 JSON 缺 `industry` → 按 `category` 推断（行为零变化）。
- 旧 JSON 缺 7 个专属物料字段 → 默认空串。
- 旧 JSON `industry` 非法值 → V8 WARNING + 回退推断。
- 旧 Excel 无「三、元件清单」/「三、配方表」区块 → 按 category 推断 → 完全兼容。
- 现有食品 Excel（有「三、配料表」）→ 推断 industry=食品 → 配料表照常 → 行为零变化。

### 文档与示例
- `references/bom-spec.md`：更新 JSON Schema（industry + 7 专属字段）、Excel 列定义（元件清单/配方表）、industry 推断逻辑、逆向规则、向后兼容表。
- `SKILL.md`：阶段零追加 industry 可选问题；阶段一物料模板按 industry 动态追加专属字段；汇总确认展示对应专属视图预览；执行标准行业建议；数据校验补 V8/W2/W3 说明。
- `README.md`：字段校验表加 industry/专属字段；Excel 结构说明加元件清单/配方表区块；向后兼容说明；已知限制更新；JSON 示例增 industry/专属字段。
- `references/bom-demo.svg`：追加电子元件清单/化工配方表区块示意图。
- `examples/`：新增 `sample_bom_v4_electronic.json`/`.xlsx`（电子行业，含 RoHS 着色）与 `sample_bom_v4_chemical.json`/`.xlsx`（化工行业，含含量% 格式）。

---

## [V2.1] - 2026-07-09（内部版本号记为 V3）

> 最小变更、向后兼容的增量；Excel 列数 7→8（A–H），新增可选元数据与配料表占比/过敏原，无新增阻断级校验。

### 新增字段（正向 / 逆向 Schema 一致）
- **`approver`**（审批人，选填）：BOM 级可选字段，默认 `""`；非空时与 `effective_date`/`standard` 拼接显示于表头行 5（单格合并 A5:H5，段间 4 空格分隔）。
- **`effective_date`**（生效日期，选填）：BOM 级可选字段，默认 `""`。
- **`standard`**（执行标准，选填）：BOM 级可选字段，默认 `""`；食品建议填 `GB 7718-2025`。
- **`materials[].allergen`**（过敏原，选填）：逗号分隔标签，∈ 八大类+其他（含麸质谷物/甲壳类/蛋类/鱼类/花生/大豆/乳/坚果/其他），默认 `""`；仅食品类「三、配料表」展示，逆向按物料名称回收回写 `material.allergen`。

### 业务规则与派生展示
- **物料区扩为 8 列（A–H）**：首列新增「序号」，按 `materials[]` 输入顺序全局连续赋值、跨工序分组不重置（纯展示，不进 JSON）。
- **物料区末「合计用量」行**：A 列写「合计」，D 列写全部物料 `usage` 求和（含包材/其他）；纯展示，逆向跳过（首列含「合计」即跳过）。
- **R4 配料表增强**：仅 `食品` 类生成「三、配料表」派生区块，新增「用量占比%」列与「过敏原」列。
  - **用量占比%**：分母 = 可食用物料用量合计，每项 `round(usage/合计×100, 1)`；采用最大余数法对末位补差（0.1 粒度），保证整列合计**恰为 100.0%**（覆盖旧版「99.9% 可接受」表述，现强制 100.0）。数字格式 `0.0"%"`.
  - **过敏原**：取该物料 `allergen`（空则留空）。
- **W1 过敏原软校验（非阻断）**：`allergen` 标签非空且任一标签 ∉ 八大类+其他集合时，打印 `WARNING: 物料『{name}』的过敏原标签…不在八大类集合内，请确认`，不进入 errors、不阻断生成。
- **H1 过敏原关键词启发式软校验（非阻断）**：物料名称含致敏物关键词（如 牛奶/奶→乳、蛋→蛋类、花生、大豆/黄豆、麸质/面筋、坚果、虾/蟹、鱼 等，共 28 条映射于 `ALLERGEN_HINTS`）但其 `allergen` 未涵盖对应类别时，打印 `WARNING: 物料『{name}』名称疑似含致敏物『{class}』但未在过敏原中标注，请确认`；已正确标注 / 名称无关键词不告警，不进入 errors、不阻断生成。与 W1 并存，为 GB 7718-2025 合规打磨项。
- **分组子标题美化**：`【工序 Sxx 名称】` / `【未归属工序】` 子标题改为浅蓝底（EAF1FB）+ 左侧加粗蓝色色条 + 加粗。
- 全产品出品率仍显示 `130.0%`（1 位小数，沿用 V2 修正）。

### Excel 输出结构（8 列 A–H）
- 行号固定：标题(1)/版本号+生成日期(2)/产品名称(3)/产品类别+全产品出品率(4)/**审批人+生效日期+执行标准(5，可选单格合并 A5:H5)**/一、物料信息(6)/表头(7)/数据(8+) 含分组子标题与「合计用量」行/二、工艺工序/三、配料表(仅食品)。
- 所有合并标题行合并 **A–H**；工序区/配料表区数据仍 7 列（A–G，H 留空）。
- 物料表头（8 列）：`序号|物料名称|单位|用量|出品率(%)|ERP物料代码|物料类型|所属工序`。
- 配料表表头（7 列）：`物料名称|物料类型|计量单位|用量|出品率(%)|用量占比%|过敏原`。
- 列宽（8 列）：序号6 / 物料名称18 / 单位10 / 用量10 / 出品率13 / ERP物料代码16 / 物料类型13 / 所属工序12。

### 逆向导入增强（import_bom.py）
- 物料名称改取 `_map_header` 映射列（**非 A 列**，因 V3 首列为序号）；忽略「序号」「用量占比%」列。
- 解析新增 `approver` / `effective_date` / `standard`（行 5 单格合并拼接场景，按 key 截取对应值）；旧版无此内容则默认 `""`。
- 物料行读取新增「首列含『合计』则跳过」逻辑；「三、配料表」仅回收「过敏原」列（按物料名称匹配回写 `material.allergen`），不重建配料表其他派生列。
- 输出 JSON 含 `approver`/`effective_date`/`standard`/`materials[].allergen`，**不含**「序号」「用量占比%」。
- 向后兼容：旧版 7 列 / 5 列 Excel 逆向完全兼容，缺新列取默认（`approver`/`effective_date`/`standard`/`allergen` 为空），不报错。

### 文档与演示
- `references/bom-spec.md`：更新输入 JSON Schema（含 `approver`/`effective_date`/`standard`/`materials[].allergen`）、Excel 8 列输出结构（行5 三字段/序号列/合计行/配料表占比%/过敏原）、逆向导入规则（列头定位+过敏原回收+合计行跳过+执行标准解析+兼容）。
- `SKILL.md`：阶段零加 `approver`/`effective_date`/`standard`（可选，空则不显示）；阶段一物料加 `allergen`（可选，逗号分隔八大类）；汇总预览补「用量占比%」「过敏原」；校验补 W1 软告警话术。
- `README.md`：字段/校验速查表增新字段；Excel 结构说明（8 列/序号/合计行/审批人/执行标准/占比%/过敏原）；`bom.json` 示例增 `allergen`/`approver`/`standard`；已知限制补充向后兼容说明。
- `references/bom-demo.svg`：还原 8 列表头 + 序号 + 合计行 + 行5 三字段 + 配料表占比%/过敏原 + 分组子标题美化。
- `examples/`：新增 `sample_bom_v3.json`（食品类多工序，含过敏原+审批人+执行标准）与生成的 `sample_bom_v3.xlsx`。

---

## [V2.0] - 2026-07-08

### 新增字段（正向 / 逆向 Schema 一致）
- **`product_name`** 由「可选」改为**必填非空**（R1），Excel 固定输出「产品名称」整行。
- **`category`**（产品类别，必填）：枚举 `食品 / 工业品 / 日化化妆品 / 医药 / 其他`；仅 `食品` 触发配料表（R1/R4）。
- **`output_rate`**（全产品出品率，必填）：数值 `> 0`，允许 `> 100`（无硬上限，如干香菇泡发增重）；Excel 表头区新增「全产品出品率」行（数字格式 `0.0"%"`）（R2）。
- **`materials[].material_type`**（物料类型，选填）：枚举 `原料 / 添加剂 / 香精香料 / 包材 / 其他`，默认 `其他`；用于配料表过滤（R4）。
- **`materials[].process`**（所属工序，选填）：引用有效 `step_no`，首道可选填；用于分工序分组呈现与 R3 流转链匹配。
- **`processes[].output`**（产物，必填）：该工序产物名称，作为工序流转链（R3）源头。

### 业务规则（R1–R5）
- **R1 非空校验**：`product_name`、`category` 必填且非空。
- **R2 出品率**：BOM 级 `output_rate` `> 0`（可 >100）；物料级 `yield_rate` 维持 `0 < 值 ≤ 100`。
- **R3 工序流转链（阻断级）**：当工序数 ≥2 时，上一工序 `output` 必填，且下一道工序的物料清单（按 `process == 该工序.step_no` 过滤）须包含上一工序 `output` 作为一条物料（名称精确匹配）；断链时纳入 `errors`，打印 `VALIDATION_FAILED` 并以**退出码 2** 结束（不降级为 WARNING）。
- **R4 配料表条件生成**：仅 `category == "食品"` 时生成「三、配料表」派生区块，仅收 `原料 / 添加剂 / 香精香料`，按用量降序排列；包材/其他/未分类被排除并输出非阻断 `WARNING`。
- **R5 向后兼容**：旧 JSON / 旧 Excel 缺新字段时按默认值处理（`material_type=其他`、`process=空`、`output_rate=空`、`category=其他`、`output=空`）；重新生成时缺必填项触发 `VALIDATION_FAILED`（待补，详见 README「已知限制」）。

### Excel 输出结构（7 列 A–G）
- 由 5 列扩展为 **7 列**：物料区新增「物料类型」「所属工序」两列；工序区新增「产物」列（G 留空）。
- 行号固定（因 `product_name` 必填）：标题(1) / 版本号+生成日期(2) / 产品名称(3) / 产品类别+全产品出品率(4) / 空行(5) / 一、物料信息(6) / 表头(7) / 物料数据(8+) / 二、工艺工序 / 三、配料表(仅食品)。
- **分工序分组呈现**：有工序且存在归属时，按 `step_no` 升序插入 `【工序 Sxx 名称】` 子标题（浅色填充），未归属物料归 `【未归属工序】`；无工序或全空则平铺。
- 列宽：A18 / B10 / C10 / D12 / E16 / F14 / G12。

### 逆向导入增强（import_bom.py）
- `parse_bom()` 由硬编码 A–E 改为**按列头文本映射列号**，对旧版 5 列 Excel 天然兼容（缺列取默认）。
- 解析新增 `category` / `output_rate` / `material_type` / `process` / `output`。
- 解析物料时跳过 `【分组】` 子标题行；定位到「三、配料表」即停止，**不解析、不回写配料表**（派生数据，重新生成时按 `category` 重新派生）。
- 错误码不变：`FILE_ERROR` / `PARSE_ERROR`（退出码 2），标题非「BOM表」仅 `WARNING`。

### 文档与演示
- `references/bom-spec.md`：重写输入 JSON Schema、Excel 7 列输出结构、逆向导入解析规则。
- `SKILL.md`：阶段零/一/二 新增字段与校验引导、汇总确认新增配料表预览与排除提示。
- `README.md`：字段校验表、Excel 结构（7 列/分组/配料表）、演示截图（`references/bom-demo.svg`）、合法示例、已知限制。
- `references/bom-demo.svg`：还原 Excel BOM 视觉的演示图（无 GUI 环境生成）。
- `examples/`：新增 `sample_bom_v2.json`（食品类多工序合法样例）与生成的 `sample_bom_v2.xlsx`。

---

## [V1.0] - 2026-07-07

- 初始版本：5 列 BOM 表（物料名称/单位/用量/出品率/ERP物料代码 + 工艺工序），正向生成与逆向导入闭环。
- `product_name` 可选；无产品类别 / 全产品出品率 / 物料类型 / 所属工序 / 工序产物等字段。
