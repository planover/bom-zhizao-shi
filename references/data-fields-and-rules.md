# 《数据字段定义与业务规则》（BOM 智造师 · 增量增强 V2）

> 文档类型：简单 PRD（需求/规则定义，不含代码）
> 适用范围：BOM 智造师正向 `generate_bom.py` 与逆向 `import_bom.py` 的字段与业务规则增强
> 配套决策：产品类别枚举 5 类、工序 BOM 平铺建模 + 分工序呈现、演示截图以 SVG/HTML 还原

---

## 1. 产品目标

本次增强为 BOM 增加**产品级元数据（名称、类别、出品率）**与**工序级物料流转链**，并针对「食品」类自动生成**配料表**（仅含食用物料），使 BOM 既能表达跨工序的物料流转，又能满足食品合规配料标注需求，同时保持对旧版 Excel 的向后兼容。

---

## 2. 数据字段定义表

### 2.1 BOM 级（顶层 JSON 字段）

| 字段名 (JSON 键) | 层级 | 类型 | 必填性 | 校验规则 | 说明 / 备注 |
|---|---|---|---|---|---|
| `product_name` | BOM 级 | string | **必填** | 非空校验，不允许空串或纯空白 | 产品名称；旧版为「可选」，本次改为**必填非空**（R1） |
| `output_rate` | BOM 级 | number | **必填** | `> 0`；允许 `> 100`（如干香菇泡发增重）；无硬性上限；字段说明注明实际业务多数 `≤ 100` | 全产品出品率(%)；本次新增 |
| `category` | BOM 级 | string(enum) | **必填** | 枚举：`食品` / `工业品` / `日化化妆品` / `医药` / `其他`；下拉选择或分类引用 | 仅 `食品` 触发配料表（R4） |
| `version` | BOM 级 | string | 选填 | 默认 `V1.0` | 现有字段，未变更 |
| `date` | BOM 级 | string | 选填 | 默认当天 `YYYY-MM-DD` | 现有字段，未变更 |

### 2.2 物料级（`materials[]` 元素字段）

| 字段名 (JSON 键) | 层级 | 类型 | 必填性 | 校验规则 | 说明 / 备注 |
|---|---|---|---|---|---|
| `name` | 物料级 | string | 必填 | 非空 | 物料名称；现有字段 |
| `unit` | 物料级 | string | 必填 | 非空 | 单位；现有字段 |
| `usage` | 物料级 | number | 必填 | `> 0` | 用量；现有字段 |
| `yield_rate` | 物料级 | number | 必填 | `0 < 值 ≤ 100` | 物料出品率(%)；现有字段，校验区间不变（R2） |
| `erp_code` | 物料级 | string | 选填 | 遵循现有物料代码命名规则分配 ERP 物料代码，允许留空（默认 `""`） | 现有字段，未变更 |
| `material_type` | 物料级 | string(enum) | 选填（食品类建议必填） | 枚举：`原料` / `添加剂` / `香精香料` / `包材` / `其他` | 本次新增；用于配料表过滤（R4） |
| `process` | 物料级 | string | 选填（首道工序可选填） | 若填写须引用已录入的有效工序 `step_no` | 本次新增；「所属工序」归属字段，实现「每条物料带 process 引用工序编号」与 Excel 分工序呈现 |

> 说明：物料清单为**平铺单数组**（`materials` 一份），每条物料以 `process` 字段归属工序；Excel 中按 `process` 分组（分工序）呈现。

### 2.3 工序级（`processes[]` 元素字段）

| 字段名 (JSON 键) | 层级 | 类型 | 必填性 | 校验规则 | 说明 / 备注 |
|---|---|---|---|---|---|
| `step_no` | 工序级 | string | 必填 | 唯一不重复（如 `S01`） | 现有字段 |
| `name` | 工序级 | string | 必填 | 非空 | 工序名称；现有字段 |
| `desc` | 工序级 | string | 选填 | — | 工序说明；现有字段 |
| `work_hours` | 工序级 | number/string | 选填 | 数值须 `≥ 0` | 工时；现有字段 |
| `note` | 工序级 | string | 选填 | — | 备注；现有字段 |
| `output` | 工序级 | string | **必填** | 非空；为该工序产物名称 | 本次新增；用于工序间物料流转链（R3） |

### 2.4 数据模型关系（Mermaid）

```mermaid
classDiagram
    class BOM {
        +string product_name  «必填非空»
        +number output_rate  «必填, >0»
        +enum category  «必填, 5类»
        +string version
        +string date
    }
    class Material {
        +string name «必填»
        +string unit «必填»
        +number usage «必填>0»
        +number yield_rate «0<值≤100»
        +string erp_code «选填»
        +enum material_type «选填»
        +string process «选填, 引用step_no»
    }
    class Process {
        +string step_no «必填唯»
        +string name «必填»
        +string desc
        +work_hours
        +string note
        +string output «必填, 产物»
    }
    BOM "1" o-- "0..*" Material : materials[]
    BOM "1" o-- "0..*" Process : processes[]
    Material "..>" Process : process 引用 step_no
    Process "产物⟶下一工序物料" Process : 流转链(R3)
```

---

## 3. 业务规则

| 规则 | 名称 | 规则定义 | 备注 |
|---|---|---|---|
| **R1** | 非空校验 | `product_name`、`category` **必填且非空**（不允许空串/纯空白） | 旧版 `product_name` 可选，本次收紧 |
| **R2** | 出品率 | BOM 级 `output_rate`：`> 0`，允许 `> 100`（无硬性上限，字段说明注明多数 `≤ 100`）；物料级 `yield_rate`：`0 < 值 ≤ 100` | 两级出品率并存，区间不同 |
| **R3** | 工序流转链 | `processes` 按序；`process[i]` 的物料清单（即 `process == process[i].step_no` 的物料集合）**必须包含 `process[i-1].output` 作为一条物料**（首道工序除外）；若仅 0–1 道工序则该规则不触发 | 流转链匹配依据：物料 `name` 等于上一工序 `output`（见待确认 Q2） |
| **R4** | 配料表条件生成 | 仅当 `category == "食品"` 时生成配料表；配料表 = 物料中 `material_type ∈ {原料, 添加剂, 香精香料}` 的集合，**排除 `包材` / `其他`**（即排除纸箱、PE 袋、胶带、标签等非食用物料） | 非食品类不生成配料表 |
| **R5** | 类别联动与向后兼容 | ① `category` 与配料表联动：改类别即时影响是否生成配料表；② 旧 Excel / 旧 JSON 缺新字段时按默认值处理（`material_type` 默认 `其他`、`process` 默认空、`output_rate` 默认空串待补、`category` 默认 `其他` 且不生成配料表） | 保证旧数据可正常导入、重新生成 |

### 3.1 工序流转链示意（Mermaid）

```mermaid
flowchart LR
    P1[工序 S01<br/>产物 = O1] -->|物料清单含 O1| P2[工序 S02<br/>产物 = O2]
    P2 -->|物料清单含 O2| P3[工序 S03<br/>产物 = O3]
    style P1 fill:#e3f2fd,stroke:#1976d2
    style P2 fill:#e3f2fd,stroke:#1976d2
    style P3 fill:#e3f2fd,stroke:#1976d2
```

> 说明：每道工序的物料区都须把上一工序的产物列为一条输入物料，从而形成可追溯的物料流转链（R3）。

---

## 4. 用户故事

1. **食品类自动配料表** — As a 食品研发工程师，I want to 将产品类别选为「食品」并标注每个物料的 `material_type`，so that BOM 自动生成仅含原料/添加剂/香精香料的配料表，且自动排除纸箱、胶带、标签等包材。

2. **多工序物料流转链** — As a 工艺工程师，I want to 为每道工序填写 `output` 且下一道工序物料含上一道产物，so that 工序间物料流转链完整可追溯，校验器能自动报出断链工序。

3. **干香菇泡发出品率 > 100** — As a 配方/采购员，I want to 在 BOM 级 `output_rate` 填写 130（干香菇泡发后增重），so that 系统接受大于 100 的出品率而不报错，正确表达增重场景。

4. **纯物料无工序 BOM** — As a 简易 BOM 用户，I want to 只录入物料、不录入任何工序，so that 系统跳过 R3 流转链校验仍能产出合法 BOM，且 `materials[].process` 留空不影响生成。

---

## 5. 需求池（优先级）

| 优先级 | 项 | 说明 |
|---|---|---|
| **P0（必做字段与规则）** | `product_name` 改为必填非空 | R1 |
| | 新增 BOM 级 `output_rate`（必填，`> 0`，可 `> 100`） | R2 |
| | 新增 BOM 级 `category`（必填，枚举 5 类） | R1 / R4 |
| | 物料级 `yield_rate` 保持 `0 < 值 ≤ 100` 校验 | R2 |
| | 新增工序级 `processes[].output`（必填，产物） | R3 |
| | 新增物料级 `materials[].process`（工序归属，首道可选填） | 流转链/分工序呈现 |
| **P1（配料表 / 流转链校验）** | 新增物料级 `material_type` 枚举（原料/添加剂/香精香料/包材/其他） | R4 |
| | R4 配料表条件生成（仅 `食品`，过滤排除包材/其他） | 核心增强 |
| | R3 工序流转链校验（断链告警/拦截） | 核心增强 |
| | R5 向后兼容与类别联动（旧字段默认值处理） | 兼容保障 |
| **P2（文档与演示）** | 以 SVG/HTML 还原 BOM 视觉图并嵌入 README（演示截图） | 环境无 Excel GUI |
| | 生成 CHANGELOG 更新日志，记录本次变更 | 仓库维护 |

---

## 6. UI / 交互要点（供 SKILL.md 与工程师参考）

| 阶段 | 采集项 | 交互要求 |
|---|---|---|
| **阶段零（产品信息）** | `product_name` | 标记为**必填**，非空校验；旧版可选，本次收紧 |
| | `category` | 新增**下拉**：`食品 / 工业品 / 日化化妆品 / 医药 / 其他`，必填 |
| | `output_rate` | 新增**数值输入**（必填，`> 0`，允许 `> 100`），占位提示「如干香菇泡发增重可 >100，通常 ≤100」 |
| **阶段一（物料）** | `material_type` | 每条物料新增**下拉**：`原料 / 添加剂 / 香精香料 / 包材 / 其他` |
| | `process` | 每条物料新增**下拉**：引用已录入工序 `step_no`，**首道工序可选填** |
| | `yield_rate` | 保持 `0 < 值 ≤ 100` 校验提示 |
| **阶段二（工序）** | `output` | 每道工序新增**产物（output）必填输入**，作为流转链源头 |
| | 工序排序 | 按 `step_no` 升序维护，供 R3 顺序校验 |
| **汇总确认** | 校验提示 | 体现：① `product_name`/`category` 非空；② `output_rate > 0`；③ R3 流转链完整性；④ `category == 食品` 时预览配料表并提示已排除包材/其他 |

---

## 7. 待确认问题

1. **`material_type` 在非食品类是否强制必填？** 建议：非食品类可仅选 `包材 / 其他` 或允许留空；食品类建议必填以保证配料表完整。需用户拍板是否对所有类别强制必填。
2. **工序流转链的匹配方式**：当前按「物料 `name` == 上一工序 `output`」字符串精确匹配，存在重名/空格导致误判风险。是否引入产物引用 ID（如 `output_id`）替代名称匹配？
3. **纯物料（无工序）BOM 的 `materials[].process` 呈现**：留空时 Excel 是否统一归为「未归属工序」分组区块，还是与有工序物料合并呈现？
4. **BOM 级 `output_rate` 与物料级 `yield_rate` 的关系**：是否在 Excel 汇总区新增「全产品出品率」展示行？二者是否需一致性校验（如 BOM 级 ≈ 物料级加权）？
5. **配料表在 Excel 中的呈现位置**：作为物料区下方独立「配料表」区块，还是物料区加「是否计入配料表」标记列？需与逆向导入（`import_bom.py`）解析规则对齐。
