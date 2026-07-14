# BOM 智造师 V7 增量产品需求文档（增量 PRD）

> 作者：许清楚（产品经理）
> 日期：2026-07-15（V6 交付后增量）
> 版本：V7 增量 PRD（基于 V6，commit af72a95）
> 范围：**仅描述变更（增量）部分**。未提及的 V6 契约（物料区 8 列 A–H、industry 8 值枚举与推断逻辑、V8/W1/H1/W2/W3/含量和软校验、机械/包装/纺织/家具/电子/化工/食品七大专属视图、成本视图双编号、37 唯一键 `_SPECIAL_FIELDS`、行业模板预设、完整逆向闭环）全部沿用，不重复。
> 格式：标准**简单 PRD**（默认档，不含竞品分析）。

---

## 1. 项目信息（增量）

| 项 | 内容 |
|----|------|
| Language | 中文 |
| Project Name | bom-zhizao-shi (V7 增量) |
| 基线版本 | V6（commit af72a95）：机械/包装专属视图 + 成本视图双编号 + 行业模板预设填实；V2–V6 共 107 例测试全绿；`_SPECIAL_FIELDS`=37 唯一键 |
| 原始需求复述 | 用户明确「走增量」并确认三个方向：**A. 多语言导出（用户选「双语对照并列」）→ B1 批量空白模板 + B2 批量生成 BOM + B3 批量合并导入（用户选「以上」，即三种批量均实现）→ 全面巩固（用户选「全面巩固」，含测试/边界/UX/性能）** |
| 核心约束（硬约束，沿用 V6 并扩展） | ① 物料区 8 列（A–H）**永不变**；② 向后兼容（旧 JSON/Excel 零结构变化，新字段默认空）；③ 不引入新第三方依赖（仅 openpyxl）；④ **不新增任何软校验 WARNING**（最小变更、最低回归风险）；⑤ `industry` 枚举无需扩展；⑥ 旧单语导出默认行为不变，新增开关向后兼容旧调用 |
| 硬约束（双语形态，本 PRD 推荐并论证，列为待确认 Q1） | **推荐**：主 sheet「BOM表」保持纯中文（逆向唯一目标，零变化）；开启双语开关时**追加第二个 sheet「BOM表(英)」**，该 sheet 内每个区块用**中英双行表头**（中文在上、英文在下），区块标题采「中文 (English)」合并行。备选：单 sheet 内表头单元格改为「中文(English)」单行合并文案 |

---

## 2. 产品目标

> **V7 目标：在 V6 已形成的 8 大行业完整闭环之上，做「外向扩展 + 工程巩固」——让 BOM 智造师支持中英双语对照导出（便于跨境/外协协作）、支持三类批量作业（空白模板分发、批量生成、批量合并导入），并对 V2–V6 全量能力做边界健壮性、测试覆盖与性能基线的全面巩固。全部新增能力以「开关/新子命令」形态落地，默认行为与旧调用 100% 向后兼容。**

三个正交子目标：
1. **双语对照并列导出（A）**：新增双语开关，产出中英对照 BOM（推荐为追加双语 sheet），覆盖区块标题、表头列名、行业模板 `material_type` 建议值；翻译字典承载于 `bom_constants.py` 新增 `I18N`（eng 为键）。默认单语导出零变化。
2. **批量模板（B）**：实现 B1 批量空白模板、B2 批量生成 BOM、B3 批量合并导入三种能力；B3 合并策略明确（顺序拼接、industry 取首个非空、不去重、冲突留痕）。
3. **全面巩固（C）**：补齐 V2–V6 边界测试用例（空数据/超长字段/非法 industry/缺失必填/异常 JSON/大 BOM 性能基线）；强化异常输入防御与输出结构一致性；增强 `SKILL.md` 交互引导（双语开关提示、批量命令说明）；关注单文件大 BOM 生成/导入效率。

---

## 3. 用户故事

| # | 角色 | 故事 |
|---|------|------|
| US-A1 | 外贸/外协工程师 | 作为一名需要与海外客户或英文同事协作的工程师，我想导出一份中英对照的 BOM（表头双语），以便对方无需翻译即可读懂物料清单与工序，而不破坏现有中文 BOM 的逆向导入能力。 |
| US-A2 | 多语言制单用户 | 作为一名常做出口单的用户，我想在交互引导里看到 `material_type` 建议值的英文对照，以便更准确地标注物料分类。 |
| US-B1 | 工艺主管 | 作为一名需要向多个车间分发空白录入表的主管，我想一次性产出机械/包装/电子等所有行业的**空白 BOM 模板** xlsx，便于各车间按模板填写后回交。 |
| US-B2 | 数据工程师 | 作为一名手里有几十个产品 JSON 的工程师，我想一条命令**批量生成**对应的 BOM xlsx，而不必逐个手动跑脚本。 |
| US-B3 | 合并同类项的工程师 | 作为一名想把若干分散的 BOM Excel **合并为一个 JSON** 的工程师，我想批量导入多个 Excel 并合并其物料/工序，便于统一核算或重新生成。 |
| US-C1 | QA 工程师 | 作为一名回归测试负责人，我想 V2–V6 的边界用例（空数据、超长字段、非法 industry、缺失必填、异常 JSON、大 BOM）都被自动化覆盖，确保最小变更不引入回归。 |
| US-C2 | 交互用户 | 作为一名普通用户，我想在阶段零/阶段一清晰看到「是否双语导出」「如何批量处理」的提示与命令示例，降低新功能的使用门槛。 |

---

## 4. 需求池

> 优先级说明：P0 = 本次核心（三大块主能力）；P1 = 增强/收口（交互引导、冲突留痕、性能基线）；P2 = 可选增强（明确不新增阻断/软校验）。所有条目均为**增量新增**，未提及的 V6 能力全部沿用不变。

### A 块 · 双语对照并列导出

| 编号 | 标题 | 描述 | 验收标准 | 优先级 |
|------|------|------|----------|--------|
| P0-A1 | **双语开关 + 双语 sheet 渲染** | `generate_bom.py` 新增开关（推荐 `--bilingual` 布尔；备选 `--lang {zh,both}`，默认 `zh`）。开启时：主 sheet「BOM表」保持纯中文（零变化）；**追加第二个 sheet「BOM表(英)」**，整表结构（列数/区块顺序/数据内容）与中文 sheet 完全一致，仅表头双语化（中英双行：中文在上、英文在下；区块标题「中文 (English)」合并行）。 | ① 默认（无开关）仅生成单「BOM表」中文 sheet，行为与 V6 字节级一致；② 开启后生成「BOM表」+「BOM表(英)」两 sheet；③ 双语 sheet 各区块表头为中文行+英文行双行；④ 双语 sheet 数据内容与中文 sheet 一致；⑤ 逆向 `import_bom.py` 仍只读「BOM表」中文 sheet，不受双语 sheet 影响 | P0 |
| P0-A2 | **I18N 翻译字典（bom_constants.py）** | 新增 `I18N` 字典（**eng 为键**，值为中文），覆盖：11 个区块标题、物料区 8 列、工序区 6 列、各行业视图列名、成本列名、`material_type` 各行业建议值、`执行标准` 等标签；并提供反向 `ZH2EN = {v:k for k,v in I18N.items()}` 供中文表头查英文。 | ① `I18N` 含全部区块标题与物料/工序区列名的中英映射；② 含 7 行业 `material_types` 建议值中英映射；③ 纯 Python 标准库，无新依赖；④ 旧调用不引用 `I18N` 时无任何影响 | P0 |
| P1-A1 | **交互层双语引导（SKILL.md）** | 阶段零采集后、汇总确认前，提示用户「是否双语导出（中英对照 sheet）？」；阶段一 `material_type` 下拉在开启双语时附英文对照。仅交互提示，不写新 JSON 结构。 | ① SKILL.md 阶段零/阶段一新增双语开关提示；② 双语开启时 `material_type` 建议值展示「中文 (English)」；③ 不改变 JSON Schema | P1 |

### B 块 · 批量模板（B1/B2/B3 均实现）

| 编号 | 标题 | 描述 | 验收标准 | 优先级 |
|------|------|------|----------|--------|
| P0-B1 | **B1 批量生成空白模板** | `generate_bom.py` 新增子模式（推荐 `--blank-templates` + `--out-dir <dir>`，可选 `--industries 机械,包装,电子,...,通用`，默认全部 8 行业）。每个行业产出一份**空白 BOM 录入模板** xlsx：含正确区块标题、表头、行业专属列（机械 8 列/包装 8 列/电子 14 列/化工 13 列/纺织 8 列/家具 8 列/食品 7 列/通用 8 列）、空物料区与空工序区；无数据、不触发校验。 | ① 一条命令产出 N 个行业空白模板（文件名如 `template_机械.xlsx`）；② 每个模板区块标题/表头/列数与对应行业正式 BOM 一致；③ 空模板不触发 `VALIDATION_FAILED`（无物料、无必填校验）；④ 默认单条生成模式（`--data/--out`）不受影响 | P0 |
| P0-B2 | **B2 批量生成 BOM** | `generate_bom.py` 新增批量输入（推荐 `--batch-dir <dir>` 读取目录下 `*.json`，或 `--batch <f1.json,f2.json,...>`），配合 `--out-dir <dir>`（默认当前目录），逐个生成 BOM xlsx。单个文件失败**不中断**其余，最后汇总成功/失败清单。 | ① 给定目录多个 JSON，逐条生成 BOM xlsx；② 单文件校验失败记录错误并继续，最终打印汇总（成功数/失败数+原因）；③ 输出文件名规则明确（建议 `BOM_<产品名>_<日期>.xlsx` 或 `BOM_<序号>.xlsx`）；④ 单条 `--data/--out` 模式零变化 | P0 |
| P0-B3 | **B3 批量合并导入 → 单一 JSON** | `import_bom.py` 支持多输入（推荐 `--in f1.xlsx f2.xlsx ...`，`nargs="+"`）+ `--merge` 标志 + `--out merged.json`。合并策略见 §6.3。 | ① 多 Excel 逆向解析后合并为单 JSON；② `materials`/`processes` 顺序拼接、不去重、保留原序；③ `industry` 取首个非空文件（兼容单文件 Schema），顶层附 `merged_from` 列表追溯；④ 产品级字段取首个文件；⑤ 单 `--in` 模式零变化 | P0 |
| P1-B1 | **B3 合并冲突留痕** | 合并时若 `step_no` 跨文件重复，不自动重命名（保留原序拼接），在输出顶层 `merge_notes` 记录冲突清单（如「step_no 'S01' 在文件 2/3 重复」）；`materials[].process` 引用保持原值。 | ① 合并 JSON 含 `merge_notes` 字段，列出所有 step_no 冲突；② 无冲突时 `merge_notes=[]`；③ 不擅自改写 step_no（最小变更） | P1 |

### C 块 · 全面巩固

| 编号 | 标题 | 描述 | 验收标准 | 优先级 |
|------|------|------|----------|--------|
| P0-C1 | **边界用例补强（V2–V6）** | 在既有测试套件基础上补充：空数据 JSON、超长字段（名称/单位超长）、非法 industry（V8 回退）、缺失必填（V1/V2/V3 阻断）、循环/异常 JSON（闭合工序链、非法嵌套）、超大 BOM（如 1000+ 物料）性能基线。 | ① 新增用例覆盖上述场景且全绿；② 不修改 V6 既有断言；③ 异常输入均被既有阻断/软校验正确拦截或优雅处理 | P0 |
| P0-C2 | **异常输入防御 + 输出一致性** | `generate_bom.py`/`import_bom.py` 对异常输入（缺标记、空表、坏 JSON、非 BOM 文件）保持既有错误前缀（`VALIDATION_FAILED`/`PARSE_ERROR`/`FILE_ERROR`，退出码 2）且信息明确；正向输出结构（行号/列数/区块顺序）在双语/批量模式下与单语单文件模式一致。 | ① 异常输入不崩溃、有清晰错误前缀；② 双语 sheet 与中文 sheet 结构一致；③ 批量产物与单条产物结构一致 | P0 |
| P1-C1 | **大 BOM 性能基线** | 建立单文件大 BOM（建议 1000 物料）生成/导入的耗时与内存基线（仅作测试关注与回归参照）。注意：**Windows 连跑 100+ subprocess 测试句柄累积崩溃属 QA 分批跑可规避的环境问题，不写入 PRD 作为功能缺陷**；代码侧仅关注单文件大 BOM 的生成耗时与内存。 | ① 测试侧提供大 BOM 性能基线用例（标注耗时/内存阈值，可选）；② 代码侧无内存泄漏式实现（流式/一次性构建，不重复全表拷贝） | P1 |
| P1-C2 | **SKILL.md 批量命令引导增强** | `SKILL.md` 阶段零/正向入口补充：双语开关提示（见 P1-A1）、批量处理说明（B1/B2/B3 的命令示例与适用场景）。仅文档/交互层，不写新 JSON 结构。 | ① SKILL.md 含双语开关与三类批量命令的简明示例；② 与现有分阶段引导风格一致 | P1 |
| P2-C1 | **模板含示例行（可选）** | 作为未来增强：空白模板是否在首行放一条浅色示例物料/工序便于填写引导。**V7 不实现、不新增任何校验**。 | ① 输出方案与取舍建议，供后续版本决策 | P2 |

---

## 5. 双语导出设计（A 块）

### 5.1 推荐形态论证

| 维度 | 推荐：主 sheet 纯中文 + 追加「BOM表(英)」双语 sheet（中英双行表头） | 备选：单 sheet 表头「中文(English)」单行合并 |
|------|------|------|
| 向后兼容 | ✅ 主「BOM表」中文 sheet 字节级不变，`import_bom.py`（`wb["BOM表"]`）零改动 | ⚠️ 表头文案变化，`_map_header` 需靠候选兼容（如「物料名称」vs「物料名称(Material Name)」），有回归风险 |
| 双语对照清晰度 | ✅ 中文行/英文行上下并列，逐项对照最清晰 | ○ 中英挤同格，长表头易溢出 |
| 实现侵入性 | ○ 需重构 `build_workbook` 表头写为可复用 helper（按 lang 渲染两 sheet） | ○ 仅改表头文案一处 |
| 逆向安全 | ✅ 双语 sheet 不参与导入，导入目标恒为中文 sheet | ✅ 但若用户误导入改动后的中文 sheet 仍可解析 |

**结论**：推荐主方案（追加双语 sheet + 中英双行表头）。不破坏任何既有契约，且双语 sheet 仅为人工阅读用途，逆向导入唯一目标恒为纯中文「BOM表」。

### 5.2 需要翻译的文本清单

| 类别 | 文本（中文 → 英文示意） | 承载 |
|------|------|------|
| 区块标题（11） | 一、物料信息→I. Material Information；二、工艺工序→II. Process / Manufacturing Process；三、配料表→III. Ingredients List；三、元件清单→III. Components List；三、配方表→III. Formula List；三、面料辅料清单→III. Fabric & Trims List；三、家具物料清单→III. Furniture BOM List；三、机械物料清单→III. Mechanical BOM List；三、包装物料清单→III. Packaging BOM List；三、成本明细→III. Cost Detail；四、成本明细→IV. Cost Detail | `I18N`（eng 为键） |
| 物料区列名（8） | 序号→No.；物料名称→Material Name；单位→Unit；用量→Qty；出品率(%)→Yield(%)；ERP物料代码→ERP Code；物料类型→Material Type；所属工序→Process | `I18N` |
| 工序区列名（6） | 工序编号→Step No.；工序名称→Step Name；工序说明→Description；工时→Work Hours；备注→Note；产物→Output | `I18N` |
| 各行业视图列名 | 各视图表头（配料表 7 / 元件清单 14 / 配方表 13 / 纺织 8 / 家具 8 / 机械 8 / 包装 8 / 成本 8）按现有中文表头逐条映射英文（如 图号→Drawing No.、材质→Material、热处理→Heat Treatment、克重(g/m²)→Basis Weight(g/m²)、环保标识→Eco Label、单价→Unit Price、币种→Currency、总价→Total 等） | `I18N` |
| 标题/合计标签 | 产品名称→Product Name；产品类别→Category；全产品出品率→Overall Yield；版本号→Version；生成日期→Date；审批人→Approver；生效日期→Effective Date；执行标准→Executive Standard；合计→Total；成本合计→Cost Total | `I18N` |
| `material_type` 建议值（7 行业） | 食品：原料/添加剂/香精香料→Raw/Additive/Flavoring；电子：电阻/电容/IC/连接器/二极管/三极管/晶振/其他→Resistor/Capacitor/IC/Connector/Diode/Transistor/Crystal/Other；化工：主料/溶剂/催化剂/添加剂/包材/其他→Main Agent/Solvent/Catalyst/Additive/Packaging/Other；纺织：面料/辅料/纱线/印染/五金/其他→Fabric/Trim/Yarn/Dyeing/Hardware/Other；家具：主材/板材/辅材/五金/面料/其他→Main Material/Board/Auxiliary/Hardware/Fabric/Other；机械：零部件/标准件/型材/铸件/焊接件/其他→Part/Std Part/Profile/Casting/Welded Assy/Other；包装：纸箱/缓冲/标签/胶带/薄膜/其他→Carton/Cushion/Label/Tape/Film/Other | `I18N`（按行业分组映射） |
| 行业模板 `standard` 执行标准名 | **标准代号为国际通用代码（如 GB/T 1804-2000、GB 7718-2025），不翻译**，双语 sheet 中该单元格保留原代号；仅标签「执行标准」翻译为 Executive Standard。是否需要英文说明（如 "General Tolerances"）列为待确认 Q3，默认不翻。 | `I18N`（仅标签） |

### 5.3 I18N 字典承载（bom_constants.py 增量）

```python
# === V7 双语导出 I18N 字典（eng 为键，值为中文） ===
I18N = {
    # 区块标题
    "I. Material Information": "一、物料信息",
    "II. Process / Manufacturing Process": "二、工艺工序",
    "III. Ingredients List": "三、配料表",
    "III. Components List": "三、元件清单",
    "III. Formula List": "三、配方表",
    "III. Fabric & Trims List": "三、面料辅料清单",
    "III. Furniture BOM List": "三、家具物料清单",
    "III. Mechanical BOM List": "三、机械物料清单",
    "III. Packaging BOM List": "三、包装物料清单",
    "III. Cost Detail": "三、成本明细",
    "IV. Cost Detail": "四、成本明细",
    # 物料区列名
    "No.": "序号", "Material Name": "物料名称", "Unit": "单位",
    "Qty": "用量", "Yield(%)": "出品率(%)", "ERP Code": "ERP物料代码",
    "Material Type": "物料类型", "Process": "所属工序",
    # 工序区列名
    "Step No.": "工序编号", "Step Name": "工序名称",
    "Description": "工序说明", "Work Hours": "工时",
    "Note": "备注", "Output": "产物",
    # 标题/合计标签
    "Product Name": "产品名称", "Category": "产品类别",
    "Overall Yield": "全产品出品率", "Version": "版本号",
    "Date": "生成日期", "Approver": "审批人",
    "Effective Date": "生效日期", "Executive Standard": "执行标准",
    "Total": "合计", "Cost Total": "成本合计",
    # material_type 建议值（节选，7 行业全量见设计文档）
    "Raw": "原料", "Additive": "添加剂", "Flavoring": "香精香料",
    "Resistor": "电阻", "Capacitor": "电容", "IC": "IC",
    "Connector": "连接器", "Diode": "二极管", "Transistor": "三极管",
    "Crystal": "晶振", "Other": "其他",
    # ……其余视图列名与行业建议值按 §5.2 全量补齐
}
# 反向映射：中文表头 → 英文（渲染双语 sheet 时由中文查英文）
ZH2EN = {v: k for k, v in I18N.items()}
```

> 说明：`generate_bom.py` 抽取一个 `_write_block(ws, r, marker_zh, headers_zh, rows, lang)` helper，渲染中文 sheet 时用 `headers_zh`；渲染双语 sheet 时 `marker` 写为 `f"{marker_zh} ({ZH2EN.get(marker_zh,'')})"`，表头写为双行（`headers_zh` 在上、`[ZH2EN.get(h,'') for h in headers_zh]` 在下）。数据行两 sheet 完全一致。

---

## 6. 批量模板设计（B 块）

### 6.1 B1 批量空白模板（generate_bom.py）

- 触发：`--blank-templates` + `--out-dir <dir>`（必填输出目录）；`--industries` 可选逗号列表，默认全部 8 行业（机械/包装/电子/化工/纺织/家具/食品/通用）。
- 每个模板：调用与正式 BOM 同构的 `build_workbook` 空数据渲染（`materials=[]`、`processes=[]`），仅写区块标题 + 表头 + 空物料区（含「合计」占位行样式一致）+ 空工序区。
- **不触发 `validate()`**（无物料、无必填），避免 `VALIDATION_FAILED`；但需保证空模板结构合法、可被 `import_bom.py` 解析为空 JSON。
- 文件名：`template_<行业>.xlsx`（如 `template_机械.xlsx`、`template_通用.xlsx`）。

### 6.2 B2 批量生成 BOM（generate_bom.py）

- 触发：`--batch-dir <dir>`（读取 `*.json`）或 `--batch <f1.json,f2.json,...>` + `--out-dir <dir>`（默认当前目录）。
- 流程：逐个读取 JSON → `validate()`；通过则 `build_workbook` + 保存；失败则记录 `(文件, 错误列表)` 并继续下一个。
- 错误隔离：单文件失败**不中断**整批；结束时打印汇总：`成功 N 个 / 失败 M 个`，失败项附 `VALIDATION_FAILED` 原因。
- 输出命名：`BOM_<产品名>_<日期>.xlsx`；产品名含非法文件名字符时回落为 `BOM_<序号>.xlsx`。
- 向后兼容：单条 `--data/--out` 模式零变化。

### 6.3 B3 批量合并导入（import_bom.py）— 合并策略（明确）

- 触发：`--in f1.xlsx f2.xlsx ...`（`nargs="+"`，≥2）+ `--merge` + `--out merged.json`。
- **合并规则（推荐，最小变更）**：
  1. **materials**：顺序拼接各文件 `parse_bom` 得到的 `materials` 列表（按文件输入顺序），**不去重、保留原序**。
  2. **processes**：顺序拼接各文件 `processes` 列表（保留原序）；`step_no` 跨文件重复时**不自动重命名**，仅在 `merge_notes` 记录冲突（见 P1-B1）。
  3. **industry**：取**首个非空文件**的 `industry`（兼容单文件 Schema，便于直接回灌正向流程）；顶层附 `merged_from = [各文件 industry]` 列表供追溯。
  4. **产品级字段**（product_name/category/output_rate/version/date/approver/effective_date/standard）：取首个文件的值（合并多产品本无统一语义），`merge_notes` 注明「产品级字段采用首个文件」。
  5. **结构冲突处理**：若某文件解析失败（`PARSE_ERROR`/`FILE_ERROR`），跳过该文件并在 `merge_notes` 记录「文件 X 解析失败：原因」，不中断其余。
- 输出 JSON 顶层结构：`{ ...产品级字段, industry, merged_from, materials, processes, merge_notes }`。
- 向后兼容：单 `--in` + `--out` 模式零变化（输出不含 `merged_from`/`merge_notes`）。

---

## 7. 全面巩固设计（C 块）

- **测试补强（P0-C1）**：在 `tests/` 既有套件上新增 V7 增量测试文件（如 `test_bom_v7.py`），覆盖：空数据、超长字段、非法 industry（V8 回退）、缺失必填（V1/V2/V3 阻断）、循环/异常 JSON、1000+ 物料大 BOM 性能基线；双语 sheet 结构一致性、批量产物结构一致性。不修改 V6 既有断言。
- **异常防御（P0-C2）**：确认 `generate_bom.py`/`import_bom.py` 对异常输入保持既有错误前缀与退出码（2），信息明确；不引入新异常类型。
- **性能（P1-C1）**：仅建单文件大 BOM 基线（耗时/内存），代码侧避免全表重复拷贝；**测试环境 Windows 句柄累积崩溃属 QA 分批跑规避项，不列为功能缺陷**。
- **交互（P1-C2 / P1-A1）**：`SKILL.md` 阶段零/阶段一补充双语开关提示与三类批量命令示例，风格与现有分阶段引导一致。

---

## 8. 向后兼容与约束提醒

| 约束 | V7 处理 |
|------|---------|
| 物料区 8 列（A–H）永不变 | ✅ 沿用 V6，双语/批量模板均复用同构空物料区 |
| 旧 JSON/Excel 零结构变化 | ✅ 默认单语导出、单条生成、单文件导入行为与 V6 字节级一致 |
| 不引入新第三方依赖 | ✅ 仅 openpyxl；I18N 为纯 Python 字典 |
| **不新增任何阻断/软校验** | ✅ V7 不新增任何 V/W 类校验；双语/批量均为形态扩展，不改变既有校验集 |
| `industry` 枚举无需扩展 | ✅ 沿用 8 值 |
| 旧调用向后兼容 | ✅ 新增开关/子命令均为**可选项**；不传则完全走 V6 路径 |
| 最小变更 | ✅ 复用 `build_workbook`/`parse_bom`/`derive_*`/`INDUSTRY_TEMPLATES`；B3 合并仅顺序拼接+留痕 |
| 性能问题定位 | ⚠️ Windows 连跑 100+ subprocess 测试句柄崩溃 = QA 分批跑规避，**非功能缺陷**，不写入代码需求 |

---

## 9. UI 设计稿（ASCII）

### 9.1 双语表头布局（追加「BOM表(英)」sheet，中英双行）

```
【BOM表(英) sheet】（结构同中文 sheet，仅表头双语）
行1   ┌─────────────────────────────────────────────────────────┐
      │ BOM Table (BOM表)                          ← 标题
行2   版本号：V1.0              生成日期：2026-07-15
行3   产品名称：减速机总成
行4   产品类别：工业品          全产品出品率：100.0%
行5   审批人：…   生效日期：…   执行标准：GB/T 1804-2000
行6   ┌─────────────────────────────────────────────────────────┐
      │ I. Material Information (一、物料信息)   ← 区块标题(中(英))
行7   序号      物料名称      单位  用量  出品率(%)  ERP物料代码  物料类型  所属工序   ← 中文行
      No.       Material Name  Unit  Qty   Yield(%)   ERP Code    Type      Process    ← 英文行(双语对照)
行8   1         主轴箱体      kg    1     100.0%    RM-001      型材       S01
      ...
（工序区 / 三、机械物料清单 等区块同构：中文表头行 + 英文表头行 + 数据行）
```

> 主「BOM表」中文 sheet 布局与 V6 完全相同（见 V6 PRD §9），此处不重复。

### 9.2 批量命令交互示意

```
# B1 批量空白模板（分发填写）
python generate_bom.py --blank-templates --out-dir ./templates \
       [--industries 机械,包装,电子,化工,纺织,家具,食品,通用]
  → 产出 template_机械.xlsx / template_包装.xlsx / ... / template_通用.xlsx

# B2 批量生成 BOM（逐条跑，单失败不中断）
python generate_bom.py --batch-dir ./inputs --out-dir ./outputs
  → 逐个读取 *.json，产出 BOM_<产品名>_<日期>.xlsx
  → 结束打印：成功 N / 失败 M（附 VALIDATION_FAILED 原因）

# B3 批量合并导入（多 Excel → 单 JSON）
python import_bom.py --in a.xlsx b.xlsx c.xlsx --merge --out merged.json
  → materials/processes 顺序拼接(不去重)
  → industry 取首个非空；merged_from=[...]；merge_notes=[冲突/失败记录]
```

---

## 10. 待确认问题

| # | 问题 | 推荐默认 | 理由 |
|---|------|----------|------|
| Q1 | 双语具体形态：追加「BOM表(英)」双语 sheet（中英双行）vs 单 sheet 表头「中文(English)」合并？ | **追加双语 sheet（中英双行表头）** | 主中文 sheet 字节级不变，`import_bom.py` 零改动；双语 sheet 仅为人工阅读，逆向唯一目标恒为中文 sheet，回归风险最低 |
| Q2 | 双语开关命名：`--bilingual`（布尔）vs `--lang {zh,both}`（含 `en`）？若支持 `--lang en` 单独英文 sheet，须保证逆向目标仍为中文 sheet | **`--bilingual` 布尔**（备选 `--lang {zh,both}`，不支持独立 `en` 替换主 sheet） | 避免英文主 sheet 破坏逆向；开关语义最直白 |
| Q3 | 行业模板 `standard` 执行标准代号是否翻译（如 GB/T 1804-2000 → 加注 "General Tolerances"）？ | **不翻代号，保留原码**；仅标签「执行标准」翻译 | 标准代号为国际通用代码，翻译易歧义；如需英文说明列为 P2 未来增强 |
| Q4 | B3 合并时 `step_no` 跨文件冲突：默认保留原序+`merge_notes` 留痕 vs 按文件前缀自动重命名（如 S01→P2-S01）？ | **保留原序 + merge_notes 留痕**（不重命名） | 最小变更、不擅自改写用户编号；留痕供人工决断 |
| Q5 | B3 `industry` 取值：取首个非空文件 vs 顶层记为列表？ | **顶层 industry 取首个非空 + merged_from 列表追溯** | 兼容单文件 Schema，可直接回灌正向；列表保留多源信息 |
| Q6 | B2/B1 输出文件命名与 `--out-dir` 默认路径（当前目录 vs 输入同目录）？ | **`--out-dir` 默认当前目录；B2 命名 `BOM_<产品名>_<日期>.xlsx`，B1 命名 `template_<行业>.xlsx`** | 命名可预期、便于分发与归档 |
| Q7 | B2 单文件失败策略：中断全部 vs 记录错误继续？ | **记录错误继续，结束汇总成功/失败** | 批量场景容错优先，不因单文件阻断整批 |
| Q8 | 大 BOM 性能基线是否需在 PRD 量化阈值（如 1000 物料 ≤X 秒/≤Y MB）？ | **仅作测试关注与回归参照，不在 PRD 硬编码阈值** | 阈值随环境浮动；代码侧关注单文件大 BOM 生成耗时与内存，不把测试环境问题当功能缺陷 |

> 说明：以上 Q1–Q8 均可在不破坏「向后兼容/最小变更/不新增阻断软校验」前提下拍板；任一接受项仅影响实现细节，不改变本 PRD 三大块主能力范围。

---

## 11. 优先级分布

| 优先级 | 数量 | 条目 |
|--------|------|------|
| P0 | 7 | 双语开关+双语 sheet(P0-A1)、I18N 字典(P0-A2)、B1 批量空白模板(P0-B1)、B2 批量生成(P0-B2)、B3 批量合并导入(P0-B3)、边界用例补强(P0-C1)、异常防御+输出一致性(P0-C2) |
| P1 | 4 | 交互层双语引导(P1-A1)、B3 冲突留痕(P1-B1)、大 BOM 性能基线(P1-C1)、SKILL.md 批量命令引导(P1-C2) |
| P2 | 1 | 模板含示例行(P2-C1，可选不实现) |

> 合计 12 条增量需求，全部为 V7 新增；V6 契约（8 列物料区、7 行业视图+机械/包装、成本双编号、37 唯一键、软校验集、逆向闭环）100% 沿用不变。新增 `I18N` 字典（eng 为键）承载翻译；新增 CLI 开关/子命令（`--bilingual` / `--blank-templates` / `--batch-dir` / `--batch` / `--out-dir` / `--merge` 多 `--in`）均为可选，默认行为零变化。
