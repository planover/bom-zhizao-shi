# BOM智造师（BOM Maker）使用说明

## 1. 技能简介

**一句话**：把产品的「物料信息」和「工艺工序信息」交互式采集、校验，一键生成标准 BOM 表 Excel；也支持把已有 BOM 表 Excel 逆向解析回结构化 JSON，方便重新编辑。

**适用场景**：
- 制造、工艺、采购、成本核算等需要结构化 BOM（物料清单）的场合；
- 需要把物料（名称 / 单位 / 用量 / 出品率 / 物料类型 / 所属工序）与工序（编号 / 名称 / 说明 / 工时 / 备注 / 产物）整理成标准 Excel；
- 食品类产品可自动生成仅含食用物料的「配料表」；
- 已有一份本技能导出的 BOM 表 .xlsx，想反查其内容或在其基础上「重新编辑」。

> **V2 新增**：产品类别（5 类枚举）、全产品出品率（>0 可>100）、物料类型与所属工序、工序产物与流转链校验、食品类自动配料表、Excel 7 列与分工序分组呈现。

## 2. 安装 / 位置说明

- 本技能已随 WorkBuddy **用户级 Skill** 生效，无需额外安装；技能根目录位于
  `~/.workbuddy/skills/bom-zhizao-shi`（Windows：
  `C:\Users\<用户名>\.workbuddy\skills\bom-zhizao-shi`）。
- 生成 / 导入脚本依赖 [`openpyxl`](https://pypi.org/project/openpyxl/)，脚本在运行时会**自动安装**（缺失则执行 `pip install openpyxl`），无需手动配置 Python 环境。

## 3. 触发方式

| 方向 | 命中关键词 |
|------|-----------|
| 正向（生成 BOM） | `BOM生成`、`物料清单制作`、`BOM表`、`物料清单`、`工艺BOM`、`生成BOM` |
| 逆向（从 Excel 导入 BOM） | `BOM导入`、`从Excel导入BOM`、`反向导入BOM`、`导入BOM表`、`BOM表导入` |

也可在对话中以自然语言表达意图，例如「帮我生成一份 BOM 表」「把这份 BOM 表 Excel 解析成数据」。

## 4. 正向流程使用步骤

```
触发 → 阶段零(产品信息) → 阶段一(物料采集) → 阶段二(工序采集) → 数据校验 → 汇总确认 → 生成Excel → 输出
                 ↑ 任意阶段可取消(重新输入 / 退出)
```

1. **阶段零 · 产品信息**：采集产品名称（**必填**）、产品类别（**必填**，下拉 5 类）、全产品出品率（**必填**，>0 可>100）、版本号（默认 V1.0）、生成日期（默认当天）。
2. **阶段一 · 物料采集**：逐条录入物料，每条 7 个字段（物料名称、计量单位、用量、出品率(%)、ERP物料代码（可选）、物料类型（可选下拉）、所属工序（可选引用 step_no））。
   - 支持逐字段回复，也支持按格式整段回复；支持多条批量录入。
   - 每条校验通过后才存入临时列表；可随时「继续添加 / 结束」。
3. **阶段二 · 工序采集**：逐条录入工序，每条 6 个字段（工序编号、工序名称、工序说明、工时、备注、**产物 output（必填）**）。
   - 工序编号可自动建议（上一条 +1）；允许 0 道工序（纯物料 BOM）。
   - 多道工序时，下一道工序的物料清单须包含上一道工序的「产物」（名称精确匹配），构成流转链（R3）。
4. **数据校验**：用量必须为正数、物料出品率须为 0–100 的正数、工时数值须 ≥0、工序编号不可重复、产品名称/类别/出品率必填、流转链完整等，逐项即时校验。
5. **汇总确认**：输出结构化汇总（含配料表预览（食品类）），确认后生成。
6. **生成 Excel**：调用 `scripts/generate_bom.py` 生成 .xlsx，返回文件绝对路径。

**交互示例**：
```
用户：帮我生成一份 BOM 表
BOM智造师：请输入产品名称（如：芒果果味糖浆）：
用户：芒果果味糖浆
BOM智造师：请选择产品类别：食品 / 工业品 / 日化化妆品 / 医药 / 其他
用户：食品
BOM智造师：请输入全产品出品率(%)（必填，>0，可>100）：
用户：130
...（采集物料、工序）...
BOM智造师：确认生成？(y / 修改 / 退出)
用户：y
BOM智造师：已生成：/path/to/BOM_2026-07-07.xlsx
```

## 5. 逆向流程使用步骤

```
触发(导入类关键词) → 接收Excel路径 → 逆向解析(import_bom.py) → 展示/导出JSON → [重新生成Excel / 退出]
```

1. 用导入类关键词触发（如「把这份 BOM 表导入」），提供 .xlsx 文件路径。
2. 调用逆向解析脚本（按列头文本定位，向后兼容旧版 5 列 Excel）：
   ```
   python3 <skill_dir>/scripts/import_bom.py --in BOM_2026-07-07.xlsx [--out data.json]
   ```
3. 查看结构化汇总（版本号、日期、物料条数、工序条数及关键字段）。
4. 提供后续选项：
   ```
   [1] 重新生成 Excel（把这份 JSON 作为 --data 传回 generate_bom.py，可先编辑）
   [2] 导出 JSON 文件（方便二次处理）
   [3] 退出
   ```
   - 选 1：进入正向后处理，可先编辑再重新生成，实现「导入 → 编辑 → 重新生成」闭环；
   - 选 2：执行 `import_bom.py --in <xlsx> --out <data.json>`，把文件落盘；
   - 选 3：终止流程。

> 逆向导入是「重新编辑已有 BOM」的入口：导出的 JSON 可直接重新喂给正向流程。**注意**：「三、配料表」为派生区块，逆向不解析、不回写；重新生成时按 `category` 重新派生。

## 6. 数据校验规则速查表

| 字段 | 必填 | 校验规则 |
|------|------|----------|
| 产品名称 | **是** | 非空（R1）；留空触发 VALIDATION_FAILED |
| 产品类别 | **是** | ∈ {食品,工业品,日化化妆品,医药,其他}（R1/R4） |
| 全产品出品率(output_rate) | **是** | 数值 `> 0`，允许 `> 100`（无硬上限）（R2） |
| 物料名称 | 是 | 非空 |
| 计量单位 | 是 | 非空 |
| 用量 | 是 | 正数（>0），数值 |
| 出品率(%)（物料级 yield_rate） | 是 | 正数（>0）且 ≤100，数值 |
| ERP物料代码 | 否 | 无，可留空 |
| 物料类型 | 否 | ∈ {原料,添加剂,香精香料,包材,其他}，默认 `其他` |
| 所属工序 | 否 | 引用有效工序 `step_no`，首道可选填，默认空 |
| 工序编号 | 是 | 非空，不可重复 |
| 工序名称 | 是 | 非空 |
| 工序说明 | 否 | 文本 |
| 工时 | 否 | 若填数值须 ≥0 |
| 备注 | 否 | 文本 |
| 产物(output) | **是** | 非空；须为下一道工序物料清单中的一条（R3 流转链） |

> **R3 流转链校验（阻断级）**：仅当工序数 ≥2 时触发。上一工序 `output` 必填；下一道工序的物料清单（按 `process == 该工序.step_no` 过滤）必须包含上一工序 `output` 作为一条物料（名称精确匹配），否则生成失败（退出码 2）并提示「流转链不完整」。

## 7. Excel 结构说明

生成的 BOM 表为单工作表（标题 `BOM表`，**7 列 A–G**），含：合并标题行、版本号/生成日期行、**产品名称行（必填，整行）**、产品类别/全产品出品率行、「一、物料信息」物料区（7 列表头：物料名称/单位/用量/出品率(%)/ERP物料代码/物料类型/所属工序；有工序且存在归属时分工序分组，含 `【工序 Sxx 名称】` 与 `【未归属工序】` 子标题）、「二、工艺工序」工序区（含产物列）、「三、配料表」区块（**仅食品类**，仅含 原料/添加剂/香精香料，按用量降序）；出品率/全产品出品率单元格数字格式 `0.0"%"`。

演示图（还原 Excel 视觉，无 GUI 环境生成）：

![BOM 表演示](references/bom-demo.svg)

完整的「输入 JSON Schema」「Excel 输出结构」「逆向导入解析规则」见 **[`references/bom-spec.md`](references/bom-spec.md)**。

## 8. 已知限制（旧版数据回灌）

- 旧版 5 列 Excel / 旧版 5 字段 JSON **逆向导入不会报错**（缺列自动取默认），但解析得到的 `category` / `output_rate` / `processes[].output` 为空。
- 将这类旧数据重新喂给 `generate_bom.py` 时，会触发 V2 校验（V1/V2/V3/V5），打印 `VALIDATION_FAILED` 并以退出码 2 结束——这是 R5「待补」的**固有结果而非缺陷**，需用户补全 `product_name` / `category` / `output_rate` / 工序 `output` 后方可重新生成。
- 食品类若未给物料标注 `material_type`，配料表无法正确过滤（默认 `其他` 会被排除）。

## 9. 目录文件清单

```
bom-zhizao-shi/
├── SKILL.md                  # 技能定义、触发与正/逆向流程说明
├── README.md                 # 本使用说明
├── CHANGELOG.md              # 版本变更记录
├── scripts/
│   ├── generate_bom.py       # 正向：JSON -> BOM 表 Excel（V2，7 列 + 配料表）
│   └── import_bom.py         # 逆向：BOM 表 Excel -> JSON（V2，按列头文本定位）
├── references/
│   ├── bom-spec.md           # BOM 表 Excel 结构与输入 JSON Schema 规范
│   └── bom-demo.svg          # BOM 表 Excel 视觉演示图
└── examples/
    ├── sample_bom_v2.json    # 合法样例（食品类，多工序，触发配料表 + 流转链）
    └── sample_bom_v2.xlsx    # 由样例生成的演示 Excel
```

## 10. 本地命令行直接调用示例

> 将 `<skill_dir>` 替换为技能实际根目录，例如
> `~/.workbuddy/skills/bom-zhizao-shi`（Windows：`C:\Users\<用户名>\.workbuddy\skills\bom-zhizao-shi`）。

**正向：生成 BOM 表**
```bash
python3 <skill_dir>/scripts/generate_bom.py --data bom.json --out BOM_2026-07-07.xlsx
# 成功输出：OK:BOM_2026-07-07.xlsx
# 数据非法输出：VALIDATION_FAILED + 错误列表（退出码 2）
```

`bom.json` 示例（食品类，S01→S02 流转链成立，触发配料表）：
```json
{
  "product_name": "芒果果味糖浆",
  "category": "食品",
  "output_rate": 130,
  "version": "V1.0",
  "date": "2026-07-07",
  "materials": [
    {"name":"芒果原浆","unit":"kg","usage":46.3,"yield_rate":55,"erp_code":"RM-001","material_type":"原料","process":"S01"},
    {"name":"白砂糖","unit":"kg","usage":30.0,"yield_rate":100,"erp_code":"RM-002","material_type":"原料","process":"S01"},
    {"name":"芒果果味糖浆基料","unit":"kg","usage":70.0,"yield_rate":98,"erp_code":"RM-100","material_type":"原料","process":"S02"},
    {"name":"柠檬酸","unit":"kg","usage":0.5,"yield_rate":100,"erp_code":"RM-003","material_type":"添加剂","process":"S02"},
    {"name":"PE 瓶","unit":"个","usage":100,"yield_rate":100,"erp_code":"PK-001","material_type":"包材","process":""}
  ],
  "processes": [
    {"step_no":"S01","name":"调配","desc":"混合搅拌","work_hours":30,"note":"常温","output":"芒果果味糖浆基料"},
    {"step_no":"S02","name":"灌装","desc":"无菌灌装","work_hours":20,"note":"","output":"芒果果味糖浆"}
  ]
}
```
说明：S01 产物 = 芒果果味糖浆基料；S02 物料含「芒果果味糖浆基料」→ 流转链成立。配料表（食品）收录 芒果原浆/白砂糖/芒果果味糖浆基料/柠檬酸（按用量降序：70.0→46.3→30.0→0.5），PE 瓶（包材）排除。

**逆向：导入 BOM 表**
```bash
# 解析并打印 JSON
python3 <skill_dir>/scripts/import_bom.py --in BOM_2026-07-07.xlsx
# 解析并导出 JSON 文件
python3 <skill_dir>/scripts/import_bom.py --in BOM_2026-07-07.xlsx --out bom_back.json
# 成功输出：OK:bom_back.json
# 非 BOM 格式 / 标记缺失输出：PARSE_ERROR（退出码 2）
```

**闭环：逆向结果重新生成 Excel**
```bash
python3 <skill_dir>/scripts/import_bom.py --in BOM_2026-07-07.xlsx --out bom_back.json
python3 <skill_dir>/scripts/generate_bom.py --data bom_back.json --out BOM_v2.xlsx
```

## 11. 仓库地址

GitHub：<https://github.com/planover/bom-zhizao-shi>

变更记录详见 **[`CHANGELOG.md`](CHANGELOG.md)**。
