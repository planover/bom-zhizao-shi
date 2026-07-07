# BOM智造师（BOM Maker）使用说明

## 1. 技能简介

**一句话**：把产品的「物料信息」和「工艺工序信息」交互式采集、校验，一键生成标准 BOM 表 Excel；也支持把已有 BOM 表 Excel 逆向解析回结构化 JSON，方便重新编辑。

**适用场景**：
- 制造、工艺、采购、成本核算等需要结构化 BOM（物料清单）的场合；
- 需要把物料（名称 / 单位 / 用量 / 出品率）与工序（编号 / 名称 / 说明 / 工时 / 备注）整理成标准 Excel；
- 已有一份本技能导出的 BOM 表 .xlsx，想反查其内容或在其基础上「重新编辑」。

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
触发 → 阶段一(物料采集) → 阶段二(工序采集) → 数据校验 → 汇总确认 → 生成Excel → 输出
                 ↑ 任意阶段可取消(重新输入 / 退出)
```

1. **阶段一 · 物料采集**：逐条录入物料，每条 4 个字段（物料名称、计量单位、用量、出品率(%)）。
   - 支持逐字段回复，也支持按格式整段回复；支持多条批量录入。
   - 每条校验通过后才存入临时列表；可随时「继续添加 / 结束」。
2. **阶段二 · 工序采集**：逐条录入工序，每条 5 个字段（工序编号、工序名称、工序说明、工时、备注）。
   - 工序编号可自动建议（上一条 +1）；允许 0 道工序（纯物料 BOM）。
3. **数据校验**：用量必须为正数、出品率须为 0–100 的正数、工时数值须 ≥0、工序编号不可重复等，逐项即时校验。
4. **汇总确认**：输出结构化汇总，确认后生成。
   ```
   ===== BOM 信息汇总（请确认）=====
   版本号：V1.0   生成日期：2026-07-07
   【物料】（共 N 条）
    1. 名称 / 单位 / 用量 / 出品率
   【工序】（共 M 条）
    1. 编号 / 名称 / 说明 / 工时 / 备注
   =================================
   确认生成？(y 生成 / 修改 重新输入 / 退出)
   ```
5. **生成 Excel**：调用 `scripts/generate_bom.py` 生成 .xlsx，返回文件绝对路径。

**交互示例**：
```
用户：帮我生成一份 BOM 表
BOM智造师：请填写一条物料（可整段或逐字段回复）：
  物料名称：钢板
  计量单位：kg
  用量：2.5
  出品率(%)：98
用户：继续
...（采集若干条后）...
BOM智造师：确认生成？(y / 修改 / 退出)
用户：y
BOM智造师：已生成：/path/to/BOM_2026-07-07.xlsx
```

## 5. 逆向流程使用步骤

```
触发(导入类关键词) → 接收Excel路径 → 逆向解析(import_bom.py) → 展示/导出JSON → [重新生成Excel / 退出]
```

1. 用导入类关键词触发（如「把这份 BOM 表导入」），提供 .xlsx 文件路径。
2. 调用逆向解析脚本：
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

**命令示例**：
```
# 解析并直接打印 JSON
python3 ~/.workbuddy/skills/bom-zhizao-shi/scripts/import_bom.py --in BOM_2026-07-07.xlsx

# 解析并导出 JSON 文件
python3 ~/.workbuddy/skills/bom-zhizao-shi/scripts/import_bom.py --in BOM_2026-07-07.xlsx --out bom_back.json
```

> 逆向导入是「重新编辑已有 BOM」的入口：导出的 JSON 可直接重新喂给正向流程。

## 6. 数据校验规则速查表

| 字段 | 必填 | 校验规则 |
|------|------|----------|
| 物料名称 | 是 | 非空 |
| 计量单位 | 是 | 非空 |
| 用量 | 是 | 正数（>0），数值 |
| 出品率(%) | 是 | 正数（>0）且 ≤100，数值 |
| 工序编号 | 是 | 非空，不可重复 |
| 工序名称 | 是 | 非空 |
| 工序说明 | 否 | 文本 |
| 工时 | 否 | 若填数值须 ≥0 |
| 备注 | 否 | 文本 |

## 7. Excel 结构说明

生成的 BOM 表为单工作表（标题 `BOM表`，5 列 A–E），含：合并标题行、版本号 / 生成日期行、「一、物料信息」物料区、「二、工艺工序」工序区、相应表头与数据行；出品率单元格数字格式 `0.0"%"`。

完整的「输入 JSON Schema」与「Excel 输出结构」规范见 **[`references/bom-spec.md`](references/bom-spec.md)**，逆向导入解析规则亦在该文档的「逆向导入（import_bom.py）」节说明。

## 8. 目录文件清单

```
bom-zhizao-shi/
├── SKILL.md                  # 技能定义、触发与正/逆向流程说明
├── README.md                 # 本使用说明
├── scripts/
│   ├── generate_bom.py       # 正向：JSON -> BOM 表 Excel
│   └── import_bom.py         # 逆向：BOM 表 Excel -> JSON
└── references/
    └── bom-spec.md           # BOM 表 Excel 结构与输入 JSON Schema 规范
```

## 9. 本地命令行直接调用示例

> 将 `<skill_dir>` 替换为技能实际根目录，例如
> `~/.workbuddy/skills/bom-zhizao-shi`（Windows：`C:\Users\<用户名>\.workbuddy\skills\bom-zhizao-shi`）。

**正向：生成 BOM 表**
```bash
python3 <skill_dir>/scripts/generate_bom.py --data bom.json --out BOM_2026-07-07.xlsx
# 成功输出：OK:BOM_2026-07-07.xlsx
# 数据非法输出：VALIDATION_FAILED + 错误列表（退出码 2）
```

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

## 10. 仓库地址

GitHub：<https://github.com/planover/bom-zhizao-shi>
