# BOM 表结构与输入规范（bom-zhizao-shi）

本文件供 `scripts/generate_bom.py` 与汇总序列化阶段参考。

## 输入 JSON Schema

`generate_bom.py --data <file.json>` 读取的 JSON 结构：

```json
{
  "version": "V1.0",
  "date": "2026-07-07",
  "materials": [
    {
      "name": "钢板",
      "unit": "kg",
      "usage": 2.5,
      "yield_rate": 98
    }
  ],
  "processes": [
    {
      "step_no": "S01",
      "name": "切割",
      "desc": "激光切割下料",
      "work_hours": 15,
      "note": "注意防护"
    }
  ]
}
```

字段约束：

| 字段 | 类型 | 约束 |
|------|------|------|
| version | string | 可选，默认 V1.0 |
| date | string | 可选，默认当天 YYYY-MM-DD |
| materials[].name | string | 必填，非空 |
| materials[].unit | string | 必填，非空 |
| materials[].usage | number | 必填，正数（>0） |
| materials[].yield_rate | number | 必填，0 < 值 ≤ 100 |
| processes[].step_no | string | 必填，唯一不重复 |
| processes[].name | string | 必填，非空 |
| processes[].desc | string | 可选 |
| processes[].work_hours | number/string | 可选，数值须 ≥ 0 |
| processes[].note | string | 可选 |

## Excel 输出结构

单工作表 `BOM表`，5 列（A–E）：

```
┌───────────────────────────────────────────────────────┐
│                     BOM表（标题，A1:E1 合并）            │
├───────────────┬───────────────┬──────────┬─────────────┤
│ 版本号：V1.0（A2:C2）      │ 生成日期：2026-07-07（D2:E2）│
├───────────────┴───────────────┴──────────┴─────────────┤
│ 一、物料信息                                              │
├──────────┬───────┬───────┬──────────┬──────────────────┤
│ 物料名称  │ 单位  │ 用量  │ 出品率(%) │ （留空）          │
├──────────┼───────┼───────┼──────────┼──────────────────┤
│ ...物料行  │       │       │          │                  │
├──────────┴───────┴───────┴──────────┴──────────────────┤
│ 二、工艺工序                                            │
├──────────┬───────┬───────┬───────┬──────────────────┤
│ 工序编号  │ 工序名称│ 工序说明│ 工时  │ 备注              │
├──────────┼───────┼───────┼───────┼──────────────────┤
│ ...工序行  │       │       │       │                  │
└──────────┴───────┴───────┴───────┴──────────────────┘
```

样式约定：标题 16pt 深蓝加粗居中；区标题 10pt 加粗；表头蓝底加粗居中并带边框；数据单元格带边框，文本列左对齐、数值列居中；出品率单元格数字格式 `0.0"%"`。

## 逆向导入（import_bom.py）

`scripts/import_bom.py` 读取由本规范生成的 BOM 表 Excel，反向解析为与上方「输入 JSON Schema」**完全一致**的结构化 JSON，便于「重新编辑已有 BOM」或闭环回写。

### CLI

```
python3 import_bom.py --in <BOM.xlsx> [--out <data.json>]
```

- `--in`：必填，待解析的 BOM 表 Excel 路径。
- `--out`：可选，指定后将 JSON 写入该路径（utf-8、ensure_ascii=False、indent=2），stdout 打印 `OK:<json路径>`；不指定则 stdout 直接打印 `OK:<json字符串>`。

### 解析规则

1. **工作表**：优先取 `title == "BOM表"` 的工作表，否则取 active sheet。
2. **标题**：读取 `A1` 合并单元格，期望为 `BOM表`；若不符仅告警（WARNING）并继续解析。
3. **版本号**：扫描包含 `版本号` 的单元格（即 `A2:C2` 合并区左上角 `A2`），提取冒号（兼容 `：`/`:`）后内容；取不到则默认 `V1.0`。
4. **生成日期**：扫描包含 `生成日期` 的单元格（即 `D2:E2` 合并区左上角 `D2`），提取冒号后内容；取不到则默认空字符串 `""`。
5. **物料区**：定位含 `一、物料信息` 的行，其下一行为表头（忽略），再往下的数据行直到遇到含 `二、工艺工序` 的行或空行（物料名称列空即止）；每行取 col1=name、col2=unit、col3=usage(float)、col4=yield_rate(float)；**完全空行跳过**。
6. **工序区**：定位含 `二、工艺工序` 的行，其下一行为表头（忽略），再往下的数据行直到空行（工序编号为空即止）；每行取 col1=step_no、col2=name、col3=desc、col4=work_hours(float 或保留原值/空)、col5=note；**完全空行跳过**。
7. **数值转换**：`usage` / `yield_rate` / `work_hours` 尽量转 float（Excel 中出品率数字格式 `0.0"%"` 实际存储为原始数值，解析后即得到 float，如 `98.0`）；为空则保留空字符串 `""`。

### 错误处理

- 找不到 `一、物料信息` 或 `二、工艺工序` 标记时，打印 `PARSE_ERROR` 及说明，并以退出码 `2` 结束。

### 闭环示例

```
# 正向：JSON -> Excel
python3 scripts/generate_bom.py --data bom.json --out BOM_2026-07-07.xlsx

# 逆向：Excel -> JSON（导出的 json 可再次喂回正向流程）
python3 scripts/import_bom.py --in BOM_2026-07-07.xlsx --out bom_back.json

# 重新生成：把逆向得到的 JSON 再传给 generate_bom.py
python3 scripts/generate_bom.py --data bom_back.json --out BOM_v2.xlsx
```

逆向导入得到的 JSON 与正向输入 Schema 一致，因此可直接作为 `generate_bom.py --data` 的输入，实现「导入 → 编辑/查看 → 重新生成」的完整闭环。
