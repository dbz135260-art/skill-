---
name: pinshu-generator
description: 聘书批量生成：解析日程表.docx（含合并单元格），按老师聚合并智能时间段，填充红头模板占位符，输出合并Word文档。触发：聘书、授课专家聘书、批量生成聘书、培训聘书。
license: MIT
---

# 聘书批量生成 Skill

## 输入

1. `聘书模板.docx` — 红头文件，含占位符：
   - （会议名字）、（对方单位）、（起止时间）、（结束时间）
   - （地点）、（老师名字）、（上课时间）、（课程内容）、（授课地址）、（发函时间）
2. 日程安排表.docx — 含表格（日期、时间、内容、授课专家），可能有合并单元格

## 核心流程

### Step 1: 解析日程表

用 `python-docx` + `lxml` 解析表格，关键点：
- **合并单元格文本去重**：vMerge 单元格的 runs 内文本会重复 2x 或 3x，需 `_dedup()` 函数对切去重
- **日期继承**：vMerge=continue 时 carry_date 不变
- **课程继承**：课程列也有 vMerge，需 carry_course
- **专家继承**：vMerge=continue 时继承 prev_expert
- **时段提取**：优先检测"上午"/"下午"关键词，兜底靠时间范围（结束≤12→上午）

### Step 2: 按老师合并

- 同姓名+同单位 → 合并
- 智能时间段聚合：
  - 同日有上午+下午 → "X月X日全天"
  - 仅上午 → "X月X日上午"
  - 跨天用 `、` 连接
- 课程合并：
  - `\d+[、.）)]` 开头 → 新条目
  - 无序号开头 → 拼到上一行
  - 不同 session 的课程去重

### Step 3: 填充模板

使用 `python-docx` 在 run 级别替换占位符（保格式）：
```python
for para in doc.paragraphs:
    for run in para.runs:
        run.text = run.text.replace(placehoder, value)
```

课程多行处理：
- 第一行替换 `__COURSE__` 占位符
- 后续行：`copy.deepcopy(段落element)` → 清除原文本 → 写入新行 → `addnext()` 插入

### Step 4: 合并输出

- 第一个老师正常填充
- 后续老师：加分页符 `<w:br w:type="page"/>` → 克隆模板 → 填充 → 追加到 body
- 保存为单个 .docx

## 公共信息（用户提供）

- meeting_name: 培训名称
- start_date / end_date: 起止日期
- location / address: 地点和详细地址
- letter_date: 发函时间

## 输出

- `聘书_全部.docx` — 所有聘书合并，分页隔开
- `聘书输出/` 文件夹 — 每人单独 .docx + ZIP包

## 依赖

```bash
pip install python-docx lxml
```

## 关键陷阱

1. vMerge 文本在 w:t 节点内重复，不是段落级
2. 课程列也 vMerge，忘记继承会导致 `_skip()` 误判跳过
3. 双层括号 `（（主任））` 需循环 strip
4. 输出编码：避免 print emoji（GBK 环境炸）
