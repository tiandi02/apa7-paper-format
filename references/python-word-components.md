# APA 7 Python Word 组件工作流

本文件说明如何用 skill 内置的 `python-docx` 组件新建或修改 Word 文档。所有命令均从 `apa7-paper-format` 目录执行，或者把 `scripts` 加入 Python 模块搜索路径。

## 选择入口

| 任务 | 入口 |
|---|---|
| 从文本生成完整论文 | `scripts/create_apa7_docx.py` |
| 修改既有 DOCX | `scripts/edit_apa7_docx.py` |
| 编写定制生成程序 | `scripts/apa7_word_components.py` |
| 交付前结构检查 | `scripts/validate_apa7_docx.py` |

新建和编辑任务都应输出到新路径。编辑前先读取原稿，确认标题、表格、页眉、分页和目标段落；编辑后重新打开并验证。

## 新建带真实表格的文档

正文文件使用现有标记，并通过 `[TABLEDATA]` 插入 JSON 中的表格：

```text
[H1]results
The endpoint × language interaction was significant, χ²(5) = 194.12, p < .001 (see {REF:Table 1}).
[TABLE]
[CAPTION]stage 1 predicted probabilities and language contrasts
[TABLEDATA]stage1
[NOTE]OR = odds ratio. p values are Holm corrected; confidence intervals are unadjusted.
```

表格 JSON：

```json
{
  "stage1": {
    "headers": ["Endpoint", "English probability [95% CI]", "Chinese probability [95% CI]", "OR [95% CI]", "Holm p"],
    "rows": [
      ["DeepSeek", ".209 [.188, .232]", ".215 [.195, .236]", "1.03 [0.89, 1.20]", ".656"],
      ["Kimi", ".833 [.813, .852]", ".757 [.735, .778]", "0.62 [0.54, 0.72]", "< .001"]
    ],
    "widths": [1.0, 1.55, 1.55, 1.45, 0.75],
    "font_size": 8.5,
    "first_column_left": true
  }
}
```

`widths` 是相对英寸权重，组件会按正文可用宽度等比缩放；不得少于或多于表头列数。`font_size` 允许 8–12 pt。组件会同时设置表格、网格列和单元格宽度，避免不同 Word 渲染器自行改列宽。

生成命令：

```powershell
python scripts/create_apa7_docx.py `
  --mode professional `
  --title "Language Effects in Large Language Models" `
  --body body.txt `
  --tables tables.json `
  --references references.txt `
  --output manuscript.docx
```

### 多面板表格

当一张 APA 表包含多个面板时，用 `panels` 取代顶层 `rows`：

```json
{
  "descriptive": {
    "headers": ["Endpoint", "0", "1", "Stance", "Conditional AER"],
    "widths": [1.2, 0.6, 0.6, 1.0, 1.2],
    "panels": [
      {"label": "Panel A. English-Prompt Condition", "rows": [["DeepSeek", "473", "473", "31.25%", "50.00%"]]},
      {"label": "Panel B. Chinese-Prompt Condition", "rows": [["DeepSeek", "332", "563", "29.60%", "62.91%"]]}
    ]
  }
}
```

每个面板生成独立的 Word 表格并重复表头，面板标签保持斜体并与后表同页。

## 在定制 Python 程序中使用组件

```python
from docx import Document
from apa7_word_components import (
    add_apa_table,
    add_table_caption,
    add_table_note,
    append_stat_runs,
    apply_apa7_layout,
    format_or_ci,
    format_p,
)

doc = Document()
apply_apa7_layout(doc)

p = doc.add_paragraph()
append_stat_runs(p, "The interaction was significant, p < .001.")

add_table_caption(doc, "Table 1", "Endpoint-Specific Language Contrasts")
add_apa_table(
    doc,
    ["Endpoint", "OR [95% CI]", "Holm p"],
    [["Kimi", format_or_ci(0.6238, 0.5385, 0.7225), format_p(0.0004)]],
    widths=[2.0, 2.7, 1.8],
    font_size=9,
)
add_table_note(doc, "OR = odds ratio. p values are Holm corrected.")
doc.save("paper.docx")
```

数字格式函数接受未舍入数值并采用常规四舍五入：

- `format_probability()`：概率/比例，无前导零。
- `format_p()`：无前导零，小于 .001 时输出 `< .001`。
- `format_or()`：保留前导零。
- `format_or_ci()`：同时格式化 OR 与 95% CI。
- `append_stat_runs()`：自动设置统计符号斜体，并将 `n0`、`n1` 的数字写成 Word 下标。

## 修改既有 DOCX

编辑器使用 JSON 操作规格。输入与输出不得为同一路径。

```json
{
  "normalize_apa": true,
  "operations": [
    {
      "type": "replace_section",
      "heading": "Results",
      "content": [
        {
          "type": "paragraph",
          "text": "The interaction was significant, χ²(5) = 194.12, p < .001."
        },
        {
          "type": "table",
          "number": "Table 2",
          "title": "Type III Likelihood-Ratio Tests",
          "headers": ["Effect", "df", "χ²", "p"],
          "rows": [["Endpoint × Language", "5", "194.12", "< .001"]],
          "widths": [2.8, 0.8, 1.4, 1.0],
          "note": "Fixed effects used sum coding."
        }
      ]
    }
  ]
}
```

```powershell
python scripts/edit_apa7_docx.py `
  --input manuscript_original.docx `
  --spec edits.json `
  --output manuscript_revised.docx
```

### 编辑操作

| 操作 | 必需字段 | 行为 |
|---|---|---|
| `replace_section` | `heading`, `content` | 保留目标标题，替换到下一个同级或更高级标题之前的全部段落和表格 |
| `replace_paragraph` | `match`, `text` | 替换唯一匹配段落的正文，保留段落属性 |
| `insert_after` | `match`, `content` | 在唯一匹配段落后插入内容块 |
| `append_content` | `content` | 在正文末尾追加内容块 |

默认执行精确匹配。只有明确需要子串匹配时才设置 `"contains": true`；如果找到零处或多处，编辑器会停止，防止改错位置。

`content` 支持以下块：

- `paragraph`：`text`，可选 `style`、`alignment`、`first_line_indent`、`bold`、`italic`、`font_size`。
- `heading`：`text`、`level`。
- `table`：与表格 JSON 相同，可附加 `number`、`title`、`note`。
- `table_note`：`text`，可选 `font_size`。

`normalize_apa: true` 会设置 Letter 纸张、1 英寸页边距、12-pt Times New Roman、双倍行距、0.5 英寸首行缩进、标题样式和孤行寡行控制。页眉、自动页码及未替换章节会保留。

含修订、评论、复杂文本框、嵌入对象或自定义 XML 的文档，应先结合 `docx` skill 检查底层 XML；不要在未获用户指示时接受或删除修订。

## 验证

```powershell
python scripts/validate_apa7_docx.py manuscript_revised.docx
python scripts/validate_apa7_docx.py manuscript_revised.docx --json
python scripts/create_apa7_docx.py --selftest
python scripts/test_apa7_word_upgrade.py
```

验证器检查：

- DOCX ZIP 完整性与 `document.xml` 可读性。
- US Letter、1 英寸页边距、字体、行距、首行缩进、孤行寡行控制和自动页码。
- 表格固定布局、无竖线、重复表头、禁止跨页拆行和总列宽。
- OR/CI 前导零与统计符号斜体。
- 修订标记；发现时给出警告，不自动接受或删除。

结构检查通过后仍需转换为 PDF 做视觉检查。逐页确认没有空白页、文字或表格越界、孤立表号/表题、截断行，以及过小而难以阅读的表格文字。
