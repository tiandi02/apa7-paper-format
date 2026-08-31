# APA 7th 论文格式规范

一个面向 [Claude Code](https://claude.com/claude-code) 的 **Skill**，用于生成严格符合 **APA Style 第 7 版**（*Publication Manual of the American Psychological Association*, 7th ed., 2020）格式的学术论文。

> 本 skill 根据**最新版 APA 7th 规范（2020 年第 7 版）**制作，涵盖论文要素与版式、作者-日期文内引用、参考文献列表四要素与完整排序（9.44–9.49）、数字与统计符号、标题大小写、公式排版，以及交付前的 APA 7 合规自查清单和 Word 文档生成脚本。

版本：**v1.3**（2026-08-31） · 许可证：[MIT](LICENSE) · 适用：中英文论文

## 功能特性

- **完整论文要素**：学生版/专业版标题页（多作者多单位上标编号对应）、摘要页（≤250 词检查）、关键词（专有名词自动保留，如 COVID-19、Big Five）、五级标题、作者注、running head（自动生成、≤50 字符）
- **真实 Word 高级排版**（非文字模拟）：
  - 真正的 Word 脚注（`/word/footnotes.xml`，自动编号、页脚单倍行距）
  - OMML 公式对象（`\frac`、`\sqrt`、上下标、`\sum_{i=1}^{n}`、希腊字母），居中显示 + 右对齐编号 `(1)(2)…`
  - 题注/公式交叉引用域（REF 域，Word 打开自动更新，表图增删后编号自动跟随）
  - 附录表图自动编号（Table A1、Figure C2，每附录独立计数）
  - URL/DOI 真实超链接
- **参考文献完整排序（APA 9.44–9.49）**：单作者优先、同作者按日期（n.d. 最前 / in press 最后）、同第一作者按后续作者、无作者按标题、同年同作者自动加 a/b/c 后缀
- **版式自动校验**：生成后自动读回检查页边距、字体、行距、孤行寡行控制、页码域、逐段格式，不达标即报错
- **回归自检**：`--selftest` 一键跑 40+ 断言（已用 Word 无头打开实测通过）

## 安装

将本仓库复制到 Claude Code 的 skills 目录：

**全局**（所有项目可用）：

```bash
git clone https://github.com/<你的用户名>/apa7-paper-format.git \
  ~/.claude/skills/apa7-paper-format
```

**项目级**：复制到项目根目录的 `.claude/skills/apa7-paper-format/`。

安装后在 Claude Code 中直接说"写论文"，或使用 `/<skill-name>` 调用；任何论文写作任务（essay、课程论文、文献综述、毕业论文、研究报告、润色/改写/审查、检查引用）都会自动触发本 skill。

## 使用方法

生成 APA 7 版式 Word 文档（脚本自带完整标记语法，详见文件头 docstring）：

```bash
# 学生论文（多作者多单位）
python scripts/create_apa7_docx.py --title "the effect of X on Y" \
    --author "Ming Li^1^; Zhang San^2^" \
    --affiliation "Department of Psychology, XX University^1^; School of Education, YY University^2^" \
    --course "PSY204" --instructor "Dr. Zhang" --date "March 16, 2026" \
    --abstract abstract.txt --keywords "COVID-19 pandemic, Big Five, depression symptoms" \
    --body body.txt --references refs.txt --output paper.docx

# 专业论文（作者注 + running head）
python scripts/create_apa7_docx.py --mode professional \
    --title "..." --author "..." --author-note note.txt \
    --running-head "EFFECT OF X ON Y" --output paper.docx

# 回归自检
python scripts/create_apa7_docx.py --selftest
```

正文标记示例：

```
[H2]participants
Two hundred participants were recruited.{FN:Sample size was determined by a power analysis.}
The model appears in {EQREF:1}, and the results appear in {REF:Table 1}.
[EQ]d = \frac{\sum_{i=1}^{n} (x_i - M)^2}{n - 1}
[TABLE]
[CAPTION]descriptive statistics for study measures
[APPENDIX]Appendix A
```

## 文件结构

```
apa7-paper-format/
├── SKILL.md                     # Skill 定义：工作流程、核心速查表、APA 7 合规自查清单
├── references/
│   ├── formatting.md            # 论文要素、页序、版式（五级标题/页眉/脚注/附录/表格图）
│   ├── citations.md             # 文内引用（作者-日期制、et al.、直接引用定位信息）
│   ├── reference-list.md        # 参考文献四要素、完整排序算法（9.44–9.49）、类型模板
│   └── mechanics.md             # 数字、统计符号、标题大小写、公式排版
└── scripts/
    └── create_apa7_docx.py      # 生成 APA 7 版式 .docx（含 --selftest 回归自检）
```

## 外部要求优先

期刊投稿要求、学校模板、教师要求**优先于**本 skill 的默认规则；有冲突时一律以外部要求为准。本 skill 提供的是无外部要求时的 APA 7 官方默认。

## 免责声明

本项目的作者与 American Psychological Association 无隶属关系。"APA" 及 "Publication Manual of the American Psychological Association" 为 American Psychological Association 的商标。本项目是根据第 7 版出版手册（2020）整理的独立实现，非官方出版物。

## 许可证

[MIT](LICENSE) © 2026
