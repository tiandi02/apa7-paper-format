# APA 7th 论文格式规范

一个面向 [Claude Code](https://claude.com/claude-code) 与 Codex 的 **Skill**，用于创建、修改和验证符合 **APA Style 第 7 版**（*Publication Manual of the American Psychological Association*, 7th ed., 2020）要求的学术论文 Word 文档。

> 本 skill 根据**最新版 APA 7th 规范（2020 年第 7 版）**制作，涵盖论文要素与版式、作者-日期文内引用、参考文献列表四要素与完整排序（9.44–9.49）、数字与统计符号、标题大小写、公式排版，以及交付前的 APA 7 合规自查清单和 Word 文档生成脚本。自 v1.5 起并入手册第 1–12 章按需章节库，覆盖 JARS、写作风格、无偏见语言、表格与图、法律文献和出版流程。

版本：**v1.5**（2026-09-20） · 许可证：[MIT](LICENSE) · 适用：中英文论文

## 功能特性

- **完整论文要素**：学生版/专业版标题页（多作者多单位上标编号对应）、摘要页（≤250 词检查）、关键词（专有名词自动保留，如 COVID-19、Big Five）、五级标题、作者注、running head（自动生成、≤50 字符）
- **真实 Word 高级排版**（非文字模拟）：
  - 真正的 Word 脚注（`/word/footnotes.xml`，自动编号、页脚单倍行距）
  - OMML 公式对象（`\frac`、`\sqrt`、上下标、`\sum_{i=1}^{n}`、希腊字母），居中显示 + 右对齐编号 `(1)(2)…`
  - 题注/公式交叉引用域（REF 域，Word 打开自动更新，表图增删后编号自动跟随）
  - 附录表图自动编号（Table A1、Figure C2，每附录独立计数）
  - URL/DOI 真实超链接
- **参考文献完整排序（APA 9.44–9.49）**：单作者优先、同作者按日期（n.d. 最前 / in press 最后）、同第一作者按后续作者、无作者按标题、同年同作者自动加 a/b/c 后缀
- **真实 APA Word 表格**：支持 JSON 表格规格、多面板表格、固定列宽、无竖线、重复表头、禁止数据行跨页拆分，以及规范的表题和表注
- **保格式修改既有 DOCX**：按章节或段落替换、在指定位置插入内容和表格、追加内容，同时保留未修改章节、页眉和自动页码；强制另存为新文件
- **统计格式组件**：自动处理 *p*、*OR*、*df*、*n* 等统计符号的斜体，以及概率、*p* 值、优势比和置信区间的前导零规则
- **独立文档验证器**：检查 DOCX 包、Letter 纸张、页边距、字体、行距、首行缩进、页码、表格结构和统计格式
- **回归自检**：`--selftest` 一键跑 40+ 断言（已用 Word 无头打开实测通过）
- **手册第 1–12 章知识库**：按需加载 `chapters/`，并提供 `glossary.md`、`patterns.md`、`cheatsheet.md`；内容为结构化提炼，标注手册章节来源，不替代 `references/` 中的细则

## 安装

将本仓库复制到 Claude Code 的 skills 目录：

**全局**（所有项目可用）：

```bash
git clone https://github.com/<你的用户名>/apa7-paper-format.git \
  ~/.claude/skills/apa7-paper-format
```

**项目级**：复制到项目根目录的 `.claude/skills/apa7-paper-format/`。

安装后在 Claude Code 中直接说"写论文"，或使用 `/<skill-name>` 调用；任何论文写作任务（essay、课程论文、文献综述、毕业论文、研究报告、润色/改写/审查、检查引用）都会自动触发本 skill。

仓库的 `dist/apa7-paper-format.skill` 是已经验证的可迁移安装包；需要在支持 `.skill` 包的环境中安装时，可直接使用该文件。

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

# 使用 JSON 表格规格生成真实 Word 表格
python scripts/create_apa7_docx.py --body body.txt --tables tables.json \
    --output paper-with-tables.docx

# 修改既有文档并另存为新文件
python scripts/edit_apa7_docx.py --input original.docx \
    --spec edits.json --output revised.docx

# 独立验证 APA 7 文档
python scripts/validate_apa7_docx.py revised.docx

# 回归自检
python scripts/create_apa7_docx.py --selftest
python scripts/test_apa7_word_upgrade.py
```

正文标记示例：

```
[H2]participants
Two hundred participants were recruited.{FN:Sample size was determined by a power analysis.}
The model appears in {EQREF:1}, and the results appear in {REF:Table 1}.
[EQ]d = \frac{\sum_{i=1}^{n} (x_i - M)^2}{n - 1}
[TABLE]
[CAPTION]descriptive statistics for study measures
[TABLEDATA]descriptive_statistics
[APPENDIX]Appendix A
```

## 文件结构

```
apa7-paper-format/
├── SKILL.md                     # Skill 定义：工作流程、核心速查表、APA 7 合规自查清单
├── chapters/                    # 手册第 1–12 章按需章节库与主题索引
│   ├── README.md                # 章节导航与主题索引
│   └── ch01-*.md … ch12-*.md    # 伦理/JARS/风格/去偏/表图/引用/法律/出版流程
├── glossary.md                  # APA 7 术语表（含章节引用）
├── patterns.md                  # 可复用的写作、报告、引用、表图、版权和投稿模式
├── cheatsheet.md                # 判定规则、阈值与决策表
├── dist/
│   └── apa7-paper-format.skill  # 已验证的可迁移安装包
├── evals/
│   └── evals.json               # 新建、编辑和验证场景的标准评估任务
├── references/
│   ├── formatting.md            # 论文要素、页序、版式（五级标题/页眉/脚注/附录/表格图）
│   ├── citations.md             # 文内引用（作者-日期制、et al.、直接引用定位信息）
│   ├── reference-list.md        # 参考文献四要素、完整排序算法（9.44–9.49）、类型模板
│   ├── mechanics.md             # 数字、统计符号、标题大小写、公式排版
│   └── python-word-components.md # Python Word 组件、表格 JSON 和编辑规格说明
└── scripts/
    ├── apa7_word_components.py  # APA 版式、真实表格和统计格式公共组件
    ├── create_apa7_docx.py      # 新建 APA 7 版式 .docx
    ├── edit_apa7_docx.py        # 保格式修改既有 DOCX
    ├── validate_apa7_docx.py    # 独立结构与格式验证器
    └── test_apa7_word_upgrade.py # 新版能力回归测试
```

## 外部要求优先

期刊投稿要求、学校模板、教师要求**优先于**本 skill 的默认规则；有冲突时一律以外部要求为准。本 skill 提供的是无外部要求时的 APA 7 官方默认。

## 免责声明

本项目的作者与 American Psychological Association 无隶属关系。"APA" 及 "Publication Manual of the American Psychological Association" 为 American Psychological Association 的商标。本项目是根据第 7 版出版手册（2020）整理的独立实现，非官方出版物。

## 许可证

[MIT](LICENSE) © 2026
