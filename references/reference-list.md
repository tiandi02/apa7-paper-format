# APA 7 参考文献列表（手册第 9–10 章）

> **外部要求优先**：期刊投稿要求、学校模板、教师要求优先于本文件所述默认规则；冲突时以外部要求为准，并在交付时注明偏差。

## 目录
1. [四要素与标点](#1-四要素与标点)
2. [作者要素](#2-作者要素)
3. [日期要素](#3-日期要素)
4. [标题要素](#4-标题要素)
5. [来源要素与 DOI/URL](#5-来源要素与-doiurl)
6. [信息缺失处理](#6-信息缺失处理)
7. [排序规则](#7-排序规则)
8. [常用文献类型模板](#8-常用文献类型模板)
9. [法律文献（第 11 章）](#9-法律文献第-11-章)
10. [参考文献中的缩写（9.50）](#10-参考文献中的缩写950)
11. [注释书目（9.51）](#11-注释书目951)

## 1. 四要素与标点

每条参考文献 = **作者. (日期). 标题. 来源.**，每个要素后跟句点（DOI/URL 后不加句点，避免破坏链接）。

## 2. 作者要素（9.7–9.12）

- 全部作者姓在前、名缩写：`Author, A. A., & Author, B. B.`
- **≤20 位作者全部列出**（APA 7 变化，旧版为 7 位），末位前用 &；**≥21 位：列前 19 位 + 省略号 + 最后一位**
- 两位作者之间也用逗号分隔：`Author, A. A., & Author, B. B.`
- 编者：`Wong, C. T., & Music, K. (Eds.).`
- 群体作者：写全称，末尾句点；列表中不缩写
- 无作者：标题移入作者位置；署名 "Anonymous" 才用 Anonymous
- 姓名后不加头衔（PhD、Dr.）；带连字符的缩写保留连字符：`Xu, A.-J.`
- 用户名+真名：`Fogarty, M. [Grammar Girl].`（@ 号保留在方括号内）

## 3. 日期要素（9.13–9.17）

- `(2020).` / `(2020, August 26).` / `(2020, Spring/Summer).`
- 无日期：`(n.d.).`；已接收未出版：`(in press).`
- 多卷/连载作品：`(2015-present).` 或 `(2013-2019).`
- 仅对会持续变化的未存档网页加检索日期：`Retrieved October 11, 2020, from https://xxxxx`（放在 URL 前）

## 4. 标题要素（9.18–9.22）

- **属于更大整体的作品**（期刊论文、书籍章节）：sentence case，**不加斜体不加引号**：`The virtue gap in humor: Exploring benevolent and corrective humor.`
- **独立作品**（书、报告、网页、数据集、影片）：sentence case，**斜体**：`Adoption-specific therapy: A guide to helping adopted children thrive.`
- 版次/卷次/报告号放标题后圆括号内：`(2nd ed., Vol. 1).`、`(Report No. 123).`
- 非常规作品加方括号描述：`[Data set]`、`[Film]`、`[Computer software]`、`[Manuscript in preparation]`、`[Press release]`
- 外语作品：原文标题 + 方括号内译文：`[The psychology of the child]`

## 5. 来源要素与 DOI/URL（9.23–9.36）

- **期刊**：刊名 title case 斜体，卷号斜体，期号括号内不斜体，页码：`Journal of Applied Psychology, 104(2), 214-228.`；文章编号：`PLOS ONE, 11(7), Article e0158474.`
- **出版社**：按作品所示名称；作者与出版社同名时省略出版社；去掉 Inc./Ltd.；多出版社用分号
- **编辑书章节**：`In E. E. Editor (Ed.), Title of book (2nd ed., pp. 3-13). Publisher.`
- **学术数据库**（PsycINFO、EBSCO、PubMed 等）：**不写数据库名**；ProQuest 学位论文库、ERIC、PsyArXiv、Cochrane、UpToDate 等受限/专有库要写
- **DOI**：有 DOI 必写，统一格式 `https://doi.org/xxxxx`；有 DOI 又有 URL 时只写 DOI；结尾不加句点
- **URL**：无 DOI 的网页/网站写 URL，直接链到所引内容；URL 失效则找 Internet Archive 存档版
- 不写 ISBN/ISSN；不写 "Retrieved from"（除非有检索日期）

## 6. 信息缺失处理（Table 9.1）

| 缺什么 | 处理 |
|---|---|
| 作者 | 标题移到最前：`Title. (Date). Source.` 文内用标题 |
| 日期 | `(n.d.)` |
| 标题 | 方括号描述代替 |
| 作者+日期 | 标题在前 + n.d. |
| 来源（无法检索） | 不列条目，作为个人交流只在正文引用 |

## 7. 排序规则（9.44–9.49）

`create_apa7_docx.py --references` 默认按以下完整算法自动排序（`--no-sort` 关闭），并自动为同第一作者同年的条目分配 a/b/c 后缀：

1. **按第一作者姓的字母序**；字母比较忽略空格、标点、连字符等一切非字母数字字符——"nothing precedes something"（Marshall 在 Marshall-Petrini 前；Loft 在 Loftus 前；前缀整体在它加字母的变体前）
2. **单作者条目在共同第一作者的多作者条目之前**（9.46），无论年份：Singh (2019) 在 Singh & Shah (2016) 之前
3. **同作者（或同第一作者+同后续作者）按年份**（9.47）：从早到晚；**n.d. 最前、in press 最后**；`(2013-2019)` 排在 `(2015)` 之前
4. **同年同作者按标题字母序**（忽略标题前导 A/An/The），并给年份加 a/b/c（9.47）：`(2020a).`、`(2020b).`、`(n.d.-a).`、`(in press-a).`；后缀顺序 = 参考文献列表中的最终顺序；文内引用必须同步使用相同后缀（如 2019a）
5. **同第一作者、不同后续作者**：先按单作者优先规则排在单作者条目之后，再按第二作者姓、第三作者姓……逐个比较（9.48）
6. **无作者条目**（标题在作者位）：按标题首个实义词字母序（忽略 A/An/The），并入整体字母序（9.49）；标题中的数字按拼写排序（"5" = five）
7. 元分析：被纳入研究的条目前加 *，并在 References 标签下加注说明
8. 解析不了日期括号的条目排在最后、保持输入顺序

示例（最终顺序）：

```
Aarons, B. B., & Zhao, X. (2018). ...        ← 第一作者字母序
American Psychological Association. (2020)…   ← 群体作者按全称
The early bird catches the worm. (2020). …    ← 无作者：按 early（忽略 The）
Marshall, A. (2010). …                        ← nothing precedes something
Marshall-Petrini, S. (2011). …
Singh, R. (n.d.-a). Another no date. …        ← n.d. 最前；标题序定 a/b
Singh, R. (n.d.-b). No date. …
Singh, R. (2015). …
Singh, R. (2019a). Earlier title work. …      ← 同年按标题 + a/b/c
Singh, R. (2019b). Later work. …
Singh, R. (in press). …                       ← in press 最后（单作者仍先于多作者）
Singh, R., & Shah, S. (2016). …               ← 单作者条目之后按第二作者
```

- 中文文献排序是**扩展约定**（手册未规定）：本 skill 建议按作者拼音与西文条目统一排序；教师/机构另有要求时服从其要求
- a/b/c 后缀只解决**同第一作者+同年**的情况；若加后缀后文内 et al. 引用仍可能歧义（不同作者组合缩写后相同），按 8.18 在文内写开足够多的姓名

## 8. 常用文献类型模板

**期刊论文（有 DOI）**
```
Author, A. A., & Author, B. B. (2020). Title of article. Journal Title, 34(2), 14-25. https://doi.org/10.xxxx
```
**期刊论文（无 DOI、印刷版/学术数据库）**：同上但省略 DOI。

**整本书**
```
Author, A. A. (2020). Title of book (2nd ed.). Publisher.
```
**编辑书章节**
```
Author, A. A. (2020). Title of chapter. In E. E. Editor & F. F. Editor (Eds.), Title of book (pp. 3-13). Publisher. https://doi.org/10.xxxx
```
**学位论文**
```
Author, A. A. (2020). Title of dissertation [Doctoral dissertation, University Name]. ProQuest Dissertations and Theses Global.
Author, A. A. (2020). Title of thesis [Unpublished master's thesis]. University Name.
```
**政府/机构报告**（作者=出版社时省略出版社）
```
National Cancer Institute. (2018). Title of report (NIH Publication No. 18-2424). https://www.cancer.gov/xxx.pdf
```
**网页/网站**（作者=网站名时省略网站名）
```
Author, A. A. (2020, May 2). Title of page. Site Name. https://xxxxx
```
**维基百科**（用存档版本）
```
List of oldest companies. (2019, January 13). In Wikipedia. https://en.wikipedia.org/w/index.php?title=...&oldid=878158136
```
**报纸/杂志/博客**
```
Guarino, B. (2017, December 4). Title of article. The Washington Post. https://xxxxx
Bergeson, S. (2019, January 4). Really cool neutral plasmas. Science, 363(6422), 33-34.
Klymkowsky, M. (2018, September 15). Title of post. Blog Name. https://xxxxx
```
**会议报告**
```
Presenter, A. A. (2020, September 18-20). Title of contribution [Poster presentation]. Conference Name, City, State, Country. https://xxxxx
```
**电影 / 剧集 / 播客**
```
Forman, M. (Director). (1975). One flew over the cuckoo's nest [Film]. United Artists.
Barris, K. (Writer & Director). (2017, January 11). Lemons (Season 3, Episode 12) [TV series episode]. In K. Barris et al. (Executive Producers), Black-ish. ABC Studios.
Glass, I. (Host). (2011, August 12). Amusement park (No. 443) [Audio podcast episode]. In This American life. WBEZ Chicago. https://xxxxx
```
**YouTube 视频**（上传者即作者）
```
University of Oxford. (2018, December 6). How do geckos walk on water? [Video]. YouTube. https://www.youtube.com/watch?v=xxxx
```
**数据集 / 软件**
```
D'Souza, A., & Wiseheart, M. (2018). Title of data set (Version V1) [Data set]. ICPSR. https://doi.org/10.xxxx
Borenstein, M., Hedges, L., Higgins, J., & Rothstein, H. (2014). Comprehensive meta-analysis (Version 3.3.070) [Computer software]. Biostat. https://www.meta-analysis.com/
```
**社交媒体**
```
Author, A. A. [@username]. (2020, May 2). Content of the post up to the first 20 words [Tweet]. Twitter. https://twitter.com/xxxx
```
**预印本 / 未发表手稿**
```
Leuker, C., Samartzidis, L., Hertwig, R., & Pleskac, T. J. (2018). Title. PsyArXiv. https://doi.org/10.17605/OSF.IO/9P7CB
Author, A. A. (2020). Title [Manuscript in preparation]. Department of Psychology, University Name.
```
**译本 / 再版**
```
Piaget, J., & Inhelder, B. (1969). The psychology of the child (H. Weaver, Trans.; 2nd ed.). Basic Books. (Original work published 1966)
```
**宗教/经典著作**
```
King James Bible. (2017). King James Bible Online. https://www.kingjamesbibleonline.org/ (Original work published 1769)
```
**诊断手册**（作者=出版社，省略出版社）
```
American Psychiatric Association. (2013). Diagnostic and statistical manual of mental disorders (5th ed.). https://doi.org/10.1176/appi.books.9780890425596
```

## 9. 法律文献（第 11 章）

法律文献**不用作者-日期制**，采用法律引用格式（Bluebook 风格）；文内引用直接给出案件名/法条与年份，如 `(Brown v. Board of Education, 1954)`。

**判例（court decision）**
```
Name v. Name, Volume Reporter Page (Court Year). https://xxxxx
```
例：`Brown v. Board of Education, 347 U.S. 483 (1954).`

文内引用时案件名**斜体**（11.4）：`(Brown v. Board of Education, 1954)`；参考文献列表中用标准体。

**成文法（statute）**
```
Name of Act, Title Source § Section (Year). https://xxxxx
```
例：`Every Student Succeeds Act, 20 U.S.C. § 6301 (2015).`

**联邦规章（federal regulation）**
```
Title/Authority, Source § Section (Year). https://xxxxx
```
例：`Protection of Human Subjects, 45 C.F.R. § 46 (2009).`

**宪法（11.9）**
```
U.S. Const. art. I, § 3.
U.S. Const. amend. XIV, § 1.
```
- 引用**整部宪法**时只在正文提及，不列参考文献；引用**具体条款或修正案**时必须进参考文献（用上述 art./amend. 模板）
- 论文中引用的判例与法条须列入 References
- 州法规、行政命令、国际条约等更多模板见手册第 11 章与 APA Style 官网

## 10. 参考文献中的缩写（9.50）

| 缩写 | 含义 | 缩写 | 含义 |
|---|---|---|---|
| ed. | edition | Vol. (Vols.) | volume(s) |
| Rev. ed. | revised edition | No. | number |
| 2nd ed. | second edition | Pt. | part |
| Ed. (Eds.) | editor(s) | Tech. Rep. | technical report |
| Trans. | translator(s) | Suppl. | supplement |
| Narr. (Narrs.) | narrator(s) | n.d. | no date |
| p. (pp.) | page(s) | para. (paras.) | paragraph(s) |

## 11. 注释书目（9.51）

- 条目与排序规则同参考文献列表
- 每条注释是其条目下方的新段落：**整段左缩进 0.5 英寸（同块引用），首行不额外缩进**
- 注释多段时，第二段起首行再缩进 0.5 英寸
- 注释内一般不必再引用该条目本身；引用其他文献时用常规文内引用
**歌曲 / 专辑**
```
Beyoncé. (2016). Formation [Song]. On Lemonade. Parkwood; Columbia.
Bach, J. S. (2010). The Brandenburg concertos [Album recorded by Academy of St Martin in the Fields]. Decca. (Original work published 1721)
```
**字典/百科词条**（持续更新、无存档版本 → 需要检索日期）
```
Merriam-Webster. (n.d.). Self-report. In Merriam-Webster.com dictionary. Retrieved July 12, 2019, from https://www.merriam-webster.com/dictionary/self-report
```
**测验/量表手册**（优先引用支持文献即手册，而非测验本身）
```
Tellegen, A., & Ben-Porath, Y. S. (2011). Minnesota Multiphasic Personality Inventory-2 Restructured Form (MMPI-2-RF): Technical manual. Pearson.
```
**书评 / 影评**（按所载媒介的格式 + 方括号内被评作品信息）
```
Santos, F. (2019, January 11). Reframing refugee children's stories [Review of the book We are displaced, by M. Yousafzai]. The New York Times. https://nyti.ms/xxxx
Mirabito, L. A., & Heck, N. C. (2016). Bringing LGBTQ youth theater into the spotlight [Review of the film The year we thought about love, by E. Brodsky, Dir.]. Psychology of Sexual Orientation and Gender Diversity, 3(4), 499-500. https://doi.org/10.1037/sgd0000205
```
**PowerPoint 幻灯片 / 网络研讨会**
```
Author, A. A. (2020). Title of slides [PowerPoint slides]. Site Name. https://xxxxx
Goldberg, J. F. (2018). Evaluating adverse drug effects [Webinar]. American Psychiatric Association. https://xxxxx
```
**地图 / 信息图**
```
Cartographer, C. C. (2020). Title of map [Map]. Publisher or Site. https://xxxxx
Rossman, J., & Palmer, R. (2015). Sorting through our space junk [Infographic]. World Science Festival. https://www.worldsciencefestival.com/xxxx
```
