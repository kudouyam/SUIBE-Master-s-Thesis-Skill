# GitHub 上外贸论文相关开源项目审查与迁移矩阵（v1.1）

> 审查日期：2026-10-11。**本报告中的 GitHub 项目均为第三方非官方实现**；其中有本科模板、2018 年硕士模板、LaTeX 模板和只提供压缩包的仓库。2020 年《上海对外经贸大学硕士学位论文内容和格式要求》及用户提供的同套范例**始终高于**这些仓库的任何实现。GitHub 内容用于设计参考、错误预警、实现启发，不能倒过来修改校规。审查范围及不能审查的项目如下明确说明。

## 1. 仓库清单与实查范围

| 仓库 | 属性与许可 | 已核查资料 | 结论 |
|---|---|---|---|
| [Yangruipis/SUIBEthesis](https://github.com/Yangruipis/SUIBEthesis) | 2018-09 创建的硕士 LaTeX 模板；MIT | README、`SUIBEthesis.cls`、`main.tex`、文件树、License 元数据 | 重要参考，但 2018 版本多处参数与 2020 硕士规范冲突 |
| [wojiaobadaye/SUIBE_Latex](https://github.com/wojiaobadaye/SUIBE_Latex) | 2023；仓库没有发现许可证；论文+PPT 模板 | README、文件树；实际 LaTeX 源码在 `moban_paper.zip`，未能直接读取二进制 ZIP | README 称改进 GB/T 7714-2015 著录；**源码未验，不得当成确证** |
| [Epiphany-Leon/Bachelor-thesis-template-SUIBE](https://github.com/Epiphany-Leon/Bachelor-thesis-template-SUIBE) | 2025 本科、作者为全英文论文；MIT | README、`main.tex`、封面/声明源文件 | 有引用管理与公式排版启发；明确不适用于硕士格式基线 |
| [cf3901646/suibe-thesis-template](https://github.com/cf3901646/suibe-thesis-template) | 2026-09，本科“2027届”；未找到许可证 | README、文件树；仓库附 `.doc` 二进制样张，未能逐段提取验证 | README 不能代替校方官网最新版规范；本科 1.5 倍行距不可迁移 |
| [MadisonRuoxiTian/suibe-undergraduate-thesis-latex-template](https://github.com/MadisonRuoxiTian/suibe-undergraduate-thesis-latex-template) | 2026 本科；子目录有 MIT 许可证，顶层无仓库许可证声明 | README、子目录 README、`suibe-thesis.cls` | 参考元数据统一入口、分节与对象保护；默认版式均不迁移 |
| [ChangjiangDong/SUIBEthesis](https://github.com/ChangjiangDong/SUIBEthesis) | `Yangruipis/SUIBEthesis` 的 GitHub Fork；MIT | GitHub parent/source 元数据 | 去重，未发现需要独立替换的校规来源 |
| [JarvanShi/SUIBEthesis](https://github.com/JarvanShi/SUIBEthesis) | 同上；MIT | GitHub parent/source 元数据 | 去重 |

> **获取限制**：本次通过 GitHub 可读取的 API 与原始文本检查仓库元数据和关键源码，无法把远程 Git 仓库整体 `git clone` 到本地工作环境；二进制 ZIP / 老式 DOC 的具体内容未完成审阅。该限制不影响已读 `.cls`、`.tex`、README 的明确结论，但禁止声称“已完整检查所有仓库的每一个文件”。后续可在有 Git 网络的 Codex 设备上按第6节命令补完。

## 2. 核心差异：2018 硕士项目与校方 2020 版

项目直接引用源码：
- [2018 类文件 `SUIBEthesis.cls`](https://github.com/Yangruipis/SUIBEthesis/blob/master/SUIBEthesis.cls)
- [2018 实例 `main.tex`](https://github.com/Yangruipis/SUIBEthesis/blob/master/main.tex)
- 官方基线：本 Skill 内 `sources/上海对外经贸大学硕士学位论文内容和格式要求_2020修订.pdf`（第3—7页）、格式范例（第5—8页）。

| 检查项 | 项目真实源码 | 2020 硕士规范/范例 | v1.1 决策 |
|---|---|---|---|
| 纸张 | `a4paper` | A4 | 兼容；保留 |
| 边距 | `top=2.54cm,bottom=2.54cm,left=3.18cm,right=3.18cm` | 上37、下30、左35、右25mm | **绝不复制**2018模板 |
| 一级标题 | `\titleformat*{\section}{\bfseries\centering\sanhao}`；其三号字号定义16bp | 范例为左对齐、宋体小三15pt | 不可复制居中和字号 |
| 二、三级标题 | 二级14bp、三级12bp | 二级14pt、三级12pt | 数值相近，但仍以学校为据 |
| 正文行距 | 先后设置 `\baselinestretch` 1.5 与1.25（覆盖关系）；并无Word固定20pt机制 | Word 固定20pt | 不迁移倍数行距 |
| 图表题 | `caption` 使用 `footnotesize`，并自定义加粗 | 宋体小五9pt加粗，上表题下图题 | 参考题签对象分离，但**不可继承不确定的footnotesize** |
| 图表编号 | LaTeX 自动生成 | 按章 `图 3-1 / 表 2-7` | 引入 Word 图表编号一致性审查；**不擅自批量改号** |
| 目录 | 用 `tocloft` 定义一至三级目录字体 | 一至三级、大小15/14/12pt、固定20pt，页码12pt | 继承“目录与正文分层”思路，实际使用学校明文 |
| 文献 | `natbib` 与自制 BST，`ref.bib` | GB/T7714-2015，可顺序编码或作者-出版年 | 增强引用体例保护/检查提示，不改文献内容 |
| 表格 | 演示使用 `booktabs` 的横线命令（示例有重复上/下线命令） | 建议三线表，范例中复杂表格有竖线 | 保守三线表；不能无条件删线 |
| 目录后正文页码 | 实例调用 `\setcounter{page}{1}` | 正文页脚居中，版式要实际逐页验收 | 保留 Word 字段、正文独立分节及页码校验 |

## 3. 本科仓库的可借鉴和禁止迁移项

### Epiphany-Leon（2025，本科，全英文）

来源：[README](https://github.com/Epiphany-Leon/Bachelor-thesis-template-SUIBE/blob/main/README.md)、[main.tex](https://github.com/Epiphany-Leon/Bachelor-thesis-template-SUIBE/blob/main/main.tex)。作者明确警告“本科与研究生论文格式不同”、“本人为全英文论文”。其中 `geometry{left=3cm,right=2.5cm,top=3cm,bottom=2.5cm}`、英文图表名、图表题 `10.5pt`、`onehalfspacing`、文献自定义规则与硕士2020规范不匹配。可借鉴：图、表、正文、参考文献分区管理；分层自动编号、引用和公式命令；不宜继承其样式参数。

### cf3901646（2027届本科信息汇编）

来源：[README](https://github.com/cf3901646/suibe-thesis-template/blob/main/README.md)。提供 Word `.doc` 样张，README 称其为最新本科官方规范；但本次没有得到原始学校官网发布的可核验通知，也未能检查 `.doc` 二进制内容。README 提到全文1.5倍行距、图表题五号、Harvard参考文献体例。这些**只可当“本科项目作者的描述”**，不可覆盖本 Skill 的硕士2020年行距固定20pt、图表题小五9pt、GB/T7714-2015文献规定。

### MadisonRuoxiTian（2026本科 class）

来源：[class](https://github.com/MadisonRuoxiTian/suibe-undergraduate-thesis-latex-template/blob/main/suibe_latex_class_release_with_preview/suibe_latex_class_release/suibe-thesis.cls)、[说明](https://github.com/MadisonRuoxiTian/suibe-undergraduate-thesis-latex-template)。使用 `pgfkeys` 管理元数据和封面字段，有前后置结构、图片位置、重复表头等思路可学。但该模板正文1.5倍行距、边距30/26/28/26mm、`ctex` 黑体标题、图题`small`、图表全局序号均非2020硕士校规。可迁移的是**构建规则表与保护对象的工程结构**，不是格式参数。

### wojiaobadaye（2023源码待补）

来源：[README](https://github.com/wojiaobadaye/SUIBE_Latex/blob/main/README.md)。有两个二进制 ZIP：`moban_paper.zip`、`moban_beamer.zip`。可确认 README 自述改进 GB/T 7714—2015 格式；**不能确认二进制中的实现精度、字号及引用边距**。缺少明确开源许可证时，不能把归档内代码直接打进此 Skill。建议后续下载后只做对照审查。

## 4. 本次真正落地的程序改动

1. **语义颜色保全（新增测试）**：普通无合并表格如果正文单元格出现非白色填充，保留全表原有色彩和边框，并报告 `semantic_table_color`。避免统计热力表被粗暴洗成白色三线表。只有装饰性表头染色且正文无彩色、结构简单时才执行用户指定的简洁三线表默认规则。
2. **列表标题消歧（新增测试）**：Word 原生 `List Number`、`List Paragraph`/中文列表样式不再仅凭“1.”等字形自动变一级标题，优先保留列表结构。
3. **安全图题换位（新增测试）**：图题紧邻纯行内图片且位于其上时，移动完整图题段落到图片下；浮动图、复合图、含文字图片段落保留原位并记录待复核。
4. **图表连续性审查（新增测试）**：按图/表、章号建立索引，检测同章重复序号、缺号、题签章号与最近一级标题不匹配；保留原编号以免破坏正文交叉引用。
5. **OOXML 扩展完整性审查**：比较输入和输出中 `word/charts/`、`word/embeddings/`、`word/diagrams/`、`word/activeX/`、`customXml/` 和 Word 批注相关附属包的哈希/路径，丢失和篡改将被标严重。
6. **修正报告增加诊断计数**：明确列出保留有语义底色的表格数、安全调整图题数、核查的图表编号组数。
7. **校规硬性要求没有被源代码覆盖**：边距、页眉页脚、章首换页、目录、正文字体、脚注、标题字号、图题、表题沿用本 Skill 的官方矩阵。

**尚未自动解决**：整个引用系统 GB/T 7714 实质正确性、Word 原生自动编号与图表题字段的完美重映射、跨页续表“（续）”的自动布局、复杂脚注引用与数学公式对象的版面排布、二进制外部模板的全面审查。这些仍需人工审核并在用户报告里显示。

## 5. 版权、依赖及可信度

- 2018 硕士项目与 Epiphany 项目都标为 **MIT**；本 Skill 的实现是独立 Python 代码，仅引用事实、源路径、工程思想，**没有直接复制第三方类文件、图像、学校 Logo、版权模板或 BST**。
- `wojiaobadaye`、`cf3901646` 未检到仓库许可声明：不能因为 GitHub 可公开访问就当做可再分发。
- 研究生院原始 PDF 由用户提供、供个人学习及格式对照；若将包含学校 PDF 的 ZIP 公开再分发，**应先确认学校材料授权**。公开版建议仅附来源链接和独立规则说明，而不捆绑学校 PDF。
- 每次入参均应提醒“2020基线尚未经当届学院导师确认”；本科项目不能证明 2026/2027 硕士标准已改变。

## 6. 下轮用有 Git 网络的设备补全源文件审查

```bash
git clone --depth 1 https://github.com/Yangruipis/SUIBEthesis.git
git clone --depth 1 https://github.com/wojiaobadaye/SUIBE_Latex.git
git clone --depth 1 https://github.com/Epiphany-Leon/Bachelor-thesis-template-SUIBE.git
git clone --depth 1 https://github.com/cf3901646/suibe-thesis-template.git
git clone --depth 1 https://github.com/MadisonRuoxiTian/suibe-undergraduate-thesis-latex-template.git
```

解包 `wojiaobadaye/moban_paper.zip` 审查 `.tex/.cls/.bst` 和许可；另将 `cf3901646` 的 `.doc` 只读转换为 `.docx` / PDF 以核对 README 的描述。记录 Git commit SHA，以便来源锁定。不必重复 clone 2018 原项目的 Fork；若后续发现 Fork 有独立补丁，再针对 diff 分析。