# SUIBE Master's Thesis Skill

**上海对外经贸大学硕士学位论文 Word 格式规范化与审查** · Unofficial community project · v1.1

> 给能够读取 GitHub、操作文件并运行 Python 的 AI Agent 使用。自动识别多级标题、摘要、目录、图表与脚注，尽可能保守地将可编辑的 Word 论文转为符合所附 **2020 年修订版硕士学位论文要求** 的格式，并生成修正报告。请先确认学校/学院是否另有最新要求。

**请注意：本项目不是学校官方工具；自动格式化不代表完成最终论文验收。原始 Word 永不覆盖。**

## 给 Agent 一个链接就能用（建议复制这句话）

> 请打开 https://github.com/kudouyam/SUIBE-Master-s-Thesis-Skill ，先阅读根目录 `SKILL.md` 与 `AGENTS.md`，获取整个仓库，按 `README.md` 安装依赖。请用该 Skill 将我提供的可编辑 DOCX 规范化为上外贸硕士论文格式；不要修改研究内容，交付新的 DOCX、Markdown/JSON 格式修正报告，并把不能确定的标题、图表、分页与来源版本问题明确列出。若无法运行代码或不能访问 GitHub，请说明限制，不要假装已完成。

> **Agent prerequisites:** GitHub access, local file operations, Python 3.10+, installable dependencies, and a Word/PDF rendering option for visual QA. Merely sending a URL to an agent with no network/file tools will not execute it.

## 直接运行

```bash
git clone https://github.com/kudouyam/SUIBE-Master-s-Thesis-Skill.git
cd SUIBE-Master-s-Thesis-Skill
python -m pip install -r requirements.txt
python scripts/format_thesis.py "论文原稿.docx" --out "论文_SUIBE规范排版.docx" --title "论文真实题目"
```

输出默认含：
- `论文_SUIBE规范排版.docx`
- `论文_格式修正报告.md`
- `论文_格式修正报告.json`

更多参数：`python scripts/format_thesis.py --help`。复杂表格建议 `--table-mode preserve`；需要保留原有页眉页脚用 `--headers preserve`。

## Skill 目录与适配

| File | Usage |
|---|---|
| [SKILL.md](SKILL.md) | Agent 的执行协议，校规优先级与必要人工复核 |
| [AGENTS.md](AGENTS.md) | Codex/其他 coding agent 的仓库入口 |
| [references/](references/) | 校方规范逐项对照、识别准则、验收检查、GitHub 前辈研究 |
| [scripts/format_thesis.py](scripts/format_thesis.py) | Python 排版和报告生成引擎 |
| [tests/](tests/) | 回归验证 |
| [scripts/build_fixture.py](scripts/build_fixture.py) | 自动生成供回归验证的混乱格式测试论文 |

适合 Codex、Claude Code、GPT Work 等有读仓库及 Python/Word 文件能力的 Agent。支持可编辑 `.docx`，不承诺扫描 PDF、加密 DOC、复杂浮动公式等无损转换。

## 排版规则依据与版本声明

使用《上海对外经贸大学硕士学位论文内容和格式要求（2020 年修订）》与配套硕士论文格式范例作为规则基线。正文宋体/Times New Roman 小四、固定 20pt；一二三级标题依次小三/四号/小四号；A4，上37、下30、左35、右25mm。正式规定、范例及工程建议要分别标注，歧义项保留并写入报告。

**学校原始 PDF 不随仓库再分发**，以免在未确认授权情况下转载校方材料。请使用官方研究生院渠道获取并核对原件，或自行向 Agent 上传最新版校规。仓库的逐项对照说明并不取代原文件。

## 依赖、验证与适用限制

```bash
python -m pip install -r requirements.txt
python -m pip install -r requirements-dev.txt
python -m pytest -q tests/test_formatter.py
```

Word 字段、分节、跨页续表、浮动对象、封面授权签名仍需要人工复核和最终 Word 打印预览。不要删除论文原有内容或将未通过人工验收的文件称作“完全符合毕业要求”。

## Credits

项目设计参考 [Yangruipis/SUIBEthesis](https://github.com/Yangruipis/SUIBEthesis)、[wojiaobadaye/SUIBE_Latex](https://github.com/wojiaobadaye/SUIBE_Latex) 等上外贸开源模板的工程经验。详情见 `references/04_GitHub开源项目逐库审查与迁移矩阵.md`。外部模板的样式不自动替代本项目所依据的 2020 年硕士校规。

This is an **independent, unofficial** community skill; no affiliation with or endorsement from SUIBE is implied.

### Agent 使用边界

仓库已含真正的可编辑 DOCX 排版引擎（**无须另下 ZIP**）。在 Codex / Claude Code / GPT Work 等具备命令执行能力的环境中，克隆后直接运行即可；若宿主不允许访问 GitHub、读写文件或安装 Python 依赖，链接本身无法替代这些权限。复杂论文需要由 Agent 实际打开渲染稿逐页审查，并在报告中声明未完成的审查项目。
