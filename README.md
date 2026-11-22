# 杰哥写作 · jiege-writing

把中文课程论文从作业要求做到可提交的 Word 成稿：读要求、核验文献、写正文、排版、自检。

按 [Agent Skills](https://agentskills.io) 开放标准组织，同一份目录可直接用于 Claude Code、OpenAI Codex、Gemini CLI、DeepSeek Harness、Cursor、VS Code Copilot 等能读取 `SKILL.md` 的客户端。

> **适用范围**
> 主要面向中文**课程论文、读书报告、文献综述、方法学评述**，篇幅 3000–8000 字，一二级标题结构，10–30 条参考文献。
> **不适用于学位论文**。目录、页码、页眉页脚都没有实现。写毕业论文请先把 `SKILL.md` 末尾「已知限制」补齐。

> **参考文献请务必自行复核**
> 脚本给的是通行默认值，不是标准答案。各校模板常在这三处不一致：
> - GB/T 7714 用 **2015 版**还是 **2025 版**（西文姓氏全大写 vs 只首字母大写）
> - 中文条目用**全角标点**（`，．：（）`）还是**半角标点**
> - **列不列 DOI**
>
> 交稿前请逐条对照学校给的示例。技能里的 `references/gbt7714.md` 讲了怎么改、改哪里。

> **使用须知**
> 这是写作与排版工具，替代不了你自己读文献、判断和论证。请遵守所在学校或期刊的学术规范，**不得用于代写、代改或任何形式的学术不端**。
> `examples/` 里的示例论文只用于演示格式规范，**请勿直接提交**——它已随仓库公开发布，查重系统可能已经收录。

## 它解决什么

写一篇中文课程论文，真正耗时的不是想观点，是这四件事：

| 痛点 | 这个技能怎么做 |
|---|---|
| 作业要求散在 PDF、截图、群消息里，字号行距容易漏 | 先渲染成图看版式，再提文本抠细节；拆成可勾选清单 |
| 参考文献凭印象写，卷期页错、作者数不合规 | 强制走 DOI 回查，`format_ref()` 按语种切标点、自动截断作者、区分六种文献类型 |
| Word 排版反复调，字体混入 Arial、上标大小不一、孤行没关 | 稿件与排版分离，`build_docx.py` 一次生成；`check_docx.py` 读回校验九项 |
| 图表字体不统一、坐标轴与刻度线错位 | `plot_style.py` 把字体、线宽、轴定位定死；`audit()` 出图前自检四类问题 |
| 机制图要从文献里找，图与图注总被拆到两页 | `fetch_pmc_figs.py` 取开放获取原图；`insert_figs.py` 用无边框表格把图与图注锁在一页，编号自动对齐正文 |
| 改八轮之后格式全乱，说不清哪次引入的 | spec 与脚本分离，改文字不动格式；每轮报数字指标 |

## 安装

仓库里技能本体在 `skills/jiege-writing/`，整目录拷到对应位置即可。

```bash
git clone https://github.com/mio929/jiege-writing.git
```

**Claude Code**

```bash
cp -r jiege-writing/skills/jiege-writing ~/.claude/skills/
# 项目内共享则放到 .claude/skills/
```

**OpenAI Codex**

```bash
cp -r jiege-writing/skills/jiege-writing ~/.agents/skills/
# 也认 ~/.codex/skills/，项目级用 .agents/skills/
```

**Gemini CLI**

```bash
gemini skills install https://github.com/mio929/jiege-writing.git --path skills/jiege-writing
# 或手动拷到 ~/.gemini/skills/
```

**手动安装（通用）**

```bash
cp -r jiege-writing/skills/jiege-writing ~/.dsh/skills/       # DeepSeek Harness
cp -r jiege-writing/skills/jiege-writing ~/.agents/skills/    # 其他读取 SKILL.md 的 harness
```

**Cursor / VS Code Copilot**

Cursor 读 `.agents/skills/` 与 `~/.cursor/skills/`；Copilot 读 `~/.copilot/skills/` 与 `.github/skills/`。两者也都兼容 `.claude/skills/`。

装完重启会话。Claude Code 用 `/jiege-writing` 显式调用，其余平台靠描述自动匹配。

## 依赖

Python 3.9+，两个包：

```bash
pip install python-docx lxml
```

读 PDF 要求时另需 `pymupdf`。Word 与图片不需要额外依赖——图片走多模态直接读，不走 OCR。

## 快速开始

```bash
python skills/jiege-writing/scripts/build_docx.py \
       skills/jiege-writing/assets/example_spec.json -o 试排.docx

python skills/jiege-writing/scripts/check_docx.py 试排.docx --strict
```

试排产出是一份三段的演示文档，用来确认字体、域、编号都正常。确认无误后改 spec 写自己的。

## 目录结构

```
skills/jiege-writing/
├── SKILL.md                     # 主流程：六步工作法、格式硬指标、已知限制
├── references/
│   ├── input-formats.md         # 从 PDF / Word / 图片里读作业要求
│   ├── checklist.md             # 开工前与交稿前的双份清单
│   ├── case-study.md            # 一篇 4600 字论文的完整复盘，含表格补丁写法
│   ├── writing-guide.md         # 结构编排、篇幅配比、语体与去 AI 味
│   ├── format-spec.md           # 字号字体参数表与 OOXML 键值
│   ├── gbt7714.md               # 六类文献格式、核验流程、常见错误
│   ├── cross-reference.md       # 交叉引用的域机制与手工步骤
│   ├── figures.md               # 插图与图注：三条来路、取图、尺寸、四个故障
│   └── docx-pitfalls.md         # 踩坑清单：现象 → 成因 → 改法
├── scripts/
│   ├── build_docx.py            # 稿件 JSON → 规范 docx
│   ├── check_docx.py            # 交付前自检
│   ├── insert_figs.py           # 图占位 → 表格包裹的图 + 图注，编号对齐正文
│   ├── fetch_pmc_figs.py        # 从 PMC 开放获取论文取原图与图题
│   └── plot_style.py            # 数据图样式：中英字体分离、轴与刻度对齐
└── assets/
    └── example_spec.json        # 最小可运行模板

examples/
└── sample-paper.md              # 完整示例：一篇 4600 字方法学评述的成品全文
```

## 示例

`examples/sample-paper.md` 是一篇实际写出来的论文全文：《农田土壤水肥盐一体化智能调控方法评述》。4637 字正文、6 个一级章节、13 个二级小节、17 条参考文献（全部 DOI 复核）、2 个三线表。

这篇照着技能的六步流程走完，包括两次字数调整、一次节数拆分和一轮格式自检。过程记录在 `references/case-study.md`。

**顺带提醒一个容易被忽略的点**：这篇正文 4637 字，但 Word 字数统计是 **5338** —— 差的 701 字来自标题、摘要和表格，学校查字数时这些都算在内。写作前务必问清按哪个口径。

字数可以在 spec 里声明，构建时自动提示越界：

```json
"word_target": [4500, 5000]
```

自检时按 Word 口径复核：

```bash
python scripts/check_docx.py 论文.docx --target 4500-5000 --by word
```

注意这个示例**不含表格**——表格要靠后处理补，做法在 `references/case-study.md` 第 4 节。

## 已知限制

| 限制 | 说明 |
|---|---|
| 表格 | spec 无 table block；`check_docx.py` 对表格内段落是盲区 |
| 图 | spec 无 figure block；用 `insert_figs.py` 后处理，图注不做交叉引用域 |
| 页码 / 目录 | 未实现 |
| 页边距 | 写死上下 2.54、左右 3.17 cm |

标题层级支持到三级（`1` / `2.1` / `2.1.1`）。**但默认写两级**——课程论文的体量两级够用，多一层会把结构切碎。只有你明确要求，或单个二级小节确实超过 800 字且内部有并列子主题时，才会用到三级。

这些都在 `SKILL.md` 的「已知限制」里给了绕法。遇到时技能会先告诉你，不会硬扛。

## 说明

示例论文为虚构选题下的一次完整实跑，不含任何个人身份信息。参考文献均为公开发表的真实文献。

## License

MIT
