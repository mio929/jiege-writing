# 杰哥写作

中文学术课程论文、文献综述、研究方案类 Word 文档的全流程写作与排版技能，用于 WorkBuddy。

从作业要求到可提交的成稿一次走完：拆解要求、核验文献、撰写正文、生成格式化 docx、交叉引用、交付自检。

## 内容

| 文件 | 作用 |
|---|---|
| `SKILL.md` | 主流程：六步工作法、格式硬指标、多轮修订规矩 |
| `scripts/build_docx.py` | 稿件 JSON → 规范 docx，含自动编号、交叉引用、字体处理 |
| `scripts/check_docx.py` | 交付前自检，读回 docx 校验字体、字号、域、孤行控制、字数 |
| `assets/example_spec.json` | 稿件模板，演示 spec 结构与引用标记 |
| `references/format-spec.md` | 字号字体参数表与 OOXML 键值、分页控制 |
| `references/cross-reference.md` | 交叉引用的域机制、手工步骤、上标与排错 |
| `references/gbt7714.md` | GB/T 7714-2025 著录规则、核验流程、常见错误 |
| `references/docx-pitfalls.md` | 排版踩坑清单：现象 → 成因 → 改法 |
| `references/writing-guide.md` | 结构编排、篇幅配比、语体与去 AI 味 |

## 用法

```bash
python scripts/build_docx.py spec.json -o 论文.docx
python scripts/check_docx.py 论文.docx --strict
```

稿件写在 `spec.json` 里，正文用 `{{文献键}}` 标引用，`refs` 给结构化字段，脚本按 GB/T 7714-2025 拼装并自动编号。详细说明见 `SKILL.md`。

## 安装

把 `jiege-writing/` 整个目录放进 `~/.workbuddy/skills/`（用户级）或项目的 `.workbuddy/skills/`（项目级）。

## 依赖

Python 3.9+，`python-docx`、`lxml`。

## 说明

示例稿件为占位内容（作者「张三」），不含任何真实论文正文。参考文献示例均为公开发表的文献。
