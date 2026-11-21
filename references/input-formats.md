# 读要求：各种输入格式

## 核心原则

作业要求里的格式信息有一半在版式里：字号大小、加粗、行距、表格结构、红色标注、页边批注。纯文本提取会把这些全丢掉，只剩下一堆句子。

所以顺序固定为：**先转图看版式，再提文字抠细节**。

## 图片（.jpg / .png / .webp / 手机截图）

直接用多模态读图，不需要 OCR 脚本。OCR 会把「小四」认成「小4」，把「1.5 倍行距」认成「15 倍行距」，得不偿失。

要点：

- 一张图信息多时，分区看：先整图判断结构，再逐块放大读细节。
- 手机拍的常见问题：透视变形、反光、边缘裁切。看不清就问用户重拍，不要猜。
- 从图里抄录的字体字号，交付时标注「据截图判断」，请用户对一遍。
- 聊天记录里的转发截图，注意上下文可能缺失，缺的部分直接问。

## PDF

先判断是文字版还是扫描版：

```python
import pymupdf
doc = pymupdf.open("要求.pdf")
for i, page in enumerate(doc, 1):
    txt = page.get_text().strip()
    print(i, len(txt), len(page.get_images()))
```

`len(txt) < 50` 且页面有图像，基本就是扫描版。

**文字版**：直接提取，再渲染成图核对版式。

```python
for i, page in enumerate(doc, 1):
    print(f"--- 第 {i} 页 ---")
    print(page.get_text())
    page.get_pixmap(dpi=150).save(f"req_p{i}.png")
```

**扫描版或混合版**：整页渲染成图，按图片处理。

```python
for i, page in enumerate(doc, 1):
    page.get_pixmap(dpi=200).save(f"req_p{i}.png")
```

dpi 用 150–200。再高对读字号没有帮助，只是徒增体积。分辨「小四」和「五号」靠的是对照，不靠像素。

## Word（.docx）

```python
from docx import Document

d = Document("要求.docx")
for p in d.paragraphs:
    if p.text.strip():
        print(p.text)

for t in d.tables:                       # 表格必须单独遍历
    for row in t.rows:
        print(" | ".join(c.text.strip() for c in row.cells))

for s in d.sections:                     # 页眉页脚
    print("页眉:", " ".join(p.text for p in s.header.paragraphs))
    print("页脚:", " ".join(p.text for p in s.footer.paragraphs))
```

三个坑：

1. **`d.paragraphs` 不含表格内文字**。学校把字号规定写在表格里很常见，漏了就等于没读到。
2. **页眉页脚要单独取**，课程名和教师名常在那里。
3. **批注与修订**可能藏着补充要求，查 `word/comments.xml`。

老格式 `.doc`（97–2003）python-docx 打不开，先转：

```bash
soffice --headless --convert-to docx 要求.doc
```

没有 LibreOffice 时提示用户另存为 `.docx`。

## 其他格式

| 格式 | 读法 | 注意 |
|---|---|---|
| Markdown / txt | 直接读 | 无版式，格式靠文字描述 |
| PPT（.pptx） | `python-pptx` 读文本框 | 字号与位置信息在 shape 里，最好再渲染成图 |
| 网页链接 | WebFetch | 登录墙后的页面抓不到，请用户截图 |
| 微信/QQ 截图 | 按图片处理 | 常有上下文缺失 |

## 提取之后

把结果整理成 `references/checklist.md` 里的第一份清单，然后回读一遍。最容易漏的四类：

- 表格里的格式规定
- 脚注与页边批注
- 页眉里的课程名、教师名、提交日期
- 截图边缘被裁掉的部分

读不到或读不准的，明确告诉用户「这一项我需要你确认」，不要用默认值蒙过去——字号和字数恰恰是最容易扣分的地方。
