# 格式参数表

## 字号对照

中文字号与磅值换算，脚本里一律用磅值取整。

| 中文字号 | 磅值 | `w:sz`（半磅） |
|---|---|---|
| 小二 | 18 | 36 |
| 三号 | 16 | 32 |
| 小三 | 15 | 30 |
| 四号 | 14 | 28 |
| 小四 | 12 | 24 |
| 五号 | 10.5 | 21 |
| 小五 | 9 | 18 |

## 段落参数

| 项 | 值 | OOXML |
|---|---|---|
| 1.5 倍行距 | 1.5 | `w:spacing w:line="360" w:lineRule="auto"` |
| 单倍 | 1.0 | `w:line="240"` |
| 正文首行缩进 | 2 字符 | `w:ind w:firstLine="480" w:firstLineChars="200"` |
| 参考文献悬挂缩进 | 21 pt | `w:ind w:left="420" w:hanging="420"` + `w:tabs` 制表位 420 |
| 关闭孤行控制 | — | `w:widowControl w:val="0"` |
| 段前/段后 | 标题 6/4 pt；正文 0 | `w:spacing w:before w:after`（单位 1/20 pt） |

`firstLineChars="200"` 是关键：只给 `firstLine` 是绝对磅值，换字号会错位；带 `Chars` 才按「字符」缩进，Word 里显示为「首行缩进 2 字符」。

## 分页控制

排版最后一公里，两处必须设：

| 位置 | 设置 | 原因 |
|---|---|---|
| 各级标题 | `keep_with_next`（`w:keepNext`） | 否则标题可能孤零零落在页底，正文跑到下一页 |
| 「参考文献」标题 | `keep_with_next` | 防止标题在页底、第一条文献翻页 |

孤行控制（`w:widowControl`）默认开启，按学校要求关闭时逐段设 `w:val="0"`，`docDefaults` 里也设一份，新建段落才继承得到。

## 字体

一套字体要同时写四个属性位，缺一个就可能在某个语种上回落：

```xml
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"
          w:eastAsia="宋体" w:cs="Times New Roman" w:hint="eastAsia"/>
```

- `ascii` 管 U+0000–U+007F，`hAnsi` 管其余拉丁字符（含希腊、西里尔），`eastAsia` 管中日韩，`cs` 管复杂文种。
- `hint` 决定**两可字符**（U+2018–201D 引号、U+2026 省略号、U+2014 破折号）走哪套字体：`eastAsia` 走中文，`default` 走西文。
- 文档默认（`docDefaults/w:rPrDefault`）、`styles.xml` 里每个 `w:rFonts`、每个 run 的 `rPr` 都要改。只改 run 不够——Word 新建段落、插入题注时会取样式里的值。

## 页面

A4（`w:page` 12240 × 16838 twips，即 21 × 29.7 cm），页边距上下 2.54 cm、左右 3.17 cm 是 Word 中文版默认，交作业通常不用改。

## 常见学校变体

遇到下面这些要求，改 `build_docx.py` 顶部的常量或 spec 的 `styles` 即可：

- 标题用「一、二、三」而非阿拉伯数字：`HEAD_NUM` 正则换成 `^([一二三四五六七八九十]+)、\s*`
- 正文要求仿宋_GB2312、三号、行距 28 磅：把 `CN_FONT` 改掉，`line` 换成固定值（`w:line="560" w:lineRule="exact"`）
- 参考文献要求单倍行距、无 DOI：spec 里 `line` 调 1.0；DOI 默认不输出，不必额外处理
- 要求页脚页码：`sectPr` 里加 `footerReference`，再写一个 `PAGE` 域，本脚本未内置
