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
| 标题段前/段后 | 一级 18/6、二级 12/4、三级 6/3 pt | `w:spacing w:before w:after`（单位 1/20 pt） |

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

## 图表

图和表里的文字，字体规则与正文一致（西文 Times New Roman、中文宋体），字号另有一套：

| 位置 | 字号 | 说明 |
|---|---|---|
| 图题、表题 | 五号（10.5 pt）加粗 | 居中；表题在表上方，图题在图下方 |
| 表内文字 | 五号（10.5 pt） | 表头加粗，短文本列居中 |
| 图内轴标签 | 9 pt | 按图宽 8–9 cm 定，出图后不要再缩放 |
| 图内刻度标签 | 8.5 pt | 比轴标签小半磅 |
| 图例 | 8.5 pt | 不加边框（`frameon=False`） |

表格靠后处理逐格设 `rFonts`，含中文的 run 必须带 `hint="eastAsia"`，否则 Word 会用 TNR 去排中文，字宽对不上、单元格被撑开。

插图用 `scripts/plot_style.py`：`font.family` 设成 `["Times New Roman", "SimSun"]`，靠 matplotlib 的逐字符 fallback 做中英分离；`new_axes()` 把 spine 固定为 `('axes', 0)`，不外扩。spine 动了而刻度没动，就是「轴线与刻度线错位」的根因，`audit(ax)` 会在出图前把这类问题报出来。

## 页面

A4（`w:page` 12240 × 16838 twips，即 21 × 29.7 cm），页边距上下 2.54 cm、左右 3.17 cm 是 Word 中文版默认，交作业通常不用改。

## 常见学校变体

遇到下面这些要求，改 `build_docx.py` 顶部的常量或 spec 的 `styles` 即可：

- 标题用「一、二、三」而非阿拉伯数字：`HEAD_NUM` 正则换成 `^([一二三四五六七八九十]+)、\s*`
- 正文要求仿宋_GB2312、三号、行距 28 磅：把 `CN_FONT` 改掉，`line` 换成固定值（`w:line="560" w:lineRule="exact"`）
- 参考文献要求单倍行距、无 DOI：spec 里 `line` 调 1.0；DOI 默认不输出，不必额外处理
- 要求页脚页码：`sectPr` 里加 `footerReference`，再写一个 `PAGE` 域，本脚本未内置
