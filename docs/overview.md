# AI 帮你改 Word，为什么总是一改就「不像人写的」？

适合：经常用 AI 或脚本生成/改写 Word，却被字体、行距、页眉页脚、表格样式折磨的场景。  
工具：`python-docx`（或任何能稳定读写 OOXML 的方案）。

## 根因：Word 不是「一串字」，而是一棵树

`.docx` 内部是 **OOXML**：段落（`w:p`）下面挂着多个 **运行**（`w:r`），每个 run 可有独立的 **`rPr`（run 属性）**——东亚字体、西文字体、字号、加粗、颜色等。

同一段里「前半宋体、后半加粗」往往是 **多个 run**，不是一段纯文本。

## 常见反模式

| 做法 | 常见后果 |
|------|----------|
| `paragraph.text = "..."` | 常销毁 run 结构，`rPr` 丢失，版式漂移 |
| `Document()` 从零新建 | 默认样式，和母版不一致 |
| 只改 `doc.paragraphs` | 表格里文字没改到 |
| `pandoc` 当默认主路径 | 与「已有 Word 母版」视觉不一致 |

## 正确心法（一句话）

**先复制你的格式母版 `.docx`，再在副本上改 run 里的字。**

## 最小可执行清单

1. `shutil.copy(母版.docx, 输出.docx)`  
2. `doc = Document("输出.docx")`  
3. 替换：逐 `run.text`；必要时跨 run 合并替换  
4. 整段重写：`runs[0].text = new_text`，其余 run 清空  
5. 插入新段：`deepcopy` 母版段落 XML，不要默认 `add_paragraph()`  
6. 表格：`doc.tables` 同步遍历  
7. 保存后全文检索旧关键词残留

更细的代码与陷阱表见：[sop-python-docx-preserve-formatting.md](sop-python-docx-preserve-formatting.md)。

## 与「写作质量」分工

本 SOP 只管 **版式结构不被破坏**；语气、事实核对、合规仍由作者与流程负责。
