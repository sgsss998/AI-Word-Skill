# SOP：python-docx 原地修改保格式指南

> 目标：修改现有 `.docx` 的内容（文字替换、段落重写、插入新章节），同时**尽量保留**原文档的字体、颜色、行距、表格样式、页眉页脚等格式。
>
> 反模式：用 `Document()` 新建 → 常见后果是默认字体与段落样式漂移。

---

## 0. 核心心法（一句话）

> **Copy 原档 → 在 copy 上改 run 级别的文字 → 格式自然保留。**

---

## 1. 第一步：复制原档

```python
import shutil
from docx import Document

shutil.copy('原档.docx', '输出档.docx')
doc = Document('输出档.docx')
```

**绝不** `Document()` 从零创建作为主交付路径。

---

## 2. 第二步：文字替换（最常用）

### 2.1 原理：段落 = 多个 run

每个 run 可以有独立格式。**替换文字时必须逐 run 操作**，不能直接 `paragraph.text = ...`（这会销毁所有 run 及其格式）。

### 2.2 单 run 内替换（最安全）

```python
def replace_in_paragraph(paragraph, old_text, new_text):
    """在段落中替换文字，保留 run 格式。"""
    for run in paragraph.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)
            return True
    return False
```

### 2.3 跨 run 替换（进阶）

当“某城市名”被拆成 run[0]="某城" + run[1]="市" 时：

```python
def replace_cross_runs(paragraph, old_text, new_text):
    """跨 run 替换：把涉及的 run 合并后替换。"""
    full_text = ''.join(r.text for r in paragraph.runs)
    if old_text not in full_text:
        return False

    start = full_text.find(old_text)
    end = start + len(old_text)

    char_pos = 0
    affected = []
    for i, run in enumerate(paragraph.runs):
        r_start = char_pos
        r_end = char_pos + len(run.text)
        if r_start < end and r_end > start:
            affected.append(i)
        char_pos = r_end

    if not affected:
        return False

    merged = ''.join(paragraph.runs[i].text for i in affected)
    merged = merged.replace(old_text, new_text, 1)
    paragraph.runs[affected[0]].text = merged
    for i in affected[1:]:
        paragraph.runs[i].text = ''
    return True
```

### 2.4 全文档批量替换

```python
def replace_all(doc, old, new):
    """全文档替换（段落 + 表格单元格）。"""
    count = 0
    for p in doc.paragraphs:
        for run in p.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)
                count += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if old in run.text:
                            run.text = run.text.replace(old, new)
                            count += 1
    return count
```

---

## 3. 第三步：整段重写（保留段落格式）

```python
def rewrite_paragraph(paragraph, new_text):
    """重写段落文字，保留段落和首个 run 的格式。"""
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ''
```

**关键**：保留 `runs[0]`，只改它的 text。清空其余 run。

---

## 4. 第四步：插入新段落（最难的部分）

`doc.add_paragraph()` 常见问题：样式不存在（KeyError）或默认样式与母版不一致。

### 4.1 正确做法：deepcopy 模板段落

```python
from copy import deepcopy
from docx.oxml.ns import qn

template = doc.paragraphs[26]._element  # 选一个版式正确的段落 XML

new_p = deepcopy(template)
for r in new_p.findall(qn('w:r')):
    new_p.remove(r)

new_r = deepcopy(template.find(qn('w:r')))
for t in new_r.findall(qn('w:t')):
    t.text = "你的新内容"
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
new_p.append(new_r)
```

### 4.2 插入顺序

`addprevious` 批量插入可能导致顺序反转；可用 `reversed()` 或 `addnext` 锚点推进。

---

## 5. 第五步：表格单元格替换

同时遍历 `doc.tables`，避免只改正文段落。

---

## 6. 常见陷阱速查

| 陷阱 | 后果 | 正确做法 |
|------|------|----------|
| `doc.add_paragraph(text)` | 格式丢失 | deepcopy 模板段落 XML |
| `paragraph.text = new` | run 结构被毁 | `runs[0].text`，清空其余 |
| 新建 `Document()` 再拼 | 全部格式丢失 | `shutil.copy` 原档再改 |
| 表格只遍历段落 | 表格内文字没替换 | 同时遍历 `doc.tables` |

---

## 7. 验证清单

保存后：全文检索旧关键词；抽查前几段 `runs[0]` 的字体/字号是否与母版一致。

---

## 8. 完整工作流（从复制到保存）

```
原档.docx
    │
    ▼ shutil.copy
输出档.docx（副本）
    │
    ▼ Document('输出档.docx')
    │
    ├── replace_all / rewrite_paragraph
    ├── deepcopy 插入新段
    ├── 表格单元格替换
    │
    ▼ doc.save()
```
