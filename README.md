# AI-Word-SOP

在 **不破坏 Word 版式** 的前提下，用 **Python（`python-docx`）** 批量改 `.docx` 的可复现流程与代码模板。

## 你在解决什么问题

- `paragraph.text = "..."` 会拆掉 `w:r` / `rPr`，东亚字体、混排、局部加粗经常丢。  
- `Document()` 从零新建再拼段落，默认样式与母版不一致。  
- `pandoc` 转 docx 往往和「你那份已经调好的 Word」不是一回事。

## 正确心法（一句话）

**复制格式母版 → 在副本上改 run 里的字。**

## 仓库结构

| 路径 | 说明 |
|------|------|
| [docs/overview.md](docs/overview.md) | 原理、反模式、操作清单（偏可读） |
| [docs/sop-python-docx-preserve-formatting.md](docs/sop-python-docx-preserve-formatting.md) | 技术细则：replace / cross-run / rewrite / 插入 / 表格 / 自检 |
| [scripts/compare_sop_vs_paragraph_text.py](scripts/compare_sop_vs_paragraph_text.py) | 同一母版两份对照：SOP 改写 vs `paragraph.text` 踩坑 |

## 快速开始

```bash
python3 -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install python-docx

python scripts/compare_sop_vs_paragraph_text.py \
  --template ./your-template.docx \
  --out-dir ./out
```

## 依赖

- Python 3.9+  
- `python-docx`

## 许可

MIT — 见 [LICENSE](LICENSE)。

## 免责声明

本仓库仅提供技术流程与示例代码，不构成法律、财税或合同意见。用于商业文档前请自行复核并与法务确认。
