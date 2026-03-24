---
name: md2docx-python
description: Markdown转Word/Excel文档的Python工具。支持多Sheet、文字表格交替、完整格式解析。
version: 1.2.0
dependencies: [python-docx, openpyxl]
---

# md2docx-python

## 决策：直接调用 vs 复制修改

**直接调用**：标准Markdown转换，无需定制
```bash
python3 scripts/md2docx.py input.md output.docx
python3 scripts/md2xlsx.py input.md output.xlsx
```

**复制修改**：需要自定义样式、页眉页脚、目录、公式等定制功能

## Excel特性

- 多Sheet：一级标题创建新Sheet
- 格式解析：自动清理Markdown标记（加粗、斜体、链接、代码等）
- 单元格换行：`<br>` → 实际换行
- 数字格式：千分位、百分比自动识别
- HTML实体：`&lt;` `&gt;` `&amp;` `&nbsp;` 正确转换

## Word特性

- 支持标题、段落、表格、列表
- 中文兼容性好
- 字体：微软雅黑

## 命令行参数

**md2docx.py**: `--font` `--size` `--no-page-break`

**md2xlsx.py**: `--no-multi-sheet` `--header-color` `--font` `--no-auto-width`

## 文件结构

```
scripts/
├── md2docx.py        # Markdown转Word
├── md2xlsx.py        # Markdown转Excel
└── report_template.py # Word模板
```
