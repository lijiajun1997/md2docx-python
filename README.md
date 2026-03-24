# md2docx-python

Markdown 转 Word/Excel 的 Python 工具。支持多 Sheet、文字表格交替、完整格式解析、富文本（加粗、斜体、代码、删除线）。

## 功能特性

### Word 转换 (md2docx.py)
- ✓ 完整 Markdown 解析
- ✓ 表格、列表、引用块、代码块
- ✓ 富文本：加粗、斜体、代码、删除线、嵌套格式
- ✓ 中文引号、特殊字符支持
- ✓ 标题、段落、列表格式化
- ✓ 页眉页脚、页码、封面支持

### Excel 转换 (md2xlsx.py)
- ✓ 多 Sheet 支持
- ✓ 表格单元格内富文本
- ✓ 自动换行策略（含`\n`的内容自动换行，不含则不换行）
- ✓ 数字格式化（百分比、千分位）
- ✓ 浅灰色表头
- ✓ 首列窄宽度（3.5字符），内容从第2列开始
- ✓ 微软雅黑字体统一

## 安装

```bash
pip install python-docx openpyxl
```

## 使用方法

### Markdown 转 Word

```bash
python3 scripts/md2docx.py input.md output.docx
```

### Markdown 转 Excel

```bash
python3 scripts/md2xlsx.py input.md output.xlsx
```

### 多 Sheet Excel

```bash
python3 scripts/md2xlsx.py input.md output.xlsx
```

在 Markdown 中使用 `---` 分隔不同 Sheet。

## 示例

### Markdown 输入

```markdown
# 标题

**加粗** 和 *斜体* 文本

## 二级标题

| 列1 | 列2 | 列3 |
|-----|-----|-----|
| 数据1 | 数据2 | 数据3 |

- 列表项1
- 列表项2

> 引用内容

`代码` 内容
```

### Word 输出

- 标题层级自动格式化
- 表格自动转换为 Word 表格
- 富文本正确显示

### Excel 输出

- 表头浅灰色，黑色加粗
- 首列宽度 3.5，内容从第2列开始
- 包含换行符的单元格自动换行
- 不含换行符的单元格不换行

## 版本

- v1.3.1: XML错误修复，富文本支持
- v1.3.0: 富文本解析（Word），表格富文本（Excel）
- v1.2.0: 多Sheet支持
- v1.1.0: 基础功能

## 许可证

MIT License
