# 使用示例

## 快速开始

### 1. 安装依赖

```bash
pip install python-docx openpyxl
```

### 2. Word转换

```bash
python3 scripts/md2docx.py sample.md sample.docx
```

### 3. Excel转换

```bash
python3 scripts/md2xlsx.py sample.md sample.xlsx
```

## 示例文件

### basic_example.md

```markdown
# 标题

**加粗** 和 *斜体* 文本

## 二级标题

普通段落文本。

| 列1 | 列2 | 列3 |
|-----|-----|-----|
| 数据1 | 数据2 | 数据3 |

- 列表项1
- 列表项2

> 引用内容

`代码` 内容
```

### 运行示例

```bash
cd examples
python3 ../scripts/md2docx.py basic_example.md basic_example.docx
python3 ../scripts/md2xlsx.py basic_example.md basic_example.xlsx
```

## 高级功能

### 多Sheet Excel

在Markdown中使用`---`分隔不同Sheet：

```markdown
# Sheet 1

表格1...

---

# Sheet 2

表格2...
```

### 富文本

- 加粗：`**文本**`
- 斜体：`*文本*` 或 `_文本_`
- 代码：`` `文本` ``
- 删除线：`~~文本~~`
- 嵌套：`**加粗*斜体*加粗**`

### 智能换行（Excel）

- 包含`\n`的内容：自动换行
- 不含`\n`的内容：不换行

```markdown
| 普通文本 | 多行文本 |
|---------|---------|
| 单行 | 第1行<br>第2行<br>第3行 |
```
