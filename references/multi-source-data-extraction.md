# 多源数据提取参考

从项目文件夹中批量提取数据时使用的命令和常见问题。

## 目录探测

```bash
# 查看项目结构
ls -la "/path/to/project/"

# 查看子项目文件夹
ls "/path/to/project/各子公司项目/"

# 查看合同文件夹
ls "/path/to/project/合同/"
```

## 各格式处理

### DOCX

```python
from docx import Document

# 读取正文
doc = Document('path/to/file.docx')
for p in doc.paragraphs:
    if p.text.strip():
        print(p.text)

# 读取表格
for table in doc.tables:
    for row in table.rows:
        print(' | '.join(cell.text.strip() for cell in row.cells))
```

### XLSX

```python
import openpyxl

# data_only=True 获取计算值而非公式
wb = openpyxl.load_workbook('path/to/file.xlsx', data_only=True)
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if any(v is not None for v in row):
            print(f'Row{i+1}: {[str(v)[:60] if v else "" for v in row]}')
```

**坑**：如果 openpyxl 未安装，需要 `python3 -m pip install openpyxl`

### PDF

```python
# 方法1: PyPDF2
import PyPDF2
with open('file.pdf', 'rb') as f:
    reader = PyPDF2.PdfReader(f)
    for page in reader.pages:
        print(page.extract_text())

# 方法2: pdftotext (需预装)
import subprocess
result = subprocess.run(['pdftotext', 'file.pdf', '-'], capture_output=True, text=True)
print(result.stdout)
```

### 文本文档

```bash
# 直接读取
cat "path/to/file.txt"

# 或通过 python
with open('file.txt', 'r') as f:
    print(f.read())
```

## 关键数据提取清单

从每个源文件中，优先提取以下数据：

| 数据类型 | 来源文件 | 提取方法 |
|---------|---------|---------|
| 项目简介/背景 | DOCX 简介文档 | 全文读取 |
| 备案容量 | 简介文档、备案证 | 提取关键数值 |
| 中标容量 | 中标结果 XLSX | 汇总各公司 |
| 上网电价 | 简介文档、竞价结果 | 提取数值 |
| 年发电小时数 | 简介文档 | 提取数值 |
| 合同条款 | DOCX 合同 | 读取关键章节 |
| 法律审查意见 | 尽调报告 DOCX | 读取结论部分 |
| 统计数据 | XLSX 统计表 | 逐sheet读取 |
| 子项目明细 | 各子文件夹 | 列出清单 |

## 数据一致性检查

当多个源文件包含同一参数时，必须交叉验证：

```python
# 示例：容量一致性检查
data = {
    '简介文档': 183.18,
    '华享中标': 53.48,
    '晟享中标': 31.5,
    '融享中标': 41.9,
    '融汇中标': 29.0,
    '中标合计(不含黄标)': 155.88,  # 53.48 + 31.5 + 41.9 + 29.0
    '中标合计(含黄标)': 185.18,    # 含未中标项目
}
print(f"差距: {data['简介文档'] - data['中标合计(不含黄标)']} MW")
```

**常见容量不一致原因**：
1. 含黄标 vs 不含黄标（未中标项目）
2. 部分项目尚未计入中标统计表
3. 统计口径差异（如含建筑光伏不含道路光伏）

## 合同关键条款提取模板

```python
contract_sections = {
    '合作模式': '第X条',
    '费用分担': '第X条',
    '付款条件': '第X条',
    '违约责任': '第X条',
    '开工条件': '第X条',
    '争议解决': '第X条',
    '终止条款': '第X条',
    '免责条款': '第X条',
}
```
