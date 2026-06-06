# Word报告生成指南

使用 `python-docx` 生成投资评估 Word 报告。**格式标准与 `wind-power-evaluation-claude` 完全一致。**

## 前提

```bash
# 使用系统 Python（sandbox Python 无 python-docx）
/opt/homebrew/bin/python3 -c "from docx import Document; print('OK')"
```

## 页面布局

| 参数 | 数值 |
|:------|:----:|
| 纸张 | A4 (21cm x 29.7cm) |
| 上边距 | 2.5cm |
| 下边距 | 2.2cm |
| 左边距 | 2.5cm |
| 右边距 | 2.5cm |

```python
from docx.shared import Cm
for section in doc.sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
```

## 字体层级

| 层级 | 字体 | 字号 | 颜色 | 加粗 |
|:----|:----|:----:|:----:|:----:|
| H1 | **黑体** | 15pt | #1B3A5C | ✅ |
| H2 | **黑体** | 13pt | #1B3A5C | ✅ |
| H3 | **黑体** | 11pt | #1B3A5C | ✅ |
| 正文 | **仿宋** | 11pt | 黑色 | ❌ |
| 表格表头 | **黑体** | 9pt | #FFFFFF / 背景#1B3A5C | ✅ |
| 表格内容 | **仿宋** | 9pt | 黑色 | ❌ |
| 代码块 | Courier New | 9pt | 黑色 | ❌ |

## 核心函数

```python
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

COLOR_DB = RGBColor(0x1B, 0x3A, 0x5C)

def h1(doc, text):
    h = doc.add_heading(text, level=1)
    for r in h.runs:
        r.font.color.rgb = COLOR_DB; r.font.name = '黑体'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); r.font.size = Pt(15)

def h2(doc, text):
    h = doc.add_heading(text, level=2)
    for r in h.runs:
        r.font.color.rgb = COLOR_DB; r.font.name = '黑体'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); r.font.size = Pt(13)

def h3(doc, text):
    h = doc.add_heading(text, level=3)
    for r in h.runs:
        r.font.color.rgb = COLOR_DB; r.font.name = '黑体'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); r.font.size = Pt(11)

def para(doc, text, bold=False, sz=11):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = '仿宋'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    r.font.size = Pt(sz); r.bold = bold

def code_para(doc, text, sz=9):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = 'Courier New'; r.font.size = Pt(sz)

def table_with_header(doc, headers, rows):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = 'Table Grid'
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = h
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True; r.font.size = Pt(9); r.font.name = '黑体'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        # 表头深蓝背景
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="1B3A5C"/>')
        c._element.get_or_add_tcPr().append(shading)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = t.rows[ri+1].cells[ci]; c.text = str(val)
            for p in c.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.size = Pt(9); r.font.name = '仿宋'
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    return t
```

## 输出结构

Word报告为7章，对应 report-template.md 的7个章节：
1. 执行摘要（含关键指标表格）
2. 项目概况（基本信息表 + 项目公司分布表）
3. 合作架构与合同保护（角色表 + 保护维度表）
4. 财务评估（利润测算表 + 垫资说明）
5. 风险评估（评分表 + 重点风险）
6. 合作方评估（维度评分表）
7. 综合建议与行动清单（参与条件表 + 行动清单表）

## 免责声明（七段，固化模板）

```
本报告由江能研究院（基于 Claude Opus 大语言模型辅助）生成，仅供项目投资决策参考，不构成任何形式的投资建议或承诺。

报告中所引用的电力市场数据均来自公开渠道，部分电价参数在缺乏直接交易数据的情况下采用保守假设推算，可能与实际成交价格存在偏差。

实际项目投资决策应结合以下因素综合判断：
① 场址实测测风数据（至少一个完整年度）；② 电网接入批复及送出工程条件；③ EPC 招标实际报价；④ 项目所在地最新的机制电价竞价结果；⑤ 所在省电力交易中心公布的全年现货与中长期交易结算数据。

本报告中的财务模型基于特定假设（融资利率 4%、经营期 20 年、等额本金还款等），不同融资结构、利率环境及政策变化可能导致测算结果显著偏离。

报告中的"投资边界"和"出售边界"为理论测算阈值，不代表项目实际可实现的交易价格或融资条件。

江能能源及报告编制方不对因使用本报告而产生的任何直接或间接损失承担责任。未经授权，不得转载或用于商业用途。
```

## 江能 Logo 页眉

```python
LOGO = os.path.expanduser("~/.openclaw/workspace/skills/wind-power-analysis/assets/jianeng_logo_header.png")
for section in doc.sections:
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run()
    run.add_picture(LOGO, width=Cm(2))
```

## 常见问题

- **中文字体**：务必设置 `rFonts.set(qn('w:eastAsia'), '字体名')`，否则中文显示为方块
- **封面**：用空行 + 居中对齐实现
- **分页**：每章结束后 `doc.add_page_break()`
- **保存路径**：放在 `投资评估报告/` 子目录
- **运行环境**：必须用 `/opt/homebrew/bin/python3`，不是 sandbox Python
