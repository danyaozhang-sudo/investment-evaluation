# Changelog

## v2.0 (2026-06-06)

### 新增
- **Phase 7: PPT 生成**：双风格 PPT（投行研报 + 路演 Pitch Deck），通过 `/pptx` 技能生成，统一 Midnight Executive 配色
- **Phase 0: 模式确认**：询问用户参与角色（股权/基金GP/EPC）
- **Phase 2: 风电财务测算子技能**：委托 `wind-power-evaluation-claude` 执行完整财务测算
- **Word 报告格式规范**：与 `wind-power-evaluation-claude` 完全对齐（黑体标题 + 仿宋正文 + 深蓝表头 + 七段免责声明 + 江能Logo页眉）
- **程序化视觉 QA 脚本**：PPTX XML 解包 → 对比度/配色/Key Insight 颜色的自动化检查
- **Obsidian 双备份**：所有 4 份文件（.md/.docx/.pptx×2）自动同步到 `wiki/topics/evaluations/{项目名}/`
- **报告合并输出**：不再输出 6 份独立 MD，改为 1 Word + 1 MD + 2 PPT 的统一结构

### 变更
- 报告从 7 章扩展为 11 章（新增电力市场与电价分析[必选A-D]、限电率评估等）
- PPT 生成方式从手写 pptxgenjs 改为调用 `/pptx` 技能+提示词模板
- Python 环境明确为 `/opt/homebrew/bin/python3`（非 sandbox Python）
- 配色方案统一为 Midnight Executive（白底深蓝 #1A3A5C），两风格共享

### 修复
- 三个 PPT 配色铁律固化：冰蓝字仅深蓝背景用、Key Insight 统一深蓝斜体、白字仅深色页用

---

## v1.0 (2026-05)

### 初始版本
- 六类风险评级（市场/技术/政策/现金流/建设/商务）
- 合作方尽调模板
- 三模式支持（股权直投/基金GP/EPC）
- python-docx Word 报告生成
