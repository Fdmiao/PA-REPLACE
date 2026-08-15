# Report Plan — 交付物汇总报告 (deliverables-summary)

## Meta
- Type: 项目交付物汇总报告（handover summary）
- Topic: 天融信 NGFW × Palo Alto NGFW 功能对比项目 P0–P6 交付物清单、关键数据、质量门与遗留移交
- Audience: 研发部项目干系人、接手 9 月复验的工程师
- Language: 中文

## Design System (Solid — 企业正式)
- --bg #f5f7fa / --bg2 #ffffff / --ink #17202e / --muted #5b6b7f / --rule #dde3ec / --accent #0f4c81 / --accent2 #b45309
- 字体: 系统中文栈 "Segoe UI", "Microsoft YaHei"; 等宽 JetBrains Mono (仅编号/数字)
- 标题: 粗黑体左对齐, h2 带左侧编号 + 底部 2px accent 边线
- 正文 15px / 1.7; 最大宽 960px 单列; 段距 3rem
- 卡片: 白底 1px rule 边框; 指标卡: 大号 accent 数字 + 小标签
- 表格: 极简行边线, 表头底部 2px
- 提示框: 左侧 4px accent2 边线 + bg2 底

## Structure
1. 项目概览与完成状态 — 定位一句话 + P0–P6 阶段表 + 6 指标卡
2. 交付物清单 — 6 项文件表(大小/内容/用途) + 目录树
3. 主交付 Excel 六张工作表 — Sheet1–6 结构表
4. 关键数据统计 — 图1 大类支持率 / 图2 支持度分布 / 图3 差异定性 + 需求状态表
5. 质量门与评审结果 — 21 项 PASS 清单 + P6 修正记录(22→16)
6. P5 遗留事项移交 — 3 类事项定位索引 + 高优 14 项表
7. 使用与复验指引 — 数据时效 / 9 月复验 / 原厂确认

## Visuals
| Visual | Type | Tool |
|--------|------|------|
| 图1 八大类 TS/PA 完全支持率 | 分组柱状 | ECharts |
| 图2 TS vs PA 支持度分布 | 分组柱状 | ECharts |
| 图3 差异定性分布(139点) | 横向条形 | ECharts |
