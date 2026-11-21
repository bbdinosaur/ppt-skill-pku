# 异常与降级手册（exceptions 处置对照表）

每个异常必须进入 `result.json.exceptions`：`{code, detail, disposition}`。诚实 > 完美。

| code | 触发情形 | 处置 |
|---|---|---|
| MISSING_INPUT | case.json / workorder / material 缺失 | 按 task.md 默认约束生成；exceptions 记录缺口 |
| TEMPLATE_MISSING | 无 target-template.pptx | 用默认模板生成 + exceptions 注明"未继承指定模板" |
| BITMAP_FIGURE | 图为位图且无数据表/矢量数据 | 只引用原图 + 图注，不重绘曲线；uncertainty 记录 |
| FIGURE_MISSING | 材料引用的图片文件不存在 | 页面放"[图缺失] + ref + 图注"文本占位，不画假图 |
| SLOT_SHORTAGE | 模板缺所需版式（如两栏） | 降级为占位符+文本框组合，plan 注明降级 |
| PAGE_COUNT_MISMATCH | 页数 ≠ workorder 要求 | 优先满足 workorder；无法满足时 exceptions 说明 |
| OVERFLOW / OVERLAP / OUT_OF_BOUNDS | 检查器发现 | 精简文字或拆页后重建；修复 ≤2 轮，仍失败则如实报告 |
| ASPECT_CHANGED | 图片被拉伸 | 必须修复（重建该页），这是硬错误 |
| QUOTE_NOT_FOUND | claim_sources.quote 不在材料原文 | 删除或改写为真实原文；绝不保留编造引文 |
| NUMBER_NOT_FOUND | 页面实质性数字（小数/≥3位）不在材料原文 | 改用材料中存在的数字或删去该数字；绝不按印象取整 |
| TABLE_OVERFLOW | 表格行数×换行高度超出框 | 减行/精简单元格文本/加大框；仍失败则拆页 |
| FONT_OVERRIDE | run 显式设置了字体名 | 移除 font.name 让模板继承 |
| TABLE_OVERFLOW | 表格过高 | 减行或精简文本后重建 |
| EXPORT_FALLBACK | 无 soffice，PDF 导不出 | 保留 PIL 几何预览 PNG；exceptions 说明 PDF 缺失 |
| NOTE_MISSING | 页面缺讲稿备注 | 补 2–4 句备注后重建 |

| SCANNED_PDF | PDF 无文本层（扫描版），material_text.md 为空 | 以 material.json 为事实依据；uncertainty 记录；勿凭图猜数字 |

## 环境降级

- **python-pptx 缺失**：`pip install python-pptx`（无网络则报 BUILD_FAILED 并把计划写入
  plan.json 留档，status=partial，不得伪造 pptx）。
- **PIL 缺失**：图片插入退化为 width-only（仍等比），aspect 检查自动跳过。
- **PyMuPDF 缺失**：跳过 PDF 文本抽取，quote 核验退化为对照 material.json 原文。
- **soffice 缺失/超时(>240s)**：reportlab 渲染**真实 PDF 交付物**（近似版式、可抽取文本、
  含图片表格）+ PIL 几何预览兜底 —— PDF 交付物永不缺失。

## 判定原则

- 材料里没有的数字/结论 **不写**。宁缺毋编。
- `status` 判定：deck 存在且关键检查（bounds/aspect/page_count/claim_sources）全过 + PDF
  导出成功 → `ok`；deck 存在但有失败项 → `partial`；无 deck → `failed`。
- 部分结果也要早保存：先落一版 partial 的 result.json，再迭代。
