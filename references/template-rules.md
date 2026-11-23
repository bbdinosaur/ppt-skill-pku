# 模板继承与版式规则

## 铁律（对应评分：模板约束 + 导出渲染）

1. **母版只读**：任何情况下不修改 slide master、theme、layout。所有改动只发生在 slide 层。
2. **占位符优先**：文字放进版式自带的 title/body 占位符 → 自动继承模板字体、字号层级、
   项目符号、配色。构建脚本不给 run 设置 `font.name`（唯一豁免：无占位符时的兜底文本框，
   也不设字体名，继承主题字体）。
3. **图片不裁不拉**：`add_picture_fit` 等比缩放居中放入槽位，绝不 `insert_picture`（那会裁剪）。
   `aspect` 检查容差 2%。
4. **页脚/页码/日期**：占位符留在 layout 上不动，导出后自动显示。检查器忽略这些区域。
5. **字号下限**：正文默认下限 `11pt`（自动紧缩 `_tighten` 最低 `10.0pt`）、表格最低 `10.0pt`、来源行/图注 `10.5pt–11pt`。低于 `min_font=10.0pt` 会触发 `min_font` 自动钳制或告警。
6. **版式选择与多视觉分槽**：`layout_hint` 由 `pick_layout` 按名称关键词 + 占位符结构打分匹配模板真实版式；模板无两栏版式时自动降级为"单栏占位符 + 右侧文本框"；单页含 ≥2 种视觉（表/图/插图）时自动按 `visual_slots` 垂直或水平切分杜绝重叠。
7. **深浅背景与对比度自适应**：通过只读 XML 判定深色背景（`slide_is_dark`，绝不触发 `python-pptx` 背景覆写），自动对 `highlights` KPI 卡、`steps` 流程图、`takeaway` 结论条、表格头、图表轴标签、来源行执行 WCAG 对比度自适应。

## 文字防溢出

- 每条 bullet ≤2 行（约 ≤40 个全角字符 / ≤80 个拉丁字符）；要点 ≤6 条/页。
- `shrink_body_to_fit` 先收紧段距与上下边距再步进降字号，配合 `run_pipeline.py` 内置 2 轮自动修补（`_tighten=1/2`）确保收敛；超长封面副标题亦自动缩排。
- 标题主张化：一行放得下（≤22 全角字符）；`shrink_title` 兜底防止标题溢出。

## 版式内容分配

| 页型 | 布局 |
|---|---|
| 标题页 / Q&A | title 版式，副标题=作者/场合/日期（超长自动缩排） |
| 方法流程页 | 顶部 `steps` 横向矢量流程图（2–5 步，顶色条+箭头）+ 下方要点或原图 |
| 图/表结果页 | 顶部可选 `highlights` KPI 卡 + 左栏要点 + 右栏表/图/插图 + 页底 `takeaway` 结论条 |
| 两栏对比 | two_content 版式（无则自动降级为左占位符+右文本框） |
| 来源行 | 内容区底部 10.5pt 斜体，逐页列出 ref+locator（深浅背景自动切换字色） |

## 导出与像素自检

```bash
python3 scripts/export_preview.py --deck /app/output/presentation.pptx --work /app/work --pdf /app/output/presentation.pdf --dpi 100
```
自动执行 LibreOffice 隔离 `.lohome` 导出（55s 快失败，自动降级 `reportlab` 真 PDF）→ PyMuPDF 核验页数/字体/标题可抽取性 → 100 DPI 逐页 PNG + `contact-sheet.png` → Pillow 像素空白检测（`visual_ok`）。纯文本模型以程序化检查与 `visual_ok` 为准，严禁读取二进制 PNG。
