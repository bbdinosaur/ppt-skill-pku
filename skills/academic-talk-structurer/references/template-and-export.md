# 模板/母版合规与导出

## 1. 布局选择（build_deck.py 自动 + hint 覆盖）
- 角色→布局名关键词映射内置（title→Title Slide，result→Picture with Caption 等）；中文模板名同样命中（标题和内容/两块内容/仅标题/空白）。
- `layout_hint` 精确匹配布局名（忽略大小写）优先；找不到回退角色映射，再回退第一个含 content 的布局。
- **始终用 layout 的 placeholder 承载标题/正文**（继承母版字体、字号、项目符号），不自造满屏 textbox；仅 claim 条与图注允许新 textbox。

## 2. 配色与主张条
- 主张条文字色自动取模板主题 accent1（暗化 62% 保证对比度）；主题色不可读时回退深灰。
- 正文/标题颜色一律继承母版，页面仅主张条与图注允许独立配色。

## 3. 字体合规
- 不在脚本里硬编码字体名；正文/标题字体一律继承模板。
- 若 case 明确规定字体而模板不符：以模板为准并在 uncertainty 注明冲突，勿手工改母版（易破坏布局）。
- 中文缺字（导出后豆腐块）：检查模板字体是否含 CJK；若 soffice 导出缺失，改用模板中已有 CJK 字体占位并记录。

## 4. 起始 deck（edit 模式）守则
- **原位微调**（默认 `meta.keep_only_spec_slides: false`）：只改明确要求的页（`reuse.slide_index` 0-based 指定），其余旧页保留原样并自动纳入 `slide_plan.json` 保持页数一致。
- **精简与重排**（`meta.keep_only_spec_slides: true`）：当任务要求把冗长草稿（如 16 页）改造成指定页数/受众汇报（如 10 页）时，按 `spec.slides` 顺序保留/改写所需旧页、删除未列出的冗余旧页，并在 `meta.tradeoffs` 与 `result.changes` 写明删减与重排依据。
- `reuse.action`: `"keep"` 完全保留原页版式与图形（仅更新讲稿备注/页底来源）；`"edit"`（默认）改写标题/要点/主张条（旧主张条自动替换为新版），若传入新 `figure`/`table` 或设 `clear_visuals: true`，自动清除该区域旧图/旧表避免双层重叠。

## 5. 深浅主题与对齐合规
- `build_deck.py` 自动检测母版/布局深浅底色与 `<a:schemeClr>`：深色底板自动切换高对比浅色页脚、图注、卡片与暗底原生表格；高亮度主题色自动切换深色表头字色保证对比度。
- 非复用内容页自动左对齐标题、清理布局残留空占位符（避免双栏布局空框），并自适应缩放表格/KPI/步骤卡片字号防止溢出。

## 6. 导出链路
```sh
soffice --headless -env:UserInstallation=file:///tmp/lo_profile --convert-to pdf --outdir <dir> <pptx>
pdftoppm -png -r 100 <pdf> <dir>/preview
```
- 失败重试一次；再失败：交付 PPTX + lint 报告，result.rendering 标注 failed 与原因（字体/媒体/超时）。
- 导出后必须目检 PNG（≤108 DPI，优先读 2×3 接触表）：越界截断、文字重叠、图片拉伸、对比度不足逐页确认。
- 大 deck 超时：先导前 12 页验证样式，再全量导出。

## 7. 交付前 checklist
1. PPTX 可被 python-pptx 重新打开（无损编辑）。
2. 每页 notes 非空且含 time_s 标注。
3. lint 0 issues 或残留问题已写入 limits。
4. PDF/PNG 与 PPTX 同名同目录；rendering.pdf_pages == PPTX 页数。
