# Fallbacks, Exceptions & Script Reference

## 1) 脚本清单与无脚本兜底
所有脚本通过 `Path(__file__).resolve().parent` 自定位，任意工作目录可跑；第三方依赖仅平台预装的 `python-pptx`、`pymupdf`、`Pillow` 与 `soffice`。任一脚本不可用时，按右列兜底执行，不中断交付：

| 步骤 | 脚本 | 无脚本兜底 |
|---|---|---|
| 一站式简报 | `${SKILL_DIR:-/skills/weekly-skill}/scripts/prepare_brief.py --input /input --workdir /app/work` | 依次手动执行 `ls /input` + 读 `case.json`/`task.md` + 下列单步脚本 |
| 模板盘点 | `${SKILL_DIR:-/skills/weekly-skill}/scripts/inspect_template.py TEMPLATE --json ...` | 15 行 `python-pptx` 遍历 `prs.slide_masters` / `slide_layouts` / `placeholders` 记录 idx、type 与几何 |
| 起始 deck 清点 | `${SKILL_DIR:-/skills/weekly-skill}/scripts/inspect_deck.py OLD.pptx --json ... --draft ... --imgdir ...` | `python-pptx` 递归遍历 slides/shapes/tables/groups 提取文本、图片与 notes |
| 论文抽取 | `${SKILL_DIR:-/skills/weekly-skill}/scripts/extract_paper.py PAPER.pdf --out ... --figdir ...` | `pdftotext -layout` 读带页文本；`pymupdf` `page.get_pixmap(clip=..., dpi=200)` 手动裁图 |
| 构建 deck | `${SKILL_DIR:-/skills/weekly-skill}/scripts/build_deck.py plan.json --out build_report.json` | 打开模板 `Presentation(tpl)`，用模板自带 `add_slide(layout)` 填充真实占位符，不覆盖主题字体 |
| 静态检查 | `${SKILL_DIR:-/skills/weekly-skill}/scripts/check_deck.py deck.pptx --json deck_check.json` | 遍历每页 shapes 检查越界、2D 重叠、纵横比、正文 ≥12pt 与 notes，并目测渲染 PNG |
| 导出渲染 | `${SKILL_DIR:-/skills/weekly-skill}/scripts/export_render.py deck.pptx --outdir /app/output/previews --pdf-out /app/output/deck.pdf --dpi 108` | `soffice --headless --convert-to pdf --outdir /app/output /app/output/deck.pptx` + `pdftoppm -png -r 108 /app/output/deck.pdf /app/output/previews/page` |
| 合成结果 | `${SKILL_DIR:-/skills/weekly-skill}/scripts/write_result.py --case-json /input/case.json ... --out /app/output/result.json` | 按 `references/result-schema.md` 手写合法 JSON（逐字保留 `case_id` 与全部 `output_contract` 字段，禁编造） |

## 2) 异常处置矩阵（必须记入 `change_log` / `uncertainty` / `limits`，禁静默吞掉）
| 情形 | 处置 |
|---|---|
| 无模板文件 | `plan.json` 设 `"template": "none"` 用默认母版建 deck；`uncertainty` 注明无模板输入 |
| `.potx` 模板 | `inspect_template.py` / `build_deck.py` 自动归一化 content-type 生成标准 `.pptx`，记入 `normalizations` 与 `change_log` |
| 模板自带示例页 | `build_deck.py` 默认清除预置页防止页号错位；确需保留时在 `plan.json` 设 `"keep_template_slides": true` 并在 `uncertainty` 注明偏移 |
| 提供起始 deck（旧 PPT） | 运行 `inspect_deck.py --draft` 清点 `master_fingerprint`：同母版优先保留原版式映射，异母版按目标模板重建；保留旧表/浮动框/图片内容并补齐论文证据，记 `change_log` |
| 固定页数但内容过密 | `plan.json` 设 `"expected_slides": N` 锁页防自动拆页超标；精简句长、改用 `two_col` 或紧凑原生表，正文不低于 13pt |
| 版式缺 body 占位符 | `build_deck.py` 自动落手动文本框并在 `build_report` 记 `fallbacks_manual_textbox`，`uncertainty` 说明该页非占位符继承 |
| 标题占位符缺失 | 标题写入 notes 并在报告中记 warning |
| 图片未自动裁出或清晰度低 | 查 `/app/work/paper_text.txt` 页码，用 `pymupdf` 按页精确 `clip` 重裁（200 DPI）；仍无法获取则在 `uncertainty` 声明 |
| 字体回退 / 导出差异 | `export_render.py` 对比模板 `theme_fonts` 与 PDF 嵌入字体 `pdf_fonts_used`；若 `title_text_ok=false` 或出现替换框，改 ASCII 符号或排查字体并记 `font_warnings` |
| `soffice` 导出失败 | 重试 1 次；仍败则交付 `deck.pptx` + `render_log(export-failed)`，`status: "partial"` 并在 `limits` 写明原因与复现命令 |
| `case.json` 标签与实文件冲突 | 以 `/input/` 真实文件内容为准，不按字段名臆断；差异记入 `uncertainty` |
