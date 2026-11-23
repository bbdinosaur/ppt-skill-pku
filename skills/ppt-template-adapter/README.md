# Academic Slide Template Adaptor（学术 PPT 模板与格式适配 Skill）

将论文 PDF（或科研材料）+ 指定 PPTX/POTX 模板（及可选起始旧 PPT），稳定转化为 **母版/版式真实继承、版面专业美观、逐页证据可追溯、2D 防重叠校验且导出可复核** 的可编辑学术演示文稿。

## 完整调用步骤（默认 `${SKILL_DIR:-/skills/weekly-skill}`，挂载在其他路径可先设 `export SKILL_DIR=...`）
```bash
# 1. Step 1 一站式盘点与抽取简报（≤32 KB 防关键实验段截断，自动落盘全部中间 JSON 与裁图）
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/prepare_brief.py" --input /input --workdir /app/work

# （可选单步调用：如需显式指定单文件）
# python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/inspect_template.py" template.pptx --json /app/work/template_inventory.json
# python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/extract_paper.py" paper.pdf --out /app/work/paper_structure.json --figdir /app/work/figs
# python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/inspect_deck.py" old_deck.pptx --json /app/work/deck_inventory.json --draft /app/work/plan_draft.json --imgdir /app/work/deck_imgs

# 2. 按 references/template-mapping.md 编写 /app/work/plan.json
#    支持 title / section / content(+table/chart) / two_col / figure(side_by_side|stacked) 与 **加粗导语**

# 3. 一键构建 + 静态 QA + 108 DPI 渲染导出 + 合成 result.json（3–5 轮闭环）
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/build_deck.py" /app/work/plan.json --out /app/work/build_report.json && \
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/check_deck.py" /app/output/deck.pptx --json /app/work/deck_check.json && \
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/export_render.py" /app/output/deck.pptx --outdir /app/output/previews --pdf-out /app/output/deck.pdf --dpi 108 && \
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/write_result.py" --case-json /input/case.json --plan /app/work/plan.json \
    --build /app/work/build_report.json --check /app/work/deck_check.json \
    --render /app/output/previews/render_log.json --deck /app/output/deck.pptx --out /app/output/result.json
```
详细执行策略、异常处置与预算纪律见 `SKILL.md`。

## 核心方法与相对基础方案的改进
1. **真实母版继承 + 自动清理残影**：严格按几何与占位符类型选版式（分离 `TEXT` 与 `PICTURE` 占位符防错填），不覆盖主题字体，清理未使用的空占位符框防遮挡，并输出逐页 `placeholder_chain` 与 `inheritance_summary`。
2. **学术多样版式、原生可编辑表与原生数据图表**：除标准要点页外，原生支持 `**加粗导语**`、多行封面主次副标题、对称 `two_col`（含 4 占位符 `Comparison` 配对）、`figure` 自动左文右图（自动裁白边 + 深底透明暗线图白卡保护），以及随模板 `accent1..6` 自动配色的原生可编辑 PPTX 表格与原生柱/条/折线图表（`chart`）；并自动避让母版边缘装饰条/校徽横幅。
3. **双栏、罗马/附录与多格式材料抽取 + 起始旧 PPT 无损迁入**：`extract_paper.py` 支持双栏跨栏全宽图、上置图注反裁、LaTeX `booktabs` 无竖线表与 `.md/.txt/.docx/.csv`；`inspect_deck.py` 完整提取占位符内外的表格、图片、原生图表、缩进层级与双栏结构，自动生成可编辑迁移草稿。
4. **2D 重叠与批量一致性闭环**：`build_deck.py` 锁页守卫 + 0 高度占位符自愈 + 折行表高精确计算；`check_deck.py` 覆盖文本/表/图/原生图表 2D 重叠、字号层级、模板残字与图号一致校验（并支持可选只读 `--paper-text` 数值参考核对）。
5. **孤立 Profile 自动重试与字体替代分级审计**：`export_render.py` 以独立临时 LibreOffice profile + 自动重试导出 108 DPI 预览，核验 PDF 页面尺寸一致性、度量兼容字体替代分级（`font_substitutions`）与深浅底通用空白页检测。
6. **一站式简报与精准状态语义**：`prepare_brief.py` 单命令输出模板槽位 + 论文图表/数值表 + 优先保护实验/结果/局限段的 ≤32 KB 简报；`write_result.py` 将核心交付完整性与启发式几何告警解耦（核心产物齐备且页数匹配即保持 `status="complete"`，告警如实记入 `checks`/`limits`/`uncertainty`），并同步结构化 `evidence`（`locator` + `kind` + `quote`）至 speaker notes 与 `evidence_map`。

## 依赖与环境
- Python 3.8+；第三方依赖仅 `python-pptx`、`pymupdf`、`Pillow`（均预装于平台 `scientific-python` 环境）；PDF 导出依赖 `soffice`（LibreOffice）。
- 所有脚本使用 `Path(__file__).resolve().parent` 自定位，无硬编码路径；每步均在 `SKILL.md` §3 与 `references/fallbacks-and-exceptions.md` 提供无脚本兜底方案。

## 限制与已知边界
- 纯扫描版无文本层 PDF 需先 OCR；本包不内置 OCR。
- 极复杂的多层矢量混排页面若自动裁剪偏紧，需按 `paper_text.txt` 页码用 `pymupdf` 手动微调 `clip`。
- 不修改模板母版 XML 本身；模板原有占位符若本身越界会如实继承并记录在 `deck_check.json` 与 `uncertainty`。
