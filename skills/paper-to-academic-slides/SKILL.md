---
name: "paper-to-academic-slides"
description: "Convert a research paper PDF or study material into a template-compliant, editable academic slide deck with per-slide evidence mapping back to paper figures/tables/sections, distinct speaker notes, PDF/PNG render verification, and a structured result.json. Use whenever the task asks to turn a paper/report into academic PPT/slides/presentation."
---

# Paper → Academic Slides（极速执行卡）

将论文或科研材料转为**主线成立、页页可溯、模板合规、渲染可查**的可编辑学术 PPT。一页一主张并挂靠论文图/表/章节，严禁段落搬运或编造数据。

## 运行时契约（Harbor）
- 脚本路径：`${SKILL_DIR:-/skills/weekly-skill}/scripts/`（容器默认 `/skills/weekly-skill`）。
- 输入：`/input/case.json`、`/input/task.md` 及论文/材料（可含母版模板、起始 deck PPTX 或补充图/表数据）。
- 工作目录 `/app/work`；最终产物与 `result.json` 输出到 `/app/output`（对齐 `output_contract`）。
- 预算：**64 请求 / 900s / 16384 token**。**4–6 步极速闭环**：`prep` 立刻落盘保底 `result.json` 并生成一站式 `reading_pack.md`，**单次读取即可写大纲并一键构建校验**。
- 诚实原则：实际文件与原论文为唯一事实依据（若标签矛盾以原文为准）；缺料、OCR、字体替换或失败如实记入 `result.json`，严禁编造数据或状态。

## 4 步极速闭环（正常 4–6 次调用）

### Step 1：一键解析、抽取、生成一站式简报与保底落盘（1 步）
```bash
python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/run_pipeline.py prep --work /app/work --out-dir /app/output
```
执行 `resolve_case.py` + `extract_paper.py`（低文本扫描页自动 Tesseract OCR），**立刻落盘保底 `/app/output/result.json`**（`status="partial"`），并生成 **`/app/work/paper/reading_pack.md`**（≤45 KB，完整打包 `task.md`、起始 deck 逐页检查、`starting_deck_figs/`、补充材料、**图/表置顶的 `paper_index.md`**、核心页**不截断逐字原文**及剩余页清单）。
- 无脚本兜底：直接读 `/input/case.json`、`/input/task.md` 并 `ls -R /input`，用 `pdftotext` 或直读文本。

### Step 2：一次读 `reading_pack.md` → 直接写 `/app/work/slides.json`（2 步）
1. **先读 `/app/work/paper/reading_pack.md`**；若文末 `Remaining Pages` 中有目标实验/消融/局限页，一次性补读对应 `pages/pNN.txt`，确保零遗漏。
2. 按学术主线一次写完 `/app/work/slides.json`：
`cover → 问题/动机 → 方法(1–2页) → 实验设置 → 主结果(1–3页,每页挂1图/表) → 消融/对比 → 局限 → 结论`。

**强制写作规则**（契约见 `references/slides-schema.md`，细则见 `references/quality-and-evidence.md`）：
1. **主张句标题**：`title` 必须为主张句（禁"方法""实验结果""Results"等栏目词）；主结果标题含原数字+单位；页数 ≤ `max_slides`（默认 ≤12），匹配受众/时长。
2. **精炼保真**：每页 3–6 条 bullets，每条 ≤22 词/30 字，支持 `**关键词/数值**` 原生加粗（正文与表格均支持）；标题/正文/表中数字、单位、条件**逐字保留，严禁四舍五入、改单位或无中生有**。
3. **归属与图类型**：对比/基线数据须显式标明基线名或作者；框架/流程/示意图须在 `image_caption` 或 `notes` 标明“示意图，非实验数据”。
4. **必保局限页**：`limitations_pages` 非空且 `max_slides ≥ 7` 时**必留局限页**；超限时先裁背景与次要消融，必保封面/方法/主结果/结论。
5. **页页可溯**：非封面页 `evidence.paper_refs` 必填 ≥1 条 `{"type":"figure|table|section|page|equation","id":"Fig. 3","page":5,"quote":"≤25词原文"}`（**必含原文 `quote`**）；`evidence.slide_line` 填页脚溯源（漏填自动合成）。
6. **真图真表与视觉层次**：优先 `image`/`images` 填 `figs/*.png`、补充图或 `starting_deck_figs/*.png`（多面板选 `composite_panels`，图号一致，**附居中 `image_caption`**）；主表用 `table` 三线表，无图可抠的数值对比可用原生 `chart`；结果/结论页可加 `kpis` 或底部 `takeaway`；纯文对比用 `right_bullets` 双栏；`--reuse` 支持 `"source_slide": N` 重排与 `"keep_existing_media": true`。
7. **讲稿分离与控时**：`notes` 写 `[~X.X min]` 控时 + 过渡 + 强调点及成立前提 + 追问补充证据，**严禁复述页面文字**。

### Step 3：一键构建 + 版面检 + S4.5 强校验 + 渲染导出（1 步）
```bash
python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/run_pipeline.py build /app/work/slides.json --work /app/work --out-dir /app/output
```
自动按 `case_resolved.json` 适配模板（纯模板清样例页重建 / 起始 deck `--reuse` 原地适配并保护角标 Logo、页眉页脚与记录逐页 diff / 无模板生成带页码的规范 16:9），依次执行：
1. `build_deck.py` → `/app/work/deck.pptx`（只读 XML 保母版背景、零投影扁平视觉、三线学术表、原生图表、KPI 指标卡、自动导语加粗、双栏对比、数字单位防断行、单行标题、深色反色、智能避让侧边栏/母版页脚与等比防裁切）
2. `check_layout.py` → `/app/work/layout_report.json`（越界、字号、严重溢出/重叠与母版页脚碰撞升 `fail`、占位残留词）
3. **`verify_evidence.py`（S4.5，必跑）** → `/app/work/evidence_report.json`（页数上限、必有证据与原文 `quote`、支持补充材料溯源、图存在/图号匹配/防重图/图注、全量数字与**单位篡改检测**、基线归属、讲稿查重与控时、栏目词、局限页）
4. `finalize_output.py` → `/app/output/{slides.pptx, slides.pdf, previews/slide-NN.png (100 DPI), evidence_map.json, evidence_map.md, speaker_notes.md, slides_outline.json, result.json}`（含六要素审计、PDF 字体审计、空白/豆腐块质检与 `output_contract` 镜像）

**强制门控（严禁跳过）**：若 `action_required: true`，按返回的 **`polish_checklist`** 一次性改完 `slides.json` 重跑 `build` 直至清零（≤2 轮）；剩余低优 `warn` 诚实记入 `result.json`。
- 分步/兜底：6 个脚本可独立调用；无 `soffice` 标 `pdf_exported:false`；无 `python-pptx` 交 `slides.json`+`evidence_map.json` 并标 `status=partial`。

### Step 4：渲染核验与交付确认（0 额外步，兼容纯文本模型）
- **自动化门控**：核对 `layout_fail == 0`、`evidence_errors == 0`、`pdf_pages_match == true`、`render_text_ok == true`、`visual_pixel_ok == true`。
- **视觉抽检（仅多模态环境可选）**：支持读图时可选 `read` 抽检 1–2 张 `inspect_previews`；**纯文本模型严禁读 PNG**，自动指标全绿即可交付。

交付清单：
- [ ] `layout_fail == 0`、`evidence_errors == 0`、`visual_pixel_ok == true`、`status == "success"`（异常时诚实 `partial`）
- [ ] 主张句标题、原数字+单位、图号与图一致且有图注、示意图/基线区分、局限页、非复述 `notes`
- [ ] `/app/output/result.json` 字段齐备（含 `six_element_audit`）并对齐 `output_contract`

## 异常、示例与依赖
- 加密/扫描 PDF、占位符缺失、页数超限、字体缺失处置见 `references/failure-playbook.md`。
- **最小调用示例（示例 / Example，非真实案例）**：输入 `{"case_id":"demo-1","paper":"paper.pdf","template":"tpl.pptx","max_slides":10}` → `prep` → 读 `reading_pack.md` → 写 `slides.json` → `build`。示意输出：`{"case_id":"demo-1","status":"success","deck":{"n_slides":10,"notes_complete":true},"checks":{"pdf_exported":true,"geometry_fail":0,"evidence_errors":0,"visual_pixel_ok":true}}`。
- 依赖：Python ≥ 3.8、`python-pptx`、`pymupdf`、`Pillow`；可选 `libreoffice`、`pdftotext`（见 `README.md`）。
