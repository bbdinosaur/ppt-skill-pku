---
name: "academic-slide-template-adaptor"
description: "Convert a research paper (PDF/materials) plus a PPTX/POTX template (and optional starting deck) into a template-faithful, editable academic presentation with per-slide evidence mapping, speaker notes, native tables/two-column/figure layouts, 2D geometry & overlap QA, and verified PDF/PNG export."
---

# Academic Slide Template Adaptor — 执行卡

**触发**：将论文/科研材料套入 PPTX/POTX 模板（或改造起始旧 PPT），产出可编辑演示、逐页证据映射与渲染件。
**铁律**：母版/占位符继承 > 一切修饰；事实数字逐字核对源文件带页码；Turn 3 即落盘完整 v1；绝不编造数据。
**脚本路径**：命令均用无状态展开 `${SKILL_DIR:-/skills/weekly-skill}/scripts/...`（直接复制运行；挂载于别处则先 `export SKILL_DIR=/path/to/skill`）。

## 0) 输入 / 输出契约
- **禁猜文件名**：Step 1 实时扫描 `/input/` 并读 `case.json`/`task.md`，按实文件定位输入（标签与实文件冲突时以实文件为准并记 `uncertainty`）。
- 中间文件放 `/app/work/`；**最终交付全落 `/app/output/`**：`deck.pptx`、`deck.pdf`、`previews/page-XX.png`（108 DPI）、`result.json`。
- `result.json`：**必须含 `case_id` 原值与 `output_contract` 全部字段名**，外加 §3 标准段（见 `references/result-schema.md`）。

## 1) 3–5 轮极速闭环（先成稿后精修，单次定向微调）

### Step 1（Turn 1）：一站式盘点与抽取简报（≤32 KB 防截断）
```bash
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/prepare_brief.py" --input /input --workdir /app/work
```
- 单命令扫描 `/input` 实文件，一次输出：`case_id`、`output_contract`、`task.md`、模板全部版式槽位/主题色与底色、论文全部图注与裁图路径、提取的数值表行、`key_stats`、优先保护实验/结果/局限段的章节摘录，以及起始旧 PPT 摘要与迁移草稿 `/app/work/plan_draft.json`。
- 完整产物同步落盘 `/app/work/template_inventory.json`、`/app/work/paper_structure.json`、`/app/work/paper_text.txt`、`/app/work/figs/`（需手动指定源文件时可单跑 `inspect_template.py` / `extract_paper.py` / `inspect_deck.py`）。

### Step 2（Turn 2）：规划并写入 `/app/work/plan.json`
- 完整字段见 `references/template-mapping.md`（若改造起始旧 PPT，直接基于 `/app/work/plan_draft.json` 增补模板/证据/备注保存为 `/app/work/plan.json`）。
- **顶层必填**：`template`（无模板填 `"none"`）、`output: "/app/output/deck.pptx"`、`audience`、`duration_min`、`language`；有固定页数要求时必填 `expected_slides`（锁页防自动拆页）；可选 `footer` / `slide_numbers: true`。
- **版式与内容规范**：组合使用 `title`、`section`、`content`（`**加粗导语**：解释`，主实验表优先原生 `table` 或 `chart`）、`two_col`（对称对比/双栏图/表）、`figure`（`auto|side_by_side|stacked`，单图+要点自动左文右图）。
- **逐页必备**：主张式结论标题、要点 ≤6 条、2–3 句可脱稿 `notes`、结构化 `evidence`（`{"locator": "§3.2 Table 1, p.5", "quote": "≤15词逐字原文"}`；数字/单位逐字核对 `brief.txt` 或 `/app/work/paper_text.txt`）。无固定页数时按 8–12 页组织：封面 → 动机 → 方法×2 → 实验/图/表×2–3 → 局限 → 结论。

### Step 3（Turn 3）：一键构建 + 静态 QA + 108 DPI 渲染 + 合成 `result.json`
```bash
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/build_deck.py" /app/work/plan.json --out /app/work/build_report.json && \
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/check_deck.py" /app/output/deck.pptx --json /app/work/deck_check.json && \
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/export_render.py" /app/output/deck.pptx --outdir /app/output/previews --pdf-out /app/output/deck.pdf --dpi 108 && \
python3 "${SKILL_DIR:-/skills/weekly-skill}/scripts/write_result.py" --case-json /input/case.json --plan /app/work/plan.json \
  --build /app/work/build_report.json --check /app/work/deck_check.json \
  --render /app/output/previews/render_log.json --deck /app/output/deck.pptx --out /app/output/result.json
```
- 自动继承模板真实版式与主题色、避让母版边缘徽标、自愈 0 高度占位符、清未使用空框、归一化 `.potx`；同步完成 2D 重叠/出界/溢出/拉伸/图号静态检查与 108 DPI PDF/PNG 导出及字体替代审计。
- **状态口径**：`write_result.py` 将硬性交付失败与启发式版式告警解耦——只要 `deck.pptx` 非空有效、`deck.pdf` 导出成功且页数匹配，顶层保持 `status="complete"`，同时在 `checks`/`limits`/`uncertainty` 如实记录启发式告警。

### Step 4（Turn 4–5）：视觉抽检 + 最多 1 次定向几何/契约微调
- 用图像读取工具抽检 3–4 张关键预览 `/app/output/previews/page-XX.png`（封面、最密表/图页、结论页），并核对 `deck_check.json` 告警与终端 `NULL_CONTRACT_FIELDS`。
- **仅当**存在真实视觉瑕疵（重叠/截断/对比度差/漏裁图）、关键几何问题，或 `NULL_CONTRACT_FIELDS` 中有可由源文件推导的字段时：一次性定向修补 `/app/work/plan.json`（追加 `change_log`）及可选 `/app/work/fills.json`，重跑 Step 3（若有 `fills.json`，给 `write_result.py` 加 `--fills /app/work/fills.json`）；无实损则直接结束，避免无效返工。

## 2) 关键异常与无脚本兜底（详参 `references/fallbacks-and-exceptions.md`）
- **无模板 / `.potx` / 起始 deck**：无模板设 `"template": "none"` 并记 `uncertainty`；`.potx` 自动归一化入 `normalizations`；起始 deck 以 `plan_draft.json` 保留旧表/图/原生图表并补齐论文证据。
- **缺图 / 缺占位符 / 导出失败**：未裁出图按 `/app/work/paper_text.txt` 页码用 `pymupdf` 200 DPI 手动 `clip`；缺 body 占位符自动退回手动框并记 `uncertainty`；`soffice` 失败自动重试 1 次，仍败则交付 pptx + `status: "partial"` + `limits`。
- **无脚本兜底**：依赖仅 `python-pptx` / `pymupdf` / `Pillow` / `soffice`；脚本不可用时按 `references/fallbacks-and-exceptions.md` 用内联 Python + `soffice`/`pdftoppm` 走完同契约闭环。

## 3) 预算纪律与 `result.json` 最小骨架（**示例 / Example**）
预算 64 请求 / 900s：目标 3 轮成稿、4–5 轮精修收官，每轮结束 `/app/output/result.json` 均为最新。
```json
{"case_id": "<from case.json>", "status": "complete",
 "deliverables": {"pptx": "/app/output/deck.pptx", "pdf": "/app/output/deck.pdf",
                  "previews_dir": "/app/output/previews", "previews": ["/app/output/previews/page-01.png"]},
 "deck_summary": {"slide_count": 8, "layouts_used": ["Title Slide", "Title and Content", "Two Content"],
                  "placeholder_chain": [{"slide": 2, "layout": "Title and Content", "title_ph": {"idx": 0, "type": "TITLE", "pt": 28}}]},
 "evidence_map": [{"slide": 3, "kind": "content", "claim": "...",
                   "sources": [{"locator": "§2.1 p.3", "kind": "paper-section", "quote": "exact <=15 words"}]}],
 "rendering": {"pages_match": true, "title_text_ok": true, "blank_pages": [], "pdf_fonts_used": ["..."], "font_warnings": []},
 "checks": {"geometry_critical": 0, "overlap_critical": 0, "text_overflow_critical": 0, "aspect_distortion": 0, "heuristic_alerts": []},
 "uncertainty": [], "change_log": ["v1 initial build", "v2 targeted figure polish"], "limits": []}
```
按需查阅：`references/template-mapping.md`（plan schema/表/双栏/起始 deck）、`references/layout-rules.md`（字号/重叠/间距）、`references/result-schema.md`（字段口径/自检）、`references/fallbacks-and-exceptions.md`（无脚本兜底/异常全表）。
