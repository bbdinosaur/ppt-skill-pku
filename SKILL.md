---
name: "academic-slides-pku"
description: "Turn a fixed paper/research material into an editable, template-inheriting academic deck (native PPTX + PDF export + per-slide evidence map), or edit an existing deck in place with per-slide diffs. High-fidelity two-command pipeline with built-in audit and polish finishes in 5-10 steps."
---

# 学术 PPT 生成与编辑 — 高保真两步执行卡

**触发**：输入 `/input/case.json` + `task.md` + `material.json`/论文 PDF + `slide-workorder.json` +
真实 `target-template.pptx` / `starting-deck.pptx`，产出原生 PPTX、真 PDF 导出与页级证据映射。
**效率与质量平衡（预算 64 步 / 900s → 目标 5–10 步 / < 300s）**：
- **严禁**单独 `ls /input`、通读 `scripts/*.py`、手写 `python-pptx` 补丁，或用 `read` 读二进制 `.png`/`.pdf`/`.pptx`（视觉看 `export.visual_ok`）。
- **绝不盲猜原文**：Step 1 打印高保真摘要；若 `read_more_hints` 非空（长论文或大 `material.json`），务必按页码标记精准补读方法/实验/局限原文后再写计划。

## 1. 高保真闭环流程（5–10 步交付）

### Step 1 一键清点 + 高保真材料直出 + 落盘保底 `result.json`
```bash
RP=/skills/weekly-skill/scripts/run_pipeline.py
[ -f "$RP" ] || RP=$(find /skills /app /root /workspace -maxdepth 5 -name run_pipeline.py 2>/dev/null | head -n1)
python3 "$RP" prepare --case-dir /input --work /app/work --out /app/output
```
单命令完成解析、落盘保底 `partial` 版 `/app/output/result.json`，并向 stdout 打印：
1. `PREPARE_OK {...}`（`workorder` 硬约束、模板版式、真实图片路径+图注、异常与 `read_more_hints`）；
2. `=== TASK.MD ===`（≤10 KB）；3. `=== MATERIAL DIGEST ===`（论文导航、首页摘要原文、全部图表原注、PDF/CSV 表、完整 `material.json`、全篇按句切分关键句，≤18 KB）；4. `=== START DECK DUMP ===`（edit 原稿逐页内容/表格/图表/备注，≤10 KB）。
- **按需深读（0–2 步）**：`read_more_hints` 为空直接进 Step 2；非空时先扫 `material.json` 权威数字/图表，再用 `read(offset, limit)` 按 `<!-- PDF page N -->` 补读方法、主结果、消融与局限段。

### Step 2 编写学术计划 `/app/work/slides_plan.json`
直接按下方示例写入（扩展字段按需参考 `references/plan-schema.md`）：

```json
{
  "deck_title": "Paper Short Title", "mode": "create", "rationale": "6-slide defense deck",
  "slides": [
    {"role": "title", "layout_hint": "title", "title": "Main Title", "subtitle": "Authors · Venue", "notes": "Opening hook and agenda preview."},
    {"role": "result", "layout_hint": "content", "title": "Method Cuts Error by 58.4%",
     "highlights": [{"value": "58.4%", "label": "Error Drop"}, {"value": "142", "label": "Qubits"}],
     "bullets": [{"t": "Core Gain: lowers error by 58.4% vs baseline", "lvl": 0}],
     "table": {"ref": "Table 1", "caption": "Main comparison", "highlight_row": -1,
               "rows": [["Method","Error","Time"],["Base","4.82","18.5"],["Ours","1.31","7.8"]]},
     "takeaway": "Achieves 1.31 error at 7.8 us.", "notes": "Walk through Table 1, highlight the 58.4% reduction, and transition to ablation.",
     "sources": [{"ref": "Table 1", "locator": "p.1", "quote": "verbatim excerpt from material"}]}
  ],
  "uncertainty": []
}
```
- **edit/adapt 模式**（在原稿上原地修改，绝不重做无关页）：设 `"mode": "edit"`，用 `"edits"` 逐条映射任务修改指令：
  ```json
  {"mode": "edit", "rationale": "Apply requested edits in place", "edits": [
    {"op": "replace_text", "slide": 2, "old": "58.0%", "new": "58.4%", "sources": [{"ref": "Table 1", "locator": "p.1", "quote": "exact text"}]},
    {"op": "set_table", "slide": 3, "rows": [["Model","Acc"],["Ours","94.2%"]], "sources": [{"ref": "Table 2", "locator": "p.4", "quote": "exact text"}]},
    {"op": "replace_slide", "slide": 4, "slide_spec": {"role": "result", "title": "...", "bullets": [{"t": "..."}], "notes": "...", "sources": [...]}}
  ], "uncertainty": []}
  ```
  其余支持 op：`set_title`、`set_body`、`set_two_col`、`set_chart`、`set_image`（原地换图）、`set_notes`、`set_source`、`insert_at`、`append`、`delete`、`move`（`replace_text` 支持 `"slide": "all"` 且递归覆盖组合图形/表格/备注）。
- **高分学术编排准则**：
  - **硬约束**：严守 `slide_count`/`min_slides`/`max_slides`、`audience`、`language`、`duration` 与 `source_preservation`（术语符号一致）。
  - **叙事与视觉节奏**：封面 → 问题/动机 → 方法（`steps` 或原图 `image`）→ 主结果（`highlights` + `table`/`chart` + `takeaway`）→ 消融/分析 → 局限与未来（`two_col` 或模板正文槽位）→ 结论/Q&A。正文页优先填模板原生 `BODY` 槽位（首词自动加粗），单页选 1 种核心视觉，目录页自动渲染编号条。
  - **零错数与真实引文**：主张式标题（≤22 全角字，避 `→`/`·`）、≤6 条要点、自然演讲 `notes`（2–4 句，讲动机/结论/过渡）；非封面/目录/Q&A 页写 ≥1 条 `sources[].{ref,locator,quote}`（`quote` 逐字复制自材料）。位图或无完整数表只用真实图，**绝不编造或取整数字**，缺口写入 `uncertainty`。

### Step 3 一键构建 + 自动修补 + 导出 + 输出逐页审计摘要
```bash
RP=/skills/weekly-skill/scripts/run_pipeline.py
[ -f "$RP" ] || RP=$(find /skills /app /root /workspace -maxdepth 5 -name run_pipeline.py 2>/dev/null | head -n1)
python3 "$RP" deliver \
  --case-dir /input --work /app/work --out /app/output \
  --plan /app/work/slides_plan.json --dpi 100
```
单命令完成：`build_deck.py` → `check_deck.py`（13 项检查）→ 内置 ≤2 轮自动修补（排版紧缩、表格限高、字号钳制、引文对齐、缺源/缺备注补齐，绝不改数字）→ `export_preview.py`（LibreOffice 隔离 `.lohome` 90s 快失败导出真 PDF，失败降级 `reportlab` + 100 DPI PNG + Pillow `visual_ok`）→ `assemble_result.py --status auto`，并在 stdout 附带 `=== DECK AUDIT SUMMARY ===`。

### Step 4 定向精修与干净停机（最多重跑 1–2 次）
- **`DELIVER_ISSUES`**：按 `failures` 与 `closest` 提示一次性修正 `/app/work/slides_plan.json` 重跑 Step 3（≤2 轮；未消除项记入 `exceptions`，见 `references/failure-playbook.md`）。
- **`DELIVER_OK` 且 `missing_fields: []`**：直接核对 `=== DECK AUDIT SUMMARY ===` 与 `auto_repaired`（0 次额外读文件）：若无回退项且叙事/视觉/讲稿/任务要求齐备 → **立即停止并回复交付摘要**；若存在自动回退备注/引文或漏项 → 对 `slides_plan.json` 做 **1 次精准精修**并重跑 Step 3（同步沉淀 `v1 → v2` 至 `version_history.md`），随后立即停止。

## 2. 无脚本兜底与交付红线

- **无脚本降级**：若 `run_pipeline.py` 不可用，先写 `/app/work/slides_plan.json` 与诚实 `partial` 版 `/app/output/result.json`，再用内联 `python-pptx` 打开真实模板逐页 `add_slide`（不设 `font.name`、`add_picture(width=...)` 等比插图、写 `notes_slide`）→ `soffice --headless --convert-to pdf`（或 `reportlab`）+ `pdftoppm -png -r 100`；工具全缺失则输出完整页级证据的 `status=partial` `result.json` 并写明 `exceptions`。
- **交付红线**：不改母版/主题；不给 run 设 `font.name`；图片严格等比；LibreOffice 缓存放 `/app/work/.lohome` 不污染 `/app/output`；页面实质数字必在材料/任务书中可指认，绝不虚报 `status=ok`。
