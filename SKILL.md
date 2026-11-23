---
name: "academic-slides-pku"
description: "Turn a fixed paper/research material into an editable, template-inheriting academic deck (native PPTX + PDF export + per-slide evidence map), or edit an existing deck in place with per-slide diffs. High-fidelity two-command pipeline with surgical local patching finishes in 4-8 steps well within budget."
---

# 学术 PPT 生成与编辑 — 高保真极速两步执行卡

**触发**：输入 `/input/case.json` + `task.md` + `material.json`/论文 PDF + `slide-workorder.json` +
真实 `target-template.pptx` / `starting-deck.pptx`，产出原生 PPTX、真 PDF 导出与页级证据映射。
**严控墙钟与步数（上限 900s / 64 步 → 目标 4–8 步 / < 350s）**：
- **严禁**单独 `ls /input`、通读 `scripts/*.py`、手写 `python-pptx` 补丁、用 `read` 读二进制 `.png`/`.pdf`/`.pptx`（视觉看 `export.visual_ok`），或因小改动整文件重写几百行 `slides_plan.json`。
- **绝不盲猜原文**：Step 1 内联 `≤28 KB` 高保真摘要；仅当 `read_more_hints` 非空时按页码精准补读 1–2 次核心实验/局限段。

## 1. 高保真极速闭环（核心 4–6 步交付）

### Step 1 一键清点 + 紧凑摘要 + 落盘保底 `result.json`（1 步）
```bash
RP=/skills/weekly-skill/scripts/run_pipeline.py
[ -f "$RP" ] || RP=$(find /skills /app /root /workspace -maxdepth 5 -name run_pipeline.py 2>/dev/null | head -n1)
python3 "$RP" prepare --case-dir /input --work /app/work --out /app/output
```
单命令完成解析、落盘保底 `partial` 版 `/app/output/result.json`，并在 stdout 输出 `≤28 KB` 紧凑简报：
`PREPARE_OK {...}`（硬约束/版式/真实图片路径+图注/`read_more_hints`）+ `TASK.MD`（≤7.5 KB）+ `MATERIAL DIGEST`（导航/首页摘要/全部图表原注/PDF 与 CSV 表/`material.json`/全篇关键句，≤13.5 KB）+ `START DECK DUMP`（edit 原稿逐页内容，≤6.5 KB）。
- **按需补读（0–2 步）**：`read_more_hints` 为空直接进 Step 2；非空时只按 `<!-- PDF page N -->` 或 `material.json` 补读关键方法/结果/局限段，**不整篇通读**。

### Step 2 一次性编写 `/app/work/slides_plan.json`（1 步）
直接套用下方示例（扩展字段按需看 `references/plan-schema.md`，通常无需读取）：

```json
{
  "deck_title": "Paper Short Title", "mode": "create", "rationale": "6-slide defense deck",
  "slides": [
    {"role": "title", "layout_hint": "title", "title": "Main Title", "subtitle": "Authors · Venue", "notes": "Opening hook and agenda."},
    {"role": "method", "layout_hint": "content", "title": "Three-Stage Pipeline Cuts Overhead",
     "steps": [{"title": "1. Encode", "desc": "Project inputs"}, {"title": "2. Gate", "desc": "Sparse routing"}],
     "image": {"path": "/app/work/figs/fig_p2_0.png", "ref": "Fig. 1", "caption": "System overview"},
     "notes": "Walk through pipeline stages and Fig. 1.", "sources": [{"ref": "Fig. 1", "locator": "p.2", "quote": "verbatim excerpt"}]},
    {"role": "result", "layout_hint": "content", "title": "Method Cuts Error by 58.4%",
     "highlights": [{"value": "58.4%", "label": "Error Drop"}, {"value": "142", "label": "Qubits"}],
     "bullets": [{"t": "Core Gain: lowers error by 58.4% vs baseline", "lvl": 0}],
     "table": {"ref": "Table 1", "caption": "Main comparison", "highlight_row": -1,
               "rows": [["Method","Error","Time"],["Base","4.82","18.5"],["Ours","1.31","7.8"]]},
     "takeaway": "Achieves 1.31 error at 7.8 us.", "notes": "Walk through Table 1 and transition.",
     "sources": [{"ref": "Table 1", "locator": "p.1", "quote": "verbatim excerpt"}]}
  ],
  "uncertainty": []
}
```
（也可用 `"chart": {"type":"bar|hbar|stacked_bar|line|pie","x_title":"...","y_title":"...","highlight_last":true,"categories":[...],"series":[{"name":"...","values":[...]}]}`、并排双图 `"images"`、双栏 `"two_col"` 替代 `table`。）
- **edit/adapt 模式**（原地修改，不重做无关页）：设 `"mode": "edit"`，用 `"edits"` 逐条映射任务指令：
  ```json
  {"mode": "edit", "rationale": "Apply requested edits in place", "edits": [
    {"op": "replace_text", "slide": 2, "old": "58.0%", "new": "58.4%", "sources": [{"ref": "Table 1", "locator": "p.1", "quote": "exact text"}]},
    {"op": "set_table", "slide": 3, "rows": [["Model","Acc"],["Ours","94.2%"]], "sources": [{"ref": "Table 2", "locator": "p.4", "quote": "exact text"}]},
    {"op": "replace_slide", "slide": 4, "slide_spec": {"role": "result", "title": "...", "bullets": [{"t": "..."}], "notes": "...", "sources": [...]}}
  ], "uncertainty": []}
  ```
  其余 op：`set_title`、`set_body`、`set_two_col`、`set_chart`、`set_image`、`set_notes`、`set_source`、`insert_at`、`append`、`delete`、`move`（`replace_text` 支持 `"slide": "all"` 且递归覆盖组合图形/表格/备注）。
- **学术高分准则**：严守 `slide_count`/`min_slides`/`max_slides`、`audience`、`language`、`duration` 与 `source_preservation`。节奏：封面 → 问题 → 方法（`steps`/`image`）→ 主结果（`highlights`+`table`/`chart`+`takeaway`）→ 消融 → 局限（`two_col`）→ 结论/Q&A；正文页优先填模板原生 `BODY` 槽位（首词自动加粗），单页选 1 种核心视觉。主张式标题（≤22 全角字，避 `→`/`·`）、≤6 条要点、自然演讲 `notes`（2–3 句）；非封面/目录/Q&A 页写 ≥1 条 `sources[].{ref,locator,quote}`（逐字复制）；**绝不编造或取整数字**，缺口入 `uncertainty`。

### Step 3 一键构建 + 内置自修 + 快速导出 + 终版 `result.json`（1 步）
```bash
RP=/skills/weekly-skill/scripts/run_pipeline.py
[ -f "$RP" ] || RP=$(find /skills /app /root /workspace -maxdepth 5 -name run_pipeline.py 2>/dev/null | head -n1)
python3 "$RP" deliver \
  --case-dir /input --work /app/work --out /app/output \
  --plan /app/work/slides_plan.json --dpi 100
```
单命令完成：`build_deck.py` → `check_deck.py`（13 项）→ 内置 ≤2 轮自动修补（紧缩/限高/字号钳制/引文对齐/缺源备注补齐，绝不改数字）→ `export_preview.py`（LibreOffice 隔离 `.lohome` **55s 快失败**导出真 PDF，超时自动降级 `reportlab` + 100 DPI PNG + Pillow `visual_ok`）→ `assemble_result.py --status auto`。

### Step 4 局部热修与严格停机（**最多重跑 1 次，严禁整文件重写**）
- **收到 `DELIVER_OK` 且 `missing_fields: []` 且 `export.visual_ok: true`**：**立即停止并向用户汇报交付摘要**，不再调用工具（绝不为写 `version_history.md` 或微调润色额外重跑导出）。
- **收到 `DELIVER_ISSUES`**：仅针对报错的 1–2 页，用局部精确 `edit` 修改 `/app/work/slides_plan.json` 对应 10 行字段，**或直接用 `--patch-slide` 局部热修重跑**（耗时 <15s，**严禁整文件重写几百行 JSON**）：
  ```bash
  python3 "$RP" deliver --case-dir /input --work /app/work --out /app/output \
    --plan /app/work/slides_plan.json --dpi 100 \
    --patch-slide '3:{"takeaway":"...","sources":[{"ref":"Table 1","locator":"p.4","quote":"exact text"}]}'
  ```
  **最多重跑 1 次 Step 3**（第 2 次运行残余纯启发式边缘提示自动诚实记入 `exceptions`，不再诱发第 3 轮循环），随后**立即停止交付**。

## 2. 无脚本兜底与交付红线
- **无脚本降级**：若 `run_pipeline.py` 不可用，先写 `/app/work/slides_plan.json` 与诚实 `partial` 版 `/app/output/result.json`，再用内联 `python-pptx` 打开真实模板逐页 `add_slide`（不设 `font.name`、等比插图、写 `notes_slide`）→ `soffice --headless --convert-to pdf`（或 `reportlab`）+ `pdftoppm -png -r 100`；工具全缺失则输出完整页级证据的 `status=partial` `result.json` 并写明 `exceptions`。
- **交付红线**：不改母版/主题；不设 `font.name`；图片等比；LibreOffice 缓存放 `/app/work/.lohome` 不污染 `/app/output`；页面实质数字必在材料/任务书中可指认，绝不虚报 `status=ok`。
