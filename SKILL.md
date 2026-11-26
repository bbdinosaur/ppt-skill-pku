---
name: "academic-slides-pku"
description: "Turn a fixed paper/research material into an editable, template-inheriting academic deck (native PPTX + PDF export + per-slide evidence map), or edit an existing deck in place with per-slide diffs. High-fidelity two-command pipeline with surgical local patching finishes in 4-8 steps well within budget."
---

# 学术 PPT 生成与编辑 — 高保真极速两步执行卡

**触发**：输入 `/input/case.json`、`task.md`、`material.json`/PDF、`slide-workorder.json` 与 `target-template.pptx`/`starting-deck.pptx`，产出原生 PPTX、真 PDF 与页级证据映射。
**严控预算（上限 900s / 64 步 → 目标 4–8 步 / <350s）**：
- **严禁**单独 `ls /input`、读 `scripts/*.py`、手写 `python-pptx`、用 `read` 读 `.png`/`.pdf`/`.pptx`（看 `export.visual_ok`），或整文件重写 `slides_plan.json`。
- **绝不盲猜原文**：Step 1 内联 `≤28 KB` 简报；仅 `read_more_hints` 非空时按页码补读 1–2 次核心段。

## 1. 高保真极速闭环（核心 4–6 步）

### Step 1 一键清点 + 紧凑简报 + 保底 `result.json`（1 步）
```bash
RP=/skills/weekly-skill/scripts/run_pipeline.py
[ -f "$RP" ] || RP=$(find /skills /app /root /workspace -maxdepth 5 -name run_pipeline.py 2>/dev/null | head -n1)
python3 "$RP" prepare --case-dir /input --work /app/work --out /app/output
```
单命令解析全量输入、落盘保底 `partial` 版 `result.json`，并输出 `≤28 KB` 简报：`PREPARE_OK`（约束/`notes_pacing_target`/`visual_policy`/版式/图注/`read_more_hints`）+ `TASK.MD`（≤7.5 KB）+ `MATERIAL DIGEST`（导航/摘要/图注/PDF 与 CSV 表/`material.json`/关键句，≤13.5 KB）+ `START DECK DUMP`（≤6.5 KB）。
- **按需补读（0–2 步）**：`read_more_hints` 为空直进 Step 2；非空按 `<!-- PDF page N -->` 补读关键段。

### Step 2 一次性编写 `/app/work/slides_plan.json`（1 步）
套用下方示例（含方法图/KPI+表/原生数据绑定图表；扩展字段见 `references/plan-schema.md`，通常无需读取）：

```json
{
  "deck_title": "Paper Short Title", "mode": "create", "rationale": "8-slide academic talk",
  "slides": [
    {"role": "title", "layout_hint": "title", "title": "Main Title", "subtitle": "Authors · Venue",
     "notes": "Introduce the core problem, highlight our 58.4% error drop, and preview the roadmap."},
    {"role": "method", "layout_hint": "content", "title": "Three-Stage Pipeline Cuts Overhead",
     "steps": [{"title": "1. Encode", "desc": "Project inputs"}, {"title": "2. Gate", "desc": "Sparse routing"}],
     "image": {"path": "/app/work/figs/fig_p2_0.png", "ref": "Fig. 1", "caption": "System overview"},
     "notes": "Walk through Fig. 1 stages and how sparse gating lowers routing cost before main results.",
     "sources": [{"ref": "Fig. 1", "locator": "p.2", "quote": "verbatim excerpt"}]},
    {"role": "result", "layout_hint": "content", "title": "Method Cuts Error by 58.4%",
     "highlights": [{"value": "58.4%", "label": "Error Drop"}, {"value": "142", "label": "Qubits"}],
     "bullets": [{"t": "Gain: cuts error 58.4% (1.31 vs 4.82)", "lvl": 0},
                 {"t": "Speed: latency drops 18.5 us to 7.8 us", "lvl": 0}],
     "table": {"ref": "Table 1", "caption": "Main comparison", "highlight_row": -1,
               "rows": [["Method","Error","Time"],["Base","4.82","18.5"],["Ours","1.31","7.8"]]},
     "takeaway": "Achieves 1.31 error at 7.8 us.",
     "notes": "Table 1 shows error drops 4.82 to 1.31 across 142 qubits at 7.8 us. Next: scaling.",
     "sources": [{"ref": "Table 1", "locator": "p.1", "quote": "verbatim excerpt"}]},
    {"role": "analysis", "layout_hint": "content", "title": "Runtime Drops 57.8% Across Scales",
     "bullets": [{"t": "Scaling: linear speedup across regimes", "lvl": 0},
                 {"t": "Overhead: stays under 7.8 us at 142 qubits", "lvl": 0}],
     "chart": {"type": "bar", "title": "Latency (us)", "y_title": "us", "highlight_last": true,
               "categories": ["Base","Ours"], "series": [{"name": "Time", "values": [18.5, 7.8]}],
               "ref": "Table 1", "caption": "Data-bound latency comparison"},
     "notes": "This native chart compares 18.5 us baseline vs 7.8 us ours, confirming 57.8% speedup.",
     "sources": [{"ref": "Table 1", "locator": "p.1", "quote": "verbatim excerpt"}]}
  ],
  "uncertainty": []
}
```
（支持 `hbar`/`stacked_bar`/`line`/`pie`、双图 `"images"`、双栏 `"two_col":{"left_title":"...","left_bullets":[...],"right_title":"...","right_bullets":[...]}`。）
- **edit/adapt 模式**（原地改，不重做无关页）：设 `"mode":"edit"`，用 `"edits"` 映射指令：
  ```json
  {"mode": "edit", "rationale": "Apply requested edits in place", "edits": [
    {"op": "replace_text", "slide": 2, "old": "58.0%", "new": "58.4%", "sources": [{"ref": "Table 1", "locator": "p.1", "quote": "exact text"}]},
    {"op": "set_table", "slide": 3, "rows": [["Model","Acc"],["Ours","94.2%"]], "sources": [{"ref": "Table 2", "locator": "p.4", "quote": "exact text"}]},
    {"op": "replace_slide", "slide": 4, "slide_spec": {"role": "result", "title": "...", "bullets": [{"t": "..."}], "notes": "...", "sources": [...]}}
  ], "uncertainty": []}
  ```
  其余 op：`set_title`/`set_body`/`set_two_col`/`set_chart`/`set_image`/`set_notes`/`set_source`/`insert_at`/`append`/`delete`/`move`（`replace_text` 支持 `"slide":"all"` 并覆盖组合图形/表/图轴/备注）。
- **学术高分准则**：严守 `slide_count`/`min_slides`/`max_slides`、`audience`、`language`、`duration` 与 `source_preservation`。节奏：封面→问题→方法（`steps`/`image`）→主结果（`highlights`+`table`+`takeaway`）→**含数表时至少 1 页原生 `chart` 与表/图交替**→局限（`two_col`）→结论/Q&A。主张式标题（≤22 全角字，避 `→`/`·`）、每页 2–5 要点、**贴合 `duration` 的 2–4 句演讲 `notes`（动机→数值→衔接）**；非封面/目录/Q&A 页写 ≥1 条 `sources[].{ref,locator,quote}`（逐字复制）；**绝不编造或取整数字**，缺图/位图/缺口如实入 `uncertainty`。

### Step 3 一键构建 + 内置自修 + 导出 + 终版 `result.json`（1 步）
```bash
python3 "${RP:-/skills/weekly-skill/scripts/run_pipeline.py}" deliver \
  --case-dir /input --work /app/work --out /app/output --plan /app/work/slides_plan.json --dpi 100
```
自动完成：`build_deck` → `check_deck`（13–14 项含洁净/零命中）→ ≤2 轮自动修补（紧缩/限高/引文对齐/补源与备注，绝不改数字）→ `export_preview`（隔离 `.lohome` **110s** 导出真 PDF，超时降级 `reportlab` + 100 DPI PNG + Pillow `visual_ok`）→ `assemble_result --status auto`。

### Step 4 局部热修与停机（**最多重跑 1 次，严禁整文件重写**）
- **`DELIVER_OK` 且 `missing_fields:[]` 且 `export.visual_ok:true`**：**立即停止并汇报摘要**。
- **`DELIVER_ISSUES`**：仅修报错 1–2 页，用 `--patch-slide` 热修重跑（<15s）：
  ```bash
  python3 "${RP:-/skills/weekly-skill/scripts/run_pipeline.py}" deliver --case-dir /input --work /app/work \
    --out /app/output --plan /app/work/slides_plan.json --dpi 100 \
    --patch-slide '3:{"takeaway":"...","sources":[{"ref":"Table 1","locator":"p.4","quote":"exact text"}]}'
  ```
  **最多重跑 1 次**（第 2 次残余启发式边缘诚实记入 `exceptions`），随后**立即停止**。

## 2. 无脚本兜底与交付红线
- **降级**：无脚本时先写 `slides_plan.json` 与 `partial` `result.json`，用内联 `python-pptx` 继承模板 `add_slide`（不设 `font.name`、等比插图、写备注）→ `soffice`/`reportlab` + `pdftoppm -r 100`；工具全缺失则输出含页级证据的 `status=partial` `result.json` 与 `exceptions`。
- **红线**：不改母版/主题、不设 `font.name`、图片等比、LibreOffice 缓存放 `/app/work/.lohome`、数字必可回查，绝不虚报 `ok`。
