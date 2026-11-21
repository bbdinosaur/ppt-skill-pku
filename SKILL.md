---
name: "academic-slides-pku"
description: "Turn a fixed paper/research material into an editable, template-inheriting academic deck (native PPTX + PDF export + per-slide evidence map), or edit an existing deck in place. Verifies geometry, fidelity and quote authenticity with scripts before delivery."
---

# 学术 PPT 生成与编辑 — 执行卡

**触发**：运行时给 `/input/case.json` + `material.json` + `artifacts/slide-workorder.json` +
真实 `target-template.pptx` / `starting-deck.pptx`，要求产出可编辑学术演示稿、PDF 导出与
页级证据映射。**预算 64 次请求 / 900s / 16384 tokens** — 一切以"脚本干重活、模型做决策"为原则，
全程约 8–12 次请求、<3 分钟。先产出合法首版，再精修。

## 0. 路径与速览

- 输入 `/input`，工作台 `WORK=/app/work`，产物 `OUT=/app/output`。
- `S = <本 Skill 目录>`；脚本用 `python3 $S/scripts/xxx.py ...`（自带路径解析，任意 cwd 可用）。
- 首步先看 `WORK/inventory.json` 而不是猜文件内容；**材料原文是唯一事实依据**，领域标签不可信。
- 详细规范按需读：`references/plan-schema.md`（计划 schema+证据规则）、
  `references/template-rules.md`（模板/版式铁律）、`references/failure-playbook.md`（异常处置）。

## 1. 流程（每步一条命令，失败即看 stdout JSON）

**Step 1 清点输入（1 次请求）**
```bash
python3 $S/scripts/inspect_case.py --case-dir /input --work $WORK
```
输出 `INSPECT_OK {...}`：workorder 约束（mode/slide_count/audience/language）、模板版式清单、
起始 deck 页数、材料 keys、PDF 页数与抽取文本位置（`$WORK/material_text.md`）、异常清单。
**同时**：把已有的输入异常先记下来（后面进 exceptions）。

**Step 2 立即落一版诚实的 partial result（防中断，0 决策）**
```bash
python3 $S/scripts/assemble_result.py --case-dir /input --work $WORK \
  --deck $OUT/presentation.pptx --status partial
```

**Step 3 读材料 → 写计划（1–2 次请求，本 Skill 唯一重决策点）**
读 `$WORK/material_text.md`（或 material.json），提取：研究问题/方法/主结果数字/局限 +
每条结论的原文定位（章节号、页码、图号）。**先扫 material.json 的 key_numbers/tables/figures
锁定权威数字集，再读全文补上下文；页面数字直接从这些字段复制粘贴，不要凭记忆打字。**
**超长材料分段读**：先读摘要/实验/结论段与图表清单页标记，再按页码标记按需补读，
不要一次性把几万字全部读入。**扫描版 PDF（material_text.md 为空）**：以 material.json
为准，并在 uncertainty 记录 "PDF 无文本层，数字取自结构化材料"。
然后**只写一个文件** `$WORK/slides_plan.json`
（schema 见 references/plan-schema.md）：
- 每页：主张式标题、≤6 条要点（每条 ≤2 行）、`notes` 讲稿 2–4 句、
  `sources[].{ref,locator,quote}`（quote **逐字**取自材料，脚本会核验）；
- 图：优先数据绑定原生图表（材料含某图完整数值时用 `chart`，数值逐项取材料，生成可在 PowerPoint
  中直接查验/编辑的原生图表）；
  其余图只引用真实存在的图（PDF 抽取图在 `$WORK/figs/`，材料图用 inventory 给的绝对路径），
  位图无数据则只引用原图不重绘；`caption` 逐字用论文图注；
- 页数 = workorder `slide_count`（无则 10–14，不得超过 `max_slides`）；edit 模式只写 `edits`，
  不碰其他页；workorder 若给 edit_instructions 列表，逐条映射为 edits 操作；
- 材料缺口写进 `uncertainty`，**没有的数字绝不编**。

**Step 4 构建（原生可编辑、模板继承）**
```bash
python3 $S/scripts/build_deck.py --plan $WORK/slides_plan.json --work $WORK \
  --out $OUT/presentation.pptx
```
输出 `BUILD_OK slides=N`。edit 模式自动用起始 deck，create 模式用 target-template 的真实版式。

**Step 5 检查 + 导出**
```bash
python3 $S/scripts/check_deck.py --deck $OUT/presentation.pptx --work $WORK
python3 $S/scripts/export_preview.py --deck $OUT/presentation.pptx --work $WORK
```
检查 13 项：页数/页数上限/越界/文字溢出/重叠/表格溢出/图片等比/标题+备注/字号下限/
字体继承/证据来源/引文真实性/数字保真（实质性数字必须能在材料原文回查，防抄错取整）。
导出产出真实 PDF + `previews/slide-XX.png`。

**Step 6 修复循环（≤2 轮，超出即如实报告）**
看 `CHECK_ISSUES` 的 failures 列表：改 `slides_plan.json` 对应页（精简文字/拆页/换图），
重跑 Step 4–5。几何硬错误（越界/图片拉伸/重叠）必须清零；无法修复的按
references/failure-playbook.md 写入 exceptions。
若运行时支持读图：用 read 看 `previews/*.png` 逐页目检（留白、层级、图注对齐、来源行）；
不支持读图时以程序化检查为准。

**Step 7 终版 result.json（含全部契约字段）**
```bash
python3 $S/scripts/assemble_result.py --case-dir /input --work $WORK \
  --deck $OUT/presentation.pptx --status auto
```
产出：`result.json`（case_id/status/editable_source/export_artifact/content_map/
compliance_checks/exceptions + uncertainty/change_log）、`plan.json`、`change_log.md`、
`qa_report.md`（一页式审计：检查表+异常+页级证据，供复核）、`outline.md`（讲稿大纲）、
`previews/contact-sheet.png`（全 deck 缩略拼图，一眼目检）、`presentation.pptx`、
`presentation.pdf`、`previews/*.png`。核对 stdout `missing_fields:[]`。

## 2. 无脚本兜底（脚本不可用时不断链）

1. 手写 `$WORK/slides_plan.json`（schema 不变）→ 若 python-pptx 可用，用一段内联 Python：
   打开模板 → `prs.slide_layouts` 逐页 add_slide → 占位符 `text_frame` 填字（不设 font.name）→
   `add_picture(width=…)` 等比插图 → `notes_slide` 写备注 → 保存。
2. PDF：`soffice --headless --convert-to pdf`；预览：`pdftoppm -png -r 110`。
3. 全部工具缺失：把计划与证据映射完整写进 result.json（status=partial, editable_source=null），
   exceptions 注明环境缺失 —— 诚实交付优于伪造。

## 3. 环境自检（可选）

- 一次请求：：`python3 $S/scripts/selftest.py` — 在临时目录生成 mini 算例
  并全流程跑通，打印 `SELFTEST_OK`；用于快速验证脚本链可用。

## 4. 红线（交付底线）

- 不改母版/主题；不给文字设字体名；图片等比插入。
- 页面数字必须能在材料中指认；quote 必须逐字存在；编造=0 分。
- workorder 的 mode/页数/受众/语言是硬约束；`source preservation`（术语符号一致）是硬约束。
- 所有异常写 exceptions，status 用 auto 判定，不虚报 ok。
