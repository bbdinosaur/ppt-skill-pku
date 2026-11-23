---
name: "academic-talk-structurer"
description: "Organize a paper, multi-file material, or draft deck into an audience- and time-aware academic talk: one source-bound claim per slide, editable PPTX on template/starting deck (booktabs tables, charts, KPIs, step cards, two-column comparisons, deck diff & figure reuse), oral notes & audit matrix, 4-dim Q&A, visual check, and contract result.json."
---

# Academic Talk Structurer（学术汇报内容组织）

## 触发与边界
输入：`case.json` + `task.md` + 源材料（PDF/PPTX/DOCX/MD/CSV/JSON/XLSX）。目标：把已有材料组织成可上台的汇报、讲稿、证据映射与问答。**不做**论文精读批注、写作润色、科研绘图生产。

## 预算与铁律（64 req / 900s / 16k tok）
- **文件唯一事实源**：数字、图表与结论严格出自实际文件；**若 `case.json` 标签与原文冲突，以实际文件为准、禁按错标签补数**，记入 `uncertainty`。
- **极速闭环（4–6 req）**：`ingest.py` → 读 `reading_map.md` 与 `paper.txt` → 写 `deck_spec.json` → `deliver.py` → 接触表目检（≤2 轮）；`references/` 按需读。
- **路径与兜底**：中间放 `/app/work`，交付放 `/app/output`；经 `${SKILL_DIR:-/skills/weekly-skill}/scripts/` 调用（`Path(__file__)` 定位），每步有无脚本兜底。

## 四步确定性闭环

### Step 1 一键盘点与多源抽取（1 req）
```bash
python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/ingest.py --case /input/case.json --work /app/work
```
单命令执行：`inspect_case` + `extract_figs`（高清单图、全宽/分栏矢量裁剪、表题页抽 `paper_tables.json`、`paper.txt` 与大纲）+ 合并补充文件与旧稿内容/图表/原图 + 标签冲突检测 + `plan_pages`（受众策略）+ 4 维去重 `qa_draft` → 生成 `reading_map.md`（内嵌 `task.md`、表预览、图清单与策略）与 `starting_deck_dump.json`。
**无脚本兜底**：读 `/input/case.json` + `/input/task.md` + `pdftotext` + `pdftoppm -r 100` + python-pptx 列母版/旧页。

### Step 2 精读原文并写 `deck_spec.json`（1–2 req）
读 `/app/work/reading_map.md` 与 `/app/work/paper.txt`（≤22 KB 通读；更长按节页码直读缺口、方法、主结果、Limitations 与表；旧稿并读 `starting_deck_dump.json`），直接写 `/app/work/deck_spec.json`（顶层可内嵌 `qa_prep`/`result_fill`/`number_allow`/`budget_justify`；按需查 `references/deck-spec.template.json`）：
1. **受众导向叙事（禁按目录搬运）**：按“缺口 → 机制洞见 → 最强对照证据 → 机理分析 → 诚实边界 → 结论”编排；`meta` 填 `duration_minutes`/`goal`/`audience`/`language`/`tradeoffs[]`（**≥2 条可复核取舍**）；内容页填 `time_s`（45–120s/页，`qa_backup` 为 `null`）。
2. **一页一主张与强度对齐**：标题写 takeaway（非“3. Experiments”，不与 `claim` 雷同）；`claim` ≤30 字/18 英文词；`bullets ≤ 4`（≤2 行/条，`标签：` 自动加粗）；`evidence[]` 写物理定位（`Fig.2, p.4`/`Table 1`/`§3.2`/`Slide 4`/补充文件名）+ `strength`（S1/S2/S3/NA），动词强度须匹配档位（S1 才用“优于/显著”，S3 仅用“提示/待验证”）。
3. **视觉多样化（拒文字墙，禁贴未裁整页 `page_p*.png`）**：自动继承模板字体、安全框、Logo 避让与深浅底：`steps`（`[{title,detail}]`）、`kpis`（`[{value,label,delta}]`）、`two_col`（`{left_title,left_bullets,right_title,right_bullets}`）、`table`（`{headers,rows,highlight_row,caption}`，booktabs 三线表）、`chart`（`{type,categories,series:[{name,values}],y_label,caption}`）、`figure`（`{path,region:"right|full",caption}`）。
4. **起始 deck 适配**：`reuse:{"slide_index":0,"action":"keep"|"edit","clear_visuals":false}`；保留原图设 `false`（右/左/底三向自动避让，不误清原图），全页替换设 `true`，拆页复用同索引（自动克隆版式）；重组自动裁废页并出 `starting_deck_diff`（局部修补设 `keep_only_spec_slides:false`）。
5. **口语讲稿与四维问答**：`notes` 写中文 120–180 字 / 英文 90–140 词，**禁照抄 bullets**，讲动机、读图法与尾句过渡；`qa_prep` ≥4 条（`mechanism`/`data`/`baseline`/`limitation`，同口径 `answer_key`/`cite_slide`/`backup_slide`）+ 1 页 `qa_backup`（缺省自动合成双栏板）。

### Step 3 一键构建、门禁、渲染与装配（1 req）
```bash
python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/deliver.py --case /input/case.json --work /app/work --out-dir /app/output
```
自动完成：①模板安全框/字体/Logo 避让构建 PPTX；②四合一分级门（硬错拦截越界/重叠/溢出/未裁整页、悬空定位、不可溯源数、缺图与页数冲突，并出 `== NEXT FIX ==` 一步指引）；③渲染 `<deck>.pdf` + `preview-*.png` + 接触表 `contact_sheet_*.png`（100 DPI）；④导出含**页级证据总表、图表章节覆盖与 diff** 的 `speaker_notes.md`；⑤装配含逐页 `verified_numbers` 与 `tradeoffs` 的 `result.json`。
**无脚本兜底**：python-pptx 填内容 → `soffice --headless --convert-to pdf` + `pdftoppm -r 100` → 核数写 `result.json`。

### Step 4 接触表目检与精修（1–2 req）
读 `/app/work/contact_sheet_*.png`（异常页再读 `preview-*.png`）：查封面层次、标题/主张对齐、卡片/表/图零重叠截断、对比度与来源条。修完重跑 `deliver.py`（≤2 轮）。
**收尾核验**：JSON 合法、`STILL NULL` 为空、`deliverables` 全存在、`gates_failed=0`（残留限制写 `limits`）。

## 异常处置
| 情形 | 处置 |
|---|---|
| 标签与原文冲突 | 以实际文件为准，禁按错误标签补数，`uncertainty` 记录冲突 |
| 无模板/旧稿 | 默认 16:9 学术配色+页码构建 + `uncertainty` |
| 旧稿冗长/无 PDF | `ingest.py` 合成源文并抽原图；`reuse` 裁剪重排并出 diff |
| 无图/图糊/表 | 禁贴未裁整页；用 `table`/`chart`/`steps`/`two_col` + 定位 + `uncertainty` |
| 扫描 PDF / 缺字 | 优先 OCR；否则可见文本+整页图，`status=partial` + `limits`；公式文字化 |
| soffice 失败 | 双 profile 重试；仍败交 PPTX+lint，`rendering` 标 failed |

## 示例 / Example（Schematic Example）
- 输入：`{"case_id":"demo-1","audience":"答辩委员会","duration_minutes":10,"slide_count":10,"template":"/input/tpl.pptx","source_paper":"/input/paper.pdf","output_contract":{"required_fields":["case_id","status","slide_plan","evidence_map","qa_prep","rendering"]}}`
- 流程：`ingest.py` → 读 `reading_map.md` & `paper.txt` → 写 `deck_spec.json` → `deliver.py` → 读 `contact_sheet_1.png`
- 示意输出：`{"case_id":"demo-1","status":"complete","tradeoffs":["压缩相关工作至1页"],"slide_plan":[{"pptx_slide":5,"role":"result","claim":"8192 batch top-1 达 76.3%（+3.2pt）","time_s":75}],"evidence_map":[{"pptx_slide":5,"evidence":[{"type":"figure","locator":"Fig.2, p.4","strength":"S1"}],"verified_numbers":["76.3%","8192"]}],"qa_prep":[{"category":"data","question":"噪声是否伤收敛？","answer_key":"见 Fig.2 与 §3.3 方差","cite_slide":5,"backup_slide":10}]}`

## References（按需读取）
- `references/narrative-patterns.md` — 受众/时长配比、页型组件、claim/表图与强度阶梯
- `references/audience-blueprints.md` — 四类受众蓝图与上台前自审清单
- `references/deck-spec.template.json` — `deck_spec.json` 骨架（含 `qa_prep`/`result_fill`）
- `references/qa-playbook.md` — 4 维追问、answer key 与备用页规则
- `references/template-and-export.md` — 母版/字体/Logo/深浅底、`edit_deck` 与导出排错
- `references/evidence-mapping.md` — locator 规范、引用上限、不改写事实
- `references/checks.md` — 四合一门、豁免规则、双语讲稿、`result.json` schema
