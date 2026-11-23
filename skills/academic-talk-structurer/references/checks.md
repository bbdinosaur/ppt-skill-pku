# 自检体系与豁免机制（check suite policies）

## 0. 一键闭环与四合一门
- S0+S1 一键盘点与抽取：`python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/ingest.py --case /input/case.json --work /app/work`
- S3–S6 一键构建+自检+渲染(含接触表)+讲义+result.json：`python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/deliver.py --case /input/case.json --work /app/work --out-dir /app/output`
- 单独重跑四合一门：`sh ${SKILL_DIR:-/skills/weekly-skill}/scripts/preflight.sh <pptx> <workdir> <case.json>` = 下方 [1]-[4] 依序执行并汇总退出码。
## 1. 检查矩阵
| 工具 | 检查 | 退出码 1 的含义 | 处置 |
|---|---|---|---|
| check_layout.py [--plan] | 越界/溢出/重叠(含表格/图表)/图片拉伸 + 建议级告警：低DPI(<90)/原始LaTeX与乱码/双语讲稿长度/总时长/泛化或重复标题/标题与主张雷同/文字墙密度/主张过长/S3夸大/S1定量校准/无视觉组件/空目标与取舍 | 存在硬布局或讲稿硬错 | 删内容不强缩；补讲稿至 ≥35 语速单位；按建议告警精修视觉与措辞 |
| check_numbers.py | 全页(含表格/图表/卡片)数字能否在源文本+补充文件+`paper_tables.json` 精确 token 溯源（含 `10^k`/PDF Unicode `10⁻⁴`/`10−3` 科学计数、`thousand`/`million`/`billion`、`min↔s`/`ms↔s`/`GB↔MB`/`KB↔MB`/`TB↔GB`/`µs↔ms`/`hr↔min`、`pp`/百分点、分类型衍生差值/倍率/相对降幅）；同步对 `qa_plan.json` 回答要点做建议级数值口径核验 | 存在不可溯源数字 | 改回源值 / 降级 S3+uncertainty / 白名单 |
| check_evidence.py | Fig/Tab/Sec/§/Slide/页码/补充文件名 locator 校验（Eq/Alg/Appendix 编号差异与模糊定位为建议告警，防 LaTeX 数学字体断行误报） | 悬空引用 | 必须修复：改 locator 或删引用，禁止豁免 |
| plan_pages.py check | 受众配比（容差 max(10%, 0.55/内容页数)）/qa 备用页/页数一致 | 配比偏差 | 调整 spec 或 justification 豁免 |

## 2. 数字白名单（number_allow.json）
格式：`[{"value":"5.5","reason":"derived from Table 2: 75.2-69.7"}]`
仅用于：推导差值、图表读取值、单位换算。每条 reason 必须给出可复核出处，并同步写进 result.uncertainty。调用：`--allow /app/work/number_allow.json`。

## 3. 配比豁免（budget_justify.json）
格式：`{"share[method]":"理由","slide_count":"理由","qa_backup":"理由"}`（key 与输出 issue 的 [key] 一致）。理由必须指向客观约束（如页数-时长冲突），同步写 uncertainty。调用：`--justif /app/work/budget_justify.json`。

## 4. 双语讲稿 lint 规则
语速单位 = `CJK 字符数 + 1.5 × 英文词数`（1 英文词≈1.5 汉字口播耗时）。内容页（title/thanks/qa_backup 除外）：≥40 单位（约 ≥40 汉字或 ≥27 英文词）；≤ `4.8 × time_s`（约 ≤288 字/分或 ≤192 英文词/分）；必须有 `time_s`。qa_backup 页不计时。

## 5. 扫描件政策
inspect.json 的 `likely_scanned`/`low_text_pages` → 优先 OCR 工具；不可用则基于可见文本构建，status=partial，uncertainty 写明"文本层缺失，证据定位以页码为准"。

## 6. result.json 默认 schema（`build_result.py` 自动对齐 `output_contract` 同义字段）
`case_id, status, sources[], deliverables{deck,pdf,previews,speaker_notes}, report_goal, audience, audience_adaptation{audience_class,category_counts,avg_seconds_per_content_slide,budget_warnings}, tradeoffs[], timing_summary{total_timed_s,budget_s,total_slides,backup_slides}, evidence_coverage{content_slides,slides_with_evidence,coverage_pct,strength_counts,dangling_locators,untraced_numbers}, source_coverage_audit{source_items_detected,included,omitted_or_deferred}, slide_plan[{pptx_slide,role,title,claim,time_s,notes_chars,visuals,evidence_locators}], evidence_map[{pptx_slide,role,title,claim,evidence[{type,locator,strength}]}], qa_prep[{question,answer_key,cite_slide,backup_slide}], qa_integrity{total,backup_slides,out_of_range}, rendering{consistency,fonts,non_embedded_fonts,render_dpi,pdf_pages,previews_count,layout_lint,lint_warnings,low_res_figures,symbol_warnings,number_check,evidence_check,budget_check}, uncertainty[], changes[], limits[]`
- 门禁分级：硬错（`exit 1`）仅拦截真实几何重叠/溢出、悬空图/表/页码定位、不可溯源科学指标、缺图与未说明的硬页数冲突；照抄讲稿、S3 强动词、受众配比偏移与低 DPI 为建议级警告（`exit 0`，写入 `result.json` 供审阅）。
- 支持在 `deck_spec.json` 顶层直接内嵌 `qa_prep`、`result_fill`、`number_allow`、`budget_justify`（`deliver.py` 自动同步到 `/app/work/*.json`），并内置无损 schema 规范化（兼容 `kpis` 别名、`two_col` 嵌套结构、`chart` `labels/datasets`、`qa_backup` 自动 `time_s=null` 与默认证据档位）。
- 若 `case.json` 领域/主题标签与实际源文零重合，`ingest.py` 自动触发 `LABEL MISMATCH ALERT` 且 `build_result.py` 自动写入 `result.uncertainty`（确保诚实记录不遗漏）。
