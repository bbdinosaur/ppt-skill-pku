# result.json Schema — 字段字典与状态口径

## 顶层结构
| 字段 | 类型 | 口径 |
|---|---|---|
| `case_id` | str | **逐字复制 case.json 的 case_id** |
| *(output_contract 每个字段)* | — | 逐字并入顶层；优先级：`--fills` 实填 > 自动交付物/标准段映射（含 `template_path`/`template_file`、`sources`/`source_files`/`provenance`、`uncertainties`、`theme_mapping`/`layout_mapping`/`normalizations`） > 保留已有 `result.json` 非空手写字段 > `null` + `uncertainty`（禁编造）；重跑时对已存在 `result.json` 执行非覆盖合并，去重合并 `sources`/`uncertainty`/`limits`/`change_log` |
| `status` | str | `complete`（`deck.pptx` 非空有效且 `slide_count>0`、`deck.pdf` 有效导出且 `pages_match==true`；启发式几何告警与疑似空白页不误降级，而是如实记入 `checks`/`limits`/`uncertainty`）/ `partial`（`deck.pptx` 有效但 PDF 缺失损坏或页数不一致）/ `failed`（`deck.pptx` 缺失损坏或 0 页） |
| `deliverables` | obj | `pptx`、`/app/output/deck.pdf`、`previews_dir`、`previews[]`、`build_report`、`plan`、`render_log`（全部绝对路径） |
| `deck_summary` | obj | `slide_count`、`title`、`layouts_used[]`、`template`、`slide_size`、`audience`、`duration_min`、`language`、**`placeholder_chain[]`**（逐页 layout + title_ph/body_ph idx+type，自动由 `build_report` 生成） |
| `evidence_map` | list | **每页一条**（拆页后最终页号为准，由 `build_report` 权威合成），见下 |
| `rendering` | obj | `pdf`、`previews_dir`、`dpi`、`pages_match`、`pdf_page_size_in`、`size_match`、`notes_slides`、`title_text_ok`、`blank_pages[]`、`theme_fonts`、`pdf_fonts_used[]`、`font_substitutions[]`、`font_warnings[]` |
| `checks` | obj | `geometry_critical`、`text_overflow_critical`、`overlap_critical`、`aspect_distortion`、`missing_notes`、`consistency_warnings`、`inheritance_summary`、`heuristic_alerts` |
| `uncertainty` | list[str] | 不确定项 + 影响范围 + 采用假设（材料缺失/无法判断/运行失败分列，不混写） |
| `change_log` | list[str] | 真实操作与精修记录（含模板归一化、布局迭代） |
| `limits` | list[str] | `partial`/`failed` 或交付缺项时：缺什么、为什么、如何补 |

## evidence_map 单条结构
```json
{
  "slide": 4,
  "kind": "content",
  "claim": "定量结果：耗时降 32.0%，覆盖率保持 95.6%",
  "sources": [
    {"locator": "Table 1, p.5", "kind": "paper-table",
     "quote": "Proposed achieves 100.9 min (-32.0%) with 95.6% coverage"}
  ],
  "figure_refs": [
    {"caption": "Figure 1: ...", "source": "Figure 1, p.3", "asset": "/app/work/figs/fig1_p3.png"}
  ],
  "notes_present": true,
  "time_hint_s": 110
}
```
- `kind` ∈ `paper-section / paper-figure / paper-table / paper-abstract / paper-title / material-other`（`write_result.py` 自动按 locator 规范推断，也可在 `plan.json` 显式指定）。
- **数值/结论页强烈建议附 `quote`（≤15 词逐字原文）**；数字、正负号、单位必须与源论文逐字一致。

## 交付前自检清单
1. `result.json` 合法 JSON，含 `case_id` 与全部 `output_contract` 字段；终端 `NULL_CONTRACT_FIELDS` 中凡可推导的均已通过 `--fills` 填实。
2. `/app/output/` 下齐备 `deck.pptx`、`deck.pdf`、`previews/page-XX.png`（108 DPI）、`result.json`。
3. `deck_check.json` 中审阅 `critical` 与 `heuristic_alerts`（尽量零出界、零 2D 重叠、零溢出、零拉伸、无 <12pt 正文；确属启发式误报者以视觉预览为准并在 `uncertainty`/`limits` 留痕）。
4. `render_log.json` 中 `pages_match == true`、`title_text_ok == true`，目检预览确认无真空白页，`pdf_fonts_used` 与 `font_warnings` 如实记录。
5. `evidence_map` 条数 == `deck_summary.slide_count` == `placeholder_chain` 条数；`change_log` 反映真实精修轮次。
