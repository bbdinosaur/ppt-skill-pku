# slides_plan.json Schema（v1）

构建脚本 `build_deck.py` 只认这个 schema。字段全部可选，但每页至少给出 `title` 或 `bullets` 或 `image`。
写计划时保持精炼：要点 ≤6 条/页、每条 ≤2 行，禁止整段论文原文进页面。

```json
{
  "deck_title": "论文短标题（与论文一致，不改写术语）",
  "mode": "create | edit | auto",
  "rationale": "页序规划理由 1–2 句（为什么这样取舍与排序），进入 result.json 供复核",
  "slides": [
    {
      "role": "title|agenda|background|method|result|figure|table|limitation|summary|qa|section",
      "layout_hint": "title|section|content|two_content|picture|quote|blank|auto",
      "title": "页面主张式标题（≤22 字，说结论不说话题）",
      "subtitle": "仅 title/section 页使用",
      "highlights": [{"value": "47.8", "label": "ROUGE-1", "sub": "ArXiv-Long"}],
      "steps": [{"title": "1. 输入编码", "desc": "分块压缩长上下文"},
                {"title": "2. 稀疏路由", "desc": "动态门控保留 12.5%"},
                {"title": "3. 并行解码", "desc": "吞吐提升 3.42x"}],
      "bullets": [{"t": "要点文本", "lvl": 0, "bold": false, "size": 18}],
      "two_col": {"left_title": "可选左栏标题", "left_bullets": [],
                  "right_title": "可选右栏标题", "right_bullets": []},
      "image": {"path": "inputs/figures/fig3.png", "ref": "图3/Fig.3",
                 "caption": "图注逐字取自论文", "missing": false},
      "images": [{"path": "fig_a.png", "ref": "图3a", "caption": "可选并排双图左"},
                 {"path": "fig_b.png", "ref": "图3b", "caption": "可选并排双图右"}],
      "chart": {"type": "bar|hbar|stacked_bar|stacked_hbar|line|pie", "ref": "图2数据",
                 "caption": "数据来源说明", "title": "图表内标题（可选）",
                 "x_title": "横轴（可选）", "y_title": "纵轴（可选）",
                 "y_min": 0, "y_max": 100, "highlight_last": true,
                 "categories": ["完整 ASA", "去门控稀疏率"],
                 "series": [{"name": "ROUGE-1", "values": [47.8, 45.7]}]},
      "table": {"rows": [["指标","本文","基线"]], "header": true, "font_pt": 12,
                "highlight_row": -1, "ref": "表1（可选）", "caption": "表注（可选）"},
      "takeaway": "页底一句核心结论（可选，自动置于来源行上方）",
      "notes": "讲稿备注 2–4 句：这页讲什么、强调哪个数字、过渡到下一页",
      "sources": [{"ref": "图3", "locator": "论文 §4.1 p.5", "quote": "论文原句片段 ≥6 字"}]
    }
  ],
  "edits": [
    {"op": "set_title", "slide": 1, "text": "新标题", "sources": []},
    {"op": "set_body",  "slide": 2, "bullets": [],
     "sources": [{"ref": "图4", "locator": "§4.2", "quote": "论文原句"}]},
    {"op": "set_two_col", "slide": 3, "left_bullets": [], "right_bullets": [], "sources": []},
    {"op": "replace_text", "slide": 2, "old": "旧数字或术语", "new": "新数字或术语", "sources": []},
    {"op": "set_source", "slide": 2, "sources": [{"ref": "表1", "locator": "p.4", "quote": "原文"}]},
    {"op": "set_table", "slide": 3, "table": {"rows": [["指标","新值"]], "highlight_row": -1}, "sources": []},
    {"op": "set_chart", "slide": 3, "categories": ["Base","Ours"], "series": [{"name":"Acc","values":[88.0,92.8]}], "sources": []},
    {"op": "set_image", "slide": 4, "path": "/app/work/figs/fig2.png", "ref": "Fig. 2", "caption": "Updated diagram", "sources": []},
    {"op": "set_notes", "slide": 3, "text": "新备注"},
    {"op": "replace_slide", "slide": 4, "slide_spec": { /* 原地替换第 4 页，不移动后续页码 */ }},
    {"op": "insert_at", "at": 4, "slide_spec": { /* 同 slides 结构 */ }},
    {"op": "append", "slide_spec": { /* 同 slides 结构 */ }},
    {"op": "delete", "slide": 4},
    {"op": "move",   "slide": 5, "to": 2}
  ],
  "uncertainty": ["无法从材料确认的信息逐条列出（禁止编造）"]
}
```

## 规则

0. **标题符号**：标题避免 `→`、`·` 等生僻符号（PDF 文本抽取可能与渲染不一致），
   数学上标用 `^2` 或文字表述；对比关系用「由 A 降至 B」表达。
1. **证据映射（claim_sources）**：每个非 title/section/agenda/qa 页必须 ≥1 条 `sources`；
   `quote` 必须逐字取自论文/材料（脚本会做子串核验，编造会被 `quote_verification` 抓住）。
2. **数字保真**：页面上每个数字都要能在材料中找到；禁止取整、禁止"约 XX%"这种改写；
   单位、符号、缩写与论文完全一致。脚本会做数字回查（小数或 ≥3 位数字），
   回查不中的数字会被 `number_fidelity` 抓住。
3. **原生图表（数据绑定，创新点）**：当材料中存在某图的**完整数值**（表格/正文数字）时，
   优先用 `chart` 生成**原生可编辑 PPT 图表**（数据可被评审查验）；`categories/series.values`
   逐项取自材料，`number_fidelity` 会核验图表数据。位图/无数据的图仍走 `image` 只引用原图。
4. **图片**：只能引用真实存在的图片文件（PDF 抽取图或材料附带图）。位图/无数据表的图
   只引用原图，**不得凭印象重绘**；`caption` 逐字取论文图注。
5. **按论文类型选页序骨架**（严格服从 workorder `slide_count`/`min_slides`/`max_slides`）：
   - *实验论文*：Title → Agenda/Motivation → Method(1–2页) → Main Results(表/`highlights` 指标卡) → Key Figures/Ablation(原生 `chart` 或原图 `image`) → Efficiency/Analysis → Limitations → Takeaways/Q&A。
   - *综述论文*：Title → Scope/Taxonomy → Paradigm Comparison(`two_col`/Table) → Key Benchmarks/Trends → Open Challenges → Summary。
   - *理论/方法论文*：Title → Problem Setup → Core Formulation/Assumptions → Proof/Algorithm Sketch → Empirical Validation → Scope & Limitations。
   听众 3 分钟内应能从 agenda 看出主线。
6. **页数**：严格按 workorder 的 `slide_count`/`min_slides`/`max_slides`（支持 `"10-12"` 区间）；没有明确数字时按 10–14 页。
7. **edit 模式**：只改 workorder 要求的页；`edits` 之外不动任何页，绝不改母版。
   局部改数字/术语优先用 `replace_text`（原地保留全部格式；`slide: "all"` 可全稿统一术语）；
   原地整页重做优先用 `replace_slide`（避免 `delete` 导致后续页码偏移）；
   指定位置插页用 `insert_at`；`set_title`/`set_body`/`replace_text`/`set_source` 均可携带 `sources`。
   构建后自动产出 `slide_diff.md` 逐页变更对照表。
8. `uncertainty` 至少在材料缺口时给出 1 条；没有缺口就给空数组，不要硬编。
