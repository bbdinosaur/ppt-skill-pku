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
      "bullets": [{"t": "要点文本", "lvl": 0, "bold": false, "size": 18}],
      "two_col": {"left_bullets": [], "right_bullets": []},
      "image": {"path": "inputs/figures/fig3.png", "ref": "图3/Fig.3",
                 "caption": "图注逐字取自论文", "missing": false},
      "chart": {"type": "bar|line|pie", "ref": "图2数据",
                 "caption": "数据来源说明", "title": "图表内标题（可选）",
                 "categories": ["完整 ASA", "去门控稀疏率"],
                 "series": [{"name": "ROUGE-1", "values": [47.8, 45.7]}]},
      "table": {"rows": [["指标","本文","基线"]], "header": true, "font_pt": 12},
      "notes": "讲稿备注 2–4 句：这页讲什么、强调哪个数字、过渡到下一页",
      "sources": [{"ref": "图3", "locator": "论文 §4.1 p.5", "quote": "论文原句片段 ≥6 字"}]
    }
  ],
  "edits": [
    {"op": "set_title", "slide": 1, "text": "新标题"},
    {"op": "set_body",  "slide": 2, "bullets": [],
     "sources": [{"ref": "图4", "locator": "§4.2", "quote": "论文原句"}]},
    {"op": "set_notes", "slide": 3, "text": "新备注"},
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
5. **页序叙事**：title → agenda/背景 → 方法(压到 1–2 页) → 结果(每图可单独一页，先主结果)
   → 局限 → 总结/Q&A。听众 3 分钟内应能从 agenda 看出主线。
6. **页数**：严格按 workorder 的 `slide_count`/`max_slides`；没有明确数字时按 10–14 页。
7. **edit 模式**：只改 workorder 要求的页；`edits` 之外不动任何页，绝不改母版。
   `set_title`/`set_body` 操作同样可携带 `sources`（改写已有页时保持证据链完整）。
8. `uncertainty` 至少在材料缺口时给出 1 条；没有缺口就给空数组，不要硬编。
