# plan.json Schema & Template Mapping

## plan.json 完整 schema（build_deck.py 输入）
```json
{
  "template": "/input/template.pptx",          // 必填（无模板时填 "none"）
  "output": "/app/output/deck.pptx",           // 必填，交付路径
  "audience": "domain researchers",            // 可选，写入 deck_summary
  "duration_min": 15,                          // 可选，自动按页型分配讲稿建议用时
  "language": "zh",                            // 可选
  "expected_slides": 8,                        // 可选：题目固定页数时自动锁页（禁自动拆页超标）
  "footer": "Paper Short Title | Venue",       // 可选：填入模板自带 FOOTER 占位符（跳过封面）
  "slide_numbers": true,                       // 可选：填入模板自带 SLIDE_NUMBER 占位符
  "change_log": ["v1 initial build"],          // 可选：传递至 result.json
  "ladder": {"title": 36, "subtitle": 20, "section": 30, "slide_title": 28,
             "body": [18, 16, 14, 13], "caption": 11.5, "table": [13.5, 12.5, 11.5, 10.5]},
  "slides": [
    {"kind": "title", "layout_hint": "标题幻灯片|Title Slide",
     "title": "论文标题", "subtitle": "作者 · 单位 · 会议/场合",
     "notes": "开场白", "evidence": [{"locator": "Title, p.1", "quote": "verbatim title"}]},

    {"kind": "section", "layout_hint": "节标题|Section Header",
     "title": "02 方法与架构", "subtitle": "核心模块概览",
     "notes": "过渡句", "evidence": [{"locator": "§2, p.2"}]},

    {"kind": "content", "layout_hint": "标题和内容|Title and Content",
     "title": "主张式标题(结论句，≤46 有效字符)",
     "bullets": [
       {"text": "**加粗导语**：支撑结论与关键数字", "level": 0},
       {"text": "细节说明（缩进 1 级）", "level": 1}
     ],
     "notes": "讲稿 2-3 句",
     "evidence": [{"locator": "§2.1, p.3", "kind": "paper-section", "quote": "exact <=15 words"}]},

    {"kind": "content", "layout_hint": "标题和内容|Title and Content",
     "title": "定量结果：耗时降 32.0%，覆盖率保持 95.6%",
     "bullets": [{"text": "**核心对比**：5 组随机种子均值，显著优于贪心基线", "level": 0}],
     "table": {
       "headers": ["Method", "Time (min)", "Coverage (%)", "Delta"],
       "rows": [["Fixed Route", "148.2", "97.9", "Baseline"],
                ["Greedy", "121.6", "94.7", "-17.9%"],
                ["Proposed", "100.9", "95.6", "-32.0%"]],
       "highlight_rows": [-1],
       "caption": "Table 1: Main benchmark comparison", "source": "Table 1, p.5"
     },
     "notes": "重点讲解最后一行高亮数据。",
     "evidence": [{"locator": "Table 1, p.5", "kind": "paper-table", "quote": "Proposed 100.9 min (-32.0%), 95.6%"}]},

    {"kind": "two_col", "layout_hint": "两栏内容|Two Content|Comparison",
     "title": "双栏对比：现有基线 vs 本文方案",
     "left": {"heading": "现有基线局限", "bullets": [{"text": "**固定路线**：过度采样低风险区", "level": 0}]},
     "right": {"heading": "本文方案", "bullets": [{"text": "**GRU 置信估计**：动态聚焦退化区域", "level": 0}]},
     "notes": "左右对照讲差异。",
     "evidence": [{"locator": "§1-§2, p.1-2", "kind": "paper-section"}]},

    {"kind": "figure", "layout_hint": "图片与标题|Two Content|Title and Content",
     "layout_mode": "auto",   // "auto" | "side_by_side" | "stacked"
     "title": "收敛曲线：本文方法跟踪滞后更低",
     "bullets": [{"text": "**关键观察**：相位滞后显著小于固定路线基线", "level": 0}],
     "images": [{"path": "/app/work/figs/fig1_p1.png",
                 "caption": "Figure 1: Proposed vs baseline", "source": "Figure 1, p.1"}],
     "notes": "先指横纵轴，再强调橙线贴合真值。",
     "evidence": [{"locator": "Figure 1, p.1", "kind": "paper-figure", "quote": "tracks true signal with lower lag"}]}
  ]
}
```
注：
- `two_col` 的 `left` 与 `right` 均对称支持 `bullets`、`images`、`table` 或原生可编辑 `chart`，并原生配对 PowerPoint `Comparison` 的 4 占位符结构（左右栏标题框 + 左右正文框）。
- 原生可编辑数据图表 `chart` 字段（可用于 `content`、`two_col` 左右栏或 `figure`）：
  ```json
  "chart": {"type": "column", "title": "Runtime Comparison", "unit": "min",
            "categories": ["Fixed", "Greedy", "Proposed"],
            "series": [{"name": "Time", "values": [148.2, 121.6, 100.9]}],
            "show_values": true, "source": "Table 1, p.5"}
  ```
  支持 `column`/`bar`/`line`，自动按模板 `accent1..accent6` 配色并适配深/浅底轴字色；可选通过 `check_deck.py --paper-text` 进行只读数值参考核对。
- 图片默认启用智能修整（裁除 >8% 纯白外边、深底上透明黑线 PNG 自动铺洁净白卡；可在单图设 `"trim": false` 关闭）；`evidence` 也向后兼容纯字符串列表 `["§2.1 p.3"]`。

## 版式选择与占位符保护规则
1. `layout_name`（精确名）> `layout_hint`（中英文关键词）+ 占位符结构打分：
   - 标题占位符在页面上部 (`top < 2.5in`) 加分，中下部标题与竖排 (`Vertical`/`竖排`) 版式降权。
   - `two_col` 优先选择含 ≥2 个文本占位符或 4 占位符 `Comparison` 的版式（自动将 `heading` 填入顶部短框、正文填入下方主体框）。
   - `figure` 单图+要点优先选宽幅可容纳要点的双栏版式（规避底部标题/窄图注条的 `Picture with Caption` 陷阱）；双图/纯图优先选 `Title Only` 干净画布。
2. 占位符角色严格分离：`TITLE/CENTER_TITLE`（标题）、`SUBTITLE`（副标题，支持多行主次字号）、`BODY/OBJECT/TEXT/CONTENT`（文本正文）、`PICTURE/CLIP_ART`（图片框，**绝不写入正文文本**）、`FOOTER/SLIDE_NUMBER/DATE/HEADER`（页眉页脚）。内容区自动贴合模板正文占位符外包框并避开母版/版式边缘装饰条与校徽横幅。
3. **不设字体名**，只设字号/加粗/行距，确保母版主题字体完整继承；背景检测纯只读 XML 零突变；页内未使用的空内容/图片占位符在落盘前自动清理。
4. 深色底板自动切换浅色图注、深蓝灰原生表格配色、提亮栏标题与透明暗线图白卡保护，表头文字按 `accent1` 亮度自动选白/深色，保持高对比度。

## 模板异常与起始 deck 决策树
- **`.potx` 模板**：自动改写临时副本的 `[Content_Types].xml` 打开，输出标准 `.pptx` 并记入 `normalizations`。
- **模板自带示例页**：默认清除（避免页号错位）；确需保留时设 `"keep_template_slides": true` 并在 `uncertainty` 注明页号偏移。
- **起始 deck（旧 PPT）**：先运行 `inspect_deck.py OLD.pptx --json deck_inventory.json --draft plan_draft.json --imgdir deck_imgs`（完整提取占位符、浮动文本框、原生表格、组合图形、图片与备注）。比对 `master_fingerprint` 与目标模板：同母版保持版式映射，异母版按本管线在目标模板上重建，保留旧内容与图表并补齐论文证据。
