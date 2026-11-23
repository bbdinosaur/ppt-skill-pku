# slides.json Schema（build_deck.py 的输入契约）

顶层：`{"meta": {...}, "slides": [...]}`

## meta（可选，写入 result.json）
| 字段 | 类型 | 说明 |
|---|---|---|
| deck_title | str | 封面主标题（论文标题或其忠实缩写） |
| authors | str | 作者/单位行（来自论文，不得杜撰） |
| audience | str | 来自 case.json 的受众/时长约束 |
| language | str | zh / en / mixed |
| reuse | bool | 输入为起始 deck 时 true |
| max_slides | int | 页数上限（来自 case.json） |

## slides[] 每页
| 字段 | 必填 | 类型 | 说明 |
|---|---|---|---|
| kind | 否 | str | `cover`（首页）/ `content`（默认）/ `closing` |
| title | 是 | str | **主张句**：该页要证明什么。≤ 2 行；禁止"方法""实验"类栏目词 |
| subtitle | 否 | str | 封面页可选副标题（当封面 `bullets` 为空时自动渲染作者/单位/会议行） |
| bullets | 否 | list | 每项 `{"level":0/1, "text":"≤22词或30汉字"}`；3–6 条为宜；支持 `**关键词/数值**` 原生加粗 |
| right_bullets | 否 | list | 纯文本对比/局限页可选右栏要点（同 `bullets` 格式），自动双栏+竖分隔线排版 |
| left_heading / right_heading | 否 | str | 双栏对比页可选左/右栏主题色小标题 |
| kpis | 否 | list | 主结果/总结页可选 2–4 张指标卡 `[{"value":"89.1%","label":"Top-1 Acc (Table 2)"}]`，数值须逐字出自论文 |
| takeaway | 否 | str | 内容页底部可选单行左饰条要点总结（≤18词/28字），支持 `**bold**`，自动避让正文/图/表/页脚 |\n| source_slide | 否 | int | `--reuse` 模式下可选 1-based 起始 deck 页码，用于精准挑选/重排起始页版式（留空则按需新建页） |\n| keep_existing_media | 否 | bool | `--reuse` 模式下设为 `true` 保留起始 deck 当前页原有原生图表/组合图/图片 |\n| image | 否 | str | 图片绝对路径（`figs/*.png`、补充图或 `starting_deck_figs/*.png`） |
| images | 否 | list[str] | 多图并列：2 张在宽区域自动左右横排，3 张纵向堆叠 |
| image_caption | 否 | str | 可选图注（10pt 居中紧贴图下），如 `"Fig. 1: 方法框架示意图（非实验数据）"`；多图也可用 `image_captions` 列表 |
| keep_existing_media | 否 | bool | 仅 `--reuse` 模式有效：`true` 时保留起始 deck 该页原有原生图表/示意图，正文自动避让 |
| notes | 建议必填 | str | 讲稿 60–150 字：过渡 + 强调点 + 追问补充证据；**不复述页面文字** |
| evidence | 是 | obj | 见下 |

### evidence 对象
| 字段 | 类型 | 说明 |
|---|---|---|
| paper_refs | list | 每项 `{"type":"figure/table/section/page/equation", "id":"Fig. 3", "page":5, "quote":"≤25词原文片段"}`；至少 1 条 |
| slide_line | str | 页脚溯源一行：`"证据: Fig. 3 (p.5); §4.2"` |

### 铁律
1. `quote` 必须是论文原文连续片段，数字/单位/方法名**逐字保留**。
2. 每页至少 1 条 paper_refs；封面可引 `{"type":"page","id":"title","page":1}`。
3. 示意图（方法框架图）标注 `type:"figure"` 但 notes 里注明“示意图，非实验数据”。
4. bullets 不出现论文未出现的结论；比较句必须能回溯到具体表/图。

## 最小示例（示例 / Example）
```json
{"meta": {"deck_title": "Example Paper Title", "language": "zh", "max_slides": 10},
 "slides": [
  {"kind": "cover", "title": "Example Paper Title",
   "bullets": [{"level": 0, "text": "作者 · 单位"}],
   "notes": "开场30秒：一句话说清问题与贡献。过渡：先看为什么难。",
   "evidence": {"paper_refs": [{"type": "page", "id": "title", "page": 1}],
                "slide_line": "证据: p.1 标题页"}},
  {"title": "X 在 Y 上提升 12.3% 且不增加延迟",
   "bullets": [{"level": 0, "text": "主表 Table 2：精度 89.1%（+12.3%）"},
               {"level": 0, "text": "延迟持平 41ms（Fig. 5）"}],
   "image": "/app/work/paper/figs/fig_p06_0.png",
   "notes": "强调 +12.3% 的对照条件与表注；若被问基线选择，引用 §5.1 第一段。",
   "evidence": {"paper_refs": [
      {"type": "table", "id": "Table 2", "page": 6, "quote": "89.1 ± 0.4, +12.3%"},
      {"type": "figure", "id": "Fig. 5", "page": 7}],
     "slide_line": "证据: Table 2 (p.6); Fig. 5 (p.7)"}}
 ]}
```

### table 对象（可选，与 image 二选一；渲染为原生可编辑表格）
| 字段 | 类型 | 说明 |
|---|---|---|
| headers | list[str] | 表头，逐字取自论文表头行 |
| rows | list[list[str]] | 数据行，逐字拷贝；>6 行自动截断并注明"另有 N 行见论文原文" |
| note | str | 表注（单位/条件），如 "Acc: top-1 (%)，Lat: ms/query" |

规则：单元格文本不做改写/换算；表号写入 evidence.paper_refs（type=table）；
大表优先抄关键 2–4 行，完整性靠 slide_line 指回论文页码。
