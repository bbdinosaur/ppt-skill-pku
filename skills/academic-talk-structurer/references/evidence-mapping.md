# 证据映射与诚实规则

## 1. evidence 条目 schema（deck_spec / slide_plan / result.evidence_map 共用）
```json
{"type": "figure|table|section|equation|page", "locator": "Fig.2, p.4 | Table 1 | §3.2 | Eq.(5) | p.7", "strength": "S1|S2|S3|NA", "note": "可选：支撑哪个 bullet"}
```

## 2. 定位规则
- locator 必须指向源材料真实存在并可复核的位置：图/表号优先，其次章节号+页码；页码用 PDF 物理页。
- 抽取的图文件名（figs/manifest.json）与 locator 一并保留，页面 caption 写原图题（≤20 字，可截断加 …）。
- 一个 locator 只支撑同页主张相关的事实；跨页事实不合并引用。

## 3. 原文引用与改写红线
- 直接引原文 ≤25 词/条，加引号；其余一律转述，但数字/单位/符号逐字保留。
- 禁止改写：结论方向（显著→不显著）、样本量/轮次/参数、图表数值。
- 源材料缺失或不可判的内容：页面不出现；如必须提及，strength=NA 且 notes 声明"源材料未验证"。

## 4. 不确定性落盘（result.uncertainty 示例条目）
- "模板缺失，使用默认母版"
- "源 PDF 为扫描件，图 3 未能抽取，页面用文字证据替代"
- "时长约束与给定页数冲突，第 9、10 页降级为备用页"
每条 uncertainty 对应一个可复核的客观事实，不写主观感受。

## 5. 变更记录（result.changes）
按版本追加，如 `["v1: 初版 12 页，per-timed-slide 75s", "v2: 修复 p5 图片越界，压缩 p7 bullets"]`；每次重建 deck 后更新。
