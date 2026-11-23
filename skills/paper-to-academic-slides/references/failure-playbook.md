# 异常处置手册（Failure Playbook）

原则：**识别 → 降级 → 诚实记录**。任何异常都要落到 result.json 的
`gaps_and_uncertainty`，status 相应取 `partial` 或 `failed`，绝不静默编造。

| 异常 | 识别方法 | 处置 |
|---|---|---|
| PDF 加密 | pymupdf `needs_pass` / 打开报错 | 尝试空密码；仍失败→用平台 OCR/pdf 工具；再失败→status=failed 并说明 |
| 扫描件无文本层 | `paper_context.pages` 总字符 < 500 | 用平台 OCR 工具重建文本；OCR 不可用→仅按图片页组织，notes 标注来源为图 |
| 公式是图片无编号 | 题注索引缺 equation | 公式页按 `type:"page"` 引用；不手抄推导 |
| 图抠不出/图损 | `figures` 为空或文件 <1KB | 检查是否矢量图：改用整页 clip 渲染裁剪；仍失败→页面上以文字+表号引用，gaps 记录 |
| 题注跨页/编号缺失 | captions 与 figures 对不上 | 引用降级为 `page` 级；notes 注明"编号以论文原文为准" |
| 模板占位符不足 | build_report.warnings 出现 textbox 落级 | 接受降级；若 >3 页落级，改用更简单的 layout（title+content） |
| 起始 deck 页数与计划不符 | reuse 模式 warnings | 少→追加新页；多→裁掉尾部空页；保持每页证据行 |
| 页数超限 | 计划页 > max_slides | 按裁剪优先级合并：消融并入主结果、方法细节并入总览；**不删封面/主结果/结论** |
| 长标题溢出 | check_layout out_of_bounds/text_overflow | 标题压缩成主张短语（≤14 词）；禁止缩字号低于 20pt 的标题 |
| 正文字号过低(<11pt) | check_layout font_too_small | 拆页或删 bullets；11pt 是底线 |
| soffice 缺失/转换失败 | finalize_output gaps | 交付 PPTX+slides.json，`checks.pdf_exported=false`+原因；绝不断言已导出 |
| 中文字体缺失渲染方块 | PNG 目检出现豆腐块 | 换模板内置字体/系统 Noto 系；仅设置字号不设字体名可规避 |
| case.json 字段缺失 | resolve_case.missing 非空 | 用 discover 到的 /input 文件继续；constraints 用默认（≤12 页、中文），gaps 记录 |
| 时间过半未完成 S3 | 请求计数 >32 或耗时 >450s | 立即用当前 slides.json 跑 S3–S5 出 partial 版本，再迭代 |

## 部分交付的最小闭环
最坏情况（脚本全不可用）：交付
1. `slides.json`（手工整理的主线+证据映射）
2. `result.json`：status=partial、gaps 写明"无 python-pptx/soffice，无法产出 PPTX/PDF"
并保证 evidence_map 逐页完整——保证可核验性优先于版面。
