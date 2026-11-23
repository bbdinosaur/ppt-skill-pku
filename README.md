# academic-slides-pku — 学术 PPT 生成与编辑 Skill

面向科研汇报场景（组会、答辩、开题/中期、同行评审）的学术 PPT 生成与编辑 Skill：
将固定论文或研究材料组织为**原生可编辑、模板继承、数字与引文可追溯、多轮迭代可回退、导出可复核**的演示稿
（原生 PPTX + 真实 PDF 导出 + 100 DPI 逐页 PNG/缩略拼图 + 页级证据映射 + 逐页 Diff + 版本历史 + 合规审计报告）。

## 1. 核心架构与方法创新（两步极速确定性管线）

相比通用单次生成容易出现的整段堆字、数字抄错/取整、覆盖模板母版、导出塌版、编辑误伤他页及长对话超时，本 Skill 以 **8 个自包含、无状态、路径可移植的 Python 模块** 提供工程化闭环，通过 `scripts/run_pipeline.py` 将全流程收拢为 `prepare` + `deliver` 两条确定性命令（亦可按需单步复用任意模块）：

| 模块 | 核心职责与关键改进 |
|---|---|
| `scripts/run_pipeline.py` | **两步一键编排器**：`prepare` 清点全量输入、写出保底 `partial` 版 `result.json` 并在 stdout 直接输出紧凑 JSON（含诚实 `read_more_hints`）+ 高保真可读 `TASK.MD`（≤7.5 KB）/ `MATERIAL DIGEST`（≤13.5 KB）/ `START DECK DUMP`（≤6.5 KB，总计 ≤28 KB）；`deliver` 串联局部补丁热修（`--patch-slide` / `--patch-file`）→构建→13 项检查→**内置 2 轮自动修补**（排版紧缩、表格限高、`min_font` 钳制、引文逐字对齐、缺源/缺备注补齐，绝不改写幻灯片数字）→PDF/PNG 导出→终版组装、`$WORK/history/vN/` 版本归档与 Pass ≥2 启发式边缘诚实转记录，支持 4–8 步极速闭环并严控最多 1 次精修重跑。 |
| `scripts/inspect_case.py` | 递归容错解析 `case.json` / `task.md` / `material.json` / `slide-workorder.json`（支持区间页数如 `"10-12"` 与显式路径）；多 PDF 全文抽取 + 论文导航地图 + 兼容罗马数字/附录编号（`TABLE I`/`Fig. S1`/`Extended Data Fig. 2`）的标题门控 `find_tables()` 精准抽表（最多 8 表×12 行×10 列）+ 内嵌位图与双栏论文分栏感知矢量图裁切；完整 `material.json`（≤9.5 KB 原样保留）+ 首页摘要原文 + 全量图表原注 + 按句切分全篇关键句 + 补充 `.csv`/`.tsv`/`.md`/`.txt` 汇总入 `material_digest.md`；edit 模式自动生成含完整表格（前 14 行）与图表序列的 `start_deck_dump.md`。 |
| `scripts/build_deck.py` | **create 模式**：只读继承真实 `target-template.pptx` 版式/配色/主题字体（永不改母版、不覆盖 `font.name`），内置 `normalize_plan()` 容错归一常见 LLM 字段别名（单对象 `sources`/`bullets[].text`/`table.headers`/`highlights[].metric`/`steps[].label`），优先填充模板原生 `BODY`/`Two Content` 槽位（首词自动加粗、Markdown `**粗体**` 自动转原生加粗 run、剥离 `` `代码` `` 并将 LaTeX 数学如 `10^{-6}`/`R^2`/`\pm`/`\alpha`/`CO_2` 转为出版级上下标与希腊字母、双栏 `two_col` 无圆点主题色栏标题 + 稀疏页 20pt/1.18 舒适节奏与细中线分隔、稀疏纯文本页 22pt/1.18 节奏、封面/Q&A 副标题与要点互不覆盖、自动清理未使用空正文/副标题占位符防幽灵框）；仅目录页升级为左对齐编号目录条（`01..06`）、仅无正文槽位版式启用要点卡片；支持扁平无阴影圆角顶/左主题色条卡片（`highlights` KPI 卡、`steps` 2–5 步横向矢量流程图含动态大字号、`takeaway` 结论条）、数据绑定原生可编辑图表（`bar`/`hbar`/`stacked_bar`/`line`/`pie` 含饼图主题色盘、`13.5pt` 标题/`11.5pt` 轴标题/`size=7` 折线标记与 `"85.4%"`/`"1,240"` 数值解析及 `%` 标签）、动态 1–2 行防溢出居中图注（`format_caption` 消除重复前缀，图片图注紧贴实图底部）、自适应列宽表格（含 `highlight_row` 与 `**单元格**` 原生加粗）与多视觉零重叠 `visual_slots` 自动切分、深浅画布 WCAG 对比度自适应。<br>**edit 模式**：原地保留格式编辑，支持递归组合图形 `GROUP`、图表标题/轴标题与智能标点/空白归一的 `replace_text`（变长自动防溢出、`slide: "all"` 全稿统一术语、零命中自动记录 `zero_hit_edits` 杜绝静默失效）、跨维度原槽位无缝替换 `set_table` / 类型切换感知且自动继承原图标题与轴标题的 `set_chart` / 旧图注自动清理 `set_image`（杜绝新旧堆叠）、`Title Only` 页安全回退 `set_body`/`set_two_col`（保护原标题不覆写并支持稀疏 20pt 无圆点主题色栏标题）、`set_source` 底栏去重、`replace_slide`/`set_title`/`set_notes`/`insert_at`/`append`/`delete`/`move`，并基于对象身份输出精准 `slide_diff.json`/`slide_diff.md`（保留 ≤30 行完整表格快照）。 |
| `scripts/check_deck.py` | **13–14 项静态合规与事实核验**：页数/区间上下限、越界（忽略无内容全出血装饰条）、文字溢出、形状/表/图重叠、表格溢出（按各行实际字号估算）、图片等比（容差 2%）、标题与备注、字号下限、字体继承、**占位符/残留标记洁净检查**（拦截裸露 `**bold**`/未渲染 LaTeX `10^{-6}`/`\times`/`Lorem ipsum`/`Click to add`/`[TODO]`）、edit 模式 **零命中改写拦截**（`edit_application` 杜绝 `replace_text` 静默失效）、页级证据覆盖、引文逐字回查（兼容 LaTeX 跨行连字符/复合词/PDF 连字/引用括号，附 `closest real line`）、**高精度双向数字保真回查**（跨保留空白原始语料 + JSON 键值 + 补充文件 + 任务书 + 起始稿核验，支持无歧义上标还原 `10⁻⁶→10^-6`/`10⁶→10^6`、双向科学计数法、表头单位缩放、APA 小数、英中量级后缀 `k/M/B/万/亿` 精确匹配且屏蔽 `"2.5 m"` 米单位误污染，精准拦截指数编造与幻觉）；edit 模式 `bounds`/`overlap` 实查已修改页，仅对未改页豁免既有历史样式告警。 |
| `scripts/export_preview.py` | LibreOffice 隔离 `.lohome` 无头导出真 PDF（**110s** 超时充分保障冷启动 CJK/图表模板完整渲染，无 `soffice` 或超时自动降级 `reportlab` 真 PDF）；PyMuPDF 核验页数/字体/标题可抽取性并渲染 100 DPI 逐页 PNG 与 `contact-sheet.png`；内置纯 Python Pillow 像素空白检测（`visual_ok`）。 |
| `scripts/assemble_result.py` | 动态满足 `case.json` 全部 `output_contract` 字段及常见别名，诚实判定 `status`（首轮 10 项关键检查全过才标 `ok`；Pass ≥2 仍 100% 严格门禁 8 项硬事实/几何检查，仅残余纯启发式 `OVERFLOW`/`TABLE_OVERFLOW` 在 PDF+`visual_ok`+`bounds`+`overlap` 全过时诚实记入 `exceptions` 防死循环），产出实算 `rubric_coverage` 与实测 `efficiency_metrics`（耗时、轮次、通过率、证据/讲稿覆盖率）、`result.json`、`plan.json`、`qa_report.md`（含六要素覆盖、实测效率与 GFM `|` 转义）、`outline.md`（含中英文语速时长估算）、`change_log.md`、`slide_diff.md`。 |
| `scripts/deckutil.py` / `scripts/selftest.py` | 共享几何/样式/WCAG 对比度/文本度量库与一键双模式合成自测工具。 |

## 2. 命令行参数与最小调用示例（Example / 示意）

| 命令 | 关键参数 | 说明 |
|---|---|---|
| `run_pipeline.py prepare` | `--case-dir /input` `--work /app/work` `--out /app/output` | 清点输入、生成摘要与保底 `partial` `result.json` |
| `run_pipeline.py deliver` | `--case-dir /input` `--work /app/work` `--out /app/output` `--plan /app/work/slides_plan.json` `[--dpi 100]` `[--patch-slide 'N:{...}']` `[--patch-file ...]` `[--template ...]` `[--start ...]` | 一键构建、13–14 项自检与 2 轮自动修补、110s 导出 PDF/PNG、更新终态 `result.json` 与版本历史；支持单页局部热修无需重写整份计划 |
| `selftest.py` | （无参数） | 在临时目录动态合成 create + edit 双模式算例全流程自测，输出 `SELFTEST_OK` |

```bash
# 示例（Example）：标准两步极速调用流程
RP=/skills/weekly-skill/scripts/run_pipeline.py
python3 "$RP" prepare --case-dir /input --work /app/work --out /app/output
# -> 基于 stdout 输出的 TASK.MD / MATERIAL DIGEST / START DECK DUMP 编写 /app/work/slides_plan.json
python3 "$RP" deliver --case-dir /input --work /app/work --out /app/output \
  --plan /app/work/slides_plan.json --dpi 100
```

示意输出目录（Example output layout under `/app/output/`）：
```
presentation.pptx   # 原生可编辑演示稿（模板继承、零母版篡改）
presentation.pdf    # 真实渲染 PDF 导出件（LibreOffice 首选 / reportlab 兜底）
previews/           # slide-01.png .. slide-NN.png (100 DPI) + contact-sheet.png
result.json         # 含 case_id/status/editable_source/export_artifact/content_map/compliance_checks/exceptions/version_history
plan.json           # 页序规划与 rationale
qa_report.md        # 13 项合规检查 + 异常处置 + 页级证据映射一页审计报告
outline.md          # 逐页大纲、要点与演讲备注
change_log.md       # 全流程时间戳与修改日志
version_history.md  # 多轮迭代版本快照与一键回退命令
slide_diff.md       # （edit 模式）修改前后逐页变更对照表
```

## 3. 依赖清单与降级策略

- **Python ≥ 3.10**：`python-pptx`（必需）、`pillow`（等比插图、拼图与像素空白检测）、`pymupdf`（PDF 文本/位图/矢量图/表格抽取与导出校验；缺失时自动降级 `pdfplumber` → `pdftotext`）。
- **文档导出**：`libreoffice`（`soffice --headless --convert-to pdf`，首选，**110s 超时**）；缺失或超时自动降级 `reportlab` 真实 PDF 导出与 `pdftoppm`/PIL 预览兜底。

## 4. 目录结构与文档导航

```
SKILL.md                        精炼执行卡（<=8.0 KB，两步极速闭环、硬停机护栏、无脚本兜底与交付红线）
README.md                       架构说明、模块参数表、最小调用示例、依赖与适用边界
GITHUB_URL.txt                  公开仓库链接（根路径）
references/plan-schema.md       slides_plan.json 完整 schema、3 类论文页序骨架与全部 edits 操作说明
references/template-rules.md    模板继承、深浅画布对比度自适应、多视觉分槽与防溢出规范
references/failure-playbook.md  15 类异常码、环境降级与诚实 status 判定手册
scripts/                        8 个自包含可移植 Python 模块（含 run_pipeline.py 与 selftest.py）
```

## 5. 适用边界与已知限制

- **适用场景**：给定论文/研究材料与真实模板/起始稿、具备明确页数/受众/语言/证据映射约束的学术汇报演示稿生成与局部编辑迭代。
- **不适用场景**：开放式文献综述检索、科研数据可视化算法研发、论文全文润色、无文本层扫描 PDF 的复杂公式 OCR（遇扫描版 PDF 时如实标注 `SCANNED_PDF`，仅依据结构化材料字段并在 `uncertainty` 记录缺口，绝不凭图猜数）。
