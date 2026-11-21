# academic-slides-pku — 学术 PPT 生成与编辑 Skill

Agentic Science Challenge（北大专场 efficiency-3）赛题 `ppt-6ad93034` 的参赛作品：
把固定论文/研究材料组织成**可编辑、模板继承、证据可追溯**的学术演示稿
（原生 PPTX + PDF 导出 + 页级证据映射 + 合规检查报告）。

## 它解决什么

论文→幻灯片的翻译难点不在复制段落，而在取舍与重排：哪页放什么主张、图先出谁、
数字是否抄对、模板是否被改坏、导出是否塌版。本 Skill 用"脚本干重活、模型做决策"
的架构，在 64 次请求 / 900 秒预算内稳定产出合格稿：

1. `scripts/inspect_case.py` — 清点输入（case.json / material.json / slide-workorder.json /
   模板与起始 deck），抽取 PDF 全文与内嵌图，识别模板版式清单与母版字体。
2. 模型唯一重决策点 — 读材料原文，写 `slides_plan.json`（页级主张、要点、图引用、
   讲稿备注、逐条 claim_sources 引文）。
3. `scripts/build_deck.py` — 按模板真实版式构建：占位符填充（不设字体名，继承主题）、
   图片等比插入（绝不裁剪拉伸）、表格、底部来源行、讲稿备注；edit 模式只动指令页。
4. `scripts/check_deck.py` — 9 项静态合规检查：页数/越界/文字溢出估算/内容重叠/
   图片等比/标题与备注齐备/字号下限/字体继承/引文真实性（quote 必须逐字存在于材料原文，
   防幻觉编造）。
5. `scripts/export_preview.py` — LibreOffice 导出真实 PDF，PyMuPDF 校验页数/字体/标题
   可抽取性，逐页 PNG 预览；无 soffice 时 PIL 几何预览兜底。
6. `scripts/assemble_result.py` — 组装 `result.json`（case_id / status / editable_source /
   export_artifact / content_map / compliance_checks / exceptions + uncertainty /
   change_log / repro）与 plan.json、change_log.md。

## 特性

- **模板铁律**：母版只读；占位符优先；显式零几何占位符自动修复（slide→layout→master
  继承链解析）；字号收缩有下限。
- **数字保真**：页面结论的引文必须逐字回查材料原文；材料缺口写入 uncertainty，不编造。
- **诚实状态机**：先落 partial 结果再迭代；status=ok/partial/failed 由关键检查自动判定。
- **异常手册**：缺图、位图、模板槽位不足、页数超限、导出失败等 13 类异常的处置对照表
  （`references/failure-playbook.md`）。

## 快速使用

```bash
python3 scripts/inspect_case.py  --case-dir /input --work /app/work
python3 scripts/build_deck.py    --plan /app/work/slides_plan.json --work /app/work --out /app/output/presentation.pptx
python3 scripts/check_deck.py    --deck /app/output/presentation.pptx --work /app/work
python3 scripts/export_preview.py --deck /app/output/presentation.pptx --work /app/work
python3 scripts/assemble_result.py --case-dir /input --work /app/work --deck /app/output/presentation.pptx --status auto
```

依赖：`python-pptx`（必须）、`pillow`（图片等比）、`pymupdf`（PDF 抽取/校验）、
`libreoffice`（PDF 导出，可缺失有兜底）。全流程说明见 `SKILL.md`。

## 目录

```
SKILL.md                      执行卡（触发、7 步流程、预算策略、兜底）
references/plan-schema.md     slides_plan.json schema 与证据映射规则
references/template-rules.md  模板继承与版式铁律
references/failure-playbook.md 异常处置对照表
scripts/                      6 个自包含脚本（路径自适应，任意 cwd 可用）
```

## 一键自测

```bash
python3 scripts/selftest.py   # 临时目录生成 mini 算例并全流程跑通，打印 SELFTEST_OK
```

每次交付还会生成 `qa_report.md`（检查表/异常/页级证据映射的一页式审计报告）、
`outline.md`（讲稿大纲）与 `previews/contact-sheet.png`（全 deck 缩略拼图）。

## 适用边界

适用于"固定论文/材料 + 明确模板与页数约束 → 可核验演示稿"的学术汇报场景；
不负责论文精读综述、数据可视化算法与学术文字润色。测试均在本地合成算例上完成。
