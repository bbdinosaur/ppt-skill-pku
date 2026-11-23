# academic-talk-structurer — 学术汇报内容组织 Skill (`v29`)

把一篇论文、多文件科研材料（PDF/DOCX/MD/CSV/TSV/JSON/XLSX）或已有草稿 PPT 组织成可直接上台的学术汇报：受众与时长感知的叙事蓝图与页级规划、每页唯一主张绑定源材料证据定位、严格模板兼容的可编辑 PPTX（自动继承模板字体、安全框、Logo 避让与深浅底板，内置 booktabs 学术三线表、单/多系列 `Ours` 高亮对比图、KPI 指标卡、流程步骤卡、双栏对比卡、深底透明图衬垫、旧稿原图保留/提取与裁剪重排差分）、含页级证据与章节覆盖总表的口语讲稿讲义、4 维问答预演、渲染预览自检与契约对齐的 `result.json`。

## 快速使用（四步确定性闭环，详见 `SKILL.md`）
```bash
# 默认挂载 /skills/weekly-skill（若在其他路径可先设 SKILL_DIR）
# Step 1 一键盘点与抽取（inspect_case + extract_figs + plan_pages + qa_draft -> reading_map.md / paper.txt / starting_deck_dump.json）
python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/ingest.py --case /input/case.json --work /app/work

# Step 2 读 /app/work/reading_map.md 与 /app/work/paper.txt 关键章节原文，参照 references/deck-spec.template.json 编写 /app/work/deck_spec.json

# Step 3 一键构建 PPTX、四合一分级门、PDF/PNG + 2x3 接触表渲染（100 DPI）、讲义与证据总表导出、result.json 装配
python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/deliver.py --case /input/case.json --work /app/work --out-dir /app/output

# Step 4 目检 /app/work/contact_sheet_*.png，必要时微调 deck_spec.json 后重跑 deliver.py
```

### 独立可复用脚本与资源资产
- **盘点与多源抽取**：`scripts/inspect_case.py`（契约/完整 `task.md` 约束/模板/起始 deck 盘点）、`scripts/extract_figs.py`（单图原生高清提取、混合栅格+矢量复合图与上/下标矢量裁剪、IEEE/Nature 分节识别、含表题页快速抽取 `paper_tables.json`、全文与大纲）、`scripts/ingest.py`（一键串联并自动合并 `.csv/.tsv/.md/.txt/.docx/.json/.xlsx/.tex` 补充文件、起始 deck 组合形状/原生图表数据与嵌入原图 `deck_s*_p*.png`，自动检测 `case.json` 标签与源文冲突，输出含完整 `task.md`、精确行号、长文定量/局限锚点与受众策略的 `reading_map.md`）、`scripts/qa_draft.py`（追问钩子扫描）、`scripts/plan_pages.py`（受众加权分类、支持 `task.md` 约束兜底、受众叙事策略 `audience_strategy` 与页型配比校核）。
- **构建、审计与导出**：`scripts/build_deck.py`（可编辑 PPTX 构建、模板字体与 Logo/横幅避让、booktabs 学术表、科学计数法安全 `Ours` 高亮图表/卡片、自适应零溢出 KPI/步骤/双栏与图注、数学上下标排版、旧稿右/左/底保留原图自动避让与裁剪重排）、`scripts/render_preview.sh`（LibreOffice 双 profile 容错导出 PDF + `pdftoppm -r 100` + `pdffonts` 嵌入检查）、`scripts/audit_talk.py`（确定性渲染像素+语义质量审计，兼顾视觉与纯文本 Agent 的空白页/截断/文字墙/漏用图表/短讲稿/问答完整度检查并输出 Top-3 精修建议）、`scripts/export_notes.py`（含页级证据与节奏总表、源图表与章节覆盖、取舍复核与 `starting_deck_diff` 的讲义导出）、`scripts/build_result.py`（`result.json` 语义同义词对齐装配、标签冲突/默认模板自动诚实落盘、真实多轮 `changes.json` 修订轨迹与 `quality_audit` 诊断摘要）。
- **方法论与可复用规范**（`references/`）：`narrative-patterns.md`（受众×时长配比与 4 档结论强度阶梯）、`audience-blueprints.md`（四类受众可复用汇报蓝图与上台前 10 项自审清单）、`worked-example.md`（虚构论文完整端到端组织与取舍示意案例）、`deck-spec.template.json`（完整 JSON 骨架）、`qa-playbook.md`（4 维追问预演手册）、`evidence-mapping.md`（定位符与反编造规则）、`template-and-export.md`（母版/字体/Logo/深浅底与导出排错）、`checks.md`（四合一门分级策略、像素+语义审计与 schema）。

## 依赖、安装与无脚本/缺包降级
| 依赖 | 安装方式 | 用途 | 缺失时自动降级 |
|---|---|---|---|
| Python 3.8+ | 系统自带 | 运行脚本与标准库编排 | 必须 |
| `python-pptx` | `pip install python-pptx` | 可编辑 PPTX 构建与母版/旧稿适配 | 报错提示安装 |
| `pymupdf` (`fitz`) | `pip install pymupdf` | PDF 文本/大纲/单图原生与矢量裁剪 | 退回 `pdftotext` + `pdftoppm` |
| `pdfplumber` | `pip install pdfplumber` | 含表题页表格与 LaTeX booktabs 抽取 | 退回 `fitz.find_tables` / 原文阅读 |
| `Pillow` | `pip install Pillow` | 透明图检测、2×3 接触表拼接 | 跳过接触表，保留单页预览 PNG |
| `python-docx` / `openpyxl` / `reportlab` | 可选 `pip install` | `.docx`/`.xlsx` 文本表抽取 / 讲义 PDF 导出 | `.docx`/`.xlsx` 自动降级标准库 `zipfile+XML` 解析；讲义保留 Markdown |
| `libreoffice` + `poppler-utils` | `apt install libreoffice poppler-utils` | PPTX→PDF→100 DPI PNG 与字体检查 | 保留可编辑 PPTX+lint，`rendering` 标 `failed` |

所有非标准库导入均受 `try/except` 保护（AST 零未保护导入），脚本路径经 `Path(__file__).resolve().parent` 动态定位，无硬编码个人路径。

## 环境自检
```bash
python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/selftest.py
```
端到端验证 11 个核心组件（python-pptx / PIL / 原生卡片+图+表 / `edit_deck` 裁剪重排 / 双语 lint / 矢量 PDF 抽取 / `ingest+deliver+chart+notes_matrix` / 科学计数法+单位换算+Q&A 数值核验 / `audit_talk` 像素与语义审计 / `build_result` / `soffice`），11/11 PASS 即可稳定运行。

## 最小调用示例（Schematic Example，非真实评测案例）
- 示例输入 `/input/case.json`：
  ```json
  {"case_id":"demo-defense-01","audience":"答辩委员会","duration_minutes":10,"slide_count":10,"template":"/input/tpl.pptx","source_paper":"/input/paper.pdf","output_contract":{"deck_path":"/app/output/talk.pptx","required_fields":["case_id","status","slide_plan","evidence_map","qa_prep","rendering"]}}
  ```
- 示例输出 `/app/output/result.json`（片段示意）：
  ```json
  {"case_id":"demo-defense-01","status":"complete","tradeoffs":["压缩相关工作至1页，保留主结果与消融"],"evidence_coverage":{"content_slides":8,"slides_with_evidence":8,"coverage_pct":100.0,"visual_diversity_pct":75.0},"slide_plan":[{"pptx_slide":5,"role":"result","claim":"8192 batch top-1 达 76.3%（+3.2pt）","time_s":75}],"evidence_map":[{"pptx_slide":5,"evidence":[{"type":"figure","locator":"Fig.2, p.4","strength":"S1"}]}],"qa_prep":[{"category":"data","question":"噪声是否伤收敛？","answer_key":"见 Fig.2 曲线与 §3.3 方差","cite_slide":5,"backup_slide":10}]}
  ```

## 适用边界、异常处置与维护
- **边界**：聚焦“已有研究材料/草稿 PPT → 可直接上台的学术汇报与可复核证据链”；不负责论文精读批注、学术论文写作润色或科研绘图素材生产。实际输入文件与原论文为唯一事实依据，不依据错误领域标签补造数据。
- **异常**：标签冲突、缺失模板、纯起始 deck（无 PDF）、多文件补充数据、扫描 PDF、公式/缺字、导出超时等处置见 `SKILL.md` 与 `references/checks.md`。
- **版本与洁净声明**：当前版本见 `VERSION.txt`（`v29`）；包内仅含纯 UTF-8 文本脚本与方法论文档，无预制答案、无测试集或虚构跑分。
