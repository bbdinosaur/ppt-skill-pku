# Paper to Academic Slides (`paper-to-academic-slides`)

一个面向科研汇报场景的 Agent Skill：把学术论文 PDF 或研究材料转换为**主线清晰、页页可溯、模板合规、讲稿分离、可编辑且经过渲染核验**的学术演示文稿（PPTX + PDF + PNG 预览 + 页级证据映射）。

## 1. 目录结构与可复用资源

| 路径 | 用途 | 复用方式 |
|---|---|---|
| `SKILL.md` | Agent 精简执行卡（两命令闭环、硬约束、S4.5 强制门控与降级策略） | 直接放入 Agent Skills 目录触发加载 |
| `scripts/run_pipeline.py` | 一键两阶段流水线封装：`prep`（S0+S1+低文本自动 OCR+标签/正文一致性预警+立即落盘保底 `result.json`+生成 ≤45KB 零截断 `reading_pack.md`）与 `build`（S3+S4+S4.5+S5+首轮高优缺陷精修门控 `polish_checklist`），正常 4–7 步完赛并兼容纯文本模型 | 推荐默认入口，支持 `prep` / `build` 子命令 |
| `scripts/resolve_case.py` | 解析 `case.json`/`task.md`（含防静默截断标记、嵌套 `output_contract` 与 `case_labels` 提取）、智能区分双 PPTX 有效模式（`effective_template_mode`）、导出起始 deck 逐页检查 `starting_deck_dump.md` 与原图资产 `starting_deck_figs/*.png` | 独立 CLI，支持 `--case-json` / `--input-dir` / `--out` |
| `scripts/extract_paper.py` | 双栏阅读顺序保真清洗 PDF/文本/DOCX（含多行标题拼接、加密 PDF 锁检测、扫描页自动 Tesseract OCR、图片公式页检测、未抠出图号列表、多行 `\multicolumn` 表头合并与无框表回退、优先保留局限/结论页的推荐读页），生成图/表置顶的 `paper_index.md` 和 `pages/pNN.txt`，支持双栏题注匹配、多面板复合图拼合、单 XObject 矢量图裁剪与自动白边裁切 | 独立 CLI，支持 `--out` / `--dpi` |
| `scripts/build_deck.py` | 按模板母版/布局重建或 `--reuse`（支持 `source_slide` 选页重排与左右侧 `keep_existing_media` 自动避让）构建零幽灵占位符可编辑 PPTX；支持只读 XML 保母版背景、母版左侧边栏自动检测与避让、`fc-list` 字体发现与缺字 1.08x 宽补偿、无阴影现代 flat 视觉、自动导语加粗、三线学术表、零换字且自动突出 Ours 的原生可编辑图表 (`chart`)、图文表双媒自动分栏防丢弃、中英自适应溯源页脚、KPI 卡、底部 `takeaway` 与双栏对比 | 独立 CLI，支持 `--template` / `--reuse` / `--out` |
| `scripts/check_layout.py` | 静态版面几何与排版检查（越界、字号下限、自动豁免全幅背景图并检测图文/图图/图注重叠及母版底部页脚碰撞升 `fail`、段落密度、模板占位残留词） | 独立 CLI，输出 `layout_report.json` |
| `scripts/verify_evidence.py` | 幻灯片与论文及补充材料交叉校验（页数上限、非封面必有证据与 PDF 容错原文 `quote`、支持补充材料 `.csv`/`.md`/`.txt`/`.json` 溯源、图存在/防重图/图号匹配/嵌图漏引拦截、图注与示意图声明、全量数字与**表头/正文/图表全量级同量纲科学单位篡改检测**、页脚 `slide_line` 幻影引用拦截、封面防误导标签、基线归属、讲稿查重与控时、栏目词、局限页） | 独立 CLI，输出 `evidence_report.json` |
| `scripts/finalize_output.py` | 隔离临时 profile 导出 PDF 与 100 DPI（上限 ≤ 108 DPI）预览 PNG、自动校验渲染页数/标题文本/像素空白页/U+FFFD 豆腐块/PDF 实际字体替换审计、生成 `--reuse` 逐页 diff、输出 `evidence_map.json`/`evidence_map.md`/`speaker_notes.md`/`slides_outline.json`/`result.json`（含六要素审计、`repro_command` 与输入/输出 SHA-256 哈希）并镜像 `output_contract` 目标文件 | 独立 CLI，支持 `--case-resolved` 对齐 `output_contract` |
| `references/slides-schema.md` | `slides.json` 完整数据契约（含 `table`/`images`/`evidence` 字段） | 供 Agent 或人工编写大纲 JSON 复用 |
| `references/slides_scaffold.json` | 空白页级大纲脚手架（示例模板） | 复制后按目标论文填写 |
| `references/quality-and-evidence.md` | 主线提取、信息密度、讲稿写法、模板安全与可核验性细则 | 团队汇报规范参考 |
| `references/failure-playbook.md` | 加密 PDF、扫描件、矢量图、占位符缺失、字体缺失等异常处置手册 | 运维与降级处置指南 |

## 2. 环境依赖与安装

- **Python**: `>= 3.8`
- **必选 Python 包**: `python-pptx`, `pymupdf` (`fitz`), `Pillow`
- **可选系统工具**: `libreoffice` (`soffice --headless --convert-to pdf`) 用于 PDF 导出与逐页 PNG 预览渲染；`tesseract` 用于扫描 PDF 自动 OCR 兜底；`pdftotext` 作为 PDF 文本抽取兜底；`fc-list` 用于本机字体审计。
- **安装示例**:
  ```bash
  pip install python-pptx pymupdf pillow
  # Linux 安装 LibreOffice（用于导出 PDF 渲染核验）
  sudo apt-get install -y libreoffice
  ```
- 所有脚本均使用标准库与上述常见科研包，无硬编码绝对路径、无符号链接、无二进制依赖文件。

## 3. 快速调用步骤（完整闭环）

推荐一键两阶段入口（正常 4–6 步闭环，自动传递模板模式、落盘保底 JSON、生成一站式简报与校验门控）：
1. `python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/run_pipeline.py prep --case-json /input/case.json --input-dir /input --work /app/work --out-dir /app/output`（同步落盘保底 `result.json` 与 `/app/work/paper/reading_pack.md`）
2. 一次阅读 `/app/work/paper/reading_pack.md`（仅需补读其余页时查看 `pages/pNN.txt`），编写 `/app/work/slides.json`（参考 `references/slides-schema.md`）
3. `python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/run_pipeline.py build /app/work/slides.json --work /app/work --out-dir /app/output`（若 `layout_fail > 0` 或 `evidence_errors > 0`，修正后重跑本步）
4. 核对自动化渲染指标（`render_text_ok`、`visual_pixel_ok`）；多模态环境可选抽检 `inspect_previews` PNG，纯文本模型无需读取二进制 PNG。

所有底层脚本也可分步独立调用（`resolve_case.py` → `extract_paper.py` → `build_deck.py` → `check_layout.py` → `verify_evidence.py` → `finalize_output.py`）。

## 4. 最小调用示例（Example，非真实评测输出）

- **示例输入** `/input/case.json`:
  ```json
  {"case_id": "example-01", "paper": "paper.pdf", "template": "dept_template.pptx", "max_slides": 10}
  ```
- **示例产出**（`/app/output/`）:
  - `slides.pptx`：可逐页编辑的学术演示文稿（含真实 speaker notes 与页脚证据行）
  - `slides.pdf` & `previews/slide-01.png ...`：导出渲染件与 100 DPI 视觉核验图（含 PDF 字体替换与空白/豆腐块自动质检）
  - `evidence_map.json`、`evidence_map.md`、`speaker_notes.md` & `slides_outline.json`：独立机读/人读页级证据映射表、逐页讲稿与可复跑大纲 JSON
  - `result.json`：包含 `case_id`、`status`、`deliverables`、`evidence_map`、`six_element_audit`、`render_record`、`checks`、`gaps_and_uncertainty`、`change_records`

## 5. 适用范围、限制与维护

- **适用**：英文/中文学术论文 PDF（含位图或矢量图表、表格）以及 Markdown/纯文本研究材料；支持无模板 16:9、母版模板重建、起始 deck `--reuse` 原地适配。
- **限制**：纯扫描图像 PDF 在系统安装 `tesseract` 时可自动触发 200 DPI OCR 兜底（若未安装则需外部 OCR 并如实记入 `gaps_and_uncertainty`）；极复杂跨页合并单元格表格可能需要人工微调；本 Skill 不负责科研绘图美化与论文内容润色。
- **维护**：各脚本均支持 `--help` 查看参数；更新规则时需保持 `SKILL.md`、`references/slides-schema.md` 与脚本字段同步。
