# 问答准备（QA Playbook）

## 1. 追问分类与生成规则（每类至少 1 条，共 ≥4 条）
| 类别 | 追问角度 | answer key 结构 |
|---|---|---|
| 机制 | 为什么 work？哪个组件贡献最大？ | 指向 analysis 页 claim + 消融表 locator |
| 数据/规模 | 换数据集/更大规模还成立吗？ | 引用实验设置页范围 + 明确说"源材料未验证 X" |
| 基线/对照 | 基线公平吗？是否调参充分？ | 引用 setup 页对照条件原文；不足则承认 |
| 局限/外推 | 在 X 场景会失效吗？ | 引用 limitation 页对应条目，不辩解 |
| 口径 | 这个数与论文摘要一致吗？ | 复述同一 claim 原句，标 locator |

## 2. answer key 口径一致性
- 与页面 claim/evidence 使用同一数字、同一 locator；禁止临时引入源材料外的新数字。
- 源材料没覆盖的问题，answer key 模板："源材料未直接验证；最接近的证据是 <locator>；合理推测是 <一句>，需实验确认。"（推测必须标注为推测）

## 3. 备用页（qa_backup）规则
- 1 页打包 2–4 个高频追问的要点式回答（每条 ≤2 行，带 locator）；追问多时加页，标题"备用：XX 追问"。
- 备用页不占正文时长；编号放在正文最后，风格与正文一致（同模板布局）。
- slide_plan 中 role=qa_backup；result.json 的 qa_prep[].backup_slide 指向其 pptx_slide 号。

## 3b. 提交前复审三问（逐页自问）
1. 这一页的 claim 能否在 30 秒内用所引 locator 向评审现场佐证？
2. 删除这一页后主线是否仍成立？（是→考虑删/降级为备用页）
3. notes 结尾衔接句是否指向下一页的实际内容？

## 4. deck_spec 中 qa_prep 的落盘
写 `/app/work/qa_plan.json`：
```json
[{"question":"...","answer_key":"...","cite_slide":5,"backup_slide":12,"category":"mechanism|data|baseline|limitation|consistency"}]
```
（此 schema 同时是 build_result.py 读入格式。）
