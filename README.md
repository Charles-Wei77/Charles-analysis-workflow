# Charles Analysis Workflow

Local agent workflows for investment research, A-share company analysis, evidence tracking, financial modeling, and report export.

This repository contains workflow code and templates only. It does not include private market-data credentials. Public users can run the workflows with official filings and manually supplied inputs, then optionally connect their own Wudao MCP, iFinD, Tonghuashun, Eastmoney, FMP, Tavily, or other licensed data sources.

## 中文介绍

`Charles Analysis Workflow` 是一套面向 A 股研究和个人投研自动化的本地 agent 工作流。它的核心目标不是直接给出“买不买”的简单结论，而是把一次公司研究拆成更可追踪、更可复核的流程：先做证据、产业链、竞品、替代风险和跟踪指标，再进入研报、PPT、财务模型等输出环节。

### 它解决什么问题

- **🧭 研究不散**：把公司名、行业、资料来源、核心判断、反方风险和跟踪指标放进同一套结构里。
- **🔍 证据有等级**：把公告、年报、交易所互动、研报、新闻、传闻分层，避免把弱证据写成强结论。
- **🏭 更贴近 A 股语境**：保留主题因子、情绪资金、位置量价、催化剂、替代风险等更适合中国市场的分析维度。
- **📊 输出可交付**：可以生成 Markdown、Word、PPT、PDF/HTML、Excel 财务模型和估值摘要。
- **🔐 数据源自带边界**：仓库不包含作者的 iFinD 权限或任何私有 key。使用者需要自己配置 Wudao MCP、iFinD、同花顺、东方财富或其他数据源。

### 两个核心工作流

| 目录 | 中文理解 | 负责什么 |
|---|---|---|
| `local-analysis-workflow` | 研究版 / 分析版 | 证据台账、竞品分析、行业格局、替代风险、A 股因子、催化剂、跟踪计划，并输出 `export_input.json` |
| `local-research-workflow` | 输出版 / 交付版 | 接收公司输入或 `export_input.json`，生成研报、PPT、PDF/HTML、Word、财务模型和估值摘要 |

### 推荐使用方式

1. **先跑研究版**：把公司和资料整理成结构化分析包，先看逻辑、证据和风险是否站得住。
2. **再跑输出版**：确认研究框架后，再生成研报、PPT、PDF 和财务模型。
3. **最后人工复核**：模型和报告只是工作流产物，不替代投资判断、合规审查或官方披露。

### 适合谁

- 想把 A 股公司研究流程标准化的个人投资者或研究员。
- 想让 Codex、Claude 等 agent 复用同一套投研 SOP 的用户。
- 想把“研究过程”和“输出交付”分开的买方、卖方、咨询或个人知识库用户。

## What's Included

- `local-analysis-workflow`: upstream analysis workflow. It creates an evidence ledger, competitor map, industry-structure view, substitution-risk table, A-share factor workbook, catalyst calendar, tracking plan, and `export_input.json`.
- `local-research-workflow`: downstream report workflow. It turns structured inputs into Markdown, DOCX, PPTX, PDF/HTML, `financial_model.xlsx`, `valuation_summary.json`, and `model_audit.md`.
- `docs/data-sources.md`: how to configure public, Wudao, and private licensed data sources.
- `docs/publishing-checklist.md`: GitHub release and secret-scan checklist.

Generated run outputs are ignored by Git. Keep `runs/`, `.env`, and `project_paths.local.json` local.

## Quick Start

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Run the upstream analysis workflow:

```bash
python local-analysis-workflow/analysis_workflow.py \
  --input local-analysis-workflow/examples/company_input.sample.json
```

Validate the generated analysis run:

```bash
python local-analysis-workflow/validate_analysis_workflow.py \
  local-analysis-workflow/runs/<run-dir>
```

Run the downstream report workflow:

```bash
python local-research-workflow/research_workflow.py \
  --input local-research-workflow/examples/company_input.sample.json \
  --export docx pptx pdf \
  --model-depth banking
```

If you only want to test the model/report core without Office exports:

```bash
python local-research-workflow/research_workflow.py \
  --input local-research-workflow/examples/company_input.sample.json \
  --export none \
  --model-depth banking
```

## Data Source Boundary

The default workflow treats official filings and exchange disclosures as the safest source layer. Structured market and financial fields must come from data sources the user configures locally.

- Wudao: optional MCP provider. See [docs/data-sources.md](docs/data-sources.md).
- iFinD / Tonghuashun / Eastmoney: optional private or licensed providers. Bring your own account and adapter.
- FMP / Tavily / web search: optional supplemental sources. Do not let them override official disclosures without analyst review.

Never commit real keys, cookies, refresh tokens, downloaded private files, or generated runs that contain personal paths.

## Local Configuration

Copy the example environment file and fill only the providers you use:

```bash
cp .env.example .env
```

Optional project path overrides can live in:

```text
local-research-workflow/config/project_paths.local.json
```

That file is ignored by Git. The committed example is:

```text
local-research-workflow/config/project_paths.example.json
```

## Workflow Handoff

`local-analysis-workflow` writes an `export_input.json` and a `handoff_command.txt`. After reviewing the analysis package, pass `export_input.json` into `local-research-workflow` to produce the final report and model.

## License

Charles Analysis Workflow is licensed under the MIT License. See [LICENSE](LICENSE) for the English license text and Chinese reference translation.
