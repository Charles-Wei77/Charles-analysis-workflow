# Charles Analysis Workflow

Local agent workflows for investment research, A-share company analysis, evidence tracking, financial modeling, and report export.

This repository contains workflow code and templates only. It does not include private market-data credentials. Public users can run the workflows with official filings and manually supplied inputs, then optionally connect their own Wudao MCP, iFinD, Tonghuashun, Eastmoney, FMP, Tavily, or other licensed data sources.

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
