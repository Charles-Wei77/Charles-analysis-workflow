# Local Research Workflow

`local-research-workflow` turns a company or industry input into a source-retaining research report, editable Office files, and a banking-style financial model.

## Quick Run

```bash
python research_workflow.py \
  --input examples/company_input.sample.json \
  --export docx pptx pdf \
  --model-depth banking
```

Company-only mode:

```bash
python research_workflow.py \
  --company 宁德时代 \
  --industry 动力电池与储能 \
  --topic 全球动力电池竞争格局、储能增长与估值重估 \
  --template securities \
  --export all \
  --model-depth banking
```

Outputs are written to:

```text
runs/<timestamp>-<topic>/
```

## Outputs

- `deep-research-prompt.md`: prompt for Codex, Claude, GPT Researcher, Open Deep Research, browser-use, or another research agent.
- `research-pack.json`: input, research questions, evidence snippets, and local project path metadata.
- `report.md`: Markdown report with source table.
- `financial_model.xlsx`: 3-statement, WACC, DCF, comps, sensitivity, charts, and audit workbook.
- `valuation_summary.json`: Bear/Base/Bull valuation summary.
- `model_audit.md`: model quality summary.
- `report.docx`: editable Word export.
- `report.pptx`: editable PowerPoint export.
- `report.pdf` or `report.html`: PDF if a PDF path is available, otherwise HTML fallback.

## Templates

`--template securities` creates a sell-side style report:

- investment conclusion
- key summary
- company and business overview
- industry and value chain
- competitive landscape
- financial and valuation framework
- catalysts
- risk factors
- source table

`--template consulting` creates a consulting-style storyline:

- executive takeaway
- market context
- company positioning
- competitive benchmark
- strategic options
- financial / valuation lens
- action plan
- appendix sources

## Financial Model Policy

The default `--model-depth banking` workbook uses formulas for forecast areas and stores hardcoded inputs with comments. If historical data, segments, market fields, or comps are missing, placeholders are used and flagged in the audit output.

Official disclosures should be the primary evidence layer. Structured fields can come from Wudao MCP or a private licensed provider configured by the user. Credentials must stay in `.env`, shell environment variables, or a local adapter outside the committed repository.

## Data Source Environment

Copy the repository-level `.env.example` or this workflow's `.env.example`, then fill only the providers you use:

```bash
cp ../.env.example ../.env
```

Common optional variables:

```bash
export WUDAO_API_KEY="..."
export IFIND_USERNAME="..."
export IFIND_PASSWORD="..."
export OPENAI_API_KEY="..."
export TAVILY_API_KEY="..."
export FMP_API_KEY="..."
```

See `../docs/data-sources.md` for the public configuration boundary.

## Validate

```bash
python validate_workflow.py runs/<run-dir> --require-exports
```

For a lightweight local test, omit `--require-exports` or run the workflow with `--export none`.

## Optional Local Paths

If you have local copies of related research projects or skills, place paths in:

```text
config/project_paths.local.json
```

This file is ignored by Git. Use `config/project_paths.example.json` as the public template.
