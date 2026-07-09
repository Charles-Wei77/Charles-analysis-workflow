# Local Analysis Workflow

`local-analysis-workflow` is the upstream analysis layer for the report workflow. It focuses on evidence quality, company and industry structure, competitor mapping, substitution risk, A-share theme factors, catalysts, and tracking indicators. It does not create styled report exports directly.

## Quick Run

```bash
python analysis_workflow.py --company 埃斯顿
```

With a structured input file:

```bash
python analysis_workflow.py --input examples/company_input.sample.json
```

Outputs are written to:

```text
runs/<timestamp>-<company>-analysis/
```

## Core Outputs

- `analysis.md`: main analysis memo.
- `evidence_ledger.json`: A/B/C/D evidence grading for key claims.
- `competitor_map.md`: direct competitors, indirect substitutes, overseas peers, customer self-build risk, and upstream/downstream integration.
- `industry_structure.md`: concentration, supply-demand, pricing pressure, technology routes, and profit-pool migration.
- `substitution_risk.md`: substitution source, timing, intensity, impact, and defense.
- `factor_score.xlsx`: A-share theme-factor scoring workbook.
- `catalyst_calendar.md`: event and catalyst calendar.
- `tracking_plan.md`: measurable follow-up indicators.
- `export_input.json`: downstream input for `local-research-workflow`.
- `handoff_command.txt`: ready-to-run downstream command.

## Evidence Levels

```text
A: official filings, annual/interim reports, exchange interaction, SEC/IR, earnings calls
B: sell-side or industry reports with cross-source support
C: news, market narrative, concept labels, or single non-official source
D: denied, unsupported, stale, or falsified claims
```

## Data Source Policy

The workflow can run with manually provided public filings and structured JSON. Market data, capital-flow data, and exact financial fields should come from user-configured sources:

- Wudao MCP for A-share market and theme data, if the user configures it.
- Private licensed providers such as iFinD, Tonghuashun, or Eastmoney, if the user has their own permission.
- Public filings and exchange disclosures as the default evidence base.

No private credentials are required in the repository. See `../docs/data-sources.md`.

## Validate

```bash
python validate_analysis_workflow.py runs/<run-dir>
```

Validation checks required files, evidence levels, workbook sheets, formulas, and downstream handoff fields.

## Downstream Handoff

After reviewing the analysis package:

```bash
python ../local-research-workflow/research_workflow.py \
  --input runs/<run-dir>/export_input.json \
  --export all \
  --model-depth banking
```
