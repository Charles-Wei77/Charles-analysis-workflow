# Data Sources

This repository does not ship private data access. Every user must configure their own market-data and research providers.

## Default Evidence Layer

Use public and official sources as the baseline:

- company annual reports, interim reports, and announcements
- exchange announcements and interaction records
- SEC / IR / earnings-call materials for overseas companies
- user-provided PDFs, links, and structured JSON fields

The workflow should mark weak or unverified claims as lower evidence levels instead of promoting them to facts.

## Wudao MCP

Wudao is the preferred public setup path for A-share market and theme data when users do not have your private iFinD permission.

Example MCP-style configuration:

```json
{
  "mcpServers": {
    "wudao-stock": {
      "url": "https://stock.quicktiny.cn/api/mcp",
      "headers": {
        "Authorization": "Bearer ${WUDAO_API_KEY}"
      }
    }
  }
}
```

Recommended environment variables:

```bash
export WUDAO_API_KEY="your_wudao_key"
export WUDAO_MCP_URL="https://stock.quicktiny.cn/api/mcp"
```

Client configuration formats vary. Keep the endpoint and authorization secret local to the user's machine.

## iFinD / Tonghuashun / Eastmoney

iFinD access is private and account-bound. This repository should not include the maintainer's username, password, refresh token, cookies, local SDK files, or downloaded private data.

Users who have their own licensed access can configure a local adapter and pass fields into the workflow input JSON:

```bash
export IFIND_USERNAME="your_username"
export IFIND_PASSWORD="your_password"
export IFIND_REFRESH_TOKEN="optional_refresh_token"
```

The workflow does not require those variables to run. If they are absent, use official filings, Wudao, public web sources, or manual JSON inputs.

## Input JSON Is the Stable Contract

Both workflows accept structured JSON. Provider adapters should fill these fields rather than hard-coding provider calls into the report generator:

```json
{
  "company": "Example Co",
  "stock_code": "000000.SZ",
  "industry": "Example industry",
  "topic": "Research topic",
  "sources": [],
  "financial_model": {
    "historical": [],
    "segments": [],
    "market": {},
    "comps": [],
    "scenarios": {}
  }
}
```

This keeps the repository reusable even when a user's data providers differ.

## Source Quality Rules

- Do not commit keys, cookies, tokens, account IDs, or private provider outputs.
- Mark fallback, stale, or manually entered data explicitly in `source`, `note`, or evidence-level fields.
- Do not let web search or marketing pages override official disclosures.
- Keep generated `runs/` local unless you have manually reviewed them for personal paths and private data.
