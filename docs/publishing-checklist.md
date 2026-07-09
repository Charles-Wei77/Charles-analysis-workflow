# Publishing Checklist

Use this before pushing the workflows to GitHub.

## Repository Setup

From this directory:

```bash
git init
git add .gitignore README.md LICENSE requirements.txt .env.example docs \
  local-analysis-workflow local-research-workflow
git status --short
```

Check that generated outputs and local config are ignored:

```bash
git status --ignored --short
```

Do not add:

- `.env` or `.env.*` with real values
- `local-research-workflow/config/project_paths.local.json`
- `local-analysis-workflow/runs/`
- `local-research-workflow/runs/`
- downloaded private PDFs, screenshots, exports, or provider SDK caches

## Secret Scan

Run a plain text scan before the first commit:

```bash
rg -n "\\bsk-[A-Za-z0-9_-]+|api[_-]?key|token|secret|password|passwd|Authorization|Bearer|Cookie|IFIND|WUDAO|/Users/" \
  --glob '!**/runs/**' \
  --glob '!**/__pycache__/**'
```

Expected results should be placeholders, docs, or environment-variable names only.

## License

The repository uses a bilingual MIT license file. The English MIT License text controls; the Chinese text is a reference translation.

## First Commit

```bash
git diff --cached --stat
git commit -m "Prepare workflows for public release"
```

Then create the GitHub repository and push:

```bash
git remote add origin git@github.com:<user>/Charles-analysis-workflow.git
git branch -M main
git push -u origin main
```

## After Publishing

- Add a short GitHub description that says users must bring their own data sources.
- Keep issues focused on workflow bugs, data-source adapters, and documentation.
- Do not accept pull requests that include real credentials, licensed data dumps, or generated private runs.
