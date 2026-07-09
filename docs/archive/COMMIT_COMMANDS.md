# Commit Commands

These commands assume you already cloned the repo locally and copied this update into the repo root.

## Recommended GitHub Desktop flow

1. Create/switch to `dev`.
2. Create a new branch from `dev`, for example:

```text
feature/setup-docs-and-repo-hardening
```

3. Copy these files into the repo.
4. Review changes in GitHub Desktop.
5. Commit:

```text
Add setup docs and repo hygiene guidance
```

6. Push the branch.
7. Open a pull request into `dev`.

## Command-line flow

```bash
git status
git checkout dev || git checkout -b dev
git pull --ff-only origin dev || true
git checkout -b feature/setup-docs-and-repo-hardening

git add README.md SETUP_INSTRUCTIONS.md CHANGELOG.md HANDOFF.md GITHUB_ISSUE_RESPONSES.md COMMIT_COMMANDS.md .gitignore .gitattributes docs/ test-data/ tools/ src/MedParser.bas

git status
git diff --cached --stat
git diff --cached -- . ':!src/MedParser.bas'

git commit -m "Add setup docs and repo hygiene guidance"
git push -u origin feature/setup-docs-and-repo-hardening
```

## Before committing

Run:

```bash
git status --short
```

Confirm no files like these are staged:

```text
*.xlsm
*.xlsx
*.xls
*.csv
~$*
patient data
screenshots with PHI
```

## Optional line-ending check for MedParser.bas

PowerShell:

```powershell
$bytes = [System.IO.File]::ReadAllBytes("src\MedParser.bas")
$txt = [System.Text.Encoding]::ASCII.GetString($bytes)
"CRLF count: " + ([regex]::Matches($txt, "`r`n")).Count
"LF count: " + ([regex]::Matches($txt, "`n")).Count
```

The CRLF count and LF count should be equal for a CRLF-only file.
