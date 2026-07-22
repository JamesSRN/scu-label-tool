# Branching and Releases

## Branch model

Use a simple branch model:

```text
main       stable, clinic-safe source/docs only
dev        active integration branch
feature/*  focused work branches from dev
```

Examples:

```text
feature/setup-docs
feature/brother-print-fix
feature/parser-tests
feature/bootstrap-sheets
```

## Rules

1. `main` should always be safe to clone.
2. Work in `dev` or `feature/*` branches.
3. Merge to `main` only after the release checklist passes.
4. Never commit PHI.
5. Do not commit live clinic workbooks.
6. Do not have two people independently edit the same `.xlsm` workbook and expect Git to merge it.

## What belongs in Git

Commit:

```text
src/MedParser.bas
HANDOFF.md
README.md
SETUP_INSTRUCTIONS.md
CHANGELOG.md
docs/*.md
test-data/no-PHI samples
.gitignore
.gitattributes
safe logo/source assets that contain no PHI
```

Do not commit:

```text
*.xlsm, *.xlsx, *.xls
*.csv, *.tsv
live dispense logs
real patient screenshots
OneDrive temp files
backups with patient data
```

## Release process

1. Create a feature branch from `dev`.
2. Update `src/MedParser.bas` and docs.
3. Confirm `.bas` is ASCII + CRLF.
4. Import into a clean no-PHI workbook.
5. Compile the VBA project.
6. Run `SetupWorkbook`.
7. Confirm sheet event handlers are present.
8. Test parser with no-PHI samples.
9. Print a physical Brother DK-1202 label.
10. Update `CHANGELOG.md`.
11. Merge feature branch into `dev`.
12. After integration testing, merge `dev` into `main`.
13. Tag a release if appropriate.

## Suggested tags

```text
v0.1.0  first local parser + label print workflow
v0.2.0  checkbox-driven batch print + label previews
v0.3.0  print log hardening + start-new-patient workflow
```

## Stable workbook vs source repository

The source repository is not the live clinic system. The clinic workbook is built from the source and maintained locally. Keep a clean no-PHI workbook template in the approved clinic storage location, and keep source/docs in GitHub.
