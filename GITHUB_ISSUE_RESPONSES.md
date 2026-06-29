# GitHub Issue Response Drafts

## Issue: Run-time error 9: Subscript out of range

Suggested response:

```text
Good catch. This is a real bootstrap/setup issue. The current `MedParser.bas` assumes it is being imported into the existing no-PHI `MedicationDispensing.xlsm` workbook template, which already has the required sheets:

- Patient & Input
- Medications
- Label Preview
- Log
- Label Previews

If imported into a blank workbook, `SetupWorkbook` can fail on `ThisWorkbook.Sheets(SH_INPUT)` with run-time error 9.

For now, I added setup documentation making clear that developers should start from the clean no-PHI workbook template, not a blank workbook. I also added troubleshooting notes for this exact error. A future code enhancement should add an `EnsureWorksheet()` bootstrap step so the macro can create missing sheets automatically.
```

## Issue: Setup instructions not found

Suggested response:

```text
Agreed. Added `SETUP_INSTRUCTIONS.md` with de novo setup steps: Windows/Excel requirements, Brother QL-1100C driver/default media setup, cloning the repo, using the clean no-PHI workbook template, importing `MedParser.bas`, compiling, installing sheet event handlers, running `SetupWorkbook`, testing with no-PHI sample data, and printing a physical DK-1202 test label.
```

## Issue: Need stable branch and working branch

Suggested response:

```text
Agreed. Added branch/release guidance in `docs/BRANCHING_AND_RELEASES.md`. Proposed model:

- `main` = stable, tested, clinic-safe source/docs
- `dev` = active integration branch
- `feature/*` = targeted work branches

Because `.xlsm` workbooks are binary and may contain PHI, the repo should remain code/docs/no-PHI samples only. Live clinic workbooks should not be committed.
```

## Issue: Should workbook template be committed?

Suggested response:

```text
For now, no. Based on the project handoff, the repository should be code-only and exclude `.xlsm/.xlsx/.xls/.csv` files because workbooks can easily contain PHI. The clean no-PHI workbook template should remain in approved clinic storage until the team explicitly decides whether a sanitized template can safely be versioned.
```
