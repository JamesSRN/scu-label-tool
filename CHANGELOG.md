# Changelog

## Unreleased

Repository/documentation hardening:

- Clarified that the GitHub repo is code-only and must not contain live workbooks or PHI.
- Added de novo setup instructions for Windows, Excel, Brother QL-1100C, workbook template, VBA import, and sheet event handlers.
- Added branching/release guidance: `main` stable, `dev` active integration, `feature/*` targeted changes.
- Added troubleshooting guidance for run-time error 9, `.bas` import failures, Brother paper-size issues, squished labels, macro blocking, and Git/OneDrive issues.
- Added `.gitattributes` to preserve CRLF line endings for `.bas` and `.vbs` files.
- Added no-PHI sample Tebra-like medication text for parser testing.

## 2026-06-29

Current workbook/source status from project handoff:

- Added checkbox-driven selection and `Print Checked Labels` workflow.
- Added `# of Prints` column and print-count logging.
- Added `Start NEW Patient` workflow that keeps the dispense log.
- Hid the internal `Label Preview` sheet and moved volunteer-facing previews to `Label Previews`.
- Removed lavender printed row state; row color now reflects selected/validated/non-validated state.
- Disabled logo import because `Shapes.AddPicture` could abort setup.
- Removed `PageSetup.PaperSize` handling because Brother DK labels are controlled by printer defaults.
- Changed print scaling to `Zoom = 100` with no fit-to-page behavior.
- Removed runtime VBProject injection in favor of preinstalled worksheet event handlers.
