# Build Release Notes

The project handoff describes a local `Build-Release.vbs` helper that re-imports `MedParser.bas`, runs `SetupWorkbook`, and saves the workbook.

This repository update does not replace the existing local script. Before committing any build script, review it line by line and confirm it:

- opens only the intended local no-PHI workbook,
- imports only `src/MedParser.bas` or the intended local `MedParser.bas`,
- does not access the internet,
- does not delete files outside the target module replacement workflow,
- does not operate on PHI-containing workbooks unless the workflow is approved locally,
- does not require volunteers to run it during normal clinic operation.

Recommended separation:

```text
GitHub repo        source/docs/no-PHI samples
Build-Release.vbs developer convenience only
.xlsm workbook     local operational artifact, not committed
```

For normal clinic use, volunteers should open the already-built `.xlsm` workbook and should not need to run the build script.
