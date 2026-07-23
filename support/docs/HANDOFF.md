# SCU Label Printing Tool - Handoff

Last updated: 2026-07-23 (**v2.1**: Check Med column moved left + frozen panes; Medications & Print Labels banners redesigned; readable `frmReview` list dialog reused for Review / Print confirm / Print complete / Remove; Add Medication uses the `frmMedEdit` form; **live yellow<->white highlighting** with a clear **white -> blue (reviewed) -> green (checked)** state and an `EnableEvents` guard invariant; Parse clears the previous list first (keeps name/DOB); removed the only clipboard `Copy`/`PasteSpecial`. Earlier **V2** (2026-07-09): Source column (required) right of Lot #; Refills right of Rx Date + default 0; 2 copies per label; auto-check on validate; per-print **Encounters** (numbered + green-banded) in the Log; **TEBRA TEMPLATE** sheet; Medications banner/grid spruce-up + hidden internal columns; **fit-to-page** so Exp/Lot never clip; resilient `ClearMedArea`; Add Medication prompts for Exp/Lot; red/yellow missing-field highlights)

An Excel + VBA tool for the **Saturday Clinic for the Uninsured (SCU)** free
pharmacy. It turns pasted prescription text into a validated medication table and
prints professional labels on a **Brother QL-1100C** (DK-1202, 62 x 100 mm,
landscape). It runs **100% offline** and is built to be robust for high-turnover
volunteers.

---

## 1. Where everything lives (IMPORTANT - read first)

Everything is in **one folder** (code, docs, workbook, backups, and Git):

```
C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing
```

This folder is **outside OneDrive** (Documents was moved off OneDrive sync in
June 2026). It contains the code, docs, the SC emblem, the workbook
(`MedicationDispensing.xlsm`), and `_backups\`. It is also the local Git repo
(remote: `JamesSRN/scu-label-tool`). The `.xlsm` and `_backups\` are git-ignored,
so patient data can never be committed.

**Two open setup items:**

1. **Trusted Location.** The old `Desktop\SCU Label Printing` folder (the prior
   Excel Trusted Location) was removed. The new folder is **not trusted yet**, so
   macros will be blocked until you add it:
   `Excel > File > Options > Trust Center > Trust Center Settings > Trusted
   Locations > Add new location > browse to the folder above > OK`.
2. **Easy access (optional).** Right-click the folder > Send to > Desktop
   (create shortcut). It can be renamed in place to drop the `GIT_VERSION_`
   prefix (then re-point GitHub Desktop with Locate).

Code/docs are version-controlled; the `.xlsm` (PHI) is local-only by design.

---

## 2. Workbook structure (5 visible tabs + hidden)

The visible workflow tabs are **numbered and color-coded** (V2): **1. Patient & Input**
(blue), **2. Medications** (green), **3. Print Labels** (orange, the gallery), **4. Log**
(purple), **5. Tebra Notes** (teal); **Start Here** is slate. Names are driven by the
`SH_*` constants and applied at build time by `EnsureSheetName` (idempotent rename, runs
first in `SetupWorkbook`) + `ColorWorkflowTabs`. Hidden helper sheets: **Label Preview**
(print surface), **EncounterData** (snapshots). **Developer Test** is the last tab.

- **1. Patient & Input** - patient name / DOB / Rx date + paste box. Buttons:
  `PARSE MEDICATIONS`, `Clear Paste Area`, `Reset Session`, `Start NEW Patient`.
- **2. Medications** - the parsed table (**V2 column order**, all driven by `C_*`
  constants). Columns A-Q: `#`(1), Name(2), Strength(3), Dosage Form(4), SIG(5),
  Quantity(6), **Expiration**(7), **Lot**(8), **Source**(9), Rx Date(10),
  Refills(11), Confidence(12), Warnings(13), Raw Text(14, **hidden**),
  Printed?(15, **hidden**), **"# of Prints"**(16, auto-increments, edit-protected),
  **"Check Med"**(17, double-click toggles a checkmark; formerly "Print?"). **Source**
  is a required dropdown (DOH / IN HOUSE / RxAPS / Other), blank by default. A full-width
  blue banner (row 1) and a light box-grid are drawn in code by `SetupWorkbook`.
  Missing-field highlights: **red** = Exp/Lot, **yellow** = Quantity/Source.
  Buttons (col S / 19): `+ Add Medication`, `- Remove Selected`, `Review & Validate`,
  `Preview ALL Labels`, `Save for Later (draft)`, `Edit Past Encounter`, `Save Edited
  Encounter`. **Printing is done only from the gallery** (Preview ALL Labels -> Print
  Checked Labels).
- **3. Print Labels** - auto-generated gallery, one card per medication (cards mirror
  the print label's three-zone header). **Top-right buttons** (code-created each
  rebuild): `Print Checked Labels`, `Refresh Previews`. **Per-card buttons**:
  **`Check this label` / `Uncheck this label`** (state-aware; toggles the med's
  `Print?` checkbox on the Medications tab via `RowCheck` -> `ToggleRowSelect`),
  `Edit this med`, `Remove this med`. Rebuilds on tab activate; the rebuild sweeps
  stray/manually-placed buttons (deletes all `al_*` shapes and loose autoshapes).
- **4. Log** - running dispense log. Columns 1-16: timestamp, patient, DOB, med,
  strength, directions, qty, refills, expiration, lot, **source**, Rx date,
  initials, dosage form, print #, **Encounter**. Each print is one **Encounter**
  (numbered, and its rows shaded in one of three cycling greens). Mirrored to a
  dated local CSV.
- **5. Tebra Notes** - paste-ready session notes built from the Log: one block per
  patient, grouped by source (IN HOUSE / DOH & Outside / RxAPs), with name/DOB on
  the right. Rebuilt by `FillTebraTemplate` (also has a "Refresh from session log" button).
- **Start Here** - the in-app quick-start guide; the workbook opens here.
- **Label Preview** (HIDDEN) - the internal print surface. Every label renders
  here and prints from here. Do NOT delete it.
- **EncounterData** (HIDDEN) - full per-encounter snapshots (one row per med) that power
  Edit Past Encounter. Cleared on full reset / close. Do NOT delete it.

Shortcuts: `Ctrl+Shift+P` parse, `Ctrl+Shift+R` reset, `Ctrl+Shift+L` refresh
Label Previews.

---

## 3. Feature set

### V2 highlights (2026-07-09)
- **Source** column (required): DOH / IN HOUSE / RxAPS / Other dropdown, blank by
  default, right of Lot #; missing = yellow; `ApplySourceValidation` keeps it the
  only dropdown in the grid. Mirrored into the Log (col 11) and CSV.
- **Refills** sits right of Rx Date and **defaults to 0**.
- **2 copies per label** via `LABEL_COPIES = 2` (one Log row per med regardless).
- **Auto-check on validate**: passing meds get their Print? checkmark automatically.
- **Add Medication** offers the same Exp/Lot popup the parse flow uses.
- **Encounters**: each print action is one numbered Encounter (`NextEncounter`),
  Log rows shaded in 3 cycling greens.
- **Fit-to-page** label print so the bottom Exp/Lot row can't clip (see §4).
- **`ClearMedArea`**: all clear paths wipe a fixed range, never miss unnamed rows.
- Medications tab: code-drawn blue banner + box grid; internal columns hidden.
- **Encounter** column moved to Log col 2 (right after Timestamp); whole Log now runs on an `LG_*` constant map.
- **Edit a past Encounter**: each print saves a full snapshot (hidden `EncounterData` sheet, `ES_*` map). **Edit Past Encounter** reopens one (pick by number) and restores the patient + all meds; **Save Edited Encounter** replaces that encounter's Log rows + snapshot in place (no dupes), rebuilds the Tebra note, and offers to reprint. Snapshots kept across Start NEW Patient, wiped on full reset / close.

### Parsing & validation
- Block splitter starts a new medication per drug-header line, even without blank
  lines between drugs.
- Per-drug **Confidence** (High / Medium / Low / Manual) from how cleanly it parsed.
- **Validation** flags missing strength/SIG/quantity/Exp/Lot/**Source**, bad Exp
  format, and possible duplicates -> Warnings column. Missing-field cell highlights:
  **red** = Exp/Lot, **yellow** = Quantity/Source.
- Add / Remove / Renumber medications manually. **Remove Selected removes every
  checked (Print?) row**, not just the clicked one.
- Exp/Lot forced to text.

### Selection, color, printing (checkbox-driven)
- **Double-click the Check Med cell OR the Medication name to toggle a checkmark.** A checkmark = "selected".
- Row background by state (priority **Selected > Reviewed > unreviewed**): **Selected
  (checked) = green** (matches the gallery tint), **Reviewed OK = blue**, **not yet
  reviewed = white**. Missing **Qty / Exp / Lot / Source** cells are **yellow**.
  Confidence cell colored by its own triad. (No printed/lavender state - printing no
  longer recolors rows.)
- **Live highlighting:** the Medications `Worksheet_Change` handler calls `LiveRefreshRow`
  on the edited row(s) so yellow clears to white the instant a value is entered (and
  returns to yellow if cleared) - **without** turning the row blue. Blue is only set when
  **Review** (`ValidateMedications`) runs; Review then auto-checks OK rows to green.
- **Event-guard invariant:** anything that programmatically writes a Medications table
  cell whose result must survive - `ValidateMedications` (Warnings "OK" + Exp normalize),
  the Review auto-check, `ToggleRowSelect` (Check Med), `ClearMedArea` - runs with
  `Application.EnableEvents = False`, because otherwise the write re-fires
  `Worksheet_Change` -> `LiveRefreshRow`, which clears the "OK" (dropping the row to
  white). Unchecking a reviewed med therefore goes green -> blue, not green -> white.
- **Print Checked Labels** prints every checked med in sequence (skips any checked
  row missing Exp/Lot). Print confirmations list exactly what will print (single =
  full detail block; batch = numbered list flagging skips).
- **Brother auto-select** (`SelectBrotherPrinter`) finds the QL-1100C via WMI and
  confirms before printing.
- **Progress popup** (`frmBusy` + `BusyShow`/`BusyHide`) covers the printer-lookup
  and page-setup delay after the confirm dialog, hiding just before the initials
  prompt (status-bar fallback if the form is absent).
- **Long medication names wrap** to two lines (>38 chars, 11 pt) instead of shrinking
  to a tiny single line; row 7 grows to 28 pt while spacer rows 13/14 shrink so the
  A1:H15 print height is preserved. EXP/LOT footer value shrinks by length (merged
  cells ignore Excel ShrinkToFit).

### Logging & multi-patient
- **Every print is logged**: timestamp, patient, DOB, medication, strength, SIG,
  quantity, refills, Exp, Lot, Rx date, dosage form, print #, and volunteer
  initials. Single AND batch prints; initials asked once per batch.
- **# of Prints** auto-increments on each print and is protected from manual edits.
- **Start NEW Patient** clears patient + meds but keeps the Log (back-to-back
  patients). **Reset Session** is a full wipe (patient + meds + Log) with confirmation.

### Data-entry forms & session hygiene

- **Expiration + Lot are asked per medication** after parsing (paired), with a
  forgiving `MM/YYYY` format check that warns but never blocks (`PromptExpiration`,
  `IsBadExpFormat`). A two-field popup **`frmExpLot`** shows both fields at once
  (`PromptExpLotPair`); if the form is unavailable it falls back to two input boxes.
- **Edit this med** opens **`frmMedEdit`** — a prefilled editor with **Medication +
  Strength bold at the top**, then Dosage form, Quantity, Directions (multi-line),
  Expiration, Lot (Refills excluded on purpose; edit it on the sheet). `RowEdit` ->
  `EditMedWithForm`, with a prefilled-input-box fallback.
- **Add Medication** opens the **same `frmMedEdit`** form, blank and titled "Add
  medication" (`AddMedicationRow` -> `AddMedWithForm`), so adding and editing look
  identical. Falls back to a chain of input boxes if the form is unavailable. Exp/Lot are
  written as text so `05/2028` isn't coerced to a date.
- **Review** opens **`frmReview`** — a **scrolling review list** (replaces the old dense
  MsgBox, which can't size or bold individual lines). Per medication it shows the **name +
  strength large & bold (14pt)** with that med's **error(s) stacked underneath, one per
  line, in a smaller 11pt font** (passing meds read green "Ready to print."; flagged meds
  read dark with red error lines). A footer sums up how many still need attention.
  `ReviewMedications` populates it via the form's `ResetList` / `AddMed` / `SetFooter` /
  `FinishList` methods (`ReviewErrText` splits the period-separated warning string into
  brief per-line issues); a **MsgBox fallback** is used if the form is unavailable.
  Implementation note: the per-med labels are added to the scrolling frame **at run time**
  with **fixed width + fixed height and `WordWrap`, never `AutoSize`** — a runtime-added
  label with `AutoSize = True` collapses its width to one character.
- **`frmReview` is a reusable list dialog.** Besides Review it also backs the **Print
  Checked Labels** confirmation (Print/Cancel; each med Ready / SKIPPED-missing-Exp/Lot /
  DOH), the **Print Complete** summary (OK; each med's outcome), and the **Remove Selected**
  confirmation (Remove/Cancel; names only). API: `ResetList`, `AddMed(title, errText, isOK)`
  (empty `errText` = compact name-only row), `SetHeader(title)`, `SetFooter(text)`,
  `ConfigButtons(showCancel, okText, cancelText)`, `FinishList`, then `.Show` and read
  `.Result` ("OK"/"CANCEL"). Every caller sets header + buttons because the default
  instance persists between opens.
- **Auto-reset (PHI hygiene):** on **open** the workbook clears patient + meds + paste
  (`ClearSessionSilent`, Log kept); on **close** it does a **full wipe** — patient +
  meds + paste **and the Log** (`ClearSessionSilent` + `ClearLogSilent`) — then saves,
  so no patient data is left on disk. Via `Workbook_Open` / `Workbook_BeforeClose`
  (installed by Build-Release, which now **replaces** the ThisWorkbook module so a stale
  `Workbook_Open` can't block it). Note: **in-progress work is discarded on close by design.**

All four UserForms (`frmExpLot`, `frmMedEdit`, `frmBusy`, `frmReview`) and the
auto-reset handlers are **generated by `Build-Release.vbs`** into the workbook (so
volunteers need no VBA-project-trust setting). The forms are referenced **late-bound**
in `MedParser.bas`, so the module compiles even without them. See §5 (build) and §7
(gotchas).

---

## 4. The redesigned label (current)

Polished, clinic-grade DK-1202 label. Printed area = **A1:H15**, landscape.

Header is **three zones**: clinic name (two lines, left) | SC emblem (centered) |
phone-over-address (right).

```
SATURDAY CLINIC          ( SC )              (414) 588-2865
for the Uninsured       emblem      1121 E. North Ave, Milwaukee WI
-----------------------------------------------------------
Doe, Jonathan                            Rx    06/29/2026
                        (name centered)   DOB   03/14/1985
Metronidazole 500 mg          <- HERO line (largest, bold)
Tablet  .  Qty 56  .  Refills 2
DIRECTIONS
Take 1 tablet by mouth twice daily          (3-line block,
with food until the course is finished.      white on black)
-----------------------------------------------------------
EXP  05/2027                   LOT  ABC1234   <- pinned to bottom row
```

Cell map (rows 16–18 hold off-label helper text that does not print):

| Field | Cell(s) |
|---|---|
| Clinic name line 1 "SATURDAY CLINIC" (Century Gothic, bold) | **A2:C2** (merged) |
| Clinic name line 2 "FOR THE UNINSURED" (small) | **A3:C3** (merged) |
| Emblem slot (centered, no fill) | **D2:E3** (merged) |
| Phone (Arial bold, right, bottom) | **F2:H2** (merged) |
| Address (Arial, right, top) | **F3:H3** (merged) |
| Patient name (vertically centered on the Rx/DOB block) | A5:E6 (merged) |
| Rx date | F5:H5 (right) |
| DOB | F6:H6 (right) |
| Medication + strength (hero) | A7:H7 |
| Dosage form + qty **+ refills** | A8:H8 |
| "DIRECTIONS" mini-label | A9:H9 |
| SIG / directions (**3 lines**) | A10:H12 (white on black) |
| EXP (**bottom row**) | A15:D15 |
| LOT (**bottom row**) | E15:H15 (right) |

**Design / type:** Hierarchy is carried by size + weight + small uppercase labels
(never color - thermal is monochrome). The **clinic name uses Century Gothic**
(`FONT_LABEL_HDR = "Century Gothic"`, resolved by `LabelHeaderFont()` with an
**Arial** fallback if the PC lacks it); the phone/address and body use **Arial**.
Two thin hairline rules (under the header, above the footer); no heavy box. The
medication name is the focal point; EXP/LOT are pinned to the bottom row.

**Adaptive sizing:**

| Element | Logic |
|---|---|
| Patient name | `PatientNameFontSize()` — steps 17→8 pt, always **≥ med line + 1 pt**; vertically centered on the Rx/DOB block |
| Medication / SIG | `NameFontSize()` / `SigFontSize()` — steps down for long text |
| Clinic name line 1 | `FmtLblClinicName()` (Century Gothic, **bold**, left, bottom) — `CLINIC_NAME_FONT_PRINT = 14` / gallery 12; `ShrinkToFit` |
| Clinic name line 2 | `FmtLblNameSub()` (Century Gothic, left, top) — `CLINIC_NAMESUB_FONT_PRINT = 7` |
| Phone (top-right) | `FmtLblContactRight(..., True, "B")` (Arial bold, right) — `CLINIC_PHONE_FONT_PRINT = 11` |
| Address (below phone) | `FmtLblContactRight(..., False, "T")` (Arial, right) — `CLINIC_ADDR_FONT_PRINT = 8.5` |
| Header row heights | **18 pt** (row 2) + **12 pt** (row 3) = 30 pt band; print area rows 1–15 total **170 pt** |

**Print width:** `LABEL_WIDTH_PT = 242` scales columns A:H via
`ApplyLabelContentWidth` so content uses more of the 100 mm die-cut. **228 pt** was
the prior sweet spot (no bleed); **218 pt** was too narrow. Re-test on the clinic
Brother after any width change.

**Logo / emblem:**

| Item | Detail |
|---|---|
| Manual source crop | `cropped_Black SCU Logo + Transparent Background - Copy.png` |
| Workbook file | `scu_emblem.png` (built by `tools/Build-ScuEmblem.ps1`) |
| Build script | `Build-ScuEmblem.ps1` is meant to copy the crop and force ink to solid black (`#000000`) for crisp thermal output. **It is currently BROKEN (zeros the alpha -> fully transparent PNG) — do not run it** until repaired; `scu_emblem.png` is maintained by hand for now. |
| Aspect ratio | `LOGO_ASPECT = 1.488` (6150×4133 px) |
| Print height | **`LOGO_HEIGHT_PRINT = 30 pt`**, **centered in the D2:E3 slot** (both axes) |
| Gallery height | **`LOGO_HEIGHT_GALLERY = 28 pt`**, **centered in the C:D card slot** |
| Horizontal placement | **centered** — `InsertLabelLogo(..., centerHoriz:=True)` for both surfaces (was flush-right) |
| Insertion | `InsertLabelLogo` — insert at natural aspect (never `maxW × height` box); `LockAspectRatio`; `ZOrder msoBringToFront`; `xlFreeFloating`; centered in band + slot after final size |
| Print refresh | `RefreshPrintLabelLogo` runs on every preview update and before each print (gallery rebuilds logos on tab activate; print sheet does not) |
| Resolution | **`LogoFilePath()` uses only local `scu_emblem.png` beside the workbook** (embedded `LogoB64()` fallback removed). The workbook is in a local Documents folder, so `ThisWorkbook.Path` is a real path and `Dir()` finds the file. |

Do **not** re-enable automatic logo cropping without explicit approval — the
manual crop is the source of truth.

**CRITICAL emblem-blank bug (fixed 2026-07-02):** the logo vanished from BOTH the
print label and every gallery card. Root cause: **`scu_emblem.png` was a fully
transparent (blank) image** — `Build-ScuEmblem.ps1` had regenerated it with the
alpha channel zeroed, so `AddPicture` was inserting an invisible picture (the shape
existed; nothing showed). The code was fine. Fix: regenerated `scu_emblem.png` from
the source crop by forcing ink to solid black **while keeping the alpha** (133 K
opaque px, verified). **`Build-ScuEmblem.ps1` is still buggy — do NOT run it until
it's repaired**, or it will blank the emblem again. Diagnostic if it recurs:
Immediate window `?"[" & Dir(ThisWorkbook.Path & "\scu_emblem.png") & "]"` (path OK),
then check the PNG actually has opaque pixels.

**Print geometry:** `ApplyLabelPageSetup` sets `A1:H15`, **fit-to-one-page**
(`.Zoom = False`, `.FitToPagesWide = 1`, `.FitToPagesTall = 1`), **no `.PaperSize`**
(Brother driver controls media), 0.04 in margins, landscape, no gridlines/headings.
Fit-to-page (V2) replaced the old `Zoom = 100` / no-fit so the bottom **EXP/LOT row
(row 15)** can't spill onto a never-printed page 2 on printers with a shorter
printable area (the "Lot/Exp missing on print, fine in the gallery" bug). If a label
looks slightly small, trim the spacer rows rather than disabling fit-to-page.

**Blank label after each print (fixed):** Excel was sometimes sending **page 2**
(empty). `PrintLabelSurfaceSafe` now calls `PrepareLabelSheetForPrint` (only
`scuLogo` may print as a shape; preview buttons excluded; `ResetAllPageBreaks`)
and **`PrintOut From:=1, To:=1`**.

**SetupWorkbook fix:** header `PasteSpecial` runtime 1004 resolved via
`MatchHeaderFormat` (safe format copy for new Medications/Log columns).
`SetupWorkbook` also calls **`PreviewAllLabels`** after layout rebuild so gallery
and print sheet stay in sync.

See also `LABEL_REDESIGN.md` and `CHANGELOG.md` (Unreleased).

---

## 5. Applying code changes (release build)

`MedParser.bas` must stay **pure ASCII with Windows CRLF** line endings or Excel
rejects the import. **No UTF-8 BOM** on line 1. (`.gitattributes` enforces CRLF
for `.bas` / `.vbs`.)

### Quick path (recommended)

1. **Close `MedicationDispensing.xlsm` completely** (check Task Manager for stray
   `Microsoft Excel` processes). If it stays open, Build-Release works on a second
   copy and Excel asks to **"replace the already present document"** — answering
   **Yes** overwrites the freshly-built file with the stale open copy and **wipes the
   logos**. Close it first so the prompt never appears.
2. **Do NOT run `tools/Build-ScuEmblem.ps1`** — it is currently broken and blanks
   `scu_emblem.png` (see §7 gotchas). The emblem is already correct; leave it.
3. Double-click **`Build-Release.vbs`** (needs **Trust access to the VBA project
   object model** + folder as **Trusted Location**).
4. Click OK on **Setup complete!** and **Release build complete**.

`Build-Release.vbs` opens `MedicationDispensing.xlsm` in place (bootstraps from
`Broken_PrettyPrint_MedicationDispensing.xlsm` only if missing), removes the old
`MedParser` module, imports `MedParser.bas`, runs
`'MedicationDispensing.xlsm'!SetupWorkbook`, saves.

**Bootstrap warning:** if `MedicationDispensing.xlsm` is missing, the script copies
from the bootstrap template and shows a **MsgBox** that patient data is **not**
restored — check `_backups\` or Windows File History.

Optional: `tools/Run-BuildRelease.ps1` — same via Excel COM with MsgBox auto-dismiss.

See `tools/BUILD_RELEASE_NOTES.md` for full detail.

### Manual path

1. VBA editor (Alt+F11): remove the old `MedParser` module, **Import
   `MedParser.bas`**, run **`SetupWorkbook`**, save. (`SetupWorkbook` running is
   the compile check.)
2. **Handlers:** the **Medications** `Worksheet_BeforeDoubleClick` (toggle Print? /
   block # of Prints) **and** `Worksheet_Change` (live highlight refresh: any table-cell
   edit re-runs `ValidateMedications`, guarded by `EnableEvents = False`) are both
   **re-installed at build time** by `SetupWorkbook` -> `InstallMedSheetEvents`, built
   from the `C_*` constants (injected as literals, since a sheet module can't see the
   module's `Private Const`s) so they stay correct after column reorders (needs
   VBA-project trust, which Build-Release has). The **Label Previews**
   `Worksheet_Activate` -> `PreviewAllLabels` handler is preinstalled in that sheet module.
3. Keep `scu_emblem.png` beside the workbook so the logo embeds on SetupWorkbook.

### Problems solved (2026-06-30 session)

| Issue | Fix |
|---|---|
| Runtime 1004 on header PasteSpecial | `MatchHeaderFormat` |
| Labels bleed to 2nd die-cut | `LABEL_WIDTH_PT` tuning (228 safe; now 242 for wider template — re-verify) |
| Blank label after each print | `PrintOut From:=1, To:=1`; `PrepareLabelSheetForPrint`; only `scuLogo` prints as shape |
| Emblem sliver (merged cells hid logo) | Logo slot **G2:H3**; text **A2:F2** / **A3:F3**; `ZOrder msoBringToFront` |
| Gallery OK, print sheet wrong logo | `RefreshPrintLabelLogo` on preview update + before each print |
| Logo squished wide | `InsertLabelLogo` inserts at natural aspect, not `maxW × height` box |
| Gray snake on thermal | `Build-ScuEmblem.ps1` forces solid black pixels |
| Clinic name clipped top/bottom | `FmtLblHeader` (no wrap) + row heights 20 + 10 pt; bottom/top vertical align |
| Clinic name too small | Wider text merge **A2:F2** (was A2:E2); `CLINIC_NAME_FONT_PRINT = 14` |
| Clinic name not bold enough | `FmtLblClinicName` — always bold via `Font.Bold` only |
| Compile: `xlBold` / `Font.Weight` | Removed — use **`Font.Bold = True`** only (Weight/xlBold break compile) |
| Weird partial logos after rebuild | Removed embedded `LogoB64()` fallback from `LogoFilePath`; reduced inset; in-band vertical centering |
| Build replaced workbook / lost data | Bootstrap only when `.xlsm` missing + warning MsgBox |
| VBScript / VBA compile | See table below |

### Compile / build fixes (2026-06-30)

| Issue | Fix |
|---|---|
| VBScript "Expected end of statement" | No typed `Dim` in `Build-Release.vbs` |
| VBA import / compile failures | Restore `FONT_LABEL_HDR_FB`; optional defaults must be literals; strip UTF-8 BOM; `msoPictureCompressNone` → `0` |
| `xlBold` / `Font.Weight` in VBA | Do not use — `Font.Bold` only |

---

## 6. TASKS STILL TO COMPLETE / VERIFY

1. **Add the consolidated folder as an Excel Trusted Location** (section 1) so
   macros run without prompts. Without this, content is blocked.
2. **(Optional) Desktop access:** make a Desktop shortcut to the folder; rename
   to drop `GIT_VERSION_` and re-point GitHub Desktop.
3. **`LABEL_WIDTH_PT = 242`** — user wanted a wider template; confirm on the clinic
   Brother that labels still fit **one** die-cut (228 pt was the prior no-bleed
   sweet spot).
4. **Bold clinic title on thermal** — bold may still look subtle on the Brother;
   `ShrinkToFit` can limit apparent size on long clinic name text.
5. **Physical print regression** — spot-check emblem size/position and header
   after the latest logo/header changes (user confirmed screen + print path OK
   at end of session).
6. **Bootstrap / data loss** — if `MedicationDispensing.xlsm` was ever missing
   during a build, the bootstrap copy overwrote it; recover from `_backups\` if
   needed.
7. **Dead code cleanup (optional):** `LogoB64()` remains in `MedParser.bas` but is
   unused after fallback removal — safe to delete later to shrink the module.
8. **Repair `tools/Build-ScuEmblem.ps1`** — it currently **blanks** `scu_emblem.png`
   (zeros the alpha channel), which made the logo invisible everywhere on 2026-07-02.
   Until it's fixed, do **not** run it. Correct behavior: force ink to solid black
   **while keeping the alpha channel** (or emit black-on-white fully opaque).
   `scu_emblem.png` has been regenerated by hand for now.

**Done / verified (2026-06-30):** release build (`Build-Release.vbs`); label layout
and typography; emblem (aspect, pure black, print + gallery sync); one label per
print (no blank follow-on label); logo/header layout on screen and print path
(user confirmed working at end of session).

---

## 7. Challenges & gotchas (for the next dev)

- **CRLF + ASCII required; no UTF-8 BOM.** Editing `.bas` on Linux/Mac (LF endings
  or BOM) makes the VBA importer reject the file. Always save CRLF + pure ASCII.
- **`MatchHeaderFormat` for new columns.** Do not raw `PasteSpecial` header formats
  onto new Medications/Log columns — use the helper (fixes runtime 1004).
- **`Shapes.AddPicture` / logo** — use local `scu_emblem.png` only. The emblem is
  **centered in D2:E3** (print) / C:D (gallery); text in **A2:C** (name) and
  **F2:H** (phone/address); bring shape to front; never insert at `maxWidth × height`
  (squashes aspect). Do **not** re-enable `LogoB64()` fallback.
- **`Build-ScuEmblem.ps1` is currently BROKEN** — it regenerated `scu_emblem.png`
  as a fully transparent (blank) image (alpha zeroed), making the logo invisible on
  every surface even though `Dir()` finds the file and `AddPicture` succeeds. **Do
  not run it until repaired.** If the emblem disappears, confirm the PNG actually has
  opaque black pixels — not just that the file exists.
- **Gallery vs print logo** — gallery rebuilds on tab activate; print sheet needs
  `RefreshPrintLabelLogo` on every `UpdateLabelPreviewForMedRow` / print.
- **Blank label between prints** — if it returns, check Excel print preview page
  count; ensure `PrintOut From:=1, To:=1` and preview buttons have
  `PrintObject = False`.
- **`Font.Bold` only for bold headers.** Do not use `Font.Weight` or `xlBold` —
  they caused compile errors in this project.
- **VBProject injection at build only.** The Medications double-click handler is
  re-installed by `SetupWorkbook` -> `InstallMedSheetEvents` at **build time** (where
  Build-Release already has VBA-project trust), so it tracks the `C_*` columns after
  a reorder. Volunteers still need no trust setting - they run the baked-in handler.
  Do not call `InstallMedSheetEvents` from normal-use code paths.
- **`PaperSize = xlPaperUser` throws 1004** on this Brother driver - never set it;
  the driver's DK-1202 default is correct. The label sheet uses **fit-to-page**
  (`.FitToPagesWide/Tall = 1`, `.Zoom = False`) so the bottom row can't clip.
- **OneDrive dehydration:** OneDrive can mark files "cloud-only" (cloud icon).
  They re-download on open. Set the folder to "Always keep on this device" to stop
  it (helps automation/editing).
- **OneDrive + Git:** never run non-Windows Git/automation against a `.git` folder
  inside OneDrive (lock errors, can corrupt `.git/config`). And OneDrive blocks
  *moving* a folder that contains `.git` (error 0x80004005) - use GitHub Desktop on
  native Windows, and don't try to relocate the repo folder via Explorer.
- **Macros blocked by Mark-of-the-Web** when the file lives in OneDrive - fixed via
  Trusted Location.
- **Close the workbook before Build-Release.** If `MedicationDispensing.xlsm` is
  open while `Build-Release.vbs` runs, Excel prompts **"replace the already present
  document."** **Yes** lets the stale open copy overwrite the freshly-built one and
  **wipes the emblems** — always answer **No**, or (better) close the workbook first
  so the prompt never appears. (`Build-Release.vbs` and `MedParser.bas` contain no
  file-replace logic; this is purely the file being open twice.)

---

## 8. Key VBA routines (`MedParser.bas`)

| Routine | Purpose |
|---|---|
| `SetupWorkbook` | Run after each import: rebuilds buttons, headers, colors, label layout; calls `PreviewAllLabels`. Uses `MatchHeaderFormat` for new columns. |
| `ParseMedications` / `SplitMedBlocks` / `IsMedHeaderLine` / `ParseOneBlock` | Parse pasted text -> one row per med. **Clears the previous list first** (`ClearMedArea` + reset batch/encounter state; keeps name/DOB/paste), confirming when meds already exist. |
| `LiveRefreshRow(r)` | Live recolor of one edited row (yellow<->white) without validating/bluing; clears a prior "OK" so an edited reviewed row returns to white. Called by the Medications `Worksheet_Change`. |
| `ValidateMedications` / `ReviewMedications` / `ReviewErrText` | `ValidateMedications` flags issues and writes the `Warnings` cell; `ReviewMedications` shows the scrolling `frmReview` (name/strength bold + stacked errors) with a MsgBox fallback; `ReviewErrText` splits the warning string into brief per-line issues. |
| `AddMedicationRow` / `RemoveSelectedMedication` / `RenumberMeds` | List editing. Remove Selected removes all checked rows. |
| `ApplyRowState` / `ApplyAllRowStates` / `IsRowSelected` / `ToggleRowSelect` | Row color by state; confidence triad; selection checkmark. |
| `PrintCheckedLabels` | Batch-print every checked med; logs each. |
| `BuildLabelPreviewLayout` / `ApplyLabelContentWidth` / `ApplyLabelPageSetup` / `UpdateLabelPreviewForMedRow` | Label layout, width scaling (`LABEL_WIDTH_PT = 242`), page setup, value writing (incl. **Refills** on the qty line; **EXP/LOT on bottom row 15**). |
| `FmtLbl` / `FmtLblHeader` / `FmtLblClinicName` / `FmtLblNameSub` / `FmtLblContactRight` | Cell formatting. Header three-zone helpers: clinic name lines (Century Gothic), phone/address (Arial, right-aligned). |
| `InsertLabelLogo` / `RefreshPrintLabelLogo` / `EnsurePrintLabelHeaderLayout` / `LogoFilePath` | Emblem insert (30 pt print / 28 pt gallery), **centered** (`centerHoriz`), refresh before print; three-zone header merges A2:C2 / A3:C3 / **D2:E3** (emblem) / F2:H2 / F3:H3. |
| `PrepareLabelSheetForPrint` / `PrintLabelSurfaceSafe` | Suppress extra shapes; single-page `PrintOut From:=1, To:=1`. |
| `LabelHeaderFont` / `SetMiniValue` / `MedFontSize` / `NameFontSize` / `PatientNameFontSize` / `SigFontSize` | Fonts and adaptive sizing. `SetMiniValue` shrinks the EXP/LOT value by length (merged cells ignore `ShrinkToFit`). Med name >38 chars wraps to two 11 pt lines (row 7 -> 28 pt, spacers 13/14 -> 1 pt). |
| `BusyShow(pct, msg)` / `BusyHide` | "Please wait" progress popup (`frmBusy`) shown during the Print Checked Labels printer-lookup + page-setup delay; status-bar fallback if the form is missing. |
| `MatchHeaderFormat` | Safe header format copy (PasteSpecial 1004 fix). |
| `PrintLabel` / `PrintCheckedLabels` / `SelectBrotherPrinter` / `MarkPrinted` | Auto-select Brother, confirm, print, bump # of Prints. |
| `LogPrint(medRow, vol, encounter)` / `AskInitials` / `NextEncounter` | Full per-print Log row (single + batch), stamped with the Encounter # (Log col 2) and shaded by one of 3 cycling greens. `NextEncounter` = highest Encounter (base number, via `Val`) in the Log/snapshots + 1. |
| `EncLabel` / `EncounterNextVersion` / `ParseEncVersion` / `gEncLabel` | **Encounter versioning:** re-saving an edited encounter stamps its Log rows `1`, `1 (v2)`, `1 (v3)`, ... The numeric identity is unchanged - matching/next-number/banding/dividers/Edit-picker all read the base number (`Val`; snapshot store stays numeric). `gEncLabel` carries the display label into `LogPrint`; cleared after each log pass. |
| `ApplyLogEncounterBorders` | Draws a dark divider rule above each new encounter group in the Log (grouped by base number). Recomputed from scratch; called after every Log change (print, reprint, save-edited-encounter, rebuild). |
| `EncStore` / `SaveEncounterSnapshot` / `ClearEncounterStore` / `DeleteRowsByEncounter` | Hidden `EncounterData` snapshot store (full med detail per encounter); snapshot on each print; wipe on full reset/close; delete rows for one encounter. |
| `EditEncounter` / `LoadEncounter` / `SaveEditedEncounter` / `PrintEncounterLabelsNoLog` | Reopen a past encounter (list + pick), restore patient + meds; save = replace that encounter's Log + snapshot, rebuild Tebra, optional reprint (no re-log). |
| `FillTebraTemplate` / `TebraPatientBlock` / `TebraLogSection` / `TebraLogMedLine` / `TebraCardBorder` | Build the 5. Tebra Notes sheet from the Log: one **green outer card** per patient containing **two inner teal boxes** (Medication Reconciliation; Dispensed + Counseling), A:L wide. Empty sections show just the title (no placeholder text). Refresh button lives in the banner. Rebuild **keeps pictures** so a hand-added emblem persists. |
| `PlaceSetupHelpSheet` | Moves Setup & Help after Tebra and appends the "how it works" blurb. **Idempotent:** clears from the first "HOW THIS TOOL WORKS" marker down before re-appending, so blurbs don't stack across builds. Does not touch column widths / row heights (manual sizing there persists). |
| `ApplySourceValidation` / `ClearMedArea` | Source dropdown (only dropdown in the grid); fixed-range wipe of the med area used by every clear path + build. |
| `InstallMedSheetEvents` | Re-inject the Medications double-click handler at build (uses `C_*` constants). |
| `StartNewPatient` | Clear patient + meds (via `ClearMedArea`), keep the Log. |
| `PreviewAllLabels` / `BuildAllLabelsPreview` / `EnsureAllLabelsSheet` | Label Previews gallery (three-zone cards mirroring the print label; top-right `Print Checked Labels` / `Refresh Previews` created each rebuild; rebuild sweeps stray autoshapes). |
| `RowCheck` / `RowEdit` / `RowRemove` (`RowPrint` legacy) / `EditMedWithForm` | Per-card gallery actions. `RowCheck` -> `ToggleRowSelect` toggles the med's `Print?` selection; `RowEdit` -> `EditMedWithForm` opens the prefilled `frmMedEdit` editor (input-box fallback); both rebuild the gallery. |
| `PromptExpLotPair` / `PromptExpiration` / `IsBadExpFormat` | Per-med Exp+Lot prompt via `frmExpLot` (two input-box fallback), with a forgiving `MM/YYYY` format check. |
| `ClearSessionSilent` | Silent "Start NEW Patient" (clear patient+meds+paste, keep Log). Called by the `Workbook_Open` / `Workbook_BeforeClose` auto-reset. |
| `ResetSession` / `ClearPasteArea` | Full reset (incl. Log) / clear paste box. |

---

## 9. HIPAA / PHI

Patient name/DOB/meds are PHI. Keep the `.xlsm` and the dispense Log **local** -
never commit them to GitHub (the `.gitignore` excludes `*.xlsm`/`*.xlsx`/`*.csv`
and `_backups/`) and never route PHI through an external AI/API without a signed
BAA. The repo is code + docs + no-PHI samples only.

---

**Bottom line:** parsing, validation, checkbox-driven selection/printing/removal,
the dispense log, multi-patient flow, and the **redesigned DK-1202 label** (emblem,
header typography, width, single-page print) are built and **verified on screen
and print path** at end of the 2026-06-30 session. Everything lives in one
folder; code is backed up to GitHub. Remaining: **Trusted Location**, optional
**242 pt width re-test** on physical Brother, and optional cleanup of unused
`LogoB64()`. Detail: `LABEL_REDESIGN.md`, `CHANGELOG.md`,
`tools/BUILD_RELEASE_NOTES.md`.
