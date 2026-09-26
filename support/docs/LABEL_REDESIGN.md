# Label redesign

Current printed label in `MedParser.bas` (`BuildLabelPreviewLayout`). The merges and constants below match that routine. Older notes that put the clinic name in A2:F2 and the emblem in G2:H3 are out of date.

## Physical target

- **Printer:** Brother QL-1100c (Hermione), USB
- **Media:** DK-1202 die-cut, **62 x 100 mm**, landscape
- **Print area:** `A1:H15` on the hidden `Label Preview` sheet
- **Copies:** `LABEL_COPIES = 2` for Print Checked Labels. Gallery Print extra and Log-row Print use 1 copy and are not logged.

## Layout (rows 1–15)

| Rows | Content |
|------|---------|
| 2 | "SATURDAY CLINIC" — **A2:C2**, Century Gothic, 14 pt |
| 3 | "FOR THE UNINSURED" — **A3:C3**, 7 pt |
| 2–3 (center) | Emblem slot — **D2:E3**, no fill, 30 pt logo |
| 2 | Phone — **F2:H2**, 11 pt, right |
| 3 | Address — **F3:H3**, 8.5 pt, right |
| 5–6 | Patient name **A5:D6** (shrinks 12→8 pt), DOB **E5:H6** |
| 7 | Medication + strength (hero), **A7:H7** |
| 8 | Form · Qty · Source · Refills in **A8:E8**; Rx date in **F8:H8** |
| 9 | "DIRECTIONS" |
| 10–12 | SIG, white on black, **A10:H12** |
| 15 | EXP **A15:D15**, LOT **E15:H15** |

Rows 17–18 are off-label helper text and do not print.

Do not merge A2:H2. A fill on a merge that covers the emblem slot hides the logo.

## Typography

| Element | Font | Notes |
|---------|------|--------|
| Clinic name | Century Gothic, fallback Arial | `FONT_LABEL_HDR`, `CLINIC_NAME_FONT_PRINT = 14` |
| Name subtitle | Arial | `CLINIC_NAMESUB_FONT_PRINT = 7` |
| Phone / address | Arial | 11 pt phone, 8.5 pt address |
| Patient name | Arial | `LabelNameFontSize` steps 12→8 pt |
| Body | Arial | `FONT_LABEL_BODY` |

Header row heights in code: row 2 = `LOGO_HDR_ROW1_PT` (**18**), row 3 = `LOGO_HDR_ROW2_PT` (**12**).

**Compile note:** use `Font.Bold = True` only. Do not use `Font.Weight` or `xlBold`.

## Print width

- **`LABEL_WIDTH_PT = 242`** scales columns A:H.
- **228 pt** was the prior no-bleed width. Re-test 242 on the clinic Brother before treating it as final.

## Logo / emblem

| Item | Value |
|------|--------|
| Workbook file | `scu_emblem.png` next to the workbook |
| Aspect | `LOGO_ASPECT = 1.488` |
| Print height | 30 pt (`LOGO_HEIGHT_PRINT`) |
| Gallery height | 28 pt |
| Insertion | Natural aspect, brought to front, centered in D2:E3 |

**Do not run `tools/Build-ScuEmblem.ps1`.** It zeros the alpha channel and the emblem disappears. `LogoFilePath()` uses only `scu_emblem.png`. The embedded `LogoB64()` fallback is not used.

## Print pipeline

- `ApplyLabelPageSetup` — print area A1:H15, landscape, fit-to-one-page (`.Zoom = False`, `.FitToPagesWide = 1`, `.FitToPagesTall = 1`). Do not set `PaperSize`.
- `PrintLabelSurfaceSafe` — `PrintOut From:=1, To:=1`.

## Related constants

```text
LABEL_WIDTH_PT = 242
LABEL_COPIES = 2
PRINT_COPIES = 1
LOGO_ASPECT = 1.488
LOGO_HEIGHT_PRINT = 30
LOGO_HEIGHT_GALLERY = 28
LOGO_HDR_ROW1_PT = 18
LOGO_HDR_ROW2_PT = 12
CLINIC_NAME_FONT_PRINT = 14
CLINIC_NAME_FONT_GALLERY = 12
CLINIC_ADDR_FONT_PRINT = 8.5
CLINIC_NAMESUB_FONT_PRINT = 7
CLINIC_PHONE_FONT_PRINT = 11
FONT_LABEL_HDR = "Century Gothic"
FONT_LABEL_HDR_FB = "Arial"
FONT_LABEL_BODY = "Arial"
LOGO_EMBLEM_FILE = "scu_emblem.png"
```
