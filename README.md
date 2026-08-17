<div align="center">

<img src="support/assets/scu-emblem-badge-v2.png" alt="Saturday Clinic for the Uninsured emblem" width="130">

# SCU Label Printing Tool

</div>

---

<div align="center">

**Paste a patient's prescriptions → get clean, validated medication labels — printed on a Brother QL-1100c, with a running dispense log and paste-ready Tebra notes.**

Saturday Clinic for the Uninsured · Medical College of Wisconsin

`Excel + VBA` · `Brother QL-1100c — DK-1202` · `offline — no PHI committed` · `v2.1`

</div>

---

## Start here

This is an Excel + VBA tool for the **Saturday Clinic for the Uninsured (SCU)** free pharmacy. It turns a block of pasted prescription text into printed medication labels and keeps a dispense log.

It's built to be **robust for high-turnover volunteers**: it parses, checks, and flags every entry, so a first-time user is guided toward correct labels — and caught before a mistake reaches a patient.

**How do I open it?** Always start the tool with one launcher — it applies the latest code and opens the workbook for you.

| | Do this |
|---|---|
| **Open the tool** (every time) | Double-click **`OPEN LABEL TOOL (double-click me).cmd`** |
| **Fallback** (only if a build won't run) | Open **`MedicationDispensing.xlsm`** directly |

> The launcher runs `Build-Release.vbs`, which rebuilds the workbook from the latest `MedParser.bas` and opens it — so you're never on stale code. Make sure the workbook is **closed** first. Everything else (`support/`, `tools/`, the `.bas` source, the emblem, and the bootstrap workbook) is support material the build needs — don't move it.

> **Install:** grab the latest `SCU-Label-Printing-vX.Y.zip` from the **[Releases page](https://github.com/JamesSRN/scu-label-tool/releases)**, unzip, and open `MedicationDispensing.xlsm`. Full steps in **[RELEASE.md](support/docs/RELEASE.md)**.

---

## How it works — five numbered tabs

The whole workflow lives on **five color-coded, numbered tabs**. The workbook opens to a built-in guide that walks you through them:

![Start Here guide](support/assets/screenshots/start-here-v2.png)

_All screenshots below use randomly generated test patients — no real patient data._

### 1 · Patient & Input  (blue)

Enter the patient's **name + DOB**, paste the prescription text into the big box, and click **PARSE MEDICATIONS** (or press `Ctrl+Shift+P`). Parse **clears the previous patient's list first** (keeping the name/DOB and the Log) so a new patient never mixes with old rows.

### 2 · Medications  (green)

Each drug becomes one row — name, strength, form, directions (SIG), quantity, refills, expiration, and lot. **Refills default to 0** and a **Source** dropdown (DOH / IN HOUSE / RxAPS / Other) is required.

Highlights update **live as you type** — a missing **Quantity / Expiration / Lot / Source** cell is **yellow** and clears to **white** the instant you fill it. Then click **Review**: complete rows validate to **blue**. Review no longer auto-checks anything — you choose what prints: **double-click a Check Med cell** to check one (it turns **green**), or **double-click the "Check Med" column header** to check / uncheck the whole list at once. The Review pop-up lists each drug's **name + strength in bold** with any issues stacked beneath.

![Medications tab](support/assets/screenshots/medications-tab-v2.png)

![Medication Review — errors listed under each drug, missing cells highlighted yellow](support/assets/screenshots/review-dialog-v2.png)

> **Row colors at a glance:** yellow cell = missing · **white** = filled, not yet reviewed · **blue** = reviewed OK · **green** = checked to print. Double-click a **Check Med** box (or the medication name) to check/uncheck one, or **double-click the Check Med header** to check/uncheck all. **+ Add Medication** and **Edit this med** share the same one-dialog editor.

### 3 · Print Labels  (orange)

Every checked med auto-previews as a real label — drug + strength, directions, `form · qty · source · refills`, and Exp/Lot pinned to the bottom row (long patient names shrink to fit). Click **Print Checked Labels** — **2 copies of each** — enter your initials, and confirm. The confirmation lists exactly what's about to print and flags any row skipped for missing Exp/Lot.

![Label Previews gallery](support/assets/screenshots/print-gallery-v2.png)

### 4 · Log  (purple)

Every printed label is recorded automatically — timestamp, patient, DOB, medication details, **source**, initials, print count, and **Encounter #**. Each print is one **encounter**: numbered, shaded by encounter in **cycling light green / light blue / white**, and separated by a divider line so each patient reads as one block. Nothing to do here — it's your record. (The Log clears on close, so it always opens blank with no leftover highlighting.)

![Dispense Log](support/assets/screenshots/log-v2.png)

### 5 · Tebra Notes  (teal)

A ready-to-paste note for **each patient**, grouped by source, with each note **boxed** so you can select it, copy, and paste straight into that patient's Tebra chart. The notes are built **from the Log** and **rebuild every time you open the Tebra Notes tab** — so anything in the Log (even a row you typed in by hand) shows up here automatically.

![Tebra Notes](support/assets/screenshots/tebra-notes-v2.png)

> **Between patients:** **Start NEW Patient** clears the patient + med list but keeps the Log (run patient after patient). **Reset Session** wipes everything — Log included — for a fresh day.

---

## Key features

- **One paste → labels**, with a per-drug **confidence** score; Parse starts a fresh list each time so patients never mix.
- **Validation before printing** — live color-coded cells (yellow = missing → white when filled) and a clear **white → blue (reviewed) → green (checked)** row state.
- **Readable pop-ups** — Review, Print, and Remove all show each med's name + strength in bold with its status beneath.
- **Required Source** dropdown (DOH / IN HOUSE / RxAPS / Other), logged on every row and **printed on the label** (on the qty line, before Refills).
- **Checkbox-driven** selection that drives printing, row color, and removal alike.
- **2 copies per label**, batch or single, with detailed confirmations and fit-to-page so the bottom row never clips.
- **Complete dispense log** with numbered, green-banded **encounters** and a divider between each.
- **Tebra notes** sheet: paste-ready session notes grouped by source, one boxed note per patient.
- **Brother QL-1100c auto-select** with a load-the-roll confirmation.

---

## Updates happen on open

There's no separate "apply updates" step. Every time you open the tool with **`OPEN LABEL TOOL`**, it rebuilds the workbook from the latest `MedParser.bas` and opens it — so the newest code is always applied. After a `git pull`, just open the tool the usual way.

Requires, one-time in Excel's Trust Center:

- **Trust access to the VBA project object model**
- this folder registered as a **Trusted Location** (so macros aren't blocked)

<details>
<summary>Manual build, emblem updates, and event handlers</summary>

- **Manual build (VBA editor):** remove the `MedParser` module, **Import** `MedParser.bas`, run `SetupWorkbook`, save.
- **After changing the emblem crop:** run `tools/Build-ScuEmblem.ps1` first, then `Build-Release.vbs`.
- The two worksheet event handlers are pasted once into the sheet modules — see `HANDOFF.md` §5.
- Everything lives in **one folder** (outside OneDrive): `C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing`. Code + docs are version-controlled; the `.xlsm` and `support\_backups\` are git-ignored so patient data stays local.

</details>

---

## What's in the folder

<details>
<summary>Full contents (click to expand)</summary>

**Main folder**

| File | What it is |
|---|---|
| `OPEN LABEL TOOL (double-click me).cmd` | **The everyday launcher** — rebuilds from `MedParser.bas` and opens the tool. |
| `MedicationDispensing.xlsm` | The clinic workbook the launcher builds + opens (local, git-ignored). Fallback if a build won't run. |
| `Build-Release.vbs` | Imports `MedParser.bas`, runs `SetupWorkbook`, saves the `.xlsm`. |
| `MedParser.bas` | The VBA source of truth (parser, validation, label layout, printing, logging). |
| `scu_emblem.png` | SC emblem for the label header (loaded at print time; must stay here). |
| `README.md` | This file. |

**`support/docs/`** — `HANDOFF.md` (architecture / routine map), `CHANGELOG.md` (release history), `RELEASE.md`, `SETUP_INSTRUCTIONS.md`, `TROUBLESHOOTING.md`, `LABEL_REDESIGN.md`, `PRIVACY_AND_PHI.md`.

**`support/`** — `assets/` (emblem art + README screenshots), `test-data/` (no-PHI samples), `_backups/` (local snapshots, git-ignored).

**`tools/`** — `check-encoding.ps1` (pre-build ASCII/CRLF check), `make-release-zip.ps1`, `Build-ScuEmblem.ps1`, `BUILD_RELEASE_NOTES.md`.

</details>

---

## Quick reference

| I want to… | Go to |
|---|---|
| Install / update the tool | [RELEASE.md](support/docs/RELEASE.md) |
| Understand the code (routine map) | [HANDOFF.md](support/docs/HANDOFF.md) |
| See what changed | [CHANGELOG.md](support/docs/CHANGELOG.md) |
| Fix a problem | [TROUBLESHOOTING.md](support/docs/TROUBLESHOOTING.md) |
| Read the PHI / privacy rules | [PRIVACY_AND_PHI.md](support/docs/PRIVACY_AND_PHI.md) |

<details>
<summary>Known open items</summary>

- **Trusted Location** — register the folder in Excel Trust Center so macros aren't blocked.
- **Print width** — `LABEL_WIDTH_PT = 242` uses more of the 100 mm die-cut; re-test on the clinic Brother that labels still fit one die-cut (228 pt was the prior no-bleed value).
- **Bold clinic title on thermal** — may still look subtle; `ShrinkToFit` can limit apparent size.
- **Physical regression test** — spot-check the emblem and header on the Brother thermal (screen + print path verified).

</details>

---

## Privacy

**This repo holds code and docs only.** The working `.xlsm` contains patient information and is kept **local** — excluded via `.gitignore`, never committed (as is `support/_backups/`). All screenshots in this README use **randomly generated test patients**.
