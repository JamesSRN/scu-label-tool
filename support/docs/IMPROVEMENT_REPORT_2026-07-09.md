# SCU Label Printing — Improvement Report

_Prepared 2026-07-09 (evening) for morning review. No changes were made from this report — everything here is a proposal for you to accept, defer, or reject._

This covers three things you asked about:

1. **User-experience** improvements for the volunteers who run the tool.
2. **Optimization + robustness** — places we can make the code cleaner and safer *without* risking the reliability you depend on.
3. A **proposed folder structure** for the project, with exact steps and the caveats that make it non-trivial.

Everything is prioritized and tagged with rough **effort** (S/M/L) and **risk** (Low/Med/High) so you can cherry-pick. I've been deliberately conservative: this tool works and prints correctly today, so nothing here is worth destabilizing that.

---

## Part A — User experience

### A1. Name the skipped labels in the batch-complete dialog — S, Low  ⭐ top pick
Right now, a checked medication that's missing Exp or Lot is silently skipped and only surfaces as a **count** ("2 skipped"). A busy volunteer may not notice a label didn't come out. Change the end-of-batch message to list the skipped meds **by name**, e.g. "Skipped (missing Exp/Lot): Lisinopril 20 mg, Metformin 500 mg." Small, self-contained, high value — it closes the most likely "why didn't this print?" gap. (`PrintCheckedLabels` already builds a per-row list; reuse that.)

### A2. Protect the dispense Log from accidental loss — M, Med  ⭐ important
Today the workbook **wipes the Log (and saves) on close** for PHI hygiene. That's good for privacy, but it also means if a volunteer closes the file by accident mid-clinic, the day's dispensing record is gone with no recovery. Consider appending each print to a **dated local archive** (e.g. `dispense-log/2026-07-09.csv` in a git-ignored, local-only folder) at print time, so the on-close wipe clears the working sheet but a durable record survives. If a persistent dispensing record isn't required for the clinic, at minimum add a **confirmation on close** ("This clears today's log — continue?"). Worth a policy conversation before building.

### A3. "Reprint last batch" action — M, Low
Paper jams, mis-feeds, and "it printed crooked" are the most common real-world redo. A one-click reprint of the previous batch (or of the last single label) saves re-checking rows under time pressure. Store the last printed row set in a module variable or a hidden sheet.

### A4. Flag out-of-format expirations on the sheet — S, Low
The Exp check is intentionally *warn-but-allow* (correct call — never block care). But a fat-fingered "05/20227" can still reach a printed label. Add a light **visual flag** (e.g. amber cell) on the Medications tab for any Exp that fails the `MM/YYYY` check, so a second person can catch it during review without blocking anyone.

### A5. Volunteer quick-start card — S, Low
High turnover is your stated constraint. A one-page, laminated **"how to print a label in 4 steps"** (paste → review → check → Print Checked Labels) taped by the workstation removes most first-timer friction. I can generate this as a printable PDF from the existing workflow.

### A6. Keyboard-first flow in the Exp/Lot popup — S, Low
Volunteers enter a lot of these. Confirm that **Tab** moves Exp → Lot → OK and **Enter** submits (the OK button is already `Default`). Setting the initial focus to the Exp box on open shaves a click off every single entry.

### A7. Undo / confirm for "Remove this med" — S, Low
An accidental remove on the gallery is unrecoverable mid-session. A quick confirm ("Remove Lisinopril 20 mg?") or a soft "removed — click to undo" prevents re-typing.

### A8. Bigger, clearer primary buttons — S, Low
For a shared clinic PC used by many hands, oversize the two buttons that matter (**Print Checked Labels**, **Start New Patient**) and keep destructive actions (Remove, Clear) visually secondary. Pure styling, no logic risk.

### A9. Physical multi-label alignment test — S, Low  (already on your list as #12)
The last unverified link between screen and paper. Print a full batch of 3–4 varied labels (short name, very long name that now wraps, long lot number that now shrinks) on the real Brother and eyeball the margins. This is the single most valuable 10 minutes you can spend before wider rollout.

---

## Part B — Optimization & robustness

Guiding principle for this section: **the tool prints correctly today.** Every item below is chosen to *reduce* the chance of a future break, not to chase elegance. I've flagged anything with real risk as "defer unless motivated."

### B1. A pre-commit ASCII/CRLF validator — S, Low  ⭐ top pick
The single biggest fragility in this project is that `.bas`/`.vbs` **must** be pure ASCII + CRLF or the VBA importer mojibakes or rejects them. That guarantee currently lives in your head and in careful editing. Add a tiny checker script (PowerShell or Python) that scans `MedParser.bas` and `Build-Release.vbs` for any non-ASCII byte, a UTF-8 BOM, doubled `\r\r`, and unbalanced `Sub/Function/If/With` — and refuse to build if it fails. Wire it into the top of `Build-Release.vbs` (or a git pre-commit hook). This turns a whole class of "why won't it import" incidents into an instant, obvious message. Highest robustness-per-line-of-effort in the whole report.

### B2. A build/version stamp — S, Low
Write a version + build-date string into a known cell (or the workbook title bar) during `SetupWorkbook`. When a volunteer says "it's doing X," you can instantly tell **which build** they're on instead of guessing. Also helps confirm a rebuild actually took.

### B3. Structure self-check on open — M, Low
Add a defensive routine that runs on open and verifies the expected sheets and header columns exist (the `C_NAME`, `C_EXP`, … columns). If a volunteer deletes or reorders a column, today the failure is a confusing runtime error mid-print; a self-check can say "the Medications tab is missing the Exp column — restore from template" up front. Pairs well with keeping a tracked template (see C4).

### B4. Centralize the remaining label geometry constants — S, Low
Most geometry is already in named constants (good). A few row-height and position literals are still inline in `BuildLabelPreviewLayout` / `UpdateLabelPreviewForMedRow` / the gallery. Pulling them into the constants block makes future "nudge the layout" requests one-line changes and keeps the print label and gallery from drifting apart. Low risk, pure refactor — do it opportunistically, not as a big-bang.

### B5. A "debug mode" toggle that surfaces errors — S, Low
The code leans on `On Error Resume Next` for resilience (correct for production — a volunteer should never see a raw VBA error). The downside is it can also hide real bugs during development. A single module flag (`DEBUG_MODE = True/False`) that, when on, stops swallowing errors would make your own testing faster without changing production behavior.

### B6. Splitting the 156 KB `MedParser.bas` — L, Med  (defer unless motivated)
It's ~3,700 lines in one module (parsing, label layout, printing, forms, session, logging). Splitting into a few logical modules would help maintainability. **But**: it touches the build pipeline (multiple imports), the single-file simplicity is part of why the tool is easy to rebuild, and there's no functional payoff. Only worth it if the file becomes genuinely hard to navigate. If you do it, do it once, carefully, with a full print regression test after.

### B7. Faster printer lookup — S, Low
`SelectBrotherPrinter()` uses a WMI query that's the main source of the "please wait" delay you just added the progress bar for. Caching the resolved printer name for the session (re-query only if printing fails) would make the 2nd…Nth batch feel instant. Keep the progress bar for the first lookup. Low risk since a failed cached name simply falls back to a fresh query.

### B8. Keep the late-bound forms pattern — (no action, just affirming)
Referencing the UserForms late-bound (`VBA.UserForms.Add`, `f As Object`) with InputBox fallbacks is exactly right for resilience: `MedParser` compiles even if a form is missing, and prompting still works. Don't "clean this up" into early binding — the current design is the robust one.

---

## Part C — Proposed folder structure

### C1. What's cluttered today
- **Root** holds 28 tracked files mixing code, three logo PNGs, seven markdown docs, and the bootstrap workbook.
- **`_backups/` is ~10 MB** and contains not just the useful timestamped `.bas`/`.vbs` snapshots but a **full nested clone of the entire project — including its own `.git` and its own `_backups`.** That's redundant with your real git history and is the bulk of the bloat. (It's git-ignored, so it's a local-disk problem, not a GitHub one — but it's confusing and a second `.git` can trip up tooling.)
- **`tools/` tracks throwaway artifacts**: `_crop_above100.png`, `_crop_above149.png`, `_embedded_emblem.png`, `_git_orig_emblem.png`, `_git_scu_emblem.bin/.png`, `LogoB64.generated.txt`. These are debug/temp leftovers that shouldn't be in version control.
- **Three logo PNGs** at root: the 464 KB source, the 288 KB intermediate crop, and the 64 KB active `scu_emblem.png` — no separation of "source" vs "the one the app uses."
- **Docs are scattered** across root, `docs/`, and `tools/`.
- A stale **`~$MedicationDispensing.xlsm`** Excel lock file is sitting in root.

### C2. The important caveat before you move anything
`Build-Release.vbs` resolves everything relative to **its own folder** (`scriptDir`): it expects `MedParser.bas`, `MedicationDispensing.xlsm`, and the bootstrap workbook to sit **beside the script**. And the label logo is loaded by `LogoFilePath()` as **`scu_emblem.png` next to the workbook**. So a folder move is *not* free — if you relocate the build-critical files, you must update those paths in the same commit or the build and the logo silently break.

Because of that, I recommend a **light-touch reorg** (moves only the safe things) as the default, with a fuller layout as an optional follow-up that includes the matching code changes.

### C3. Recommended light-touch layout (low risk)
Keep the four build-critical files where the script expects them (root), and tidy everything else:

```
SCU Label Printing/
├─ README.md                     ← entry point (what it is, quick start, links)
├─ CHANGELOG.md
├─ MedParser.bas                 ← build-critical: leave at root
├─ Build-Release.vbs             ← build-critical: leave at root
├─ scu_emblem.png                ← build-critical: must stay beside the workbook
├─ MedicationDispensing.xlsm     ← live/PHI, git-ignored (leave at root)
├─ .gitignore  .gitattributes
│
├─ docs/                         ← ALL documentation lives here
│   ├─ HANDOFF.md
│   ├─ SETUP_INSTRUCTIONS.md
│   ├─ BUILD_RELEASE_NOTES.md    (move from tools/)
│   ├─ LABEL_REDESIGN.md
│   ├─ PRIVACY_AND_PHI.md
│   ├─ TROUBLESHOOTING.md
│   ├─ BRANCHING_AND_RELEASES.md
│   ├─ IMPROVEMENT_REPORT_2026-07-09.md   (this file)
│   └─ archive/                  ← one-off/historical notes
│       ├─ COMMIT_COMMANDS.md
│       └─ GITHUB_ISSUE_RESPONSES.md
│
├─ assets/
│   └─ logo-source/              ← the big source art, kept out of the way
│       ├─ Black SCU Logo + Transparent Background.png
│       └─ cropped_...Copy.png
│
├─ tools/
│   ├─ Build-ScuEmblem.ps1       (currently BROKEN — fix or remove; see note)
│   ├─ Run-BuildRelease.ps1
│   └─ check-encoding.ps1        ← new ASCII/CRLF validator from B1
│
├─ test-data/
│   └─ sample_tebra_pastes_no_phi.txt
│
└─ templates/
    └─ Template_MedicationDispensing.xlsm   ← renamed bootstrap (see C4)
```

### C4. Track a clean template (robustness gap worth closing)
The bootstrap `Broken_PrettyPrint_MedicationDispensing.xlsm` is git-ignored (via `*.xlsm`) and therefore **not on GitHub** — a fresh clone can't rebuild without someone hand-carrying that file. Since a *clean, PHI-free* template contains no patient data, it's safe to version it. Recommend: rename it to `templates/Template_MedicationDispensing.xlsm`, confirm it has zero PHI, and add a `.gitignore` exception:

```
*.xlsm
!templates/Template_MedicationDispensing.xlsm
```

Also rename it away from "Broken_PrettyPrint" — that name reads like a warning and confuses newcomers.

### C5. Cleanup to do (all safe — these are ignored/untracked or pure junk)
- **Delete the nested clone** `_backups/GIT_VERSION_SCU Label Printing/` — it's a full second copy of the repo (with its own `.git`) and is the bulk of the 10 MB. Your real git history already preserves everything.
- **Delete** the stale lock `~$MedicationDispensing.xlsm`.
- **Untrack the temp art** in `tools/` (`git rm --cached tools/_crop_above*.png tools/_embedded_emblem.png tools/_git_* tools/LogoB64.generated.txt`) and add `tools/_*` to `.gitignore`.
- Consider capping `_backups/` — the flat timestamped `.bas`/`.vbs` snapshots are useful, but you have ~40; keeping the last ~10 is plenty since git holds the rest.

### C6. Suggested order (do it as one reviewed commit)
1. Make the safe deletions in C5 first (instant space + clarity).
2. Move docs and logo-source art (C3). These aren't referenced by the build, so nothing breaks.
3. Rename + track the template (C4), update `.gitignore`.
4. **Only if you want the fuller `src/` layout:** move `MedParser.bas` + `Build-Release.vbs` into `src/` **and** update `scriptDir`/path resolution in `Build-Release.vbs` (and confirm `scu_emblem.png` still resolves beside the workbook) in the *same* commit. Test a full build + print before pushing.

Do the reorg through **GitHub Desktop** (or `git mv` so history follows the files), not by dragging in Explorer, and keep it a separate commit from any code change so it's easy to review and revert.

---

## Suggested first session (≈ half a day, all low-risk)
1. **B1** encoding validator — removes your biggest latent fragility.
2. **A1** named skipped labels — best UX win for the least code.
3. **A9** physical alignment test — verify the recent label changes on paper.
4. **C5** folder cleanup — delete the nested clone, stale lock, and tracked temp files.
5. **A2** decide the dispense-log retention policy — this one needs your judgment more than my code.

Everything else can wait for when you're motivated. Nothing in Part B is required for the tool to keep working as-is.
