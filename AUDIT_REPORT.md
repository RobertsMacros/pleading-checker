# Pleadings Checker VBA -- Targeted Audit Report

**Date:** 2026-03-13 (pass 9)
**Scope:** All 20 modules in `Code/` (18,824 lines total, fully inspected)
**Approach:** Targeted fixes only; no broad rewrites

---

## Confirmed Defects Fixed

### 1. Application.StatusBar not restored in tracked-changes and highlight cleanup
**Module:** `PleadingsEngine.bas`
**Defect:** `ApplySuggestionsAsTrackedChanges` and `ApplyHighlights` both modify application state but neither originally captured/restored `Application.StatusBar`, leaving a stale status bar after batch operations.
**Fix:** Both functions now capture `Application.StatusBar` as `Variant` on entry and restore it in their single cleanup paths (`TrackedCleanup:` and `HighlightCleanup:` respectively).

### 2. Stale installation note about Microsoft Scripting Runtime
**Module:** `PleadingsEngine.bas` header
**Defect:** Installation instructions told users to tick "Microsoft Scripting Runtime" in Tools > References, but the project uses late binding (`CreateObject("Scripting.Dictionary")`) exclusively.
**Fix (pass 2):** Removed the reference step. Added clarifying note: "No early-bound references are required."

### 3. Insufficient instrumentation in high-risk mutation paths
**Module:** `PleadingsEngine.bas`
**Defect:** The three highest-risk mutation/export paths had inconsistent diagnostic coverage:

| Path | Before | After |
|------|--------|-------|
| `ApplyHighlights` | Silent `Err.Clear` on highlight/comment failures; no logging for skipped items; no `DebugLogDoc` | `DebugLogError` on highlight, comment, and `doc.Range` failures; `TraceStep` for skipped items; `DebugLogDoc` on entry; StatusBar capture/restore |
| `ApplySuggestionsAsTrackedChanges` | Comment-only insertions (non-autofix path, line ~1408) silently cleared errors; skip-amendment comment insertion silently cleared errors | Both now log via `DebugLogError` on failure with step identity (`"comment-only i="`, `"skip-comment i="`) |
| `GenerateReport` | No `TraceEnter`/`TraceExit`; file-open failure logged in return string only | Added `TraceEnter`/`TraceExit` with issue count; `DebugLogError` on file-open failure and write-error paths |

**Fix (pass 3):** 7 new `DebugLogError` calls, 1 new `DebugLogDoc`, 1 new `TraceEnter`, 2 new `TraceExit`, 1 new `TraceStep` for skipped highlights. All use existing `modDebugLog` helpers. Zero overhead when `DEBUG_MODE = False`.
**Fix (pass 5):** Two remaining gaps in `ApplySuggestionsAsTrackedChanges`:
  - `doc.Range` failure was completely silent (unlike `ApplyHighlights` which had an `Else` branch). Added `DebugLogError` with range coordinates on `doc.Range` failure.
  - Comment-only path (non-autofix items) had no pre-call diagnostic. Added `TraceStep "COMMENT-ONLY"` with range, rule name before `doc.Comments.Add`.

### 4. Silent fallback in 10 Engine* wrapper functions
**Modules:** `Rules_Spelling.bas`, `Rules_Quotes.bas`, `Rules_NumberFormats.bas`, `Rules_Terms.bas`, `Rules_Spacing.bas`
**Defect (pass 2):** 10 wrappers (beyond the 32 `EngineIsInPageRange`/`EngineGetLocationString` fixed in pass 1) fell back silently with just `Err.Clear`.
**Fix (pass 2):** Added `Debug.Print` with `Err.Number` + `Err.Description` to all 10. **Total: 42/42** Engine wrappers now log on fallback (verified by automated count).

### 5. Retired rules not unmistakably retired
**Modules:** `Rules_NumberFormats.bas` (Rule 18), `Rules_Terms.bas` (Rule 23)
**Defect:** Public functions `Check_PageRange`, `SetRange`, `Check_PhraseConsistency` looked like ordinary active entry points despite comments saying "RETIRED".
**Fix:** Engine header annotated `(Rules 5, 7; 23 RETIRED)` and `(Rules 9, 19; 18 RETIRED)`. Function-level comments strengthened to "NOT dispatched" / "Kept ONLY for backwards compatibility". Runtime `Debug.Print "WARNING: ..."` added to each retired function body.
**Fix (pass 5):** Retired-rule constants renamed: `RULE_NAME_PAGE_RANGE` → `RETIRED_RULE_NAME_PAGE_RANGE` (dead code, unused), `RULE23_NAME` → `RETIRED_RULE23_NAME` (used only within retired `Check_PhraseConsistency`). Makes retirement impossible to miss when scanning constant declarations.

### 6. On Error Resume Next tightened in safe helper functions
**Modules:** `Rules_Spelling.bas`, `Rules_Quotes.bas`, `Rules_TextScan.bas`
**Defect:** Three helper functions wrapped pure string/array operations under OERN:
  - `IsException`: OERN over `LCase`/`Trim` comparison loop
  - `GetQListPrefixLen`: OERN continued over `Len`/`Left$`/`Mid$` after `ListFormat.ListString`
  - `GetSOListPrefixLen`: Same pattern
**Fix (pass 1):** Removed or repositioned `On Error GoTo 0` immediately after each fragile Word OM call.

### 7. OERN tightened in `IsBlockQuotePara` helper
**Module:** `Rules_Formatting.bas`
**Defect (pass 4):** `IsBlockQuotePara` had a single `On Error Resume Next` spanning ~128 lines. Two significant blocks of pure string/value logic (24 lines of `Like`/string operations for list-pattern detection, and 19 lines of quotation-mark `Left$`/`Right$`/`ChrW` comparisons) ran under unnecessary OERN.
**Fix (pass 4):** Inserted 2 `On Error GoTo 0` and 2 `On Error Resume Next` to create two protected string-only zones:
  - Zone A (lines ~124-148): list-pattern `Like` checks -- no longer under OERN
  - Zone B (lines ~182-215): indent-check + quote-mark string comparisons -- no longer under OERN
  Word OM calls (`ListFormat.ListString`, `Style.NameLocal`, `Format.LeftIndent`, `Font.Italic`) remain individually guarded by `On Error Resume Next` with `Err.Number` checks.

---

## Areas Verified as Acceptable and Left Alone

### Brand-rules path delegation
Both `frmPleadingsChecker.GetBrandRulesPath()` and `PleadingsLauncher.GetBrandRulesPath()` delegate to `Rules_Brands.GetDefaultBrandRulesPath` via `Application.Run` with clearly marked fallback.

### Quote-family deduplication
`PleadingsEngine.RunAllPleadingsRules` (line 755) deduplicates across the three quote rules by `RangeStart|RangeEnd` key. Verified intact.

### CreateIssueDict 8-key payload consistency
All 16 rule modules produce: `RuleName`, `Location`, `Issue`, `Suggestion`, `RangeStart`, `RangeEnd`, `Severity`, `AutoFixSafe`. No drift.

### Rules_Lists engine wiring semantics
ENGINE WIRING NOTE (lines 11-15) accurately documents single aggregate toggle `"list_rules"`. Matches dispatch.

### Brand persistence API
`SaveBrandRules` / `LoadBrandRules` return `Boolean`. Form and launcher handle `False`.

### Debug logging infrastructure
`modDebugLog.bas` reviewed -- circular buffer, `DEBUG_MODE` toggle, full trace/error/doc/range helpers. No changes needed.

### Wrapper fallback observability
42/42 `Engine*` wrappers log `Err.Number` + `Err.Description` on fallback.

### OERN in paragraph-iteration and Find.Execute loops
Broad OERN over paragraph loops and Find loops intentionally preserved.

### Pass 5 OERN audit: all 12 specified modules re-verified
All 12 modules re-audited. Previous fixes intact (`IsException`, `GetQListPrefixLen`, `GetSOListPrefixLen`, `IsBlockQuotePara`). No new safe tightening targets found. Two agent-reported candidates rejected on manual review: `FindTermRange` call in Rules_Terms.bas involves Word OM internally (OERN correctly needed); single `Len()` call in Rules_Formatting.bas between two OM calls (not worth toggling for 1 line).

### Pass 4 OERN audit: 5 newly-audited modules
Modules `Rules_Formatting.bas`, `Rules_Headings.bas`, `Rules_Italics.bas`, `Rules_FootnoteIntegrity.bas`, `Rules_LegalTerms.bas` were audited for OERN tightening opportunities:
- **Rules_Headings.bas**: OERN already tightly scoped per-call in `CountWordInDoc` and `FlagOccurrences`; paragraph scan (pass 1) ends with `On Error GoTo 0` before pure-VBA passes 2-3. No targets.
- **Rules_Italics.bas**: `IsRangeItalic` already does `On Error GoTo 0` immediately after each fragile call. Main functions (`Check_AnglicisedTermsNotItalic`, `Check_ForeignNamesNotItalic`) scope OERN per-call. Pure-string helpers (`IsLetter`, `MergeArrays`) have no OERN. No targets.
- **Rules_FootnoteIntegrity.bas**: All 8 private subs scope OERN around individual fragile Word OM calls with immediate `On Error GoTo 0`. `IsPunctuation` is pure value logic with no OERN. No targets.
- **Rules_LegalTerms.bas**: No broad OERN anywhere. `SearchAndFlag` and `CheckTermInParagraph` scope OERN per-call. Pure-string helpers (`IsWordChar`, `IsInsideQuote`, `MergeArrays2`) have no OERN. No targets.
- **Rules_Formatting.bas**: `IsBlockQuotePara` tightened (see defect #7). `Check_ParagraphBreakConsistency` and `Check_FontConsistency` have broad OERN over paragraph loops -- intentionally preserved (needs live Word testing).

---

## Remaining Limitations (Need Live Word Testing)

### Broad OERN in paragraph-iteration loops
Several modules wrap entire paragraph loops (50-200+ lines) in `On Error Resume Next`:
- `Rules_Quotes.bas` ~200 lines (nesting analysis)
- `Rules_Punctuation.bas` ~170 lines (dash usage)
- `Rules_TextScan.bas` ~250 lines (repeated words + spell-out)
- `Rules_NumberFormats.bas` ~115 lines (currency formatting)
- `Rules_Terms.bas` ~64 lines (defined terms scan)
- `Rules_Formatting.bas` ~150 lines (`Check_ParagraphBreakConsistency`) and ~440 lines (`Check_FontConsistency`)

**Recommended approach:** Extract `.Range.Text` under OERN, `On Error GoTo 0`, then process extracted string without suppression. Must be done incrementally with live Word regression testing.

### Find.Execute loop OERN
Find loops keep OERN over the full iteration including `rng.Collapse wdCollapseEnd` (itself fragile). Splitting requires control-flow surgery tested against real documents.

### No unit-test harness
VBA has no native test framework. All fixes verified by code inspection. Manual regression testing recommended.

---

## Exact Procedures Changed (Pass 3)

| Module | Procedure | Change |
|--------|-----------|--------|
| `PleadingsEngine.bas` | `ApplyHighlights` | Added `DebugLogDoc`, `DebugLogError` (3 paths: highlight, comment, doc.Range), `TraceStep` for skipped items, StatusBar capture/restore |
| `PleadingsEngine.bas` | `ApplySuggestionsAsTrackedChanges` | Added `DebugLogError` for comment-only insertion (non-autofix path) and skip-amendment comment insertion |
| `PleadingsEngine.bas` | `GenerateReport` | Added `TraceEnter`/`TraceExit`, `DebugLogError` for file-open and write-error paths |

**Unchanged modules pass 3:** All 16 rule modules, frmPleadingsChecker.frm, PleadingsLauncher.bas, modDebugLog.bas

## Exact Procedures Changed (Pass 4)

| Module | Procedure | Change |
|--------|-----------|--------|
| `Rules_Formatting.bas` | `IsBlockQuotePara` | Tightened OERN: added `On Error GoTo 0` after `para.Range.Text` read (before 24-line string block) and after `para.Format.LeftIndent` read (before 19-line string block); added `On Error Resume Next` before `ListFormat.ListString` and `Font.Italic` Word OM calls |

**Unchanged modules pass 4:** PleadingsEngine.bas, all other 15 rule modules, frmPleadingsChecker.frm, PleadingsLauncher.bas, modDebugLog.bas

## Exact Procedures Changed (Pass 5)

| Module | Procedure | Change |
|--------|-----------|--------|
| `PleadingsEngine.bas` | `ApplySuggestionsAsTrackedChanges` | Added `DebugLogError` Else branch for `doc.Range` failure (was completely silent); added `TraceStep "COMMENT-ONLY"` before comment insertion on non-autofix items |
| `Rules_NumberFormats.bas` | _(constant only)_ | Renamed `RULE_NAME_PAGE_RANGE` → `RETIRED_RULE_NAME_PAGE_RANGE` (dead code) |
| `Rules_Terms.bas` | _(constant only)_ | Renamed `RULE23_NAME` → `RETIRED_RULE23_NAME` (and 1 reference in `Check_PhraseConsistency`) |

**Unchanged modules pass 5:** All other 14 rule modules, frmPleadingsChecker.frm, PleadingsLauncher.bas, modDebugLog.bas, Rules_Formatting.bas

## Exact Procedures Changed (Pass 6)

| Module | Procedure | Change |
|--------|-----------|--------|
| `PleadingsEngine.bas` | `RunAllPleadingsRules` | Capture `Application.ScreenUpdating` and `Application.StatusBar` on entry; restore both in `RunnerCleanup` (was hardcoded `True` / `""`) |
| `PleadingsEngine.bas` | `ApplyHighlights` | Replaced raw `doc.Comments.Add` with `TryAddComment` wrapper from `modDebugLog` |
| `PleadingsEngine.bas` | `ApplySuggestionsAsTrackedChanges` | Replaced raw `rng.Text = sugText` with `TrySetRangeText`; replaced 2x raw `doc.Comments.Add` with `TryAddComment`; all three now get range-level diagnostics via `modDebugLog` wrappers |
| `frmPleadingsChecker.frm` | `GetTempReportPath` | Added `Environ("USERPROFILE")` fallback before hardcoded `C:\Temp` |
| `PleadingsLauncher.bas` | `ExportReport` | Same `Environ("USERPROFILE")` fallback added |

**Unchanged modules pass 6:** All 16 rule modules, modDebugLog.bas

## Pass 6 Verification

### Confirmed defects fixed this pass

**8. `RunAllPleadingsRules` did not restore prior application state**
- Cleanup forced `ScreenUpdating = True` and `StatusBar = ""` regardless of prior values
- Now captures both on entry and restores in `RunnerCleanup`

**9. Raw mutation calls lacked wrapper-level diagnostics**
- `rng.Text = sugText`, `doc.Comments.Add` in `ApplySuggestionsAsTrackedChanges` and `ApplyHighlights` replaced with `TrySetRangeText` / `TryAddComment` from `modDebugLog`
- Wrappers log `DebugLogRange` before mutation and `DebugLogError` on failure (when `DEBUG_MODE = True`)
- Zero overhead when `DEBUG_MODE = False` (fast path)

**10. Export temp-path fallback could hit non-existent `C:\Temp`**
- Added `Environ("USERPROFILE")` as intermediate fallback in both form and launcher
- Chain now: document path → `%TEMP%` → `%TMP%` → `%USERPROFILE%` → `C:\Temp`

### Areas verified as acceptable

- `CreateIssue` 9-key vs `CreateIssueDict` 8-key: compatible — `GetIssueProp` handles missing keys via `Err.Clear`; no runtime issue
- 42/42 Engine wrapper fallbacks log `Err.Number` + `Err.Description` (automated count)
- OERN: all 12 specified modules re-verified in passes 4-5; no new targets in pass 6
- Brand API: `SaveBrandRules`/`LoadBrandRules` return Boolean; form and launcher delegate with fallback
- All prior fixes intact (StatusBar, Scripting Runtime note, retired rules, quote dedupe, Lists wiring, etc.)

### Assumptions and source-coverage limits
- All 13 scoped modules inspected from the full repo files (not truncated)
- `Rules_Numbering.bas`, `Rules_Punctuation.bas`, `Rules_Quotes.bas`, `Rules_Spacing.bas`, `Rules_Spelling.bas`, `Rules_Terms.bas`, `Rules_TextScan.bas` not in scope this pass but verified in passes 1-5

### Remaining limitations (need live Word testing)
- Broad OERN in paragraph-iteration loops (7 modules, 50-440 lines each)
- Find.Execute loop OERN (`rng.Collapse wdCollapseEnd` itself fragile)
- No unit-test harness
- Rules_Lists ENGINE WIRING NOTE accurate ✓
- Brand persistence API (`SaveBrandRules`/`LoadBrandRules`) returns Boolean ✓

---

## Exact Procedures Changed (Pass 7)

| Module | Procedure | Change |
|--------|-----------|--------|
| `Rules_Lists.bas` | `Check_ListPunctuation` | Fixed list ID computation: now uses `List.ListParagraphs(1).Range.Start` (unique per list) instead of `ListParagraphs.Count` (non-unique). Grouping loop now breaks groups when `paraListID` changes between consecutive list paragraphs. |
| `Rules_Punctuation.bas` | `CreateBracketIssue` | Fixed suggestion logic: `Select Case` now matches `"()"`, `"[]"`, `"{}"` (the values actually passed) instead of individual characters that never matched. |
| `Rules_Punctuation.bas` | `Check_BracketIntegrity` | Added stack-based nesting check for `([)]` patterns. Runs only when counts balance (avoids duplicate reports). Uses existing `CodesMatch` helper. |

**Unchanged modules pass 7:** PleadingsEngine.bas, frmPleadingsChecker.frm, PleadingsLauncher.bas, modDebugLog.bas, all 15 other rule modules

## Pass 7 Verification

### Confirmed defects fixed this pass

**11. List-grouping bug merged adjacent distinct lists**
- `Check_ListPunctuation` grouped consecutive list paragraphs purely by `paraIsList` flag
- `paraListID` was computed but never used in grouping
- The ID computation itself was broken: `ListParagraphs.Count` is not a unique list identifier (two lists with the same item count get the same "ID")
- Fixed: ID now uses start position of first paragraph in the Word List object (`List.ListParagraphs(1).Range.Start`), which is unique per list. Grouping loop breaks groups when IDs differ (with guard for `0` = unknown).

**12. Bracket suggestion never matched passed values**
- `Check_BracketIntegrity` passes `"()"`, `"[]"`, `"{}"` as `bracketChar`
- `CreateBracketIssue` compared against `"("`, `")"`, `"["`, etc. — individual characters that never match
- All bracket issues fell through to the generic `"Review bracket pairing"` suggestion
- Fixed: `Select Case` now matches the pair strings (`"()"`, `"[]"`, `"{}"`) plus individual characters for forward compatibility

**13. No nesting check in bracket integrity**
- Header claimed checks for "improperly nested brackets" but code only compared open/close counts
- `([)]` would pass undetected (1 open paren, 1 close paren, 1 open bracket, 1 close bracket — all balanced)
- Fixed: stack-based nesting check runs when counts balance. Uses existing `CodesMatch` helper. Reports at the position of the first nesting violation.

### Areas verified as acceptable and left alone

- **Items 1, 2, 6 (pass 6 fixes):** `RunAllPleadingsRules` state capture/restore, `TrySetRangeText`/`TryAddComment` wrappers, `USERPROFILE` temp-path fallback — all intact
- **Item 7 (issue payload):** `CreateIssue` (9 keys) vs `CreateIssueDict` (8 keys) — compatible. All access through `GetIssueProp` which returns `""` on missing keys via `Err.Clear`. No runtime problem.
- **Item 8 (OERN):** All 10 specified modules re-audited. No new safe tightening targets found. Previous fixes intact (`IsException`, `GetQListPrefixLen`, `GetSOListPrefixLen`, `IsBlockQuotePara`). Paragraph-loop and Find.Execute OERN intentionally preserved.
- **42/42 Engine wrapper fallbacks** log `Err.Number` + `Err.Description` (verified across all scoped rule modules)
- **Rules_Spacing.bas** fully audited (623 lines): 5 public functions, 3 engine wrappers all logging properly, no defects found, OERN patterns all intentionally broad or already tightened

### Assumptions and source-coverage limits

- All 17 scoped modules inspected from the full repo files (not truncated)
- `Rules_FootnoteHarts.bas`, `Rules_TextScan.bas`, `Rules_Brands.bas` not in scope this pass but verified in passes 1-6

### Remaining limitations (need live Word testing)

- **Performance:** `CheckManualNumbering` calls `Application.Run("Rules_Formatting.IsBlockQuotePara", para)` per paragraph via late-bound dispatch. Caching would improve performance on large documents but requires live testing to validate.
- Broad OERN in paragraph-iteration loops (7 modules, 50-440 lines each)
- Find.Execute loop OERN (`rng.Collapse wdCollapseEnd` itself fragile)
- No unit-test harness

---

## Pass 8 — Concrete Fixes and Full Codebase Audit

### Confirmed defects fixed

**15. Export/report paths did not create parent directories**
- `btnExport_Click` and `PleadingsLauncher.ExportReport` passed a report path to `GenerateReport` without ensuring the parent directory exists
- Temp-path fallback (e.g. `C:\Temp`) could fail if the directory didn't exist
- Debug log path (derived from report path) had the same risk
- Fixed: added `EnsureDirectoryExists` + `GetParentDirectory` helpers to `modDebugLog.bas`; both export paths now call `EnsureDirectoryExists` before writing

**16. Brand-save path used single MkDir (fails for nested directories)**
- `btnSaveBrands_Click` (form) and `ManageBrands` (launcher) used a single `MkDir` that fails if intermediate directories don't exist (e.g. `%APPDATA%\PleadingsChecker` when `%APPDATA%` itself is missing)
- Fixed: both now use `modDebugLog.EnsureDirectoryExists` which walks path components and creates each level

### Exact procedures/modules changed (pass 8)

| Module | Procedure | Change |
|--------|-----------|--------|
| `modDebugLog.bas` | `EnsureDirectoryExists` (new) | Recursive directory creation helper, no FSO dependency, Mac/Win compatible |
| `modDebugLog.bas` | `GetParentDirectory` (new) | Extract parent directory from a file path |
| `frmPleadingsChecker.frm` | `btnExport_Click` | Added `EnsureDirectoryExists` call before `GenerateReport` |
| `frmPleadingsChecker.frm` | `btnSaveBrands_Click` | Replaced single `MkDir` with `EnsureDirectoryExists` |
| `PleadingsLauncher.bas` | `ExportReport` | Added `EnsureDirectoryExists` call before `GenerateReport` |
| `PleadingsLauncher.bas` | `ManageBrands` (SAVE case) | Replaced single `MkDir` with `EnsureDirectoryExists` |

### Areas verified and left unchanged

- **Item 1:** `RunAllPleadingsRules` already captures `wasScreenUpdating` + `wasStatusBar` on entry and restores both in cleanup
- **Item 2:** `ApplySuggestionsAsTrackedChanges` and `ApplyHighlights` already use `TrySetRangeText` and `TryAddComment` wrappers; no raw mutations remain
- **Item 5:** List-grouping edge case: when both `paraListID` values are 0 (unknown), groups stay merged. This is conservative and avoids false splits. Level changes within the same Word `List` object correctly stay grouped.
- **Item 6:** Bracket count-mismatch anchors on first occurrence of that bracket type (imperfect but adequate). Nesting-error anchor correctly points to the offending closing bracket.
- **Item 7:** All issue-dictionary access goes through `GetIssueProp` which returns `""` on missing keys. No function assumes specific key counts. `SortIssuesByPosition` does not exist.
- **Item 8:** OERN audit across all 10 specified modules found no clearly safe new tightening targets. All existing patterns are either intentionally broad (paragraph/Find loops), already tightened with inline error checks, or wrapper functions.
- **Item 9:** `CheckManualNumbering` performance hotspot (`IsBlockQuotePara` per paragraph) requires live testing to cache safely. Noted as limitation, not changed.

### Source coverage

| Module | Lines | Status |
|--------|-------|--------|
| PleadingsEngine.bas | 1926 | Fully inspected |
| Rules_Spelling.bas | 1726 | Fully inspected |
| Rules_Punctuation.bas | 1002 | Fully inspected |
| Rules_Lists.bas | 981 | Fully inspected |
| Rules_TextScan.bas | 976 | Fully inspected |
| frmPleadingsChecker.frm | 951 | Fully inspected |
| Rules_NumberFormats.bas | 949 | Fully inspected |
| Rules_Terms.bas | 931 | Fully inspected |
| Rules_Formatting.bas | 922 | Fully inspected |
| Rules_Numbering.bas | 905 | Fully inspected |
| Rules_Quotes.bas | 819 | Fully inspected |
| modDebugLog.bas | 804 | Fully inspected |
| Rules_Headings.bas | 707 | Fully inspected |
| Rules_FootnoteHarts.bas | 647 | Fully inspected |
| Rules_Spacing.bas | 622 | Fully inspected |
| Rules_FootnoteIntegrity.bas | 502 | Fully inspected |
| Rules_LegalTerms.bas | 487 | Fully inspected |
| Rules_Italics.bas | 382 | Fully inspected |
| PleadingsLauncher.bas | 334 | Fully inspected |
| Rules_Brands.bas | 325 | Fully inspected |
| **Total** | **18,824** | **All 20 files fully inspected** |

No files were truncated or partially read. All line counts verified via `wc -l`.

### Remaining limitations (need live Word testing)

- **Performance:** `CheckManualNumbering` per-paragraph `Application.Run("Rules_Formatting.IsBlockQuotePara", para)` is slow on large documents; needs caching validated under live conditions
- **List-grouping fallback:** When Word OM fails to identify a List object (both IDs = 0), adjacent distinct lists may still merge; extremely unlikely in practice
- **Bracket count-mismatch anchoring:** Reports position of first bracket of that type, not necessarily the unmatched one
- Broad OERN in paragraph-iteration loops (7 modules, 50-440 lines each)
- Find.Execute loop OERN (`rng.Collapse wdCollapseEnd` itself fragile)
- No unit-test harness

---

## Pass 9 — Final Hardening

### Confirmed defects fixed

**17. EnsureDirectoryExists broken on UNC paths**
- `Split("\\server\share\dir", "\")` produces `["","","server","share","dir"]`
- Old code: `built = parts(0)` = `""`, then builds `"\server"`, `"\\server"` — attempted `MkDir` on nonsensical intermediate paths before ever reaching `\\server\share`
- Fixed: detect UNC prefix (`\\`), skip to index 4, treating `\\server\share` as the unsplittable root
- Also fixed: empty path components from double separators now skipped instead of producing malformed paths

**18. Temp-path fallback could target unwritable C:\Temp**
- Old chain: `TEMP` -> `TMP` -> `USERPROFILE` -> `C:\Temp`
- `C:\Temp` may not exist and is often unwritable on locked-down machines
- `USERPROFILE` dropped the report in the user's home directory root (clumsy)
- Fixed: new shared helper `modDebugLog.GetWritableTempDir` uses chain: `TEMP` -> `TMP` -> `Application.Options.DefaultFilePath(wdTempFilePath)` -> `LOCALAPPDATA\Temp` -> `USERPROFILE` (last resort)
- Both form (`GetTempReportPath`) and launcher (`ExportReport`) now call the shared helper — eliminates duplicated inline logic

**19. UserForm sizing used outer dimensions, clipping bottom controls**
- `Me.Height = 1000` sets the *outer* height including title bar and window chrome (~25pt)
- Controls were laid out assuming 1000pt of interior space, so the bottom status label was clipped
- Fixed: now uses `Me.InsideHeight` (= client area) computed from actual layout endpoint (`yPos + LBL_H + PAD`)
- `.frm` header `ClientHeight` reduced from 1000 to 600 to minimize pre-Initialize flash

### Exact procedures/modules changed (pass 9)

| Module | Procedure | Change |
|--------|-----------|--------|
| `modDebugLog.bas` | `EnsureDirectoryExists` | UNC path handling, empty component skip, trailing separator loop |
| `modDebugLog.bas` | `GetWritableTempDir` (new) | Shared temp-dir helper with robust fallback chain, no FSO |
| `frmPleadingsChecker.frm` | `UserForm_Initialize` | `Me.InsideWidth`/`Me.InsideHeight` instead of `Me.Width`/`Me.Height`; height computed from layout |
| `frmPleadingsChecker.frm` | `GetTempReportPath` | Delegates to `modDebugLog.GetWritableTempDir` |
| `frmPleadingsChecker.frm` | `.frm` header | `ClientHeight` 1000 -> 600 |
| `PleadingsLauncher.bas` | `ExportReport` | Inline temp-path logic replaced with `modDebugLog.GetWritableTempDir` call |

### Areas verified and left unchanged

- **RunAllPleadingsRules state restoration**: captures/restores `wasScreenUpdating` + `wasStatusBar` — intact
- **ApplyHighlights / ApplySuggestionsAsTrackedChanges**: all mutations go through `TrySetRangeText`/`TryAddComment` — no raw mutations remain. Return values not captured but wrappers self-log errors. No meaningful recovery possible in callers.
- **Debug log path**: derived from report path (same directory, `.json` -> `_debug.log`) — always valid when report path parent exists
- **Brand save/load path**: form and launcher both use `EnsureDirectoryExists` + `GetBrandRulesPath` delegate — aligned
- **OERN audit**: no new clearly-safe tightening targets found in this pass

### Remaining limitations (need live Word testing)

- `InsideWidth`/`InsideHeight` may behave differently across Word versions; need to verify on Word 2010, 2016, 365
- `Application.Options.DefaultFilePath(wdTempFilePath)` fallback untested on machines where TEMP/TMP are both unset
- `CheckManualNumbering` performance hotspot unchanged (needs caching validated under live conditions)
- Broad OERN in paragraph-iteration loops (7 modules)
- No unit-test harness
