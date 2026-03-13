# Pleadings Checker VBA -- Targeted Audit Report

**Date:** 2026-03-13 (pass 3)
**Scope:** All 20 modules in `Combined/`
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

### 4. Silent fallback in 10 Engine* wrapper functions
**Modules:** `Rules_Spelling.bas`, `Rules_Quotes.bas`, `Rules_NumberFormats.bas`, `Rules_Terms.bas`, `Rules_Spacing.bas`
**Defect (pass 2):** 10 wrappers (beyond the 32 `EngineIsInPageRange`/`EngineGetLocationString` fixed in pass 1) fell back silently with just `Err.Clear`.
**Fix (pass 2):** Added `Debug.Print` with `Err.Number` + `Err.Description` to all 10. **Total: 42/42** Engine wrappers now log on fallback (verified by automated count).

### 5. Retired rules not unmistakably retired
**Modules:** `Rules_NumberFormats.bas` (Rule 18), `Rules_Terms.bas` (Rule 23)
**Defect:** Public functions `Check_PageRange`, `SetRange`, `Check_PhraseConsistency` looked like ordinary active entry points despite comments saying "RETIRED".
**Fix:** Engine header annotated `(Rules 5, 7; 23 RETIRED)` and `(Rules 9, 19; 18 RETIRED)`. Function-level comments strengthened to "NOT dispatched" / "Kept ONLY for backwards compatibility". Runtime `Debug.Print "WARNING: ..."` added to each retired function body.

### 6. On Error Resume Next tightened in safe helper functions
**Modules:** `Rules_Spelling.bas`, `Rules_Quotes.bas`, `Rules_TextScan.bas`
**Defect:** Three helper functions wrapped pure string/array operations under OERN:
  - `IsException`: OERN over `LCase`/`Trim` comparison loop
  - `GetQListPrefixLen`: OERN continued over `Len`/`Left$`/`Mid$` after `ListFormat.ListString`
  - `GetSOListPrefixLen`: Same pattern
**Fix (pass 1):** Removed or repositioned `On Error GoTo 0` immediately after each fragile Word OM call.

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
Broad OERN over paragraph loops and Find loops intentionally preserved. No safe narrowing opportunities remain in small helper functions.

---

## Remaining Limitations (Need Live Word Testing)

### Broad OERN in paragraph-iteration loops
Several modules wrap entire paragraph loops (50-200+ lines) in `On Error Resume Next`:
- `Rules_Quotes.bas` ~200 lines (nesting analysis)
- `Rules_Punctuation.bas` ~170 lines (dash usage)
- `Rules_TextScan.bas` ~250 lines (repeated words + spell-out)
- `Rules_NumberFormats.bas` ~115 lines (currency formatting)
- `Rules_Terms.bas` ~64 lines (defined terms scan)

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

**Unchanged modules this pass:** All 16 rule modules, frmPleadingsChecker.frm, PleadingsLauncher.bas, modDebugLog.bas
