# Pleadings Checker VBA -- Targeted Audit Report

**Date:** 2026-03-13 (pass 2)
**Scope:** All 20 modules in `Combined/`
**Approach:** Targeted fixes only; no broad rewrites

---

## Confirmed Defects Fixed

### 1. Application.StatusBar not restored in tracked-changes cleanup
**Module:** `PleadingsEngine.bas` `ApplySuggestionsAsTrackedChanges`
**Defect:** `doc.TrackRevisions` and `Application.ScreenUpdating` were captured/restored, but `Application.StatusBar` was not, leaving a stale status bar message after tracked-changes application.
**Fix (pass 1):** Added `Dim wasStatusBar As Variant` / `wasStatusBar = Application.StatusBar` on entry and `Application.StatusBar = wasStatusBar` in the single `TrackedCleanup:` path. Uses `Variant` because StatusBar can hold a `String` or `False`.

### 2. Stale installation note about Microsoft Scripting Runtime
**Module:** `PleadingsEngine.bas` header comment (was line 33)
**Defect:** The installation instructions told users to tick "Microsoft Scripting Runtime" in Tools > References, but the project uses late binding (`CreateObject("Scripting.Dictionary")`) exclusively. No early-bound reference is required.
**Fix (pass 2):** Removed the Scripting Runtime step from the numbered install instructions. Added a clarifying note: "No early-bound references are required."

### 3. Retired rules not unmistakably retired
**Modules:** `Rules_NumberFormats.bas` (Rule 18), `Rules_Terms.bas` (Rule 23)
**Defect:** Both retired rules had comments saying "RETIRED" but their public functions (`Check_PageRange`, `SetRange`, `Check_PhraseConsistency`) still looked like ordinary active entry points. No runtime signal if something accidentally called them.
**Fix (pass 1):** Updated engine header to `(Rules 5, 7; 23 RETIRED)` and `(Rules 9, 19; 18 RETIRED)`.
**Fix (pass 2):** Strengthened the function-level documentation with "NOT dispatched" / "Kept ONLY for backwards compatibility" language, and added `Debug.Print "WARNING: ..."` at the top of each retired function body so any accidental invocation is immediately visible in the Immediate window.

### 4. Silent fallback in 10 Engine* wrapper functions
**Modules:** `Rules_Spelling.bas`, `Rules_Quotes.bas`, `Rules_NumberFormats.bas`, `Rules_Terms.bas`, `Rules_Spacing.bas`
**Defect:** While `EngineIsInPageRange` and `EngineGetLocationString` wrappers (32 instances, 16 files) were fixed in pass 1 to log `Err.Number` + `Err.Description`, 10 other Engine wrappers fell back silently with just `Err.Clear` -- no `Debug.Print`:
  - `Rules_Spelling.EngineIsWhitelistedTerm`
  - `Rules_Spelling.EngineGetSpellingMode`
  - `Rules_Quotes.EngineGetQuoteNesting`
  - `Rules_Quotes.EngineGetSmartQuotePref`
  - `Rules_NumberFormats.EngineGetDateFormatPref`
  - `Rules_NumberFormats.EngineSetPageRange`
  - `Rules_Terms.EngineGetTermQuotePref`
  - `Rules_Terms.EngineGetTermFormatPref`
  - `Rules_Terms.EngineSetWhitelist`
  - `Rules_Spacing.EngineGetSpaceStylePref`
**Fix (pass 2):** Added `Debug.Print "<WrapperName>: fallback (Err " & Err.Number & ": " & Err.Description & ")"` to all 10 wrappers. Total wrapper fallback coverage is now **42/42** (verified by automated count: 42 Engine wrapper definitions, 42 fallback log lines).

### 5. On Error Resume Next tightened in safe helper functions
**Modules:** `Rules_Spelling.bas`, `Rules_Quotes.bas`, `Rules_TextScan.bas`
**Defect:** Three helper functions wrapped pure string/array operations under OERN unnecessarily:
  - `Rules_Spelling.IsException`: OERN over `For` loop doing `LCase`/`Trim` string comparisons
  - `Rules_Quotes.GetQListPrefixLen`: OERN continued over `Len`/`Left$`/`Mid$` after `ListFormat.ListString` call
  - `Rules_TextScan.GetSOListPrefixLen`: Same pattern as Rules_Quotes
**Fix (pass 1):** Removed or repositioned `On Error GoTo 0` immediately after each fragile Word OM call so pure string operations execute without error suppression.

---

## Areas Verified as Acceptable and Left Alone

### Brand-rules path delegation
Both `frmPleadingsChecker.GetBrandRulesPath()` and `PleadingsLauncher.GetBrandRulesPath()` already delegate to `Rules_Brands.GetDefaultBrandRulesPath` via `Application.Run` with a clearly marked fallback. Comments are accurate.

### Quote-family deduplication
`PleadingsEngine.RunAllPleadingsRules` (lines 755-788) deduplicates across the three quote rules (`quotation_mark_consistency`, `single_quotes_default`, `smart_quote_consistency`) by `RangeStart|RangeEnd` key. First finding per position wins; later duplicates are discarded. Verified intact.

### CreateIssueDict 8-key payload consistency
All 16 rule modules produce the standard 8-key dictionary: `RuleName`, `Location`, `Issue`, `Suggestion`, `RangeStart`, `RangeEnd`, `Severity`, `AutoFixSafe`. Each module's private `CreateIssueDict` function follows the same signature. No drift found.

### Rules_Lists engine wiring semantics
`Rules_Lists.bas` has an accurate ENGINE WIRING NOTE (lines 11-15) documenting the single aggregate toggle `"list_rules"` which dispatches both `Check_InlineListFormat` and `Check_ListPunctuation`. This matches the engine dispatch at lines 682-689.

### Brand persistence API
`Rules_Brands.SaveBrandRules` / `LoadBrandRules` both return `Boolean`. The form and launcher check the return value and show appropriate feedback on `False`.

### Debug logging infrastructure
`modDebugLog.bas` reviewed -- circular buffer, `DEBUG_MODE` toggle, `DebugLog`/`DebugLogError`/`TraceEnter`/`TraceExit`/`TraceFail` all present and functional. No changes needed.

---

## Remaining Limitations (Need Live Word Testing)

### Broad On Error Resume Next in paragraph-iteration loops
Several modules wrap entire paragraph-iteration loops (50-200+ lines) in `On Error Resume Next`. This is a defensive pattern against the Word Object Model, where any paragraph access can throw at any iteration. Tightening these would require restructuring each loop into multiple error-handler blocks, which risks introducing control-flow bugs without live testing.

**Most significant instances:**
- `Rules_Quotes.bas` lines 246-447 (~200 lines, nesting analysis)
- `Rules_Punctuation.bas` lines 766-934 (~170 lines, dash usage)
- `Rules_TextScan.bas` lines 49-167 and 204-341 (~250 lines total, repeated words + spell-out)
- `Rules_NumberFormats.bas` lines 819-934 (~115 lines, currency formatting)
- `Rules_Terms.bas` lines 495-559 (~64 lines, defined terms scan)

**Recommended approach:** Extract `.Range.Text` under OERN, restore `On Error GoTo 0`, then process the extracted string without error suppression. Should be done incrementally with regression testing.

### Find.Execute loop OERN patterns
Find loops in Rules_Terms, Rules_NumberFormats, and other modules keep OERN over the full iteration including `rng.Collapse wdCollapseEnd`, which is itself fragile on deleted/malformed ranges. Splitting these requires careful control-flow surgery that should be tested against real documents.

### No unit-test harness
VBA has no native test framework. All fixes were verified by code inspection only. Recommend manual regression testing in Word with a document that exercises all rule categories.

---

## Exact Modules Changed

| Module | Pass 1 | Pass 2 |
|--------|--------|--------|
| PleadingsEngine.bas | StatusBar capture/restore, retired-rule header annotations | Scripting Runtime install note corrected |
| Rules_Brands.bas | Err.Description in wrappers | -- |
| Rules_FootnoteHarts.bas | Err.Description in wrappers | -- |
| Rules_FootnoteIntegrity.bas | Err.Description in wrappers | -- |
| Rules_Formatting.bas | Err.Description in wrappers | -- |
| Rules_Headings.bas | Err.Description in wrappers | -- |
| Rules_Italics.bas | Err.Description in wrappers | -- |
| Rules_LegalTerms.bas | Err.Description in wrappers | -- |
| Rules_Lists.bas | Err.Description in wrappers | -- |
| Rules_NumberFormats.bas | Err.Description in wrappers | Retired-rule Debug.Print warnings, EngineSetPageRange + EngineGetDateFormatPref logging |
| Rules_Numbering.bas | Err.Description in wrappers | -- |
| Rules_Punctuation.bas | Err.Description in wrappers | -- |
| Rules_Quotes.bas | Err.Description in wrappers, OERN tightened in GetQListPrefixLen | EngineGetQuoteNesting + EngineGetSmartQuotePref logging |
| Rules_Spacing.bas | Err.Description in wrappers | EngineGetSpaceStylePref logging |
| Rules_Spelling.bas | Err.Description in wrappers, OERN removed from IsException | EngineIsWhitelistedTerm + EngineGetSpellingMode logging |
| Rules_Terms.bas | Err.Description in wrappers | Retired-rule Debug.Print warning, EngineGetTermQuotePref + EngineGetTermFormatPref + EngineSetWhitelist logging |
| Rules_TextScan.bas | Err.Description in wrappers, OERN tightened in GetSOListPrefixLen | -- |

**Unchanged modules:** frmPleadingsChecker.frm, PleadingsLauncher.bas, modDebugLog.bas

---

## Assumptions

1. The `Combined/` directory contains the canonical source for all modules.
2. `modDebugLog.bas` is correct as-is (reviewed, no issues found).
3. `Application.StatusBar` may hold a `String` or `False`; using `Variant` capture handles both.
4. The existing `Application.Run` dispatch architecture is intentional and correct.
5. Retired rules (23, 18) should remain in source for backwards compatibility but must not appear active.
6. All 42 Engine* wrappers should log on fallback for diagnosability.
