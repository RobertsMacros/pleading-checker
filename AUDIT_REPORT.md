# Pleadings Checker VBA -- Targeted Audit Report

**Date:** 2026-03-13
**Scope:** All 20 modules in `Combined/`
**Approach:** Targeted fixes only; no broad rewrites

---

## Fixes Applied

### Fix 1 -- Brand-rules path delegation (verified, no change needed)
Both `frmPleadingsChecker.GetBrandRulesPath()` (line 924) and `PleadingsLauncher.GetBrandRulesPath()` (line 313) already delegate to `Rules_Brands.GetDefaultBrandRulesPath` via `Application.Run` with a clearly marked fallback. Comments are accurate. No drift found.

### Fix 2 -- Restore fallback observability (32 edits across 16 files)
All `EngineIsInPageRange` and `EngineGetLocationString` wrappers across every rule module logged only `Err.Number` in their fallback branch. Added `Err.Description` to all 32 `Debug.Print` statements.

**Before:** `Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ")"`
**After:** `Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"`

Files changed: Rules_Brands, Rules_FootnoteHarts, Rules_FootnoteIntegrity, Rules_Formatting, Rules_Headings, Rules_Italics, Rules_LegalTerms, Rules_Lists, Rules_NumberFormats, Rules_Numbering, Rules_Punctuation, Rules_Quotes, Rules_Spacing, Rules_Spelling, Rules_Terms, Rules_TextScan.

### Fix 3 -- Tracked-changes state handling (PleadingsEngine.bas)
`ApplySuggestionsAsTrackedChanges` already captured and restored `doc.TrackRevisions` and `Application.ScreenUpdating`. Added `Application.StatusBar` capture on entry and restore in the cleanup block, so all three state variables are now saved/restored in a single cleanup path.

### Fix 4 -- Retired rules honestly marked (PleadingsEngine.bas header)
Updated engine header comments to clearly indicate retired rules:
- `Rules_Terms.bas (Rules 5, 7; 23 RETIRED)` (was `Rules 5, 7, 23`)
- `Rules_NumberFormats.bas (Rules 9, 19; 18 RETIRED)` (was `Rules 9, 18, 19`)

The rule modules themselves (`Rules_Terms.RULE23_NAME` and `Rules_NumberFormats.RULE_NAME_PAGE_RANGE`) were already properly marked as retired with clear comments and `Check_PhraseConsistency` / `Check_PageRange` annotated as not dispatched. No engine dispatch wiring exists for either.

### Fix 5 -- Tighten broad On Error Resume Next (3 targeted fixes)
- **Rules_Spelling.bas `IsException`**: Removed unnecessary `On Error Resume Next` around a pure string-comparison loop (array iteration of `LCase`/`Trim` comparisons cannot fail).
- **Rules_Quotes.bas `GetQListPrefixLen`**: Moved `On Error GoTo 0` immediately after the fragile `ListFormat.ListString` call. The subsequent `Len`/`Left$`/`Mid$` string operations no longer run under error suppression.
- **Rules_TextScan.bas `GetSOListPrefixLen`**: Same pattern fix as Rules_Quotes.

### Fix 6 -- Quote-family deduplication (verified, no change needed)
`PleadingsEngine.RunAllPleadingsRules` (lines 753-786) already deduplicates across the three quote rules (`quotation_mark_consistency`, `single_quotes_default`, `smart_quote_consistency`) by `RangeStart|RangeEnd` key. First finding per position wins; later duplicates are discarded.

### Fix 7 -- Issue payload consistency (verified, no change needed)
All modules produce the standard 8-key dictionary: `RuleName`, `Location`, `Issue`, `Suggestion`, `RangeStart`, `RangeEnd`, `Severity`, `AutoFixSafe`. Each module's `CreateIssueDict` function follows the same signature. No missing or extra keys found.

### Fix 8 -- Rules_Lists comments match engine wiring (verified, no change needed)
`Rules_Lists.bas` has an accurate ENGINE WIRING NOTE (lines 12-15) documenting that the engine uses a single aggregate toggle `"list_rules"` which dispatches both `Check_InlineListFormat` and `Check_ListPunctuation`. This matches the engine dispatch at lines 682-689.

### Fix 9 -- Stale comments (fixed via Fix 4)
The engine header was the only stale comment found -- it listed retired rules without annotation. Fixed in Fix 4. All other module-level comments verified as accurate.

### Fix 10 -- Brand persistence API (verified, no change needed)
`Rules_Brands.SaveBrandRules` and `LoadBrandRules` both return `Boolean`. `frmPleadingsChecker.btnSaveBrands_Click` / `btnLoadBrands_Click` and `PleadingsLauncher` both check the return value and show appropriate feedback on `False`.

---

## Known Limitations / Not Fixed

### Broad On Error Resume Next in paragraph-iteration loops
Several modules wrap entire paragraph-iteration loops (50-200+ lines) in `On Error Resume Next`. This is a defensive pattern against the Word Object Model, where any paragraph access can throw at any iteration. Tightening these would require restructuring each loop into multiple error-handler blocks, which risks introducing control-flow bugs without the ability to test in a live Word environment.

**Most significant instances:**
- `Rules_Quotes.bas` lines 246-447 (~200 lines, nesting analysis)
- `Rules_Punctuation.bas` lines 766-934 (~170 lines, dash usage)
- `Rules_TextScan.bas` lines 49-167 and 204-341 (~250 lines total, repeated words + spell-out)
- `Rules_NumberFormats.bas` lines 819-934 (~115 lines, currency formatting)
- `Rules_Terms.bas` lines 495-559 (~64 lines, defined terms scan)

**Recommendation:** These should be refactored incrementally with Word-environment testing. The pattern to follow: extract `.Range.Text` under OERN, restore `On Error GoTo 0`, then process the extracted string without error suppression.

### No unit-test harness
VBA has no native test framework. All fixes were verified by code inspection only. Recommend manual regression testing in Word with a document that exercises all rule categories.

---

## Assumptions

1. The `Combined/` directory contains the canonical source for all modules.
2. `modDebugLog.bas` is correct as-is (reviewed, no issues found).
3. `Application.StatusBar` may hold a string or `False`; using `Variant` capture handles both.
4. The existing `Application.Run` dispatch architecture is intentional and correct.
5. Retired rules (23, 18) should remain in source for reference but must not appear active.
