# Refactoring Plan: Shared Helper Module Architecture

## Goal
Consolidate duplicated logic into TextAnchoring.bas (the existing shared utility module) and rewrite all rule modules to call these shared helpers instead of duplicating boilerplate.

## Phase 1: Add Shared Helpers to TextAnchoring.bas

Add the following new Public functions:

### 1.1 `SafeLocationString(rng, doc) As String`
Replaces the 4-line On Error / GetLocationString / Err.Clear / On Error GoTo 0 pattern used ~30 times across all rule files.

### 1.2 `SafeRange(doc, startPos, endPos) As Range`
Creates a Range with error handling, returns Nothing on failure. Replaces the repeated Set rng = doc.Range / If Err.Number pattern.

### 1.3 `IsWhitespaceChar(ch) As Boolean`
Centralises `ch = " " Or ch = vbTab Or ch = ChrW(160) Or ch = vbCr Or ch = vbLf Or ch = Chr(11)` used in Rules_TextScan.bas and Rules_Spacing.bas.

### 1.4 `CreateRegex(pattern, global, ignoreCase) As Object`
Factory for VBScript.RegExp. Replaces 4-line CreateObject/set properties blocks used in Rules_Spacing.bas, Rules_TextScan.bas, Rules_Punctuation.bas, Rules_NumberFormats.bas.

### 1.5 `AddIssue(issues, ruleName, rng, doc, msg, suggestion, startPos, endPos, severity, autoFixSafe, replacementText, matchedText, anchorKind, confidence)`
Combines SafeLocationString + CreateIssueDict + issues.Add into one call. This is the single biggest line-saver across the entire codebase.

### 1.6 `IterateParagraphs(doc, moduleName, procName) As Collection`
Generic paragraph iteration loop that:
- Iterates `doc.Paragraphs`
- Gets paraRange with error handling
- Checks IsPastPageFilter / IsInPageRange
- Gets paraText via StripParaMarkChar
- Calculates listPrefixLen
- Dispatches to `Application.Run moduleName & "." & procName` with the standard ProcessParagraph_ signature: `(doc, paraRange, paraText, paraStart, listPrefixLen, issues)`
- Returns the issues collection

This eliminates ~20-30 lines of boilerplate from every Check_ function that iterates paragraphs.

### 1.7 `FindAll(doc, searchText, wholeWord, matchCase, useWildcards, Optional searchRange) As Collection`
Generic Find loop with stall guard. Returns a Collection of 3-element Arrays: `Array(startPos, endPos, matchText)`. Replaces the duplicated Find/Execute/stall-guard/Collapse pattern in Rules_Brands.bas, Rules_LegalTerms.bas, Rules_Punctuation.bas, Rules_Spacing.bas, Rules_NumberFormats.bas.

### 1.8 `ForEachFootnote(doc, moduleName, procName) As Collection`
Generic footnote iteration that:
- Iterates `doc.Footnotes(i)` with error handling
- Checks IsInPageRange on fn.Reference
- Gets noteText = fn.Range.Text
- Dispatches to `Application.Run moduleName & "." & procName, doc, fn, noteText, issues`
- Returns the issues collection

## Phase 2: Create Missing ProcessParagraph_ Handlers

The following Check_ functions use paragraph iteration but don't yet have ProcessParagraph_ equivalents:
- Check_SpaceBeforePunct (Rules_Spacing.bas) — uses Find loop, not para loop, so stays as-is but uses FindAll
- Check_MandatedLegalTermForms (Rules_LegalTerms.bas) — uses Find loop via SearchAndFlag, stays as-is but uses FindAll

No new ProcessParagraph_ handlers needed — the missing ones all use Find loops, not paragraph iteration.

## Phase 3: Rewrite Check_ Functions

For each Check_ function, replace boilerplate with shared helper calls:

### Paragraph-loop rules (use IterateParagraphs):
- `Check_DoubleSpaces` → `IterateParagraphs(doc, "Rules_Spacing", "ProcessParagraph_DoubleSpaces")`
- `Check_DoubleCommas` → `IterateParagraphs(doc, "Rules_Spacing", "ProcessParagraph_DoubleCommas")`
- `Check_MissingSpaceAfterDot` → `IterateParagraphs(doc, "Rules_Spacing", "ProcessParagraph_MissingSpaceAfterDot")`
- `Check_RepeatedWords` → `IterateParagraphs(doc, "Rules_TextScan", "ProcessParagraph_RepeatedWords")`
- `Check_SpellOutUnderTen` → `IterateParagraphs(doc, "Rules_TextScan", "ProcessParagraph_SpellOutUnderTen")`
- `Check_TriplicatePunctuation` → `IterateParagraphs(doc, "Rules_Punctuation", "ProcessParagraph_TriplicatePunctuation")`
- `Check_DashUsage` → `IterateParagraphs(doc, "Rules_Punctuation", "ProcessParagraph_DashUsage")`
- `Check_BracketIntegrity` → keeps its global pre-check, then `IterateParagraphs` for per-paragraph work
- `Check_AnglicisedTermsNotItalic` → `IterateParagraphs(doc, "Rules_Italics", "ProcessParagraph_AnglicisedTerms")`
- `Check_ForeignNamesNotItalic` → `IterateParagraphs(doc, "Rules_Italics", "ProcessParagraph_ForeignNames")`
- `Check_AlwaysCapitaliseTerms` → `IterateParagraphs(doc, "Rules_LegalTerms", "ProcessParagraph_AlwaysCapitalise")`

### Find-loop rules (use FindAll + AddIssue):
- `SearchAndFlag` in Rules_Brands.bas → use FindAll + AddIssue
- `SearchAndFlag` in Rules_LegalTerms.bas → use FindAll + AddIssue
- `Check_SpaceBeforePunct` → use FindAll + AddIssue
- `FlagSpacedSlashes / FlagTightSlashes / FlagBackslashes / CountTightSlashes / CountSpacedSlashes` in Rules_Punctuation.bas → use FindAll
- `FindWithWildcard` in Rules_NumberFormats.bas → use FindAll

### Footnote rules (use ForEachFootnote):
- `Check_FootnoteTerminalFullStop` → ForEachFootnote + ProcessFootnote_TerminalFullStop
- `Check_FootnoteInitialCapital` → ForEachFootnote + ProcessFootnote_InitialCapital
- `Check_FootnoteAbbreviationDictionary` → ForEachFootnote + ProcessFootnote_AbbreviationDictionary
- `Check_FootnoteIntegrity` → ForEachFootnote (if applicable, may need special handling)
- `Check_DuplicateFootnotes` → builds a dictionary across all footnotes, so can still use ForEachFootnote for iteration

### Rules that stay largely unchanged:
- `Check_Spelling` — uses its own optimised single-pass tokeniser, just simplify with SafeLocationString/AddIssue
- `Check_LicenceLicense` / `Check_CheckCheque` — use Find loops, simplify with FindAll
- `Check_FootnotesNotEndnotes` — trivial, no iteration pattern to simplify
- `Check_DateTimeFormat` / `Check_CurrencyNumberFormat` — already use FindWithWildcard helper, just replace it with FindAll
- `Check_CustomTermWhitelist` — uses Find loop, simplify with FindAll

## Phase 4: Simplify ProcessParagraph_ Internals

Within each ProcessParagraph_ function, replace:
- Location string fetch + CreateIssueDict + Add → `AddIssue`
- doc.Range creation → `SafeRange`
- Regex creation → `CreateRegex`
- Whitespace checks → `IsWhitespaceChar`

## Phase 5: Clean Up Dead Code

After rewriting, remove:
- Duplicated paragraph iteration boilerplate from Check_ functions
- Duplicated Find loop boilerplate
- Private SearchAndFlag functions that are replaced by FindAll
- FindWithWildcard in Rules_NumberFormats.bas (replaced by FindAll)

## Ordering

1. Phase 1 first (add helpers) — no existing code changes, purely additive
2. Phase 3+4 together per file — rewrite each Rules_*.bas file one at a time
3. Phase 5 last — clean up after all rewrites verified
