# VBA_learnings.md — Word VBA Macro Development Guide

## Project overview

This is a Word VBA macro project (`.bas` / `.frm` modules). All source lives in `Code/`. The macro is loaded into a Word document or template and executed from the VBA editor or a ribbon button.

There is no build step, test suite, or CI pipeline. Validation is manual: import the `.bas` files into Word's VBA editor, compile (Debug > Compile), and run against test documents.

---

## VBA pitfalls — lessons learned

These are concrete issues encountered during development that caused compile errors, runtime errors, silent failures, or Word freezes. They apply to **any Word VBA project**.

### 1. Late-bind all Scripting.Dictionary references

VBA early binding (`Dim d As Scripting.Dictionary`) requires the "Microsoft Scripting Runtime" reference to be set in every target machine. Use late binding instead:

```vba
' WRONG — compile error if reference not set
Dim d As Scripting.Dictionary
Set d = New Scripting.Dictionary

' RIGHT — works everywhere
Dim d As Object
Set d = CreateObject("Scripting.Dictionary")
```

### 2. Reserved words as variable names

VBA has many reserved words that are not always caught at compile time but cause runtime errors or silent corruption. Words we hit:

- `variant` — reserved type name. Renamed to `brandVariant`
- `Stop` — reserved statement (also `sTop` is case-insensitive match). Renamed to `stackTop`
- `input` — shadows `VBA.Input`. Renamed to `pageInput`
- `issue` — not reserved but conflicts with common patterns; renamed to `finding` for clarity

**Rule of thumb:** If a variable name matches any VBA keyword, built-in function, or type name, rename it. VBA is case-insensitive, so `sTop` matches `Stop`.

### 3. `Const` cannot use function calls

```vba
' WRONG — VBA Const must be a literal, not a function call
Const PUNCT_CHARS As String = Chr(8212) & Chr(8211)

' RIGHT — use Dim instead
Dim PUNCT_CHARS As String
PUNCT_CHARS = Chr(8212) & Chr(8211)
```

### 4. Non-ASCII characters in source code

The VBA editor uses the system's ANSI code page. Non-ASCII characters (em-dashes, curly quotes, box-drawing chars) become mojibake when the `.bas` file is imported on a different locale. Always use `Chr()` / `ChrW()` for non-ASCII, or stick to ASCII equivalents (e.g. `"--"` instead of an em-dash in user-facing strings).

### 5. `Range.Runs` does not exist in Word VBA

Word's `Range` object has no `.Runs` collection (unlike PowerPoint). There is also no `Ranges` type. To walk character formatting runs, use `wdCharacterFormatting`:

```vba
Dim rn As Range
Set rn = doc.Content.Duplicate
rn.Collapse wdCollapseStart
Do While rn.Start < doc.Content.End
    rn.MoveEnd wdCharacterFormatting, 1
    ' ... inspect rn.Font properties ...
    rn.Collapse wdCollapseEnd
Loop
```

### 6. `Application.EnableEvents` is Excel-only

`Application.EnableEvents` does not exist in Word VBA. Using it causes a compile error that **silently prevents the entire module from loading** — the macro appears to do nothing (0 issues found) with no visible error. Similarly, `Application.DisplayAlerts` save/restore is unnecessary in Word rule-running context.

### 7. Duplicate `Dim` statements and duplicate functions

VBA does not allow two `Dim` statements for the same variable name in the same procedure scope, even inside different `If` blocks. Similarly, two functions with the same name in one module is a compile error. Both are easy to introduce during automated refactoring.

### 8. Max 25 line continuations per statement

VBA limits a single statement to 25 line-continuation characters (`_`). Large `Array()` literals easily exceed this. Split into batches and merge:

```vba
' WRONG — 30+ continuations
Dim arr As Variant
arr = Array("a", "b", ... 60 items ... "zz")

' RIGHT — split and merge
Dim a1 As Variant, a2 As Variant
a1 = Array("a", "b", ... 20 items)
a2 = Array("u", "v", ... 20 items)
arr = MergeArrays(a1, a2)
```

### 9. `Dim x As Type: x = value` on one line

Some VBA versions reject single-line `Dim` + assignment. Always split into separate lines for reliability.

### 10. Fixed-size array `Dim` inside loops

Declaring a fixed-size array (`Dim arr(0 To 63) As Long`) inside a loop causes a compile error on re-entry. Declare at procedure top with no bounds; use `ReDim` to reset inside the loop.

### 11. Function call syntax — no trailing parentheses

```vba
' WRONG — extra () causes type mismatch or runtime error 5
result = SomeFunc(arg1)()

' RIGHT
result = SomeFunc(arg1)
```

### 12. Cross-module calls must use `Application.Run`

Direct calls like `ModuleName.FunctionName(args)` fail if the module isn't loaded yet or has a compile error. Use `Application.Run` for resilience:

```vba
Dim result As Object
Set result = Application.Run("ModuleName.FunctionName", doc)
```

### 13. `Chr(8212)` and em-dash runtime issues

Some Word/VBA environments throw runtime errors on `Chr(8212)` in string concatenation contexts. Prefer `ChrW(8212)` or avoid em-dashes in user-facing strings entirely (use `"--"` instead).

### 14. MsgBox string length limit (Error 5)

`MsgBox` has an approximate 1024-character limit for the prompt string. Exceeding it throws runtime error 5 ("Invalid procedure call or argument"). Truncate long output or use a UserForm instead.

### 15. `Find.Execute` infinite loops

If `Find.Execute` matches but `Range.Collapse` doesn't advance past the match, the loop runs forever and Word freezes. Always add a stall guard:

```vba
Dim lastPos As Long
lastPos = -1
Do While rng.Find.Execute(...)
    If rng.Start = lastPos Then Exit Do  ' stall guard
    lastPos = rng.Start
    ' ... process match ...
    rng.Collapse wdCollapseEnd
Loop
```

### 16. Screen updating and responsiveness

Always wrap long-running operations. Every `Find.Execute` redraws the screen without this:

```vba
Application.ScreenUpdating = False
' ... do work ...
DoEvents  ' between major steps to prevent "Not Responding"
Application.ScreenUpdating = True
```

Use `On Error GoTo cleanup` to guarantee `ScreenUpdating = True` is restored even on error.

### 17. `On Error Resume Next` leaks across statements

After using `On Error Resume Next`, always reset with `On Error GoTo 0` before returning to normal flow. Forgetting this masks all subsequent errors silently.

### 18. Self-referencing Select Case causes infinite recursion

```vba
' WRONG — calls itself forever, stack overflow
Function GetProp(obj, propName)
    Select Case propName
        Case "Name": GetProp = GetProp(obj, "Name")  ' oops
    End Select
End Function

' RIGHT — use CallByName or direct access
Case "Name": GetProp = CallByName(obj, propName, VbGet)
```

### 19. Truncated arguments in automated refactoring

When using find-and-replace or scripts to refactor function calls (e.g. changing from class constructors to helper functions), arguments can be silently lost if the replacement pattern doesn't capture all parameters. Always verify argument counts match after automated changes.

### 20. Silent rule failures with no diagnostics

`On Error Resume Next` around `Application.Run` for rule dispatch means failures are completely invisible — rules fail and report "0 issues found" with no indication anything went wrong. Always add error logging:

```vba
On Error Resume Next
Set result = Application.Run(funcName, doc)
If Err.Number <> 0 Then
    Debug.Print "RULE ERROR: " & funcName & " -- " & Err.Description
    errorLog = errorLog & funcName & vbCrLf
    Err.Clear
End If
On Error GoTo 0
```

### 21. Character-by-character Range allocation is catastrophically slow

Creating a new `doc.Range(pos, pos+1)` for each character in a document to inspect font properties causes Word to freeze on any real document. Use `wdCharacterFormatting` walk or byte-array scans instead. Orders of magnitude faster.

### 22. Class dependencies prevent standalone compilation

If rule modules type variables `As SomeClass`, they won't compile unless that `.cls` file is also imported. Use Dictionary-based objects with a factory helper function instead, so every module compiles independently.

### 23. `wdCharacterFormatting` MoveEnd can infinite-loop

`Range.MoveEnd wdCharacterFormatting, 1` can return 0 (no movement) on certain paragraph structures without raising an error. Always check the return value or add a position guard to prevent infinite loops.

### 24. `AutoFixSafe = True` with descriptive Suggestion corrupts documents

When `AutoFixSafe = True`, the engine applies the `Suggestion` field as a tracked change via `rng.Text = sugText`. If the Suggestion is a human-readable message like `"Add a full stop at the end."`, that literal text replaces the target range, corrupting the document. **Suggestions for auto-fixable issues must be literal replacement values** (e.g. `","`, `ChrW(8211)`, `""` for deletions) — never descriptive text. If the fix is complex, set `AutoFixSafe = False`.

```vba
' WRONG — applies message text as tracked change
Set finding = CreateIssueDict(rule, loc, msg, _
    "Replace ',,' with a single ','.", start, end, "error", True)

' RIGHT — literal replacement value
Set finding = CreateIssueDict(rule, loc, msg, _
    ",", start, end, "error", True)

' RIGHT — descriptive suggestion, not auto-fixable
Set finding = CreateIssueDict(rule, loc, msg, _
    "Add a full stop at the end.", start, end, "warning", False)
```

### 25. Tracked change comment anchoring on deletions

When a tracked change deletes text (empty suggestion), the range collapses to zero length after `rng.Text = ""`. Word then anchors any subsequent comment on the nearest word instead of the deletion mark. Preserve the original range length before the tracked change and re-anchor:

```vba
origStart = rng.Start
origLen = rng.End - rng.Start
doc.TrackRevisions = True
rng.Text = sugText
doc.TrackRevisions = wasTracking
' For comment anchor:
If Len(sugText) > 0 Then
    Set commentRng = doc.Range(origStart, origStart + Len(sugText))
Else
    Set commentRng = doc.Range(origStart, origStart + origLen)
End If
```

### 26. Space-targeting rules must only affect spaces

Rules that fix spaces (double spaces, trailing spaces, space-before-punctuation) must:
- Set range to cover **only the spaces**, not surrounding text
- Set suggestion to `""` (empty = delete) — never to a descriptive message
- For double spaces: target `dsStart + 1` to `dsEnd` (keep first space, delete extras)
- For trailing spaces: target `paraRange.End - 1 - numSpaces` to `paraRange.End - 1`
- For space-before-punct: target `rng.Start` to `rng.Start + 1` (just the space)

### 27. Block quote detection must not catch lists

Indented paragraphs are not necessarily block quotes. Numbered/bulleted lists in legal documents are often heavily indented. Block quote detection requires **at least one** of:
- A quote/block/extract style name (definitive)
- Quotation marks at start or end of the indented text
- Entirely italic formatting
- Font size clearly smaller than body text (e.g. 9pt vs 12pt body)

**Pure indentation alone — even heavy (>72pt / 1 inch) — is NOT sufficient.** If the paragraph has body-sized font, is not italic, and has no quotation marks, it is a list or other indented content, not a block quote.

### 28. VBA `Trim` does not strip tabs or non-breaking spaces

`Trim$()` only removes ASCII spaces (Chr(32)). Tabs (`vbTab`), non-breaking spaces (`ChrW(160)`), and other whitespace survive. When checking if paragraph text starts/ends with specific characters (like quotation marks), strip these first:

```vba
pText = Replace(Replace(Replace(para.Range.Text, vbCr, ""), vbTab, ""), ChrW(160), "")
pText = Trim$(pText)
```

### 29. `.MatchWholeWord` treats hyphens as word boundaries

Word's `Find.MatchWholeWord` considers hyphens as word separators. Searching for "check" with `MatchWholeWord = True` will match "double-check" (finding "check" after the hyphen). Build exception arrays for compound words and check before/after context manually.

### 30. VBA is entirely case-insensitive — variable names must not collide with any identifier

VBA treats `MyVar`, `myvar`, `MYVAR`, and `myVar` as the same symbol. This means a variable name that happens to match a built-in function, method, property, enum constant, or type — even in a different case — will shadow or collide with it. The VBA editor will silently "auto-correct" the casing of one to match the other everywhere in the module, making the conflict invisible.

**Dangerous collisions include:**
- `text` / `Text` — shadows `Range.Text`, `TextBox.Text`, etc.
- `name` / `Name` — shadows `Name` statement and `.Name` properties
- `type` / `Type` — reserved keyword for user-defined types
- `value` / `Value` — shadows `.Value` on controls, cells, fields
- `count` / `Count` — shadows `.Count` on collections
- `replace` / `Replace` — shadows `VBA.Replace()` function
- `format` / `Format` — shadows `VBA.Format()` function
- `left` / `Left` — shadows `VBA.Left()` function
- `right` / `Right` — shadows `VBA.Right()` function
- `mid` / `Mid` — shadows `VBA.Mid()` function
- `len` / `Len` — shadows `VBA.Len()` function
- `trim` / `Trim` — shadows `VBA.Trim()` function
- `date` / `Date` — shadows `VBA.Date` function/type

**Rule:** Always prefix or qualify variable names to avoid any ambiguity: `paraText` not `text`, `itemCount` not `count`, `sugLen` not `len`, `startPos` not `start`, etc.

---

## Filtering false positives

When checking documents, certain regions should be excluded from content-based rules:

### Cover pages
Page 1 of a multi-page document when it has no numbered paragraphs (tribunal cover sheets, title pages). Detected via first section break or page-1 paragraph analysis. **All rules suppressed.**

### Contents / Table of Contents pages
Detected via Word's built-in `TablesOfContents` collection, TOC-styled paragraphs, or dot/tab leader patterns (text followed by `....` or tabs then a page number). **All rules suppressed.**

### Block quotes / quoted text
Detected by style name ("quote", "block", "extract"), or by a combination of indentation with at least one distinguishing feature: smaller font than body, italic formatting, or quotation marks at start/end. **Indentation alone is not sufficient** — heavily indented paragraphs at body font size without quotes or italics are lists, not block quotes. **Content rules suppressed** (spelling, grammar, numbers, quotes) but **formatting rules still apply** (font consistency, colour formatting, paragraph breaks) because formatting belongs to the author, not the source.

### Footnotes and endnotes
Have their own dedicated rule set. Body-text rules should not flag footnote content; footnote rules should not flag body text.

---

## Code style conventions

- All rule functions return a `Collection` of `Dictionary` objects (created via `CreateIssueDict` / `CreateIssue`)
- Each issue dictionary has keys: `RuleName`, `Location`, `Issue`, `Suggestion`, `RangeStart`, `RangeEnd`, `Severity`, `AutoFixSafe`
- **AutoFixSafe rule:** When `AutoFixSafe = True`, `Suggestion` must be a **literal replacement value** (or `""` for deletion). Never a human-readable description. When the fix is too complex for a literal, use `AutoFixSafe = False`.
- **Range targeting:** `RangeStart`/`RangeEnd` must target exactly the text being flagged or replaced — spaces target only spaces, dashes target only the dash character, etc.
- Use `On Error Resume Next` with `Err.Clear` around any Range/Paragraph property access (documents can have corrupt ranges)
- Prefer `LCase()` comparisons over case-sensitive string matching
- Keep individual rule modules independent — no cross-rule imports
