# CLAUDE.md — Word VBA Macro Development Guide

## Project overview

This is a Word VBA macro project (`.bas` / `.frm` modules). All source lives in `Combined/`. The macro is loaded into a Word document or template and executed from the VBA editor or a ribbon button.

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

**Rule of thumb:** If a variable name matches any VBA keyword, built-in function, or type name, rename it.

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

Word's `Range` object has no `.Runs` collection (unlike PowerPoint). To walk character formatting runs, use `Range.Characters` or the `wdCharacterFormatting` find approach:

```vba
Dim rng As Range
Set rng = doc.Content.Duplicate
rng.Collapse wdCollapseStart
Do While rng.Start < doc.Content.End
    rng.MoveEndUntil Cset:="", Count:=wdForward  ' etc.
Loop
```

### 6. `Application.EnableEvents` is Excel-only

`Application.EnableEvents` does not exist in Word VBA. Using it causes a compile error that **silently prevents the entire module from loading** — the macro appears to do nothing (0 issues found) with no visible error.

### 7. Duplicate `Dim` statements

VBA does not allow two `Dim` statements for the same variable name in the same procedure scope, even inside different `If` blocks. This is a compile error.

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

Some VBA versions reject single-line `Dim` + assignment. Always split:

```vba
' Safer
Dim x As Long
x = 0
```

### 10. Fixed-size array `Dim` inside loops

Declaring a fixed-size array (`Dim arr(0 To 63) As Long`) inside a loop causes a compile error on re-entry. Declare at procedure top; use `ReDim` to reset inside the loop.

### 11. Function call syntax — no trailing parentheses for Subs

```vba
' WRONG — extra () causes type mismatch or compile error
Call SomeSub(arg1, arg2)()
result = SomeFunc(arg1)()

' RIGHT
Call SomeSub(arg1, arg2)
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

Always wrap long-running operations:

```vba
Application.ScreenUpdating = False
' ... do work ...
DoEvents  ' between major steps to prevent "Not Responding"
Application.ScreenUpdating = True
```

### 17. `On Error Resume Next` leaks across statements

After using `On Error Resume Next`, always reset with `On Error GoTo 0` before returning to normal flow. Forgetting this masks all subsequent errors silently.

---

## Filtering false positives

When checking documents, certain regions should be excluded from content-based rules:

### Cover pages
Page 1 of a multi-page document when it has no numbered paragraphs (tribunal cover sheets, title pages). Detected via first section break or page-1 paragraph analysis. **All rules suppressed.**

### Contents / Table of Contents pages
Detected via Word's built-in `TablesOfContents` collection, TOC-styled paragraphs, or dot/tab leader patterns (text followed by `....` or tabs then a page number). **All rules suppressed.**

### Block quotes / quoted text
Detected by style name ("quote", "block", "extract"), significant left indentation with smaller font, or text wrapped in quotation marks. **Content rules suppressed** (spelling, grammar, numbers, quotes) but **formatting rules still apply** (font consistency, colour formatting, paragraph breaks) because formatting belongs to the author, not the source.

### Footnotes and endnotes
Have their own dedicated rule set. Body-text rules should not flag footnote content; footnote rules should not flag body text.

---

## Code style conventions

- All rule functions return a `Collection` of `Dictionary` objects (created via `CreateIssueDict` / `CreateIssue`)
- Each issue dictionary has keys: `RuleName`, `Message`, `Suggestion`, `RangeStart`, `RangeEnd`, `Severity`
- Use `On Error Resume Next` with `Err.Clear` around any Range/Paragraph property access (documents can have corrupt ranges)
- Prefer `LCase()` comparisons over case-sensitive string matching
- Keep individual rule modules independent — no cross-rule imports
