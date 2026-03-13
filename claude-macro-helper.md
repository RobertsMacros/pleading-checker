# VBA Macro Project Guidelines

## Architecture

### Late Binding Over Early Binding
- Use `As Object` + `CreateObject()` instead of typed references like `As Scripting.Dictionary`
- Eliminates compile errors when references aren't set up on the target machine
- For cross-module calls, use `Application.Run("ModuleName.FunctionName", args)` to avoid compile-time dependencies

### Dictionary-Based Data Structures
- Prefer `CreateObject("Scripting.Dictionary")` over custom classes for data transfer objects
- Create a `CreateIssueDict()` or similar factory function in each module to standardise construction
- Access properties with `d("Key")` syntax

### Module Organisation
- One canonical source folder — never duplicate modules across directories
- Group related rules/checks into combined modules to reduce file count
- Use `Private Const` at module top for rule names and configuration
- Use `Option Explicit` in every module

## Performance (Large Documents)

### Critical: Avoid Per-Character Mid$ on Full Document Text
- `Mid(docText, i, 1)` over a 600K-character document creates 600K string allocations
- Convert to byte array instead: `Dim b() As Byte: b = docText`
- Read code points: `code = b(i) Or (CLng(b(i + 1)) * 256&)` (VBA strings are UTF-16LE)
- Per-paragraph `Mid$` loops are acceptable — paragraphs are typically 50-500 chars

### Critical: Reuse Range Objects
- `doc.Range(pos, pos + 1)` creates a new COM object each call — devastating in loops
- Create once, reposition with `.SetRange`: `fontRng.SetRange pos, pos + 1`
- Only create new Range objects for results that need to persist

### Word Find/Replace is Fast
- `Range.Find.Execute` with wildcards uses Word's optimised internal search
- Prefer Find loops over manual character scanning when possible
- Always set `.Wrap = wdFindStop` and collapse after each match to avoid infinite loops

### Paragraph Iteration is Safe
- `For Each para In doc.Paragraphs` is fine even for 5000+ paragraphs
- The overhead is in what you do inside the loop, not the iteration itself

## VBA Language Constraints

### Line Continuation Limit
- VBA allows a maximum of 25 `_` line continuations per logical statement
- Long `Array()` literals are the most common offender
- Split into batches and merge with a helper function:
```vba
Private Function MergeArrays(a As Variant, b As Variant) As Variant
    Dim result() As String
    Dim i As Long, n As Long
    n = UBound(a) - LBound(a) + UBound(b) - LBound(b) + 2
    ReDim result(0 To n - 1)
    For i = LBound(a) To UBound(a)
        result(i - LBound(a)) = a(i)
    Next i
    Dim offset As Long: offset = UBound(a) - LBound(a) + 1
    For i = LBound(b) To UBound(b)
        result(offset + i - LBound(b)) = b(i)
    Next i
    MergeArrays = result
End Function
```

### No GoTo Across Blocks
- `GoTo` labels must be in the same procedure
- Use `GoTo ContinueLabel` for skip-to-next-iteration patterns (VBA has no `Continue For`)

### Error Handling
- `On Error Resume Next` / `On Error GoTo 0` for expected failures (COM calls, Range access)
- Always `Err.Clear` after checking `Err.Number`
- Restore error handling with `On Error GoTo 0` as soon as the risky section ends

### Variable Declarations
- `Dim` inside a `For` loop body is legal but the variable is scoped to the entire procedure
- Avoid `Dim` inside conditionals/loops for clarity — declare at procedure top

## Common Patterns

### Safe Range Access
```vba
On Error Resume Next
Err.Clear
Set rng = doc.Range(startPos, endPos)
If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Sub
End If
On Error GoTo 0
```

### Byte-Array Character Scanning
```vba
Dim b() As Byte: b = text
Dim bMax As Long: bMax = UBound(b) - 1
Dim i As Long, code As Long

For i = 0 To bMax Step 2
    code = b(i) Or (CLng(b(i + 1)) * 256&)
    ' code is the Unicode code point
    ' document position = i \ 2
Next i
```

### Reusable Range for Property Checks
```vba
Dim fontRng As Range
Set fontRng = doc.Range(0, 1)

' Inside loop:
fontRng.SetRange pos, pos + 1
fontName = fontRng.Font.Name
```

### Late-Bound Cross-Module Calls
```vba
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function
```

### Dictionary-Based Finding Factory
```vba
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function
```

## Automated Refactoring Safety

### Multi-Line Statement Handling
- VBA uses `_` for line continuation — any regex/script operating on VBA source must join continuation lines before matching
- A single-line regex on a multi-line VBA statement will silently truncate the statement
- Always reconstruct the full logical line before transforming, then re-split if needed

### Variable Rename Patterns
- `Set x = New ClassName` and `Dim x As New ClassName` are distinct patterns — handle both
- Module prefixes vary (`Rules_`, `Rule01_`, etc.) — match flexibly
- After renaming, verify no references to the old name remain

### Validation Checks After Bulk Edits
- Count quotes per non-continuation line — odd count means unclosed string
- Count line continuations per statement — must stay under 25
- Verify `CreateObject`/factory calls have the correct argument count
- Check for duplicate `Dim` statements (common artifact of automated insertion)
