Attribute VB_Name = "Rule21_TitleFormatting"
' ============================================================
' Rule21_TitleFormatting.bas
' Proofreading rule: detects inconsistent use of titles/
' honorifics with or without trailing periods.
'
' Checks pairs like Mr/Mr., Mrs/Mrs., Dr/Dr., QC/Q.C. etc.
' Determines dominant style and flags minority occurrences.
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "title_formatting"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_TitleFormatting(doc As Document) As Collection
    Dim issues As New Collection

    ' ── Define title pairs: noDot / withDot ──────────────────
    Dim noDot As Variant
    Dim withDot As Variant

    noDot = Array("Mr", "Mrs", "Ms", "Dr", "Prof", "QC", "KC", "MP", "JP")
    withDot = Array("Mr.", "Mrs.", "Ms.", "Dr.", "Prof.", "Q.C.", "K.C.", "M.P.", "J.P.")

    Dim i As Long
    Dim noDotCount As Long
    Dim withDotCount As Long

    For i = LBound(noDot) To UBound(noDot)
        noDotCount = CountWordInDoc(doc, CStr(noDot(i)))
        withDotCount = CountWordInDoc(doc, CStr(withDot(i)))

        ' Only flag if both forms exist
        If noDotCount > 0 And withDotCount > 0 Then
            If noDotCount >= withDotCount Then
                ' noDot is dominant — flag all withDot occurrences
                FlagOccurrences doc, CStr(withDot(i)), _
                    "Inconsistent title formatting: '" & withDot(i) & "' used", _
                    "Use '" & noDot(i) & "' without period (dominant style)", _
                    issues
            Else
                ' withDot is dominant — flag all noDot occurrences
                FlagOccurrences doc, CStr(noDot(i)), _
                    "Inconsistent title formatting: '" & noDot(i) & "' used", _
                    "Use '" & withDot(i) & "' with period (dominant style)", _
                    issues
            End If
        End If
    Next i

    Set Check_TitleFormatting = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Count occurrences of a word in the document
'  Uses Find with MatchWholeWord and MatchCase.
' ════════════════════════════════════════════════════════════
Private Function CountWordInDoc(doc As Document, word As String) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean
    Dim useWildcards As Boolean

    cnt = 0

    ' For dotted abbreviations like "Q.C.", use non-wildcard search
    ' For simple words like "Mr", use whole-word matching
    useWildcards = False

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = word
        .MatchWholeWord = True
        .MatchCase = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        If PleadingsEngine.IsInPageRange(rng) Then
            cnt = cnt + 1
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    CountWordInDoc = cnt
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Flag all occurrences of a minority form
' ════════════════════════════════════════════════════════════
Private Sub FlagOccurrences(doc As Document, _
                             word As String, _
                             issueText As String, _
                             suggestionText As String, _
                             ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = word
        .MatchWholeWord = True
        .MatchCase = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        If PleadingsEngine.IsInPageRange(rng) Then
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       issueText, _
                       suggestionText, _
                       rng.Start, _
                       rng.End, _
                       "error"
            issues.Add issue
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunTitleFormatting()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Title Formatting"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_TitleFormatting(doc)

    ' ── Highlight issues in document ─────────────────────────
    Dim iss As PleadingsIssue
    Dim rng As Range
    Dim i As Long
    For i = 1 To issues.Count
        Set iss = issues(i)
        If iss.RangeStart >= 0 And iss.RangeEnd > iss.RangeStart Then
            On Error Resume Next
            Set rng = doc.Range(iss.RangeStart, iss.RangeEnd)
            rng.HighlightColorIndex = wdYellow
            doc.Comments.Add Range:=rng, _
                Text:="[" & iss.RuleName & "] " & iss.Issue & _
                      " " & Chr(8212) & " Suggestion: " & iss.Suggestion
            On Error GoTo 0
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Title Formatting"
End Sub
