Attribute VB_Name = "Rule21_title_formatting"
' ============================================================
' Rule21_title-formatting.bas
' Proofreading rule: detects inconsistent use of titles/
' honorifics with or without trailing periods.
'
' Checks pairs like Mr/Mr., Mrs/Mrs., Dr/Dr., QC/Q.C. etc.
' Determines dominant style and flags minority occurrences.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "title_formatting"

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_TitleFormatting(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Define title pairs: noDot / withDot ------------------
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
                ' noDot is dominant -- flag all withDot occurrences
                FlagOccurrences doc, CStr(withDot(i)), _
                    "Inconsistent title formatting: '" & withDot(i) & "' used", _
                    "Use '" & noDot(i) & "' without period (dominant style)", _
                    issues
            Else
                ' withDot is dominant -- flag all noDot occurrences
                FlagOccurrences doc, CStr(noDot(i)), _
                    "Inconsistent title formatting: '" & noDot(i) & "' used", _
                    "Use '" & withDot(i) & "' with period (dominant style)", _
                    issues
            End If
        End If
    Next i

    Set Check_TitleFormatting = issues
End Function

' ============================================================
'  PRIVATE: Count occurrences of a word in the document
'  Uses Find with MatchWholeWord and MatchCase.
' ============================================================
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

        If EngineIsInPageRange(rng) Then
            cnt = cnt + 1
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    CountWordInDoc = cnt
End Function

' ============================================================
'  PRIVATE: Flag all occurrences of a minority form
' ============================================================
Private Sub FlagOccurrences(doc As Document, _
                             word As String, _
                             issueText As String, _
                             suggestionText As String, _
                             ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As Object
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

        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = CreateIssueDict(RULE_NAME, locStr, issueText, suggestionText, rng.Start, rng.End, "error")
            issues.Add issue
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based issue (no class dependency)
' ----------------------------------------------------------------
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

' ----------------------------------------------------------------
'  Late-bound wrapper: EngineIsInPageRange
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  Late-bound wrapper: EngineGetLocationString
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function
