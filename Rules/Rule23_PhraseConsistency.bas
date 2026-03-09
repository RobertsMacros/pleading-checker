Attribute VB_Name = "Rule23_PhraseConsistency"
' ============================================================
' Rule23_PhraseConsistency.bas
' Proofreading rule: detects inconsistent use of synonymous
' legal phrases. Groups of interchangeable phrases are
' checked; the dominant phrase in each group is identified,
' and minority usages are flagged.
'
' Phrase groups cover common legal/formal alternatives:
'   "not later than" vs "no later than"
'   "pursuant to" vs "in accordance with"
'   "prior to" vs "before"
'   etc.
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "phrase_consistency"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_PhraseConsistency(doc As Document) As Collection
    Dim issues As New Collection

    ' ── Define phrase groups ─────────────────────────────────
    ' Each group is an array of synonymous phrases
    Dim groups(0 To 9) As Variant

    groups(0) = Array("not later than", "no later than")
    groups(1) = Array("in respect of", "with respect to", "in relation to")
    groups(2) = Array("pursuant to", "in accordance with")
    groups(3) = Array("notwithstanding", "despite", "regardless of")
    groups(4) = Array("prior to", "before")
    groups(5) = Array("subsequent to", "after", "following")
    groups(6) = Array("in the event that", "if", "where")
    groups(7) = Array("save that", "except that", "provided that")
    groups(8) = Array("forthwith", "immediately", "without delay")
    groups(9) = Array("hereby", "by this")

    ' ── Process each group ───────────────────────────────────
    Dim g As Long
    For g = 0 To 9
        CheckPhraseGroup doc, groups(g), issues
    Next g

    Set Check_PhraseConsistency = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check a single phrase group for consistency
'  Counts each phrase, determines dominant, flags minorities.
' ════════════════════════════════════════════════════════════
Private Sub CheckPhraseGroup(doc As Document, _
                              phrases As Variant, _
                              ByRef issues As Collection)
    Dim counts() As Long
    Dim phraseCount As Long
    Dim p As Long
    Dim dominantIdx As Long
    Dim dominantCount As Long
    Dim usedCount As Long

    phraseCount = UBound(phrases) - LBound(phrases) + 1
    ReDim counts(LBound(phrases) To UBound(phrases))

    ' ── Count occurrences of each phrase ─────────────────────
    For p = LBound(phrases) To UBound(phrases)
        counts(p) = CountPhrase(doc, CStr(phrases(p)))
    Next p

    ' ── Determine how many phrases in this group are used ────
    usedCount = 0
    dominantIdx = LBound(phrases)
    dominantCount = counts(LBound(phrases))

    For p = LBound(phrases) To UBound(phrases)
        If counts(p) > 0 Then usedCount = usedCount + 1
        If counts(p) > dominantCount Then
            dominantCount = counts(p)
            dominantIdx = p
        End If
    Next p

    ' Only flag if more than one phrase in the group is used
    If usedCount < 2 Then Exit Sub

    ' ── Flag all minority phrase occurrences ─────────────────
    For p = LBound(phrases) To UBound(phrases)
        If counts(p) > 0 And p <> dominantIdx Then
            FlagPhraseOccurrences doc, CStr(phrases(p)), CStr(phrases(dominantIdx)), issues
        End If
    Next p
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Count occurrences of a phrase in the document
'  Uses Find with MatchWildcards=False, MatchWholeWord=False
'  (necessary for multi-word phrases).
' ════════════════════════════════════════════════════════════
Private Function CountPhrase(doc As Document, phrase As String) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean

    cnt = 0

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = phrase
        .MatchWholeWord = False
        .MatchCase = False
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

    CountPhrase = cnt
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Flag all occurrences of a minority phrase
' ════════════════════════════════════════════════════════════
Private Sub FlagPhraseOccurrences(doc As Document, _
                                   minorityPhrase As String, _
                                   dominantPhrase As String, _
                                   ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = minorityPhrase
        .MatchWholeWord = False
        .MatchCase = False
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
                       "Inconsistent phrase: '" & rng.Text & "' used", _
                       "Use '" & dominantPhrase & "' for consistency (dominant style)", _
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
Public Sub RunPhraseConsistency()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Phrase Consistency"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_PhraseConsistency(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Phrase Consistency"
End Sub
