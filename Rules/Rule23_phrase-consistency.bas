Attribute VB_Name = "Rule23_phrase_consistency"
' ============================================================
' Rule23_phrase-consistency.bas
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
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "phrase_consistency"

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_PhraseConsistency(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Define phrase groups ---------------------------------
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

    ' -- Process each group -----------------------------------
    Dim g As Long
    For g = 0 To 9
        CheckPhraseGroup doc, groups(g), issues
    Next g

    Set Check_PhraseConsistency = issues
End Function

' ============================================================
'  PRIVATE: Check a single phrase group for consistency
'  Counts each phrase, determines dominant, flags minorities.
' ============================================================
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

    ' -- Count occurrences of each phrase ---------------------
    For p = LBound(phrases) To UBound(phrases)
        counts(p) = CountPhrase(doc, CStr(phrases(p)))
    Next p

    ' -- Determine how many phrases in this group are used ----
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

    ' -- Flag all minority phrase occurrences -----------------
    For p = LBound(phrases) To UBound(phrases)
        If counts(p) > 0 And p <> dominantIdx Then
            FlagPhraseOccurrences doc, CStr(phrases(p)), CStr(phrases(dominantIdx)), issues
        End If
    Next p
End Sub

' ============================================================
'  PRIVATE: Count occurrences of a phrase in the document
'  Uses Find with MatchWildcards=False, MatchWholeWord=False
'  (necessary for multi-word phrases).
' ============================================================
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

        If EngineIsInPageRange(rng) Then
            cnt = cnt + 1
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    CountPhrase = cnt
End Function

' ============================================================
'  PRIVATE: Flag all occurrences of a minority phrase
' ============================================================
Private Sub FlagPhraseOccurrences(doc As Document, _
                                   minorityPhrase As String, _
                                   dominantPhrase As String, _
                                   ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As Object
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

        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = CreateIssueDict(RULE_NAME, locStr, "Inconsistent phrase: '" & rng.Text & "' used", "Use '" & dominantPhrase & "' for consistency (dominant style)", rng.Start, rng.End, "error")
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
