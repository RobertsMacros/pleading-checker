Attribute VB_Name = "Rule26_footnote_initial_capital"
' ============================================================
' Rule26_footnote-initial-capital.bas
' Proofreading rule: footnotes should begin with a capital
' letter, except for approved lower-case abbreviations.
'
' Approved lower-case starts:
'   c, cf, cp, eg, ie, p, pp, ibid
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "footnote_initial_capital"

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_FootnoteInitialCapital(doc As Document) As Collection
    Dim issues As New Collection
    Dim allowed As Object
    Dim fn As Footnote
    Dim issue As Object
    Dim locStr As String
    Dim noteText As String
    Dim trimmed As String
    Dim token As String
    Dim firstCharCode As Long
    Dim i As Long
    Dim j As Long
    Dim ch As String

    ' -- Build allowed lower-case starts dictionary -----------
    Set allowed = CreateObject("Scripting.Dictionary")
    allowed.CompareMode = vbTextCompare
    allowed.Add "c", True
    allowed.Add "cf", True
    allowed.Add "cp", True
    allowed.Add "eg", True
    allowed.Add "ie", True
    allowed.Add "p", True
    allowed.Add "pp", True
    allowed.Add "ibid", True

    For i = 1 To doc.Footnotes.Count
        On Error Resume Next
        Set fn = doc.Footnotes(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote
        End If
        On Error GoTo 0

        ' -- Check page range on the reference mark -----------
        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFootnote
        End If
        On Error GoTo 0

        ' -- Get footnote text --------------------------------
        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote
        End If
        On Error GoTo 0

        ' -- Trim leading whitespace --------------------------
        trimmed = LTrim(noteText)
        If Len(trimmed) = 0 Then GoTo NextFootnote

        ' -- Skip past leading punctuation (quotes, brackets) -
        j = 1
        Do While j <= Len(trimmed)
            ch = Mid(trimmed, j, 1)
            If IsLeadingPunctuation(ch) Then
                j = j + 1
            Else
                Exit Do
            End If
        Loop

        If j > Len(trimmed) Then GoTo NextFootnote
        trimmed = Mid(trimmed, j)
        If Len(trimmed) = 0 Then GoTo NextFootnote

        ' -- Extract first lexical token (letters only) -------
        token = ExtractFirstToken(trimmed)
        If Len(token) = 0 Then GoTo NextFootnote

        ' -- Check if token is in allowed list ----------------
        If allowed.Exists(LCase(token)) Then GoTo NextFootnote

        ' -- Check if first character is lower-case -----------
        firstCharCode = AscW(Mid(token, 1, 1))
        If firstCharCode >= 97 And firstCharCode <= 122 Then
            ' Lower-case and not in allowed list: flag
            On Error Resume Next
            locStr = EngineGetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = CreateIssueDict(RULE_NAME, locStr, "Footnote begins with lower-case text outside the approved exceptions.", "Begin the footnote with a capital letter, unless it starts with an approved lower-case abbreviation.", fn.Range.Start, fn.Range.End, "warning", False)
            issues.Add issue
        End If

NextFootnote:
    Next i

    Set Check_FootnoteInitialCapital = issues
End Function

' ============================================================
'  PRIVATE: Check if character is leading punctuation to skip
' ============================================================
Private Function IsLeadingPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case "(", "[", ChrW(8216), ChrW(8220), """", "'"
            IsLeadingPunctuation = True
        Case Else
            IsLeadingPunctuation = False
    End Select
End Function

' ============================================================
'  PRIVATE: Extract the first token of letters from a string
' ============================================================
Private Function ExtractFirstToken(ByVal s As String) As String
    Dim i As Long
    Dim charCode As Long
    Dim result As String
    result = ""

    For i = 1 To Len(s)
        charCode = AscW(Mid(s, i, 1))
        ' A-Z = 65-90, a-z = 97-122
        If (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Then
            result = result & Mid(s, i, 1)
        Else
            Exit For
        End If
    Next i

    ExtractFirstToken = result
End Function

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
