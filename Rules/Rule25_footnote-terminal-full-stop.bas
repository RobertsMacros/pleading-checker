Attribute VB_Name = "Rule25_footnote_terminal_full_stop"
' ============================================================
' Rule25_footnote-terminal-full-stop.bas
' Proofreading rule: every footnote should end with a full stop.
'
' Handles trailing whitespace/paragraph marks and closing
' brackets/quotes (checks character before them for ".").
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "footnote_terminal_full_stop"

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_FootnoteTerminalFullStop(doc As Document) As Collection
    Dim issues As New Collection
    Dim fn As Footnote
    Dim issue As Object
    Dim locStr As String
    Dim noteText As String
    Dim trimmed As String
    Dim lastChar As String
    Dim penultChar As String
    Dim i As Long

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

        ' -- Trim trailing whitespace / paragraph marks -------
        trimmed = noteText
        trimmed = TrimTrailingWhitespace(trimmed)

        ' -- Skip empty footnotes -----------------------------
        If Len(trimmed) = 0 Then GoTo NextFootnote

        ' -- Get last character -------------------------------
        lastChar = Mid(trimmed, Len(trimmed), 1)

        ' -- If last char is closing bracket/quote, check penultimate --
        If IsClosingPunctuation(lastChar) Then
            If Len(trimmed) >= 2 Then
                penultChar = Mid(trimmed, Len(trimmed) - 1, 1)
                If penultChar = "." Then GoTo NextFootnote
            End If
            ' Fall through to flag
        ElseIf lastChar = "." Then
            GoTo NextFootnote
        End If

        ' -- Flag missing full stop ---------------------------
        On Error Resume Next
        locStr = EngineGetLocationString(fn.Reference, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
        On Error GoTo 0

        Set issue = CreateIssueDict(RULE_NAME, locStr, "Footnote does not end with a full stop.", "Add a full stop at the end of the footnote.", fn.Range.Start, fn.Range.End, "warning", True)
        issues.Add issue

NextFootnote:
    Next i

    Set Check_FootnoteTerminalFullStop = issues
End Function

' ============================================================
'  PRIVATE: Strip trailing CR, LF, VT, and spaces
' ============================================================
Private Function TrimTrailingWhitespace(ByVal s As String) As String
    Dim ch As String
    Do While Len(s) > 0
        ch = Mid(s, Len(s), 1)
        Select Case ch
            Case vbCr, vbLf, Chr(13), Chr(10), Chr(11), " ", vbTab
                s = Left(s, Len(s) - 1)
            Case Else
                Exit Do
        End Select
    Loop
    TrimTrailingWhitespace = s
End Function

' ============================================================
'  PRIVATE: Check if character is a closing bracket or quote
' ============================================================
Private Function IsClosingPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case ")", "]", ChrW(8217), ChrW(8221)
            IsClosingPunctuation = True
        Case Else
            IsClosingPunctuation = False
    End Select
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
