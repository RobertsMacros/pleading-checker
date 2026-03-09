Attribute VB_Name = "Rule25_FootnoteTerminalFullStop"
' ============================================================
' Rule25_FootnoteTerminalFullStop.bas
' Proofreading rule: every footnote should end with a full stop.
'
' Handles trailing whitespace/paragraph marks and closing
' brackets/quotes (checks character before them for ".").
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "footnote_terminal_full_stop"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_FootnoteTerminalFullStop(doc As Document) As Collection
    Dim issues As New Collection
    Dim fn As Footnote
    Dim issue As PleadingsIssue
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

        ' ── Check page range on the reference mark ───────────
        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFootnote
        End If
        On Error GoTo 0

        ' ── Get footnote text ────────────────────────────────
        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote
        End If
        On Error GoTo 0

        ' ── Trim trailing whitespace / paragraph marks ───────
        trimmed = noteText
        trimmed = TrimTrailingWhitespace(trimmed)

        ' ── Skip empty footnotes ─────────────────────────────
        If Len(trimmed) = 0 Then GoTo NextFootnote

        ' ── Get last character ───────────────────────────────
        lastChar = Mid(trimmed, Len(trimmed), 1)

        ' ── If last char is closing bracket/quote, check penultimate ──
        If IsClosingPunctuation(lastChar) Then
            If Len(trimmed) >= 2 Then
                penultChar = Mid(trimmed, Len(trimmed) - 1, 1)
                If penultChar = "." Then GoTo NextFootnote
            End If
            ' Fall through to flag
        ElseIf lastChar = "." Then
            GoTo NextFootnote
        End If

        ' ── Flag missing full stop ───────────────────────────
        On Error Resume Next
        locStr = PleadingsEngine.GetLocationString(fn.Reference, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
        On Error GoTo 0

        Set issue = New PleadingsIssue
        issue.Init RULE_NAME, _
                   locStr, _
                   "Footnote does not end with a full stop.", _
                   "Add a full stop at the end of the footnote.", _
                   fn.Range.Start, _
                   fn.Range.End, _
                   "warning", _
                   True
        issues.Add issue

NextFootnote:
    Next i

    Set Check_FootnoteTerminalFullStop = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Strip trailing CR, LF, VT, and spaces
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if character is a closing bracket or quote
' ════════════════════════════════════════════════════════════
Private Function IsClosingPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case ")", "]", ChrW(8217), ChrW(8221)
            IsClosingPunctuation = True
        Case Else
            IsClosingPunctuation = False
    End Select
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunFootnoteTerminalFullStop()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Footnote Terminal Full Stop"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_FootnoteTerminalFullStop(doc)

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
           vbInformation, "Footnote Terminal Full Stop"
End Sub
