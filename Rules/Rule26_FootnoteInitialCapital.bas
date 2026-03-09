Attribute VB_Name = "Rule26_FootnoteInitialCapital"
' ============================================================
' Rule26_FootnoteInitialCapital.bas
' Proofreading rule: footnotes should begin with a capital
' letter, except for approved lower-case abbreviations.
'
' Approved lower-case starts:
'   c, cf, cp, eg, ie, p, pp, ibid
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "footnote_initial_capital"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_FootnoteInitialCapital(doc As Document) As Collection
    Dim issues As New Collection
    Dim allowed As Object
    Dim fn As Footnote
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim noteText As String
    Dim trimmed As String
    Dim token As String
    Dim firstCharCode As Long
    Dim i As Long
    Dim j As Long
    Dim ch As String

    ' ── Build allowed lower-case starts dictionary ───────────
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

        ' ── Trim leading whitespace ──────────────────────────
        trimmed = LTrim(noteText)
        If Len(trimmed) = 0 Then GoTo NextFootnote

        ' ── Skip past leading punctuation (quotes, brackets) ─
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

        ' ── Extract first lexical token (letters only) ───────
        token = ExtractFirstToken(trimmed)
        If Len(token) = 0 Then GoTo NextFootnote

        ' ── Check if token is in allowed list ────────────────
        If allowed.Exists(LCase(token)) Then GoTo NextFootnote

        ' ── Check if first character is lower-case ───────────
        firstCharCode = AscW(Mid(token, 1, 1))
        If firstCharCode >= 97 And firstCharCode <= 122 Then
            ' Lower-case and not in allowed list: flag
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       "Footnote begins with lower-case text outside the approved exceptions.", _
                       "Begin the footnote with a capital letter, unless it starts with an approved lower-case abbreviation.", _
                       fn.Range.Start, _
                       fn.Range.End, _
                       "warning", _
                       False
            issues.Add issue
        End If

NextFootnote:
    Next i

    Set Check_FootnoteInitialCapital = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if character is leading punctuation to skip
' ════════════════════════════════════════════════════════════
Private Function IsLeadingPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case "(", "[", ChrW(8216), ChrW(8220), """", "'"
            IsLeadingPunctuation = True
        Case Else
            IsLeadingPunctuation = False
    End Select
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Extract the first token of letters from a string
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunFootnoteInitialCapital()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Footnote Initial Capital"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_FootnoteInitialCapital(doc)

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
           vbInformation, "Footnote Initial Capital"
End Sub
