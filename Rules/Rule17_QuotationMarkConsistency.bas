Attribute VB_Name = "Rule17_QuotationMarkConsistency"
' ============================================================
' Rule17_QuotationMarkConsistency.bas
' Proofreading rule: checks for mixed usage of straight and
' curly quotation marks (both single and double).
'
' Distinguishes apostrophes from single quotation marks by
' checking if the character is mid-word (preceded AND followed
' by a letter).
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "quotation_mark_consistency"

' Quotation mark character constants
Private Const STRAIGHT_DOUBLE As Long = 34        ' Chr(34) "
Private Const CURLY_DOUBLE_OPEN As Long = 8220     ' ChrW(8220)
Private Const CURLY_DOUBLE_CLOSE As Long = 8221    ' ChrW(8221)
Private Const STRAIGHT_SINGLE As Long = 39         ' Chr(39) '
Private Const CURLY_SINGLE_OPEN As Long = 8216     ' ChrW(8216)
Private Const CURLY_SINGLE_CLOSE As Long = 8217    ' ChrW(8217)

' Cached document boundary — set once per run
Private m_docEnd As Long

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_QuotationMarkConsistency(doc As Document) As Collection
    Dim issues As New Collection
    Dim docText As String
    Dim textLen As Long
    Dim i As Long
    Dim ch As String
    Dim charCode As Long

    ' Counters
    Dim straightDoubleCount As Long
    Dim curlyDoubleCount As Long
    Dim straightSingleCount As Long
    Dim curlySingleCount As Long

    straightDoubleCount = 0
    curlyDoubleCount = 0
    straightSingleCount = 0
    curlySingleCount = 0

    ' ── Get full document text ───────────────────────────────
    On Error Resume Next
    docText = doc.Content.Text
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set Check_QuotationMarkConsistency = issues
        Exit Function
    End If
    On Error GoTo 0

    textLen = Len(docText)
    If textLen = 0 Then
        Set Check_QuotationMarkConsistency = issues
        Exit Function
    End If

    ' ── Cache document boundary once ─────────────────────────
    m_docEnd = doc.Content.End

    ' ── First pass: count quotation mark types ───────────────
    For i = 1 To textLen
        charCode = AscW(Mid(docText, i, 1))

        Select Case charCode
            Case STRAIGHT_DOUBLE
                straightDoubleCount = straightDoubleCount + 1

            Case CURLY_DOUBLE_OPEN, CURLY_DOUBLE_CLOSE
                curlyDoubleCount = curlyDoubleCount + 1

            Case STRAIGHT_SINGLE
                ' Check if it is an apostrophe (mid-word)
                If Not IsApostrophe(docText, i, textLen) Then
                    straightSingleCount = straightSingleCount + 1
                End If

            Case CURLY_SINGLE_OPEN
                curlySingleCount = curlySingleCount + 1

            Case CURLY_SINGLE_CLOSE
                ' Check if it is an apostrophe (mid-word)
                If Not IsApostrophe(docText, i, textLen) Then
                    curlySingleCount = curlySingleCount + 1
                End If
        End Select
    Next i

    ' ── Determine dominant styles ────────────────────────────
    Dim dominantDouble As String ' "straight" or "curly"
    Dim dominantSingle As String ' "straight" or "curly"

    If straightDoubleCount >= curlyDoubleCount Then
        dominantDouble = "straight"
    Else
        dominantDouble = "curly"
    End If

    If straightSingleCount >= curlySingleCount Then
        dominantSingle = "straight"
    Else
        dominantSingle = "curly"
    End If

    ' ── Flag minority double quotation marks ─────────────────
    If dominantDouble = "straight" And curlyDoubleCount > 0 Then
        ' Flag curly doubles
        FlagQuotationMarks doc, issues, ChrW(CURLY_DOUBLE_OPEN), _
            "Curly double quotation mark found; document predominantly uses straight", _
            "Change to straight double quotation mark (" & Chr(STRAIGHT_DOUBLE) & ")"
        FlagQuotationMarks doc, issues, ChrW(CURLY_DOUBLE_CLOSE), _
            "Curly double quotation mark found; document predominantly uses straight", _
            "Change to straight double quotation mark (" & Chr(STRAIGHT_DOUBLE) & ")"
    ElseIf dominantDouble = "curly" And straightDoubleCount > 0 Then
        ' Flag straight doubles
        FlagQuotationMarks doc, issues, Chr(STRAIGHT_DOUBLE), _
            "Straight double quotation mark found; document predominantly uses curly", _
            "Change to curly double quotation marks (" & ChrW(CURLY_DOUBLE_OPEN) & _
            ChrW(CURLY_DOUBLE_CLOSE) & ")"
    End If

    ' ── Flag minority single quotation marks ─────────────────
    If dominantSingle = "straight" And curlySingleCount > 0 Then
        ' Flag curly singles (excluding apostrophes)
        FlagSingleQuotationMarks doc, issues, ChrW(CURLY_SINGLE_OPEN), _
            "Curly single quotation mark found; document predominantly uses straight", _
            "Change to straight single quotation mark (" & Chr(STRAIGHT_SINGLE) & ")", _
            False
        FlagSingleQuotationMarks doc, issues, ChrW(CURLY_SINGLE_CLOSE), _
            "Curly single quotation mark found; document predominantly uses straight", _
            "Change to straight single quotation mark (" & Chr(STRAIGHT_SINGLE) & ")", _
            True
    ElseIf dominantSingle = "curly" And straightSingleCount > 0 Then
        ' Flag straight singles (excluding apostrophes)
        FlagSingleQuotationMarks doc, issues, Chr(STRAIGHT_SINGLE), _
            "Straight single quotation mark found; document predominantly uses curly", _
            "Change to curly single quotation marks (" & ChrW(CURLY_SINGLE_OPEN) & _
            ChrW(CURLY_SINGLE_CLOSE) & ")", _
            True
    End If

    Set Check_QuotationMarkConsistency = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a character at position is an apostrophe
'  (preceded AND followed by a letter = mid-word)
' ════════════════════════════════════════════════════════════
Private Function IsApostrophe(ByRef docText As String, _
                               ByVal pos As Long, _
                               ByVal textLen As Long) As Boolean
    Dim prevChar As String
    Dim nextChar As String

    IsApostrophe = False

    ' Check character before
    If pos <= 1 Then Exit Function
    prevChar = Mid(docText, pos - 1, 1)
    If Not IsLetterChar(prevChar) Then Exit Function

    ' Check character after
    If pos >= textLen Then Exit Function
    nextChar = Mid(docText, pos + 1, 1)
    If Not IsLetterChar(nextChar) Then Exit Function

    ' Both sides are letters — this is an apostrophe
    IsApostrophe = True
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a character is a letter (A-Z, a-z)
' ════════════════════════════════════════════════════════════
Private Function IsLetterChar(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsLetterChar = (code >= 65 And code <= 90) Or _
                   (code >= 97 And code <= 122) Or _
                   (code >= 192 And code <= 687) ' Extended Latin
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Flag double quotation marks using Find
' ════════════════════════════════════════════════════════════
Private Sub FlagQuotationMarks(doc As Document, _
                                ByRef issues As Collection, _
                                ByVal searchChar As String, _
                                ByVal issueText As String, _
                                ByVal suggestion As String)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
    Dim locStr As String

    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = searchChar
        .MatchWildcards = False
        .MatchCase = True
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

        If Not PleadingsEngine.IsInPageRange(rng) Then GoTo ContinueFlag

        locStr = PleadingsEngine.GetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set issue = New PleadingsIssue
        issue.Init RULE_NAME, _
                   locStr, _
                   issueText, _
                   suggestion, _
                   rng.Start, _
                   rng.End, _
                   "possible_error"
        issues.Add issue

ContinueFlag:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Flag single quotation marks, skipping apostrophes
' ════════════════════════════════════════════════════════════
Private Sub FlagSingleQuotationMarks(doc As Document, _
                                      ByRef issues As Collection, _
                                      ByVal searchChar As String, _
                                      ByVal issueText As String, _
                                      ByVal suggestion As String, _
                                      ByVal checkApostrophe As Boolean)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
    Dim locStr As String

    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = searchChar
        .MatchWildcards = False
        .MatchCase = True
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

        If Not PleadingsEngine.IsInPageRange(rng) Then GoTo ContinueSingle

        ' Skip apostrophes: check if preceded AND followed by a letter.
        ' Uses a single 3-char range read (1 COM call) instead of
        ' two per-character Range objects (2 COM calls per match).
        If checkApostrophe Then
            Dim isApost As Boolean
            isApost = False

            If rng.Start > 0 And rng.End < m_docEnd Then
                Dim ctxRng As Range
                Set ctxRng = doc.Range(rng.Start - 1, rng.End + 1)
                If Err.Number = 0 Then
                    Dim ctxText As String
                    ctxText = ctxRng.Text
                    If Len(ctxText) >= 3 Then
                        If IsLetterChar(Left$(ctxText, 1)) And _
                           IsLetterChar(Right$(ctxText, 1)) Then
                            isApost = True
                        End If
                    End If
                End If
                If Err.Number <> 0 Then Err.Clear
            End If

            If isApost Then GoTo ContinueSingle
        End If

        locStr = PleadingsEngine.GetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set issue = New PleadingsIssue
        issue.Init RULE_NAME, _
                   locStr, _
                   issueText, _
                   suggestion, _
                   rng.Start, _
                   rng.End, _
                   "possible_error"
        issues.Add issue

ContinueSingle:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunQuotationMarkConsistency()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Quotation Mark Consistency"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_QuotationMarkConsistency(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Quotation Mark Consistency"
End Sub
