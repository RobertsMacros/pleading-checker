Attribute VB_Name = "Rule33_SmartQuoteConsistency"
' ============================================================
' Rule33_SmartQuoteConsistency.bas
' Proofreading rule: detects inconsistent use of straight and
' curly quotation marks across the document.
'
' Distinguishes apostrophes from single quotation marks by
' checking if the character is mid-word (preceded AND followed
' by a letter).
'
' When mixed styles are found, prefers curly as the dominant
' style and flags straight quotes as the minority.
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "smart_quote_consistency"

' Quotation mark character constants
Private Const STRAIGHT_DOUBLE As Long = 34         ' Chr(34)
Private Const CURLY_DOUBLE_OPEN As Long = 8220     ' ChrW(8220)
Private Const CURLY_DOUBLE_CLOSE As Long = 8221    ' ChrW(8221)
Private Const STRAIGHT_SINGLE As Long = 39         ' Chr(39)
Private Const CURLY_SINGLE_OPEN As Long = 8216     ' ChrW(8216)
Private Const CURLY_SINGLE_CLOSE As Long = 8217    ' ChrW(8217)

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_SmartQuoteConsistency(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim i As Long
    Dim charCode As Long
    Dim textLen As Long

    ' Counters for straight vs curly
    Dim straightCount As Long
    Dim curlyCount As Long
    straightCount = 0
    curlyCount = 0

    ' ── First pass: count straight vs curly quotes ─────────
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass1
        End If

        ' Skip paragraphs outside the configured page range
        If Not PleadingsEngine.IsInPageRange(paraRange) Then
            GoTo NextParaPass1
        End If

        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass1
        End If

        textLen = Len(paraText)
        If textLen = 0 Then GoTo NextParaPass1

        For i = 1 To textLen
            charCode = AscW(Mid(paraText, i, 1))

            Select Case charCode
                Case STRAIGHT_DOUBLE
                    straightCount = straightCount + 1

                Case CURLY_DOUBLE_OPEN, CURLY_DOUBLE_CLOSE
                    curlyCount = curlyCount + 1

                Case STRAIGHT_SINGLE
                    ' Only count if not an apostrophe
                    If Not IsApostrophe(paraText, i, textLen) Then
                        straightCount = straightCount + 1
                    End If

                Case CURLY_SINGLE_OPEN
                    curlyCount = curlyCount + 1

                Case CURLY_SINGLE_CLOSE
                    ' Only count if not an apostrophe
                    If Not IsApostrophe(paraText, i, textLen) Then
                        curlyCount = curlyCount + 1
                    End If
            End Select
        Next i

NextParaPass1:
    Next para
    On Error GoTo 0

    ' ── Determine if there is a mix ────────────────────────
    ' If only one style or no quotes at all, no issue
    If straightCount = 0 Or curlyCount = 0 Then
        Set Check_SmartQuoteConsistency = issues
        Exit Function
    End If

    ' Per spec: prefer curly as dominant when both exist
    ' Emit document-level summary issue
    Dim summaryIssue As PleadingsIssue
    Set summaryIssue = New PleadingsIssue
    summaryIssue.Init RULE_NAME, _
                       "Document", _
                       "Quotation mark style is inconsistent. Found " & _
                       straightCount & " straight and " & curlyCount & _
                       " curly quotation marks.", _
                       "Use curly quotation marks consistently throughout the document.", _
                       0, _
                       0, _
                       "warning", _
                       False
    issues.Add summaryIssue

    ' ── Second pass: flag each straight quote occurrence ───
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass2
        End If

        ' Skip paragraphs outside the configured page range
        If Not PleadingsEngine.IsInPageRange(paraRange) Then
            GoTo NextParaPass2
        End If

        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass2
        End If

        textLen = Len(paraText)
        If textLen = 0 Then GoTo NextParaPass2

        For i = 1 To textLen
            charCode = AscW(Mid(paraText, i, 1))

            Dim isStraightQuote As Boolean
            isStraightQuote = False

            Select Case charCode
                Case STRAIGHT_DOUBLE
                    isStraightQuote = True

                Case STRAIGHT_SINGLE
                    ' Only flag if not an apostrophe
                    If Not IsApostrophe(paraText, i, textLen) Then
                        isStraightQuote = True
                    End If
            End Select

            If isStraightQuote Then
                Dim rangeStart As Long
                Dim rangeEnd As Long
                Dim locStr As String
                Dim charRange As Range
                Dim issue As PleadingsIssue

                rangeStart = paraRange.Start + i - 1
                rangeEnd = rangeStart + 1

                Err.Clear
                Set charRange = doc.Range(rangeStart, rangeEnd)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                Else
                    locStr = PleadingsEngine.GetLocationString(charRange, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If
                End If

                Set issue = New PleadingsIssue
                issue.Init RULE_NAME, _
                           locStr, _
                           "Straight quotation mark found in otherwise curly-quoted document.", _
                           "Replace with curly quotation mark.", _
                           rangeStart, _
                           rangeEnd, _
                           "warning", _
                           False
                issues.Add issue
            End If
        Next i

NextParaPass2:
    Next para
    On Error GoTo 0

    Set Check_SmartQuoteConsistency = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a character at position is an apostrophe
'  (preceded AND followed by a letter = mid-word)
' ════════════════════════════════════════════════════════════
Private Function IsApostrophe(ByRef txt As String, _
                               ByVal pos As Long, _
                               ByVal textLen As Long) As Boolean
    Dim prevChar As String
    Dim nextChar As String

    IsApostrophe = False

    ' Check character before
    If pos <= 1 Then Exit Function
    prevChar = Mid(txt, pos - 1, 1)
    If Not IsLetterChar(prevChar) Then Exit Function

    ' Check character after
    If pos >= textLen Then Exit Function
    nextChar = Mid(txt, pos + 1, 1)
    If Not IsLetterChar(nextChar) Then Exit Function

    ' Both sides are letters -- this is an apostrophe
    IsApostrophe = True
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a character is a letter (A-Z, a-z,
'  extended Latin)
' ════════════════════════════════════════════════════════════
Private Function IsLetterChar(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsLetterChar = (code >= 65 And code <= 90) Or _
                   (code >= 97 And code <= 122) Or _
                   (code >= 192 And code <= 687) ' Extended Latin
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunSmartQuoteConsistency()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Smart Quote Consistency"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_SmartQuoteConsistency(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Smart Quote Consistency"
End Sub
