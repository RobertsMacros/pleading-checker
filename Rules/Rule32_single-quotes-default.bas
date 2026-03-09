Attribute VB_Name = "Rule32_single_quotes_default"
' ============================================================
' Rule32_single-quotes-default.bas
' Proofreading rule: requires single quotation marks as the
' default outer quotation marks in ordinary prose.
'
' Excludes block quotations, code/data regions by checking
' paragraph style names for "Block", "Quote", or "Code".
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "single_quotes_default"

' Double quotation mark character constants
Private Const CURLY_DOUBLE_OPEN As Long = 8220    ' ChrW(8220)
Private Const CURLY_DOUBLE_CLOSE As Long = 8221   ' ChrW(8221)
Private Const STRAIGHT_DOUBLE As Long = 34         ' Chr(34)

' Single quotation mark character constants (for nesting check)
Private Const CURLY_SINGLE_OPEN As Long = 8216     ' ChrW(8216)
Private Const CURLY_SINGLE_CLOSE As Long = 8217    ' ChrW(8217)
Private Const STRAIGHT_SINGLE As Long = 39         ' Chr(39)

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_SingleQuotesDefault(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim styleName As String
    Dim i As Long
    Dim charCode As Long
    Dim singleDepth As Long
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim charRange As Range

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph
        End If

        ' Skip paragraphs outside the configured page range
        If Not PleadingsEngine.IsInPageRange(paraRange) Then
            GoTo NextParagraph
        End If

        ' ── Check paragraph style for exclusions ────────────
        styleName = ""
        styleName = paraRange.ParagraphStyle
        If Err.Number <> 0 Then
            Err.Clear
            styleName = ""
        End If

        If IsExcludedStyle(styleName) Then
            GoTo NextParagraph
        End If

        ' ── Get paragraph text ──────────────────────────────
        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph
        End If

        If Len(paraText) = 0 Then
            GoTo NextParagraph
        End If

        ' ── Scan for double quotation marks ─────────────────
        ' Track single-quote nesting depth to determine if a
        ' double quote is "outer" (depth = 0) or "inner" (depth > 0).
        singleDepth = 0

        For i = 1 To Len(paraText)
            charCode = AscW(Mid(paraText, i, 1))

            Select Case charCode
                ' ── Single quote openers ────────────────────
                Case CURLY_SINGLE_OPEN
                    singleDepth = singleDepth + 1

                Case STRAIGHT_SINGLE
                    ' Straight single: treat as opener if not apostrophe
                    If Not IsApostrophe(paraText, i, Len(paraText)) Then
                        ' Toggle: if depth is 0 treat as open, else close
                        If singleDepth = 0 Then
                            singleDepth = singleDepth + 1
                        Else
                            singleDepth = singleDepth - 1
                            If singleDepth < 0 Then singleDepth = 0
                        End If
                    End If

                ' ── Single quote closers ────────────────────
                Case CURLY_SINGLE_CLOSE
                    ' Only treat as closer if not an apostrophe
                    If Not IsApostrophe(paraText, i, Len(paraText)) Then
                        singleDepth = singleDepth - 1
                        If singleDepth < 0 Then singleDepth = 0
                    End If

                ' ── Double quote characters ─────────────────
                Case CURLY_DOUBLE_OPEN, CURLY_DOUBLE_CLOSE, STRAIGHT_DOUBLE
                    ' Flag all double quotes in non-excluded paragraphs.
                    ' The rule mandates single quotes as default outer marks.
                    Dim rangeStart As Long
                    Dim rangeEnd As Long

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
                               "Outer quotation marks should use single quotation marks.", _
                               "Use single quotation marks instead of double quotation marks.", _
                               rangeStart, _
                               rangeEnd, _
                               "warning", _
                               False
                    issues.Add issue
            End Select
        Next i

NextParagraph:
    Next para
    On Error GoTo 0

    Set Check_SingleQuotesDefault = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a paragraph style should be excluded
'  Excludes styles containing "Block", "Quote", or "Code"
' ════════════════════════════════════════════════════════════
Private Function IsExcludedStyle(ByVal styleName As String) As Boolean
    Dim lStyle As String
    lStyle = LCase(styleName)

    IsExcludedStyle = (InStr(lStyle, "block") > 0) Or _
                      (InStr(lStyle, "quote") > 0) Or _
                      (InStr(lStyle, "code") > 0)
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
