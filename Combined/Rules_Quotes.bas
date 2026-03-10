Attribute VB_Name = "Rules_Quotes"
' ============================================================
' Rules_Quotes.bas
' Combined proofreading rules for quotation marks:
'   - Rule17: quotation mark consistency (straight vs curly)
'   - Rule32: single quotes as default outer marks
'   - Rule33: smart quote consistency (prefers curly)
'
' Shared helpers (IsApostrophe, IsLetterChar) are defined once.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' -- Rule name constants (one per rule) ------------------------
Private Const RULE_NAME_17 As String = "quotation_mark_consistency"
Private Const RULE_NAME_32 As String = "single_quotes_default"
Private Const RULE_NAME_33 As String = "smart_quote_consistency"

' -- Quotation mark character constants ------------------------
Private Const STRAIGHT_DOUBLE As Long = 34        ' Chr(34) "
Private Const CURLY_DOUBLE_OPEN As Long = 8220     ' ChrW(8220)
Private Const CURLY_DOUBLE_CLOSE As Long = 8221    ' ChrW(8221)
Private Const STRAIGHT_SINGLE As Long = 39         ' Chr(39) '
Private Const CURLY_SINGLE_OPEN As Long = 8216     ' ChrW(8216)
Private Const CURLY_SINGLE_CLOSE As Long = 8217    ' ChrW(8217)

' ================================================================
' ================================================================
'  RULE 17 - QUOTATION MARK CONSISTENCY
' ================================================================
' ================================================================

' ============================================================
'  MAIN ENTRY POINT (Rule 17)
' ============================================================
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

    ' -- Get full document text -------------------------------
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

    ' -- First pass: count quotation mark types ---------------
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

    ' -- Determine dominant styles ----------------------------
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

    ' -- Flag minority double quotation marks -----------------
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

    ' -- Flag minority single quotation marks -----------------
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

' ================================================================
' ================================================================
'  RULE 32 - SINGLE QUOTES DEFAULT
' ================================================================
' ================================================================

' ============================================================
'  MAIN ENTRY POINT (Rule 32)
' ============================================================
Public Function Check_SingleQuotesDefault(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim styleName As String
    Dim i As Long
    Dim charCode As Long
    Dim singleDepth As Long
    Dim finding As Object
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
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextParagraph
        End If

        ' -- Check paragraph style for exclusions ------------
        styleName = ""
        styleName = paraRange.ParagraphStyle
        If Err.Number <> 0 Then
            Err.Clear
            styleName = ""
        End If

        If IsExcludedStyle(styleName) Then
            GoTo NextParagraph
        End If

        ' -- Get paragraph text ------------------------------
        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph
        End If

        If Len(paraText) = 0 Then
            GoTo NextParagraph
        End If

        ' -- Scan for double quotation marks -----------------
        ' Track single-quote nesting depth to determine if a
        ' double quote is "outer" (depth = 0) or "inner" (depth > 0).
        singleDepth = 0

        For i = 1 To Len(paraText)
            charCode = AscW(Mid(paraText, i, 1))

            Select Case charCode
                ' -- Single quote openers --------------------
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

                ' -- Single quote closers --------------------
                Case CURLY_SINGLE_CLOSE
                    ' Only treat as closer if not an apostrophe
                    If Not IsApostrophe(paraText, i, Len(paraText)) Then
                        singleDepth = singleDepth - 1
                        If singleDepth < 0 Then singleDepth = 0
                    End If

                ' -- Double quote characters -----------------
                Case CURLY_DOUBLE_OPEN, CURLY_DOUBLE_CLOSE, STRAIGHT_DOUBLE
                    ' Flag all double quotes in non-excluded paragraphs.
                    ' The rule mandates single quotes as default outer marks.
                    Dim rangeStart32 As Long
                    Dim rangeEnd32 As Long

                    rangeStart32 = paraRange.Start + i - 1
                    rangeEnd32 = rangeStart32 + 1

                    Err.Clear
                    Set charRange = doc.Range(rangeStart32, rangeEnd32)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    Else
                        locStr = EngineGetLocationString(charRange, doc)
                        If Err.Number <> 0 Then
                            locStr = "unknown location"
                            Err.Clear
                        End If
                    End If

                    Set finding = CreateIssueDict(RULE_NAME_32, locStr, "Outer quotation marks should use single quotation marks.", "Use single quotation marks instead of double quotation marks.", rangeStart32, rangeEnd32, "warning", False)
                    issues.Add finding
            End Select
        Next i

NextParagraph:
    Next para
    On Error GoTo 0

    Set Check_SingleQuotesDefault = issues
End Function

' ================================================================
' ================================================================
'  RULE 33 - SMART QUOTE CONSISTENCY
' ================================================================
' ================================================================

' ============================================================
'  MAIN ENTRY POINT (Rule 33)
' ============================================================
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

    ' -- First pass: count straight vs curly quotes ---------
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass1
        End If

        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(paraRange) Then
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

    ' -- Determine if there is a mix ------------------------
    ' If only one style or no quotes at all, no finding
    If straightCount = 0 Or curlyCount = 0 Then
        Set Check_SmartQuoteConsistency = issues
        Exit Function
    End If

    ' Per spec: prefer curly as dominant when both exist
    ' Emit document-level summary finding
    Dim summaryFinding As Object
    Set summaryFinding = CreateIssueDict(RULE_NAME_33, "Document", "Quotation mark style is inconsistent. Found " & straightCount & " straight and " & curlyCount & " curly quotation marks.", "Use curly quotation marks consistently throughout the document.", 0, 0, "warning", False)
    issues.Add summaryFinding

    ' -- Second pass: flag each straight quote occurrence ---
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass2
        End If

        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(paraRange) Then
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
                Dim rangeStart33 As Long
                Dim rangeEnd33 As Long
                Dim locStr33 As String
                Dim charRange33 As Range
                Dim finding33 As Object

                rangeStart33 = paraRange.Start + i - 1
                rangeEnd33 = rangeStart33 + 1

                Err.Clear
                Set charRange33 = doc.Range(rangeStart33, rangeEnd33)
                If Err.Number <> 0 Then
                    locStr33 = "unknown location"
                    Err.Clear
                Else
                    locStr33 = EngineGetLocationString(charRange33, doc)
                    If Err.Number <> 0 Then
                        locStr33 = "unknown location"
                        Err.Clear
                    End If
                End If

                Set finding33 = CreateIssueDict(RULE_NAME_33, locStr33, "Straight quotation mark found in otherwise curly-quoted document.", "Replace with curly quotation mark.", rangeStart33, rangeEnd33, "warning", False)
                issues.Add finding33
            End If
        Next i

NextParaPass2:
    Next para
    On Error GoTo 0

    Set Check_SmartQuoteConsistency = issues
End Function

' ================================================================
' ================================================================
'  SHARED PRIVATE HELPERS
' ================================================================
' ================================================================

' ============================================================
'  PRIVATE: Check if a character at position is an apostrophe
'  (preceded AND followed by a letter = mid-word)
'  Used by Rule17, Rule32, and Rule33
' ============================================================
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

' ============================================================
'  PRIVATE: Check if a character is a letter (A-Z, a-z,
'  extended Latin)
'  Used by Rule17, Rule32, and Rule33 (via IsApostrophe)
' ============================================================
Private Function IsLetterChar(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsLetterChar = (code >= 65 And code <= 90) Or _
                   (code >= 97 And code <= 122) Or _
                   (code >= 192 And code <= 687) ' Extended Latin
End Function

' ================================================================
' ================================================================
'  RULE 17 PRIVATE HELPERS
' ================================================================
' ================================================================

' ============================================================
'  PRIVATE: Flag double quotation marks using Find (Rule 17)
' ============================================================
Private Sub FlagQuotationMarks(doc As Document, _
                                ByRef issues As Collection, _
                                ByVal searchChar As String, _
                                ByVal issueText As String, _
                                ByVal suggestion As String)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
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

        If Not EngineIsInPageRange(rng) Then GoTo ContinueFlag

        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set finding = CreateIssueDict(RULE_NAME_17, locStr, issueText, suggestion, rng.Start, rng.End, "possible_error")
        issues.Add finding

ContinueFlag:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Flag single quotation marks, skipping apostrophes
'  (Rule 17)
' ============================================================
Private Sub FlagSingleQuotationMarks(doc As Document, _
                                      ByRef issues As Collection, _
                                      ByVal searchChar As String, _
                                      ByVal issueText As String, _
                                      ByVal suggestion As String, _
                                      ByVal checkApostrophe As Boolean)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
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

        If Not EngineIsInPageRange(rng) Then GoTo ContinueSingle

        ' Skip apostrophes: check if preceded AND followed by a letter
        If checkApostrophe Then
            Dim prevChar As String
            Dim nextChar As String
            Dim isApost As Boolean
            isApost = False

            If rng.Start > 0 Then
                Dim prevRng As Range
                Set prevRng = doc.Range(rng.Start - 1, rng.Start)
                If Err.Number = 0 Then
                    prevChar = prevRng.Text
                    If IsLetterChar(prevChar) Then
                        Dim nextRng As Range
                        If rng.End < doc.Content.End Then
                            Set nextRng = doc.Range(rng.End, rng.End + 1)
                            If Err.Number = 0 Then
                                nextChar = nextRng.Text
                                If IsLetterChar(nextChar) Then
                                    isApost = True
                                End If
                            End If
                            If Err.Number <> 0 Then Err.Clear
                        End If
                    End If
                End If
                If Err.Number <> 0 Then Err.Clear
            End If

            If isApost Then GoTo ContinueSingle
        End If

        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set finding = CreateIssueDict(RULE_NAME_17, locStr, issueText, suggestion, rng.Start, rng.End, "possible_error")
        issues.Add finding

ContinueSingle:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Check if a paragraph style should be excluded
'  Excludes styles containing "Block", "Quote", or "Code"
'  (Rule 32)
' ============================================================
Private Function IsExcludedStyle(ByVal styleName As String) As Boolean
    Dim lStyle As String
    lStyle = LCase(styleName)

    IsExcludedStyle = (InStr(lStyle, "block") > 0) Or _
                      (InStr(lStyle, "quote") > 0) Or _
                      (InStr(lStyle, "code") > 0)
End Function

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineIsInPageRange
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineGetLocationString
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
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
