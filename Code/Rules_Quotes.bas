Attribute VB_Name = "Rules_Quotes"
' ============================================================
' Rules_Quotes.bas
' Quotation-mark rules for UK legal proofreading:
'   Rule 17: quotation mark consistency (straight vs smart)
'   Rule 32: single quotes as default outer marks (nesting-aware)
'   Rule 33: smart quote consistency (prefers smart)
'
' Performance notes:
'   - All character scanning uses byte arrays, not Mid$/AscW
'   - Apostrophe detection is inlined on the byte data
'   - Rule 17 collects positions in one pass, flags from the array
'   - Rule 33 merged from two paragraph passes into one
'   - Location ranges are reused via SetRange, not re-created
'
' Dependency: PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' -- Rule name constants ----------------------------------------
Private Const RULE17 As String = "quotation_mark_consistency"
Private Const RULE32 As String = "single_quotes_default"
Private Const RULE33 As String = "smart_quote_consistency"

' -- Unicode code points ----------------------------------------
Private Const QD  As Long = 34     ' straight double  "
Private Const QDO As Long = 8220   ' smart double open
Private Const QDC As Long = 8221   ' smart double close
Private Const QS  As Long = 39     ' straight single  '
Private Const QSO As Long = 8216   ' smart single open
Private Const QSC As Long = 8217   ' smart single close

' ================================================================
'  RULE 17 -- QUOTATION MARK CONSISTENCY
'
'  One byte-array pass over doc.Content.Text to count + collect
'  positions of every quote type.  Determines dominant style for
'  doubles and singles independently (ties -> straight).  Emits
'  findings for each minority occurrence within the page range.
' ================================================================
Public Function Check_QuotationMarkConsistency( _
        doc As Document) As Collection

    Dim issues As New Collection

    ' -- Grab full-document text once ---------------------------
    Dim docText As String
    On Error Resume Next
    docText = doc.Content.Text
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        Set Check_QuotationMarkConsistency = issues
        Exit Function
    End If
    On Error GoTo 0

    If LenB(docText) = 0 Then
        Set Check_QuotationMarkConsistency = issues
        Exit Function
    End If

    ' -- Convert to byte array for fast scanning ----------------
    '    VBA strings are UTF-16LE: two bytes per character.
    '    Byte(i) is low byte, Byte(i+1) is high byte.
    '    Character's document position = i \ 2.
    Dim b() As Byte
    b = docText
    Dim bMax As Long
    bMax = UBound(b) - 1   ' last even index

    ' -- Counters -----------------------------------------------
    Dim cSD As Long   ' straight double
    Dim cCD As Long   ' smart double
    Dim cSS As Long   ' straight single (excluding apostrophes)
    Dim cCS As Long   ' smart single   (excluding apostrophes)

    ' -- Position collectors (grow-on-demand) --------------------
    Dim pSD() As Long
    ReDim pSD(0 To 127)
    Dim pCD() As Long
    ReDim pCD(0 To 127)
    Dim pSS() As Long
    ReDim pSS(0 To 127)
    Dim pCS() As Long
    ReDim pCS(0 To 127)
    Dim capSD As Long
    capSD = 128
    Dim capCD As Long
    capCD = 128
    Dim capSS As Long
    capSS = 128
    Dim capCS As Long
    capCS = 128

    ' -- Single pass: count + collect positions ------------------
    Dim i As Long, code As Long

    For i = 0 To bMax Step 2
        code = b(i) Or (CLng(b(i + 1)) * 256&)

        Select Case code
        Case QD
            If cSD >= capSD Then
                capSD = capSD * 2
                ReDim Preserve pSD(0 To capSD - 1)
            End If
            pSD(cSD) = i \ 2: cSD = cSD + 1

        Case QDO, QDC
            If cCD >= capCD Then
                capCD = capCD * 2
                ReDim Preserve pCD(0 To capCD - 1)
            End If
            pCD(cCD) = i \ 2: cCD = cCD + 1

        Case QS
            If Not ByteIsApostrophe(b, i, bMax) Then
                If cSS >= capSS Then
                    capSS = capSS * 2
                    ReDim Preserve pSS(0 To capSS - 1)
                End If
                pSS(cSS) = i \ 2: cSS = cSS + 1
            End If

        Case QSO
            If cCS >= capCS Then
                capCS = capCS * 2
                ReDim Preserve pCS(0 To capCS - 1)
            End If
            pCS(cCS) = i \ 2: cCS = cCS + 1

        Case QSC
            If Not ByteIsApostrophe(b, i, bMax) Then
                If cCS >= capCS Then
                    capCS = capCS * 2
                    ReDim Preserve pCS(0 To capCS - 1)
                End If
                pCS(cCS) = i \ 2: cCS = cCS + 1
            End If
        End Select
    Next i

    ' -- Determine dominant styles (tie -> straight) ------------
    Dim dblStraight As Boolean
    dblStraight = (cSD >= cCD)
    Dim sglStraight As Boolean
    sglStraight = (cSS >= cCS)

    ' -- Flag minority doubles ----------------------------------
    If dblStraight And cCD > 0 Then
        EmitFromPositions doc, issues, pCD, cCD, RULE17, _
            "Smart double quotation mark found; " & _
            "document predominantly uses straight", _
            "Change to straight double quotation mark (" & _
            Chr$(QD) & ")"

    ElseIf (Not dblStraight) And cSD > 0 Then
        EmitFromPositions doc, issues, pSD, cSD, RULE17, _
            "Straight double quotation mark found; " & _
            "document predominantly uses smart", _
            "Change to smart double quotation marks (" & _
            ChrW$(QDO) & ChrW$(QDC) & ")"
    End If

    ' -- Flag minority singles ----------------------------------
    If sglStraight And cCS > 0 Then
        EmitFromPositions doc, issues, pCS, cCS, RULE17, _
            "Smart single quotation mark found; " & _
            "document predominantly uses straight", _
            "Change to straight single quotation mark (" & _
            Chr$(QS) & ")"

    ElseIf (Not sglStraight) And cSS > 0 Then
        EmitFromPositions doc, issues, pSS, cSS, RULE17, _
            "Straight single quotation mark found; " & _
            "document predominantly uses smart", _
            "Change to smart single quotation marks (" & _
            ChrW$(QSO) & ChrW$(QSC) & ")"
    End If

    Set Check_QuotationMarkConsistency = issues
End Function

' ================================================================
'  RULE 32 -- SINGLE / DOUBLE QUOTES DEFAULT (nesting-aware)
'
'  Proper nesting-aware scanner that classifies each non-apostrophe
'  quote character as outer-level or inner-level, then only flags
'  quotes that use the wrong type AT THEIR NESTING LEVEL.
'
'  UK convention (nestMode="SINGLE"):
'    Depth 0 (outer) should use single quotes.
'    Depth 1 (inner) should use double quotes.
'    Any deeper level alternates.
'
'  The scanner uses explicit nesting stacks per paragraph.
'  Apostrophes (letter-flanked ' chars) are always skipped.
'  Straight quotes toggle; smart quotes have open/close directionality.
' ================================================================
Public Function Check_SingleQuotesDefault( _
        doc As Document) As Collection

    Dim issues As New Collection
    Dim para As Paragraph
    Dim pRng As Range
    Dim pText As String
    Dim pStart As Long
    Dim styleName As String
    Dim b() As Byte
    Dim bMax As Long
    Dim i As Long, code As Long, pos As Long
    Dim locStr As String

    ' Determine which quote type to flag based on user preference
    Dim nestMode As String
    nestMode = EngineGetQuoteNesting()  ' "SINGLE" or "DOUBLE"

    ' Outer and inner quote codes based on nesting mode
    ' For SINGLE outer: even depths (0,2,4..) use single, odd depths use double
    ' For DOUBLE outer: even depths use double, odd depths use single
    Dim outerOpen As Long, outerClose As Long, outerStraight As Long
    Dim innerOpen As Long, innerClose As Long, innerStraight As Long
    If nestMode = "DOUBLE" Then
        outerOpen = QDO: outerClose = QDC: outerStraight = QD
        innerOpen = QSO: innerClose = QSC: innerStraight = QS
    Else
        outerOpen = QSO: outerClose = QSC: outerStraight = QS
        innerOpen = QDO: innerClose = QDC: innerStraight = QD
    End If

    ' Reusable range -- created once, repositioned via SetRange
    Dim locRng As Range
    Set locRng = doc.Range(0, 1)

    ' Nesting state persists across paragraphs for multi-paragraph quotes
    Dim nestDepth As Long
    nestDepth = 0
    ' Track what type each nesting level was opened with:
    ' True = outer-type, False = inner-type.  Max 10 levels.
    Dim levelIsOuter(0 To 9) As Boolean
    ' Count of nesting anomalies (underflow / overflow).
    ' If too many anomalies we suppress further flagging.
    Dim nestAnomalies As Long
    nestAnomalies = 0

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set pRng = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NxtP32

        ' Page-range gate (once per paragraph, not per character)
        If Not EngineIsInPageRange(pRng) Then GoTo NxtP32

        ' Style exclusion gate
        Err.Clear
        styleName = pRng.ParagraphStyle
        If Err.Number <> 0 Then styleName = "": Err.Clear
        If IsExcludedStyle(styleName) Then GoTo NxtP32

        ' Fetch paragraph text
        Err.Clear
        pText = pRng.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NxtP32
        If LenB(pText) = 0 Then GoTo NxtP32

        pStart = pRng.Start
        b = pText
        bMax = UBound(b) - 1

        ' Compute list prefix length for position correction
        Dim r32ListPrefixLen As Long
        r32ListPrefixLen = GetQListPrefixLen(para, pText)

        ' ======================================================
        '  NESTING-AWARE SCAN
        '
        '  nestDepth tracks how many quote levels deep we are.
        '  Even depths (0, 2, ...) expect the "outer" type.
        '  Odd  depths (1, 3, ...) expect the "inner" type.
        '
        '  Only FLAG wrong-type OPENING quotes.  Closing quotes
        '  just pop the stack. Straight-quote open/close is
        '  determined by preceding-character context (space/SOL
        '  = opening, letter/digit = closing).
        '
        '  Soft failure: if nestDepth underflows on close, reset
        '  to 0 and increment anomaly counter; if anomaly count
        '  is high, stop flagging for this document.
        ' ======================================================

        For i = 0 To bMax Step 2
            code = b(i) Or (CLng(b(i + 1)) * 256&)

            ' -- Apostrophe skip for any single-quote character --
            Dim isApostrophe As Boolean
            isApostrophe = False
            If code = QS Or code = QSC Or code = QSO Then
                isApostrophe = ByteIsApostropheExt(b, i, bMax)
            End If
            If isApostrophe Then GoTo NxtChar32

            ' -- Classify this character --
            Dim isOuterOpen As Boolean
            Dim isOuterClose As Boolean
            Dim isInnerOpen As Boolean
            Dim isInnerClose As Boolean
            Dim isStraightOuter As Boolean
            Dim isStraightInner As Boolean
            isOuterOpen = (code = outerOpen)
            isOuterClose = (code = outerClose)
            isInnerOpen = (code = innerOpen)
            isInnerClose = (code = innerClose)
            isStraightOuter = (code = outerStraight)
            isStraightInner = (code = innerStraight)

            ' Skip non-quote characters
            If Not (isOuterOpen Or isOuterClose Or isInnerOpen Or _
                    isInnerClose Or isStraightOuter Or isStraightInner) Then
                GoTo NxtChar32
            End If

            ' -- Determine if this is an opening or closing quote --
            Dim isOpening As Boolean
            Dim isClosing As Boolean
            Dim isOuterType As Boolean  ' True = outer-style char

            isOuterType = (isOuterOpen Or isOuterClose Or isStraightOuter)

            ' Smart quotes have directionality built in
            If isOuterOpen Or isInnerOpen Then
                isOpening = True
                isClosing = False
            ElseIf isOuterClose Or isInnerClose Then
                isOpening = False
                isClosing = True
            Else
                ' Straight quote: use preceding-character context
                ' Space, tab, newline, SOL, open-paren = opening
                ' Letter, digit, punctuation = closing
                Dim prevIsSpace As Boolean
                prevIsSpace = True  ' default: start-of-line = opening
                If i >= 2 Then
                    Dim prevCode As Long
                    prevCode = b(i - 2) Or (CLng(b(i - 1)) * 256&)
                    prevIsSpace = (prevCode = 32 Or prevCode = 9 Or _
                                   prevCode = 13 Or prevCode = 10 Or _
                                   prevCode = 160 Or prevCode = 40 Or _
                                   prevCode = 91 Or prevCode = 8212 Or _
                                   prevCode = 8211)
                End If
                If prevIsSpace Then
                    isOpening = True
                    isClosing = False
                Else
                    ' Check if the current nesting level was opened with same type
                    If nestDepth > 0 Then
                        Dim curLevelOuter As Boolean
                        If nestDepth <= 10 Then
                            curLevelOuter = levelIsOuter(nestDepth - 1)
                        Else
                            curLevelOuter = ((nestDepth Mod 2) = 0)
                        End If
                        If curLevelOuter = isOuterType Then
                            isOpening = False
                            isClosing = True
                        Else
                            isOpening = True
                            isClosing = False
                        End If
                    Else
                        ' At depth 0, non-space-preceded straight quote
                        ' is probably closing an untracked opening
                        isOpening = False
                        isClosing = True
                    End If
                End If
            End If

            ' -- Process opening quotes --
            If isOpening Then
                Dim expectOuter As Boolean
                expectOuter = ((nestDepth Mod 2) = 0)

                Dim wrongTypeOpen As Boolean
                wrongTypeOpen = (isOuterType <> expectOuter)

                ' Only flag if we have low anomaly count (reliable nesting state)
                If wrongTypeOpen And nestAnomalies < 5 Then
                    Dim issMsg As String
                    If expectOuter Then
                        If nestMode = "DOUBLE" Then
                            issMsg = "Outer quotation marks should use double quotation marks, not single."
                        Else
                            issMsg = "Outer quotation marks should use single quotation marks, not double."
                        End If
                    Else
                        If nestMode = "DOUBLE" Then
                            issMsg = "Inner quotation marks should use single quotation marks, not double."
                        Else
                            issMsg = "Inner quotation marks should use double quotation marks, not single."
                        End If
                    End If

                    pos = pStart + (i \ 2) - r32ListPrefixLen
                    Err.Clear
                    locRng.SetRange pos, pos + 1
                    If Err.Number <> 0 Then
                        locStr = "unknown location": Err.Clear
                    Else
                        Err.Clear
                        locStr = EngineGetLocationString(locRng, doc)
                        If Err.Number <> 0 Then
                            locStr = "unknown location": Err.Clear
                        End If
                    End If

                    issues.Add CreateIssueDict(RULE32, locStr, _
                        issMsg, "", pos, pos + 1, "warning")
                End If

                ' Push nesting level regardless (so its close is tracked)
                If nestDepth < 10 Then
                    levelIsOuter(nestDepth) = isOuterType
                End If
                nestDepth = nestDepth + 1
                GoTo NxtChar32
            End If

            ' -- Process closing quotes --
            If isClosing Then
                If nestDepth > 0 Then
                    nestDepth = nestDepth - 1
                Else
                    ' Underflow: malformed sequence.  Soft-reset.
                    nestAnomalies = nestAnomalies + 1
                    nestDepth = 0
                End If
                GoTo NxtChar32
            End If

NxtChar32:
        Next i

NxtP32:
    Next para
    On Error GoTo 0

    Set Check_SingleQuotesDefault = issues
End Function

' ================================================================
'  RULE 33 -- SMART QUOTE CONSISTENCY
'
'  Single pass over paragraphs: counts straight vs smart quotes
'  AND collects minority-style positions simultaneously.
'  Preference (smart or straight) is read from the engine toggle.
'  If both styles exist, flags the non-preferred style.
' ================================================================
Public Function Check_SmartQuoteConsistency( _
        doc As Document) As Collection

    Dim issues As New Collection
    Dim para As Paragraph
    Dim pRng As Range
    Dim pText As String
    Dim pStart As Long
    Dim b() As Byte
    Dim bMax As Long
    Dim i As Long, code As Long

    Dim prefStyle As String
    prefStyle = EngineGetSmartQuotePref()  ' "SMART" or "STRAIGHT"
    Dim preferSmart As Boolean
    preferSmart = (prefStyle <> "STRAIGHT")

    ' Counters
    Dim cStraight As Long, cSmart As Long

    ' Collect positions of the non-preferred style
    Dim fPos() As Long
    Dim fCnt As Long, fCap As Long
    fCap = 256
    ReDim fPos(0 To fCap - 1)

    ' -- Single pass: count + collect positions ------------------
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set pRng = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NxtP33

        If Not EngineIsInPageRange(pRng) Then GoTo NxtP33

        Err.Clear
        pText = pRng.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NxtP33
        If LenB(pText) = 0 Then GoTo NxtP33

        pStart = pRng.Start
        b = pText
        bMax = UBound(b) - 1

        ' Compute list prefix length for position correction
        Dim r33ListPrefixLen As Long
        r33ListPrefixLen = GetQListPrefixLen(para, pText)

        For i = 0 To bMax Step 2
            code = b(i) Or (CLng(b(i + 1)) * 256&)

            Select Case code
            Case QD
                cStraight = cStraight + 1
                If Not preferSmart Then GoTo NxtCode33
                ' Prefer smart -> collect straight positions
                If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1

            Case QDO, QDC
                cSmart = cSmart + 1
                If preferSmart Then GoTo NxtCode33
                ' Prefer straight -> collect smart positions
                If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1

            Case QS
                If Not ByteIsApostrophe(b, i, bMax) Then
                    cStraight = cStraight + 1
                    If preferSmart Then
                        If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                        fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1
                    End If
                End If

            Case QSO
                cSmart = cSmart + 1
                If Not preferSmart Then
                    If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                    fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1
                End If

            Case QSC
                If Not ByteIsApostrophe(b, i, bMax) Then
                    cSmart = cSmart + 1
                    If Not preferSmart Then
                        If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                        fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1
                    End If
                End If
            End Select
NxtCode33:
        Next i

NxtP33:
    Next para
    On Error GoTo 0

    ' -- No mix? Nothing to report ------------------------------
    If cStraight = 0 Or cSmart = 0 Then
        Set Check_SmartQuoteConsistency = issues
        Exit Function
    End If

    ' -- Summary finding ----------------------------------------
    Dim prefName As String, wrongName As String
    If preferSmart Then prefName = "smart": wrongName = "straight" _
    Else prefName = "straight": wrongName = "smart"

    issues.Add CreateIssueDict(RULE33, "Document", _
        "Quotation mark style is inconsistent. Found " & _
        cStraight & " straight and " & cSmart & _
        " smart quotation marks.", _
        "Use " & prefName & " quotation marks consistently " & _
        "throughout the document.", 0, 0, "warning")

    ' -- Flag each non-preferred quote ---------------------------
    Dim locRng As Range
    Set locRng = doc.Range(0, 1)

    Dim j As Long, pos As Long, locStr As String
    On Error Resume Next
    For j = 0 To fCnt - 1
        pos = fPos(j)
        Err.Clear
        locRng.SetRange pos, pos + 1
        If Err.Number <> 0 Then Err.Clear: GoTo SkipP33

        Err.Clear
        locStr = EngineGetLocationString(locRng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

        issues.Add CreateIssueDict(RULE33, locStr, _
            UCase(Left(wrongName, 1)) & Mid(wrongName, 2) & _
            " quotation mark found in document.", _
            "Replace with " & prefName & " quotation mark.", _
            pos, pos + 1, "warning")
SkipP33:
    Next j
    On Error GoTo 0

    Set Check_SmartQuoteConsistency = issues
End Function

' ================================================================
'  PRIVATE HELPERS
' ================================================================

' ------------------------------------------------------------
'  Apostrophe check on raw byte data (original strict version).
'  True when the character at byte offset bi is flanked by
'  letters on both sides (= mid-word = apostrophe, not quote).
'  Works directly on the byte array -- no Mid$/AscW overhead.
' ------------------------------------------------------------
Private Function ByteIsApostrophe(b() As Byte, _
        ByVal bi As Long, ByVal bMax As Long) As Boolean
    Dim pc As Long, nc As Long
    If bi < 2 Or bi + 3 > bMax Then Exit Function  ' False
    pc = b(bi - 2) Or (CLng(b(bi - 1)) * 256&)
    nc = b(bi + 2) Or (CLng(b(bi + 3)) * 256&)
    ByteIsApostrophe = IsLetterCode(pc) And IsLetterCode(nc)
End Function

' ------------------------------------------------------------
'  Extended apostrophe check: letter or digit flanked.
'  Treats letter+digit and digit+letter combos as apostrophes
'  too (e.g. 90's, '80s).  Used by the nesting scanner.
' ------------------------------------------------------------
Private Function ByteIsApostropheExt(b() As Byte, _
        ByVal bi As Long, ByVal bMax As Long) As Boolean
    Dim pc As Long, nc As Long
    If bi < 2 Or bi + 3 > bMax Then Exit Function  ' False
    pc = b(bi - 2) Or (CLng(b(bi - 1)) * 256&)
    nc = b(bi + 2) Or (CLng(b(bi + 3)) * 256&)
    ' Both sides must be letter or digit
    ByteIsApostropheExt = (IsLetterCode(pc) Or IsDigitCode(pc)) And _
                          (IsLetterCode(nc) Or IsDigitCode(nc))
End Function

' ------------------------------------------------------------
'  Letter test by code point (A-Z, a-z, extended Latin U+00C0
'  through U+02AF).  Covers accented characters common in UK
'  legal text (cafe, naive, resume, etc.).
' ------------------------------------------------------------
Private Function IsLetterCode(ByVal c As Long) As Boolean
    IsLetterCode = (c >= 65 And c <= 90) Or _
                   (c >= 97 And c <= 122) Or _
                   (c >= 192 And c <= 687)
End Function

' ------------------------------------------------------------
'  Digit test by code point (0-9).
' ------------------------------------------------------------
Private Function IsDigitCode(ByVal c As Long) As Boolean
    IsDigitCode = (c >= 48 And c <= 57)
End Function

' ------------------------------------------------------------
'  Style exclusion for Rule 32.  Paragraphs with "Block",
'  "Quote", or "Code" in their style name are skipped.
' ------------------------------------------------------------
Private Function IsExcludedStyle(ByVal sn As String) As Boolean
    If Len(sn) = 0 Then Exit Function  ' False
    Dim ls As String
    ls = LCase$(sn)
    IsExcludedStyle = (InStr(1, ls, "block", vbBinaryCompare) > 0) _
                   Or (InStr(1, ls, "quote", vbBinaryCompare) > 0) _
                   Or (InStr(1, ls, "code", vbBinaryCompare) > 0)
End Function

' ------------------------------------------------------------
'  Emit findings from a pre-collected position array (Rule 17).
'  Uses a single reusable Range to avoid per-finding allocation.
'  Checks page range per position (Rule 17 counts document-wide
'  but only flags within the configured page range).
' ------------------------------------------------------------
Private Sub EmitFromPositions(doc As Document, _
        issues As Collection, _
        positions() As Long, _
        cnt As Long, _
        ruleName As String, _
        issueText As String, _
        suggestion As String)

    If cnt = 0 Then Exit Sub

    Dim locRng As Range
    Set locRng = doc.Range(0, 1)

    Dim j As Long, pos As Long, locStr As String
    On Error Resume Next
    For j = 0 To cnt - 1
        pos = positions(j)
        Err.Clear
        locRng.SetRange pos, pos + 1
        If Err.Number <> 0 Then Err.Clear: GoTo SkipEmit

        If Not EngineIsInPageRange(locRng) Then GoTo SkipEmit

        Err.Clear
        locStr = EngineGetLocationString(locRng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

        issues.Add CreateIssueDict(ruleName, locStr, _
            issueText, suggestion, pos, pos + 1, "possible_error")
SkipEmit:
    Next j
    On Error GoTo 0
End Sub

' ------------------------------------------------------------
'  Create a dictionary-based finding (no class dependency).
' ------------------------------------------------------------
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

' ------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run( _
        "PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, _
        doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run( _
        "PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ------------------------------------------------------------
'  List prefix length for byte-array position correction.
'  para.Range.Text includes auto-generated list numbering
'  (e.g. "1." & vbTab) but para.Range.Start does NOT account
'  for it, so byte-array positions must subtract this offset.
' ------------------------------------------------------------
Private Function GetQListPrefixLen(para As Paragraph, ByVal paraText As String) As Long
    GetQListPrefixLen = 0
    On Error Resume Next
    Dim lStr As String
    lStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0
    If Len(lStr) = 0 Then Exit Function
    ' Verify the text actually starts with the list string
    If Len(paraText) > Len(lStr) Then
        If Left$(paraText, Len(lStr)) = lStr Then
            GetQListPrefixLen = Len(lStr)
            ' Account for tab separator after list number
            If Mid$(paraText, GetQListPrefixLen + 1, 1) = vbTab Then
                GetQListPrefixLen = GetQListPrefixLen + 1
            End If
        End If
    End If
End Function

' ------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetQuoteNesting
' ------------------------------------------------------------
Private Function EngineGetQuoteNesting() As String
    On Error Resume Next
    EngineGetQuoteNesting = Application.Run( _
        "PleadingsEngine.GetQuoteNesting")
    If Err.Number <> 0 Then
        Debug.Print "EngineGetQuoteNesting: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetQuoteNesting = "SINGLE"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetSmartQuotePref
' ------------------------------------------------------------
Private Function EngineGetSmartQuotePref() As String
    On Error Resume Next
    EngineGetSmartQuotePref = Application.Run( _
        "PleadingsEngine.GetSmartQuotePref")
    If Err.Number <> 0 Then
        Debug.Print "EngineGetSmartQuotePref: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetSmartQuotePref = "SMART"
        Err.Clear
    End If
    On Error GoTo 0
End Function
