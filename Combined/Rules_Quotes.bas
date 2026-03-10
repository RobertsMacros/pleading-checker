Attribute VB_Name = "Rules_Quotes"
' ============================================================
' Rules_Quotes.bas
' Quotation-mark rules for UK legal proofreading:
'   Rule 17: quotation mark consistency (straight vs curly)
'   Rule 32: single quotes as default outer marks
'   Rule 33: smart quote consistency (prefers curly)
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
Private Const QDO As Long = 8220   ' curly double open
Private Const QDC As Long = 8221   ' curly double close
Private Const QS  As Long = 39     ' straight single  '
Private Const QSO As Long = 8216   ' curly single open
Private Const QSC As Long = 8217   ' curly single close

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
    Dim cCD As Long   ' curly double
    Dim cSS As Long   ' straight single (excluding apostrophes)
    Dim cCS As Long   ' curly single   (excluding apostrophes)

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
            "Curly double quotation mark found; " & _
            "document predominantly uses straight", _
            "Change to straight double quotation mark (" & _
            Chr$(QD) & ")"

    ElseIf (Not dblStraight) And cSD > 0 Then
        EmitFromPositions doc, issues, pSD, cSD, RULE17, _
            "Straight double quotation mark found; " & _
            "document predominantly uses curly", _
            "Change to curly double quotation marks (" & _
            ChrW$(QDO) & ChrW$(QDC) & ")"
    End If

    ' -- Flag minority singles ----------------------------------
    If sglStraight And cCS > 0 Then
        EmitFromPositions doc, issues, pCS, cCS, RULE17, _
            "Curly single quotation mark found; " & _
            "document predominantly uses straight", _
            "Change to straight single quotation mark (" & _
            Chr$(QS) & ")"

    ElseIf (Not sglStraight) And cSS > 0 Then
        EmitFromPositions doc, issues, pSS, cSS, RULE17, _
            "Straight single quotation mark found; " & _
            "document predominantly uses curly", _
            "Change to curly single quotation marks (" & _
            ChrW$(QSO) & ChrW$(QSC) & ")"
    End If

    Set Check_QuotationMarkConsistency = issues
End Function

' ================================================================
'  RULE 32 -- SINGLE / DOUBLE QUOTES DEFAULT
'
'  Configurable: Single outer (UK convention) or Double outer (US).
'  One pass over paragraphs.  Per paragraph: one page-range check,
'  one style check, then a byte-array scan for the wrong quote type.
'  Uses a single reusable Range for location lookups.
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

    Dim issueMsg As String
    Dim suggMsg As String
    If nestMode = "DOUBLE" Then
        issueMsg = "Outer quotation marks should use double quotation marks."
        suggMsg = "Use double quotation marks instead of single quotation marks."
    Else
        issueMsg = "Outer quotation marks should use single quotation marks."
        suggMsg = "Use single quotation marks instead of double quotation marks."
    End If

    ' Reusable range -- created once, repositioned via SetRange
    Dim locRng As Range
    Set locRng = doc.Range(0, 1)

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

        ' Byte-array scan: flag the wrong quote type
        For i = 0 To bMax Step 2
            code = b(i) Or (CLng(b(i + 1)) * 256&)

            Dim isWrongQuote As Boolean
            isWrongQuote = False

            If nestMode = "DOUBLE" Then
                ' Flag single quotes (but skip apostrophes)
                If code = QS Or code = QSO Or code = QSC Then
                    If Not ByteIsApostrophe(b, i, bMax) Then
                        isWrongQuote = True
                    End If
                End If
            Else
                ' Flag double quotes
                If code = QD Or code = QDO Or code = QDC Then
                    isWrongQuote = True
                End If
            End If

            If isWrongQuote Then
                pos = pStart + (i \ 2)
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
                    issueMsg, suggMsg, pos, pos + 1, "warning")
            End If
        Next i

NxtP32:
    Next para
    On Error GoTo 0

    Set Check_SingleQuotesDefault = issues
End Function

' ================================================================
'  RULE 33 -- SMART QUOTE CONSISTENCY
'
'  Single pass over paragraphs: counts straight vs curly quotes
'  AND collects minority-style positions simultaneously.
'  Preference (curly or straight) is read from the engine toggle.
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
    prefStyle = EngineGetSmartQuotePref()  ' "CURLY" or "STRAIGHT"
    Dim preferCurly As Boolean
    preferCurly = (prefStyle <> "STRAIGHT")

    ' Counters
    Dim cStraight As Long, cCurly As Long

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

        For i = 0 To bMax Step 2
            code = b(i) Or (CLng(b(i + 1)) * 256&)

            Select Case code
            Case QD
                cStraight = cStraight + 1
                If Not preferCurly Then GoTo NxtCode33
                ' Prefer curly -> collect straight positions
                If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                fPos(fCnt) = pStart + (i \ 2): fCnt = fCnt + 1

            Case QDO, QDC
                cCurly = cCurly + 1
                If preferCurly Then GoTo NxtCode33
                ' Prefer straight -> collect curly positions
                If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                fPos(fCnt) = pStart + (i \ 2): fCnt = fCnt + 1

            Case QS
                If Not ByteIsApostrophe(b, i, bMax) Then
                    cStraight = cStraight + 1
                    If preferCurly Then
                        If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                        fPos(fCnt) = pStart + (i \ 2): fCnt = fCnt + 1
                    End If
                End If

            Case QSO
                cCurly = cCurly + 1
                If Not preferCurly Then
                    If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                    fPos(fCnt) = pStart + (i \ 2): fCnt = fCnt + 1
                End If

            Case QSC
                If Not ByteIsApostrophe(b, i, bMax) Then
                    cCurly = cCurly + 1
                    If Not preferCurly Then
                        If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                        fPos(fCnt) = pStart + (i \ 2): fCnt = fCnt + 1
                    End If
                End If
            End Select
NxtCode33:
        Next i

NxtP33:
    Next para
    On Error GoTo 0

    ' -- No mix? Nothing to report ------------------------------
    If cStraight = 0 Or cCurly = 0 Then
        Set Check_SmartQuoteConsistency = issues
        Exit Function
    End If

    ' -- Summary finding ----------------------------------------
    Dim prefName As String, wrongName As String
    If preferCurly Then prefName = "curly": wrongName = "straight" _
    Else prefName = "straight": wrongName = "curly"

    issues.Add CreateIssueDict(RULE33, "Document", _
        "Quotation mark style is inconsistent. Found " & _
        cStraight & " straight and " & cCurly & _
        " curly quotation marks.", _
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
'  Apostrophe check on raw byte data.
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
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetQuoteNesting
' ------------------------------------------------------------
Private Function EngineGetQuoteNesting() As String
    On Error Resume Next
    EngineGetQuoteNesting = Application.Run( _
        "PleadingsEngine.GetQuoteNesting")
    If Err.Number <> 0 Then
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
        EngineGetSmartQuotePref = "CURLY"
        Err.Clear
    End If
    On Error GoTo 0
End Function
