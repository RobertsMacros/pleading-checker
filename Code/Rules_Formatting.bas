Attribute VB_Name = "Rules_Formatting"
' ============================================================
' Rules_Formatting.bas
' Combined module for formatting-related rules:
'   - Rule06: Paragraph break consistency (headings)
'   - Rule11: Font consistency (headings, body, footnotes)
'
' IsBlockQuotePara is a public helper used by other modules.
' It requires STRONG indicators beyond mere indentation:
'   - Quote-related style name (definitive)
'   - Indentation + quotation-mark wrapping
'   - Indentation + entirely italic text
' Indentation + smaller font alone is NOT sufficient.
' ============================================================
Option Explicit

Private Const RULE_NAME_PARAGRAPH_BREAK As String = "paragraph_break_consistency"
Private Const RULE_NAME_FONT As String = "font_consistency"

' ============================================================
'  RULE 06 HELPERS
' ============================================================

' -- Classify spacing pattern after a heading ----------------
' Returns: "no_spacing", "spacing_Npt", or "manual_double_break"
Private Function ClassifyAfterSpacing(para As Paragraph, doc As Document, paraIdx As Long) As String
    Dim spAfter As Single
    spAfter = para.Format.SpaceAfter

    ' Check if the next paragraph is empty (manual double break)
    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    If paraIdx < totalParas Then
        Dim nextPara As Paragraph
        Set nextPara = doc.Paragraphs(paraIdx + 1)
        Dim nextText As String
        nextText = nextPara.Range.Text
        ' An empty paragraph contains only vbCr
        If nextText = vbCr Then
            ClassifyAfterSpacing = "manual_double_break"
            Exit Function
        End If
    End If

    If spAfter = 0 Then
        ClassifyAfterSpacing = "no_spacing"
    Else
        ClassifyAfterSpacing = "spacing_" & CStr(CLng(spAfter)) & "pt"
    End If
End Function

' -- Classify SpaceBefore pattern ----------------------------
Private Function ClassifyBeforeSpacing(para As Paragraph) As String
    Dim spBefore As Single
    spBefore = para.Format.SpaceBefore
    If spBefore = 0 Then
        ClassifyBeforeSpacing = "before_0pt"
    Else
        ClassifyBeforeSpacing = "before_" & CStr(CLng(spBefore)) & "pt"
    End If
End Function

' ============================================================
'  RULE 11 HELPERS
' ============================================================

' -- Helper: build a font profile key ------------------------
Private Function FontKey(ByVal fontName As String, ByVal fontSize As Single) As String
    FontKey = fontName & "|" & CStr(fontSize)
End Function

' -- Helper: find dominant key in a dictionary of counts -----
Private Function GetDominant(counts As Object) As String
    Dim k As Variant
    Dim maxCnt As Long
    Dim domKey As String
    maxCnt = 0
    domKey = ""
    For Each k In counts.keys
        If counts(k) > maxCnt Then
            maxCnt = counts(k)
            domKey = CStr(k)
        End If
    Next k
    GetDominant = domKey
End Function

' -- Helper: parse font key back to readable description -----
' ------------------------------------------------------------
'  PUBLIC: Detect block quote / indented extract paragraphs.
'
'  STRICT RULE: Indentation alone is NEVER enough.
'  Smaller font + indentation alone is NEVER enough.
'  A block quote must have at least one of:
'    1. A block-quote style (name contains "quote"/"block"/"extract")
'    2. Enclosing quotation marks AND indentation
'    3. Entirely italic text AND indentation
'  Lists, numbered paragraphs, and bullet items are explicitly excluded.
' ------------------------------------------------------------
Public Function IsBlockQuotePara(para As Paragraph) As Boolean
    IsBlockQuotePara = False
    On Error Resume Next

    ' ==========================================================
    '  CHECK 0: Exclude list paragraphs (numbered, bulleted, etc.)
    '  Lists must NEVER be treated as block quotes.
    ' ==========================================================
    Dim listLvl As Long
    listLvl = 0
    listLvl = para.Range.ListFormat.ListLevelNumber
    If Err.Number <> 0 Then listLvl = 0: Err.Clear
    ' ListLevelNumber > 0 means this paragraph is in a list
    If listLvl > 0 Then
        On Error GoTo 0
        Exit Function
    End If

    ' Also check for list-like text patterns (manual numbering)
    Dim pTextRaw As String
    pTextRaw = ""
    pTextRaw = para.Range.Text
    If Err.Number <> 0 Then pTextRaw = "": Err.Clear
    On Error GoTo 0
    Dim pTextTrimmed As String
    pTextTrimmed = Replace(Replace(Replace(pTextRaw, vbCr, ""), vbTab, ""), ChrW(160), " ")
    pTextTrimmed = Trim$(pTextTrimmed)

    ' Check for bullet-like or number-list-like starts
    If Len(pTextTrimmed) > 1 Then
        Dim firstTwo As String
        firstTwo = Left$(pTextTrimmed, 2)
        ' Bullet characters: bullet, en-dash, em-dash, hyphen
        If Left$(pTextTrimmed, 1) = ChrW(8226) Or _
           Left$(pTextTrimmed, 1) = ChrW(8211) & " " Or _
           firstTwo = "- " Or firstTwo = "* " Then
            On Error GoTo 0
            Exit Function
        End If
        ' Numbered list pattern: "(a)", "(i)", "(1)", "1.", "a.", "i."
        If pTextTrimmed Like "(#)*" Or pTextTrimmed Like "(##)*" Or _
           pTextTrimmed Like "([a-z])*" Or pTextTrimmed Like "([ivx])*" Or _
           pTextTrimmed Like "#.*" Or pTextTrimmed Like "##.*" Or _
           pTextTrimmed Like "[a-z].*" Then
            On Error GoTo 0
            Exit Function
        End If
    End If

    ' Also check ListFormat.ListString for auto-numbered lists
    Dim listStr As String
    listStr = ""
    On Error Resume Next
    listStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then listStr = "": Err.Clear
    If Len(listStr) > 0 Then
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  CHECK 1: Style name for quote/block/extract keywords
    '  (Definitive indicator - no other checks needed)
    ' ==========================================================
    Dim sn As String
    sn = LCase(para.Style.NameLocal)
    If Err.Number <> 0 Then sn = "": Err.Clear
    If InStr(sn, "quote") > 0 Or InStr(sn, "block") > 0 Or _
       InStr(sn, "extract") > 0 Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  INDENTATION CHECK
    '  All remaining indicators require indentation.
    ' ==========================================================
    Dim leftInd As Single
    leftInd = para.Format.LeftIndent
    If Err.Number <> 0 Then leftInd = 0: Err.Clear
    On Error GoTo 0

    ' No indentation = not a block quote (style check already done above)
    If leftInd <= 18 Then
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  CHECK 2: Indentation + quotation marks wrapping
    '  Starts or ends with a quotation mark character.
    ' ==========================================================
    Dim startsWithQuote As Boolean
    Dim endsWithQuote As Boolean
    startsWithQuote = False
    endsWithQuote = False
    If Len(pTextTrimmed) > 1 Then
        Dim fcChar As String
        Dim lcChar As String
        fcChar = Left$(pTextTrimmed, 1)
        lcChar = Right$(pTextTrimmed, 1)
        startsWithQuote = (fcChar = Chr(34) Or fcChar = ChrW(8220) Or fcChar = ChrW(8216))
        endsWithQuote = (lcChar = Chr(34) Or lcChar = ChrW(8221) Or lcChar = ChrW(8217))
    End If

    ' Block quote if indented AND wrapped in quotation marks
    If startsWithQuote Or endsWithQuote Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  CHECK 3: Indentation + entirely italic
    '  wdTrue (-1) means ALL text in the range is italic.
    ' ==========================================================
    Dim italVal As Long
    On Error Resume Next
    italVal = para.Range.Font.Italic
    If Err.Number <> 0 Then italVal = 0: Err.Clear
    If italVal = -1 Then  ' wdTrue = -1 means ALL italic
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  DEFAULT: Indented but no strong indicator = NOT a block quote.
    '  Smaller font + indentation alone is deliberately NOT enough.
    '  This prevents indented lists, definitions, and body text
    '  from being misclassified.
    ' ==========================================================

    On Error GoTo 0
End Function

Private Function FontDescription(ByVal fKey As String) As String
    Dim parts() As String
    parts = Split(fKey, "|")
    If UBound(parts) >= 1 Then
        FontDescription = parts(0) & " " & parts(1) & "pt"
    Else
        FontDescription = fKey
    End If
End Function

' ============================================================
'  RULE 06: PARAGRAPH BREAK CONSISTENCY
' ============================================================
Public Function Check_ParagraphBreakConsistency(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim lvl As Long
    Dim info() As Variant

    On Error Resume Next

    ' -- Dictionaries keyed by outline level -----------------
    ' afterPatterns:  level -> Dictionary(pattern -> count)
    ' beforePatterns: level -> Dictionary(pattern -> count)
    ' headingInfos:   level -> Collection of Array(paraIdx, afterPattern, beforePattern, rangeStart, rangeEnd, text)
    Dim afterPatterns As Object
    Set afterPatterns = CreateObject("Scripting.Dictionary")
    Dim beforePatterns As Object
    Set beforePatterns = CreateObject("Scripting.Dictionary")
    Dim headingInfos As Object
    Set headingInfos = CreateObject("Scripting.Dictionary")

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1
        If paraIdx Mod 200 = 0 Then DoEvents

        lvl = para.OutlineLevel
        If lvl < wdOutlineLevel1 Or lvl > wdOutlineLevel9 Then GoTo NextPara

        ' Page range filter
        If Not EngineIsInPageRange(para.Range) Then GoTo NextPara

        ' Classify after-spacing
        Dim aftPat As String
        aftPat = ClassifyAfterSpacing(para, doc, paraIdx)

        ' Classify before-spacing
        Dim befPat As String
        befPat = ClassifyBeforeSpacing(para)

        ' -- Track after-spacing counts ---------------------
        If Not afterPatterns.Exists(lvl) Then
            afterPatterns.Add lvl, CreateObject("Scripting.Dictionary")
        End If
        Dim aftDict As Object
        Set aftDict = afterPatterns(lvl)
        If aftDict.Exists(aftPat) Then
            aftDict(aftPat) = aftDict(aftPat) + 1
        Else
            aftDict.Add aftPat, 1
        End If

        ' -- Track before-spacing counts --------------------
        If Not beforePatterns.Exists(lvl) Then
            beforePatterns.Add lvl, CreateObject("Scripting.Dictionary")
        End If
        Dim befDict As Object
        Set befDict = beforePatterns(lvl)
        If befDict.Exists(befPat) Then
            befDict(befPat) = befDict(befPat) + 1
        Else
            befDict.Add befPat, 1
        End If

        ' -- Store heading info -----------------------------
        If Not headingInfos.Exists(lvl) Then
            headingInfos.Add lvl, New Collection
        End If
        ReDim info(0 To 5)
        info(0) = paraIdx
        info(1) = aftPat
        info(2) = befPat
        info(3) = para.Range.Start
        info(4) = para.Range.End
        info(5) = Trim$(Replace(para.Range.Text, vbCr, ""))
        headingInfos(lvl).Add info
NextPara:
    Next para

    ' -- Determine dominant patterns and flag deviations -----
    Dim lvlKey As Variant
    For Each lvlKey In headingInfos.keys
        Dim hdgs As Collection
        Set hdgs = headingInfos(lvlKey)
        If hdgs.Count <= 1 Then GoTo NextLevel

        ' Find dominant after-pattern
        Dim domAfter As String
        domAfter = ""
        Dim maxCnt As Long
        maxCnt = 0
        If afterPatterns.Exists(lvlKey) Then
            Set aftDict = afterPatterns(lvlKey)
            Dim pk As Variant
            For Each pk In aftDict.keys
                If aftDict(pk) > maxCnt Then
                    maxCnt = aftDict(pk)
                    domAfter = CStr(pk)
                End If
            Next pk
        End If

        ' Find dominant before-pattern
        Dim domBefore As String
        domBefore = ""
        maxCnt = 0
        If beforePatterns.Exists(lvlKey) Then
            Set befDict = beforePatterns(lvlKey)
            For Each pk In befDict.keys
                If befDict(pk) > maxCnt Then
                    maxCnt = befDict(pk)
                    domBefore = CStr(pk)
                End If
            Next pk
        End If

        ' Flag outliers
        Dim h As Long
        For h = 1 To hdgs.Count
            Dim hInfo As Variant
            hInfo = hdgs(h)

            Dim hAft As String
            hAft = CStr(hInfo(1))
            Dim hBef As String
            hBef = CStr(hInfo(2))
            Dim hText As String
            hText = CStr(hInfo(5))

            ' Check after-spacing deviation
            If hAft <> domAfter And Len(domAfter) > 0 Then
                Dim findingA As Object
                Dim rngA As Range
                Set rngA = doc.Range(CLng(hInfo(3)), CLng(hInfo(4)))
                Dim locA As String
                locA = EngineGetLocationString(rngA, doc)

                Set findingA = CreateIssueDict(RULE_NAME_PARAGRAPH_BREAK, locA, "After-heading spacing inconsistency at '" & hText & "': uses " & hAft & " but dominant pattern for level " & CLng(lvlKey) & " headings is " & domAfter, "Change spacing after this heading to match: " & domAfter, CLng(hInfo(3)), CLng(hInfo(4)), "possible_error")
                issues.Add findingA
            End If

            ' Check before-spacing deviation
            If hBef <> domBefore And Len(domBefore) > 0 Then
                Dim findingB As Object
                Dim rngB As Range
                Set rngB = doc.Range(CLng(hInfo(3)), CLng(hInfo(4)))
                Dim locB As String
                locB = EngineGetLocationString(rngB, doc)

                Set findingB = CreateIssueDict(RULE_NAME_PARAGRAPH_BREAK, locB, "Before-heading spacing inconsistency at '" & hText & "': uses " & hBef & " but dominant pattern for level " & CLng(lvlKey) & " headings is " & domBefore, "Change spacing before this heading to match: " & domBefore, CLng(hInfo(3)), CLng(hInfo(4)), "possible_error")
                issues.Add findingB
            End If
        Next h
NextLevel:
    Next lvlKey

    If Err.Number <> 0 Then
        Debug.Print "Check_ParagraphBreakConsistency: exiting with Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Set Check_ParagraphBreakConsistency = issues
End Function

' ============================================================
'  RULE 11: FONT CONSISTENCY
'  Type-based approach: classify paragraphs into heading,
'  body, block-quote; compute dominant font per type;
'  flag outliers within each type.
'
'  Block-quote classification in font consistency uses the
'  same strict criteria as IsBlockQuotePara: indentation
'  plus italic, or indentation plus quotation wrapping.
'  Indentation + smaller font alone does NOT classify as
'  block quote here either.
' ============================================================
Public Function Check_FontConsistency(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next
    PerfTimerStart "font_consist:scan"

    ' ==========================================================
    '  SINGLE MERGED PASS: Classify paragraphs and collect font
    '  tallies in one scan.
    ' ==========================================================
    Dim bodyIndents As Object   ' LeftIndent (rounded) -> count
    Set bodyIndents = CreateObject("Scripting.Dictionary")
    Dim bodySizes As Object     ' FontSize (rounded) -> count
    Set bodySizes = CreateObject("Scripting.Dictionary")

    Dim headingFonts As Object
    Set headingFonts = CreateObject("Scripting.Dictionary")
    Dim bodyFonts As Object
    Set bodyFonts = CreateObject("Scripting.Dictionary")
    Dim bqFonts As Object
    Set bqFonts = CreateObject("Scripting.Dictionary")
    Dim footnoteFonts As Object
    Set footnoteFonts = CreateObject("Scripting.Dictionary")

    ' Cache paragraph metadata in arrays to avoid re-scanning
    Dim paraCap As Long
    paraCap = 512
    Dim pLevels() As Long       ' outline level
    Dim pIndents() As Single    ' left indent
    Dim pFontNames() As String  ' font name
    Dim pFontSizes() As Single  ' font size
    Dim pStarts() As Long       ' range start
    Dim pEnds() As Long         ' range end
    Dim pTypes() As String      ' "heading"/"body"/"block_quote"/""
    Dim pInRange() As Boolean   ' in page range
    ReDim pLevels(0 To paraCap - 1)
    ReDim pIndents(0 To paraCap - 1)
    ReDim pFontNames(0 To paraCap - 1)
    ReDim pFontSizes(0 To paraCap - 1)
    ReDim pStarts(0 To paraCap - 1)
    ReDim pEnds(0 To paraCap - 1)
    ReDim pTypes(0 To paraCap - 1)
    ReDim pInRange(0 To paraCap - 1)

    Dim para As Paragraph
    Dim paraIdx As Long
    Dim fk As String

    ' -- Single scan: collect all paragraph metadata --
    paraIdx = 0
    For Each para In doc.Paragraphs
        If paraIdx Mod 200 = 0 And paraIdx > 0 Then DoEvents
        ' Grow arrays if needed
        If paraIdx >= paraCap Then
            paraCap = paraCap * 2
            ReDim Preserve pLevels(0 To paraCap - 1)
            ReDim Preserve pIndents(0 To paraCap - 1)
            ReDim Preserve pFontNames(0 To paraCap - 1)
            ReDim Preserve pFontSizes(0 To paraCap - 1)
            ReDim Preserve pStarts(0 To paraCap - 1)
            ReDim Preserve pEnds(0 To paraCap - 1)
            ReDim Preserve pTypes(0 To paraCap - 1)
            ReDim Preserve pInRange(0 To paraCap - 1)
        End If

        pTypes(paraIdx) = ""
        pInRange(paraIdx) = EngineIsInPageRange(para.Range)
        If Not pInRange(paraIdx) Then
            paraIdx = paraIdx + 1
            GoTo NextScanPara
        End If

        Dim lvl As Long
        lvl = para.OutlineLevel
        If Err.Number <> 0 Then lvl = wdOutlineLevelBodyText: Err.Clear
        pLevels(paraIdx) = lvl

        Dim curInd As Single
        curInd = para.Format.LeftIndent
        If Err.Number <> 0 Then curInd = 0: Err.Clear
        pIndents(paraIdx) = curInd

        pStarts(paraIdx) = para.Range.Start
        pEnds(paraIdx) = para.Range.End

        ' Font info (read once, cache for reuse)
        Dim curFontName As String
        Dim curFontSize As Single
        curFontName = para.Range.Font.Name
        If Err.Number <> 0 Then curFontName = "": Err.Clear
        curFontSize = para.Range.Font.Size
        If Err.Number <> 0 Then curFontSize = 0: Err.Clear
        If curFontSize > 1000 Then curFontSize = 0
        pFontNames(paraIdx) = curFontName
        pFontSizes(paraIdx) = curFontSize

        ' Tally body-text indent and size (for dominant calculation)
        If lvl = wdOutlineLevelBodyText And curFontSize > 0 Then
            Dim indKey As String
            indKey = CStr(CLng(curInd))
            If bodyIndents.Exists(indKey) Then
                bodyIndents(indKey) = bodyIndents(indKey) + 1
            Else
                bodyIndents.Add indKey, 1
            End If
            Dim szKey As String
            szKey = CStr(CLng(curFontSize * 10))
            If bodySizes.Exists(szKey) Then
                bodySizes(szKey) = bodySizes(szKey) + 1
            Else
                bodySizes.Add szKey, 1
            End If
        End If

        paraIdx = paraIdx + 1
NextScanPara:
    Next para
    Dim totalParas As Long
    totalParas = paraIdx

    ' Determine dominant body indent and font size
    Dim domBodyIndent As Single
    Dim domBodySizeTenths As Long
    Dim domBodySize As Single
    Dim tmpDomKey As String
    tmpDomKey = GetDominant(bodyIndents)
    If Len(tmpDomKey) > 0 Then domBodyIndent = CSng(tmpDomKey) Else domBodyIndent = 0
    tmpDomKey = GetDominant(bodySizes)
    If Len(tmpDomKey) > 0 Then domBodySizeTenths = CLng(tmpDomKey) Else domBodySizeTenths = 0
    domBodySize = CSng(domBodySizeTenths) / 10#

    PerfTimerEnd "font_consist:scan"
    PerfTimerStart "font_consist:classify"
    ' -- Classify paragraphs and tally fonts --
    ' Block-quote classification uses STRICT criteria:
    '   - Indentation + full italic (font.Italic = wdTrue)
    '   - Indentation + quote wrapping (first/last char is quote mark)
    '   - Style name with "quote"/"block"/"extract"
    ' Indentation + smaller font alone = classified as body, NOT block_quote.
    Dim pi As Long
    For pi = 0 To totalParas - 1
        If Not pInRange(pi) Then GoTo NextClassify

        Dim paraType As String
        paraType = ""
        Dim isHeading As Boolean
        isHeading = (pLevels(pi) >= wdOutlineLevel1 And pLevels(pi) <= wdOutlineLevel9)

        If isHeading Then
            paraType = "heading"
        ElseIf pLevels(pi) = wdOutlineLevelBodyText Then
            ' Lightweight block-quote classification.
            ' Only create a Range object when indentation is high enough
            ' to potentially qualify as a block quote (saves COM calls).
            Dim isBQ As Boolean
            isBQ = False

            If pIndents(pi) > domBodyIndent + 18 Then
                ' Indented enough — need Range for italic/quote checks
                Dim bqRng As Range
                Set bqRng = doc.Range(pStarts(pi), pEnds(pi))
                If Err.Number = 0 Then
                    ' Check style name
                    Dim paraStyleName As String
                    paraStyleName = ""
                    paraStyleName = bqRng.ParagraphStyle
                    If Err.Number <> 0 Then paraStyleName = "": Err.Clear
                    Dim lsn As String
                    lsn = LCase$(paraStyleName)
                    If InStr(lsn, "quote") > 0 Or InStr(lsn, "block") > 0 Or _
                       InStr(lsn, "extract") > 0 Then
                        isBQ = True
                    End If

                    ' Check indentation + italic
                    If Not isBQ Then
                        Dim italCheck As Long
                        italCheck = bqRng.Font.Italic
                        If Err.Number <> 0 Then italCheck = 0: Err.Clear
                        If italCheck = -1 Then isBQ = True
                    End If

                    ' Check indentation + quotation wrapping
                    If Not isBQ Then
                        Dim bqText As String
                        bqText = ""
                        bqText = bqRng.Text
                        If Err.Number <> 0 Then bqText = "": Err.Clear
                        bqText = Trim$(Replace(Replace(bqText, vbCr, ""), vbTab, ""))
                        If Len(bqText) > 1 Then
                            Dim bqFirst As String
                            Dim bqLast As String
                            bqFirst = Left$(bqText, 1)
                            bqLast = Right$(bqText, 1)
                            If bqFirst = Chr(34) Or bqFirst = ChrW(8220) Or bqFirst = ChrW(8216) Or _
                               bqLast = Chr(34) Or bqLast = ChrW(8221) Or bqLast = ChrW(8217) Then
                                isBQ = True
                            End If
                        End If
                    End If
                Else
                    Err.Clear
                End If
            End If

            If isBQ Then
                paraType = "block_quote"
            Else
                paraType = "body"
            End If
        End If

        pTypes(pi) = paraType

        ' Tally font for this type
        If Len(pFontNames(pi)) > 0 And pFontSizes(pi) > 0 Then
            fk = FontKey(pFontNames(pi), pFontSizes(pi))
            Select Case paraType
                Case "heading"
                    If headingFonts.Exists(fk) Then
                        headingFonts(fk) = headingFonts(fk) + 1
                    Else
                        headingFonts.Add fk, 1
                    End If
                Case "body"
                    If bodyFonts.Exists(fk) Then
                        bodyFonts(fk) = bodyFonts(fk) + 1
                    Else
                        bodyFonts.Add fk, 1
                    End If
                Case "block_quote"
                    If bqFonts.Exists(fk) Then
                        bqFonts(fk) = bqFonts(fk) + 1
                    Else
                        bqFonts.Add fk, 1
                    End If
            End Select
        End If
NextClassify:
    Next pi

    ' -- Footnotes ------------------------------------------
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        If Not EngineIsInPageRange(fn.Range) Then GoTo NextFootnote

        Dim fnFontName As String
        Dim fnFontSize As Single
        fnFontName = fn.Range.Font.Name
        fnFontSize = fn.Range.Font.Size

        If Len(fnFontName) > 0 And fnFontSize > 0 And fnFontSize < 1000 Then
            fk = FontKey(fnFontName, fnFontSize)
            If footnoteFonts.Exists(fk) Then
                footnoteFonts(fk) = footnoteFonts(fk) + 1
            Else
                footnoteFonts.Add fk, 1
            End If
        End If
NextFootnote:
    Next fn

    PerfTimerEnd "font_consist:classify"
    PerfTimerStart "font_consist:flag"
    ' ==========================================================
    '  PASS 2: Determine dominant fonts per type
    ' ==========================================================
    Dim domHeading As String
    Dim domBody As String
    Dim domBQ As String
    Dim domFootnote As String

    domHeading = GetDominant(headingFonts)
    domBody = GetDominant(bodyFonts)
    domBQ = GetDominant(bqFonts)
    domFootnote = GetDominant(footnoteFonts)

    ' Only check block_quote type if there are at least 2 paragraphs
    ' (too small a sample otherwise)
    Dim bqTotalCount As Long
    bqTotalCount = 0
    Dim bqK As Variant
    For Each bqK In bqFonts.keys
        bqTotalCount = bqTotalCount + bqFonts(bqK)
    Next bqK
    If bqTotalCount < 2 Then domBQ = ""

    ' ==========================================================
    '  PASS 3: Flag deviations using cached data.
    ' ==========================================================
    Dim paraFontName As String
    Dim paraFontSize As Single

    For pi = 0 To totalParas - 1
        If Not pInRange(pi) Then GoTo NextParaFont2
        If Len(pTypes(pi)) = 0 Then GoTo NextParaFont2

        Dim expectedFont As String
        Dim context As String
        expectedFont = ""
        context = ""

        Select Case pTypes(pi)
            Case "heading"
                If Len(domHeading) > 0 Then
                    expectedFont = domHeading
                    context = "heading"
                End If
            Case "body"
                If Len(domBody) > 0 Then
                    expectedFont = domBody
                    context = "body"
                End If
            Case "block_quote"
                If Len(domBQ) > 0 Then
                    expectedFont = domBQ
                    context = "block quote"
                End If
        End Select

        If Len(expectedFont) = 0 Then GoTo NextParaFont2

        ' -- Check at paragraph level using cached data ----
        paraFontName = pFontNames(pi)
        paraFontSize = pFontSizes(pi)

        If Len(paraFontName) > 0 And paraFontSize > 0 Then
            fk = FontKey(paraFontName, paraFontSize)
            If fk <> expectedFont Then
                Dim findingPara As Object
                Dim locP As String
                Dim paraRng As Range
                Set paraRng = doc.Range(pStarts(pi), pEnds(pi))
                locP = EngineGetLocationString(paraRng, doc)

                Dim cleanParaText As String
                cleanParaText = Trim$(Replace(Left$(paraRng.Text, 60), vbCr, ""))

                Set findingPara = CreateIssueDict(RULE_NAME_FONT, locP, _
                    "Font inconsistency in " & context & ": '" & cleanParaText & _
                    "...' uses " & FontDescription(fk) & " but dominant " & _
                    context & " font is " & FontDescription(expectedFont), _
                    "Change to " & FontDescription(expectedFont), _
                    pStarts(pi), pEnds(pi), "error")
                issues.Add findingPara
                GoTo NextParaFont2
            End If
        End If

        ' -- Run-level check only for mixed-font paragraphs --
        ' (Font info was 0/empty = mixed formatting detected in scan)
        If Len(paraFontName) = 0 Or paraFontSize <= 0 Then
            If pEnds(pi) - pStarts(pi) > 1 Then
                Dim runRange As Range
                Dim runText As String
                Dim isField As Boolean

                Set runRange = doc.Range(pStarts(pi), pEnds(pi))
                runRange.Collapse wdCollapseStart

                On Error Resume Next
                Do While runRange.Start < pEnds(pi)
                    runRange.MoveEnd wdCharacterFormatting, 1
                    If runRange.Start >= pEnds(pi) Then Exit Do

                    Err.Clear
                    runText = runRange.Text
                    If Err.Number <> 0 Then Err.Clear: GoTo AdvanceFontRun

                    If Len(Trim$(Replace(Replace(runText, vbCr, ""), vbLf, ""))) = 0 Then
                        GoTo AdvanceFontRun
                    End If

                    isField = False
                    If runRange.Fields.Count > 0 Then isField = True
                    If Err.Number <> 0 Then Err.Clear: isField = False

                    If Not isField Then
                        fk = FontKey(runRange.Font.Name, runRange.Font.Size)
                        If fk <> expectedFont And Len(runRange.Font.Name) > 0 And _
                           runRange.Font.Size > 0 And runRange.Font.Size < 1000 Then
                            Dim findingRun As Object
                            Dim locR As String
                            Dim cleanRunText As String
                            locR = EngineGetLocationString(runRange, doc)
                            cleanRunText = Trim$(Replace(Left$(runText, 40), vbCr, ""))

                            Set findingRun = CreateIssueDict(RULE_NAME_FONT, locR, _
                                "Mid-paragraph font change in " & context & ": '" & cleanRunText & _
                                "' uses " & FontDescription(fk) & " instead of " & FontDescription(expectedFont), _
                                "Change to " & FontDescription(expectedFont), _
                                runRange.Start, runRange.End, "error")
                            issues.Add findingRun
                            On Error GoTo 0
                            GoTo NextParaFont2
                        End If
                    End If

AdvanceFontRun:
                    runRange.Collapse wdCollapseEnd
                Loop
                On Error GoTo 0
            End If
        End If

NextParaFont2:
    Next pi

    ' ==========================================================
    '  PASS 4: Check footnote font deviations
    ' ==========================================================
    If Len(domFootnote) > 0 Then
        For Each fn In doc.Footnotes
            If Not EngineIsInPageRange(fn.Range) Then GoTo NextFN2

            fnFontName = fn.Range.Font.Name
            fnFontSize = fn.Range.Font.Size

            If Len(fnFontName) > 0 And fnFontSize > 0 And fnFontSize < 1000 Then
                fk = FontKey(fnFontName, fnFontSize)
                If fk <> domFootnote Then
                    Dim findingFN As Object
                    Dim locFN As String
                    locFN = EngineGetLocationString(fn.Range, doc)

                    Dim cleanFNText As String
                    cleanFNText = Trim$(Replace(Left$(fn.Range.Text, 50), vbCr, ""))

                    Set findingFN = CreateIssueDict(RULE_NAME_FONT, locFN, _
                        "Footnote font inconsistency: '" & cleanFNText & _
                        "...' uses " & FontDescription(fk) & " but dominant " & _
                        "footnote font is " & FontDescription(domFootnote), _
                        "Change to " & FontDescription(domFootnote), _
                        fn.Range.Start, fn.Range.End, "error")
                    issues.Add findingFN
                End If
            End If
NextFN2:
        Next fn
    End If

    PerfTimerEnd "font_consist:flag"

    If Err.Number <> 0 Then
        Debug.Print "Check_FontConsistency: exiting with Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Set Check_FontConsistency = issues
End Function


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
                                 Optional ByVal autoFixSafe_ As Boolean = False, _
                                 Optional ByVal replacementText_ As String = "") As Object
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
    If autoFixSafe_ Then d("ReplacementText") = replacementText_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
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
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function
