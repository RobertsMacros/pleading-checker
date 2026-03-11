Attribute VB_Name = "Rules_Formatting"
' ============================================================
' Rules_Formatting.bas
' Combined module for formatting-related rules:
'   - Rule06: Paragraph break consistency (headings)
'   - Rule11: Font consistency (headings, body, footnotes)
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
'  PRIVATE: Detect block quote / indented extract paragraphs.
'  Checks style name and left indentation.
' ------------------------------------------------------------
Public Function IsBlockQuotePara(para As Paragraph) As Boolean
    IsBlockQuotePara = False
    On Error Resume Next

    ' Check style name for quote/block/extract keywords
    Dim sn As String
    sn = LCase(para.Style.NameLocal)
    If Err.Number <> 0 Then sn = "": Err.Clear
    If InStr(sn, "quote") > 0 Or InStr(sn, "block") > 0 Or _
       InStr(sn, "extract") > 0 Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' Check for significant left indentation (> 36pt = 0.5 inch)
    ' Block quotes typically have extra indentation on both sides
    Dim leftInd As Single
    leftInd = para.Format.LeftIndent
    If Err.Number <> 0 Then leftInd = 0: Err.Clear

    ' Get font size — for mixed-format paragraphs, Font.Size returns
    ' wdUndefined (9999999). In that case, sample the first run's font size.
    Dim fontSize As Single
    fontSize = para.Range.Font.Size
    If Err.Number <> 0 Then fontSize = 0: Err.Clear
    If fontSize <= 0 Or fontSize > 1000 Then
        ' Mixed formatting — sample first character's font size
        Dim sampleRng As Range
        Set sampleRng = para.Range.Duplicate
        If Err.Number = 0 Then
            sampleRng.Collapse wdCollapseStart
            sampleRng.MoveEnd wdCharacter, 1
            fontSize = sampleRng.Font.Size
            If Err.Number <> 0 Then fontSize = 0: Err.Clear
            If fontSize > 1000 Then fontSize = 0
        Else
            Err.Clear
        End If
    End If

    ' Check if paragraph text starts/ends with quotation marks
    ' (strong block-quote indicator when combined with indentation)
    Dim pText As String
    pText = Replace(Replace(Replace(para.Range.Text, vbCr, ""), vbTab, ""), ChrW(160), "")
    pText = Trim$(pText)
    If Err.Number <> 0 Then pText = "": Err.Clear
    Dim startsWithQuote As Boolean
    Dim endsWithQuote As Boolean
    startsWithQuote = False
    endsWithQuote = False
    If Len(pText) > 1 Then
        Dim fc As String
        Dim lc As String
        fc = Left$(pText, 1)
        lc = Right$(pText, 1)
        startsWithQuote = (fc = Chr(34) Or fc = ChrW(8220) Or fc = ChrW(8216))
        endsWithQuote = (lc = Chr(34) Or lc = ChrW(8221) Or lc = ChrW(8217))
    End If

    ' Block quote if significantly indented AND smaller font
    If leftInd > 36 And fontSize > 0 And fontSize < 11 Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' Block quote if indented at all and font is noticeably smaller
    If leftInd > 18 And fontSize > 0 And fontSize < 10 Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' Block quote if indented and wrapped in quotation marks
    If leftInd > 18 And startsWithQuote And endsWithQuote Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' Block quote if indented and wrapped in quotation marks with smaller font
    If leftInd > 18 And fontSize > 0 And fontSize < 12 And _
       (startsWithQuote Or endsWithQuote) Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' Block quote if very significantly indented (>72pt = 1 inch),
    ' regardless of font size — heavy indentation alone signals a quote
    If leftInd > 72 Then
        IsBlockQuotePara = True
    End If

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

    On Error GoTo 0
    Set Check_ParagraphBreakConsistency = issues
End Function

' ============================================================
'  RULE 11: FONT CONSISTENCY
'  Type-based approach: classify paragraphs into heading,
'  body, block-quote; compute dominant font per type;
'  flag outliers within each type.
' ============================================================
Public Function Check_FontConsistency(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' ==========================================================
    '  PASS 1a: Quick scan of body-text paragraphs to determine
    '  the dominant body indent and font size (needed later to
    '  classify block-quote paragraphs by relative metrics).
    ' ==========================================================
    Dim bodyIndents As Object   ' LeftIndent (rounded) -> count
    Set bodyIndents = CreateObject("Scripting.Dictionary")
    Dim bodySizes As Object     ' FontSize (rounded) -> count
    Set bodySizes = CreateObject("Scripting.Dictionary")

    Dim para As Paragraph
    Dim paraIdx As Long
    Dim fk As String

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1
        If Not EngineIsInPageRange(para.Range) Then GoTo NextPre
        Dim preLvl As Long
        preLvl = para.OutlineLevel
        If Err.Number <> 0 Then Err.Clear: GoTo NextPre
        If preLvl <> wdOutlineLevelBodyText Then GoTo NextPre

        Dim preInd As Single
        preInd = para.Format.LeftIndent
        If Err.Number <> 0 Then preInd = 0: Err.Clear

        Dim preSize As Single
        preSize = para.Range.Font.Size
        If Err.Number <> 0 Then preSize = 0: Err.Clear
        If preSize <= 0 Or preSize > 1000 Then GoTo NextPre

        ' Tally indent (round to nearest pt)
        Dim indKey As String
        indKey = CStr(CLng(preInd))
        If bodyIndents.Exists(indKey) Then
            bodyIndents(indKey) = bodyIndents(indKey) + 1
        Else
            bodyIndents.Add indKey, 1
        End If

        ' Tally font size
        Dim szKey As String
        szKey = CStr(CLng(preSize * 10))  ' tenths of a point for precision
        If bodySizes.Exists(szKey) Then
            bodySizes(szKey) = bodySizes(szKey) + 1
        Else
            bodySizes.Add szKey, 1
        End If
NextPre:
    Next para

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

    ' ==========================================================
    '  PASS 1b: Classify each paragraph into a type and collect
    '  font tallies per type.
    '  Types: "heading", "body", "block_quote"
    ' ==========================================================
    Dim headingFonts As Object
    Set headingFonts = CreateObject("Scripting.Dictionary")
    Dim bodyFonts As Object
    Set bodyFonts = CreateObject("Scripting.Dictionary")
    Dim bqFonts As Object
    Set bqFonts = CreateObject("Scripting.Dictionary")
    Dim footnoteFonts As Object
    Set footnoteFonts = CreateObject("Scripting.Dictionary")

    ' Store each paragraph's assigned type for Pass 3
    ' Key = paraIndex (1-based), Value = "heading"/"body"/"block_quote"/""
    Dim paraTypes As Object
    Set paraTypes = CreateObject("Scripting.Dictionary")

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1
        paraTypes.Add CStr(paraIdx), ""

        If Not EngineIsInPageRange(para.Range) Then GoTo NextParaFont1

        Dim lvl As Long
        lvl = para.OutlineLevel
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaFont1

        ' --- Determine type ---
        Dim paraType As String
        paraType = ""

        Dim isHeading As Boolean
        isHeading = (lvl >= wdOutlineLevel1 And lvl <= wdOutlineLevel9)

        If isHeading Then
            paraType = "heading"
        Else
            ' Check if this is a block quote (style-based or metric-based)
            Dim bqByStyle As Boolean
            bqByStyle = False
            Dim sn As String
            sn = ""
            sn = LCase(para.Style.NameLocal)
            If Err.Number <> 0 Then sn = "": Err.Clear
            If InStr(sn, "quote") > 0 Or InStr(sn, "block") > 0 Or _
               InStr(sn, "extract") > 0 Then
                bqByStyle = True
            End If

            If bqByStyle Then
                paraType = "block_quote"
            ElseIf lvl = wdOutlineLevelBodyText Then
                ' Check relative indentation and font size vs body dominant
                Dim pInd As Single
                pInd = para.Format.LeftIndent
                If Err.Number <> 0 Then pInd = 0: Err.Clear

                Dim pSize As Single
                pSize = para.Range.Font.Size
                If Err.Number <> 0 Then pSize = 0: Err.Clear
                ' Handle mixed formatting — sample first character
                If pSize <= 0 Or pSize > 1000 Then
                    Dim smpRng As Range
                    Set smpRng = para.Range.Duplicate
                    If Err.Number = 0 Then
                        smpRng.Collapse wdCollapseStart
                        smpRng.MoveEnd wdCharacter, 1
                        pSize = smpRng.Font.Size
                        If Err.Number <> 0 Then pSize = 0: Err.Clear
                        If pSize > 1000 Then pSize = 0
                    Else
                        Err.Clear
                    End If
                End If

                ' Block quote if: indented beyond body median + 18pt
                ' AND font smaller than body dominant font size
                Dim isExtraIndented As Boolean
                isExtraIndented = (pInd > domBodyIndent + 18)

                Dim isSmallerFont As Boolean
                isSmallerFont = (pSize > 0 And domBodySize > 0 And pSize < domBodySize - 0.5)

                ' Also check quotation marks (with tab stripping)
                Dim pText As String
                pText = ""
                pText = para.Range.Text
                If Err.Number <> 0 Then pText = "": Err.Clear
                pText = Replace(Replace(Replace(pText, vbCr, ""), vbTab, ""), ChrW(160), "")
                pText = Trim$(pText)
                Dim hasCurlyQuotes As Boolean
                hasCurlyQuotes = False
                If Len(pText) > 2 Then
                    Dim fcCh As String
                    Dim lcCh As String
                    fcCh = Left$(pText, 1)
                    lcCh = Right$(pText, 1)
                    If (fcCh = ChrW(8220) Or fcCh = Chr(34)) And _
                       (lcCh = ChrW(8221) Or lcCh = Chr(34)) Then
                        hasCurlyQuotes = True
                    End If
                End If

                If isExtraIndented And isSmallerFont Then
                    paraType = "block_quote"
                ElseIf isExtraIndented And hasCurlyQuotes Then
                    paraType = "block_quote"
                ElseIf isSmallerFont And hasCurlyQuotes And pInd > domBodyIndent Then
                    paraType = "block_quote"
                ElseIf pInd > domBodyIndent + 54 Then
                    ' Very heavy indentation alone (>~0.75 inch beyond body)
                    paraType = "block_quote"
                Else
                    paraType = "body"
                End If
            End If
        End If

        paraTypes(CStr(paraIdx)) = paraType

        ' --- Collect font tally for this type ---
        Dim paraFontName As String
        Dim paraFontSize As Single
        paraFontName = para.Range.Font.Name
        If Err.Number <> 0 Then paraFontName = "": Err.Clear
        paraFontSize = para.Range.Font.Size
        If Err.Number <> 0 Then paraFontSize = 0: Err.Clear

        ' Skip if font info is indeterminate (mixed within paragraph)
        If Len(paraFontName) = 0 Or paraFontSize <= 0 Or paraFontSize > 1000 Then GoTo NextParaFont1

        fk = FontKey(paraFontName, paraFontSize)

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
NextParaFont1:
    Next para

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
    '  PASS 3: Flag deviations at paragraph and run level
    ' ==========================================================
    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        If Not EngineIsInPageRange(para.Range) Then GoTo NextParaFont2

        Dim curType As String
        curType = ""
        If paraTypes.Exists(CStr(paraIdx)) Then curType = paraTypes(CStr(paraIdx))
        If Len(curType) = 0 Then GoTo NextParaFont2

        Dim expectedFont As String
        Dim context As String
        expectedFont = ""
        context = ""

        Select Case curType
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

        ' -- Check at paragraph level -----------------------
        paraFontName = para.Range.Font.Name
        If Err.Number <> 0 Then paraFontName = "": Err.Clear
        paraFontSize = para.Range.Font.Size
        If Err.Number <> 0 Then paraFontSize = 0: Err.Clear

        If Len(paraFontName) > 0 And paraFontSize > 0 And paraFontSize < 1000 Then
            fk = FontKey(paraFontName, paraFontSize)
            If fk <> expectedFont Then
                Dim findingPara As Object
                Dim locP As String
                locP = EngineGetLocationString(para.Range, doc)

                Dim cleanParaText As String
                cleanParaText = Trim$(Replace(Left$(para.Range.Text, 60), vbCr, ""))

                Set findingPara = CreateIssueDict(RULE_NAME_FONT, locP, _
                    "Font inconsistency in " & context & ": '" & cleanParaText & _
                    "...' uses " & FontDescription(fk) & " but dominant " & _
                    context & " font is " & FontDescription(expectedFont), _
                    "Change to " & FontDescription(expectedFont), _
                    para.Range.Start, para.Range.End, "error")
                issues.Add findingPara
                ' Skip run-level check if paragraph-level already flagged
                GoTo NextParaFont2
            End If
        End If

        ' -- Check at run level for mid-paragraph changes ---
        ' Walk formatting runs using wdCharacterFormatting (fast)
        Dim runRange As Range
        Dim runText As String
        Dim isField As Boolean

        If para.Range.End - para.Range.Start > 1 Then
            Set runRange = para.Range.Duplicate
            runRange.Collapse wdCollapseStart

            On Error Resume Next
            Do While runRange.Start < para.Range.End
                runRange.MoveEnd wdCharacterFormatting, 1
                If runRange.Start >= para.Range.End Then Exit Do

                Err.Clear
                runText = runRange.Text
                If Err.Number <> 0 Then Err.Clear: GoTo AdvanceFontRun

                ' Skip whitespace-only runs
                If Len(Trim$(Replace(Replace(runText, vbCr, ""), vbLf, ""))) = 0 Then
                    GoTo AdvanceFontRun
                End If

                ' Skip field codes
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

NextParaFont2:
    Next para

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
