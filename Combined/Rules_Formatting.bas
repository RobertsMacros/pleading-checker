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

' ── Classify spacing pattern after a heading ────────────────
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

' ── Classify SpaceBefore pattern ────────────────────────────
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

' ── Helper: build a font profile key ────────────────────────
Private Function FontKey(ByVal fontName As String, ByVal fontSize As Single) As String
    FontKey = fontName & "|" & CStr(fontSize)
End Function

' ── Helper: find dominant key in a dictionary of counts ─────
Private Function GetDominant(counts As Scripting.Dictionary) As String
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

' ── Helper: parse font key back to readable description ─────
Private Function FontDescription(ByVal fKey As String) As String
    Dim parts() As String
    parts = Split(fKey, "|")
    If UBound(parts) >= 1 Then
        FontDescription = parts(0) & " " & parts(1) & "pt"
    Else
        FontDescription = fKey
    End If
End Function

' ════════════════════════════════════════════════════════════
'  RULE 06: PARAGRAPH BREAK CONSISTENCY
' ════════════════════════════════════════════════════════════
Public Function Check_ParagraphBreakConsistency(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim lvl As Long

    On Error Resume Next

    ' ── Dictionaries keyed by outline level ─────────────────
    ' afterPatterns:  level -> Dictionary(pattern -> count)
    ' beforePatterns: level -> Dictionary(pattern -> count)
    ' headingInfos:   level -> Collection of Array(paraIdx, afterPattern, beforePattern, rangeStart, rangeEnd, text)
    Dim afterPatterns As New Scripting.Dictionary
    Dim beforePatterns As New Scripting.Dictionary
    Dim headingInfos As New Scripting.Dictionary

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        lvl = para.OutlineLevel
        If lvl < wdOutlineLevel1 Or lvl > wdOutlineLevel9 Then GoTo NextPara

        ' Page range filter
        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextPara

        ' Classify after-spacing
        Dim aftPat As String
        aftPat = ClassifyAfterSpacing(para, doc, paraIdx)

        ' Classify before-spacing
        Dim befPat As String
        befPat = ClassifyBeforeSpacing(para)

        ' ── Track after-spacing counts ─────────────────────
        If Not afterPatterns.Exists(lvl) Then
            afterPatterns.Add lvl, New Scripting.Dictionary
        End If
        Dim aftDict As Scripting.Dictionary
        Set aftDict = afterPatterns(lvl)
        If aftDict.Exists(aftPat) Then
            aftDict(aftPat) = aftDict(aftPat) + 1
        Else
            aftDict.Add aftPat, 1
        End If

        ' ── Track before-spacing counts ────────────────────
        If Not beforePatterns.Exists(lvl) Then
            beforePatterns.Add lvl, New Scripting.Dictionary
        End If
        Dim befDict As Scripting.Dictionary
        Set befDict = beforePatterns(lvl)
        If befDict.Exists(befPat) Then
            befDict(befPat) = befDict(befPat) + 1
        Else
            befDict.Add befPat, 1
        End If

        ' ── Store heading info ─────────────────────────────
        If Not headingInfos.Exists(lvl) Then
            headingInfos.Add lvl, New Collection
        End If
        Dim info(0 To 5) As Variant
        info(0) = paraIdx
        info(1) = aftPat
        info(2) = befPat
        info(3) = para.Range.Start
        info(4) = para.Range.End
        info(5) = Trim$(Replace(para.Range.Text, vbCr, ""))
        headingInfos(lvl).Add info
NextPara:
    Next para

    ' ── Determine dominant patterns and flag deviations ─────
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
                Dim issueA As New PleadingsIssue
                Dim rngA As Range
                Set rngA = doc.Range(CLng(hInfo(3)), CLng(hInfo(4)))
                Dim locA As String
                locA = PleadingsEngine.GetLocationString(rngA, doc)

                issueA.Init RULE_NAME_PARAGRAPH_BREAK, locA, _
                    "After-heading spacing inconsistency at '" & hText & _
                    "': uses " & hAft & " but dominant pattern for level " & _
                    CLng(lvlKey) & " headings is " & domAfter, _
                    "Change spacing after this heading to match: " & domAfter, _
                    CLng(hInfo(3)), CLng(hInfo(4)), "possible_error"
                issues.Add issueA
            End If

            ' Check before-spacing deviation
            If hBef <> domBefore And Len(domBefore) > 0 Then
                Dim issueB As New PleadingsIssue
                Dim rngB As Range
                Set rngB = doc.Range(CLng(hInfo(3)), CLng(hInfo(4)))
                Dim locB As String
                locB = PleadingsEngine.GetLocationString(rngB, doc)

                issueB.Init RULE_NAME_PARAGRAPH_BREAK, locB, _
                    "Before-heading spacing inconsistency at '" & hText & _
                    "': uses " & hBef & " but dominant pattern for level " & _
                    CLng(lvlKey) & " headings is " & domBefore, _
                    "Change spacing before this heading to match: " & domBefore, _
                    CLng(hInfo(3)), CLng(hInfo(4)), "possible_error"
                issues.Add issueB
            End If
        Next h
NextLevel:
    Next lvlKey

    On Error GoTo 0
    Set Check_ParagraphBreakConsistency = issues
End Function

' ════════════════════════════════════════════════════════════
'  RULE 11: FONT CONSISTENCY
' ════════════════════════════════════════════════════════════
Public Function Check_FontConsistency(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' ══════════════════════════════════════════════════════════
    '  PASS 1: Build font profiles per context
    ' ══════════════════════════════════════════════════════════
    Dim headingFonts As New Scripting.Dictionary  ' FontKey -> count
    Dim bodyFonts As New Scripting.Dictionary     ' FontKey -> count
    Dim footnoteFonts As New Scripting.Dictionary ' FontKey -> count

    Dim para As Paragraph
    Dim paraIdx As Long
    Dim fk As String

    ' ── Headings and body text ─────────────────────────────
    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextParaFont1

        Dim lvl As Long
        lvl = para.OutlineLevel

        ' Determine context
        Dim isHeading As Boolean
        isHeading = (lvl >= wdOutlineLevel1 And lvl <= wdOutlineLevel9)

        Dim isBody As Boolean
        isBody = (lvl = wdOutlineLevelBodyText)

        ' Get font info from the paragraph range
        Dim paraFontName As String
        Dim paraFontSize As Single
        paraFontName = para.Range.Font.Name
        paraFontSize = para.Range.Font.Size

        ' Skip if font info is indeterminate (mixed within paragraph)
        If Len(paraFontName) = 0 Or paraFontSize <= 0 Then GoTo NextParaFont1

        fk = FontKey(paraFontName, paraFontSize)

        If isHeading Then
            If headingFonts.Exists(fk) Then
                headingFonts(fk) = headingFonts(fk) + 1
            Else
                headingFonts.Add fk, 1
            End If
        ElseIf isBody Then
            If bodyFonts.Exists(fk) Then
                bodyFonts(fk) = bodyFonts(fk) + 1
            Else
                bodyFonts.Add fk, 1
            End If
        End If
NextParaFont1:
    Next para

    ' ── Footnotes ──────────────────────────────────────────
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        If Not PleadingsEngine.IsInPageRange(fn.Range) Then GoTo NextFootnote

        Dim fnFontName As String
        Dim fnFontSize As Single
        fnFontName = fn.Range.Font.Name
        fnFontSize = fn.Range.Font.Size

        If Len(fnFontName) > 0 And fnFontSize > 0 Then
            fk = FontKey(fnFontName, fnFontSize)
            If footnoteFonts.Exists(fk) Then
                footnoteFonts(fk) = footnoteFonts(fk) + 1
            Else
                footnoteFonts.Add fk, 1
            End If
        End If
NextFootnote:
    Next fn

    ' ══════════════════════════════════════════════════════════
    '  PASS 2: Determine dominant fonts per context
    ' ══════════════════════════════════════════════════════════
    Dim domHeading As String
    Dim domBody As String
    Dim domFootnote As String

    domHeading = GetDominant(headingFonts)
    domBody = GetDominant(bodyFonts)
    domFootnote = GetDominant(footnoteFonts)

    ' ══════════════════════════════════════════════════════════
    '  PASS 3: Flag deviations at paragraph and run level
    ' ══════════════════════════════════════════════════════════
    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextParaFont2

        lvl = para.OutlineLevel
        isHeading = (lvl >= wdOutlineLevel1 And lvl <= wdOutlineLevel9)
        isBody = (lvl = wdOutlineLevelBodyText)

        Dim expectedFont As String
        Dim context As String
        If isHeading And Len(domHeading) > 0 Then
            expectedFont = domHeading
            context = "heading"
        ElseIf isBody And Len(domBody) > 0 Then
            expectedFont = domBody
            context = "body"
        Else
            GoTo NextParaFont2
        End If

        ' ── Check at paragraph level ───────────────────────
        paraFontName = para.Range.Font.Name
        paraFontSize = para.Range.Font.Size

        If Len(paraFontName) > 0 And paraFontSize > 0 Then
            fk = FontKey(paraFontName, paraFontSize)
            If fk <> expectedFont Then
                Dim issuePara As New PleadingsIssue
                Dim locP As String
                locP = PleadingsEngine.GetLocationString(para.Range, doc)

                Dim cleanParaText As String
                cleanParaText = Trim$(Replace(Left$(para.Range.Text, 60), vbCr, ""))

                issuePara.Init RULE_NAME_FONT, locP, _
                    "Font inconsistency in " & context & ": '" & cleanParaText & _
                    "...' uses " & FontDescription(fk) & " but dominant " & _
                    context & " font is " & FontDescription(expectedFont), _
                    "Change to " & FontDescription(expectedFont), _
                    para.Range.Start, para.Range.End, "error"
                issues.Add issuePara
                ' Skip run-level check if paragraph-level already flagged
                GoTo NextParaFont2
            End If
        End If

        ' ── Check at run level for mid-paragraph changes ───
        Dim run As Range
        Dim runIdx As Long
        runIdx = 0
        Dim runs As Ranges

        ' Iterate through character runs in the paragraph
        Dim runRange As Range
        Set runRange = para.Range.Duplicate

        ' Use the Words/Characters approach via Runs if available
        ' VBA doesn't have a native Runs collection on Range,
        ' so we iterate using the paragraph range and check
        ' font changes character by character in blocks
        Dim runStart As Long
        Dim runEnd As Long
        Dim currentFontName As String
        Dim currentFontSize As Single
        Dim charPos As Long

        If para.Range.End - para.Range.Start > 1 Then
            runStart = para.Range.Start
            Set runRange = doc.Range(runStart, runStart + 1)
            currentFontName = runRange.Font.Name
            currentFontSize = runRange.Font.Size

            ' Scan through the paragraph in character blocks
            Dim blockSize As Long
            blockSize = 1
            For charPos = para.Range.Start + 1 To para.Range.End - 1
                Set runRange = doc.Range(charPos, charPos + 1)
                If runRange.Font.Name <> currentFontName Or _
                   runRange.Font.Size <> currentFontSize Then

                    ' End of a run — check the previous run
                    runEnd = charPos

                    ' Skip whitespace-only runs
                    Dim runText As String
                    Set runRange = doc.Range(runStart, runEnd)
                    runText = runRange.Text
                    If Len(Trim$(runText)) > 0 Then
                        ' Skip field codes
                        Dim isField As Boolean
                        isField = False
                        If runRange.Fields.Count > 0 Then isField = True

                        If Not isField Then
                            fk = FontKey(currentFontName, currentFontSize)
                            If fk <> expectedFont And Len(currentFontName) > 0 And currentFontSize > 0 Then
                                Dim issueRun As New PleadingsIssue
                                Dim locR As String
                                locR = PleadingsEngine.GetLocationString(runRange, doc)

                                Dim cleanRunText As String
                                cleanRunText = Trim$(Replace(Left$(runText, 40), vbCr, ""))

                                issueRun.Init RULE_NAME_FONT, locR, _
                                    "Mid-paragraph font change in " & context & _
                                    ": '" & cleanRunText & "' uses " & FontDescription(fk) & _
                                    " instead of " & FontDescription(expectedFont), _
                                    "Change to " & FontDescription(expectedFont), _
                                    runStart, runEnd, "error"
                                issues.Add issueRun
                                ' Only flag once per paragraph for run-level
                                GoTo NextParaFont2
                            End If
                        End If
                    End If

                    ' Start new run
                    runStart = charPos
                    Set runRange = doc.Range(charPos, charPos + 1)
                    currentFontName = runRange.Font.Name
                    currentFontSize = runRange.Font.Size
                End If
            Next charPos
        End If

NextParaFont2:
    Next para

    ' ══════════════════════════════════════════════════════════
    '  PASS 4: Check footnote font deviations
    ' ══════════════════════════════════════════════════════════
    If Len(domFootnote) > 0 Then
        For Each fn In doc.Footnotes
            If Not PleadingsEngine.IsInPageRange(fn.Range) Then GoTo NextFN2

            fnFontName = fn.Range.Font.Name
            fnFontSize = fn.Range.Font.Size

            If Len(fnFontName) > 0 And fnFontSize > 0 Then
                fk = FontKey(fnFontName, fnFontSize)
                If fk <> domFootnote Then
                    Dim issueFN As New PleadingsIssue
                    Dim locFN As String
                    locFN = PleadingsEngine.GetLocationString(fn.Range, doc)

                    Dim cleanFNText As String
                    cleanFNText = Trim$(Replace(Left$(fn.Range.Text, 50), vbCr, ""))

                    issueFN.Init RULE_NAME_FONT, locFN, _
                        "Footnote font inconsistency: '" & cleanFNText & _
                        "...' uses " & FontDescription(fk) & " but dominant " & _
                        "footnote font is " & FontDescription(domFootnote), _
                        "Change to " & FontDescription(domFootnote), _
                        fn.Range.Start, fn.Range.End, "error"
                    issues.Add issueFN
                End If
            End If
NextFN2:
        Next fn
    End If

    On Error GoTo 0
    Set Check_FontConsistency = issues
End Function
