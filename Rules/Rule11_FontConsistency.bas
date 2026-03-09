Attribute VB_Name = "Rule11_FontConsistency"
' ============================================================
' Rule11_FontConsistency.bas
' Checks font consistency across three contexts: headings,
' body text, and footnotes. Detects the dominant font profile
' (name + size) for each context and flags deviations at both
' paragraph and run level (to catch mid-paragraph changes).
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "font_consistency"

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
'  MAIN RULE FUNCTION
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

        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextPara

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
        If Len(paraFontName) = 0 Or paraFontSize <= 0 Then GoTo NextPara

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
NextPara:
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

        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextPara2

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
            GoTo NextPara2
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

                issuePara.Init RULE_NAME, locP, _
                    "Font inconsistency in " & context & ": '" & cleanParaText & _
                    "...' uses " & FontDescription(fk) & " but dominant " & _
                    context & " font is " & FontDescription(expectedFont), _
                    "Change to " & FontDescription(expectedFont), _
                    para.Range.Start, para.Range.End, "error"
                issues.Add issuePara
                ' Skip run-level check if paragraph-level already flagged
                GoTo NextPara2
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

                                issueRun.Init RULE_NAME, locR, _
                                    "Mid-paragraph font change in " & context & _
                                    ": '" & cleanRunText & "' uses " & FontDescription(fk) & _
                                    " instead of " & FontDescription(expectedFont), _
                                    "Change to " & FontDescription(expectedFont), _
                                    runStart, runEnd, "error"
                                issues.Add issueRun
                                ' Only flag once per paragraph for run-level
                                GoTo NextPara2
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

NextPara2:
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

                    issueFN.Init RULE_NAME, locFN, _
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

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunFontConsistency()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Font Consistency"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_FontConsistency(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Font Consistency"
End Sub
