Attribute VB_Name = "Rule06_ParagraphBreakConsistency"
' ============================================================
' Rule06_ParagraphBreakConsistency.bas
' Checks that headings at each outline level use consistent
' spacing: SpaceAfter, SpaceBefore, and whether manual double
' breaks (empty paragraphs) are used after headings.
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "paragraph_break_consistency"

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

' ════════════════════════════════════════════════════════════
'  MAIN RULE FUNCTION
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

                issueA.Init RULE_NAME, locA, _
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

                issueB.Init RULE_NAME, locB, _
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
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunParagraphBreakConsistency()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Paragraph Break Consistency"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_ParagraphBreakConsistency(doc)

    ' ── Highlight issues in document ─────────────────────────
    Dim iss As PleadingsIssue
    Dim rng As Range
    Dim i As Long
    For i = 1 To issues.Count
        Set iss = issues(i)
        If iss.RangeStart >= 0 And iss.RangeEnd > iss.RangeStart Then
            On Error Resume Next
            Set rng = doc.Range(iss.RangeStart, iss.RangeEnd)
            rng.HighlightColorIndex = wdYellow
            doc.Comments.Add Range:=rng, _
                Text:="[" & iss.RuleName & "] " & iss.Issue & _
                      " " & Chr(8212) & " Suggestion: " & iss.Suggestion
            On Error GoTo 0
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Paragraph Break Consistency"
End Sub
