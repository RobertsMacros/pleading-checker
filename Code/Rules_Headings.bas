Attribute VB_Name = "Rules_Headings"
' ============================================================
' Rules_Headings.bas
' Combined module for heading / title rules:
'   - Rule 04: Heading capitalisation consistency
'   - Rule 21: Title (honorific) formatting consistency
'
' Rule 04 uses LOCAL heading families rather than one global
' dominant per outline level.  Headings are grouped into
' contiguous runs separated by structural boundaries
' (appendix/schedule/annex headings, or large gaps of body
' text).  Each family is judged independently.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_CAPITALISATION As String = "heading_capitalisation"
Private Const RULE_NAME_TITLE As String = "title_formatting"

' Maximum body-text paragraphs between headings before we treat
' the next heading as a new structural family.
Private Const MAX_GAP_PARAS As Long = 40

' --------------------------------------------------------------
'  PRIVATE HELPERS  (from Rule04 - heading capitalisation)
' --------------------------------------------------------------

' -- Minor words to skip when checking Title Case ------------
Private Function GetMinorWords() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim w As Variant
    For Each w In Array("the", "a", "an", "in", "on", "at", "to", _
                        "for", "of", "and", "but", "or", "nor", _
                        "with", "by")
        d.Add CStr(w), True
    Next w
    Set GetMinorWords = d
End Function

' -- Proper nouns that are always capitalised ----------------
Private Function GetProperNouns() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim w As Variant
    For Each w In Array("Court", "Claimant", "Defendant", "Respondent", _
                        "Applicant", "Tribunal", "Parliament", "Crown", _
                        "State", "Government", "Minister")
        d.Add CStr(w), True
    Next w
    Set GetProperNouns = d
End Function

' -- Classify a heading's capitalisation pattern -------------
' Returns "ALL_CAPS", "TITLE_CASE", "SENTENCE_CASE", or "MIXED"
Private Function ClassifyCapitalisation(ByVal headingText As String) As String
    Dim cleanText As String
    Dim i As Long
    Dim ch As String
    Dim hasLower As Boolean
    Dim hasUpper As Boolean

    ' Strip trailing paragraph mark and whitespace
    cleanText = Trim$(Replace(headingText, vbCr, ""))
    cleanText = Trim$(Replace(cleanText, vbLf, ""))
    If Len(cleanText) = 0 Then
        ClassifyCapitalisation = "MIXED"
        Exit Function
    End If

    ' Check ALL CAPS: every alpha character is uppercase
    hasLower = False
    For i = 1 To Len(cleanText)
        ch = Mid$(cleanText, i, 1)
        If ch Like "[a-z]" Then
            hasLower = True
            Exit For
        End If
    Next i
    If Not hasLower Then
        ' Verify there is at least one alpha character
        hasUpper = False
        For i = 1 To Len(cleanText)
            ch = Mid$(cleanText, i, 1)
            If ch Like "[A-Z]" Then
                hasUpper = True
                Exit For
            End If
        Next i
        If hasUpper Then
            ClassifyCapitalisation = "ALL_CAPS"
            Exit Function
        End If
    End If

    ' Split into words and analyse
    Dim words() As String
    words = Split(cleanText, " ")

    Dim minorWords As Object
    Set minorWords = GetMinorWords()

    Dim properNouns As Object
    Set properNouns = GetProperNouns()

    ' Check Title Case: significant words start with uppercase
    Dim titleCaseCount As Long
    Dim significantCount As Long
    Dim wordIdx As Long

    For wordIdx = LBound(words) To UBound(words)
        Dim w As String
        w = Trim$(words(wordIdx))
        If Len(w) = 0 Then GoTo NextWordTitle

        ' Strip leading punctuation
        Dim firstAlpha As String
        Dim charPos As Long
        firstAlpha = ""
        For charPos = 1 To Len(w)
            If Mid$(w, charPos, 1) Like "[A-Za-z]" Then
                firstAlpha = Mid$(w, charPos, 1)
                Exit For
            End If
        Next charPos
        If Len(firstAlpha) = 0 Then GoTo NextWordTitle

        ' Skip minor words (except first word)
        If wordIdx > LBound(words) And minorWords.Exists(LCase(w)) Then
            GoTo NextWordTitle
        End If

        ' Skip proper nouns (always capitalised, not diagnostic)
        If properNouns.Exists(w) Then GoTo NextWordTitle

        significantCount = significantCount + 1
        If firstAlpha Like "[A-Z]" Then
            titleCaseCount = titleCaseCount + 1
        End If
NextWordTitle:
    Next wordIdx

    ' Check Sentence Case: only first word capitalised
    ' First word must start uppercase, rest should be lowercase (except proper nouns)
    Dim firstWord As String
    firstWord = ""
    For wordIdx = LBound(words) To UBound(words)
        If Len(Trim$(words(wordIdx))) > 0 Then
            firstWord = Trim$(words(wordIdx))
            Exit For
        End If
    Next wordIdx

    Dim firstCharOfFirst As String
    firstCharOfFirst = ""
    For charPos = 1 To Len(firstWord)
        If Mid$(firstWord, charPos, 1) Like "[A-Za-z]" Then
            firstCharOfFirst = Mid$(firstWord, charPos, 1)
            Exit For
        End If
    Next charPos

    Dim sentenceCaseViolations As Long
    sentenceCaseViolations = 0
    If firstCharOfFirst Like "[a-z]" Then
        ' First word not capitalised -- not sentence case
        sentenceCaseViolations = significantCount ' force fail
    Else
        ' Check that subsequent significant words start lowercase
        Dim pastFirst As Boolean
        pastFirst = False
        For wordIdx = LBound(words) To UBound(words)
            w = Trim$(words(wordIdx))
            If Len(w) = 0 Then GoTo NextWordSentence
            If Not pastFirst Then
                pastFirst = True
                GoTo NextWordSentence
            End If
            ' Skip proper nouns
            If properNouns.Exists(w) Then GoTo NextWordSentence

            firstAlpha = ""
            For charPos = 1 To Len(w)
                If Mid$(w, charPos, 1) Like "[A-Za-z]" Then
                    firstAlpha = Mid$(w, charPos, 1)
                    Exit For
                End If
            Next charPos
            If Len(firstAlpha) > 0 Then
                If firstAlpha Like "[A-Z]" Then
                    sentenceCaseViolations = sentenceCaseViolations + 1
                End If
            End If
NextWordSentence:
        Next wordIdx
    End If

    ' Determine pattern
    If significantCount > 0 And titleCaseCount = significantCount Then
        ClassifyCapitalisation = "TITLE_CASE"
    ElseIf sentenceCaseViolations = 0 Then
        ClassifyCapitalisation = "SENTENCE_CASE"
    Else
        ClassifyCapitalisation = "MIXED"
    End If
End Function

' -- Count words in a heading (excluding trailing marks) -----
Private Function CountWords(ByVal txt As String) As Long
    Dim cleanText As String
    cleanText = Trim$(Replace(txt, vbCr, ""))
    cleanText = Trim$(Replace(cleanText, vbLf, ""))
    If Len(cleanText) = 0 Then
        CountWords = 0
        Exit Function
    End If
    Dim wParts() As String
    wParts = Split(cleanText, " ")
    Dim cnt As Long
    Dim p As Variant
    For Each p In wParts
        If Len(Trim$(CStr(p))) > 0 Then cnt = cnt + 1
    Next p
    CountWords = cnt
End Function

' -- Check if a heading text indicates a structural boundary --
'  Returns True for schedule, appendix, annex, part, exhibit etc.
Private Function IsStructuralBoundary(ByVal headingText As String) As Boolean
    Dim lText As String
    lText = LCase$(Trim$(Replace(headingText, vbCr, "")))
    IsStructuralBoundary = False

    ' Check for section-divider keywords at the start
    If Left$(lText, 8) = "schedule" Or Left$(lText, 8) = "appendix" Or _
       Left$(lText, 5) = "annex" Or Left$(lText, 7) = "exhibit" Or _
       Left$(lText, 10) = "attachment" Then
        IsStructuralBoundary = True
        Exit Function
    End If

    ' Also match "SCHEDULE", "APPENDIX" etc. in ALL_CAPS
    If Left$(lText, 4) = "part" Then
        ' "Part" followed by a number or letter is structural
        If Len(lText) > 4 Then
            Dim afterPart As String
            afterPart = Trim$(Mid$(lText, 5))
            If Len(afterPart) > 0 Then
                Dim fc As String
                fc = Left$(afterPart, 1)
                If (fc >= "0" And fc <= "9") Or _
                   (fc >= "a" And fc <= "z") Then
                    IsStructuralBoundary = True
                End If
            End If
        End If
    End If
End Function

' --------------------------------------------------------------
'  PRIVATE HELPERS  (from Rule21 - title formatting)
' --------------------------------------------------------------

' -- Count occurrences of a word in the document -------------
'  Uses Find with MatchWholeWord and MatchCase.
Private Function CountWordInDoc(doc As Document, word As String) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean

    cnt = 0

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = word
        .MatchWholeWord = True
        .MatchCase = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        If EngineIsInPageRange(rng) Then
            cnt = cnt + 1
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    CountWordInDoc = cnt
End Function

' -- Flag all occurrences of a minority form -----------------
Private Sub FlagOccurrences(doc As Document, _
                             word As String, _
                             issueText As String, _
                             suggestionText As String, _
                             ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = word
        .MatchWholeWord = True
        .MatchCase = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_TITLE, locStr, issueText, suggestionText, rng.Start, rng.End, "error")
            issues.Add finding
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ============================================================
'  PUBLIC: Check heading capitalisation  (Rule 04)
'
'  LOCAL-FAMILY APPROACH:
'  1. Collect all headings into an ordered list.
'  2. Walk the ordered list and split into "families" whenever:
'     a) A structural boundary heading is encountered
'        (schedule, appendix, annex, etc.), or
'     b) More than MAX_GAP_PARAS non-heading paragraphs
'        separate two consecutive headings.
'  3. Within each family, determine dominant capitalisation
'     per outline level and flag outliers.
' ============================================================
Public Function Check_HeadingCapitalisation(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim lvl As Long

    On Error Resume Next

    ' -------------------------------------------------------
    '  PASS 1: Collect all headings into an ordered array
    ' -------------------------------------------------------
    ' Each entry: Array(paraIdx, headingText, pattern, rangeStart, rangeEnd, outlineLevel)
    Dim hCap As Long
    hCap = 128
    Dim hCount As Long
    hCount = 0
    Dim hParaIdx() As Long
    Dim hTexts() As String
    Dim hPatterns() As String
    Dim hStarts() As Long
    Dim hEnds() As Long
    Dim hLevels() As Long
    ReDim hParaIdx(0 To hCap - 1)
    ReDim hTexts(0 To hCap - 1)
    ReDim hPatterns(0 To hCap - 1)
    ReDim hStarts(0 To hCap - 1)
    ReDim hEnds(0 To hCap - 1)
    ReDim hLevels(0 To hCap - 1)

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Check if this is a heading (outline levels 1-9)
        lvl = para.OutlineLevel
        If Err.Number <> 0 Then lvl = wdOutlineLevelBodyText: Err.Clear
        If lvl >= wdOutlineLevel1 And lvl <= wdOutlineLevel9 Then

            ' Page range filter
            If Not EngineIsInPageRange(para.Range) Then GoTo NextPara

            Dim headingText As String
            headingText = para.Range.Text
            If Err.Number <> 0 Then headingText = "": Err.Clear

            ' Skip single-word headings
            If CountWords(headingText) <= 1 Then GoTo NextPara

            ' Classify capitalisation
            Dim pattern As String
            pattern = ClassifyCapitalisation(headingText)

            ' Grow arrays if needed
            If hCount >= hCap Then
                hCap = hCap * 2
                ReDim Preserve hParaIdx(0 To hCap - 1)
                ReDim Preserve hTexts(0 To hCap - 1)
                ReDim Preserve hPatterns(0 To hCap - 1)
                ReDim Preserve hStarts(0 To hCap - 1)
                ReDim Preserve hEnds(0 To hCap - 1)
                ReDim Preserve hLevels(0 To hCap - 1)
            End If

            hParaIdx(hCount) = paraIdx
            hTexts(hCount) = headingText
            hPatterns(hCount) = pattern
            hStarts(hCount) = para.Range.Start
            If Err.Number <> 0 Then hStarts(hCount) = 0: Err.Clear
            hEnds(hCount) = para.Range.End
            If Err.Number <> 0 Then hEnds(hCount) = 0: Err.Clear
            hLevels(hCount) = lvl
            hCount = hCount + 1
        End If
NextPara:
    Next para

    If hCount < 2 Then
        On Error GoTo 0
        Set Check_HeadingCapitalisation = issues
        Exit Function
    End If
    On Error GoTo 0   ' Pass 1 complete; Pass 2 is pure VBA

    ' -------------------------------------------------------
    '  PASS 2: Split headings into local families
    '
    '  familyStarts() and familyEnds() mark index ranges
    '  within the heading arrays.
    ' -------------------------------------------------------
    Dim fsCap As Long
    fsCap = 32
    Dim fsCount As Long
    fsCount = 0
    Dim familyStarts() As Long
    Dim familyEnds() As Long
    ReDim familyStarts(0 To fsCap - 1)
    ReDim familyEnds(0 To fsCap - 1)

    Dim curFamilyStart As Long
    curFamilyStart = 0

    Dim hi As Long
    For hi = 1 To hCount - 1
        Dim newFamily As Boolean
        newFamily = False

        ' Check for structural boundary
        If IsStructuralBoundary(hTexts(hi)) Then
            newFamily = True
        End If

        ' Check for large gap between consecutive headings
        If Not newFamily Then
            Dim gap As Long
            gap = hParaIdx(hi) - hParaIdx(hi - 1)
            If gap > MAX_GAP_PARAS Then
                newFamily = True
            End If
        End If

        If newFamily Then
            ' Close the current family
            If fsCount >= fsCap Then
                fsCap = fsCap * 2
                ReDim Preserve familyStarts(0 To fsCap - 1)
                ReDim Preserve familyEnds(0 To fsCap - 1)
            End If
            familyStarts(fsCount) = curFamilyStart
            familyEnds(fsCount) = hi - 1
            fsCount = fsCount + 1
            curFamilyStart = hi
        End If
    Next hi

    ' Close the last family
    If fsCount >= fsCap Then
        fsCap = fsCap * 2
        ReDim Preserve familyStarts(0 To fsCap - 1)
        ReDim Preserve familyEnds(0 To fsCap - 1)
    End If
    familyStarts(fsCount) = curFamilyStart
    familyEnds(fsCount) = hCount - 1
    fsCount = fsCount + 1

    ' -------------------------------------------------------
    '  PASS 3: Within each family, find dominant per level
    '  and flag outliers
    ' -------------------------------------------------------
    Dim fi As Long
    For fi = 0 To fsCount - 1
        Dim fStart As Long
        fStart = familyStarts(fi)
        Dim fEnd As Long
        fEnd = familyEnds(fi)

        ' Build pattern counts per level within this family
        Dim levelPats As Object
        Set levelPats = CreateObject("Scripting.Dictionary")
        ' levelPats: level -> Dictionary(pattern -> count)

        Dim hj As Long
        For hj = fStart To fEnd
            lvl = hLevels(hj)
            If Not levelPats.Exists(lvl) Then
                levelPats.Add lvl, CreateObject("Scripting.Dictionary")
            End If
            Dim patDict As Object
            Set patDict = levelPats(lvl)
            If patDict.Exists(hPatterns(hj)) Then
                patDict(hPatterns(hj)) = patDict(hPatterns(hj)) + 1
            Else
                patDict.Add hPatterns(hj), 1
            End If
        Next hj

        ' For each level in this family, find dominant and flag outliers
        Dim lvlKey As Variant
        For Each lvlKey In levelPats.keys
            Set patDict = levelPats(lvlKey)

            ' Count total headings at this level in this family
            Dim levelTotal As Long
            levelTotal = 0
            Dim patKey As Variant
            For Each patKey In patDict.keys
                levelTotal = levelTotal + patDict(patKey)
            Next patKey

            ' Need at least 2 headings to compare
            If levelTotal < 2 Then GoTo NextFamilyLevel

            ' Find dominant pattern
            Dim dominantPattern As String
            Dim maxCount As Long
            dominantPattern = ""
            maxCount = 0
            For Each patKey In patDict.keys
                If patDict(patKey) > maxCount Then
                    maxCount = patDict(patKey)
                    dominantPattern = CStr(patKey)
                End If
            Next patKey

            ' Flag headings in this family+level that deviate
            For hj = fStart To fEnd
                If hLevels(hj) = CLng(lvlKey) Then
                    If hPatterns(hj) <> dominantPattern Then
                        Dim finding As Object
                        Dim loc As String
                        Dim rng As Range
                        On Error Resume Next
                        Set rng = doc.Range(hStarts(hj), hEnds(hj))
                        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFamilyHeading
                        loc = EngineGetLocationString(rng, doc)
                        If Err.Number <> 0 Then loc = "unknown location": Err.Clear
                        On Error GoTo 0

                        Dim cleanHText As String
                        cleanHText = Trim$(Replace(hTexts(hj), vbCr, ""))

                        Dim suggn As String
                        Select Case dominantPattern
                            Case "ALL_CAPS"
                                suggn = "Convert to ALL CAPS to match nearby level " & CLng(lvlKey) & " headings"
                            Case "TITLE_CASE"
                                suggn = "Convert to Title Case to match nearby level " & CLng(lvlKey) & " headings"
                            Case "SENTENCE_CASE"
                                suggn = "Convert to Sentence case to match nearby level " & CLng(lvlKey) & " headings"
                            Case Else
                                suggn = "Review capitalisation for consistency with nearby level " & CLng(lvlKey) & " headings"
                        End Select

                        Set finding = CreateIssueDict(RULE_NAME_CAPITALISATION, loc, "Heading capitalisation mismatch: '" & cleanHText & "' uses " & hPatterns(hj) & " but nearby dominant pattern is " & dominantPattern, suggn, hStarts(hj), hEnds(hj), "possible_error")
                        issues.Add finding
                    End If
                End If
NextFamilyHeading:
            Next hj
NextFamilyLevel:
        Next lvlKey
    Next fi

    On Error GoTo 0
    Set Check_HeadingCapitalisation = issues
End Function

' ============================================================
'  PUBLIC: Check title formatting  (Rule 21)
' ============================================================
Public Function Check_TitleFormatting(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Define title pairs: noDot / withDot ------------------
    Dim noDot As Variant
    Dim withDot As Variant

    noDot = Array("Mr", "Mrs", "Ms", "Dr", "Prof", "QC", "KC", "MP", "JP")
    withDot = Array("Mr.", "Mrs.", "Ms.", "Dr.", "Prof.", "Q.C.", "K.C.", "M.P.", "J.P.")

    Dim i As Long
    Dim noDotCount As Long
    Dim withDotCount As Long

    For i = LBound(noDot) To UBound(noDot)
        noDotCount = CountWordInDoc(doc, CStr(noDot(i)))
        withDotCount = CountWordInDoc(doc, CStr(withDot(i)))

        ' Only flag if both forms exist
        If noDotCount > 0 And withDotCount > 0 Then
            If noDotCount >= withDotCount Then
                ' noDot is dominant -- flag all withDot occurrences
                FlagOccurrences doc, CStr(withDot(i)), _
                    "Inconsistent title formatting: '" & withDot(i) & "' used", _
                    "Use '" & noDot(i) & "' without full stop (dominant style)", _
                    issues
            Else
                ' withDot is dominant -- flag all noDot occurrences
                FlagOccurrences doc, CStr(noDot(i)), _
                    "Inconsistent title formatting: '" & noDot(i) & "' used", _
                    "Use '" & withDot(i) & "' with full stop (dominant style)", _
                    issues
            End If
        End If
    Next i

    Set Check_TitleFormatting = issues
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
