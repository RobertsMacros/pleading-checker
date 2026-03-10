Attribute VB_Name = "Rules_Headings"
' ============================================================
' Rules_Headings.bas
' Combined module for heading / title rules:
'   - Rule 04: Heading capitalisation consistency
'   - Rule 21: Title (honorific) formatting consistency
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_CAPITALISATION As String = "heading_capitalisation"
Private Const RULE_NAME_TITLE As String = "title_formatting"

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
    Dim parts() As String
    parts = Split(cleanText, " ")
    Dim cnt As Long
    Dim p As Variant
    For Each p In parts
        If Len(Trim$(CStr(p))) > 0 Then cnt = cnt + 1
    Next p
    CountWords = cnt
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
    Dim useWildcards As Boolean

    cnt = 0

    ' For dotted abbreviations like "Q.C.", use non-wildcard search
    ' For simple words like "Mr", use whole-word matching
    useWildcards = False

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
    Dim issue As Object
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

            Set issue = CreateIssueDict(RULE_NAME_TITLE, locStr, issueText, suggestionText, rng.Start, rng.End, "error")
            issues.Add issue
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ============================================================
'  PUBLIC: Check heading capitalisation  (Rule 04)
' ============================================================
Public Function Check_HeadingCapitalisation(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim lvl As Long

    On Error Resume Next

    ' -- Dictionaries keyed by outline level -----------------
    ' levelPatterns: level -> Dictionary(pattern -> count)
    ' levelHeadings: level -> Collection of Array(paraIdx, text, pattern, rangeStart, rangeEnd)
    Dim levelPatterns As Object
    Set levelPatterns = CreateObject("Scripting.Dictionary")
    Dim levelHeadings As Object
    Set levelHeadings = CreateObject("Scripting.Dictionary")

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Check if this is a heading (outline levels 1-9)
        lvl = para.OutlineLevel
        If lvl >= wdOutlineLevel1 And lvl <= wdOutlineLevel9 Then

            ' Page range filter
            If Not EngineIsInPageRange(para.Range) Then GoTo NextPara

            Dim headingText As String
            headingText = para.Range.Text

            ' Skip single-word headings
            If CountWords(headingText) <= 1 Then GoTo NextPara

            ' Classify capitalisation
            Dim pattern As String
            pattern = ClassifyCapitalisation(headingText)

            ' Store pattern count per level
            If Not levelPatterns.Exists(lvl) Then
                levelPatterns.Add lvl, CreateObject("Scripting.Dictionary")
            End If
            Dim patDict As Object
            Set patDict = levelPatterns(lvl)
            If patDict.Exists(pattern) Then
                patDict(pattern) = patDict(pattern) + 1
            Else
                patDict.Add pattern, 1
            End If

            ' Store heading info per level
            If Not levelHeadings.Exists(lvl) Then
                levelHeadings.Add lvl, New Collection
            End If
            Dim info(0 To 4) As Variant
            info(0) = paraIdx
            info(1) = headingText
            info(2) = pattern
            info(3) = para.Range.Start
            info(4) = para.Range.End
            levelHeadings(lvl).Add info
        End If
NextPara:
    Next para

    ' -- Determine dominant pattern per level and flag outliers --
    Dim lvlKey As Variant
    For Each lvlKey In levelPatterns.keys
        Set patDict = levelPatterns(lvlKey)

        ' Find dominant pattern
        Dim dominantPattern As String
        Dim maxCount As Long
        Dim patKey As Variant
        dominantPattern = ""
        maxCount = 0
        For Each patKey In patDict.keys
            If patDict(patKey) > maxCount Then
                maxCount = patDict(patKey)
                dominantPattern = CStr(patKey)
            End If
        Next patKey

        ' Skip if only one heading at this level
        If Not levelHeadings.Exists(lvlKey) Then GoTo NextLevel
        Dim headings As Collection
        Set headings = levelHeadings(lvlKey)
        If headings.Count <= 1 Then GoTo NextLevel

        ' Flag headings that don't match dominant pattern
        Dim h As Long
        For h = 1 To headings.Count
            Dim hInfo As Variant
            hInfo = headings(h)
            Dim hPattern As String
            hPattern = CStr(hInfo(2))
            If hPattern <> dominantPattern Then
                Dim issue As Object
                Dim loc As String
                Dim rng As Range
                Set rng = doc.Range(CLng(hInfo(3)), CLng(hInfo(4)))
                loc = EngineGetLocationString(rng, doc)

                Dim cleanHText As String
                cleanHText = Trim$(Replace(CStr(hInfo(1)), vbCr, ""))

                Dim suggestion As String
                Select Case dominantPattern
                    Case "ALL_CAPS"
                        suggestion = "Convert to ALL CAPS to match other level " & CLng(lvlKey) & " headings"
                    Case "TITLE_CASE"
                        suggestion = "Convert to Title Case to match other level " & CLng(lvlKey) & " headings"
                    Case "SENTENCE_CASE"
                        suggestion = "Convert to Sentence case to match other level " & CLng(lvlKey) & " headings"
                    Case Else
                        suggestion = "Review capitalisation for consistency with other level " & CLng(lvlKey) & " headings"
                End Select

                Set issue = CreateIssueDict(RULE_NAME_CAPITALISATION, loc, "Heading capitalisation mismatch:)
                issues.Add issue
            End If
        Next h
NextLevel:
    Next lvlKey

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
                    "Use '" & noDot(i) & "' without period (dominant style)", _
                    issues
            Else
                ' withDot is dominant -- flag all noDot occurrences
                FlagOccurrences doc, CStr(noDot(i)), _
                    "Inconsistent title formatting: '" & noDot(i) & "' used", _
                    "Use '" & withDot(i) & "' with period (dominant style)", _
                    issues
            End If
        End If
    Next i

    Set Check_TitleFormatting = issues
End Function

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineIsInPageRange
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineGetLocationString
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based issue (no class dependency)
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
