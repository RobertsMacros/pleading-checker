Attribute VB_Name = "Rule04_HeadingCapitalisation"
' ============================================================
' Rule04_HeadingCapitalisation.bas
' Checks that headings at each outline level use a consistent
' capitalisation pattern (ALL CAPS, Title Case, or Sentence case).
' Flags outliers that deviate from the dominant pattern at their level.
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "heading_capitalisation"

' ── Minor words to skip when checking Title Case ────────────
Private Function GetMinorWords() As Scripting.Dictionary
    Dim d As New Scripting.Dictionary
    Dim w As Variant
    For Each w In Array("the", "a", "an", "in", "on", "at", "to", _
                        "for", "of", "and", "but", "or", "nor", _
                        "with", "by")
        d.Add CStr(w), True
    Next w
    Set GetMinorWords = d
End Function

' ── Proper nouns that are always capitalised ────────────────
Private Function GetProperNouns() As Scripting.Dictionary
    Dim d As New Scripting.Dictionary
    Dim w As Variant
    For Each w In Array("Court", "Claimant", "Defendant", "Respondent", _
                        "Applicant", "Tribunal", "Parliament", "Crown", _
                        "State", "Government", "Minister")
        d.Add CStr(w), True
    Next w
    Set GetProperNouns = d
End Function

' ── Classify a heading's capitalisation pattern ─────────────
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

    Dim minorWords As Scripting.Dictionary
    Set minorWords = GetMinorWords()

    Dim properNouns As Scripting.Dictionary
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
        ' First word not capitalised — not sentence case
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

' ── Count words in a heading (excluding trailing marks) ─────
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

' ════════════════════════════════════════════════════════════
'  MAIN RULE FUNCTION
' ════════════════════════════════════════════════════════════
Public Function Check_HeadingCapitalisation(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim lvl As Long

    On Error Resume Next

    ' ── Dictionaries keyed by outline level ─────────────────
    ' levelPatterns: level -> Dictionary(pattern -> count)
    ' levelHeadings: level -> Collection of Array(paraIdx, text, pattern, rangeStart, rangeEnd)
    Dim levelPatterns As New Scripting.Dictionary
    Dim levelHeadings As New Scripting.Dictionary

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Check if this is a heading (outline levels 1-9)
        lvl = para.OutlineLevel
        If lvl >= wdOutlineLevel1 And lvl <= wdOutlineLevel9 Then

            ' Page range filter
            If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextPara

            Dim headingText As String
            headingText = para.Range.Text

            ' Skip single-word headings
            If CountWords(headingText) <= 1 Then GoTo NextPara

            ' Classify capitalisation
            Dim pattern As String
            pattern = ClassifyCapitalisation(headingText)

            ' Store pattern count per level
            If Not levelPatterns.Exists(lvl) Then
                levelPatterns.Add lvl, New Scripting.Dictionary
            End If
            Dim patDict As Scripting.Dictionary
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

    ' ── Determine dominant pattern per level and flag outliers ──
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
                Dim issue As New PleadingsIssue
                Dim loc As String
                Dim rng As Range
                Set rng = doc.Range(CLng(hInfo(3)), CLng(hInfo(4)))
                loc = PleadingsEngine.GetLocationString(rng, doc)

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

                issue.Init RULE_NAME, loc, _
                           "Heading capitalisation mismatch: '" & cleanHText & _
                           "' uses " & hPattern & " but dominant pattern is " & dominantPattern, _
                           suggestion, CLng(hInfo(3)), CLng(hInfo(4)), "possible_error"
                issues.Add issue
            End If
        Next h
NextLevel:
    Next lvlKey

    On Error GoTo 0
    Set Check_HeadingCapitalisation = issues
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunHeadingCapitalisation()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Heading Capitalisation"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_HeadingCapitalisation(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Heading Capitalisation"
End Sub
