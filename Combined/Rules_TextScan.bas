Attribute VB_Name = "Rules_TextScan"
' ============================================================
' Rules_TextScan.bas
' Combined text-scanning proofreading rules:
'   - Check_RepeatedWords (from Rule02)
'   - Check_SpellOutUnderTen (from Rule34)
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_REPEATED As String = "repeated_words"
Private Const RULE_NAME_SPELL_OUT As String = "spell_out_under_ten"

' ============================================================
'  PUBLIC: Check_RepeatedWords
'  Detects consecutive repeated words (e.g. "the the").
'  Known-valid repetitions (e.g. "that that", "had had") are
'  flagged as "possible_error" rather than "error".
' ============================================================
Public Function Check_RepeatedWords(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim words() As String
    Dim cleanWords() As String
    Dim wordCount As Long
    Dim i As Long
    Dim prevWord As String
    Dim currWord As String
    Dim severity As String
    Dim issueText As String
    Dim suggestion As String
    Dim locStr As String
    Dim charPos As Long
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim finding As Object
    Dim paraRange As Range

    ' -- Known-valid repetitions that may be intentional ---
    ' These get flagged as "possible_error" with a note
    ' to review context, rather than a hard "error".
    Dim knownValid As Variant
    knownValid = Array("that", "had", "is", "was", "can")

    ' -- Iterate all paragraphs ----------------------------
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph_RW
        End If

        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextParagraph_RW
        End If

        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph_RW
        End If

        ' Skip very short or empty paragraphs
        If Len(Trim(paraText)) < 3 Then
            GoTo NextParagraph_RW
        End If

        ' -- Split paragraph into words --------------------
        words = Split(paraText, " ")
        wordCount = UBound(words) - LBound(words) + 1

        If wordCount < 2 Then
            GoTo NextParagraph_RW
        End If

        ' -- Clean words: strip leading/trailing punctuation -
        ReDim cleanWords(LBound(words) To UBound(words))
        For i = LBound(words) To UBound(words)
            cleanWords(i) = StripPunctuation(words(i))
        Next i

        ' -- Compare consecutive words ---------------------
        prevWord = ""
        For i = LBound(cleanWords) To UBound(cleanWords)
            currWord = LCase(cleanWords(i))

            ' Skip empty tokens (multiple spaces, punctuation-only)
            If Len(currWord) = 0 Then
                prevWord = ""
                GoTo NextWordInPara
            End If

            ' Check for repetition
            If currWord = prevWord And Len(currWord) > 0 Then

                ' -- Determine severity ----------------
                If IsKnownValidRepetition(currWord, knownValid) Then
                    severity = "possible_error"
                    issueText = "Repeated word '" & cleanWords(i) & "' " & _
                                "-- review context; may be intentional"
                Else
                    severity = "error"
                    issueText = "Repeated word '" & cleanWords(i) & "' detected"
                End If

                suggestion = "Remove the duplicate '" & cleanWords(i) & "'"

                ' -- Calculate character position within paragraph -
                ' Find the second occurrence of the repeated word
                ' by searching after the first occurrence
                Dim searchStart As Long
                Dim firstPos As Long
                Dim secondPos As Long

                firstPos = InStr(1, paraText, words(i - 1), vbTextCompare)
                If firstPos > 0 Then
                    secondPos = InStr(firstPos + Len(words(i - 1)), _
                                      paraText, words(i), vbTextCompare)
                Else
                    secondPos = 0
                End If

                ' Map to document-level character positions
                If secondPos > 0 Then
                    rangeStart = paraRange.Start + secondPos - 1
                    rangeEnd = rangeStart + Len(words(i))
                Else
                    ' Fallback: use paragraph start
                    rangeStart = paraRange.Start
                    rangeEnd = paraRange.End
                End If

                ' -- Build location string -------------
                Err.Clear
                Dim matchRange As Range
                Set matchRange = doc.Range(rangeStart, rangeEnd)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                Else
                    locStr = EngineGetLocationString(matchRange, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If
                End If

                ' -- Create the finding ------------------
                Set finding = CreateIssueDict(RULE_NAME_REPEATED, locStr, issueText, suggestion, rangeStart, rangeEnd, severity)
                issues.Add finding
            End If

            prevWord = currWord

NextWordInPara:
        Next i

NextParagraph_RW:
    Next para
    On Error GoTo 0

    Set Check_RepeatedWords = issues
End Function

' ============================================================
'  PUBLIC: Check_SpellOutUnderTen
'  In running prose, numbers under 10 should be written in
'  words (e.g. "seven" instead of "7").
' ============================================================
Public Function Check_SpellOutUnderTen(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim styleName As String
    Dim i As Long
    Dim ch As String
    Dim digitVal As Long
    Dim finding As Object
    Dim locStr As String
    Dim charRange As Range
    Dim textLen As Long

    ' Number word map
    Dim numberWords(0 To 9) As String
    numberWords(0) = "zero"
    numberWords(1) = "one"
    numberWords(2) = "two"
    numberWords(3) = "three"
    numberWords(4) = "four"
    numberWords(5) = "five"
    numberWords(6) = "six"
    numberWords(7) = "seven"
    numberWords(8) = "eight"
    numberWords(9) = "nine"

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph_SO
        End If

        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextParagraph_SO
        End If

        ' -- Check paragraph style for exclusions ------------
        styleName = ""
        styleName = paraRange.ParagraphStyle
        If Err.Number <> 0 Then
            Err.Clear
            styleName = ""
        End If

        If IsExcludedStyle(styleName) Then
            GoTo NextParagraph_SO
        End If

        ' -- Get paragraph text ------------------------------
        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph_SO
        End If

        textLen = Len(paraText)
        If textLen = 0 Then GoTo NextParagraph_SO

        ' -- Scan character by character for digits 0-9 ------
        For i = 1 To textLen
            ch = Mid(paraText, i, 1)

            ' Check if character is a digit 0-9
            If ch >= "0" And ch <= "9" Then
                digitVal = CInt(ch)

                ' -- Check: isolated digit (not part of larger number) --
                If IsPartOfLargerNumber(paraText, i, textLen) Then
                    GoTo NextChar
                End If

                ' -- Check: preceded by structural reference word --
                If IsPrecededByStructuralRef(paraText, i) Then
                    GoTo NextChar
                End If

                ' -- Check: part of a range pattern --
                If IsPartOfRange(paraText, i, textLen) Then
                    GoTo NextChar
                End If

                ' -- Check: citation context --
                If IsInCitationContext(paraText, i) Then
                    GoTo NextChar
                End If

                ' -- Check: preceded by currency/unit symbols --
                If IsPrecededByCurrencyOrUnit(paraText, i) Then
                    GoTo NextChar
                End If

                ' -- All checks passed: flag this digit ------
                Dim rangeStart As Long
                Dim rangeEnd As Long

                rangeStart = paraRange.Start + i - 1
                rangeEnd = rangeStart + 1

                Err.Clear
                Set charRange = doc.Range(rangeStart, rangeEnd)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                Else
                    locStr = EngineGetLocationString(charRange, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If
                End If

                Set finding = CreateIssueDict(RULE_NAME_SPELL_OUT, locStr, "Number under 10 is given as a figure in running prose.", "Write)
                issues.Add finding
            End If

NextChar:
        Next i

NextParagraph_SO:
    Next para
    On Error GoTo 0

    Set Check_SpellOutUnderTen = issues
End Function

' ============================================================
'  HELPERS FOR Check_RepeatedWords
' ============================================================

' ------------------------------------------------------------
'  PRIVATE: Strip leading and trailing punctuation from a word
'  Removes characters like . , ; : ! ? " ' ( ) [ ] etc.
' ------------------------------------------------------------
Private Function StripPunctuation(ByVal word As String) As String
    Dim ch As String
    Dim startPos As Long
    Dim endPos As Long

    word = Trim(word)
    If Len(word) = 0 Then
        StripPunctuation = ""
        Exit Function
    End If

    ' Strip from start
    startPos = 1
    Do While startPos <= Len(word)
        ch = Mid(word, startPos, 1)
        If IsPunctuation(ch) Then
            startPos = startPos + 1
        Else
            Exit Do
        End If
    Loop

    ' Strip from end
    endPos = Len(word)
    Do While endPos >= startPos
        ch = Mid(word, endPos, 1)
        If IsPunctuation(ch) Then
            endPos = endPos - 1
        Else
            Exit Do
        End If
    Loop

    If startPos > endPos Then
        StripPunctuation = ""
    Else
        StripPunctuation = Mid(word, startPos, endPos - startPos + 1)
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if a character is punctuation
' ------------------------------------------------------------
Private Function IsPunctuation(ByVal ch As String) As Boolean
    Const PUNCT_CHARS As String = ".,;:!?""'()[]{}/-" & Chr(8220) & Chr(8221) & _
                                   Chr(8216) & Chr(8217) & Chr(8212) & Chr(8211)
    IsPunctuation = (InStr(1, PUNCT_CHARS, ch) > 0)
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if a word is in the known-valid list
' ------------------------------------------------------------
Private Function IsKnownValidRepetition(ByVal word As String, _
                                         ByRef knownValid As Variant) As Boolean
    Dim i As Long
    Dim lWord As String
    lWord = LCase(word)

    For i = LBound(knownValid) To UBound(knownValid)
        If LCase(CStr(knownValid(i))) = lWord Then
            IsKnownValidRepetition = True
            Exit Function
        End If
    Next i

    IsKnownValidRepetition = False
End Function

' ============================================================
'  HELPERS FOR Check_SpellOutUnderTen
' ============================================================

' ------------------------------------------------------------
'  PRIVATE: Check if paragraph style should be excluded
'  Excludes: Table, Code, Data, Technical, Footnote
' ------------------------------------------------------------
Private Function IsExcludedStyle(ByVal styleName As String) As Boolean
    Dim lStyle As String
    lStyle = LCase(styleName)

    IsExcludedStyle = (InStr(lStyle, "table") > 0) Or _
                      (InStr(lStyle, "code") > 0) Or _
                      (InStr(lStyle, "data") > 0) Or _
                      (InStr(lStyle, "technical") > 0) Or _
                      (InStr(lStyle, "footnote") > 0)
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if the digit is part of a larger number
'  (preceded or followed by another digit or decimal point)
' ------------------------------------------------------------
Private Function IsPartOfLargerNumber(ByRef txt As String, _
                                       ByVal pos As Long, _
                                       ByVal textLen As Long) As Boolean
    Dim prevChar As String
    Dim nextChar As String

    IsPartOfLargerNumber = False

    ' Check character before
    If pos > 1 Then
        prevChar = Mid(txt, pos - 1, 1)
        If (prevChar >= "0" And prevChar <= "9") Or _
           prevChar = "." Or prevChar = "," Then
            IsPartOfLargerNumber = True
            Exit Function
        End If
    End If

    ' Check character after
    If pos < textLen Then
        nextChar = Mid(txt, pos + 1, 1)
        If (nextChar >= "0" And nextChar <= "9") Or _
           nextChar = "." Or nextChar = "," Then
            IsPartOfLargerNumber = True
            Exit Function
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is preceded by a structural
'  reference word (section, para, clause, etc.)
' ------------------------------------------------------------
Private Function IsPrecededByStructuralRef(ByRef txt As String, _
                                            ByVal pos As Long) As Boolean
    Dim refWords As Variant
    refWords = Array("section", "sect", "para", "paragraph", "clause", _
                     "article", "art", "rule", "reg", "regulation", _
                     "chapter", "page", "part", "schedule", "sch", _
                     "annex", "appendix", "item", "figure", "fig", _
                     "table", "tab", "footnote", "endnote", "version", _
                     "vol", "no", "ch", "cl", "fn", "pt", "pp", "p", "r", "s")

    IsPrecededByStructuralRef = False

    ' Extract the word immediately before the digit
    Dim prevWord As String
    prevWord = GetPrecedingWord(txt, pos)
    If Len(prevWord) = 0 Then Exit Function

    Dim lWord As String
    lWord = LCase(prevWord)

    Dim j As Long
    For j = LBound(refWords) To UBound(refWords)
        If lWord = LCase(CStr(refWords(j))) Then
            IsPrecededByStructuralRef = True
            Exit Function
        End If
    Next j
End Function

' ------------------------------------------------------------
'  PRIVATE: Get the word immediately preceding position pos
'  Looks back from pos, skipping whitespace, then collecting
'  letters until a non-letter is found.
' ------------------------------------------------------------
Private Function GetPrecedingWord(ByRef txt As String, _
                                   ByVal pos As Long) As String
    Dim k As Long
    Dim ch As String
    Dim wordEnd As Long
    Dim wordStart As Long

    GetPrecedingWord = ""

    ' Skip whitespace before the digit
    k = pos - 1
    Do While k >= 1
        ch = Mid(txt, k, 1)
        If ch <> " " And ch <> vbTab Then Exit Do
        k = k - 1
    Loop

    If k < 1 Then Exit Function

    ' Check we landed on a letter or period (for abbreviations like "s.")
    ' Skip trailing period/dot
    If ch = "." Then
        k = k - 1
        If k < 1 Then Exit Function
    End If

    ' Now collect the word (letters only) going backwards
    wordEnd = k
    Do While k >= 1
        ch = Mid(txt, k, 1)
        If IsLetterChar(ch) Then
            k = k - 1
        Else
            Exit Do
        End If
    Loop
    wordStart = k + 1

    If wordStart > wordEnd Then Exit Function

    GetPrecedingWord = Mid(txt, wordStart, wordEnd - wordStart + 1)
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is part of a range pattern
'  e.g. "7-12", "3--9", digit followed by en-dash/hyphen
'  and another digit, or preceded by digit+dash
' ------------------------------------------------------------
Private Function IsPartOfRange(ByRef txt As String, _
                                ByVal pos As Long, _
                                ByVal textLen As Long) As Boolean
    Dim nextPos As Long
    Dim nextChar As String
    Dim prevPos As Long
    Dim prevChar As String

    IsPartOfRange = False

    ' Check forward: digit followed by dash/en-dash then digit
    nextPos = pos + 1
    If nextPos <= textLen Then
        nextChar = Mid(txt, nextPos, 1)
        ' Hyphen, en-dash (ChrW(8211)), or em-dash (ChrW(8212))
        If nextChar = "-" Or AscW(nextChar) = 8211 Or AscW(nextChar) = 8212 Then
            ' Check if next-next is a digit
            If nextPos + 1 <= textLen Then
                Dim afterDash As String
                afterDash = Mid(txt, nextPos + 1, 1)
                If afterDash >= "0" And afterDash <= "9" Then
                    IsPartOfRange = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' Check backward: preceded by dash then digit (we are the end of a range)
    prevPos = pos - 1
    If prevPos >= 1 Then
        prevChar = Mid(txt, prevPos, 1)
        If prevChar = "-" Or AscW(prevChar) = 8211 Or AscW(prevChar) = 8212 Then
            If prevPos - 1 >= 1 Then
                Dim beforeDash As String
                beforeDash = Mid(txt, prevPos - 1, 1)
                If beforeDash >= "0" And beforeDash <= "9" Then
                    IsPartOfRange = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' Check for "to" pattern: digit + space + "to" + space + digit
    ' Forward check -- need at least 5 chars after pos: " to X"
    If pos + 5 <= textLen Then
        If Mid(txt, pos + 1, 4) = " to " Then
            Dim afterTo As String
            afterTo = Mid(txt, pos + 5, 1)
            If afterTo >= "0" And afterTo <= "9" Then
                IsPartOfRange = True
                Exit Function
            End If
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is in a citation context
'  Look for "[" within 10 characters before
' ------------------------------------------------------------
Private Function IsInCitationContext(ByRef txt As String, _
                                      ByVal pos As Long) As Boolean
    Dim startSearch As Long
    Dim k As Long

    IsInCitationContext = False

    startSearch = pos - 10
    If startSearch < 1 Then startSearch = 1

    For k = startSearch To pos - 1
        If Mid(txt, k, 1) = "[" Then
            IsInCitationContext = True
            Exit Function
        End If
    Next k
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is preceded by currency symbols,
'  percentage, or unit markers
' ------------------------------------------------------------
Private Function IsPrecededByCurrencyOrUnit(ByRef txt As String, _
                                             ByVal pos As Long) As Boolean
    Dim prevChar As String
    Dim prevCode As Long

    IsPrecededByCurrencyOrUnit = False

    If pos <= 1 Then Exit Function

    prevChar = Mid(txt, pos - 1, 1)
    prevCode = AscW(prevChar)

    ' Currency symbols: $, pound sign (163), euro (8364), yen (165)
    ' Unit markers: %, #
    Select Case prevCode
        Case 36    ' $
            IsPrecededByCurrencyOrUnit = True
        Case 163   ' pound sign
            IsPrecededByCurrencyOrUnit = True
        Case 8364  ' euro sign
            IsPrecededByCurrencyOrUnit = True
        Case 165   ' yen sign
            IsPrecededByCurrencyOrUnit = True
        Case 37    ' %
            IsPrecededByCurrencyOrUnit = True
        Case 35    ' #
            IsPrecededByCurrencyOrUnit = True
    End Select

    ' Also check if the character after the digit is %
    If Not IsPrecededByCurrencyOrUnit Then
        If pos < Len(txt) Then
            Dim nextChar As String
            nextChar = Mid(txt, pos + 1, 1)
            If nextChar = "%" Then
                IsPrecededByCurrencyOrUnit = True
            End If
        End If
    End If
End Function

' ============================================================
'  SHARED HELPER (used by both rules' helpers)
' ============================================================

' ------------------------------------------------------------
'  PRIVATE: Check if a character is a letter (A-Z, a-z,
'  extended Latin)
' ------------------------------------------------------------
Private Function IsLetterChar(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsLetterChar = (code >= 65 And code <= 90) Or _
                   (code >= 97 And code <= 122) Or _
                   (code >= 192 And code <= 687) ' Extended Latin
End Function

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineIsInPageRange
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineGetLocationString
' ----------------------------------------------------------------

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
