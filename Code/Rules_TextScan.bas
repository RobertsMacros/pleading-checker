Attribute VB_Name = "Rules_TextScan"
' ============================================================
' Rules_TextScan.bas
' Combined text-scanning proofreading rules:
'   - Check_RepeatedWords (from Rule02)
'   - Check_SpellOutUnderTen (from Rule34)
'
' Dependencies:
'   - TextAnchoring.bas (IterateParagraphs, AddIssue, SafeRange,
'                        IsWhitespaceChar, IsLetterChar,
'                        StripPunctuation, IsPunctuation, PerfCount)
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
    Set Check_RepeatedWords = TextAnchoring.IterateParagraphs(doc, "Rules_TextScan", "ProcessParagraph_RepeatedWords")
End Function

' ============================================================
'  PUBLIC: Check_SpellOutUnderTen
'  In running prose, numbers under 10 should be written in
'  words (e.g. "seven" instead of "7").
' ============================================================
Public Function Check_SpellOutUnderTen(doc As Document) As Collection
    Set Check_SpellOutUnderTen = TextAnchoring.IterateParagraphs(doc, "Rules_TextScan", "ProcessParagraph_SpellOutUnderTen")
End Function

' ============================================================
'  HELPERS FOR Check_RepeatedWords
' ============================================================


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

    ' Strip trailing "s" to handle plurals (e.g. "Rules" -> "rule")
    Dim lWordBase As String
    lWordBase = lWord
    If Len(lWordBase) > 2 And Right$(lWordBase, 1) = "s" Then
        lWordBase = Left$(lWordBase, Len(lWordBase) - 1)
    End If

    Dim j As Long
    For j = LBound(refWords) To UBound(refWords)
        If lWord = LCase(CStr(refWords(j))) Or _
           lWordBase = LCase(CStr(refWords(j))) Then
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

    ' Check we landed on a letter or full stop (for abbreviations like "s.")
    ' Skip trailing full stop/dot
    If ch = "." Then
        k = k - 1
        If k < 1 Then Exit Function
    End If

    ' Now collect the word (letters only) going backwards
    wordEnd = k
    Do While k >= 1
        ch = Mid(txt, k, 1)
        If TextAnchoring.IsLetterChar(ch) Then
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
'  PRIVATE: Check if digit is inside parentheses -- catches
'  clause sub-numbers like "34(3)(e)", "(iv)", "s.2(1)" etc.
' ------------------------------------------------------------
Private Function IsInsideParentheses(ByRef txt As String, _
                                      ByVal pos As Long) As Boolean
    IsInsideParentheses = False

    ' Check for opening paren before (skipping digits and letters)
    Dim k As Long
    k = pos - 1
    If k >= 1 Then
        If Mid(txt, k, 1) = "(" Then
            IsInsideParentheses = True
            Exit Function
        End If
    End If

    ' Check for closing paren after (skipping ahead past the digit)
    k = pos + 1
    If k <= Len(txt) Then
        If Mid(txt, k, 1) = ")" Then
            IsInsideParentheses = True
            Exit Function
        End If
    End If
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

' ------------------------------------------------------------
'  PRIVATE: Check if digit is linked via conjunction (and/or/to)
'  to another digit that IS preceded by a structural reference.
'  Catches patterns like "paragraphs 4 and 5", "rules 3 to 7",
'  "sections 2 or 3", "paragraphs 4, 5 and 6".
' ------------------------------------------------------------
Private Function IsConjunctionLinkedRef(ByRef txt As String, _
                                         ByVal pos As Long) As Boolean
    IsConjunctionLinkedRef = False

    ' Get the word before this digit
    Dim prevWord As String
    prevWord = LCase(GetPrecedingWord(txt, pos))
    If Len(prevWord) = 0 Then Exit Function

    ' Must be preceded by "and", "or", "to", or a comma
    Dim isConj As Boolean
    isConj = (prevWord = "and" Or prevWord = "or" Or prevWord = "to")

    ' Also handle comma-separated: "paragraphs 4, 5 and 6"
    If Not isConj Then
        ' Check if preceded by comma (skip spaces)
        Dim k As Long
        k = pos - 1
        Do While k >= 1
            Dim c As String
            c = Mid$(txt, k, 1)
            If c = " " Or c = vbTab Then
                k = k - 1
            Else
                Exit Do
            End If
        Loop
        If k >= 1 And Mid$(txt, k, 1) = "," Then
            isConj = True
        End If
    End If

    If Not isConj Then Exit Function

    ' Now scan backwards past the conjunction to find a preceding digit
    ' For "and"/"or"/"to": skip back past the conjunction word + spaces + the digit
    ' For comma: already at the comma, skip back past it + spaces + the digit
    Dim scanPos As Long
    scanPos = pos

    ' Skip back to before the preceding word / comma
    scanPos = scanPos - 1  ' space before digit
    Do While scanPos >= 1 And (Mid$(txt, scanPos, 1) = " " Or Mid$(txt, scanPos, 1) = vbTab)
        scanPos = scanPos - 1
    Loop
    ' Skip back past the conjunction word or comma
    If isConj And (prevWord = "and" Or prevWord = "or" Or prevWord = "to") Then
        scanPos = scanPos - Len(prevWord)
    ElseIf isConj Then
        ' comma case — scanPos is already past the comma
    End If
    ' Skip spaces before the conjunction
    Do While scanPos >= 1 And (Mid$(txt, scanPos, 1) = " " Or Mid$(txt, scanPos, 1) = vbTab)
        scanPos = scanPos - 1
    Loop

    ' Check if there's a digit at scanPos
    If scanPos >= 1 Then
        Dim prevCh As String
        prevCh = Mid$(txt, scanPos, 1)
        If prevCh >= "0" And prevCh <= "9" Then
            ' Found a digit — check if THAT digit is preceded by a structural ref
            If IsPrecededByStructuralRef(txt, scanPos) Then
                IsConjunctionLinkedRef = True
                Exit Function
            End If
            ' Or if THAT digit is also conjunction-linked (recursive chain)
            If IsConjunctionLinkedRef(txt, scanPos) Then
                IsConjunctionLinkedRef = True
            End If
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is adjacent to a letter
'  (postcodes like SO50 2ZH, codes like ET1, etc.)
' ------------------------------------------------------------
Private Function IsAdjacentToLetter(ByRef txt As String, _
                                     ByVal pos As Long, _
                                     ByVal textLen As Long) As Boolean
    IsAdjacentToLetter = False

    ' Check character before
    If pos > 1 Then
        If TextAnchoring.IsLetterChar(Mid(txt, pos - 1, 1)) Then
            IsAdjacentToLetter = True
            Exit Function
        End If
    End If

    ' Check character after
    If pos < textLen Then
        If TextAnchoring.IsLetterChar(Mid(txt, pos + 1, 1)) Then
            IsAdjacentToLetter = True
            Exit Function
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is followed by opening bracket
'  (clause references like 1(4), 3(a), etc.)
' ------------------------------------------------------------
Private Function IsFollowedByBracket(ByRef txt As String, _
                                      ByVal pos As Long, _
                                      ByVal textLen As Long) As Boolean
    IsFollowedByBracket = False

    If pos < textLen Then
        If Mid(txt, pos + 1, 1) = "(" Then
            IsFollowedByBracket = True
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is followed by a month name
'  (date patterns like "1 October 2004")
' ------------------------------------------------------------
Private Function IsFollowedByMonthName(ByRef txt As String, _
                                        ByVal pos As Long, _
                                        ByVal textLen As Long) As Boolean
    IsFollowedByMonthName = False

    ' Need at least a space + 3 chars after the digit
    If pos + 4 > textLen Then Exit Function

    ' Must be followed by a space
    If Mid(txt, pos + 1, 1) <> " " Then Exit Function

    ' Extract the next word after the space
    Dim wordStart As Long
    wordStart = pos + 2
    Dim wordEnd As Long
    wordEnd = wordStart
    Do While wordEnd <= textLen
        If Not TextAnchoring.IsLetterChar(Mid(txt, wordEnd, 1)) Then Exit Do
        wordEnd = wordEnd + 1
    Loop

    If wordEnd <= wordStart Then Exit Function

    Dim nextWord As String
    nextWord = LCase(Mid(txt, wordStart, wordEnd - wordStart))

    Dim months As Variant
    months = Array("january", "february", "march", "april", "may", _
                   "june", "july", "august", "september", "october", _
                   "november", "december")

    Dim m As Long
    For m = LBound(months) To UBound(months)
        If nextWord = CStr(months(m)) Then
            IsFollowedByMonthName = True
            Exit Function
        End If
    Next m
End Function

' Check if the digit is effectively at the start of paragraph text
' (possibly after whitespace/tab), which typically means paragraph numbering
Private Function IsAtParagraphStart(ByRef txt As String, _
                                     ByVal pos As Long, _
                                     ByVal listPrefixLen As Long) As Boolean
    IsAtParagraphStart = False
    Dim effectivePos As Long
    effectivePos = pos - listPrefixLen
    If effectivePos > 5 Then Exit Function  ' not near start
    ' Check that everything before this digit is whitespace/tab
    Dim k As Long
    For k = 1 + listPrefixLen To pos - 1
        Dim c As String
        c = Mid$(txt, k, 1)
        If c <> " " And c <> vbTab And c <> ChrW(160) Then
            Exit Function  ' non-whitespace before digit = not paragraph start
        End If
    Next k
    ' Also check if digit is followed by "." or ")" which is numbering
    If pos < Len(txt) Then
        Dim nextCh As String
        nextCh = Mid$(txt, pos + 1, 1)
        If nextCh = "." Or nextCh = ")" Or nextCh = " " Then
            IsAtParagraphStart = True
        End If
    End If
End Function

' ============================================================
'  PUBLIC: ProcessParagraph_RepeatedWords
'  Per-paragraph handler extracted from Check_RepeatedWords.
'  Scans a single paragraph for consecutive repeated words.
' ============================================================
Public Sub ProcessParagraph_RepeatedWords(doc As Document, paraRange As Range, paraText As String, paraStart As Long, listPrefixLen As Long, ByRef issues As Collection)
    Dim knownValid As Variant
    knownValid = Array("that", "had", "is", "was", "can")

    Dim tLen As Long
    tLen = Len(paraText)
    If tLen < 3 Then Exit Sub

    Dim prevWord As String
    prevWord = ""
    Dim prevTokenStart As Long, prevTokenEnd As Long
    prevTokenStart = 0: prevTokenEnd = 0

    Dim scanPos As Long
    scanPos = 1  ' 1-based position in paraText

    Dim currWord As String
    Dim rawToken As String
    Dim severity As String
    Dim issueText As String
    Dim suggestion As String
    Dim rangeStart As Long
    Dim rangeEnd As Long

    On Error Resume Next
    Do While scanPos <= tLen
        ' Skip whitespace
        Dim sc As String
        sc = Mid$(paraText, scanPos, 1)
        If TextAnchoring.IsWhitespaceChar(sc) Then
            scanPos = scanPos + 1
            GoTo NextScanPos_PPR
        End If

        ' Found start of a token
        Dim tokStart As Long
        tokStart = scanPos
        Do While scanPos <= tLen
            sc = Mid$(paraText, scanPos, 1)
            If TextAnchoring.IsWhitespaceChar(sc) Then Exit Do
            scanPos = scanPos + 1
        Loop
        Dim tokEnd As Long
        tokEnd = scanPos  ' one past end (exclusive)

        rawToken = Mid$(paraText, tokStart, tokEnd - tokStart)
        currWord = LCase(TextAnchoring.StripPunctuation(rawToken))

        If Len(currWord) = 0 Then
            prevWord = ""
            GoTo NextScanPos_PPR
        End If

        ' Check for repetition with previous token
        If currWord = prevWord And Len(currWord) > 0 Then
            ' tokStart is 1-based in paraText; convert to document position
            rangeStart = paraStart + (tokStart - 1) - listPrefixLen
            rangeEnd = rangeStart + (tokEnd - tokStart)

            ' -- Whole-word verification --
            Err.Clear
            Dim matchRange As Range
            Set matchRange = doc.Range(rangeStart, rangeEnd)
            If Err.Number <> 0 Then Err.Clear: GoTo NextScanPos_PPR

            Dim actualCurr As String
            actualCurr = LCase(TextAnchoring.StripPunctuation(matchRange.Text))
            If Err.Number <> 0 Then Err.Clear: GoTo NextScanPos_PPR
            If actualCurr <> currWord Then
                ' Try alternative offsets
                Dim rwOffset As Long
                Dim rwFixed As Boolean
                rwFixed = False
                Dim rwTryOffsets As Variant
                rwTryOffsets = Array(-1, 1, -2, 2)
                Dim rwOff As Long
                For rwOff = LBound(rwTryOffsets) To UBound(rwTryOffsets)
                    rwOffset = CLng(rwTryOffsets(rwOff))
                    Dim rwTryStart As Long
                    rwTryStart = rangeStart + rwOffset
                    If rwTryStart >= 0 Then
                        Set matchRange = doc.Range(rwTryStart, rwTryStart + (tokEnd - tokStart))
                        If Err.Number = 0 Then
                            actualCurr = LCase(TextAnchoring.StripPunctuation(matchRange.Text))
                            If Err.Number = 0 And actualCurr = currWord Then
                                rangeStart = rwTryStart
                                rangeEnd = rwTryStart + (tokEnd - tokStart)
                                rwFixed = True
                                TextAnchoring.PerfCount "anchoring_corrections"
                                Exit For
                            End If
                            If Err.Number <> 0 Then Err.Clear
                        Else
                            Err.Clear
                        End If
                    End If
                Next rwOff
                If Not rwFixed Then GoTo NextScanPos_PPR
            End If

            Dim prevRngStart As Long, prevRngEnd As Long
            prevRngStart = paraStart + (prevTokenStart - 1) - listPrefixLen
            prevRngEnd = prevRngStart + (prevTokenEnd - prevTokenStart)

            Dim prevMatchRange As Range
            Set prevMatchRange = doc.Range(prevRngStart, prevRngEnd)
            If Err.Number <> 0 Then Err.Clear: GoTo NextScanPos_PPR

            Dim actualPrev As String
            actualPrev = LCase(TextAnchoring.StripPunctuation(prevMatchRange.Text))
            If Err.Number <> 0 Then Err.Clear: GoTo NextScanPos_PPR
            If actualPrev <> currWord Then GoTo NextScanPos_PPR

            ' Also verify the gap between the two words has no
            ' hidden content (only whitespace/punctuation).
            If prevRngEnd < rangeStart Then
                Dim gapRange As Range
                Set gapRange = doc.Range(prevRngEnd, rangeStart)
                If Err.Number = 0 Then
                    Dim gapText As String
                    gapText = gapRange.Text
                    If Err.Number = 0 Then
                        Dim gIdx As Long
                        Dim gCh As String
                        For gIdx = 1 To Len(gapText)
                            gCh = Mid$(gapText, gIdx, 1)
                            If Not TextAnchoring.IsWhitespaceChar(gCh) And _
                               Not TextAnchoring.IsPunctuation(gCh) Then
                                ' Non-whitespace content between the two
                                ' words - this is not a real repetition.
                                GoTo NextScanPos_PPR
                            End If
                        Next gIdx
                    Else
                        Err.Clear
                    End If
                Else
                    Err.Clear
                End If
            End If
            ' -- End whole-word verification --

            ' Determine severity
            If IsKnownValidRepetition(currWord, knownValid) Then
                severity = "possible_error"
                issueText = "Repeated word '" & currWord & "'"
            Else
                severity = "error"
                issueText = "Repeated word '" & currWord & "'"
            End If

            suggestion = "Remove the duplicate '" & currWord & "'"

            TextAnchoring.AddIssue issues, RULE_NAME_REPEATED, doc, matchRange, issueText, suggestion, rangeStart, rangeEnd, severity, False, "", rawToken, "token"
        End If

        prevWord = currWord
        prevTokenStart = tokStart
        prevTokenEnd = tokEnd
NextScanPos_PPR:
    Loop
    On Error GoTo 0
End Sub

' ============================================================
'  PUBLIC: ProcessParagraph_SpellOutUnderTen
'  Per-paragraph handler extracted from Check_SpellOutUnderTen.
'  Scans a single paragraph for digits 0-9 that should be
'  spelled out in running prose.
' ============================================================
Public Sub ProcessParagraph_SpellOutUnderTen(doc As Document, paraRange As Range, paraText As String, paraStart As Long, listPrefixLen As Long, ByRef issues As Collection)
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

    Dim styleName As String
    Dim i As Long
    Dim ch As String
    Dim digitVal As Long
    Dim charRange As Range
    Dim textLen As Long

    On Error Resume Next

    ' -- Check paragraph style for exclusions --
    styleName = ""
    styleName = paraRange.ParagraphStyle
    If Err.Number <> 0 Then
        Err.Clear
        styleName = ""
    End If

    If IsExcludedStyle(styleName) Then
        Exit Sub
    End If

    ' -- Skip block quotes / indented extracts --
    Dim isBlockQ As Boolean
    isBlockQ = False
    isBlockQ = Application.Run("Rules_Formatting.IsBlockQuotePara", paraRange.Paragraphs(1))
    If Err.Number <> 0 Then isBlockQ = False: Err.Clear
    If isBlockQ Then Exit Sub

    ' -- Skip headings (numbering is structural, not prose) --
    Dim soOutlineLevel As Long
    soOutlineLevel = 10
    soOutlineLevel = paraRange.Paragraphs(1).OutlineLevel
    If Err.Number <> 0 Then soOutlineLevel = 10: Err.Clear
    If soOutlineLevel >= 1 And soOutlineLevel <= 9 Then Exit Sub

    textLen = Len(paraText)
    If textLen = 0 Then Exit Sub

    ' -- Scan character by character for digits 0-9 --
    For i = 1 To textLen
        ch = Mid(paraText, i, 1)

        ' Check if character is a digit 0-9
        If ch >= "0" And ch <= "9" Then
            ' -- Check: digit at start of paragraph (likely numbering) --
            If IsAtParagraphStart(paraText, i, listPrefixLen) Then
                GoTo NextChar_PPS
            End If

            digitVal = CInt(ch)

            ' -- Check: isolated digit (not part of larger number) --
            If IsPartOfLargerNumber(paraText, i, textLen) Then
                GoTo NextChar_PPS
            End If

            ' -- Check: digit adjacent to a letter (postcodes, codes) --
            If IsAdjacentToLetter(paraText, i, textLen) Then
                GoTo NextChar_PPS
            End If

            ' -- Check: preceded by structural reference word --
            If IsPrecededByStructuralRef(paraText, i) Then
                GoTo NextChar_PPS
            End If

            ' -- Check: inside parentheses (clause sub-numbers) --
            If IsInsideParentheses(paraText, i) Then
                GoTo NextChar_PPS
            End If

            ' -- Check: digit followed by opening bracket (clause ref like 1(4)) --
            If IsFollowedByBracket(paraText, i, textLen) Then
                GoTo NextChar_PPS
            End If

            ' -- Check: digit followed by month name (date like 1 October) --
            If IsFollowedByMonthName(paraText, i, textLen) Then
                GoTo NextChar_PPS
            End If

            ' -- Check: part of a range pattern --
            If IsPartOfRange(paraText, i, textLen) Then
                GoTo NextChar_PPS
            End If

            ' -- Check: citation context --
            If IsInCitationContext(paraText, i) Then
                GoTo NextChar_PPS
            End If

            ' -- Check: preceded by currency/unit symbols --
            If IsPrecededByCurrencyOrUnit(paraText, i) Then
                GoTo NextChar_PPS
            End If

            ' -- Check: conjunction-linked structural ref --
            If IsConjunctionLinkedRef(paraText, i) Then
                GoTo NextChar_PPS
            End If

            ' -- All checks passed: flag this digit --
            Dim rangeStart As Long
            Dim rangeEnd As Long

            rangeStart = paraStart + i - 1 - listPrefixLen
            rangeEnd = rangeStart + 1

            ' -- Stale-anchor validation: verify document text matches --
            Err.Clear
            Set charRange = TextAnchoring.SafeRange(doc, rangeStart, rangeEnd)
            If charRange Is Nothing Then
                GoTo NextChar_PPS
            End If
            Dim anchorText As String
            anchorText = charRange.Text
            If Err.Number <> 0 Then Err.Clear: GoTo NextChar_PPS

            If anchorText <> ch Then
                ' Try alternative offsets +/-1, +/-2
                Dim soOffset As Long
                Dim soFixed As Boolean
                soFixed = False
                Dim soTryOffsets As Variant
                soTryOffsets = Array(-1, 1, -2, 2)
                Dim soOff As Long
                For soOff = LBound(soTryOffsets) To UBound(soTryOffsets)
                    soOffset = CLng(soTryOffsets(soOff))
                    Dim soTryStart As Long
                    soTryStart = rangeStart + soOffset
                    If soTryStart >= 0 Then
                        Set charRange = TextAnchoring.SafeRange(doc, soTryStart, soTryStart + 1)
                        If Not charRange Is Nothing Then
                            If charRange.Text = ch Then
                                rangeStart = soTryStart
                                rangeEnd = soTryStart + 1
                                soFixed = True
                                TextAnchoring.PerfCount "anchoring_corrections"
                                Exit For
                            End If
                        End If
                    End If
                Next soOff
                If Not soFixed Then
                    Debug.Print "ANCHOR_WARN: SpellOutUnderTen anchor mismatch at pos " & rangeStart & _
                                ": expected '" & ch & "', got '" & anchorText & "'"
                    GoTo NextChar_PPS
                End If
            End If

            TextAnchoring.AddIssue issues, RULE_NAME_SPELL_OUT, doc, charRange, _
                "Number under 10 is given as a figure in running prose.", _
                "Write '" & numberWords(digitVal) & "' instead of '" & ch & "'.", _
                rangeStart, rangeEnd, "warning", False, "", _
                ch, "token", "medium"
        End If

NextChar_PPS:
    Next i
    On Error GoTo 0
End Sub







