Attribute VB_Name = "Rule02_repeated_words"
' ============================================================
' Rule02_repeated-words.bas
' Proofreading rule: detects consecutive repeated words
' (e.g. "the the", "is is") within paragraphs.
'
' Known-valid repetitions (e.g. "that that", "had had") are
' flagged with severity "possible_error" rather than "error"
' to prompt manual review of context.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "repeated_words"

' ============================================================
'  MAIN ENTRY POINT
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
    Dim issue As Object
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
            GoTo NextParagraph
        End If

        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextParagraph
        End If

        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph
        End If

        ' Skip very short or empty paragraphs
        If Len(Trim(paraText)) < 3 Then
            GoTo NextParagraph
        End If

        ' -- Split paragraph into words --------------------
        words = Split(paraText, " ")
        wordCount = UBound(words) - LBound(words) + 1

        If wordCount < 2 Then
            GoTo NextParagraph
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

                ' -- Create the issue ------------------
                Set issue = CreateIssueDict(RULE_NAME, locStr, issueText, suggestion, rangeStart, rangeEnd, severity)
                issues.Add issue
            End If

            prevWord = currWord

NextWordInPara:
        Next i

NextParagraph:
    Next para
    On Error GoTo 0

    Set Check_RepeatedWords = issues
End Function

' ============================================================
'  PRIVATE: Strip leading and trailing punctuation from a word
'  Removes characters like . , ; : ! ? " ' ( ) [ ] etc.
' ============================================================
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

' ============================================================
'  PRIVATE: Check if a character is punctuation
' ============================================================
Private Function IsPunctuation(ByVal ch As String) As Boolean
    Const PUNCT_CHARS As String = ".,;:!?""'()[]{}/-" & Chr(8220) & Chr(8221) & _
                                   Chr(8216) & Chr(8217) & Chr(8212) & Chr(8211)
    IsPunctuation = (InStr(1, PUNCT_CHARS, ch) > 0)
End Function

' ============================================================
'  PRIVATE: Check if a word is in the known-valid list
' ============================================================
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
