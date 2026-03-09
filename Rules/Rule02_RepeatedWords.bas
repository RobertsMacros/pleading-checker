Attribute VB_Name = "Rule02_RepeatedWords"
' ============================================================
' Rule02_RepeatedWords.bas
' Proofreading rule: detects consecutive repeated words
' (e.g. "the the", "is is") within paragraphs.
'
' Known-valid repetitions (e.g. "that that", "had had") are
' flagged with severity "possible_error" rather than "error"
' to prompt manual review of context.
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "repeated_words"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
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
    Dim issue As PleadingsIssue
    Dim paraRange As Range

    ' ── Known-valid repetitions that may be intentional ───
    ' These get flagged as "possible_error" with a note
    ' to review context, rather than a hard "error".
    Dim knownValid As Variant
    knownValid = Array("that", "had", "is", "was", "can")

    ' ── Iterate all paragraphs ────────────────────────────
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph
        End If

        ' Skip paragraphs outside the configured page range
        If Not PleadingsEngine.IsInPageRange(paraRange) Then
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

        ' ── Split paragraph into words ────────────────────
        words = Split(paraText, " ")
        wordCount = UBound(words) - LBound(words) + 1

        If wordCount < 2 Then
            GoTo NextParagraph
        End If

        ' ── Clean words: strip leading/trailing punctuation ─
        ReDim cleanWords(LBound(words) To UBound(words))
        For i = LBound(words) To UBound(words)
            cleanWords(i) = StripPunctuation(words(i))
        Next i

        ' ── Compare consecutive words ─────────────────────
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

                ' ── Determine severity ────────────────
                If IsKnownValidRepetition(currWord, knownValid) Then
                    severity = "possible_error"
                    issueText = "Repeated word '" & cleanWords(i) & "' " & _
                                "-- review context; may be intentional"
                Else
                    severity = "error"
                    issueText = "Repeated word '" & cleanWords(i) & "' detected"
                End If

                suggestion = "Remove the duplicate '" & cleanWords(i) & "'"

                ' ── Calculate character position within paragraph ─
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

                ' ── Build location string ─────────────
                Err.Clear
                Dim matchRange As Range
                Set matchRange = doc.Range(rangeStart, rangeEnd)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                Else
                    locStr = PleadingsEngine.GetLocationString(matchRange, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If
                End If

                ' ── Create the issue ──────────────────
                Set issue = New PleadingsIssue
                issue.Init RULE_NAME, _
                           locStr, _
                           issueText, _
                           suggestion, _
                           rangeStart, _
                           rangeEnd, _
                           severity
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Strip leading and trailing punctuation from a word
'  Removes characters like . , ; : ! ? " ' ( ) [ ] etc.
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a character is punctuation
' ════════════════════════════════════════════════════════════
Private Function IsPunctuation(ByVal ch As String) As Boolean
    Const PUNCT_CHARS As String = ".,;:!?""'()[]{}/-" & Chr(8220) & Chr(8221) & _
                                   Chr(8216) & Chr(8217) & Chr(8212) & Chr(8211)
    IsPunctuation = (InStr(1, PUNCT_CHARS, ch) > 0)
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a word is in the known-valid list
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunRepeatedWords()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Repeated Words"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_RepeatedWords(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Repeated Words"
End Sub
