Attribute VB_Name = "Rule34_SpellOutUnderTen"
' ============================================================
' Rule34_SpellOutUnderTen.bas
' Proofreading rule: in running prose, numbers under 10 should
' be written in words (e.g. "seven" instead of "7").
'
' Exceptions:
'   - Preceded by structural reference words (section, para, etc.)
'   - Part of a range pattern (e.g. "7-12", "3--9")
'   - In a table, code, data, technical, or footnote style
'   - In a citation context (near "[" or footnote style)
'   - Part of a larger number (adjacent digits or decimal)
'   - Preceded by currency symbols, %, or unit markers
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "spell_out_under_ten"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_SpellOutUnderTen(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim styleName As String
    Dim i As Long
    Dim ch As String
    Dim digitVal As Long
    Dim issue As PleadingsIssue
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
            GoTo NextParagraph
        End If

        ' Skip paragraphs outside the configured page range
        If Not PleadingsEngine.IsInPageRange(paraRange) Then
            GoTo NextParagraph
        End If

        ' ── Check paragraph style for exclusions ────────────
        styleName = ""
        styleName = paraRange.ParagraphStyle
        If Err.Number <> 0 Then
            Err.Clear
            styleName = ""
        End If

        If IsExcludedStyle(styleName) Then
            GoTo NextParagraph
        End If

        ' ── Get paragraph text ──────────────────────────────
        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph
        End If

        textLen = Len(paraText)
        If textLen = 0 Then GoTo NextParagraph

        ' ── Scan character by character for digits 0-9 ──────
        For i = 1 To textLen
            ch = Mid(paraText, i, 1)

            ' Check if character is a digit 0-9
            If ch >= "0" And ch <= "9" Then
                digitVal = CInt(ch)

                ' ── Check: isolated digit (not part of larger number) ──
                If IsPartOfLargerNumber(paraText, i, textLen) Then
                    GoTo NextChar
                End If

                ' ── Check: preceded by structural reference word ──
                If IsPrecededByStructuralRef(paraText, i) Then
                    GoTo NextChar
                End If

                ' ── Check: part of a range pattern ──
                If IsPartOfRange(paraText, i, textLen) Then
                    GoTo NextChar
                End If

                ' ── Check: citation context ──
                If IsInCitationContext(paraText, i) Then
                    GoTo NextChar
                End If

                ' ── Check: preceded by currency/unit symbols ──
                If IsPrecededByCurrencyOrUnit(paraText, i) Then
                    GoTo NextChar
                End If

                ' ── All checks passed: flag this digit ──────
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
                    locStr = PleadingsEngine.GetLocationString(charRange, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If
                End If

                Set issue = New PleadingsIssue
                issue.Init RULE_NAME, _
                           locStr, _
                           "Number under 10 is given as a figure in running prose.", _
                           "Write '" & numberWords(digitVal) & "' instead of '" & ch & "'.", _
                           rangeStart, _
                           rangeEnd, _
                           "warning", _
                           False
                issues.Add issue
            End If

NextChar:
        Next i

NextParagraph:
    Next para
    On Error GoTo 0

    Set Check_SpellOutUnderTen = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if paragraph style should be excluded
'  Excludes: Table, Code, Data, Technical, Footnote
' ════════════════════════════════════════════════════════════
Private Function IsExcludedStyle(ByVal styleName As String) As Boolean
    Dim lStyle As String
    lStyle = LCase(styleName)

    IsExcludedStyle = (InStr(lStyle, "table") > 0) Or _
                      (InStr(lStyle, "code") > 0) Or _
                      (InStr(lStyle, "data") > 0) Or _
                      (InStr(lStyle, "technical") > 0) Or _
                      (InStr(lStyle, "footnote") > 0)
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if the digit is part of a larger number
'  (preceded or followed by another digit or decimal point)
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if digit is preceded by a structural
'  reference word (section, para, clause, etc.)
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Get the word immediately preceding position pos
'  Looks back from pos, skipping whitespace, then collecting
'  letters until a non-letter is found.
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if digit is part of a range pattern
'  e.g. "7-12", "3--9", digit followed by en-dash/hyphen
'  and another digit, or preceded by digit+dash
' ════════════════════════════════════════════════════════════
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
    ' Forward check
    If pos + 1 <= textLen Then
        If Mid(txt, pos + 1, 1) = " " Then
            If pos + 4 <= textLen Then
                If LCase(Mid(txt, pos + 2, 2)) = "to" Then
                    If pos + 4 <= textLen Then
                        If Mid(txt, pos + 4, 1) = " " Then
                            If pos + 5 <= textLen Then
                                Dim afterTo As String
                                afterTo = Mid(txt, pos + 5, 1)
                                If afterTo >= "0" And afterTo <= "9" Then
                                    IsPartOfRange = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if digit is in a citation context
'  Look for "[" within 10 characters before
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if digit is preceded by currency symbols,
'  percentage, or unit markers
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a character is a letter (A-Z, a-z,
'  extended Latin)
' ════════════════════════════════════════════════════════════
Private Function IsLetterChar(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsLetterChar = (code >= 65 And code <= 90) Or _
                   (code >= 97 And code <= 122) Or _
                   (code >= 192 And code <= 687) ' Extended Latin
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunSpellOutUnderTen()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Spell Out Under Ten"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_SpellOutUnderTen(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Spell Out Under Ten"
End Sub
