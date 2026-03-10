Attribute VB_Name = "Rule30_anglicised_terms_not_italic"
' ============================================================
' Rule30_anglicised-terms-not-italic.bas
' Proofreading rule: flags italicisation of known foreign-origin
' terms that Hart's Rules treats as absorbed into English.
' These terms should be set in roman (upright) type, not italics.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "known_anglicised_terms_not_italic"

' -- Seed list of anglicised terms -----------------------------
Private seedTerms As Variant
Private seedInitialised As Boolean

' ============================================================
'  PRIVATE: Initialise the seed term list
' ============================================================
Private Sub InitSeedTerms()
    If seedInitialised Then Exit Sub
    seedTerms = Array( _
        "amicus curiae", _
        "a priori", _
        "a fortiori", _
        "bona fide", _
        "de facto", _
        "de jure", _
        "ex parte", _
        "ex post", _
        "ex post facto", _
        "indicia", _
        "in situ", _
        "inter alia", _
        "laissez-faire", _
        "mutatis mutandis", _
        "novus actus interveniens", _
        "obiter dicta", _
        "per se", _
        "prima facie", _
        "quantum meruit", _
        "quid pro quo", _
        "raison d'etre", _
        "ratio decidendi", _
        "stare decisis", _
        "terra nullius", _
        "ultra vires", _
        "vice versa", _
        "vis-a-vis", _
        "viz")
    seedInitialised = True
End Sub

' ============================================================
'  PRIVATE: Check whether a character is a letter (A-Z, a-z)
' ============================================================
Private Function IsLetter(ByVal ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then
        IsLetter = False
        Exit Function
    End If
    c = AscW(Left$(ch, 1))
    IsLetter = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

' ============================================================
'  PRIVATE: Check whether a range span is italic
'  Returns True if any part of the range is italic.
' ============================================================
Private Function IsRangeItalic(rng As Range) As Boolean
    On Error Resume Next

    ' If Font.Italic is True the whole range is italic
    If rng.Font.Italic = True Then
        IsRangeItalic = True
        Exit Function
    End If

    ' If Font.Italic is wdUndefined (9999999) the range has
    ' mixed formatting -- check individual characters
    If rng.Font.Italic = wdUndefined Then
        Dim i As Long
        Dim charRng As Range
        For i = rng.Start To rng.End - 1
            Set charRng = rng.Document.Range(i, i + 1)
            If charRng.Font.Italic = True Then
                IsRangeItalic = True
                Exit Function
            End If
        Next i
    End If

    ' wdToggle treated as italic present
    If rng.Font.Italic = wdToggle Then
        IsRangeItalic = True
        Exit Function
    End If

    IsRangeItalic = False
    On Error GoTo 0
End Function

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_AnglicisedTermsNotItalic(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    InitSeedTerms

    Dim para As Paragraph
    Dim paraText As String
    Dim pos As Long
    Dim termIdx As Long
    Dim term As String
    Dim termLen As Long
    Dim charBefore As String
    Dim charAfter As String
    Dim rng As Range
    Dim locStr As String
    Dim issue As Object

    For Each para In doc.Paragraphs
        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(para.Range) Then GoTo NextPara

        paraText = para.Range.Text
        If Len(paraText) = 0 Then GoTo NextPara

        For termIdx = LBound(seedTerms) To UBound(seedTerms)
            term = CStr(seedTerms(termIdx))
            termLen = Len(term)

            ' Search for all occurrences of the term in this paragraph
            pos = InStr(1, paraText, term, vbTextCompare)
            Do While pos > 0
                ' -- Word-boundary check -----------------------
                ' Character before match must be non-letter (or match starts at position 1)
                If pos > 1 Then
                    charBefore = Mid$(paraText, pos - 1, 1)
                    If IsLetter(charBefore) Then GoTo NextMatch
                Else
                    charBefore = ""
                End If

                ' Character after match must be non-letter (or match ends at string end)
                If pos + termLen <= Len(paraText) Then
                    charAfter = Mid$(paraText, pos + termLen, 1)
                    If IsLetter(charAfter) Then GoTo NextMatch
                End If

                ' -- Build Range for the matched span ----------
                Set rng = doc.Range( _
                    para.Range.Start + pos - 1, _
                    para.Range.Start + pos - 1 + termLen)

                ' -- Check italic formatting -------------------
                If IsRangeItalic(rng) Then
                    locStr = EngineGetLocationString(rng, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

                    Set issue = CreateIssueDict(RULE_NAME, locStr, "Anglicised foreign term is italicised.", "Set '" & term & "' in roman, not italics.", rng.Start, rng.End, "warning", False)
                    issues.Add issue
                End If

NextMatch:
                ' Search for next occurrence after this one
                pos = InStr(pos + 1, paraText, term, vbTextCompare)
            Loop
        Next termIdx

NextPara:
    Next para

    On Error GoTo 0
    Set Check_AnglicisedTermsNotItalic = issues
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
