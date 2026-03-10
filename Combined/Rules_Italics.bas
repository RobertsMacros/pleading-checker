Attribute VB_Name = "Rules_Italics"
' ============================================================
' Rules_Italics.bas
' Combined italics-related proofreading rules:
'   - Rule 30: flags italicisation of known anglicised foreign
'     terms that should be set in roman (upright) type.
'   - Rule 31: flags italicisation of foreign names, institutions,
'     places or courts that should not be italicised.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' -- Rule-name constants -------------------------------------
Private Const RULE_NAME_ANGLICISED As String = "known_anglicised_terms_not_italic"
Private Const RULE_NAME_FOREIGN   As String = "foreign_names_not_italic"

' -- Seed list of anglicised terms (Rule 30) -----------------
Private seedTerms As Variant
Private seedInitialised As Boolean

' -- Module-level dictionary of protected foreign names (Rule 31) -
' Key = name (String), Value = True (Boolean) -- used as a set
Private foreignNames As Object

' ============================================================
'  SHARED PRIVATE HELPERS
' ============================================================

' ------------------------------------------------------------
'  Check whether a character is a letter (A-Z, a-z)
' ------------------------------------------------------------
Private Function IsLetter(ByVal ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then
        IsLetter = False
        Exit Function
    End If
    c = AscW(Left$(ch, 1))
    IsLetter = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

' ------------------------------------------------------------
'  Check whether a range span is italic
'  Returns True if any part of the range is italic.
' ------------------------------------------------------------
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
'  RULE 30 -- ANGLICISED TERMS NOT ITALIC
' ============================================================

' ------------------------------------------------------------
'  Initialise the seed term list
' ------------------------------------------------------------
Private Sub InitSeedTerms()
    If seedInitialised Then Exit Sub
    Dim batch1 As Variant, batch2 As Variant, batch3 As Variant
    batch1 = Array( _
        "amicus curiae", "a priori", "a fortiori", "bona fide", _
        "de facto", "de jure", "ex parte", "ex post", _
        "ex post facto", "indicia")
    batch2 = Array( _
        "in situ", "inter alia", "laissez-faire", "mutatis mutandis", _
        "novus actus interveniens", "obiter dicta", "per se", _
        "prima facie", "quantum meruit", "quid pro quo")
    batch3 = Array( _
        "raison d'etre", "ratio decidendi", "stare decisis", _
        "terra nullius", "ultra vires", "vice versa", _
        "vis-a-vis", "viz")
    seedTerms = MergeArrays(batch1, batch2, batch3)
    seedInitialised = True
End Sub

' ------------------------------------------------------------
'  MAIN ENTRY POINT -- Rule 30
' ------------------------------------------------------------
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
    Dim finding As Object

    For Each para In doc.Paragraphs
        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(para.Range) Then GoTo NextParaR30

        paraText = para.Range.Text
        If Len(paraText) = 0 Then GoTo NextParaR30

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
                    If IsLetter(charBefore) Then GoTo NextMatchR30
                Else
                    charBefore = ""
                End If

                ' Character after match must be non-letter (or match ends at string end)
                If pos + termLen <= Len(paraText) Then
                    charAfter = Mid$(paraText, pos + termLen, 1)
                    If IsLetter(charAfter) Then GoTo NextMatchR30
                End If

                ' -- Build Range for the matched span ----------
                Set rng = doc.Range( _
                    para.Range.Start + pos - 1, _
                    para.Range.Start + pos - 1 + termLen)

                ' -- Check italic formatting -------------------
                If IsRangeItalic(rng) Then
                    locStr = EngineGetLocationString(rng, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

                    Set finding = CreateIssueDict(RULE_NAME_ANGLICISED, locStr, "Anglicised foreign term is italicised.", "Set)
                    issues.Add finding
                End If

NextMatchR30:
                ' Search for next occurrence after this one
                pos = InStr(pos + 1, paraText, term, vbTextCompare)
            Loop
        Next termIdx

NextParaR30:
    Next para

    On Error GoTo 0
    Set Check_AnglicisedTermsNotItalic = issues
End Function

' ============================================================
'  RULE 31 -- FOREIGN NAMES NOT ITALIC
' ============================================================

' ------------------------------------------------------------
'  Initialise the seed name dictionary
' ------------------------------------------------------------
Private Sub InitSeedNames()
    Set foreignNames = CreateObject("Scripting.Dictionary")
    foreignNames.CompareMode = vbTextCompare

    foreignNames.Add "Cour de cassation", True
    foreignNames.Add "Conseil d'Etat", True
    foreignNames.Add "Bayerisches Staatsministerium der Justiz", True
End Sub

' ------------------------------------------------------------
'  PUBLIC: Add a foreign name to the protected list
' ------------------------------------------------------------
Public Sub AddForeignName(ByVal termName As String)
    If foreignNames Is Nothing Then
        InitSeedNames
    End If

    If Not foreignNames.Exists(termName) Then
        foreignNames.Add termName, True
    End If
End Sub

' ------------------------------------------------------------
'  MAIN ENTRY POINT -- Rule 31
' ------------------------------------------------------------
Public Function Check_ForeignNamesNotItalic(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' Initialise defaults if not yet loaded
    If foreignNames Is Nothing Then
        InitSeedNames
    End If

    Dim para As Paragraph
    Dim paraText As String
    Dim pos As Long
    Dim nameKey As Variant
    Dim term As String
    Dim termLen As Long
    Dim charBefore As String
    Dim charAfter As String
    Dim rng As Range
    Dim locStr As String
    Dim finding As Object
    Dim keys As Variant

    keys = foreignNames.keys

    For Each para In doc.Paragraphs
        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(para.Range) Then GoTo NextParaR31

        paraText = para.Range.Text
        If Len(paraText) = 0 Then GoTo NextParaR31

        Dim k As Long
        For k = 0 To foreignNames.Count - 1
            term = CStr(keys(k))
            termLen = Len(term)

            ' Search for all occurrences of the name in this paragraph
            pos = InStr(1, paraText, term, vbTextCompare)
            Do While pos > 0
                ' -- Word-boundary check -----------------------
                ' Character before match must be non-letter (or match starts at position 1)
                If pos > 1 Then
                    charBefore = Mid$(paraText, pos - 1, 1)
                    If IsLetter(charBefore) Then GoTo NextMatchR31
                Else
                    charBefore = ""
                End If

                ' Character after match must be non-letter (or match ends at string end)
                If pos + termLen <= Len(paraText) Then
                    charAfter = Mid$(paraText, pos + termLen, 1)
                    If IsLetter(charAfter) Then GoTo NextMatchR31
                End If

                ' -- Build Range for the matched span ----------
                Set rng = doc.Range( _
                    para.Range.Start + pos - 1, _
                    para.Range.Start + pos - 1 + termLen)

                ' -- Check italic formatting -------------------
                If IsRangeItalic(rng) Then
                    locStr = EngineGetLocationString(rng, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

                    Set finding = CreateIssueDict(RULE_NAME_FOREIGN, locStr, "Foreign name or institution should not be italicised.", "Set)
                    issues.Add finding
                End If

NextMatchR31:
                ' Search for next occurrence after this one
                pos = InStr(pos + 1, paraText, term, vbTextCompare)
            Loop
        Next k

NextParaR31:
    Next para

    On Error GoTo 0
    Set Check_ForeignNamesNotItalic = issues
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

' ----------------------------------------------------------------
'  Merge up to 3 Variant arrays into one flat Variant array
' ----------------------------------------------------------------
Private Function MergeArrays(a1 As Variant, a2 As Variant, a3 As Variant) As Variant
    Dim total As Long
    total = UBound(a1) - LBound(a1) + 1 _
          + UBound(a2) - LBound(a2) + 1 _
          + UBound(a3) - LBound(a3) + 1
    Dim out() As Variant
    ReDim out(0 To total - 1)
    Dim idx As Long: idx = 0
    Dim v As Variant
    For Each v In a1: out(idx) = v: idx = idx + 1: Next v
    For Each v In a2: out(idx) = v: idx = idx + 1: Next v
    For Each v In a3: out(idx) = v: idx = idx + 1: Next v
    MergeArrays = out
End Function
