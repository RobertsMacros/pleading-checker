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
'   - TextAnchoring.bas (IterateParagraphs, SafeRange, AddIssue,
'     IsLetterChar, MergeArrays3)
' ============================================================
Option Explicit

' -- Rule-name constants -------------------------------------
Private Const RULE_NAME_ANGLICISED As String = "non_english_terms"
Private Const RULE_NAME_FOREIGN   As String = "non_english_terms"

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
    Dim italVal As Long

    On Error Resume Next
    italVal = rng.Font.Italic
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: IsRangeItalic = False: Exit Function
    On Error GoTo 0

    ' If Font.Italic is True the whole range is italic
    If italVal = True Then
        IsRangeItalic = True
        Exit Function
    End If

    ' wdToggle treated as italic present
    If italVal = wdToggle Then
        IsRangeItalic = True
        Exit Function
    End If

    ' If Font.Italic is wdUndefined (9999999) the range has
    ' mixed formatting -- check individual characters
    If italVal = wdUndefined Then
        Dim i As Long
        Dim charRng As Range
        For i = rng.Start To rng.End - 1
            On Error Resume Next
            Set charRng = rng.Document.Range(i, i + 1)
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextCharItalic
            Dim charItal As Long
            charItal = charRng.Font.Italic
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextCharItalic
            On Error GoTo 0
            If charItal = True Then
                IsRangeItalic = True
                Exit Function
            End If
NextCharItalic:
        Next i
    End If

    IsRangeItalic = False
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
    seedTerms = TextAnchoring.MergeArrays3(batch1, batch2, batch3)
    seedInitialised = True
End Sub

' ------------------------------------------------------------
'  MAIN ENTRY POINT -- Rule 30
' ------------------------------------------------------------
Public Function Check_AnglicisedTermsNotItalic(doc As Document) As Collection
    InitSeedTerms
    Set Check_AnglicisedTermsNotItalic = TextAnchoring.IterateParagraphs(doc, "Rules_Italics", "ProcessParagraph_AnglicisedTerms")
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
    ' Initialise defaults if not yet loaded
    If foreignNames Is Nothing Then
        InitSeedNames
    End If
    Set Check_ForeignNamesNotItalic = TextAnchoring.IterateParagraphs(doc, "Rules_Italics", "ProcessParagraph_ForeignNames")
End Function

' ============================================================
'  ProcessParagraph_AnglicisedTerms
'  Extracts per-paragraph logic from Check_AnglicisedTermsNotItalic.
' ============================================================
Public Sub ProcessParagraph_AnglicisedTerms(doc As Document, paraRange As Range, paraText As String, paraStart As Long, listPrefixLen As Long, ByRef issues As Collection)
    InitSeedTerms

    Dim termIdx As Long
    Dim term As String
    Dim termLen As Long
    Dim pos As Long
    Dim charBefore As String
    Dim charAfter As String
    Dim rng As Range
    Dim adjStart As Long

    For termIdx = LBound(seedTerms) To UBound(seedTerms)
        term = CStr(seedTerms(termIdx))
        termLen = Len(term)

        pos = InStr(1, paraText, term, vbTextCompare)
        Do While pos > 0
            If pos > 1 Then
                charBefore = Mid$(paraText, pos - 1, 1)
                If TextAnchoring.IsLetterChar(charBefore) Then GoTo NextMatchAT
            End If

            If pos + termLen <= Len(paraText) Then
                charAfter = Mid$(paraText, pos + termLen, 1)
                If TextAnchoring.IsLetterChar(charAfter) Then GoTo NextMatchAT
            End If

            ' Anchor model: paraText includes the list prefix, so we
            ' subtract listPrefixLen to map back to document positions.
            ' Formula: docPos = paraStart + (pos - 1) - listPrefixLen
            adjStart = paraStart + (pos - 1) - listPrefixLen

            Set rng = TextAnchoring.SafeRange(doc, adjStart, adjStart + termLen)
            If rng Is Nothing Then GoTo NextMatchAT

            If IsRangeItalic(rng) Then
                TextAnchoring.AddIssue issues, RULE_NAME_ANGLICISED, doc, rng, _
                    "Anglicised foreign term is italicised.", _
                    "Set '" & term & "' in roman, not italics.", _
                    rng.Start, rng.End, "warning"
            End If

NextMatchAT:
            pos = InStr(pos + 1, paraText, term, vbTextCompare)
        Loop
    Next termIdx
End Sub

' ============================================================
'  ProcessParagraph_ForeignNames
'  Extracts per-paragraph logic from Check_ForeignNamesNotItalic.
' ============================================================
Public Sub ProcessParagraph_ForeignNames(doc As Document, paraRange As Range, paraText As String, paraStart As Long, listPrefixLen As Long, ByRef issues As Collection)
    If foreignNames Is Nothing Then
        InitSeedNames
    End If

    Dim keys As Variant
    Dim k As Long
    Dim term As String
    Dim termLen As Long
    Dim pos As Long
    Dim charBefore As String
    Dim charAfter As String
    Dim rng As Range
    Dim adjStart As Long

    keys = foreignNames.keys

    For k = 0 To foreignNames.Count - 1
        term = CStr(keys(k))
        termLen = Len(term)

        pos = InStr(1, paraText, term, vbTextCompare)
        Do While pos > 0
            If pos > 1 Then
                charBefore = Mid$(paraText, pos - 1, 1)
                If TextAnchoring.IsLetterChar(charBefore) Then GoTo NextMatchFN
            End If

            If pos + termLen <= Len(paraText) Then
                charAfter = Mid$(paraText, pos + termLen, 1)
                If TextAnchoring.IsLetterChar(charAfter) Then GoTo NextMatchFN
            End If

            ' Anchor model: paraText includes the list prefix, so we
            ' subtract listPrefixLen to map back to document positions.
            adjStart = paraStart + (pos - 1) - listPrefixLen

            Set rng = TextAnchoring.SafeRange(doc, adjStart, adjStart + termLen)
            If rng Is Nothing Then GoTo NextMatchFN

            If IsRangeItalic(rng) Then
                TextAnchoring.AddIssue issues, RULE_NAME_FOREIGN, doc, rng, _
                    "Foreign name or institution should not be italicised.", _
                    "Set '" & term & "' in roman, not italics.", _
                    rng.Start, rng.End, "warning"
            End If

NextMatchFN:
            pos = InStr(pos + 1, paraText, term, vbTextCompare)
        Loop
    Next k
End Sub

