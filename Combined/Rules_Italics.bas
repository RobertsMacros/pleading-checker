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
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
'   - Microsoft Scripting Runtime (Scripting.Dictionary)
' ============================================================
Option Explicit

' ── Rule-name constants ─────────────────────────────────────
Private Const RULE_NAME_ANGLICISED As String = "known_anglicised_terms_not_italic"
Private Const RULE_NAME_FOREIGN   As String = "foreign_names_not_italic"

' ── Seed list of anglicised terms (Rule 30) ─────────────────
Private seedTerms As Variant
Private seedInitialised As Boolean

' ── Module-level dictionary of protected foreign names (Rule 31) ─
' Key = name (String), Value = True (Boolean) — used as a set
Private foreignNames As Scripting.Dictionary

' ════════════════════════════════════════════════════════════
'  SHARED PRIVATE HELPERS
' ════════════════════════════════════════════════════════════

' ────────────────────────────────────────────────────────────
'  Check whether a character is a letter (A-Z, a-z)
' ────────────────────────────────────────────────────────────
Private Function IsLetter(ByVal ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then
        IsLetter = False
        Exit Function
    End If
    c = AscW(Left$(ch, 1))
    IsLetter = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

' ────────────────────────────────────────────────────────────
'  Check whether a range span is italic
'  Returns True if any part of the range is italic.
' ────────────────────────────────────────────────────────────
Private Function IsRangeItalic(rng As Range) As Boolean
    On Error Resume Next

    ' If Font.Italic is True the whole range is italic
    If rng.Font.Italic = True Then
        IsRangeItalic = True
        Exit Function
    End If

    ' If Font.Italic is wdUndefined (9999999) the range has
    ' mixed formatting — check individual characters
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

' ════════════════════════════════════════════════════════════
'  RULE 30 — ANGLICISED TERMS NOT ITALIC
' ════════════════════════════════════════════════════════════

' ────────────────────────────────────────────────────────────
'  Initialise the seed term list
' ────────────────────────────────────────────────────────────
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

' ────────────────────────────────────────────────────────────
'  MAIN ENTRY POINT — Rule 30
' ────────────────────────────────────────────────────────────
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
    Dim issue As PleadingsIssue

    For Each para In doc.Paragraphs
        ' Skip paragraphs outside the configured page range
        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextParaR30

        paraText = para.Range.Text
        If Len(paraText) = 0 Then GoTo NextParaR30

        For termIdx = LBound(seedTerms) To UBound(seedTerms)
            term = CStr(seedTerms(termIdx))
            termLen = Len(term)

            ' Search for all occurrences of the term in this paragraph
            pos = InStr(1, paraText, term, vbTextCompare)
            Do While pos > 0
                ' ── Word-boundary check ───────────────────────
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

                ' ── Build Range for the matched span ──────────
                Set rng = doc.Range( _
                    para.Range.Start + pos - 1, _
                    para.Range.Start + pos - 1 + termLen)

                ' ── Check italic formatting ───────────────────
                If IsRangeItalic(rng) Then
                    locStr = PleadingsEngine.GetLocationString(rng, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

                    Set issue = New PleadingsIssue
                    issue.Init RULE_NAME_ANGLICISED, _
                               locStr, _
                               "Anglicised foreign term is italicised.", _
                               "Set '" & term & "' in roman, not italics.", _
                               rng.Start, _
                               rng.End, _
                               "warning", _
                               False
                    issues.Add issue
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

' ════════════════════════════════════════════════════════════
'  RULE 31 — FOREIGN NAMES NOT ITALIC
' ════════════════════════════════════════════════════════════

' ────────────────────────────────────────────────────────────
'  Initialise the seed name dictionary
' ────────────────────────────────────────────────────────────
Private Sub InitSeedNames()
    Set foreignNames = New Scripting.Dictionary
    foreignNames.CompareMode = vbTextCompare

    foreignNames.Add "Cour de cassation", True
    foreignNames.Add "Conseil d'Etat", True
    foreignNames.Add "Bayerisches Staatsministerium der Justiz", True
End Sub

' ────────────────────────────────────────────────────────────
'  PUBLIC: Add a foreign name to the protected list
' ────────────────────────────────────────────────────────────
Public Sub AddForeignName(ByVal termName As String)
    If foreignNames Is Nothing Then
        InitSeedNames
    End If

    If Not foreignNames.Exists(termName) Then
        foreignNames.Add termName, True
    End If
End Sub

' ────────────────────────────────────────────────────────────
'  MAIN ENTRY POINT — Rule 31
' ────────────────────────────────────────────────────────────
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
    Dim issue As PleadingsIssue
    Dim keys As Variant

    keys = foreignNames.keys

    For Each para In doc.Paragraphs
        ' Skip paragraphs outside the configured page range
        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextParaR31

        paraText = para.Range.Text
        If Len(paraText) = 0 Then GoTo NextParaR31

        Dim k As Long
        For k = 0 To foreignNames.Count - 1
            term = CStr(keys(k))
            termLen = Len(term)

            ' Search for all occurrences of the name in this paragraph
            pos = InStr(1, paraText, term, vbTextCompare)
            Do While pos > 0
                ' ── Word-boundary check ───────────────────────
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

                ' ── Build Range for the matched span ──────────
                Set rng = doc.Range( _
                    para.Range.Start + pos - 1, _
                    para.Range.Start + pos - 1 + termLen)

                ' ── Check italic formatting ───────────────────
                If IsRangeItalic(rng) Then
                    locStr = PleadingsEngine.GetLocationString(rng, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

                    Set issue = New PleadingsIssue
                    issue.Init RULE_NAME_FOREIGN, _
                               locStr, _
                               "Foreign name or institution should not be italicised.", _
                               "Set '" & term & "' in roman, not italics.", _
                               rng.Start, _
                               rng.End, _
                               "warning", _
                               False
                    issues.Add issue
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
