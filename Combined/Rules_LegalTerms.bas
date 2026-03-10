Attribute VB_Name = "Rules_LegalTerms"
' ============================================================
' Rules_LegalTerms.bas
' Combined module for Rule28 (mandated legal term forms) and
' Rule29 (always capitalise terms).
'
' Rule28: enforces fixed hyphenation for specific legal and
'   governmental terms. Flags unhyphenated variants and suggests
'   the approved hyphenated form.
'   Default mandatory list:
'     "Solicitor-General", "Attorney-General"
'   Additional terms can be added at runtime via AddMandatedTerm.
'
' Rule29: enforces capitalisation for specified Hart-style terms.
'   Scans each paragraph for case-insensitive matches and flags
'   any occurrence whose capitalisation does not match the
'   approved form. Matches inside quoted material are skipped.
'   Context-sensitive terms (Province, State, party names) are
'   intentionally omitted -- the engine does not yet have
'   reliable context handling.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE28_NAME As String = "mandated_legal_term_forms"
Private Const RULE29_NAME As String = "always_capitalise_terms"

' -- Module-level dictionary for Rule28 ----------------------
' Key = LCase(correct form), Value = correct form (String)
Private mandatedTerms As Object

' ============================================================
'  RULE 28 -- MAIN ENTRY POINT
' ============================================================
Public Function Check_MandatedLegalTermForms(doc As Document) As Collection
    Dim issues As New Collection

    ' Initialise defaults if not yet loaded
    If mandatedTerms Is Nothing Then
        InitDefaultTerms
    End If

    Dim keys As Variant
    Dim k As Long
    Dim correctForm As String
    Dim searchPhrase As String

    keys = mandatedTerms.keys

    For k = 0 To mandatedTerms.Count - 1
        correctForm = CStr(mandatedTerms(keys(k)))

        ' Build the unhyphenated search variant by replacing hyphens with spaces
        searchPhrase = Replace(correctForm, "-", " ")

        ' Only search if the unhyphenated form is actually different
        If StrComp(searchPhrase, correctForm, vbBinaryCompare) <> 0 Then
            SearchAndFlag doc, searchPhrase, correctForm, issues
        End If
    Next k

    Set Check_MandatedLegalTermForms = issues
End Function

' ============================================================
'  RULE 28 -- PRIVATE: Search for an unhyphenated variant and flag matches
' ============================================================
Private Sub SearchAndFlag(doc As Document, _
                           searchPhrase As String, _
                           correctForm As String, _
                           ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = searchPhrase
        .MatchWholeWord = True
        .MatchCase = False
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        ' Skip if the matched text already has the correct hyphenated form
        If StrComp(rng.Text, correctForm, vbTextCompare) = 0 Then
            GoTo SkipMatch
        End If

        ' Verify it is not actually the hyphenated form by checking
        ' the surrounding context -- the Find matched with MatchCase=False
        ' and spaces, so an exact binary comparison rules out false positives
        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE28_NAME, locStr, "Mandatory term is not hyphenated in the approved form.", "Use)
            issues.Add finding
        End If

SkipMatch:
        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ============================================================
'  RULE 28 -- PRIVATE: Populate default mandatory terms
' ============================================================
Private Sub InitDefaultTerms()
    Set mandatedTerms = CreateObject("Scripting.Dictionary")

    mandatedTerms.Add LCase("Solicitor-General"), "Solicitor-General"
    mandatedTerms.Add LCase("Attorney-General"), "Attorney-General"
End Sub

' ============================================================
'  RULE 28 -- PUBLIC: Add a mandated term at runtime
'  The term must contain a hyphen (e.g. "Director-General").
'  If the term already exists it is silently ignored.
' ============================================================
Public Sub AddMandatedTerm(term As String)
    If mandatedTerms Is Nothing Then
        InitDefaultTerms
    End If

    Dim lcKey As String
    lcKey = LCase(term)

    If Not mandatedTerms.Exists(lcKey) Then
        mandatedTerms.Add lcKey, term
    End If
End Sub

' ============================================================
'  RULE 29 -- MAIN ENTRY POINT
' ============================================================
Public Function Check_AlwaysCapitaliseTerms(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Seed dictionary of correct forms -------------------
    Dim terms As Variant
    terms = Array( _
        "Act", _
        "Bill", _
        "Attorney-General", _
        "Cabinet", _
        "Commonwealth", _
        "Constitution", _
        "Crown", _
        "Executive Council", _
        "Governor", _
        "Governor-General", _
        "Her Majesty", _
        "the Queen", _
        "his Honour", _
        "her Honour", _
        "their Honours", _
        "Law Lords", _
        "their Lordships", _
        "Lords Justices", _
        "Member States", _
        "Parliament", _
        "Labour Party", _
        "Prime Minister", _
        "Vice-Chancellor" _
    )

    ' -- Iterate paragraphs ---------------------------------
    Dim para As Paragraph
    Dim paraRng As Range
    Dim paraText As String
    Dim paraStart As Long

    For Each para In doc.Paragraphs
        On Error Resume Next
        Set paraRng = para.Range
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextPara
        On Error GoTo 0

        ' Check page range filter
        If Not EngineIsInPageRange(paraRng) Then GoTo NextPara

        paraText = paraRng.Text
        paraStart = paraRng.Start

        If Len(paraText) = 0 Then GoTo NextPara

        ' -- Check each term against this paragraph ---------
        Dim t As Long
        For t = LBound(terms) To UBound(terms)
            CheckTermInParagraph doc, CStr(terms(t)), paraText, paraStart, paraRng, issues
        Next t

NextPara:
    Next para

    Set Check_AlwaysCapitaliseTerms = issues
End Function

' ============================================================
'  RULE 29 -- PRIVATE: Search for a single term within one paragraph
' ============================================================
Private Sub CheckTermInParagraph(doc As Document, _
                                  correctForm As String, _
                                  paraText As String, _
                                  paraStart As Long, _
                                  paraRng As Range, _
                                  ByRef issues As Collection)
    Dim termLen As Long
    Dim pos As Long
    Dim actualText As String
    Dim matchStart As Long
    Dim matchEnd As Long
    Dim finding As Object
    Dim locStr As String
    Dim charBefore As String
    Dim charAfter As String

    termLen = Len(correctForm)

    ' Walk through all case-insensitive matches in the paragraph
    pos = InStr(1, paraText, correctForm, vbTextCompare)

    Do While pos > 0
        ' -- Word boundary check ----------------------------
        ' Ensure we are not matching a substring of a longer word
        If pos > 1 Then
            charBefore = Mid(paraText, pos - 1, 1)
            If IsWordChar(charBefore) Then GoTo NextMatch
        End If

        If pos + termLen <= Len(paraText) Then
            charAfter = Mid(paraText, pos + termLen, 1)
            If IsWordChar(charAfter) Then GoTo NextMatch
        End If

        ' -- Extract the actual text at the match position --
        actualText = Mid(paraText, pos, termLen)

        ' -- Skip if capitalisation already matches ---------
        If StrComp(actualText, correctForm, vbBinaryCompare) = 0 Then
            GoTo NextMatch
        End If

        ' -- Skip if inside quoted material -----------------
        If IsInsideQuote(paraText, pos) Then GoTo NextMatch

        ' -- Calculate range positions ----------------------
        matchStart = paraStart + pos - 1
        matchEnd = matchStart + termLen

        On Error Resume Next
        Dim matchRng As Range
        Set matchRng = doc.Range(matchStart, matchEnd)
        locStr = EngineGetLocationString(matchRng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
        On Error GoTo 0

        Set finding = CreateIssueDict(RULE29_NAME, locStr, "Term should be capitalised in the approved form.", "Use)
        issues.Add finding

NextMatch:
        ' Search for next occurrence after current position
        If pos + 1 > Len(paraText) Then Exit Do
        pos = InStr(pos + 1, paraText, correctForm, vbTextCompare)
    Loop
End Sub

' ============================================================
'  PRIVATE: Check whether a character is a word character
'  (letter, digit, hyphen, or underscore)
' ============================================================
Private Function IsWordChar(ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then
        IsWordChar = False
        Exit Function
    End If

    c = AscW(ch)

    ' A-Z, a-z
    If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
        IsWordChar = True
        Exit Function
    End If

    ' 0-9
    If c >= 48 And c <= 57 Then
        IsWordChar = True
        Exit Function
    End If

    ' Hyphen or underscore (treat as word chars for compound terms)
    If c = 45 Or c = 95 Then
        IsWordChar = True
        Exit Function
    End If

    IsWordChar = False
End Function

' ============================================================
'  PRIVATE: Determine if position is inside quoted material
'  Checks for smart quotes and straight quotes in a window
'  before the match position.
' ============================================================
Private Function IsInsideQuote(paraText As String, matchPos As Long) As Boolean
    Dim openCount As Long
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim windowStart As Long

    IsInsideQuote = False
    openCount = 0

    ' Scan from start of paragraph to match position
    ' to count unmatched opening quotes
    windowStart = 1
    If matchPos <= 1 Then Exit Function

    For i = windowStart To matchPos - 1
        ch = Mid(paraText, i, 1)
        code = AscW(ch)

        Select Case code
            Case 8220  ' left double smart quote
                openCount = openCount + 1
            Case 8221  ' right double smart quote
                If openCount > 0 Then openCount = openCount - 1
            Case 8216  ' left single smart quote
                openCount = openCount + 1
            Case 8217  ' right single smart quote
                If openCount > 0 Then openCount = openCount - 1
            Case 34    ' straight double quote -- toggle
                If openCount > 0 Then
                    openCount = openCount - 1
                Else
                    openCount = openCount + 1
                End If
            Case 39    ' straight single quote / apostrophe
                ' Only treat as quote if preceded by whitespace or at start
                If i = 1 Then
                    openCount = openCount + 1
                Else
                    Dim prevCh As String
                    prevCh = Mid(paraText, i - 1, 1)
                    If prevCh = " " Or prevCh = vbTab Or AscW(prevCh) = 10 Or AscW(prevCh) = 13 Then
                        openCount = openCount + 1
                    ElseIf openCount > 0 Then
                        openCount = openCount - 1
                    End If
                End If
        End Select
    Next i

    ' If there are unmatched opening quotes, the match is inside quoted material
    If openCount > 0 Then
        IsInsideQuote = True
    End If
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
