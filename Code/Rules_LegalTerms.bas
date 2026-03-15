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
'   - TextAnchoring.bas (IsInPageRange, GetLocationString,
'     IsPastPageFilter, CreateIssueDict)
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

    Dim lastPos As Long
    lastPos = -1
    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        ' Skip if the matched text already has the correct hyphenated form
        If StrComp(rng.Text, correctForm, vbTextCompare) = 0 Then
            GoTo SkipMatch
        End If

        ' Verify it is not actually the hyphenated form by checking
        ' the surrounding context -- the Find matched with MatchCase=False
        ' and spaces, so an exact binary comparison rules out false positives
        If TextAnchoring.IsInPageRange(rng) Then
            On Error Resume Next
            locStr = TextAnchoring.GetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = TextAnchoring.CreateIssueDict(RULE28_NAME, locStr, "Mandatory term is not hyphenated in the approved form.", "Use '" & correctForm & "'.", rng.Start, rng.End, "warning", False)
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
    Dim batch1 As Variant, batch2 As Variant
    batch1 = Array( _
        "Act", "Bill", "Attorney-General", "Cabinet", _
        "Commonwealth", "Constitution", "Crown", _
        "Executive Council", "Governor", "Governor-General", _
        "Her Majesty", "the Queen")
    batch2 = Array( _
        "his Honour", "her Honour", "their Honours", _
        "Law Lords", "their Lordships", "Lords Justices", _
        "Member States", "Parliament", "Labour Party", _
        "Prime Minister", "Vice-Chancellor")
    terms = MergeArrays2(batch1, batch2)

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
        If TextAnchoring.IsPastPageFilter(paraRng.Start) Then Exit For
        If Not TextAnchoring.IsInPageRange(paraRng) Then GoTo NextPara

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
        locStr = TextAnchoring.GetLocationString(matchRng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
        On Error GoTo 0

        Set finding = TextAnchoring.CreateIssueDict(RULE29_NAME, locStr, "Term should be capitalised in the approved form.", "Use '" & correctForm & "'.", matchStart, matchEnd, "warning", False)
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
                ' Skip if flanked by letters (apostrophe in smart-quote mode)
                If i > 1 And i < Len(paraText) Then
                    Dim ls16Prev As String, ls16Next As String
                    ls16Prev = Mid(paraText, i - 1, 1)
                    ls16Next = Mid(paraText, i + 1, 1)
                    Dim ls16PrevA As Boolean, ls16NextA As Boolean
                    ls16PrevA = IsWordChar(ls16Prev) And ls16Prev <> "-" And ls16Prev <> "_"
                    ls16NextA = IsWordChar(ls16Next) And ls16Next <> "-" And ls16Next <> "_"
                    If Not (ls16PrevA And ls16NextA) Then
                        openCount = openCount + 1
                    End If
                Else
                    openCount = openCount + 1
                End If
            Case 8217  ' right single smart quote
                ' Skip if flanked by letters (apostrophe: it's, don't)
                If i > 1 And i < Len(paraText) Then
                    Dim rs17Prev As String, rs17Next As String
                    rs17Prev = Mid(paraText, i - 1, 1)
                    rs17Next = Mid(paraText, i + 1, 1)
                    Dim rs17PrevA As Boolean, rs17NextA As Boolean
                    rs17PrevA = IsWordChar(rs17Prev) And rs17Prev <> "-" And rs17Prev <> "_"
                    rs17NextA = IsWordChar(rs17Next) And rs17Next <> "-" And rs17Next <> "_"
                    If rs17PrevA And rs17NextA Then
                        ' Apostrophe - skip
                    ElseIf openCount > 0 Then
                        openCount = openCount - 1
                    End If
                ElseIf openCount > 0 Then
                    openCount = openCount - 1
                End If
            Case 34    ' straight double quote -- toggle
                If openCount > 0 Then
                    openCount = openCount - 1
                Else
                    openCount = openCount + 1
                End If
            Case 39    ' straight single quote / apostrophe
                ' Distinguish apostrophe from quote delimiter:
                ' If flanked by letters/digits on both sides, it's an apostrophe -> skip.
                ' Otherwise use whitespace heuristic for open/close.
                Dim prevCh As String
                Dim nextCh As String
                Dim prevIsAlpha As Boolean, nextIsAlpha As Boolean
                prevIsAlpha = False
                nextIsAlpha = False
                If i > 1 Then
                    prevCh = Mid(paraText, i - 1, 1)
                    prevIsAlpha = IsWordChar(prevCh) And prevCh <> "-" And prevCh <> "_"
                End If
                If i < Len(paraText) Then
                    nextCh = Mid(paraText, i + 1, 1)
                    nextIsAlpha = IsWordChar(nextCh) And nextCh <> "-" And nextCh <> "_"
                End If
                ' Letter/digit on both sides = apostrophe (it's, don't, 90's)
                If prevIsAlpha And nextIsAlpha Then
                    ' Skip: this is an apostrophe, not a quote
                Else
                    If i = 1 Then
                        openCount = openCount + 1
                    ElseIf Not prevIsAlpha Then
                        ' Preceded by space/punct = opening quote
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
'  Merge 2 Variant arrays into one flat Variant array
' ----------------------------------------------------------------
Private Function MergeArrays2(a1 As Variant, a2 As Variant) As Variant
    Dim total As Long
    total = UBound(a1) - LBound(a1) + 1 _
          + UBound(a2) - LBound(a2) + 1
    Dim out() As Variant
    ReDim out(0 To total - 1)
    Dim idx As Long
    idx = 0
    Dim v As Variant
    For Each v In a1: out(idx) = v: idx = idx + 1: Next v
    For Each v In a2: out(idx) = v: idx = idx + 1: Next v
    MergeArrays2 = out
End Function
