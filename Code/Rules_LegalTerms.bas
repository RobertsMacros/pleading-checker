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
'   - TextAnchoring.bas (SafeRange, SafeLocationString, AddIssue,
'     FindAll, IterateParagraphs)
' ============================================================
Option Explicit

Private Const RULE28_NAME As String = "mandated_legal_term_forms"
Private Const RULE29_NAME As String = "always_capitalise_terms"

' -- Module-level cached term list for Rule29 ----------------
' Built once, reused across all paragraphs.
Private r29Terms As Variant
Private r29Initialised As Boolean

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
    ' Use FindAll: wholeWord=True, matchCase=False (to catch all case variants)
    Dim matches As Collection
    Set matches = TextAnchoring.FindAll(doc, searchPhrase, True, False)

    Dim i As Long
    Dim matchArr As Variant
    Dim matchText As String
    Dim startPos As Long
    Dim endPos As Long
    Dim rng As Range

    For i = 1 To matches.Count
        matchArr = matches(i)
        startPos = matchArr(0)
        endPos = matchArr(1)
        matchText = matchArr(2)

        ' Skip if the matched text already has the correct hyphenated form
        If StrComp(matchText, correctForm, vbTextCompare) = 0 Then
            GoTo NextMatch
        End If

        Set rng = TextAnchoring.SafeRange(doc, startPos, endPos)
        TextAnchoring.AddIssue issues, RULE28_NAME, doc, rng, _
            "Mandatory term is not hyphenated in the approved form.", _
            "Use '" & correctForm & "'.", _
            startPos, endPos, "warning"

NextMatch:
    Next i
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
    ' -- Seed dictionary of correct forms -------------------
    ' (kept here for documentation; the actual per-paragraph
    '  work re-builds the same array inside ProcessParagraph_AlwaysCapitalise)
    Set Check_AlwaysCapitaliseTerms = TextAnchoring.IterateParagraphs(doc, "Rules_LegalTerms", "ProcessParagraph_AlwaysCapitalise")
End Function

' ============================================================
'  RULE 29 -- PRIVATE: Search for a single term within one paragraph
' ============================================================
Private Sub CheckTermInParagraph(doc As Document, _
                                  correctForm As String, _
                                  paraText As String, _
                                  paraStart As Long, _
                                  paraRng As Range, _
                                  ByRef issues As Collection, _
                                  Optional ByVal listPrefixLen As Long = 0)
    Dim termLen As Long
    Dim pos As Long
    Dim actualText As String
    Dim matchStart As Long
    Dim matchEnd As Long
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
        ' Anchor model: paraText includes the list prefix, so we
        ' subtract listPrefixLen to map back to document positions.
        matchStart = paraStart + (pos - 1) - listPrefixLen
        matchEnd = matchStart + termLen

        Dim rng As Range
        Set rng = TextAnchoring.SafeRange(doc, matchStart, matchEnd)
        TextAnchoring.AddIssue issues, RULE29_NAME, doc, rng, _
            "Term should be capitalised in the approved form.", _
            "Use '" & correctForm & "'.", _
            matchStart, matchEnd, "warning"

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

' ============================================================
'  ProcessParagraph_AlwaysCapitalise
'  Extracts per-paragraph logic from Check_AlwaysCapitaliseTerms.
' ============================================================
Public Sub ProcessParagraph_AlwaysCapitalise(doc As Document, paraRange As Range, paraText As String, paraStart As Long, listPrefixLen As Long, ByRef issues As Collection)
    ' Build the term list once at module scope and reuse it.
    If Not r29Initialised Then
        Dim batch1 As Variant, batch2 As Variant
        batch1 = Array("Act", "Bill", "Attorney-General", "Cabinet", "Commonwealth", "Constitution", "Crown", _
            "Executive Council", "Governor", "Governor-General", "Her Majesty", "the Queen")
        batch2 = Array("his Honour", "her Honour", "their Honours", "Law Lords", "their Lordships", _
            "Lords Justices", "Member States", "Parliament", "Labour Party", "Prime Minister", "Vice-Chancellor")
        r29Terms = TextAnchoring.MergeArrays2(batch1, batch2)
        r29Initialised = True
    End If

    Dim t As Long
    For t = LBound(r29Terms) To UBound(r29Terms)
        CheckTermInParagraph doc, CStr(r29Terms(t)), paraText, paraStart, paraRange, issues, listPrefixLen
    Next t
End Sub
