Attribute VB_Name = "Rules_Terms"
' ============================================================
' Rules_Terms.bas
' Combined module for term-related rules:
'   Rule05 - Custom term whitelist
'   Rule07 - Defined terms
'   Rule23 - Phrase consistency
' ============================================================
Option Explicit

Private Const RULE05_NAME As String = "custom_term_whitelist"
Private Const RULE07_NAME As String = "defined_terms"
Private Const RULE23_NAME As String = "phrase_consistency"

' ============================================================
'  PRIVATE HELPERS (Rule07)
' ============================================================

' -- Helper: remove hyphens from a term ----------------------
Private Function RemoveHyphens(ByVal term As String) As String
    RemoveHyphens = Replace(term, "-", "")
End Function

' -- Helper: count occurrences of a term in document text ----
Private Function CountTermInDoc(doc As Document, ByVal searchTerm As String) As Long
    Dim rng As Range
    Dim cnt As Long
    cnt = 0

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = searchTerm
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
    End With

    Do While rng.Find.Execute
        cnt = cnt + 1
        rng.Collapse wdCollapseEnd
    Loop

    CountTermInDoc = cnt
End Function

' -- Helper: find first occurrence of a term and return range -
Private Function FindTermRange(doc As Document, ByVal searchTerm As String, _
                                matchCase As Boolean) As Range
    Dim rng As Range
    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = searchTerm
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = matchCase
        .MatchWholeWord = True
        .MatchWildcards = False
    End With

    If rng.Find.Execute Then
        Set FindTermRange = rng
    Else
        Set FindTermRange = Nothing
    End If
End Function

' ============================================================
'  PRIVATE HELPERS (Rule23)
' ============================================================

' -- Check a single phrase group for consistency -------------
'  Counts each phrase, determines dominant, flags minorities.
' ------------------------------------------------------------
Private Sub CheckPhraseGroup(doc As Document, _
                              phrases As Variant, _
                              ByRef issues As Collection)
    Dim counts() As Long
    Dim phraseCount As Long
    Dim p As Long
    Dim dominantIdx As Long
    Dim dominantCount As Long
    Dim usedCount As Long

    phraseCount = UBound(phrases) - LBound(phrases) + 1
    ReDim counts(LBound(phrases) To UBound(phrases))

    ' -- Count occurrences of each phrase ---------------------
    For p = LBound(phrases) To UBound(phrases)
        counts(p) = CountPhrase(doc, CStr(phrases(p)))
    Next p

    ' -- Determine how many phrases in this group are used ----
    usedCount = 0
    dominantIdx = LBound(phrases)
    dominantCount = counts(LBound(phrases))

    For p = LBound(phrases) To UBound(phrases)
        If counts(p) > 0 Then usedCount = usedCount + 1
        If counts(p) > dominantCount Then
            dominantCount = counts(p)
            dominantIdx = p
        End If
    Next p

    ' Only flag if more than one phrase in the group is used
    If usedCount < 2 Then Exit Sub

    ' -- Flag all minority phrase occurrences -----------------
    For p = LBound(phrases) To UBound(phrases)
        If counts(p) > 0 And p <> dominantIdx Then
            FlagPhraseOccurrences doc, CStr(phrases(p)), CStr(phrases(dominantIdx)), issues
        End If
    Next p
End Sub

' -- Count occurrences of a phrase in the document -----------
'  Uses Find with MatchWildcards=False, MatchWholeWord=False
'  (necessary for multi-word phrases).
' ------------------------------------------------------------
Private Function CountPhrase(doc As Document, phrase As String) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean

    cnt = 0

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = phrase
        .MatchWholeWord = False
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

        If EngineIsInPageRange(rng) Then
            cnt = cnt + 1
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    CountPhrase = cnt
End Function

' -- Flag all occurrences of a minority phrase ---------------
Private Sub FlagPhraseOccurrences(doc As Document, _
                                   minorityPhrase As String, _
                                   dominantPhrase As String, _
                                   ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = minorityPhrase
        .MatchWholeWord = False
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

        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE23_NAME, locStr, "Inconsistent phrase: '" & rng.Text & "' used", "Use '" & dominantPhrase & "' for consistency (dominant style)", rng.Start, rng.End, "error")
            issues.Add finding
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ============================================================
'  RULE 05: CUSTOM TERM WHITELIST
' ============================================================
Public Function Check_CustomTermWhitelist(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' -- Define default whitelist terms ----------------------
    Dim terms As Variant
    Dim batch1 As Variant, batch2 As Variant
    batch1 = Array( _
        "co-counsel", "anti-suit injunction", "pre-action", _
        "re-examination", "cross-examination", "counter-claim", _
        "sub-contract", "non-disclosure", "inter-partes", _
        "ex-parte", "bona fide")
    batch2 = Array( _
        "prima facie", "pro rata", "ad hoc", "de facto", _
        "de jure", "inter alia", "mutatis mutandis", _
        "pari passu", "ultra vires", "vis-a-vis")
    terms = MergeArrays2(batch1, batch2)

    ' -- Build the dictionary -------------------------------
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim t As Variant
    For Each t In terms
        Dim lcTerm As String
        lcTerm = LCase(CStr(t))
        If Not dict.Exists(lcTerm) Then
            dict.Add lcTerm, True
        End If
    Next t

    ' -- Store in the engine for other rules to query -------
    EngineSetWhitelist dict

    On Error GoTo 0

    ' This rule returns no issues -- it is purely a setup rule
    Set Check_CustomTermWhitelist = issues
End Function

' ============================================================
'  RULE 07: DEFINED TERMS
' ============================================================
Public Function Check_DefinedTerms(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' Dictionary: term (String) -> Array(definitionParaIdx, rangeStart, rangeEnd)
    Dim definedTerms As Object
    Set definedTerms = CreateObject("Scripting.Dictionary")
    Dim defInfo() As Variant
    Dim mInfo() As Variant
    Dim hInfo() As Variant
    Dim pInfo() As Variant
    Dim rng As Range
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim paraText As String

    ' ==========================================================
    '  PASS 1: Scan for defined terms
    ' ==========================================================

    ' -- Pattern A: Curly-quoted defined terms ----------------
    ' Look for opening curly quote followed by uppercase letter
    Dim leftCurly As String
    Dim rightCurly As String
    leftCurly = ChrW(8220)   ' left double curly quote
    rightCurly = ChrW(8221)  ' right double curly quote

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = leftCurly & "[A-Z]"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = True
    End With

    Do While rng.Find.Execute
        If Not EngineIsInPageRange(rng) Then
            rng.Collapse wdCollapseEnd
            GoTo NextCurlyFind
        End If

        ' Expand to find the closing curly quote
        Dim startPos As Long
        startPos = rng.Start
        Dim expandedRng As Range
        Set expandedRng = doc.Range(startPos, startPos)

        ' Search forward for closing curly quote (max 100 chars)
        Dim endSearch As Long
        endSearch = startPos + 100
        If endSearch > doc.Content.End Then endSearch = doc.Content.End
        Set expandedRng = doc.Range(startPos, endSearch)
        Dim fullText As String
        fullText = expandedRng.Text

        Dim closePos As Long
        closePos = InStr(2, fullText, rightCurly)
        If closePos > 1 Then
            Dim termText As String
            ' Extract between quotes (skip the opening quote)
            termText = Mid$(fullText, 2, closePos - 2)
            If Len(Trim$(termText)) > 0 And Not definedTerms.Exists(termText) Then
                ReDim defInfo(0 To 2)
                defInfo(0) = 0 ' paragraph index (approximate)
                defInfo(1) = startPos + 1  ' range start of term
                defInfo(2) = startPos + CLng(closePos) - 1  ' range end of term
                definedTerms.Add termText, defInfo
            End If
        End If

        rng.Collapse wdCollapseEnd
NextCurlyFind:
    Loop

    ' -- Pattern B: "X means " or "X has the meaning " -------
    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1
        If Not EngineIsInPageRange(para.Range) Then GoTo NextParaMeans

        paraText = para.Range.Text
        Dim meansPos As Long
        meansPos = InStr(1, paraText, " means ", vbTextCompare)
        If meansPos > 1 Then
            ' Extract term before " means "
            Dim beforeMeans As String
            beforeMeans = Trim$(Left$(paraText, meansPos - 1))
            ' Take last quoted or significant phrase
            Dim lastQuoteStart As Long
            lastQuoteStart = InStrRev(beforeMeans, leftCurly)
            If lastQuoteStart = 0 Then lastQuoteStart = InStrRev(beforeMeans, """")
            If lastQuoteStart > 0 Then
                Dim afterQuote As String
                afterQuote = Mid$(beforeMeans, lastQuoteStart + 1)
                Dim endQuote As Long
                endQuote = InStr(1, afterQuote, rightCurly)
                If endQuote = 0 Then endQuote = InStr(1, afterQuote, """")
                If endQuote > 1 Then
                    Dim meansTerm As String
                    meansTerm = Left$(afterQuote, endQuote - 1)
                    If Len(meansTerm) > 0 And Not definedTerms.Exists(meansTerm) Then
                        ReDim mInfo(0 To 2)
                        mInfo(0) = paraIdx
                        mInfo(1) = para.Range.Start
                        mInfo(2) = para.Range.Start + meansPos
                        definedTerms.Add meansTerm, mInfo
                    End If
                End If
            End If
        End If

        ' Check for "has the meaning"
        Dim htmPos As Long
        htmPos = InStr(1, paraText, " has the meaning ", vbTextCompare)
        If htmPos > 1 Then
            Dim beforeHTM As String
            beforeHTM = Trim$(Left$(paraText, htmPos - 1))
            lastQuoteStart = InStrRev(beforeHTM, leftCurly)
            If lastQuoteStart = 0 Then lastQuoteStart = InStrRev(beforeHTM, """")
            If lastQuoteStart > 0 Then
                afterQuote = Mid$(beforeHTM, lastQuoteStart + 1)
                endQuote = InStr(1, afterQuote, rightCurly)
                If endQuote = 0 Then endQuote = InStr(1, afterQuote, """")
                If endQuote > 1 Then
                    Dim htmTerm As String
                    htmTerm = Left$(afterQuote, endQuote - 1)
                    If Len(htmTerm) > 0 And Not definedTerms.Exists(htmTerm) Then
                        ReDim hInfo(0 To 2)
                        hInfo(0) = paraIdx
                        hInfo(1) = para.Range.Start
                        hInfo(2) = para.Range.Start + htmPos
                        definedTerms.Add htmTerm, hInfo
                    End If
                End If
            End If
        End If
NextParaMeans:
    Next para

    ' -- Pattern C: Parenthetical definitions (the "Term") ---
    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = "(the " & leftCurly & "[A-Z]"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = True
    End With

    Do While rng.Find.Execute
        If Not EngineIsInPageRange(rng) Then
            rng.Collapse wdCollapseEnd
            GoTo NextParenFind
        End If

        startPos = rng.Start
        endSearch = startPos + 120
        If endSearch > doc.Content.End Then endSearch = doc.Content.End
        Set expandedRng = doc.Range(startPos, endSearch)
        fullText = expandedRng.Text

        ' Find closing curly quote then closing paren
        closePos = InStr(6, fullText, rightCurly)
        If closePos > 6 Then
            ' Extract between the curly quotes
            Dim pOpenQ As Long
            pOpenQ = InStr(1, fullText, leftCurly)
            If pOpenQ > 0 Then
                Dim parenTerm As String
                parenTerm = Mid$(fullText, pOpenQ + 1, closePos - pOpenQ - 1)
                If Len(parenTerm) > 0 And Not definedTerms.Exists(parenTerm) Then
                    ReDim pInfo(0 To 2)
                    pInfo(0) = 0
                    pInfo(1) = startPos + pOpenQ
                    pInfo(2) = startPos + CLng(closePos)
                    definedTerms.Add parenTerm, pInfo
                End If
            End If
        End If

        rng.Collapse wdCollapseEnd
NextParenFind:
    Loop

    ' ==========================================================
    '  PASS 2: Validate each defined term
    ' ==========================================================
    Dim termKey As Variant
    For Each termKey In definedTerms.keys
        Dim term As String
        term = CStr(termKey)
        Dim tInfo As Variant
        tInfo = definedTerms(termKey)

        ' -- Check A: Lowercase variant (inconsistent capitalisation) --
        Dim lcTerm2 As String
        lcTerm2 = LCase(Left$(term, 1)) & Mid$(term, 2)
        If lcTerm2 <> term Then
            Dim lcRng As Range
            Set lcRng = FindTermRange(doc, lcTerm2, True)
            If Not lcRng Is Nothing Then
                If EngineIsInPageRange(lcRng) Then
                    Dim findingLC As Object
                    Dim locLC As String
                    locLC = EngineGetLocationString(lcRng, doc)
                    Set findingLC = CreateIssueDict(RULE07_NAME, locLC, "Inconsistent capitalisation: '" & lcTerm2 & "' found but '" & term & "' is the defined term", "Use '" & term & "' consistently", lcRng.Start, lcRng.End, "error")
                    issues.Add findingLC
                End If
            End If
        End If

        ' -- Check B: Hyphenated/unhyphenated variant ----------
        If InStr(1, term, "-") > 0 Then
            Dim noHyphen As String
            noHyphen = RemoveHyphens(term)
            Dim nhRng As Range
            Set nhRng = FindTermRange(doc, noHyphen, False)
            If Not nhRng Is Nothing Then
                If EngineIsInPageRange(nhRng) Then
                    Dim findingH As Object
                    Dim locH As String
                    locH = EngineGetLocationString(nhRng, doc)
                    Set findingH = CreateIssueDict(RULE07_NAME, locH, "Hyphenation variant: '" & noHyphen & "' found but defined term uses hyphen: '" & term & "'", "Use the defined form: '" & term & "'", nhRng.Start, nhRng.End, "error")
                    issues.Add findingH
                End If
            End If
        Else
            ' Term has no hyphen -- check if hyphenated variant exists
            ' Try common hyphenation points (before common prefixes)
            ' This is a best-effort check
        End If

        ' -- Check C: Defined term never referenced ------------
        Dim totalCount As Long
        totalCount = CountTermInDoc(doc, term)
        If totalCount <= 1 Then
            ' Only appears at the definition site
            Dim findingUnused As Object
            Dim unusedRng As Range
            Set unusedRng = doc.Range(CLng(tInfo(1)), CLng(tInfo(2)))
            Dim locUnused As String
            locUnused = EngineGetLocationString(unusedRng, doc)
            Set findingUnused = CreateIssueDict(RULE07_NAME, locUnused, "Defined term never referenced: '" & term & "' is defined but not used elsewhere in the document", "Remove the definition or use the term in the document", CLng(tInfo(1)), CLng(tInfo(2)), "possible_error")
            issues.Add findingUnused
        End If
    Next termKey

    On Error GoTo 0
    Set Check_DefinedTerms = issues
End Function

' ============================================================
'  RULE 23: PHRASE CONSISTENCY
' ============================================================
Public Function Check_PhraseConsistency(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Define phrase groups ---------------------------------
    ' Each group is an array of synonymous phrases
    Dim groups(0 To 9) As Variant

    groups(0) = Array("not later than", "no later than")
    groups(1) = Array("in respect of", "with respect to", "in relation to")
    groups(2) = Array("pursuant to", "in accordance with")
    groups(3) = Array("notwithstanding", "despite", "regardless of")
    groups(4) = Array("prior to", "before")
    groups(5) = Array("subsequent to", "after", "following")
    groups(6) = Array("in the event that", "if", "where")
    groups(7) = Array("save that", "except that", "provided that")
    groups(8) = Array("forthwith", "immediately", "without delay")
    groups(9) = Array("hereby", "by this")

    ' -- Process each group -----------------------------------
    Dim g As Long
    For g = 0 To 9
        CheckPhraseGroup doc, groups(g), issues
    Next g

    Set Check_PhraseConsistency = issues
End Function


' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineSetWhitelist ' ----------------------------------------------------------------

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
'  Late-bound wrapper: EngineSetWhitelist ' ----------------------------------------------------------------

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
'  Late-bound wrapper: PleadingsEngine.SetWhitelist
' ----------------------------------------------------------------
Private Sub EngineSetWhitelist(dict As Object)
    On Error Resume Next
    Application.Run "PleadingsEngine.SetWhitelist", dict
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

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
