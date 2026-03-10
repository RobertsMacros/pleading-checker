Attribute VB_Name = "Rules_Spelling"
' ============================================================
' Rules_Spelling.bas
' Combined proofreading rules for UK/US English spelling.
'
' Rule 1 -- British/US Spelling:
'   Detects ~95 spelling differences between US and UK English,
'   with a configurable direction (UK or US mode).
'   Categories: -or/-our, -ize/-ise, -ization/-isation,
'   -er/-re, -se/-ce, -og/-ogue, -ment variants, misc.
'
'   Text in italics or inside quotation marks is NOT auto-fixed
'   but is flagged as a "possible_error" for manual review.
'
' Rule 12 -- Licence/License:
'   Checks correct UK usage of licence (noun) vs license (verb).
'   Also handles compounds and derivatives.
'   UK convention:
'     licence = noun ("a licence", "the licence holder")
'     license = verb ("to license", "shall license")
'     licensed, licensing = always -s- (verb derivatives)
'
' Rule 13 -- Colour Formatting:
'   Detects non-standard font colours in the document body.
'   Identifies the dominant text colour and flags any runs
'   using a different colour (excluding hyperlinks and
'   heading-styled paragraphs).
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, IsWhitelistedTerm,
'                          GetLocationString, GetSpellingMode)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "spelling"
Private Const RULE_NAME_LICENCE As String = "licence_license"
Private Const RULE_NAME_COLOUR As String = "colour_formatting"

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_Spelling(doc As Document) As Collection
    Dim issues As New Collection
    Dim usWords() As String
    Dim ukWords() As String
    Dim searchWords() As String
    Dim targetWords() As String
    Dim exceptions() As String
    Dim spellingMode As String
    Dim direction As String

    ' -- Build the US <-> UK mapping arrays ----------------
    BuildSpellingArrays usWords, ukWords

    ' -- Determine spelling mode -------------------------
    spellingMode = EngineGetSpellingMode()

    If spellingMode = "US" Then
        ' Search for UK words, suggest US replacements
        searchWords = ukWords
        targetWords = usWords
        direction = "US"

        ' In US mode, no special legal exceptions
        exceptions = Split("program,practice", ",")
    Else
        ' Default: "UK" -- search for US words, suggest UK replacements
        searchWords = usWords
        targetWords = ukWords
        direction = "UK"

        ' "judgment" is standard in UK legal writing (not "judgement")
        ' "practice" is the correct UK noun form (verb: "practise")
        exceptions = Split("program,judgment,practice", ",")
    End If

    ' -- Search main document body -----------------------
    SearchRangeForSpellingIssues doc.Content, doc, searchWords, targetWords, exceptions, direction, issues

    ' -- Search footnotes --------------------------------
    On Error Resume Next
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        Err.Clear
        SearchRangeForSpellingIssues fn.Range, doc, searchWords, targetWords, exceptions, direction, issues
        If Err.Number <> 0 Then Err.Clear
    Next fn
    On Error GoTo 0

    ' -- Search endnotes ---------------------------------
    On Error Resume Next
    Dim en As Endnote
    For Each en In doc.Endnotes
        Err.Clear
        SearchRangeForSpellingIssues en.Range, doc, searchWords, targetWords, exceptions, direction, issues
        If Err.Number <> 0 Then Err.Clear
    Next en
    On Error GoTo 0

    Set Check_Spelling = issues
End Function

' ============================================================
'  PRIVATE: Search a Range for spelling issues
'  Iterates every search/target pair, uses Word's Find to
'  locate whole-word, case-insensitive matches, then filters
'  by page range and whitelist before creating issues.
'
'  direction = "UK" or "US" -- controls the issue text:
'    "UK" -> "US spelling detected: '...'"
'    "US" -> "UK spelling detected: '...'"
' ============================================================
Private Sub SearchRangeForSpellingIssues(searchRange As Range, _
                                         doc As Document, _
                                         ByRef searchWords() As String, _
                                         ByRef targetWords() As String, _
                                         ByRef exceptions() As String, _
                                         ByVal direction As String, _
                                         ByRef issues As Collection)
    Dim i As Long
    Dim rng As Range
    Dim foundText As String
    Dim issue As Object
    Dim locStr As String
    Dim issueText As String
    Dim sourceLabel As String

    ' Determine the label for the detected spelling variant
    If direction = "UK" Then
        sourceLabel = "US"
    Else
        sourceLabel = "UK"
    End If

    For i = LBound(searchWords) To UBound(searchWords)

        ' Reset a fresh range for each search term
        On Error Resume Next
        Set rng = searchRange.Duplicate
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextWord
        End If
        On Error GoTo 0

        With rng.Find
            .ClearFormatting
            .Text = searchWords(i)
            .MatchWholeWord = True
            .MatchCase = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With

        ' Loop through all occurrences of this term
        Do
            On Error Resume Next
            Dim found As Boolean
            found = rng.Find.Execute
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Do
            End If
            On Error GoTo 0

            If Not found Then Exit Do

            foundText = rng.Text

            ' -- Skip exceptions -----------------------
            If IsException(foundText, exceptions) Then
                GoTo ContinueSearch
            End If

            ' -- Skip whitelisted terms ----------------
            If EngineIsWhitelistedTerm(foundText) Then
                GoTo ContinueSearch
            End If

            ' -- Skip if outside configured page range -
            If Not EngineIsInPageRange(rng) Then
                GoTo ContinueSearch
            End If

            ' -- Create the issue ----------------------
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            issueText = sourceLabel & " spelling detected: '" & foundText & "'"

            ' -- Downgrade italic / quoted text -------
            Dim severity As String
            Dim suggestion As String
            severity = "error"
            suggestion = targetWords(i)

            If IsRangeItalic(rng) Then
                severity = "possible_error"
                suggestion = ""
                issueText = issueText & " (in italic text " & Chr(8212) & " review manually)"
            ElseIf IsInsideQuotes(rng, doc) Then
                severity = "possible_error"
                suggestion = ""
                issueText = issueText & " (in quoted text " & Chr(8212) & " review manually)"
            End If

            Set issue = CreateIssueDict(RULE_NAME, locStr, issueText, suggestion, rng.Start, rng.End, severity)
            issues.Add issue

ContinueSearch:
            ' Collapse range to end of current match to find next
            On Error Resume Next
            rng.Collapse wdCollapseEnd
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Do
            End If
            On Error GoTo 0
        Loop

NextWord:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check if a found term is in the exceptions list
' ============================================================
Private Function IsException(ByVal term As String, _
                              ByRef exceptions() As String) As Boolean
    Dim i As Long
    Dim lTerm As String
    lTerm = LCase(Trim(term))

    On Error Resume Next
    For i = LBound(exceptions) To UBound(exceptions)
        If LCase(Trim(exceptions(i))) = lTerm Then
            IsException = True
            Exit Function
        End If
    Next i
    On Error GoTo 0

    IsException = False
End Function

' ============================================================
'  PRIVATE: Build the parallel US/UK spelling arrays
'  ~95 pairs across all categories.
' ============================================================
Private Sub BuildSpellingArrays(ByRef usWords() As String, _
                                 ByRef ukWords() As String)
    Const PAIR_COUNT As Long = 95

    ReDim usWords(0 To PAIR_COUNT - 1)
    ReDim ukWords(0 To PAIR_COUNT - 1)

    Dim idx As Long
    idx = 0

    ' -- -or -> -our (25 pairs) ----------------------------
    usWords(idx) = "color":       ukWords(idx) = "colour":       idx = idx + 1
    usWords(idx) = "favor":       ukWords(idx) = "favour":       idx = idx + 1
    usWords(idx) = "honor":       ukWords(idx) = "honour":       idx = idx + 1
    usWords(idx) = "humor":       ukWords(idx) = "humour":       idx = idx + 1
    usWords(idx) = "labor":       ukWords(idx) = "labour":       idx = idx + 1
    usWords(idx) = "neighbor":    ukWords(idx) = "neighbour":    idx = idx + 1
    usWords(idx) = "behavior":    ukWords(idx) = "behaviour":    idx = idx + 1
    usWords(idx) = "endeavor":    ukWords(idx) = "endeavour":    idx = idx + 1
    usWords(idx) = "harbor":      ukWords(idx) = "harbour":      idx = idx + 1
    usWords(idx) = "vigor":       ukWords(idx) = "vigour":       idx = idx + 1
    usWords(idx) = "valor":       ukWords(idx) = "valour":       idx = idx + 1
    usWords(idx) = "candor":      ukWords(idx) = "candour":      idx = idx + 1
    usWords(idx) = "clamor":      ukWords(idx) = "clamour":      idx = idx + 1
    usWords(idx) = "glamor":      ukWords(idx) = "glamour":      idx = idx + 1
    usWords(idx) = "parlor":      ukWords(idx) = "parlour":      idx = idx + 1
    usWords(idx) = "rancor":      ukWords(idx) = "rancour":      idx = idx + 1
    usWords(idx) = "rigor":       ukWords(idx) = "rigour":       idx = idx + 1
    usWords(idx) = "rumor":       ukWords(idx) = "rumour":       idx = idx + 1
    usWords(idx) = "savior":      ukWords(idx) = "saviour":      idx = idx + 1
    usWords(idx) = "splendor":    ukWords(idx) = "splendour":    idx = idx + 1
    usWords(idx) = "tumor":       ukWords(idx) = "tumour":       idx = idx + 1
    usWords(idx) = "vapor":       ukWords(idx) = "vapour":       idx = idx + 1
    usWords(idx) = "fervor":      ukWords(idx) = "fervour":      idx = idx + 1
    usWords(idx) = "armor":       ukWords(idx) = "armour":       idx = idx + 1
    usWords(idx) = "flavor":      ukWords(idx) = "flavour":      idx = idx + 1

    ' -- -ize -> -ise (20 pairs) ---------------------------
    usWords(idx) = "organize":      ukWords(idx) = "organise":      idx = idx + 1
    usWords(idx) = "realize":       ukWords(idx) = "realise":       idx = idx + 1
    usWords(idx) = "recognize":     ukWords(idx) = "recognise":     idx = idx + 1
    usWords(idx) = "authorize":     ukWords(idx) = "authorise":     idx = idx + 1
    usWords(idx) = "characterize":  ukWords(idx) = "characterise":  idx = idx + 1
    usWords(idx) = "customize":     ukWords(idx) = "customise":     idx = idx + 1
    usWords(idx) = "emphasize":     ukWords(idx) = "emphasise":     idx = idx + 1
    usWords(idx) = "finalize":      ukWords(idx) = "finalise":      idx = idx + 1
    usWords(idx) = "maximize":      ukWords(idx) = "maximise":      idx = idx + 1
    usWords(idx) = "minimize":      ukWords(idx) = "minimise":      idx = idx + 1
    usWords(idx) = "normalize":     ukWords(idx) = "normalise":     idx = idx + 1
    usWords(idx) = "optimize":      ukWords(idx) = "optimise":      idx = idx + 1
    usWords(idx) = "prioritize":    ukWords(idx) = "prioritise":    idx = idx + 1
    usWords(idx) = "standardize":   ukWords(idx) = "standardise":   idx = idx + 1
    usWords(idx) = "summarize":     ukWords(idx) = "summarise":     idx = idx + 1
    usWords(idx) = "symbolize":     ukWords(idx) = "symbolise":     idx = idx + 1
    usWords(idx) = "utilize":       ukWords(idx) = "utilise":       idx = idx + 1
    usWords(idx) = "apologize":     ukWords(idx) = "apologise":     idx = idx + 1
    usWords(idx) = "capitalize":    ukWords(idx) = "capitalise":    idx = idx + 1
    usWords(idx) = "criticize":     ukWords(idx) = "criticise":     idx = idx + 1

    ' -- -ization -> -isation (10 pairs) -------------------
    usWords(idx) = "organization":     ukWords(idx) = "organisation":     idx = idx + 1
    usWords(idx) = "authorization":    ukWords(idx) = "authorisation":    idx = idx + 1
    usWords(idx) = "characterization": ukWords(idx) = "characterisation": idx = idx + 1
    usWords(idx) = "customization":    ukWords(idx) = "customisation":    idx = idx + 1
    usWords(idx) = "optimization":     ukWords(idx) = "optimisation":     idx = idx + 1
    usWords(idx) = "normalization":    ukWords(idx) = "normalisation":    idx = idx + 1
    usWords(idx) = "realization":      ukWords(idx) = "realisation":      idx = idx + 1
    usWords(idx) = "utilization":      ukWords(idx) = "utilisation":      idx = idx + 1
    usWords(idx) = "specialization":   ukWords(idx) = "specialisation":   idx = idx + 1
    usWords(idx) = "globalization":    ukWords(idx) = "globalisation":    idx = idx + 1

    ' -- -er -> -re (10 pairs) -----------------------------
    usWords(idx) = "center":   ukWords(idx) = "centre":   idx = idx + 1
    usWords(idx) = "fiber":    ukWords(idx) = "fibre":    idx = idx + 1
    usWords(idx) = "liter":    ukWords(idx) = "litre":    idx = idx + 1
    usWords(idx) = "meter":    ukWords(idx) = "metre":    idx = idx + 1
    usWords(idx) = "theater":  ukWords(idx) = "theatre":  idx = idx + 1
    usWords(idx) = "somber":   ukWords(idx) = "sombre":   idx = idx + 1
    usWords(idx) = "caliber":  ukWords(idx) = "calibre":  idx = idx + 1
    usWords(idx) = "saber":    ukWords(idx) = "sabre":    idx = idx + 1
    usWords(idx) = "specter":  ukWords(idx) = "spectre":  idx = idx + 1
    usWords(idx) = "meager":   ukWords(idx) = "meagre":   idx = idx + 1

    ' -- -se -> -ce (3 pairs) ------------------------------
    usWords(idx) = "defense":   ukWords(idx) = "defence":   idx = idx + 1
    usWords(idx) = "offense":   ukWords(idx) = "offence":   idx = idx + 1
    usWords(idx) = "pretense":  ukWords(idx) = "pretence":  idx = idx + 1

    ' -- -og -> -ogue (6 pairs) ----------------------------
    usWords(idx) = "analog":   ukWords(idx) = "analogue":   idx = idx + 1
    usWords(idx) = "catalog":  ukWords(idx) = "catalogue":  idx = idx + 1
    usWords(idx) = "dialog":   ukWords(idx) = "dialogue":   idx = idx + 1
    usWords(idx) = "monolog":  ukWords(idx) = "monologue":  idx = idx + 1
    usWords(idx) = "prolog":   ukWords(idx) = "prologue":   idx = idx + 1
    usWords(idx) = "epilog":   ukWords(idx) = "epilogue":   idx = idx + 1

    ' -- -ment variants (5 pairs) -------------------------
    usWords(idx) = "judgment":        ukWords(idx) = "judgement":        idx = idx + 1
    usWords(idx) = "acknowledgment":  ukWords(idx) = "acknowledgement":  idx = idx + 1
    usWords(idx) = "fulfillment":     ukWords(idx) = "fulfilment":       idx = idx + 1
    usWords(idx) = "enrollment":      ukWords(idx) = "enrolment":        idx = idx + 1
    usWords(idx) = "installment":     ukWords(idx) = "instalment":       idx = idx + 1

    ' -- Other / miscellaneous (16 pairs) -----------------
    usWords(idx) = "gray":        ukWords(idx) = "grey":        idx = idx + 1
    usWords(idx) = "plow":        ukWords(idx) = "plough":      idx = idx + 1
    usWords(idx) = "tire":        ukWords(idx) = "tyre":        idx = idx + 1
    usWords(idx) = "check":       ukWords(idx) = "cheque":      idx = idx + 1
    usWords(idx) = "skeptic":     ukWords(idx) = "sceptic":     idx = idx + 1
    usWords(idx) = "aluminum":    ukWords(idx) = "aluminium":   idx = idx + 1
    usWords(idx) = "maneuver":    ukWords(idx) = "manoeuvre":    idx = idx + 1
    usWords(idx) = "artifact":    ukWords(idx) = "artefact":    idx = idx + 1
    usWords(idx) = "pediatric":   ukWords(idx) = "paediatric":  idx = idx + 1
    usWords(idx) = "anesthetic":  ukWords(idx) = "anaesthetic": idx = idx + 1
    usWords(idx) = "estrogen":    ukWords(idx) = "oestrogen":   idx = idx + 1
    usWords(idx) = "aging":       ukWords(idx) = "ageing":      idx = idx + 1
    usWords(idx) = "ax":          ukWords(idx) = "axe":         idx = idx + 1
    usWords(idx) = "program":     ukWords(idx) = "programme":   idx = idx + 1
    usWords(idx) = "curb":        ukWords(idx) = "kerb":        idx = idx + 1
    usWords(idx) = "draft":       ukWords(idx) = "draught":     idx = idx + 1

    ' idx should now equal PAIR_COUNT (95)
End Sub

' ============================================================
'  PRIVATE: Check if a range is italic
' ============================================================
Private Function IsRangeItalic(rng As Range) As Boolean
    On Error Resume Next
    Dim italicVal As Long
    italicVal = rng.Font.Italic
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsRangeItalic = False
        Exit Function
    End If
    On Error GoTo 0

    ' wdTrue = -1, True = -1; wdUndefined = 9999999 (mixed)
    IsRangeItalic = (italicVal = -1)
End Function

' ============================================================
'  PRIVATE: Check if a range is inside quotation marks
'  Looks at the character immediately before and after the
'  range for smart quotes, straight quotes, or single quotes.
' ============================================================
Private Function IsInsideQuotes(rng As Range, doc As Document) As Boolean
    Dim charBefore As String
    Dim charAfter As String

    On Error Resume Next

    ' Get character before range
    If rng.Start > 0 Then
        charBefore = doc.Range(rng.Start - 1, rng.Start).Text
    Else
        charBefore = ""
    End If
    If Err.Number <> 0 Then
        charBefore = ""
        Err.Clear
    End If

    ' Get character after range
    If rng.End < doc.Content.End Then
        charAfter = doc.Range(rng.End, rng.End + 1).Text
    Else
        charAfter = ""
    End If
    If Err.Number <> 0 Then
        charAfter = ""
        Err.Clear
    End If
    On Error GoTo 0

    ' Check for opening + closing quotes around the word
    ' This catches "word" and 'word' and similar
    If IsOpeningQuote(charBefore) And IsClosingQuote(charAfter) Then
        IsInsideQuotes = True
        Exit Function
    End If

    ' Broader check: scan backward for an unmatched opening quote
    ' within 200 characters
    Dim lookbackStart As Long
    lookbackStart = rng.Start - 200
    If lookbackStart < 0 Then lookbackStart = 0

    On Error Resume Next
    Dim beforeText As String
    beforeText = doc.Range(lookbackStart, rng.Start).Text
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsInsideQuotes = False
        Exit Function
    End If
    On Error GoTo 0

    ' Count open vs close quotes in the preceding text
    Dim openCount As Long
    Dim closeCount As Long
    Dim ch As String
    Dim c As Long
    openCount = 0: closeCount = 0
    For c = 1 To Len(beforeText)
        ch = Mid(beforeText, c, 1)
        If IsOpeningQuote(ch) Then openCount = openCount + 1
        If IsClosingQuote(ch) Then closeCount = closeCount + 1
    Next c

    ' If there are more opens than closes, we're inside quotes
    IsInsideQuotes = (openCount > closeCount)
End Function

' ============================================================
'  PRIVATE: Check if a character is an opening quote
' ============================================================
Private Function IsOpeningQuote(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsOpeningQuote = False
        Exit Function
    End If
    Select Case AscW(ch)
        Case 8220  ' left double smart quote "
            IsOpeningQuote = True
        Case 8216  ' left single smart quote '
            IsOpeningQuote = True
        Case Else
            IsOpeningQuote = False
    End Select
End Function

' ============================================================
'  PRIVATE: Check if a character is a closing quote
' ============================================================
Private Function IsClosingQuote(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsClosingQuote = False
        Exit Function
    End If
    Select Case AscW(ch)
        Case 8221  ' right double smart quote "
            IsClosingQuote = True
        Case 8217  ' right single smart quote '
            IsClosingQuote = True
        Case Else
            IsClosingQuote = False
    End Select
End Function

' ================================================================
' ================================================================
'  RULE 12 -- LICENCE / LICENSE
' ================================================================
' ================================================================

' ============================================================
'  MAIN ENTRY POINT -- Licence/License
' ============================================================
Public Function Check_LicenceLicense(doc As Document) As Collection
    Dim issues As New Collection

    ' Search for both spellings in the document body
    SearchForLicenceIssues doc.Content, doc, issues

    ' Search footnotes
    On Error Resume Next
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        Err.Clear
        SearchForLicenceIssues fn.Range, doc, issues
        If Err.Number <> 0 Then Err.Clear
    Next fn
    On Error GoTo 0

    ' Search endnotes
    On Error Resume Next
    Dim en As Endnote
    For Each en In doc.Endnotes
        Err.Clear
        SearchForLicenceIssues en.Range, doc, issues
        If Err.Number <> 0 Then Err.Clear
    Next en
    On Error GoTo 0

    Set Check_LicenceLicense = issues
End Function

' ============================================================
'  PRIVATE: Search a range for licence/license issues
' ============================================================
Private Sub SearchForLicenceIssues(searchRange As Range, _
                                    doc As Document, _
                                    ByRef issues As Collection)
    Dim searchTerms As Variant
    Dim t As Long

    ' Search for the base forms; skip derivatives that are always correct
    searchTerms = Array("licence", "license", "sub-licence", "sub-license", _
                        "re-licence", "re-license")

    For t = LBound(searchTerms) To UBound(searchTerms)
        SearchSingleLicenceTerm CStr(searchTerms(t)), searchRange, doc, issues
    Next t
End Sub

' ============================================================
'  PRIVATE: Search for a single licence/license term and
'  analyse context
' ============================================================
Private Sub SearchSingleLicenceTerm(ByVal term As String, _
                              searchRange As Range, _
                              doc As Document, _
                              ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As Object
    Dim locStr As String
    Dim contextBefore As String
    Dim contextAfter As String
    Dim wordBefore As String
    Dim wordAfter As String
    Dim issueText As String
    Dim suggestion As String
    Dim usesS As Boolean
    Dim baseIsNoun As Boolean
    Dim baseIsVerb As Boolean

    On Error Resume Next
    Set rng = searchRange.Duplicate
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    With rng.Find
        .ClearFormatting
        .Text = term
        .MatchWholeWord = True
        .MatchCase = False
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Do
        End If
        On Error GoTo 0

        If Not found Then Exit Do

        ' Skip if outside page range
        If Not EngineIsInPageRange(rng) Then
            GoTo ContinueLicenceSearch
        End If

        ' Determine if the found word uses -s- or -c-
        usesS = (InStr(1, LCase(rng.Text), "license") > 0)

        ' Skip "licensed" and "licensing" -- always correct with -s-
        Dim foundLower As String
        foundLower = LCase(Trim(rng.Text))
        If foundLower = "licensed" Or foundLower = "licensing" Then
            GoTo ContinueLicenceSearch
        End If

        ' -- Downgrade italic / quoted text ------------------
        Dim licSeverity As String
        licSeverity = "possible_error"

        If IsRangeItalic(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = CreateIssueDict(RULE_NAME_LICENCE, locStr, ")
            issues.Add issue
            GoTo ContinueLicenceSearch
        End If

        If IsInsideQuotes(rng, doc) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = CreateIssueDict(RULE_NAME_LICENCE, locStr, ")
            issues.Add issue
            GoTo ContinueLicenceSearch
        End If

        ' -- Get surrounding context --------------------------
        contextBefore = GetLicenceContextBefore(rng, doc, 50)
        contextAfter = GetLicenceContextAfter(rng, doc, 50)

        ' Extract the last word before the match
        wordBefore = GetLastWordFromContext(contextBefore)

        ' Extract the first word after the match
        wordAfter = GetFirstWordFromContext(contextAfter)

        ' -- Determine noun or verb context -------------------
        baseIsVerb = IsVerbIndicator(wordBefore)
        baseIsNoun = IsNounIndicator(wordBefore) Or IsNounFollower(wordAfter)

        ' -- Decide if there is an issue ----------------------
        issueText = ""
        suggestion = ""

        If usesS And baseIsNoun And Not baseIsVerb Then
            ' "license" used in noun context -- should be "licence"
            issueText = "'" & rng.Text & "' appears in a noun context; " & _
                        "UK convention uses 'licence' for the noun"
            suggestion = ReplaceSWithC(rng.Text)
        ElseIf Not usesS And baseIsVerb And Not baseIsNoun Then
            ' "licence" used in verb context -- should be "license"
            issueText = "'" & rng.Text & "' appears in a verb context; " & _
                        "UK convention uses 'license' for the verb"
            suggestion = ReplaceCWithS(rng.Text)
        ElseIf (usesS And Not baseIsVerb And Not baseIsNoun) Or _
               (Not usesS And Not baseIsVerb And Not baseIsNoun) Then
            ' Context ambiguous
            issueText = "'" & rng.Text & "' " & Chr(8212) & " unable to determine noun/verb context; " & _
                        "review context to ensure correct UK spelling"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present -- ambiguous
            issueText = "'" & rng.Text & "' " & Chr(8212) & " conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf Not usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present -- ambiguous
            issueText = "'" & rng.Text & "' " & Chr(8212) & " conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        End If

        ' Only create issue if we found something to flag
        If Len(issueText) > 0 Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            Set issue = CreateIssueDict(RULE_NAME_LICENCE, locStr, issueText, suggestion, rng.Start, rng.End, "possible_error")
            issues.Add issue
        End If

ContinueLicenceSearch:
        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Do
        End If
        On Error GoTo 0
    Loop
End Sub

' ============================================================
'  PRIVATE: Get text before the match range (up to N chars)
' ============================================================
Private Function GetLicenceContextBefore(rng As Range, doc As Document, _
                                   ByVal charCount As Long) As String
    Dim startPos As Long
    Dim contextRng As Range

    On Error Resume Next
    startPos = rng.Start - charCount
    If startPos < 0 Then startPos = 0

    Set contextRng = doc.Range(startPos, rng.Start)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        GetLicenceContextBefore = ""
        Exit Function
    End If
    On Error GoTo 0

    GetLicenceContextBefore = contextRng.Text
End Function

' ============================================================
'  PRIVATE: Get text after the match range (up to N chars)
' ============================================================
Private Function GetLicenceContextAfter(rng As Range, doc As Document, _
                                  ByVal charCount As Long) As String
    Dim endPos As Long
    Dim contextRng As Range
    Dim docEnd As Long

    On Error Resume Next
    docEnd = doc.Content.End
    endPos = rng.End + charCount
    If endPos > docEnd Then endPos = docEnd

    Set contextRng = doc.Range(rng.End, endPos)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        GetLicenceContextAfter = ""
        Exit Function
    End If
    On Error GoTo 0

    GetLicenceContextAfter = contextRng.Text
End Function

' ============================================================
'  PRIVATE: Extract the last word from a context string
' ============================================================
Private Function GetLastWordFromContext(ByVal text As String) As String
    Dim trimmed As String
    Dim i As Long
    Dim ch As String

    trimmed = Trim(text)
    If Len(trimmed) = 0 Then
        GetLastWordFromContext = ""
        Exit Function
    End If

    ' Walk backward from end to find last word boundary
    For i = Len(trimmed) To 1 Step -1
        ch = Mid(trimmed, i, 1)
        If ch = " " Or ch = vbCr Or ch = vbLf Or ch = vbTab Then
            GetLastWordFromContext = LCase(Mid(trimmed, i + 1))
            Exit Function
        End If
    Next i

    GetLastWordFromContext = LCase(trimmed)
End Function

' ============================================================
'  PRIVATE: Extract the first word from a context string
' ============================================================
Private Function GetFirstWordFromContext(ByVal text As String) As String
    Dim trimmed As String
    Dim spacePos As Long

    trimmed = Trim(text)
    If Len(trimmed) = 0 Then
        GetFirstWordFromContext = ""
        Exit Function
    End If

    spacePos = InStr(1, trimmed, " ")
    If spacePos > 0 Then
        GetFirstWordFromContext = LCase(Left(trimmed, spacePos - 1))
    Else
        GetFirstWordFromContext = LCase(trimmed)
    End If

    ' Strip trailing punctuation
    Dim result As String
    Dim pch As String
    result = GetFirstWordFromContext
    Do While Len(result) > 0
        pch = Right(result, 1)
        If pch Like "[A-Za-z]" Then Exit Do
        result = Left(result, Len(result) - 1)
    Loop
    GetFirstWordFromContext = result
End Function

' ============================================================
'  PRIVATE: Check if a word is a verb indicator
' ============================================================
Private Function IsVerbIndicator(ByVal word As String) As Boolean
    Dim indicators As Variant
    Dim i As Long

    indicators = Array("to", "will", "shall", "may", "must", _
                       "can", "should", "would", "not")

    word = LCase(Trim(word))
    For i = LBound(indicators) To UBound(indicators)
        If word = CStr(indicators(i)) Then
            IsVerbIndicator = True
            Exit Function
        End If
    Next i

    IsVerbIndicator = False
End Function

' ============================================================
'  PRIVATE: Check if a word is a noun indicator
' ============================================================
Private Function IsNounIndicator(ByVal word As String) As Boolean
    Dim indicators As Variant
    Dim i As Long

    indicators = Array("a", "an", "the", "this", "that", "such", _
                       "said", "its", "their", "our", "your", "his", "her")

    word = LCase(Trim(word))
    For i = LBound(indicators) To UBound(indicators)
        If word = CStr(indicators(i)) Then
            IsNounIndicator = True
            Exit Function
        End If
    Next i

    IsNounIndicator = False
End Function

' ============================================================
'  PRIVATE: Check if the word after indicates noun usage
' ============================================================
Private Function IsNounFollower(ByVal word As String) As Boolean
    Dim followers As Variant
    Dim i As Long

    followers = Array("agreement", "holder", "fee", "number", _
                      "plate", "condition")

    word = LCase(Trim(word))
    For i = LBound(followers) To UBound(followers)
        If word = CStr(followers(i)) Then
            IsNounFollower = True
            Exit Function
        End If
    Next i

    IsNounFollower = False
End Function

' ============================================================
'  PRIVATE: Replace -s- with -c- in licence/license words
' ============================================================
Private Function ReplaceSWithC(ByVal word As String) As String
    ReplaceSWithC = Replace(word, "license", "licence", , , vbTextCompare)
    ReplaceSWithC = Replace(ReplaceSWithC, "License", "Licence", , , vbBinaryCompare)
End Function

' ============================================================
'  PRIVATE: Replace -c- with -s- in licence/license words
' ============================================================
Private Function ReplaceCWithS(ByVal word As String) As String
    ReplaceCWithS = Replace(word, "licence", "license", , , vbTextCompare)
    ReplaceCWithS = Replace(ReplaceCWithS, "Licence", "License", , , vbBinaryCompare)
End Function

' ================================================================
' ================================================================
'  RULE 13 -- COLOUR FORMATTING
' ================================================================
' ================================================================

' ============================================================
'  MAIN ENTRY POINT -- Colour Formatting
' ============================================================
Public Function Check_ColourFormatting(doc As Document) As Collection
    Dim issues As New Collection
    Dim colourCounts As Object ' Scripting.Dictionary
    Dim para As Paragraph
    Dim rn As Range
    Dim runColor As Long
    Dim dominantColour As Long
    Dim maxCount As Long
    Dim runText As String

    ' -- First pass: count colour usage per run ---------------
    Set colourCounts = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Dim paraRange As Range
        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass1
        End If

        ' Skip paragraphs outside page range
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextParaPass1
        End If

        ' Iterate runs within the paragraph
        Dim r As Long
        For r = 1 To paraRange.Runs.Count
            Err.Clear
            Set rn = paraRange.Runs(r)
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextRunPass1
            End If

            ' Skip whitespace-only runs
            runText = rn.Text
            If Len(Trim(Replace(Replace(runText, vbCr, ""), vbLf, ""))) = 0 Then
                GoTo NextRunPass1
            End If

            runColor = rn.Font.Color
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextRunPass1
            End If

            If colourCounts.Exists(runColor) Then
                colourCounts(runColor) = colourCounts(runColor) + 1
            Else
                colourCounts.Add runColor, 1
            End If

NextRunPass1:
        Next r

NextParaPass1:
    Next para
    On Error GoTo 0

    ' -- Determine dominant colour ----------------------------
    If colourCounts.Count = 0 Then
        Set Check_ColourFormatting = issues
        Exit Function
    End If

    dominantColour = 0
    maxCount = 0
    Dim colourKey As Variant
    For Each colourKey In colourCounts.keys
        If colourCounts(colourKey) > maxCount Then
            maxCount = colourCounts(colourKey)
            dominantColour = CLng(colourKey)
        End If
    Next colourKey

    ' -- Second pass: flag non-dominant, non-automatic colours -
    Const WD_COLOR_AUTOMATIC As Long = -16777216

    ' Tracking for grouping consecutive same-colour runs
    Dim groupStartPos As Long
    Dim groupEndPos As Long
    Dim groupColour As Long
    Dim groupActive As Boolean
    Dim groupParaRange As Range

    groupActive = False

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass2
        End If

        ' Skip paragraphs outside page range
        If Not EngineIsInPageRange(paraRange) Then
            ' Flush any active group before skipping
            If groupActive Then
                FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                groupActive = False
            End If
            GoTo NextParaPass2
        End If

        ' Skip heading-styled paragraphs (may have intentional colour)
        Dim styleName As String
        styleName = ""
        styleName = para.Style.NameLocal
        If Err.Number <> 0 Then
            Err.Clear
            styleName = ""
        End If
        If LCase(Left(styleName, 7)) = "heading" Then
            If groupActive Then
                FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                groupActive = False
            End If
            GoTo NextParaPass2
        End If

        ' Iterate runs
        For r = 1 To paraRange.Runs.Count
            Err.Clear
            Set rn = paraRange.Runs(r)
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextRunPass2
            End If

            ' Skip whitespace-only runs
            runText = rn.Text
            If Len(Trim(Replace(Replace(runText, vbCr, ""), vbLf, ""))) = 0 Then
                GoTo NextRunPass2
            End If

            runColor = rn.Font.Color
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextRunPass2
            End If

            ' Skip if colour matches dominant or is automatic
            If runColor = dominantColour Or runColor = WD_COLOR_AUTOMATIC Then
                ' Flush any active group
                If groupActive Then
                    FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                    groupActive = False
                End If
                GoTo NextRunPass2
            End If

            ' Skip hyperlinks
            If IsRunInsideHyperlink(rn, doc) Then
                If groupActive Then
                    FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                    groupActive = False
                End If
                GoTo NextRunPass2
            End If

            ' -- This run has a non-standard colour -----------
            If groupActive And runColor = groupColour And _
               rn.Start = groupEndPos Then
                ' Extend existing group
                groupEndPos = rn.End
            Else
                ' Flush previous group if any
                If groupActive Then
                    FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                End If
                ' Start new group
                groupStartPos = rn.Start
                groupEndPos = rn.End
                groupColour = runColor
                groupActive = True
            End If

NextRunPass2:
        Next r

NextParaPass2:
    Next para

    ' Flush final group
    If groupActive Then
        FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
    End If
    On Error GoTo 0

    Set Check_ColourFormatting = issues
End Function

' ============================================================
'  PRIVATE: Flush a grouped colour issue
' ============================================================
Private Sub FlushColourGroup(doc As Document, _
                              ByRef issues As Collection, _
                              ByVal startPos As Long, _
                              ByVal endPos As Long, _
                              ByVal fontColor As Long)
    Dim issue As Object
    Dim locStr As String
    Dim hexStr As String
    Dim rng As Range

    hexStr = ColourToHex(fontColor)

    On Error Resume Next
    Set rng = doc.Range(startPos, endPos)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    locStr = EngineGetLocationString(rng, doc)
    If Err.Number <> 0 Then
        locStr = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0

    Dim previewText As String
    On Error Resume Next
    previewText = Left(rng.Text, 60)
    If Err.Number <> 0 Then
        previewText = "(text unavailable)"
        Err.Clear
    End If
    On Error GoTo 0

    Set issue = CreateIssueDict(RULE_NAME_COLOUR, locStr, "Non-standard font colour " & hexStr & " detected:)
    issues.Add issue
End Sub

' ============================================================
'  PRIVATE: Convert a Long colour value to hex string
' ============================================================
Private Function ColourToHex(ByVal colorVal As Long) As String
    Dim cR As Long
    Dim cG As Long
    Dim cB As Long

    ' Word stores colours as BGR in Long format
    cR = colorVal Mod 256
    cG = (colorVal \ 256) Mod 256
    cB = (colorVal \ 65536) Mod 256

    ColourToHex = "#" & Right("0" & Hex(cR), 2) & _
                        Right("0" & Hex(cG), 2) & _
                        Right("0" & Hex(cB), 2)
End Function

' ============================================================
'  PRIVATE: Check if a run is inside a hyperlink
' ============================================================
Private Function IsRunInsideHyperlink(rn As Range, doc As Document) As Boolean
    Dim hl As Hyperlink

    On Error Resume Next
    For Each hl In doc.Hyperlinks
        Err.Clear
        If hl.Range.Start <= rn.Start And hl.Range.End >= rn.End Then
            IsRunInsideHyperlink = True
            Exit Function
        End If
        If Err.Number <> 0 Then
            Err.Clear
        End If
    Next hl
    On Error GoTo 0

    IsRunInsideHyperlink = False
End Function

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineIsInPageRange
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineGetLocationString
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for PleadingsEngine.IsWhitelistedTerm
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for PleadingsEngine.GetSpellingMode
' ----------------------------------------------------------------

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
'  Late-bound wrapper: PleadingsEngine.IsWhitelistedTerm
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetSpellingMode
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
'  Late-bound wrapper: PleadingsEngine.IsWhitelistedTerm
' ----------------------------------------------------------------
Private Function EngineIsWhitelistedTerm(ByVal term As String) As Boolean
    On Error Resume Next
    EngineIsWhitelistedTerm = Application.Run("PleadingsEngine.IsWhitelistedTerm", term)
    If Err.Number <> 0 Then
        EngineIsWhitelistedTerm = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetSpellingMode
' ----------------------------------------------------------------
Private Function EngineGetSpellingMode() As String
    On Error Resume Next
    EngineGetSpellingMode = Application.Run("PleadingsEngine.GetSpellingMode")
    If Err.Number <> 0 Then
        EngineGetSpellingMode = "UK"
        Err.Clear
    End If
    On Error GoTo 0
End Function
