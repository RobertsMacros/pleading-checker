Attribute VB_Name = "Rules_Spelling"
' ============================================================
' Rules_Spelling.bas
' Combined proofreading rules for UK/US English spelling.
'
' Rule 1 -- British/US Spelling:
'   Detects ~133 spelling differences between US and UK English,
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
Private Const RULE_NAME_CHECK As String = "check_cheque"

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
'  direction = "UK" or "US" -- controls the finding text:
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
    Dim finding As Object
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

            ' -- Create the finding ----------------------
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
                issueText = issueText & " (in italic text -- review manually)"
            ElseIf IsInsideQuotes(rng, doc) Then
                severity = "possible_error"
                suggestion = ""
                issueText = issueText & " (in quoted text -- review manually)"
            End If

            Set finding = CreateIssueDict(RULE_NAME, locStr, issueText, suggestion, rng.Start, rng.End, severity)
            issues.Add finding

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

    For i = LBound(exceptions) To UBound(exceptions)
        If LCase(Trim(exceptions(i))) = lTerm Then
            IsException = True
            Exit Function
        End If
    Next i

    IsException = False
End Function

' ============================================================
'  PRIVATE: Build the parallel US/UK spelling arrays
'  ~95 pairs across all categories.
' ============================================================
Private Sub BuildSpellingArrays(ByRef usWords() As String, _
                                 ByRef ukWords() As String)
    ' Dynamic pair building -- no hard-coded PAIR_COUNT needed.
    ' Only low-risk, non-contentious US-to-UK variants are included.
    ' Excluded: check/cheque, practice/practise, license/licence,
    '   judgment/judgement, program/programme, draft/draught,
    '   tire/tyre, curb/kerb, story/storey, meter/metre,
    '   sulphur/sulfur, medical/scientific variants.
    Dim pairCount As Long
    pairCount = 0
    ReDim usWords(0 To 255)
    ReDim ukWords(0 To 255)

    ' -- -or -> -our and inflections --
    AddSpellingPair usWords, ukWords, pairCount, "color", "colour"
    AddSpellingPair usWords, ukWords, pairCount, "colors", "colours"
    AddSpellingPair usWords, ukWords, pairCount, "colored", "coloured"
    AddSpellingPair usWords, ukWords, pairCount, "coloring", "colouring"
    AddSpellingPair usWords, ukWords, pairCount, "favor", "favour"
    AddSpellingPair usWords, ukWords, pairCount, "favored", "favoured"
    AddSpellingPair usWords, ukWords, pairCount, "favoring", "favouring"
    AddSpellingPair usWords, ukWords, pairCount, "favorite", "favourite"
    AddSpellingPair usWords, ukWords, pairCount, "favorites", "favourites"
    AddSpellingPair usWords, ukWords, pairCount, "honor", "honour"
    AddSpellingPair usWords, ukWords, pairCount, "honors", "honours"
    AddSpellingPair usWords, ukWords, pairCount, "honored", "honoured"
    AddSpellingPair usWords, ukWords, pairCount, "honoring", "honouring"
    AddSpellingPair usWords, ukWords, pairCount, "humor", "humour"
    AddSpellingPair usWords, ukWords, pairCount, "labor", "labour"
    AddSpellingPair usWords, ukWords, pairCount, "labored", "laboured"
    AddSpellingPair usWords, ukWords, pairCount, "laboring", "labouring"
    AddSpellingPair usWords, ukWords, pairCount, "neighbor", "neighbour"
    AddSpellingPair usWords, ukWords, pairCount, "neighbors", "neighbours"
    AddSpellingPair usWords, ukWords, pairCount, "neighboring", "neighbouring"
    AddSpellingPair usWords, ukWords, pairCount, "neighborhood", "neighbourhood"
    AddSpellingPair usWords, ukWords, pairCount, "behavior", "behaviour"
    AddSpellingPair usWords, ukWords, pairCount, "behaviors", "behaviours"
    AddSpellingPair usWords, ukWords, pairCount, "behavioral", "behavioural"
    AddSpellingPair usWords, ukWords, pairCount, "endeavor", "endeavour"
    AddSpellingPair usWords, ukWords, pairCount, "endeavored", "endeavoured"
    AddSpellingPair usWords, ukWords, pairCount, "endeavoring", "endeavouring"
    AddSpellingPair usWords, ukWords, pairCount, "harbor", "harbour"
    AddSpellingPair usWords, ukWords, pairCount, "harbors", "harbours"
    AddSpellingPair usWords, ukWords, pairCount, "vigor", "vigour"
    AddSpellingPair usWords, ukWords, pairCount, "valor", "valour"
    AddSpellingPair usWords, ukWords, pairCount, "candor", "candour"
    AddSpellingPair usWords, ukWords, pairCount, "clamor", "clamour"
    AddSpellingPair usWords, ukWords, pairCount, "glamor", "glamour"
    AddSpellingPair usWords, ukWords, pairCount, "parlor", "parlour"
    AddSpellingPair usWords, ukWords, pairCount, "rancor", "rancour"
    AddSpellingPair usWords, ukWords, pairCount, "rigor", "rigour"
    AddSpellingPair usWords, ukWords, pairCount, "rumor", "rumour"
    AddSpellingPair usWords, ukWords, pairCount, "rumors", "rumours"
    AddSpellingPair usWords, ukWords, pairCount, "savior", "saviour"
    AddSpellingPair usWords, ukWords, pairCount, "splendor", "splendour"
    AddSpellingPair usWords, ukWords, pairCount, "tumor", "tumour"
    AddSpellingPair usWords, ukWords, pairCount, "tumors", "tumours"
    AddSpellingPair usWords, ukWords, pairCount, "vapor", "vapour"
    AddSpellingPair usWords, ukWords, pairCount, "fervor", "fervour"
    AddSpellingPair usWords, ukWords, pairCount, "armor", "armour"
    AddSpellingPair usWords, ukWords, pairCount, "armored", "armoured"
    AddSpellingPair usWords, ukWords, pairCount, "flavor", "flavour"
    AddSpellingPair usWords, ukWords, pairCount, "flavors", "flavours"
    AddSpellingPair usWords, ukWords, pairCount, "flavored", "flavoured"
    AddSpellingPair usWords, ukWords, pairCount, "flavoring", "flavouring"

    ' -- -er -> -re where generally safe --
    AddSpellingPair usWords, ukWords, pairCount, "center", "centre"
    AddSpellingPair usWords, ukWords, pairCount, "centers", "centres"
    AddSpellingPair usWords, ukWords, pairCount, "centered", "centred"
    AddSpellingPair usWords, ukWords, pairCount, "centering", "centring"
    AddSpellingPair usWords, ukWords, pairCount, "fiber", "fibre"
    AddSpellingPair usWords, ukWords, pairCount, "fibers", "fibres"
    AddSpellingPair usWords, ukWords, pairCount, "theater", "theatre"
    AddSpellingPair usWords, ukWords, pairCount, "theaters", "theatres"
    AddSpellingPair usWords, ukWords, pairCount, "somber", "sombre"
    AddSpellingPair usWords, ukWords, pairCount, "caliber", "calibre"
    AddSpellingPair usWords, ukWords, pairCount, "saber", "sabre"
    AddSpellingPair usWords, ukWords, pairCount, "specter", "spectre"
    AddSpellingPair usWords, ukWords, pairCount, "meager", "meagre"
    AddSpellingPair usWords, ukWords, pairCount, "luster", "lustre"
    AddSpellingPair usWords, ukWords, pairCount, "maneuver", "manoeuvre"
    AddSpellingPair usWords, ukWords, pairCount, "maneuvered", "manoeuvred"
    AddSpellingPair usWords, ukWords, pairCount, "maneuvering", "manoeuvring"
    AddSpellingPair usWords, ukWords, pairCount, "reconnoiter", "reconnoitre"
    AddSpellingPair usWords, ukWords, pairCount, "goiter", "goitre"
    AddSpellingPair usWords, ukWords, pairCount, "ocher", "ochre"

    ' -- -se -> -ce where safe --
    AddSpellingPair usWords, ukWords, pairCount, "defense", "defence"
    AddSpellingPair usWords, ukWords, pairCount, "defenses", "defences"
    AddSpellingPair usWords, ukWords, pairCount, "offense", "offence"
    AddSpellingPair usWords, ukWords, pairCount, "offenses", "offences"
    AddSpellingPair usWords, ukWords, pairCount, "pretense", "pretence"

    ' -- -og -> -ogue --
    AddSpellingPair usWords, ukWords, pairCount, "analog", "analogue"
    AddSpellingPair usWords, ukWords, pairCount, "catalog", "catalogue"
    AddSpellingPair usWords, ukWords, pairCount, "dialog", "dialogue"
    AddSpellingPair usWords, ukWords, pairCount, "monolog", "monologue"
    AddSpellingPair usWords, ukWords, pairCount, "prolog", "prologue"
    AddSpellingPair usWords, ukWords, pairCount, "epilog", "epilogue"

    ' -- -ment and similar safe variants --
    AddSpellingPair usWords, ukWords, pairCount, "acknowledgment", "acknowledgement"
    AddSpellingPair usWords, ukWords, pairCount, "acknowledgments", "acknowledgements"
    AddSpellingPair usWords, ukWords, pairCount, "fulfillment", "fulfilment"
    AddSpellingPair usWords, ukWords, pairCount, "fulfill", "fulfil"
    AddSpellingPair usWords, ukWords, pairCount, "enrollment", "enrolment"
    AddSpellingPair usWords, ukWords, pairCount, "enroll", "enrol"
    AddSpellingPair usWords, ukWords, pairCount, "installment", "instalment"
    AddSpellingPair usWords, ukWords, pairCount, "installments", "instalments"

    ' -- Doubled consonant variants --
    AddSpellingPair usWords, ukWords, pairCount, "traveled", "travelled"
    AddSpellingPair usWords, ukWords, pairCount, "traveling", "travelling"
    AddSpellingPair usWords, ukWords, pairCount, "traveler", "traveller"
    AddSpellingPair usWords, ukWords, pairCount, "travelers", "travellers"
    AddSpellingPair usWords, ukWords, pairCount, "canceled", "cancelled"
    AddSpellingPair usWords, ukWords, pairCount, "canceling", "cancelling"
    AddSpellingPair usWords, ukWords, pairCount, "labeled", "labelled"
    AddSpellingPair usWords, ukWords, pairCount, "labeling", "labelling"
    AddSpellingPair usWords, ukWords, pairCount, "modeled", "modelled"
    AddSpellingPair usWords, ukWords, pairCount, "modeling", "modelling"
    AddSpellingPair usWords, ukWords, pairCount, "counselor", "counsellor"
    AddSpellingPair usWords, ukWords, pairCount, "counselors", "counsellors"
    AddSpellingPair usWords, ukWords, pairCount, "counseling", "counselling"
    AddSpellingPair usWords, ukWords, pairCount, "signaled", "signalled"
    AddSpellingPair usWords, ukWords, pairCount, "signaling", "signalling"
    AddSpellingPair usWords, ukWords, pairCount, "fueled", "fuelled"
    AddSpellingPair usWords, ukWords, pairCount, "fueling", "fuelling"

    ' -- -ize -> -ise (safe subset) --
    AddSpellingPair usWords, ukWords, pairCount, "organize", "organise"
    AddSpellingPair usWords, ukWords, pairCount, "realize", "realise"
    AddSpellingPair usWords, ukWords, pairCount, "recognize", "recognise"
    AddSpellingPair usWords, ukWords, pairCount, "authorize", "authorise"
    AddSpellingPair usWords, ukWords, pairCount, "characterize", "characterise"
    AddSpellingPair usWords, ukWords, pairCount, "customize", "customise"
    AddSpellingPair usWords, ukWords, pairCount, "emphasize", "emphasise"
    AddSpellingPair usWords, ukWords, pairCount, "finalize", "finalise"
    AddSpellingPair usWords, ukWords, pairCount, "maximize", "maximise"
    AddSpellingPair usWords, ukWords, pairCount, "minimize", "minimise"
    AddSpellingPair usWords, ukWords, pairCount, "normalize", "normalise"
    AddSpellingPair usWords, ukWords, pairCount, "optimize", "optimise"
    AddSpellingPair usWords, ukWords, pairCount, "prioritize", "prioritise"
    AddSpellingPair usWords, ukWords, pairCount, "standardize", "standardise"
    AddSpellingPair usWords, ukWords, pairCount, "summarize", "summarise"
    AddSpellingPair usWords, ukWords, pairCount, "symbolize", "symbolise"
    AddSpellingPair usWords, ukWords, pairCount, "utilize", "utilise"
    AddSpellingPair usWords, ukWords, pairCount, "apologize", "apologise"
    AddSpellingPair usWords, ukWords, pairCount, "capitalize", "capitalise"
    AddSpellingPair usWords, ukWords, pairCount, "criticize", "criticise"
    AddSpellingPair usWords, ukWords, pairCount, "legalize", "legalise"
    AddSpellingPair usWords, ukWords, pairCount, "memorize", "memorise"
    AddSpellingPair usWords, ukWords, pairCount, "patronize", "patronise"
    AddSpellingPair usWords, ukWords, pairCount, "penalize", "penalise"
    AddSpellingPair usWords, ukWords, pairCount, "privatize", "privatise"
    AddSpellingPair usWords, ukWords, pairCount, "harmonize", "harmonise"
    AddSpellingPair usWords, ukWords, pairCount, "economize", "economise"
    AddSpellingPair usWords, ukWords, pairCount, "immunize", "immunise"
    AddSpellingPair usWords, ukWords, pairCount, "neutralize", "neutralise"
    AddSpellingPair usWords, ukWords, pairCount, "stabilize", "stabilise"

    ' -- -ization -> -isation --
    AddSpellingPair usWords, ukWords, pairCount, "organization", "organisation"
    AddSpellingPair usWords, ukWords, pairCount, "authorization", "authorisation"
    AddSpellingPair usWords, ukWords, pairCount, "characterization", "characterisation"
    AddSpellingPair usWords, ukWords, pairCount, "customization", "customisation"
    AddSpellingPair usWords, ukWords, pairCount, "optimization", "optimisation"
    AddSpellingPair usWords, ukWords, pairCount, "normalization", "normalisation"
    AddSpellingPair usWords, ukWords, pairCount, "realization", "realisation"
    AddSpellingPair usWords, ukWords, pairCount, "utilization", "utilisation"
    AddSpellingPair usWords, ukWords, pairCount, "specialization", "specialisation"
    AddSpellingPair usWords, ukWords, pairCount, "globalization", "globalisation"
    AddSpellingPair usWords, ukWords, pairCount, "legalization", "legalisation"
    AddSpellingPair usWords, ukWords, pairCount, "privatization", "privatisation"
    AddSpellingPair usWords, ukWords, pairCount, "harmonization", "harmonisation"
    AddSpellingPair usWords, ukWords, pairCount, "neutralization", "neutralisation"
    AddSpellingPair usWords, ukWords, pairCount, "stabilization", "stabilisation"

    ' -- Safe miscellaneous pairs --
    AddSpellingPair usWords, ukWords, pairCount, "gray", "grey"
    AddSpellingPair usWords, ukWords, pairCount, "plow", "plough"
    AddSpellingPair usWords, ukWords, pairCount, "skeptic", "sceptic"
    AddSpellingPair usWords, ukWords, pairCount, "skeptical", "sceptical"
    AddSpellingPair usWords, ukWords, pairCount, "aluminum", "aluminium"
    AddSpellingPair usWords, ukWords, pairCount, "artifact", "artefact"
    AddSpellingPair usWords, ukWords, pairCount, "aging", "ageing"
    AddSpellingPair usWords, ukWords, pairCount, "pajamas", "pyjamas"
    AddSpellingPair usWords, ukWords, pairCount, "cozy", "cosy"
    AddSpellingPair usWords, ukWords, pairCount, "donut", "doughnut"

    ' Trim arrays to actual size
    ReDim Preserve usWords(0 To pairCount - 1)
    ReDim Preserve ukWords(0 To pairCount - 1)
End Sub

Private Sub AddSpellingPair(ByRef usWords() As String, _
                             ByRef ukWords() As String, _
                             ByRef pairCount As Long, _
                             ByVal usWord As String, _
                             ByVal ukWord As String)
    ' Grow arrays if needed
    If pairCount > UBound(usWords) Then
        ReDim Preserve usWords(0 To UBound(usWords) + 128)
        ReDim Preserve ukWords(0 To UBound(ukWords) + 128)
    End If
    usWords(pairCount) = usWord
    ukWords(pairCount) = ukWord
    pairCount = pairCount + 1
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
    Dim finding As Object
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

            Set finding = CreateIssueDict(RULE_NAME_LICENCE, locStr, "'" & rng.Text & "' -- in italic text, review manually", "", rng.Start, rng.End, "possible_error")
            issues.Add finding
            GoTo ContinueLicenceSearch
        End If

        If IsInsideQuotes(rng, doc) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_LICENCE, locStr, "'" & rng.Text & "' -- in quoted text, review manually", "", rng.Start, rng.End, "possible_error")
            issues.Add finding
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

        ' -- Decide if there is an finding ----------------------
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
            issueText = "'" & rng.Text & "' -- unable to determine noun/verb context; " & _
                        "review context to ensure correct UK spelling"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present -- ambiguous
            issueText = "'" & rng.Text & "' -- conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf Not usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present -- ambiguous
            issueText = "'" & rng.Text & "' -- conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        End If

        ' Only create finding if we found something to flag
        If Len(issueText) > 0 Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_LICENCE, locStr, issueText, suggestion, rng.Start, rng.End, "possible_error")
            issues.Add finding
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
'  RULE 14 -- CHECK / CHEQUE (UK mode only)
'  "check" as a verb (to verify) is valid UK English.
'  Only the financial-instrument noun should be "cheque" in UK.
'  Detects "check" when used as a noun (not a verb) and suggests
'  "cheque". Verb detection uses preceding word context.
' ================================================================
' ================================================================

Public Function Check_CheckCheque(doc As Document) As Collection
    Dim issues As New Collection
    Dim spellingMode As String
    spellingMode = EngineGetSpellingMode()

    ' Only applies in UK mode (US uses "check" for everything)
    If spellingMode <> "UK" Then
        Set Check_CheckCheque = issues
        Exit Function
    End If

    ' Search body text for "check" / "checks" (context-aware)
    SearchCheckCheque doc.Content, doc, issues

    ' Search body text for financial compound phrases
    SearchFinancialCheckCompounds doc.Content, doc, issues

    ' Search footnotes
    On Error Resume Next
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        Err.Clear
        SearchCheckCheque fn.Range, doc, issues
        If Err.Number <> 0 Then Err.Clear
        Err.Clear
        SearchFinancialCheckCompounds fn.Range, doc, issues
        If Err.Number <> 0 Then Err.Clear
    Next fn
    On Error GoTo 0

    Set Check_CheckCheque = issues
End Function

Private Sub SearchCheckCheque(searchRange As Range, doc As Document, _
                               ByRef issues As Collection)
    Dim rng As Range
    Dim foundText As String
    Dim finding As Object
    Dim locStr As String

    ' Search for "check" as whole word
    Dim searchTerms As Variant
    searchTerms = Array("check", "checks")

    Dim si As Long
    For si = LBound(searchTerms) To UBound(searchTerms)
        On Error Resume Next
        Set rng = searchRange.Duplicate
        If Err.Number <> 0 Then Err.Clear: GoTo NextSearchTerm
        On Error GoTo 0

        With rng.Find
            .ClearFormatting
            .Text = CStr(searchTerms(si))
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
            Dim foundIt As Boolean
            foundIt = rng.Find.Execute
            If Err.Number <> 0 Then Err.Clear: Exit Do
            On Error GoTo 0

            If Not foundIt Then Exit Do
            If rng.Start <= lastPos Then Exit Do
            lastPos = rng.Start

            If Not EngineIsInPageRange(rng) Then
                rng.Collapse wdCollapseEnd
                GoTo NextCheckMatch
            End If

            foundText = rng.Text

            ' Determine if this is a verb usage (skip) or noun (flag)
            If IsCheckUsedAsVerb(rng, doc) Then
                rng.Collapse wdCollapseEnd
                GoTo NextCheckMatch
            End If

            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Dim suggestion As String
            If LCase(foundText) = "checks" Then
                suggestion = "cheques"
            Else
                suggestion = "cheque"
            End If

            Set finding = CreateIssueDict(RULE_NAME_CHECK, locStr, _
                "UK spelling: '" & foundText & "' appears to be a noun (financial instrument). Use '" & suggestion & "' in UK English.", _
                suggestion, rng.Start, rng.End, "possible_error")
            issues.Add finding

NextCheckMatch:
            On Error Resume Next
            rng.Collapse wdCollapseEnd
            If Err.Number <> 0 Then Err.Clear: Exit Do
            On Error GoTo 0
        Loop
NextSearchTerm:
    Next si
End Sub

' Determine if "check" is used as a verb by looking at surrounding context.
' Returns True if likely a verb (should NOT be flagged).
Private Function IsCheckUsedAsVerb(rng As Range, doc As Document) As Boolean
    IsCheckUsedAsVerb = False

    ' Get up to 30 chars before the word
    Dim lookStart As Long
    lookStart = rng.Start - 30
    If lookStart < 0 Then lookStart = 0
    Dim beforeText As String
    beforeText = ""
    On Error Resume Next
    If rng.Start > lookStart Then
        beforeText = LCase(doc.Range(lookStart, rng.Start).Text)
    End If
    If Err.Number <> 0 Then beforeText = "": Err.Clear
    On Error GoTo 0

    ' Get up to 20 chars after the word
    Dim afterText As String
    afterText = ""
    Dim lookEnd As Long
    lookEnd = rng.End + 20
    On Error Resume Next
    If lookEnd > doc.Content.End Then lookEnd = doc.Content.End
    If lookEnd > rng.End Then
        afterText = LCase(doc.Range(rng.End, lookEnd).Text)
    End If
    If Err.Number <> 0 Then afterText = "": Err.Clear
    On Error GoTo 0

    ' Extract last word before "check"
    beforeText = Trim(beforeText)
    Dim lastWord As String
    Dim sp As Long
    sp = InStrRev(beforeText, " ")
    If sp > 0 Then
        lastWord = Mid$(beforeText, sp + 1)
    Else
        lastWord = beforeText
    End If

    ' --- Compound prefix check (before "check") ---
    ' Words like "double-check", "cross-check", "spot-check" etc.
    ' are NOT financial. Also handles "double check" (space-separated)
    ' where lastWord = "double" via the space split above.
    Dim compoundPrefixes As Variant
    Dim cp1 As Variant, cp2 As Variant, cp3 As Variant
    cp1 = Array("double", "triple", "quadruple", "spot", _
        "fact", "reality", "re", "counter", "body", "rain")
    cp2 = Array("sound", "spell", "health", "quality", "background", _
        "reference", "security", "safety", "compliance", "system")
    cp3 = Array("gut", "sense", "temperature", "sanity", "mic", _
        "mike", "status", "progress", "wellness", "vibe", _
        "stock", "price", "proof", "ground", "over", _
        "under", "un", "pre", "self")
    Dim vi As Long

    ' Check if lastWord contains a hyphen (e.g. "double-")
    Dim compPrefix As String
    If InStr(lastWord, "-") > 0 Then
        compPrefix = Left$(lastWord, InStr(lastWord, "-") - 1)
    Else
        compPrefix = lastWord
    End If

    ' Check against all compound prefix arrays
    For vi = LBound(cp1) To UBound(cp1)
        If compPrefix = CStr(cp1(vi)) Then
            IsCheckUsedAsVerb = True  ' Not financial
            Exit Function
        End If
    Next vi
    For vi = LBound(cp2) To UBound(cp2)
        If compPrefix = CStr(cp2(vi)) Then
            IsCheckUsedAsVerb = True
            Exit Function
        End If
    Next vi
    For vi = LBound(cp3) To UBound(cp3)
        If compPrefix = CStr(cp3(vi)) Then
            IsCheckUsedAsVerb = True
            Exit Function
        End If
    Next vi

    ' --- Compound suffix check (after "check") ---
    ' Words like "check-in", "check-out", "check-mate" etc.
    ' are NOT financial. "check-book" IS financial (excluded).
    Dim firstCharAfter As String
    afterText = Trim(afterText)
    firstCharAfter = ""
    If Len(afterText) > 0 Then firstCharAfter = Left$(afterText, 1)

    If firstCharAfter = "-" And Len(afterText) > 1 Then
        ' Extract the word after the hyphen
        Dim suffixWord As String
        Dim restAfter As String
        restAfter = Mid$(afterText, 2)
        sp = InStr(1, restAfter, " ")
        If sp > 0 Then
            suffixWord = Left$(restAfter, sp - 1)
        Else
            suffixWord = restAfter
        End If
        suffixWord = LCase(suffixWord)

        Dim nfSuffix1 As Variant, nfSuffix2 As Variant
        nfSuffix1 = Array("in", "out", "up", "list", "mark", "mate", _
            "point", "sum", "box", "off", "room", "er", "ers", "ed", "ing")
        nfSuffix2 = Array("able", "board", "down", "through", "ride", _
            "rein", "bone", "flag", "gate", "land", "line", _
            "pattern", "piece", "rail", "row", "side", "weight", "work")
        ' Note: "book" deliberately absent — check-book IS financial

        For vi = LBound(nfSuffix1) To UBound(nfSuffix1)
            If suffixWord = CStr(nfSuffix1(vi)) Then
                IsCheckUsedAsVerb = True
                Exit Function
            End If
        Next vi
        For vi = LBound(nfSuffix2) To UBound(nfSuffix2)
            If suffixWord = CStr(nfSuffix2(vi)) Then
                IsCheckUsedAsVerb = True
                Exit Function
            End If
        Next vi
    End If

    ' --- Standard verb/noun context analysis ---
    ' Verb indicators: preceded by modal verbs, auxiliaries, etc.
    Dim verbPrecedes As Variant
    verbPrecedes = Array("to", "will", "shall", "must", "should", _
                         "would", "could", "can", "may", "might", _
                         "please", "let", "did", "does", "do", _
                         "not", "always", "also", "then", "and", _
                         "or", "we", "they", "you", "i")
    For vi = LBound(verbPrecedes) To UBound(verbPrecedes)
        If lastWord = CStr(verbPrecedes(vi)) Then
            IsCheckUsedAsVerb = True
            Exit Function
        End If
    Next vi

    ' Verb indicator: followed by certain words
    Dim firstWordAfter As String
    sp = InStr(1, afterText, " ")
    If sp > 0 Then
        firstWordAfter = Left$(afterText, sp - 1)
    Else
        firstWordAfter = afterText
    End If
    ' Strip leading hyphen if present (already handled above for compounds)
    If Left$(firstWordAfter, 1) = "-" Then firstWordAfter = Mid$(firstWordAfter, 2)

    Dim verbFollows As Variant
    verbFollows = Array("that", "whether", "if", "the", "this", _
                        "for", "with", "on", "your", "our", _
                        "his", "her", "its", "their", "my", _
                        "each", "every", "all", "any")
    For vi = LBound(verbFollows) To UBound(verbFollows)
        If firstWordAfter = CStr(verbFollows(vi)) Then
            IsCheckUsedAsVerb = True
            Exit Function
        End If
    Next vi

    ' Noun indicators: preceded by determiners/prepositions
    Dim nounPrecedes As Variant
    nounPrecedes = Array("a", "the", "this", "that", "each", _
                         "every", "your", "our", "his", "her", _
                         "my", "its", "their", "by", "per", _
                         "no", "any", "one", "blank")
    For vi = LBound(nounPrecedes) To UBound(nounPrecedes)
        If lastWord = CStr(nounPrecedes(vi)) Then
            IsCheckUsedAsVerb = False
            Exit Function
        End If
    Next vi

    ' Default: treat as possible noun (flag it as possible_error for review)
    IsCheckUsedAsVerb = False
End Function

' ----------------------------------------------------------------
' Search for financial compound words/phrases containing "check"
' that should use "cheque" in UK English. These are searched as
' literal phrases and flagged unconditionally (no verb/noun analysis).
' ----------------------------------------------------------------
Private Sub SearchFinancialCheckCompounds(searchRange As Range, _
                                          doc As Document, _
                                          ByRef issues As Collection)
    ' Parallel arrays: search terms and their UK suggestions
    ' Split into batches to stay under 25 line-continuation limit
    Dim terms1 As Variant, sugs1 As Variant
    terms1 = Array("checkbook", "check-book", "checkbooks", "check-books", _
        "paycheck", "pay-check", "paychecks", "pay-checks")
    sugs1 = Array("chequebook", "cheque-book", "chequebooks", "cheque-books", _
        "pay cheque", "pay cheque", "pay cheques", "pay cheques")

    Dim terms2 As Variant, sugs2 As Variant
    terms2 = Array("blank check", "blank checks", "bad check", "bad checks", _
        "bounced check", "bounced checks", "rubber check", "rubber checks")
    sugs2 = Array("blank cheque", "blank cheques", "bad cheque", "bad cheques", _
        "bounced cheque", "bounced cheques", "rubber cheque", "rubber cheques")

    Dim terms3 As Variant, sugs3 As Variant
    terms3 = Array("cancelled check", "canceled check", _
        "certified check", "certified checks", _
        "cashier's check", "cashiers check")
    sugs3 = Array("cancelled cheque", "cancelled cheque", _
        "certified cheque", "certified cheques", _
        "cashier's cheque", "cashier's cheque")

    Dim terms4 As Variant, sugs4 As Variant
    terms4 = Array("traveller's check", "traveler's check", _
        "travellers check", "travelers check", _
        "traveller's checks", "traveler's checks")
    sugs4 = Array("traveller's cheque", "traveller's cheque", _
        "travellers' cheque", "travellers' cheque", _
        "traveller's cheques", "traveller's cheques")

    Dim terms5 As Variant, sugs5 As Variant
    terms5 = Array("travellers checks", "travelers checks", _
        "personal check", "personal checks", _
        "bank check", "bank checks")
    sugs5 = Array("travellers' cheques", "travellers' cheques", _
        "personal cheque", "personal cheques", _
        "bank cheque", "bank cheques")

    Dim terms6 As Variant, sugs6 As Variant
    terms6 = Array("post-dated check", "postdated check", _
        "stale check", "stale checks", _
        "dishonoured check", "dishonored check")
    sugs6 = Array("post-dated cheque", "post-dated cheque", _
        "stale cheque", "stale cheques", _
        "dishonoured cheque", "dishonoured cheque")

    Dim terms7 As Variant, sugs7 As Variant
    terms7 = Array("check stub", "check stubs", "check fraud", _
        "check forgery", "check clearing", _
        "check guarantee", "check number", "check numbers")
    sugs7 = Array("cheque stub", "cheque stubs", "cheque fraud", _
        "cheque forgery", "cheque clearing", _
        "cheque guarantee", "cheque number", "cheque numbers")

    ' Process each batch
    SearchFinancialBatch searchRange, doc, issues, terms1, sugs1, True
    SearchFinancialBatch searchRange, doc, issues, terms2, sugs2, False
    SearchFinancialBatch searchRange, doc, issues, terms3, sugs3, False
    SearchFinancialBatch searchRange, doc, issues, terms4, sugs4, False
    SearchFinancialBatch searchRange, doc, issues, terms5, sugs5, False
    SearchFinancialBatch searchRange, doc, issues, terms6, sugs6, False
    SearchFinancialBatch searchRange, doc, issues, terms7, sugs7, False
End Sub

Private Sub SearchFinancialBatch(searchRange As Range, _
                                  doc As Document, _
                                  ByRef issues As Collection, _
                                  terms As Variant, _
                                  suggestions As Variant, _
                                  wholeWord As Boolean)
    Dim ti As Long
    Dim rng As Range
    Dim finding As Object
    Dim locStr As String

    For ti = LBound(terms) To UBound(terms)
        On Error Resume Next
        Set rng = searchRange.Duplicate
        If Err.Number <> 0 Then Err.Clear: GoTo NextFinTerm
        On Error GoTo 0

        With rng.Find
            .ClearFormatting
            .Text = CStr(terms(ti))
            .MatchWholeWord = wholeWord
            .MatchCase = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With

        Dim lastPos As Long
        lastPos = -1
        Do
            On Error Resume Next
            Dim foundIt As Boolean
            foundIt = rng.Find.Execute
            If Err.Number <> 0 Then Err.Clear: Exit Do
            On Error GoTo 0

            If Not foundIt Then Exit Do
            If rng.Start <= lastPos Then Exit Do
            lastPos = rng.Start

            If Not EngineIsInPageRange(rng) Then
                rng.Collapse wdCollapseEnd
                GoTo NextFinMatch
            End If

            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_CHECK, locStr, _
                "UK spelling: '" & rng.Text & "' should be '" & _
                CStr(suggestions(ti)) & "' in UK English.", _
                "Use '" & CStr(suggestions(ti)) & "'", rng.Start, rng.End, _
                "possible_error", True, CStr(suggestions(ti)))
            issues.Add finding

NextFinMatch:
            On Error Resume Next
            rng.Collapse wdCollapseEnd
            If Err.Number <> 0 Then Err.Clear: Exit Do
            On Error GoTo 0
        Loop
NextFinTerm:
    Next ti
End Sub

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
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraColor As Long
    Dim colourCounts As Object
    Dim dominantColour As Long
    Dim maxCount As Long

    Const WD_COLOR_AUTOMATIC As Long = -16777216

    ' -- Build hyperlink position set once (avoid O(n^2)) ------
    Dim hlStarts As Object, hlEnds As Object
    Set hlStarts = CreateObject("Scripting.Dictionary")
    Set hlEnds = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim hl As Hyperlink
    Dim hlIdx As Long: hlIdx = 0
    For Each hl In doc.Hyperlinks
        Err.Clear
        hlStarts.Add hlIdx, hl.Range.Start
        hlEnds.Add hlIdx, hl.Range.End
        If Err.Number <> 0 Then Err.Clear
        hlIdx = hlIdx + 1
    Next hl
    On Error GoTo 0

    ' -- Pass 1: count paragraph-level colours -----------------
    Set colourCounts = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextPC1

        If Not EngineIsInPageRange(paraRange) Then GoTo NextPC1

        paraColor = paraRange.Font.Color
        If Err.Number <> 0 Then Err.Clear: GoTo NextPC1

        ' Skip indeterminate (mixed-colour paragraphs counted in pass 2)
        If paraColor = 9999999 Then GoTo NextPC1

        If colourCounts.Exists(paraColor) Then
            colourCounts(paraColor) = colourCounts(paraColor) + 1
        Else
            colourCounts.Add paraColor, 1
        End If
NextPC1:
    Next para
    On Error GoTo 0

    ' -- Determine dominant colour -----------------------------
    If colourCounts.Count = 0 Then
        Set Check_ColourFormatting = issues
        Exit Function
    End If

    dominantColour = 0: maxCount = 0
    Dim colourKey As Variant
    For Each colourKey In colourCounts.keys
        If colourCounts(colourKey) > maxCount Then
            maxCount = colourCounts(colourKey)
            dominantColour = CLng(colourKey)
        End If
    Next colourKey

    ' -- Pass 2: flag paragraphs with non-standard colours -----
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextPC2

        If Not EngineIsInPageRange(paraRange) Then GoTo NextPC2

        ' Skip heading-styled paragraphs
        Dim styleName As String
        styleName = ""
        styleName = para.Style.NameLocal
        If Err.Number <> 0 Then Err.Clear: styleName = ""
        If LCase(Left(styleName, 7)) = "heading" Then GoTo NextPC2

        paraColor = paraRange.Font.Color
        If Err.Number <> 0 Then Err.Clear: GoTo NextPC2

        ' Skip dominant, automatic, or indeterminate
        If paraColor = dominantColour Or _
           paraColor = WD_COLOR_AUTOMATIC Or _
           paraColor = 9999999 Then GoTo NextPC2

        ' Skip if inside a hyperlink
        If IsRangeInsideHyperlink(paraRange, hlStarts, hlEnds) Then GoTo NextPC2

        ' Flag this paragraph
        FlushColourGroup doc, issues, paraRange.Start, paraRange.End, paraColor

NextPC2:
    Next para
    On Error GoTo 0

    Set Check_ColourFormatting = issues
End Function

' ============================================================
'  PRIVATE: Flush a grouped colour finding
' ============================================================
Private Sub FlushColourGroup(doc As Document, _
                              ByRef issues As Collection, _
                              ByVal startPos As Long, _
                              ByVal endPos As Long, _
                              ByVal fontColor As Long)
    Dim finding As Object
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

    Set finding = CreateIssueDict(RULE_NAME_COLOUR, locStr, "Non-standard font colour " & hexStr & " detected: '" & previewText & "'", "Change font colour to match document default", startPos, endPos, "possible_error")
    issues.Add finding
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
Private Function IsRangeInsideHyperlink(rng As Range, _
                                        hlStarts As Object, _
                                        hlEnds As Object) As Boolean
    Dim i As Long
    For i = 0 To hlStarts.Count - 1
        If hlStarts(i) <= rng.Start And hlEnds(i) >= rng.End Then
            IsRangeInsideHyperlink = True
            Exit Function
        End If
    Next i
    IsRangeInsideHyperlink = False
End Function


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
                                 Optional ByVal autoFixSafe_ As Boolean = False, _
                                 Optional ByVal replacementText_ As String = "") As Object
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
    d("ReplacementText") = replacementText_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
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
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
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
        Debug.Print "EngineIsWhitelistedTerm: fallback (Err " & Err.Number & ": " & Err.Description & ")"
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
        Debug.Print "EngineGetSpellingMode: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetSpellingMode = "UK"
        Err.Clear
    End If
    On Error GoTo 0
End Function
