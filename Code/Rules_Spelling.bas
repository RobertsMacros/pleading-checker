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
' Dependencies:
'   - TextAnchoring.bas (AddIssue, SafeRange, FindAll, IsWhitespaceChar,
'                        IsInPageRange, IsPastPageFilter, IsWhitelistedTerm,
'                        GetSpellingMode, GetListPrefixLen,
'                        PerfTimerStart/End, PerfCount)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "spellchecker"
Private Const RULE_NAME_LICENCE As String = "spellchecker"
Private Const RULE_NAME_CHECK As String = "spellchecker"

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

    ' -- Build the US <-> UK mapping arrays (cached once per call) --
    BuildSpellingArrays usWords, ukWords

    ' -- Determine spelling mode -------------------------
    spellingMode = TextAnchoring.GetSpellingMode()

    If spellingMode = "US" Then
        searchWords = ukWords
        targetWords = usWords
        direction = "US"
        exceptions = Split("program,practice", ",")
    Else
        searchWords = usWords
        targetWords = ukWords
        direction = "UK"
        exceptions = Split("program,judgment,practice", ",")
    End If

    ' -- Search main document body -----------------------
    TextAnchoring.PerfTimerStart "spelling_body"
    SearchRangeForSpellingIssues doc.Content, doc, searchWords, targetWords, exceptions, direction, issues
    TextAnchoring.PerfTimerEnd "spelling_body"

    ' -- Search footnotes via story range (single pass, not per-footnote) --
    TextAnchoring.PerfTimerStart "spelling_footnotes"
    On Error Resume Next
    If doc.Footnotes.Count > 0 Then
        Dim fnStory As Range
        Set fnStory = doc.StoryRanges(wdFootnotesStory)
        If Err.Number = 0 And Not fnStory Is Nothing Then
            Err.Clear
            SearchRangeForSpellingIssues fnStory, doc, searchWords, targetWords, exceptions, direction, issues
            If Err.Number <> 0 Then Err.Clear
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0
    TextAnchoring.PerfTimerEnd "spelling_footnotes"

    ' -- Search endnotes via story range (single pass) --
    TextAnchoring.PerfTimerStart "spelling_endnotes"
    On Error Resume Next
    If doc.Endnotes.Count > 0 Then
        Dim enStory As Range
        Set enStory = doc.StoryRanges(wdEndnotesStory)
        If Err.Number = 0 And Not enStory Is Nothing Then
            Err.Clear
            SearchRangeForSpellingIssues enStory, doc, searchWords, targetWords, exceptions, direction, issues
            If Err.Number <> 0 Then Err.Clear
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0
    TextAnchoring.PerfTimerEnd "spelling_endnotes"

    TextAnchoring.PerfCount "spelling_find_passes", CLng(UBound(searchWords) - LBound(searchWords) + 1)

    Set Check_Spelling = issues
End Function

' ============================================================
'  PRIVATE: Search a Range for spelling issues using a
'  single-pass paragraph scan with dictionary lookup.
'  Replaces the old O(N x pairs) Range.Find approach with
'  O(N) tokenisation where N = total document text length.
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
    Dim issueText As String
    Dim sourceLabel As String

    ' Determine the label for the detected spelling variant
    If direction = "UK" Then
        sourceLabel = "US"
    Else
        sourceLabel = "UK"
    End If

    ' -- Build dictionary: LCase(searchWord) -> index into arrays --
    Dim lookupDict As Object
    Set lookupDict = CreateObject("Scripting.Dictionary")
    For i = LBound(searchWords) To UBound(searchWords)
        Dim lcWord As String
        lcWord = LCase$(searchWords(i))
        If Not lookupDict.Exists(lcWord) Then
            lookupDict.Add lcWord, i
        End If
    Next i

    ' -- Build exceptions dictionary for O(1) lookup --
    Dim exceptDict As Object
    Set exceptDict = CreateObject("Scripting.Dictionary")
    For i = LBound(exceptions) To UBound(exceptions)
        Dim lcExc As String
        lcExc = LCase$(Trim$(exceptions(i)))
        If Len(lcExc) > 0 And Not exceptDict.Exists(lcExc) Then
            exceptDict.Add lcExc, True
        End If
    Next i

    ' -- Iterate paragraphs in the search range --
    Dim para As Paragraph
    Dim paraText As String
    Dim paraStart As Long
    Dim tLen As Long
    Dim scanPos As Long
    Dim tokStart As Long
    Dim sc As String
    Dim rawToken As String
    Dim cleanToken As String
    Dim matchIdx As Long

    On Error Resume Next
    For Each para In searchRange.Paragraphs
        Err.Clear
        Dim paraRange As Range
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextSpellPara

        paraStart = paraRange.Start

        ' Page-range filter
        If TextAnchoring.IsPastPageFilter(paraStart) Then Exit For
        If Not TextAnchoring.IsInPageRange(paraRange) Then GoTo NextSpellPara

        paraText = paraRange.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NextSpellPara
        tLen = Len(paraText)
        If tLen < 2 Then GoTo NextSpellPara

        ' Calculate list prefix offset
        Dim spListPrefixLen As Long
        spListPrefixLen = TextAnchoring.GetListPrefixLen(para, paraText)

        ' -- Tokenise by scanning character positions --
        scanPos = 1
        Do While scanPos <= tLen
            sc = Mid$(paraText, scanPos, 1)
            ' Skip non-word characters (whitespace, punctuation)
            If Not IsWordCharSpelling(sc) Then
                scanPos = scanPos + 1
            Else
                ' Found start of a token
                tokStart = scanPos
                Do While scanPos <= tLen
                    sc = Mid$(paraText, scanPos, 1)
                    If Not IsWordCharSpelling(sc) Then Exit Do
                    scanPos = scanPos + 1
                Loop
                ' Extract token
                rawToken = Mid$(paraText, tokStart, scanPos - tokStart)
                cleanToken = LCase$(rawToken)

                ' -- Look up in dictionary --
                If lookupDict.Exists(cleanToken) Then
                    matchIdx = CLng(lookupDict(cleanToken))

                    ' -- Skip exceptions --
                    If exceptDict.Exists(cleanToken) Then GoTo NextSpellToken

                    ' -- Skip whitelisted terms --
                    If TextAnchoring.IsWhitelistedTerm(rawToken) Then GoTo NextSpellToken

                    ' -- Compute document position --
                    Dim spRangeStart As Long, spRangeEnd As Long
                    spRangeStart = paraStart + (tokStart - 1) - spListPrefixLen
                    spRangeEnd = spRangeStart + Len(rawToken)

                    ' Create a range for the match
                    Dim matchRng As Range
                    Set matchRng = TextAnchoring.SafeRange(doc, spRangeStart, spRangeEnd)
                    If matchRng Is Nothing Then GoTo NextSpellToken

                    ' Verify the document text matches (guard against list prefix offset issues)
                    Dim actualText As String
                    actualText = matchRng.Text
                    If Err.Number <> 0 Then Err.Clear: GoTo NextSpellToken
                    If LCase$(actualText) <> cleanToken Then GoTo NextSpellToken

                    ' -- Create the finding --
                    Dim foundText As String
                    foundText = actualText

                    issueText = sourceLabel & " spelling detected: '" & foundText & "'"

                    ' -- Downgrade italic / quoted text --
                    Dim severity As String
                    Dim suggestion As String
                    severity = "error"
                    suggestion = targetWords(matchIdx)

                    If IsRangeItalic(matchRng) Then
                        severity = "possible_error"
                        suggestion = ""
                        issueText = issueText & " (in italic text -- review manually)"
                    ElseIf IsInsideQuotes(matchRng, doc) Then
                        severity = "possible_error"
                        suggestion = ""
                        issueText = issueText & " (in quoted text -- review manually)"
                    End If

                    ' Mark as auto-fix safe when severity is "error"
                    Dim spAutoFix As Boolean
                    Dim spReplacement As String
                    spAutoFix = False
                    spReplacement = ""
                    If severity = "error" And Len(suggestion) > 0 Then
                        spAutoFix = True
                        spReplacement = suggestion
                    End If

                    TextAnchoring.AddIssue issues, RULE_NAME, doc, matchRng, issueText, suggestion, _
                        spRangeStart, spRangeEnd, severity, spAutoFix, spReplacement, _
                        foundText, "exact_text", "high"
                End If
NextSpellToken:
            End If
        Loop
NextSpellPara:
    Next para
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Check if a character is a word character for
'  spelling tokenisation (letter, digit, apostrophe, hyphen)
' ============================================================
Private Function IsWordCharSpelling(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsWordCharSpelling = False
        Exit Function
    End If
    Dim c As Long
    c = AscW(ch)
    ' A-Z, a-z
    If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
        IsWordCharSpelling = True
        Exit Function
    End If
    ' 0-9
    If c >= 48 And c <= 57 Then
        IsWordCharSpelling = True
        Exit Function
    End If
    ' Apostrophe (for contractions like "don't" - these won't match spelling words)
    ' Hyphen (for compound words)
    If c = 39 Or c = 45 Then
        IsWordCharSpelling = True
        Exit Function
    End If
    ' Smart apostrophes
    If c = 8217 Or c = 8216 Then
        IsWordCharSpelling = True
        Exit Function
    End If
    IsWordCharSpelling = False
End Function

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

    ' Get character before range
    charBefore = ""
    If rng.Start > 0 Then
        Dim rngBefore As Range
        Set rngBefore = TextAnchoring.SafeRange(doc, rng.Start - 1, rng.Start)
        If Not rngBefore Is Nothing Then
            On Error Resume Next
            charBefore = rngBefore.Text
            If Err.Number <> 0 Then charBefore = "": Err.Clear
            On Error GoTo 0
        End If
    End If

    ' Get character after range
    charAfter = ""
    On Error Resume Next
    Dim docEnd As Long
    docEnd = doc.Content.End
    If Err.Number <> 0 Then docEnd = 0: Err.Clear
    On Error GoTo 0
    If rng.End < docEnd Then
        Dim rngAfter As Range
        Set rngAfter = TextAnchoring.SafeRange(doc, rng.End, rng.End + 1)
        If Not rngAfter Is Nothing Then
            On Error Resume Next
            charAfter = rngAfter.Text
            If Err.Number <> 0 Then charAfter = "": Err.Clear
            On Error GoTo 0
        End If
    End If

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

    Dim lookbackRng As Range
    Set lookbackRng = TextAnchoring.SafeRange(doc, lookbackStart, rng.Start)
    If lookbackRng Is Nothing Then
        IsInsideQuotes = False
        Exit Function
    End If
    Dim beforeText As String
    On Error Resume Next
    beforeText = lookbackRng.Text
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

    ' Search body
    SearchForLicenceIssues doc.Content, doc, issues

    ' Search footnotes via story range
    On Error Resume Next
    If doc.Footnotes.Count > 0 Then
        Dim fnStory As Range
        Set fnStory = doc.StoryRanges(wdFootnotesStory)
        If Err.Number = 0 And Not fnStory Is Nothing Then
            Err.Clear
            SearchForLicenceIssues fnStory, doc, issues
            If Err.Number <> 0 Then Err.Clear
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0

    ' Search endnotes via story range
    On Error Resume Next
    If doc.Endnotes.Count > 0 Then
        Dim enStory As Range
        Set enStory = doc.StoryRanges(wdEndnotesStory)
        If Err.Number = 0 And Not enStory Is Nothing Then
            Err.Clear
            SearchForLicenceIssues enStory, doc, issues
            If Err.Number <> 0 Then Err.Clear
        Else
            Err.Clear
        End If
    End If
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
'  analyse context.  Uses TextAnchoring.FindAll to replace
'  the manual Find loop.
' ============================================================
Private Sub SearchSingleLicenceTerm(ByVal term As String, _
                              searchRange As Range, _
                              doc As Document, _
                              ByRef issues As Collection)
    Dim matches As Collection
    Set matches = TextAnchoring.FindAll(doc, term, True, False, False, searchRange)

    Dim contextBefore As String
    Dim contextAfter As String
    Dim wordBefore As String
    Dim wordAfter As String
    Dim issueText As String
    Dim suggestion As String
    Dim usesS As Boolean
    Dim baseIsNoun As Boolean
    Dim baseIsVerb As Boolean
    Dim matchArr As Variant
    Dim startPos As Long, endPos As Long, matchText As String
    Dim m As Long

    For m = 1 To matches.Count
        matchArr = matches(m)
        startPos = matchArr(0)
        endPos = matchArr(1)
        matchText = matchArr(2)

        ' Determine if the found word uses -s- or -c-
        usesS = (InStr(1, LCase(matchText), "license") > 0)

        ' Skip "licensed" and "licensing" -- always correct with -s-
        Dim foundLower As String
        foundLower = LCase(Trim(matchText))
        If foundLower = "licensed" Or foundLower = "licensing" Then
            GoTo ContinueLicenceSearch
        End If

        ' Create a range for context and italic/quote checks
        Dim rng As Range
        Set rng = TextAnchoring.SafeRange(doc, startPos, endPos)
        If rng Is Nothing Then GoTo ContinueLicenceSearch

        ' -- Downgrade italic / quoted text ------------------
        If IsRangeItalic(rng) Then
            TextAnchoring.AddIssue issues, RULE_NAME_LICENCE, doc, rng, _
                "'" & matchText & "' -- in italic text, review manually", "", _
                startPos, endPos, "possible_error"
            GoTo ContinueLicenceSearch
        End If

        If IsInsideQuotes(rng, doc) Then
            TextAnchoring.AddIssue issues, RULE_NAME_LICENCE, doc, rng, _
                "'" & matchText & "' -- in quoted text, review manually", "", _
                startPos, endPos, "possible_error"
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
            issueText = "'" & matchText & "' appears in a noun context; " & _
                        "UK convention uses 'licence' for the noun"
            suggestion = ReplaceSWithC(matchText)
        ElseIf Not usesS And baseIsVerb And Not baseIsNoun Then
            ' "licence" used in verb context -- should be "license"
            issueText = "'" & matchText & "' appears in a verb context; " & _
                        "UK convention uses 'license' for the verb"
            suggestion = ReplaceCWithS(matchText)
        ElseIf (usesS And Not baseIsVerb And Not baseIsNoun) Or _
               (Not usesS And Not baseIsVerb And Not baseIsNoun) Then
            ' Context ambiguous
            issueText = "'" & matchText & "' -- unable to determine noun/verb context; " & _
                        "review context to ensure correct UK spelling"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present -- ambiguous
            issueText = "'" & matchText & "' -- conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf Not usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present -- ambiguous
            issueText = "'" & matchText & "' -- conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        End If

        ' Only create finding if we found something to flag
        If Len(issueText) > 0 Then
            TextAnchoring.AddIssue issues, RULE_NAME_LICENCE, doc, rng, _
                issueText, suggestion, startPos, endPos, "possible_error"
        End If

ContinueLicenceSearch:
    Next m
End Sub

' ============================================================
'  PRIVATE: Get text before the match range (up to N chars)
' ============================================================
Private Function GetLicenceContextBefore(rng As Range, doc As Document, _
                                   ByVal charCount As Long) As String
    Dim startPos As Long
    startPos = rng.Start - charCount
    If startPos < 0 Then startPos = 0

    Dim contextRng As Range
    Set contextRng = TextAnchoring.SafeRange(doc, startPos, rng.Start)
    If contextRng Is Nothing Then
        GetLicenceContextBefore = ""
        Exit Function
    End If

    On Error Resume Next
    GetLicenceContextBefore = contextRng.Text
    If Err.Number <> 0 Then
        GetLicenceContextBefore = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ============================================================
'  PRIVATE: Get text after the match range (up to N chars)
' ============================================================
Private Function GetLicenceContextAfter(rng As Range, doc As Document, _
                                  ByVal charCount As Long) As String
    Dim endPos As Long
    Dim docEnd As Long

    On Error Resume Next
    docEnd = doc.Content.End
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        GetLicenceContextAfter = ""
        Exit Function
    End If
    On Error GoTo 0

    endPos = rng.End + charCount
    If endPos > docEnd Then endPos = docEnd

    Dim contextRng As Range
    Set contextRng = TextAnchoring.SafeRange(doc, rng.End, endPos)
    If contextRng Is Nothing Then
        GetLicenceContextAfter = ""
        Exit Function
    End If

    On Error Resume Next
    GetLicenceContextAfter = contextRng.Text
    If Err.Number <> 0 Then
        GetLicenceContextAfter = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ============================================================
'  PRIVATE: Extract the last word from a context string
' ============================================================
Private Function GetLastWordFromContext(ByVal contextStr As String) As String
    Dim trimmed As String
    Dim i As Long
    Dim ch As String

    trimmed = Trim(contextStr)
    If Len(trimmed) = 0 Then
        GetLastWordFromContext = ""
        Exit Function
    End If

    ' Walk backward from end to find last word boundary
    For i = Len(trimmed) To 1 Step -1
        ch = Mid(trimmed, i, 1)
        If TextAnchoring.IsWhitespaceChar(ch) Then
            GetLastWordFromContext = LCase(Mid(trimmed, i + 1))
            Exit Function
        End If
    Next i

    GetLastWordFromContext = LCase(trimmed)
End Function

' ============================================================
'  PRIVATE: Extract the first word from a context string
' ============================================================
Private Function GetFirstWordFromContext(ByVal contextStr As String) As String
    Dim trimmed As String
    Dim spacePos As Long

    trimmed = Trim(contextStr)
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
    spellingMode = TextAnchoring.GetSpellingMode()

    ' Only applies in UK mode (US uses "check" for everything)
    If spellingMode <> "UK" Then
        Set Check_CheckCheque = issues
        Exit Function
    End If

    ' Search body text
    SearchCheckCheque doc.Content, doc, issues
    SearchFinancialCheckCompounds doc.Content, doc, issues

    ' Search footnotes via story range
    On Error Resume Next
    If doc.Footnotes.Count > 0 Then
        Dim fnStory As Range
        Set fnStory = doc.StoryRanges(wdFootnotesStory)
        If Err.Number = 0 And Not fnStory Is Nothing Then
            Err.Clear
            SearchCheckCheque fnStory, doc, issues
            If Err.Number <> 0 Then Err.Clear
            Err.Clear
            SearchFinancialCheckCompounds fnStory, doc, issues
            If Err.Number <> 0 Then Err.Clear
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0

    Set Check_CheckCheque = issues
End Function

Private Sub SearchCheckCheque(searchRange As Range, doc As Document, _
                               ByRef issues As Collection)
    ' Search for "check" and "checks" as whole words
    Dim searchTerms As Variant
    searchTerms = Array("check", "checks")

    Dim si As Long
    For si = LBound(searchTerms) To UBound(searchTerms)
        Dim matches As Collection
        Set matches = TextAnchoring.FindAll(doc, CStr(searchTerms(si)), True, False, False, searchRange)

        Dim m As Long
        For m = 1 To matches.Count
            Dim matchArr As Variant
            matchArr = matches(m)
            Dim startPos As Long, endPos As Long, matchText As String
            startPos = matchArr(0)
            endPos = matchArr(1)
            matchText = matchArr(2)

            ' Create a range for verb/noun context analysis
            Dim rng As Range
            Set rng = TextAnchoring.SafeRange(doc, startPos, endPos)
            If rng Is Nothing Then GoTo NextCheckMatch

            ' Determine if this is a verb usage (skip) or noun (flag)
            If IsCheckUsedAsVerb(rng, doc) Then GoTo NextCheckMatch

            Dim suggestion As String
            If LCase(matchText) = "checks" Then
                suggestion = "cheques"
            Else
                suggestion = "cheque"
            End If

            TextAnchoring.AddIssue issues, RULE_NAME_CHECK, doc, rng, _
                "UK spelling: '" & matchText & "' appears to be a noun (financial instrument). Use '" & suggestion & "' in UK English.", _
                suggestion, startPos, endPos, "possible_error"

NextCheckMatch:
        Next m
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
    If rng.Start > lookStart Then
        Dim beforeRng As Range
        Set beforeRng = TextAnchoring.SafeRange(doc, lookStart, rng.Start)
        If Not beforeRng Is Nothing Then
            On Error Resume Next
            beforeText = LCase(beforeRng.Text)
            If Err.Number <> 0 Then beforeText = "": Err.Clear
            On Error GoTo 0
        End If
    End If

    ' Get up to 20 chars after the word
    Dim afterText As String
    afterText = ""
    Dim lookEnd As Long
    lookEnd = rng.End + 20
    On Error Resume Next
    If lookEnd > doc.Content.End Then lookEnd = doc.Content.End
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If lookEnd > rng.End Then
        Dim afterRng As Range
        Set afterRng = TextAnchoring.SafeRange(doc, rng.End, lookEnd)
        If Not afterRng Is Nothing Then
            On Error Resume Next
            afterText = LCase(afterRng.Text)
            If Err.Number <> 0 Then afterText = "": Err.Clear
            On Error GoTo 0
        End If
    End If

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

    For ti = LBound(terms) To UBound(terms)
        Dim matches As Collection
        Set matches = TextAnchoring.FindAll(doc, CStr(terms(ti)), wholeWord, False, False, searchRange)

        Dim m As Long
        For m = 1 To matches.Count
            Dim matchArr As Variant
            matchArr = matches(m)
            Dim startPos As Long, endPos As Long, matchText As String
            startPos = matchArr(0)
            endPos = matchArr(1)
            matchText = matchArr(2)

            Dim rng As Range
            Set rng = TextAnchoring.SafeRange(doc, startPos, endPos)
            If rng Is Nothing Then GoTo NextFinMatch

            TextAnchoring.AddIssue issues, RULE_NAME_CHECK, doc, rng, _
                "UK spelling: '" & matchText & "' should be '" & _
                CStr(suggestions(ti)) & "' in UK English.", _
                "Use '" & CStr(suggestions(ti)) & "'", startPos, endPos, _
                "possible_error", True, CStr(suggestions(ti))

NextFinMatch:
        Next m
    Next ti
End Sub



