Attribute VB_Name = "Rules_Spelling"
' ============================================================
' Rules_Spelling.bas
' Combined proofreading rule: detects spelling differences
' between US and UK English, with a configurable direction.
'
' Covers ~95 mappings across categories:
'   -or/-our, -ize/-ise, -ization/-isation, -er/-re,
'   -se/-ce, -og/-ogue, -ment variants, and miscellaneous.
'
' Toggle:
'   Calls PleadingsEngine.GetSpellingMode() to determine
'   direction:
'     "UK" → search for US words, suggest UK replacements
'     "US" → search for UK words, suggest US replacements
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, IsWhitelistedTerm,
'                          GetLocationString, GetSpellingMode)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "spelling"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_Spelling(doc As Document) As Collection
    Dim issues As New Collection
    Dim usWords() As String
    Dim ukWords() As String
    Dim searchWords() As String
    Dim targetWords() As String
    Dim exceptions() As String
    Dim spellingMode As String
    Dim direction As String

    ' ── Build the US ↔ UK mapping arrays ────────────────
    BuildSpellingArrays usWords, ukWords

    ' ── Determine spelling mode ─────────────────────────
    spellingMode = PleadingsEngine.GetSpellingMode()

    If spellingMode = "US" Then
        ' Search for UK words, suggest US replacements
        searchWords = ukWords
        targetWords = usWords
        direction = "US"

        ' In US mode, no special legal exceptions
        exceptions = Split("program,practice", ",")
    Else
        ' Default: "UK" — search for US words, suggest UK replacements
        searchWords = usWords
        targetWords = ukWords
        direction = "UK"

        ' "judgment" is standard in UK legal writing (not "judgement")
        ' "practice" is the correct UK noun form (verb: "practise")
        exceptions = Split("program,judgment,practice", ",")
    End If

    ' ── Search main document body ───────────────────────
    SearchRangeForSpellingIssues doc.Content, doc, searchWords, targetWords, exceptions, direction, issues

    ' ── Search footnotes ────────────────────────────────
    On Error Resume Next
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        Err.Clear
        SearchRangeForSpellingIssues fn.Range, doc, searchWords, targetWords, exceptions, direction, issues
        If Err.Number <> 0 Then Err.Clear
    Next fn
    On Error GoTo 0

    ' ── Search endnotes ─────────────────────────────────
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Search a Range for spelling issues
'  Iterates every search/target pair, uses Word's Find to
'  locate whole-word, case-insensitive matches, then filters
'  by page range and whitelist before creating issues.
'
'  direction = "UK" or "US" — controls the issue text:
'    "UK" → "US spelling detected: '...'"
'    "US" → "UK spelling detected: '...'"
' ════════════════════════════════════════════════════════════
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
    Dim issue As PleadingsIssue
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

            ' ── Skip exceptions ───────────────────────
            If IsException(foundText, exceptions) Then
                GoTo ContinueSearch
            End If

            ' ── Skip whitelisted terms ────────────────
            If PleadingsEngine.IsWhitelistedTerm(foundText) Then
                GoTo ContinueSearch
            End If

            ' ── Skip if outside configured page range ─
            If Not PleadingsEngine.IsInPageRange(rng) Then
                GoTo ContinueSearch
            End If

            ' ── Create the issue ──────────────────────
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            issueText = sourceLabel & " spelling detected: '" & foundText & "'"

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       issueText, _
                       targetWords(i), _
                       rng.Start, _
                       rng.End, _
                       "error"
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a found term is in the exceptions list
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Build the parallel US/UK spelling arrays
'  ~95 pairs across all categories.
' ════════════════════════════════════════════════════════════
Private Sub BuildSpellingArrays(ByRef usWords() As String, _
                                 ByRef ukWords() As String)
    Const PAIR_COUNT As Long = 95

    ReDim usWords(0 To PAIR_COUNT - 1)
    ReDim ukWords(0 To PAIR_COUNT - 1)

    Dim idx As Long
    idx = 0

    ' ── -or → -our (25 pairs) ────────────────────────────
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

    ' ── -ize → -ise (20 pairs) ───────────────────────────
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

    ' ── -ization → -isation (10 pairs) ───────────────────
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

    ' ── -er → -re (10 pairs) ─────────────────────────────
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

    ' ── -se → -ce (3 pairs) ──────────────────────────────
    usWords(idx) = "defense":   ukWords(idx) = "defence":   idx = idx + 1
    usWords(idx) = "offense":   ukWords(idx) = "offence":   idx = idx + 1
    usWords(idx) = "pretense":  ukWords(idx) = "pretence":  idx = idx + 1

    ' ── -og → -ogue (6 pairs) ────────────────────────────
    usWords(idx) = "analog":   ukWords(idx) = "analogue":   idx = idx + 1
    usWords(idx) = "catalog":  ukWords(idx) = "catalogue":  idx = idx + 1
    usWords(idx) = "dialog":   ukWords(idx) = "dialogue":   idx = idx + 1
    usWords(idx) = "monolog":  ukWords(idx) = "monologue":  idx = idx + 1
    usWords(idx) = "prolog":   ukWords(idx) = "prologue":   idx = idx + 1
    usWords(idx) = "epilog":   ukWords(idx) = "epilogue":   idx = idx + 1

    ' ── -ment variants (5 pairs) ─────────────────────────
    usWords(idx) = "judgment":        ukWords(idx) = "judgement":        idx = idx + 1
    usWords(idx) = "acknowledgment":  ukWords(idx) = "acknowledgement":  idx = idx + 1
    usWords(idx) = "fulfillment":     ukWords(idx) = "fulfilment":       idx = idx + 1
    usWords(idx) = "enrollment":      ukWords(idx) = "enrolment":        idx = idx + 1
    usWords(idx) = "installment":     ukWords(idx) = "instalment":       idx = idx + 1

    ' ── Other / miscellaneous (16 pairs) ─────────────────
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
