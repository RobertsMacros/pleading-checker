Attribute VB_Name = "Rule07_DefinedTerms"
' ============================================================
' Rule07_DefinedTerms.bas
' Scans for defined terms (curly-quoted, "means" definitions,
' and parenthetical definitions), then checks:
'   1. Inconsistent capitalisation of the defined term
'   2. Hyphenated/unhyphenated variant inconsistency
'   3. Defined terms that are never referenced after definition
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "defined_terms"

' ── Helper: remove hyphens from a term ──────────────────────
Private Function RemoveHyphens(ByVal term As String) As String
    RemoveHyphens = Replace(term, "-", "")
End Function

' ── Helper: count occurrences of a term in document text ────
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

' ── Helper: find first occurrence of a term and return range ─
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

' ════════════════════════════════════════════════════════════
'  MAIN RULE FUNCTION
' ════════════════════════════════════════════════════════════
Public Function Check_DefinedTerms(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' Dictionary: term (String) -> Array(definitionParaIdx, rangeStart, rangeEnd)
    Dim definedTerms As New Scripting.Dictionary
    Dim rng As Range
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim paraText As String

    ' ══════════════════════════════════════════════════════════
    '  PASS 1: Scan for defined terms
    ' ══════════════════════════════════════════════════════════

    ' ── Pattern A: Curly-quoted defined terms ────────────────
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
        If Not PleadingsEngine.IsInPageRange(rng) Then
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
                Dim defInfo(0 To 2) As Variant
                defInfo(0) = 0 ' paragraph index (approximate)
                defInfo(1) = startPos + 1  ' range start of term
                defInfo(2) = startPos + CLng(closePos) - 1  ' range end of term
                definedTerms.Add termText, defInfo
            End If
        End If

        rng.Collapse wdCollapseEnd
NextCurlyFind:
    Loop

    ' ── Pattern B: "X means " or "X has the meaning " ───────
    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1
        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextParaMeans

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
                        Dim mInfo(0 To 2) As Variant
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
                        Dim hInfo(0 To 2) As Variant
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

    ' ── Pattern C: Parenthetical definitions (the "Term") ───
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
        If Not PleadingsEngine.IsInPageRange(rng) Then
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
                    Dim pInfo(0 To 2) As Variant
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

    ' ══════════════════════════════════════════════════════════
    '  PASS 2: Validate each defined term
    ' ══════════════════════════════════════════════════════════
    Dim termKey As Variant
    For Each termKey In definedTerms.keys
        Dim term As String
        term = CStr(termKey)
        Dim tInfo As Variant
        tInfo = definedTerms(termKey)

        ' ── Check A: Lowercase variant (inconsistent capitalisation) ──
        Dim lcTerm As String
        lcTerm = LCase(Left$(term, 1)) & Mid$(term, 2)
        If lcTerm <> term Then
            Dim lcRng As Range
            Set lcRng = FindTermRange(doc, lcTerm, True)
            If Not lcRng Is Nothing Then
                If PleadingsEngine.IsInPageRange(lcRng) Then
                    Dim issueLC As New PleadingsIssue
                    Dim locLC As String
                    locLC = PleadingsEngine.GetLocationString(lcRng, doc)
                    issueLC.Init RULE_NAME, locLC, _
                        "Inconsistent capitalisation: '" & lcTerm & _
                        "' found but '" & term & "' is the defined term", _
                        "Use '" & term & "' consistently", _
                        lcRng.Start, lcRng.End, "error"
                    issues.Add issueLC
                End If
            End If
        End If

        ' ── Check B: Hyphenated/unhyphenated variant ──────────
        If InStr(1, term, "-") > 0 Then
            Dim noHyphen As String
            noHyphen = RemoveHyphens(term)
            Dim nhRng As Range
            Set nhRng = FindTermRange(doc, noHyphen, False)
            If Not nhRng Is Nothing Then
                If PleadingsEngine.IsInPageRange(nhRng) Then
                    Dim issueH As New PleadingsIssue
                    Dim locH As String
                    locH = PleadingsEngine.GetLocationString(nhRng, doc)
                    issueH.Init RULE_NAME, locH, _
                        "Hyphenation variant: '" & noHyphen & _
                        "' found but defined term uses hyphen: '" & term & "'", _
                        "Use the defined form: '" & term & "'", _
                        nhRng.Start, nhRng.End, "error"
                    issues.Add issueH
                End If
            End If
        Else
            ' Term has no hyphen — check if hyphenated variant exists
            ' Try common hyphenation points (before common prefixes)
            ' This is a best-effort check
        End If

        ' ── Check C: Defined term never referenced ────────────
        Dim totalCount As Long
        totalCount = CountTermInDoc(doc, term)
        If totalCount <= 1 Then
            ' Only appears at the definition site
            Dim issueUnused As New PleadingsIssue
            Dim unusedRng As Range
            Set unusedRng = doc.Range(CLng(tInfo(1)), CLng(tInfo(2)))
            Dim locUnused As String
            locUnused = PleadingsEngine.GetLocationString(unusedRng, doc)
            issueUnused.Init RULE_NAME, locUnused, _
                "Defined term never referenced: '" & term & _
                "' is defined but not used elsewhere in the document", _
                "Remove the definition or use the term in the document", _
                CLng(tInfo(1)), CLng(tInfo(2)), "possible_error"
            issues.Add issueUnused
        End If
    Next termKey

    On Error GoTo 0
    Set Check_DefinedTerms = issues
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunDefinedTerms()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Defined Terms"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_DefinedTerms(doc)

    ' ── Highlight issues in document ─────────────────────────
    Dim iss As PleadingsIssue
    Dim rng As Range
    Dim i As Long
    For i = 1 To issues.Count
        Set iss = issues(i)
        If iss.RangeStart >= 0 And iss.RangeEnd > iss.RangeStart Then
            On Error Resume Next
            Set rng = doc.Range(iss.RangeStart, iss.RangeEnd)
            rng.HighlightColorIndex = wdYellow
            doc.Comments.Add Range:=rng, _
                Text:="[" & iss.RuleName & "] " & iss.Issue & _
                      " " & Chr(8212) & " Suggestion: " & iss.Suggestion
            On Error GoTo 0
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Defined Terms"
End Sub
