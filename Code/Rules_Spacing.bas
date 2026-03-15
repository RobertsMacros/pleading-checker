Attribute VB_Name = "Rules_Spacing"
' ============================================================
' Rules_Spacing.bas
' Spacing and whitespace proofreading rules:
'   - Check_DoubleSpaces      : Flag runs of 2+ spaces
'                                (mode-aware: ONE space or TWO after full stop)
'   - Check_DoubleCommas      : Flag ",," sequences
'   - Check_SpaceBeforePunct  : Flag "word ," patterns
'   - Check_MissingSpaceAfterDot : Flag ".X" (missing space)
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString,
'                          GetSpaceStylePref)
' ============================================================
Option Explicit

Private Const RULE_DOUBLE_SPACES As String = "double_spaces"
Private Const RULE_DOUBLE_COMMAS As String = "double_commas"
Private Const RULE_SPACE_BEFORE_PUNCT As String = "space_before_punct"
Private Const RULE_MISSING_SPACE_DOT As String = "missing_space_after_dot"

' Known abbreviations (delimited for InStr lookup)
Private Const ABBREV_LIST As String = _
    "|mr|mrs|ms|dr|prof|sr|jr|st|no|nos|" & _
    "|vs|etc|al|approx|dept|govt|inc|ltd|" & _
    "|corp|co|assn|ave|blvd|rd|ct|ft|" & _
    "|vol|rev|gen|sgt|cpl|pvt|lt|capt|" & _
    "|maj|col|cmdr|adm|jan|feb|mar|apr|" & _
    "|jun|jul|aug|sep|oct|nov|dec|mon|" & _
    "|tue|wed|thu|fri|sat|sun|fig|eq|" & _
    "|ref|para|paras|cl|pt|sch|art|reg|v|"

' ============================================================
'  PUBLIC: Check_DoubleSpaces
'  Flags runs of 2+ spaces. In TWO mode, allows double space
'  after sentence-ending full stops (but not after abbreviations).
' ============================================================
Public Function Check_DoubleSpaces(doc As Document) As Collection
    Dim issues As New Collection
    Dim reDouble As Object
    Set reDouble = CreateObject("VBScript.RegExp")
    reDouble.Global = True
    reDouble.Pattern = " {2,}"

    Dim reSingle As Object
    Set reSingle = CreateObject("VBScript.RegExp")
    reSingle.Global = True
    reSingle.Pattern = "\.( )([A-Z])"

    Dim spaceStyle As String
    spaceStyle = TextAnchoring.GetSpaceStylePref()

    Dim para As Paragraph
    Dim paraText As String
    Dim paraRange As Range
    Dim finding As Object
    Dim locStr As String

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDS

        If TextAnchoring.IsPastPageFilter(paraRange.Start) Then Exit For
        If Not TextAnchoring.IsInPageRange(paraRange) Then GoTo NextParaDS

        ' Block quotes are filtered at engine level by FilterBlockQuoteIssues.
        ' Removed per-paragraph Application.Run("IsBlockQuotePara") call here
        ' to eliminate heavy object-model traffic (font/italic/text/style
        ' per paragraph via cross-module dispatch was a major regression cause).

        paraText = TextAnchoring.StripParaMarkChar(paraRange.Text)
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDS
        If Len(paraText) < 2 Then GoTo NextParaDS

        ' Calculate auto-number prefix offset
        Dim listPrefixLen As Long
        listPrefixLen = TextAnchoring.GetListPrefixLen(para, paraText)

        ' --- Pass 1: Flag runs of 2+ spaces ---
        Dim mDoubles As Object
        Set mDoubles = reDouble.Execute(paraText)
        Dim md As Object
        For Each md In mDoubles
            Dim mStart As Long
            mStart = md.FirstIndex   ' 0-based

            If spaceStyle = "TWO" And mStart > 0 Then
                ' In two-space mode, exactly 2 spaces after sentence-end full stop = correct.
                ' 3+ spaces after full stop still flagged (excess beyond the allowed 2).
                Dim charBefore As String
                charBefore = Mid(paraText, mStart, 1)   ' char at 0-based mStart-1
                If charBefore = "." And md.Length = 2 Then
                    Dim dotPos As Long
                    dotPos = mStart - 1   ' 0-based index of the full stop
                    Dim wb As String
                    wb = GetWordBeforePos(paraText, dotPos)
                    If Not IsLikelyAbbreviation(paraText, dotPos, wb) Then
                        GoTo NextDoubleMatch   ' sentence-end + exactly 2 spaces = correct
                    End If
                End If
            End If

            ' Flag this double space
            Dim dsStart As Long
            Dim dsEnd As Long
            dsStart = paraRange.Start + mStart - listPrefixLen
            dsEnd = dsStart + md.Length

            Err.Clear
            Dim dsRng As Range
            Set dsRng = doc.Range(dsStart, dsEnd)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            Else
                locStr = TextAnchoring.GetLocationString(dsRng, doc)
            End If

            Dim dsMsg As String
            If md.Length = 2 Then
                dsMsg = "Double space found."
            Else
                dsMsg = md.Length & " consecutive spaces found."
            End If

            ' Range covers only the EXTRA space(s) — keep the first one
            ' Store the actual space characters as MatchedText
            Dim dsMatchedText As String
            dsMatchedText = String(md.Length - 1, " ")
            Set finding = TextAnchoring.CreateIssueDict(RULE_DOUBLE_SPACES, locStr, _
                dsMsg, "Remove extra space(s)", dsStart + 1, dsEnd, "error", True, "", _
                dsMatchedText, "exact_text", "high")
            issues.Add finding

NextDoubleMatch:
        Next md

        ' --- Pass 2 (TWO mode only): Flag missing second space after sentence-end ---
        If spaceStyle = "TWO" Then
            Dim mSingles As Object
            Set mSingles = reSingle.Execute(paraText)
            Dim ms As Object
            For Each ms In mSingles
                Dim sdotPos As Long
                sdotPos = ms.FirstIndex   ' 0-based index of the full stop
                Dim swb As String
                swb = GetWordBeforePos(paraText, sdotPos)
                If Not IsLikelyAbbreviation(paraText, sdotPos, swb) Then
                    ' Sentence-end with only one space -- flag it
                    ' Anchor the issue on the full stop + single space
                    Dim msStart As Long
                    msStart = paraRange.Start + sdotPos - listPrefixLen
                    Dim msEnd As Long
                    msEnd = msStart + 2  ' full stop + space

                    Err.Clear
                    Dim msRng As Range
                    Set msRng = doc.Range(msStart, msEnd)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    Else
                        locStr = TextAnchoring.GetLocationString(msRng, doc)
                    End If

                    ' Suggestion replaces ". " with ".  " (insert extra space)
                    Set finding = TextAnchoring.CreateIssueDict(RULE_DOUBLE_SPACES, locStr, _
                        "Missing second space after sentence-ending full stop.", _
                        "Add a second space after the full stop", msStart, msEnd, _
                        "warning", True, ".  ", ". ", "exact_text", "high")
                    issues.Add finding
                End If
            Next ms
        End If

NextParaDS:
    Next para
    On Error GoTo 0

    Set Check_DoubleSpaces = issues
End Function

' ============================================================
'  PUBLIC: Check_DoubleCommas
'  Flags ",," sequences in paragraph text.
' ============================================================
Public Function Check_DoubleCommas(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim paraRange As Range
    Dim finding As Object
    Dim locStr As String
    Dim pos As Long

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDC

        If TextAnchoring.IsPastPageFilter(paraRange.Start) Then Exit For
        If Not TextAnchoring.IsInPageRange(paraRange) Then GoTo NextParaDC

        paraText = paraRange.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDC

        Dim dcListPrefixLen As Long
        dcListPrefixLen = TextAnchoring.GetListPrefixLen(para, paraText)

        pos = InStr(1, paraText, ",,")
        Do While pos > 0
            Dim dcStart As Long
            dcStart = paraRange.Start + pos - 1 - dcListPrefixLen
            Dim dcEnd As Long
            dcEnd = dcStart + 2

            Err.Clear
            Dim dcRng As Range
            Set dcRng = doc.Range(dcStart, dcEnd)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            Else
                locStr = TextAnchoring.GetLocationString(dcRng, doc)
            End If

            Set finding = TextAnchoring.CreateIssueDict(RULE_DOUBLE_COMMAS, locStr, _
                "Double comma found.", "Replace with a single comma", _
                dcStart, dcEnd, "error", True, ",", _
                ",,", "exact_text", "high")
            issues.Add finding

            pos = InStr(pos + 2, paraText, ",,")
        Loop

NextParaDC:
    Next para
    On Error GoTo 0

    Set Check_DoubleCommas = issues
End Function

' ============================================================
'  PUBLIC: Check_SpaceBeforePunct
'  Flags "word ," / "word ;" / "word :" etc. patterns.
' ============================================================
Public Function Check_SpaceBeforePunct(doc As Document) As Collection
    Dim issues As New Collection
    Dim rng As Range
    Set rng = doc.Content.Duplicate
    Dim finding As Object
    Dim locStr As String

    On Error Resume Next
    With rng.Find
        .ClearFormatting
        .Text = " [,;:!?]"
        .MatchCase = True
        .MatchWildcards = True
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    Do While rng.Find.Execute
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        If Not TextAnchoring.IsInPageRange(rng) Then
            rng.Collapse wdCollapseEnd
            GoTo NextSBP
        End If

        Err.Clear
        locStr = TextAnchoring.GetLocationString(rng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

        Dim punctChar As String
        punctChar = Mid(rng.Text, 2, 1)

        ' Range covers only the space (not the punctuation character)
        ' Store the space as MatchedText
        Set finding = TextAnchoring.CreateIssueDict(RULE_SPACE_BEFORE_PUNCT, locStr, _
            "Unexpected space before '" & punctChar & "'", _
            "Remove the space before punctuation", rng.Start, rng.Start + 1, "error", True, "", _
            " ", "exact_text", "high")
        issues.Add finding

        rng.Collapse wdCollapseEnd
NextSBP:
    Loop
    On Error GoTo 0

    Set Check_SpaceBeforePunct = issues
End Function

' ============================================================
'  PUBLIC: Check_MissingSpaceAfterDot
'  Flags ".X" where X is uppercase and the dot is not an
'  abbreviation full stop. Uses per-paragraph regex scanning.
' ============================================================
Public Function Check_MissingSpaceAfterDot(doc As Document) As Collection
    Dim issues As New Collection
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.Pattern = "\.([A-Z])"

    Dim para As Paragraph
    Dim paraText As String
    Dim paraRange As Range
    Dim finding As Object
    Dim locStr As String

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaMSD

        If TextAnchoring.IsPastPageFilter(paraRange.Start) Then Exit For
        If Not TextAnchoring.IsInPageRange(paraRange) Then GoTo NextParaMSD

        paraText = TextAnchoring.StripParaMarkChar(paraRange.Text)
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaMSD
        If Len(paraText) < 2 Then GoTo NextParaMSD

        Dim msdListPrefixLen As Long
        msdListPrefixLen = TextAnchoring.GetListPrefixLen(para, paraText)

        Dim matches As Object
        Set matches = re.Execute(paraText)
        Dim m As Object
        For Each m In matches
            Dim dotIdx As Long
            dotIdx = m.FirstIndex   ' 0-based position of the full stop
            Dim wordBefore As String
            wordBefore = GetWordBeforePos(paraText, dotIdx)
            If Not IsLikelyAbbreviation(paraText, dotIdx, wordBefore) Then
                Dim msdStart As Long
                msdStart = paraRange.Start + dotIdx - msdListPrefixLen
                Dim msdEnd As Long
                msdEnd = msdStart + 2   ' "." + capital letter

                Err.Clear
                Dim msdRng As Range
                Set msdRng = doc.Range(msdStart, msdEnd)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                Else
                    locStr = TextAnchoring.GetLocationString(msdRng, doc)
                End If

                Set finding = TextAnchoring.CreateIssueDict(RULE_MISSING_SPACE_DOT, locStr, _
                    "Missing space after full stop before '" & _
                    Mid(paraText, dotIdx + 2, 1) & "'.", _
                    "Insert a space after the full stop.", _
                    msdStart, msdEnd, "error", False)
                issues.Add finding
            End If
        Next m

NextParaMSD:
    Next para
    On Error GoTo 0

    Set Check_MissingSpaceAfterDot = issues
End Function


' ============================================================
'  PRIVATE HELPERS
' ============================================================

' Return the word (letters only) immediately before 0-based position pos
Private Function GetWordBeforePos(ByVal s As String, ByVal pos As Long) As String
    Dim result As String
    result = ""
    Dim i As Long
    Dim c As String
    For i = pos - 1 To 0 Step -1
        c = Mid$(s, i + 1, 1)   ' convert 0-based to 1-based for Mid
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Then
            result = c & result
        Else
            Exit For
        End If
    Next i
    GetWordBeforePos = result
End Function

' Check if word is a known abbreviation
Private Function IsAbbrevWord(ByVal word As String) As Boolean
    IsAbbrevWord = (InStr(1, ABBREV_LIST, "|" & LCase(word) & "|", vbTextCompare) > 0)
End Function

' Extended abbreviation detection:
'   1. Known abbreviation list (Mr, Dr, etc, vs ...)
'   2. Dotted abbreviation: wordBefore is 1-2 chars preceded by a dot (e.g. "e" in "i.e.")
'   3. Ellipsis: empty wordBefore preceded by a dot ("...Word")
'   4. First dot of dotted abbreviation: 1-char wordBefore followed by letter+dot
'
' Index arithmetic: pos is 0-based; Mid uses 1-based.
'   char at 0-based N = Mid(s, N+1, 1)
Private Function IsLikelyAbbreviation(ByVal paraText As String, _
                                       ByVal pos As Long, _
                                       ByVal wordBefore As String) As Boolean
    IsLikelyAbbreviation = False

    ' 1. Standard abbreviation list
    If IsAbbrevWord(wordBefore) Then
        IsLikelyAbbreviation = True
        Exit Function
    End If

    ' 1b. Single uppercase letter (initial: "J. Smith", "A. Jones")
    If Len(wordBefore) = 1 Then
        Dim wbCode As Long
        wbCode = AscW(wordBefore)
        If wbCode >= 65 And wbCode <= 90 Then  ' A-Z
            IsLikelyAbbreviation = True
            Exit Function
        End If
    End If

    ' 2. Dotted abbreviation: wordBefore is 1-2 chars and char before it is a dot
    If Len(wordBefore) >= 1 And Len(wordBefore) <= 2 And _
       pos > Len(wordBefore) Then
        If Mid$(paraText, pos - Len(wordBefore), 1) = "." Then
            IsLikelyAbbreviation = True
            Exit Function
        End If
    End If

    ' 3. Ellipsis: empty wordBefore and char before this dot is also a dot
    If Len(wordBefore) = 0 And pos >= 1 Then
        If Mid$(paraText, pos, 1) = "." Then
            IsLikelyAbbreviation = True
            Exit Function
        End If
    End If

    ' 4. First dot of dotted abbreviation: 1-char wordBefore and
    '    char after this dot is letter followed by another dot
    If Len(wordBefore) = 1 And pos + 2 < Len(paraText) Then
        If Mid$(paraText, pos + 2, 1) Like "[A-Za-z]" And _
           Mid$(paraText, pos + 3, 1) = "." Then
            IsLikelyAbbreviation = True
            Exit Function
        End If
    End If
End Function

