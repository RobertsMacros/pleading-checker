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
'   - TextAnchoring.bas (IterateParagraphs, AddIssue,
'                        SafeRange, CreateRegex, GetSpaceStylePref)
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
    Set Check_DoubleSpaces = TextAnchoring.IterateParagraphs(doc, "Rules_Spacing", "ProcessParagraph_DoubleSpaces")
End Function

' ============================================================
'  PUBLIC: Check_DoubleCommas
'  Flags ",," sequences in paragraph text.
' ============================================================
Public Function Check_DoubleCommas(doc As Document) As Collection
    Set Check_DoubleCommas = TextAnchoring.IterateParagraphs(doc, "Rules_Spacing", "ProcessParagraph_DoubleCommas")
End Function

' ============================================================
'  PUBLIC: Check_SpaceBeforePunct
'  Flags "word ," / "word ;" / "word :" etc. patterns.
'  Standalone entry point -- delegates to IterateParagraphs.
'  (In the engine, ProcessParagraph_SpaceBeforePunct is called
'  directly from RunParagraphRules for single-pass efficiency.)
' ============================================================
Public Function Check_SpaceBeforePunct(doc As Document) As Collection
    Set Check_SpaceBeforePunct = TextAnchoring.IterateParagraphs(doc, "Rules_Spacing", "ProcessParagraph_SpaceBeforePunct")
End Function

' ============================================================
'  PUBLIC: Check_MissingSpaceAfterDot
'  Flags ".X" where X is uppercase and the dot is not an
'  abbreviation full stop. Uses per-paragraph regex scanning.
' ============================================================
Public Function Check_MissingSpaceAfterDot(doc As Document) As Collection
    Set Check_MissingSpaceAfterDot = TextAnchoring.IterateParagraphs(doc, "Rules_Spacing", "ProcessParagraph_MissingSpaceAfterDot")
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

' ============================================================
'  PUBLIC: ProcessParagraph_DoubleSpaces
'  Per-paragraph handler extracted from Check_DoubleSpaces.
'  Flags runs of 2+ spaces and (in TWO mode) missing second
'  space after sentence-ending full stops.
' ============================================================
Public Sub ProcessParagraph_DoubleSpaces(doc As Document, paraRange As Range, paraText As String, paraStart As Long, listPrefixLen As Long, ByRef issues As Collection)
    Dim reDouble As Object: Set reDouble = TextAnchoring.CreateRegex(" {2,}")
    Dim reSingle As Object: Set reSingle = TextAnchoring.CreateRegex("\.( )([A-Z])")
    Dim spaceStyle As String: spaceStyle = TextAnchoring.GetSpaceStylePref()

    On Error Resume Next

    ' --- Pass 1: Flag runs of 2+ spaces ---
    Dim md As Object
    For Each md In reDouble.Execute(paraText)
        Dim mStart As Long: mStart = md.FirstIndex
        If spaceStyle = "TWO" And mStart > 0 Then
            Dim charBefore As String: charBefore = Mid(paraText, mStart, 1)
            If charBefore = "." And md.Length = 2 Then
                Dim dotPos As Long: dotPos = mStart - 1
                If Not IsLikelyAbbreviation(paraText, dotPos, GetWordBeforePos(paraText, dotPos)) Then GoTo NextDoubleMatchPP
            End If
        End If

        Dim dsStart As Long: dsStart = paraStart + mStart - listPrefixLen
        Dim dsEnd As Long: dsEnd = dsStart + md.Length
        Dim dsMsg As String
        If md.Length = 2 Then dsMsg = "Double space found." Else dsMsg = md.Length & " consecutive spaces found."
        Dim dsRng As Range: Set dsRng = TextAnchoring.SafeRange(doc, dsStart, dsEnd)
        TextAnchoring.AddIssue issues, RULE_DOUBLE_SPACES, doc, dsRng, dsMsg, "Remove extra space(s)", dsStart, dsEnd, "error", True, " ", String(md.Length, " "), "exact_text", "high"
NextDoubleMatchPP:
    Next md

    ' --- Pass 2 (TWO mode only): Flag missing second space after sentence-end ---
    If spaceStyle = "TWO" Then
        Dim ms As Object
        For Each ms In reSingle.Execute(paraText)
            Dim sdotPos As Long: sdotPos = ms.FirstIndex
            If Not IsLikelyAbbreviation(paraText, sdotPos, GetWordBeforePos(paraText, sdotPos)) Then
                Dim msStart As Long: msStart = paraStart + sdotPos - listPrefixLen
                Dim msEnd As Long: msEnd = msStart + 2
                Dim msRng As Range: Set msRng = TextAnchoring.SafeRange(doc, msStart, msEnd)
                TextAnchoring.AddIssue issues, RULE_DOUBLE_SPACES, doc, msRng, "Missing second space after sentence-ending full stop.", "Add a second space after the full stop", msStart, msEnd, "warning", True, ".  ", ". ", "exact_text", "high"
            End If
        Next ms
    End If
    On Error GoTo 0
End Sub

' ============================================================
'  PUBLIC: ProcessParagraph_DoubleCommas
'  Per-paragraph handler extracted from Check_DoubleCommas.
'  Flags ",," sequences in paragraph text.
' ============================================================
Public Sub ProcessParagraph_DoubleCommas(doc As Document, paraRange As Range, paraText As String, paraStart As Long, listPrefixLen As Long, ByRef issues As Collection)
    Dim pos As Long
    pos = InStr(1, paraText, ",,")
    Do While pos > 0
        Dim dcStart As Long: dcStart = paraStart + pos - 1 - listPrefixLen
        Dim dcEnd As Long: dcEnd = dcStart + 2
        Dim rng As Range: Set rng = TextAnchoring.SafeRange(doc, dcStart, dcEnd)
        TextAnchoring.AddIssue issues, RULE_DOUBLE_COMMAS, doc, rng, "Double comma found.", "Replace with a single comma", dcStart, dcEnd, "error", True, ",", ",,", "exact_text", "high"
        pos = InStr(pos + 2, paraText, ",,")
    Loop
End Sub

' ============================================================
'  PUBLIC: ProcessParagraph_MissingSpaceAfterDot
'  Per-paragraph handler extracted from Check_MissingSpaceAfterDot.
'  Flags ".X" where X is uppercase and the dot is not an
'  abbreviation full stop.
' ============================================================
Public Sub ProcessParagraph_MissingSpaceAfterDot(doc As Document, paraRange As Range, paraText As String, paraStart As Long, listPrefixLen As Long, ByRef issues As Collection)
    Dim re As Object: Set re = TextAnchoring.CreateRegex("\.([A-Z])")

    Dim m As Object
    For Each m In re.Execute(paraText)
        Dim dotIdx As Long: dotIdx = m.FirstIndex
        If Not IsLikelyAbbreviation(paraText, dotIdx, GetWordBeforePos(paraText, dotIdx)) Then
            Dim msdStart As Long: msdStart = paraStart + dotIdx - listPrefixLen
            Dim msdEnd As Long: msdEnd = msdStart + 2
            Dim rng As Range: Set rng = TextAnchoring.SafeRange(doc, msdStart, msdEnd)
            TextAnchoring.AddIssue issues, RULE_MISSING_SPACE_DOT, doc, rng, "Missing space after full stop before '" & Mid(paraText, dotIdx + 2, 1) & "'.", "Insert a space after the full stop.", msdStart, msdEnd, "error", False
        End If
    Next m
End Sub

' ============================================================
'  PUBLIC: ProcessParagraph_SpaceBeforePunct
'  Per-paragraph handler for space-before-punctuation detection.
'  Scans paraText with regex for " [,;:!?]" patterns.
' ============================================================
Public Sub ProcessParagraph_SpaceBeforePunct(doc As Document, paraRange As Range, paraText As String, paraStart As Long, listPrefixLen As Long, ByRef issues As Collection)
    Dim re As Object: Set re = TextAnchoring.CreateRegex(" [,;:!?]")
    Dim m As Object
    For Each m In re.Execute(paraText)
        Dim mIdx As Long: mIdx = m.FirstIndex
        Dim spStart As Long: spStart = paraStart + mIdx - listPrefixLen
        Dim spEnd As Long: spEnd = spStart + 2
        Dim punctChar As String: punctChar = Mid(paraText, mIdx + 2, 1)
        Dim rng As Range: Set rng = TextAnchoring.SafeRange(doc, spStart, spEnd)
        ' Conservative: AutoFixSafe=False because false positives near
        ' URLs, code snippets, or stylistic spacing make deletion risky.
        TextAnchoring.AddIssue issues, RULE_SPACE_BEFORE_PUNCT, doc, rng, "Unexpected space before '" & punctChar & "'", "Remove the space before punctuation", spStart, spStart + 1, "error", False, "", " ", "exact_text", "high"
    Next m
End Sub

