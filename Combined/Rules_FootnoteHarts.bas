Attribute VB_Name = "Rules_FootnoteHarts"
' ============================================================
' Rules_FootnoteHarts.bas
' Combined proofreading rules for footnotes per Hart's Rules:
'   - Rule24: flags documents that use endnotes instead of footnotes
'   - Rule25: every footnote should end with a full stop
'   - Rule26: footnotes should begin with a capital letter
'   - Rule27: flags unapproved footnote abbreviation variants
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' ------------------------------------------------------------
'  Rule-name constants
' ------------------------------------------------------------
Private Const RULE24_NAME As String = "footnotes_not_endnotes"
Private Const RULE25_NAME As String = "footnote_terminal_full_stop"
Private Const RULE26_NAME As String = "footnote_initial_capital"
Private Const RULE27_NAME As String = "footnote_abbreviation_dictionary"

' ============================================================
'  RULE 24 -- FOOTNOTES NOT ENDNOTES
' ============================================================

Public Function Check_FootnotesNotEndnotes(doc As Document) As Collection
    Dim issues As New Collection
    Dim finding As Object

    On Error Resume Next

    Dim endCount As Long
    Dim fnCount As Long
    endCount = doc.Endnotes.Count
    fnCount = doc.Footnotes.Count

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set Check_FootnotesNotEndnotes = issues
        Exit Function
    End If
    On Error GoTo 0

    If endCount > 0 And fnCount = 0 Then
        ' Document uses only endnotes
        Set finding = CreateIssueDict(RULE24_NAME, "document level", "Document uses endnotes instead of footnotes.", "Use footnotes rather than endnotes.", 0, 0, "error", False)
        issues.Add finding

    ElseIf endCount > 0 And fnCount > 0 Then
        ' Document uses both
        Set finding = CreateIssueDict(RULE24_NAME, "document level", "Document uses both footnotes and endnotes.", "Use footnotes rather than endnotes.", 0, 0, "error", False)
        issues.Add finding
    End If

    ' If only footnotes exist (endCount = 0): no finding

    Set Check_FootnotesNotEndnotes = issues
End Function

' ============================================================
'  RULE 25 -- FOOTNOTE TERMINAL FULL STOP
' ============================================================

Public Function Check_FootnoteTerminalFullStop(doc As Document) As Collection
    Dim issues As New Collection
    Dim fn As Footnote
    Dim finding As Object
    Dim locStr As String
    Dim noteText As String
    Dim trimmed As String
    Dim lastChar As String
    Dim penultChar As String
    Dim i As Long

    For i = 1 To doc.Footnotes.Count
        On Error Resume Next
        Set fn = doc.Footnotes(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote25
        End If
        On Error GoTo 0

        ' -- Check page range on the reference mark -----------
        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFootnote25
        End If
        On Error GoTo 0

        ' -- Get footnote text --------------------------------
        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote25
        End If
        On Error GoTo 0

        ' -- Trim trailing whitespace / paragraph marks -------
        trimmed = noteText
        trimmed = TrimTrailingWhitespace(trimmed)

        ' -- Skip empty footnotes -----------------------------
        If Len(trimmed) = 0 Then GoTo NextFootnote25

        ' -- Get last character -------------------------------
        lastChar = Mid(trimmed, Len(trimmed), 1)

        ' -- If last char is closing bracket/quote, check penultimate --
        If IsClosingPunctuation(lastChar) Then
            If Len(trimmed) >= 2 Then
                penultChar = Mid(trimmed, Len(trimmed) - 1, 1)
                If penultChar = "." Then GoTo NextFootnote25
            End If
            ' Fall through to flag
        ElseIf lastChar = "." Then
            GoTo NextFootnote25
        End If

        ' -- Flag missing full stop ---------------------------
        On Error Resume Next
        locStr = EngineGetLocationString(fn.Reference, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
        On Error GoTo 0

        Set finding = CreateIssueDict(RULE25_NAME, locStr, "Footnote does not end with a full stop.", "Add a full stop at the end of the footnote.", fn.Range.Start, fn.Range.End, "warning", False)
        issues.Add finding

NextFootnote25:
    Next i

    Set Check_FootnoteTerminalFullStop = issues
End Function

' ============================================================
'  RULE 26 -- FOOTNOTE INITIAL CAPITAL
' ============================================================

Public Function Check_FootnoteInitialCapital(doc As Document) As Collection
    Dim issues As New Collection
    Dim allowed As Object
    Dim fn As Footnote
    Dim finding As Object
    Dim locStr As String
    Dim noteText As String
    Dim trimmed As String
    Dim token As String
    Dim firstCharCode As Long
    Dim i As Long
    Dim j As Long
    Dim ch As String

    ' -- Build allowed lower-case starts dictionary -----------
    Set allowed = CreateObject("Scripting.Dictionary")
    allowed.CompareMode = vbTextCompare
    allowed.Add "c", True
    allowed.Add "cf", True
    allowed.Add "cp", True
    allowed.Add "eg", True
    allowed.Add "ie", True
    allowed.Add "p", True
    allowed.Add "pp", True
    allowed.Add "ibid", True

    For i = 1 To doc.Footnotes.Count
        On Error Resume Next
        Set fn = doc.Footnotes(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote26
        End If
        On Error GoTo 0

        ' -- Check page range on the reference mark -----------
        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFootnote26
        End If
        On Error GoTo 0

        ' -- Get footnote text --------------------------------
        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote26
        End If
        On Error GoTo 0

        ' -- Trim leading whitespace --------------------------
        trimmed = LTrim(noteText)
        If Len(trimmed) = 0 Then GoTo NextFootnote26

        ' -- Skip past leading punctuation (quotes, brackets) -
        j = 1
        Do While j <= Len(trimmed)
            ch = Mid(trimmed, j, 1)
            If IsLeadingPunctuation(ch) Then
                j = j + 1
            Else
                Exit Do
            End If
        Loop

        If j > Len(trimmed) Then GoTo NextFootnote26
        trimmed = Mid(trimmed, j)
        If Len(trimmed) = 0 Then GoTo NextFootnote26

        ' -- Extract first lexical token (letters only) -------
        token = ExtractFirstToken(trimmed)
        If Len(token) = 0 Then GoTo NextFootnote26

        ' -- Check if token is in allowed list ----------------
        If allowed.Exists(LCase(token)) Then GoTo NextFootnote26

        ' -- Check if first character is lower-case -----------
        firstCharCode = AscW(Mid(token, 1, 1))
        If firstCharCode >= 97 And firstCharCode <= 122 Then
            ' Lower-case and not in allowed list: flag
            On Error Resume Next
            locStr = EngineGetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE26_NAME, locStr, "Footnote begins with lower-case text outside the approved exceptions.", "Begin the footnote with a capital letter, unless it starts with an approved lower-case abbreviation.", fn.Range.Start, fn.Range.End, "warning", False)
            issues.Add finding
        End If

NextFootnote26:
    Next i

    Set Check_FootnoteInitialCapital = issues
End Function

' ============================================================
'  RULE 27 -- FOOTNOTE ABBREVIATION DICTIONARY
' ============================================================

Public Function Check_FootnoteAbbreviationDictionary(doc As Document) As Collection
    Dim issues As New Collection
    Dim approved As Object
    Dim approvedLC As Object
    Dim unapproved As Object
    Dim fn As Footnote
    Dim i As Long

    ' -- Build approved abbreviations set (case-sensitive) ----
    Set approved = CreateObject("Scripting.Dictionary")
    approved.CompareMode = vbBinaryCompare
    BuildApprovedDict approved

    ' -- Build approved lower-case set for dotted-form check --
    Set approvedLC = CreateObject("Scripting.Dictionary")
    approvedLC.CompareMode = vbTextCompare
    BuildApprovedLCDict approvedLC

    ' -- Build unapproved variant mapping (LCase key) --------
    Set unapproved = CreateObject("Scripting.Dictionary")
    unapproved.CompareMode = vbTextCompare
    BuildUnapprovedDict unapproved

    ' -- Process each footnote --------------------------------
    For i = 1 To doc.Footnotes.Count
        On Error Resume Next
        Set fn = doc.Footnotes(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote27
        End If
        On Error GoTo 0

        ' -- Check page range on the reference mark -----------
        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFootnote27
        End If
        On Error GoTo 0

        CheckFootnoteText doc, fn, approved, approvedLC, unapproved, issues

NextFootnote27:
    Next i

    Set Check_FootnoteAbbreviationDictionary = issues
End Function

' ============================================================
'  PRIVATE HELPERS -- Rule 25
' ============================================================

' Strip trailing CR, LF, VT, and spaces
Private Function TrimTrailingWhitespace(ByVal s As String) As String
    Dim ch As String
    Do While Len(s) > 0
        ch = Mid(s, Len(s), 1)
        Select Case ch
            Case vbCr, vbLf, Chr(13), Chr(10), Chr(11), " ", vbTab
                s = Left(s, Len(s) - 1)
            Case Else
                Exit Do
        End Select
    Loop
    TrimTrailingWhitespace = s
End Function

' Check if character is a closing bracket or quote
Private Function IsClosingPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case ")", "]", ChrW(8217), ChrW(8221)
            IsClosingPunctuation = True
        Case Else
            IsClosingPunctuation = False
    End Select
End Function

' ============================================================
'  PRIVATE HELPERS -- Rule 26
' ============================================================

' Check if character is leading punctuation to skip
Private Function IsLeadingPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case "(", "[", ChrW(8216), ChrW(8220), """", "'"
            IsLeadingPunctuation = True
        Case Else
            IsLeadingPunctuation = False
    End Select
End Function

' Extract the first token of letters from a string
Private Function ExtractFirstToken(ByVal s As String) As String
    Dim i As Long
    Dim charCode As Long
    Dim result As String
    result = ""

    For i = 1 To Len(s)
        charCode = AscW(Mid(s, i, 1))
        ' A-Z = 65-90, a-z = 97-122
        If (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Then
            result = result & Mid(s, i, 1)
        Else
            Exit For
        End If
    Next i

    ExtractFirstToken = result
End Function

' ============================================================
'  PRIVATE HELPERS -- Rule 27
' ============================================================

' Check a single footnote's text for abbreviation issues
Private Sub CheckFootnoteText(doc As Document, _
                               fn As Footnote, _
                               ByRef approved As Object, _
                               ByRef approvedLC As Object, _
                               ByRef unapproved As Object, _
                               ByRef issues As Collection)
    Dim noteText As String
    Dim tokens() As String
    Dim token As String
    Dim stripped As String
    Dim noDots As String
    Dim lcToken As String
    Dim preferred As String
    Dim finding As Object
    Dim locStr As String
    Dim j As Long
    Dim issueText As String
    Dim suggText As String

    On Error Resume Next
    noteText = fn.Range.Text
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' -- Tokenize on spaces -----------------------------------
    tokens = Split(noteText, " ")

    For j = LBound(tokens) To UBound(tokens)
        token = Trim(tokens(j))
        If Len(token) = 0 Then GoTo NextToken

        ' Clean token boundaries: strip leading/trailing non-letter, non-dot chars
        token = CleanTokenBoundaries(token)
        If Len(token) = 0 Then GoTo NextToken

        ' -- Check 1: Unapproved variant (without trailing dot) --
        stripped = StripTrailingDot(token)
        lcToken = LCase(stripped)

        If unapproved.Exists(lcToken) Then
            preferred = unapproved(lcToken)

            On Error Resume Next
            locStr = EngineGetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            issueText = "Unapproved footnote abbreviation."
            suggText = "Use '" & preferred & "' instead of '" & stripped & "'."

            Set finding = CreateIssueDict(RULE27_NAME, locStr, issueText, suggText, fn.Range.Start, fn.Range.End, "warning", False)
            issues.Add finding
            GoTo NextToken
        End If

        ' -- Check 2: Dotted form of approved abbreviation -------
        ' Only flag tokens that contain dots
        If InStr(1, token, ".") > 0 Then
            ' Strip trailing dot and check
            stripped = StripTrailingDot(token)

            ' Remove all internal dots (e.g. "e.g." -> "eg", "i.e." -> "ie")
            noDots = Replace(stripped, ".", "")

            If Len(noDots) > 0 Then
                ' Check if the undotted form is an approved abbreviation
                If approvedLC.Exists(noDots) Then
                    ' This is a dotted form of an approved abbrev -- flag it
                    On Error Resume Next
                    locStr = EngineGetLocationString(fn.Reference, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
                    On Error GoTo 0

                    issueText = "Unapproved footnote abbreviation."
                    suggText = "Use '" & noDots & "' instead of '" & token & "'."

                    Set finding = CreateIssueDict(RULE27_NAME, locStr, issueText, suggText, fn.Range.Start, fn.Range.End, "warning", False)
                    issues.Add finding
                    GoTo NextToken
                End If
            End If
        End If

NextToken:
    Next j
End Sub

' Build approved abbreviations dictionary (case-sensitive binary compare)
Private Sub BuildApprovedDict(ByRef d As Object)
    d.Add "Art", True
    d.Add "art", True
    d.Add "Arts", True
    d.Add "arts", True
    d.Add "ch", True
    d.Add "chs", True
    d.Add "c", True
    d.Add "cc", True
    d.Add "cl", True
    d.Add "cls", True
    d.Add "cp", True
    d.Add "cf", True
    d.Add "ed", True
    d.Add "eds", True
    d.Add "edn", True
    d.Add "edns", True
    d.Add "eg", True
    d.Add "etc", True
    d.Add "f", True
    d.Add "ff", True
    d.Add "fn", True
    d.Add "fns", True
    d.Add "ibid", True
    d.Add "ie", True
    d.Add "MS", True
    d.Add "MSS", True
    d.Add "n", True
    d.Add "nn", True
    d.Add "no", True
    d.Add "No", True
    d.Add "p", True
    d.Add "pp", True
    d.Add "para", True
    d.Add "paras", True
    d.Add "pt", True
    d.Add "reg", True
    d.Add "regs", True
    d.Add "r", True
    d.Add "rr", True
    d.Add "sch", True
    d.Add "s", True
    d.Add "ss", True
    d.Add "sub-s", True
    d.Add "sub-ss", True
    d.Add "trans", True
    d.Add "vol", True
    d.Add "vols", True
End Sub

' Build approved lower-case dictionary for dotted form checks (case-insensitive)
Private Sub BuildApprovedLCDict(ByRef d As Object)
    Dim abbrevs As Variant
    Dim k As Long
    abbrevs = Array("art", "arts", "ch", "chs", "c", "cc", "cl", "cls", _
                    "cp", "cf", "ed", "eds", "edn", "edns", "eg", "etc", _
                    "f", "ff", "fn", "fns", "ibid", "ie", "ms", "mss", _
                    "n", "nn", "no", "p", "pp", "para", "paras", "pt", _
                    "reg", "regs", "r", "rr", "sch", "s", "ss", _
                    "trans", "vol", "vols")
    For k = LBound(abbrevs) To UBound(abbrevs)
        If Not d.Exists(CStr(abbrevs(k))) Then
            d.Add CStr(abbrevs(k)), True
        End If
    Next k
End Sub

' Build unapproved variant mapping
Private Sub BuildUnapprovedDict(ByRef d As Object)
    d.Add "pgs", "pp"
    d.Add "sec", "s"
    d.Add "secs", "ss"
    d.Add "sect", "s"
    d.Add "sects", "ss"
    d.Add "para.", "para"
    d.Add "paras.", "paras"
End Sub

' Strip a single trailing dot from a token
Private Function StripTrailingDot(ByVal s As String) As String
    If Len(s) > 0 Then
        If Right(s, 1) = "." Then
            StripTrailingDot = Left(s, Len(s) - 1)
        Else
            StripTrailingDot = s
        End If
    Else
        StripTrailingDot = s
    End If
End Function

' Clean token boundaries -- strip leading/trailing characters
' that are not letters, digits, dots, or hyphens
Private Function CleanTokenBoundaries(ByVal s As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim ch As String
    Dim code As Long

    ' Strip leading non-word chars (keep letters, digits, dots, hyphens)
    startPos = 1
    Do While startPos <= Len(s)
        ch = Mid(s, startPos, 1)
        code = AscW(ch)
        If IsWordChar(code) Or ch = "." Or ch = "-" Then
            Exit Do
        End If
        startPos = startPos + 1
    Loop

    ' Strip trailing non-word chars (keep letters, digits, dots, hyphens)
    endPos = Len(s)
    Do While endPos >= startPos
        ch = Mid(s, endPos, 1)
        code = AscW(ch)
        If IsWordChar(code) Or ch = "." Or ch = "-" Then
            Exit Do
        End If
        endPos = endPos - 1
    Loop

    If startPos > endPos Then
        CleanTokenBoundaries = ""
    Else
        CleanTokenBoundaries = Mid(s, startPos, endPos - startPos + 1)
    End If
End Function

' Check if a character code is a letter or digit
Private Function IsWordChar(ByVal code As Long) As Boolean
    ' A-Z = 65-90, a-z = 97-122, 0-9 = 48-57
    IsWordChar = (code >= 65 And code <= 90) Or _
                 (code >= 97 And code <= 122) Or _
                 (code >= 48 And code <= 57)
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
