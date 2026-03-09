Attribute VB_Name = "Rule27_FootnoteAbbreviationDictionary"
' ============================================================
' Rule27_FootnoteAbbreviationDictionary.bas
' Proofreading rule: flags unapproved footnote abbreviation
' variants and dotted forms of approved abbreviations.
'
' Approved abbreviations (without dots):
'   Art, art, Arts, arts, ch, chs, c, cc, cl, cls, cp, cf,
'   ed, eds, edn, edns, eg, etc, f, ff, fn, fns, ibid, ie,
'   MS, MSS, n, nn, no, No, p, pp, para, paras, pt, reg,
'   regs, r, rr, sch, s, ss, sub-s, sub-ss, trans, vol, vols
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "footnote_abbreviation_dictionary"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_FootnoteAbbreviationDictionary(doc As Document) As Collection
    Dim issues As New Collection
    Dim approved As Object
    Dim approvedLC As Object
    Dim unapproved As Object
    Dim fn As Footnote
    Dim i As Long

    ' ── Build approved abbreviations set (case-sensitive) ────
    Set approved = CreateObject("Scripting.Dictionary")
    approved.CompareMode = vbBinaryCompare
    BuildApprovedDict approved

    ' ── Build approved lower-case set for dotted-form check ──
    Set approvedLC = CreateObject("Scripting.Dictionary")
    approvedLC.CompareMode = vbTextCompare
    BuildApprovedLCDict approvedLC

    ' ── Build unapproved variant mapping (LCase key) ────────
    Set unapproved = CreateObject("Scripting.Dictionary")
    unapproved.CompareMode = vbTextCompare
    BuildUnapprovedDict unapproved

    ' ── Process each footnote ────────────────────────────────
    For i = 1 To doc.Footnotes.Count
        On Error Resume Next
        Set fn = doc.Footnotes(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote
        End If
        On Error GoTo 0

        ' ── Check page range on the reference mark ───────────
        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFootnote
        End If
        On Error GoTo 0

        CheckFootnoteText doc, fn, approved, approvedLC, unapproved, issues

NextFootnote:
    Next i

    Set Check_FootnoteAbbreviationDictionary = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check a single footnote's text for abbreviation issues
' ════════════════════════════════════════════════════════════
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
    Dim issue As PleadingsIssue
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

    ' ── Tokenize on spaces ───────────────────────────────────
    tokens = Split(noteText, " ")

    For j = LBound(tokens) To UBound(tokens)
        token = Trim(tokens(j))
        If Len(token) = 0 Then GoTo NextToken

        ' Clean token boundaries: strip leading/trailing non-letter, non-dot chars
        token = CleanTokenBoundaries(token)
        If Len(token) = 0 Then GoTo NextToken

        ' ── Check 1: Unapproved variant (without trailing dot) ──
        stripped = StripTrailingDot(token)
        lcToken = LCase(stripped)

        If unapproved.Exists(lcToken) Then
            preferred = unapproved(lcToken)

            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            issueText = "Unapproved footnote abbreviation."
            suggText = "Use '" & preferred & "' instead of '" & stripped & "'."

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       issueText, _
                       suggText, _
                       fn.Range.Start, _
                       fn.Range.End, _
                       "warning", _
                       False
            issues.Add issue
            GoTo NextToken
        End If

        ' ── Check 2: Dotted form of approved abbreviation ───────
        ' Only flag tokens that contain dots
        If InStr(1, token, ".") > 0 Then
            ' Strip trailing dot and check
            stripped = StripTrailingDot(token)

            ' Remove all internal dots (e.g. "e.g." -> "eg", "i.e." -> "ie")
            noDots = Replace(stripped, ".", "")

            If Len(noDots) > 0 Then
                ' Check if the undotted form is an approved abbreviation
                If approvedLC.Exists(noDots) Then
                    ' This is a dotted form of an approved abbrev — flag it
                    On Error Resume Next
                    locStr = PleadingsEngine.GetLocationString(fn.Reference, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
                    On Error GoTo 0

                    issueText = "Unapproved footnote abbreviation."
                    suggText = "Use '" & noDots & "' instead of '" & token & "'."

                    Set issue = New PleadingsIssue
                    issue.Init RULE_NAME, _
                               locStr, _
                               issueText, _
                               suggText, _
                               fn.Range.Start, _
                               fn.Range.End, _
                               "warning", _
                               False
                    issues.Add issue
                    GoTo NextToken
                End If
            End If
        End If

NextToken:
    Next j
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Build approved abbreviations dictionary
'  (case-sensitive binary compare)
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Build approved lower-case dictionary for dotted
'  form checks (case-insensitive text compare)
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Build unapproved variant mapping
' ════════════════════════════════════════════════════════════
Private Sub BuildUnapprovedDict(ByRef d As Object)
    d.Add "pgs", "pp"
    d.Add "sec", "s"
    d.Add "secs", "ss"
    d.Add "sect", "s"
    d.Add "sects", "ss"
    d.Add "para.", "para"
    d.Add "paras.", "paras"
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Strip a single trailing dot from a token
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Clean token boundaries — strip leading/trailing
'  characters that are not letters, digits, dots, or hyphens
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a character code is a letter or digit
' ════════════════════════════════════════════════════════════
Private Function IsWordChar(ByVal code As Long) As Boolean
    ' A-Z = 65-90, a-z = 97-122, 0-9 = 48-57
    IsWordChar = (code >= 65 And code <= 90) Or _
                 (code >= 97 And code <= 122) Or _
                 (code >= 48 And code <= 57)
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunFootnoteAbbreviationDictionary()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Footnote Abbreviation Dictionary"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_FootnoteAbbreviationDictionary(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Footnote Abbreviation Dictionary"
End Sub
