Attribute VB_Name = "Rule28_MandatedLegalTermForms"
' ============================================================
' Rule28_MandatedLegalTermForms.bas
' Proofreading rule: enforces fixed hyphenation for specific
' legal and governmental terms. Flags unhyphenated variants
' and suggests the approved hyphenated form.
'
' Default mandatory list:
'   "Solicitor-General", "Attorney-General"
'
' Additional terms can be added at runtime via AddMandatedTerm.
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
'   - Microsoft Scripting Runtime (Scripting.Dictionary)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "mandated_legal_term_forms"

' ── Module-level dictionary ───────────────────────────────
' Key = LCase(correct form), Value = correct form (String)
Private mandatedTerms As Scripting.Dictionary

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_MandatedLegalTermForms(doc As Document) As Collection
    Dim issues As New Collection

    ' Initialise defaults if not yet loaded
    If mandatedTerms Is Nothing Then
        InitDefaultTerms
    End If

    Dim keys As Variant
    Dim k As Long
    Dim correctForm As String
    Dim searchPhrase As String

    keys = mandatedTerms.keys

    For k = 0 To mandatedTerms.Count - 1
        correctForm = CStr(mandatedTerms(keys(k)))

        ' Build the unhyphenated search variant by replacing hyphens with spaces
        searchPhrase = Replace(correctForm, "-", " ")

        ' Only search if the unhyphenated form is actually different
        If StrComp(searchPhrase, correctForm, vbBinaryCompare) <> 0 Then
            SearchAndFlag doc, searchPhrase, correctForm, issues
        End If
    Next k

    Set Check_MandatedLegalTermForms = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Search for an unhyphenated variant and flag matches
' ════════════════════════════════════════════════════════════
Private Sub SearchAndFlag(doc As Document, _
                           searchPhrase As String, _
                           correctForm As String, _
                           ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = searchPhrase
        .MatchWholeWord = True
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

        ' Skip if the matched text already has the correct hyphenated form
        If StrComp(rng.Text, correctForm, vbTextCompare) = 0 Then
            GoTo SkipMatch
        End If

        ' Verify it is not actually the hyphenated form by checking
        ' the surrounding context — the Find matched with MatchCase=False
        ' and spaces, so an exact binary comparison rules out false positives
        If PleadingsEngine.IsInPageRange(rng) Then
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       "Mandatory term is not hyphenated in the approved form.", _
                       "Use '" & correctForm & "'.", _
                       rng.Start, _
                       rng.End, _
                       "warning", _
                       False
            issues.Add issue
        End If

SkipMatch:
        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Populate default mandatory terms
' ════════════════════════════════════════════════════════════
Private Sub InitDefaultTerms()
    Set mandatedTerms = New Scripting.Dictionary

    mandatedTerms.Add LCase("Solicitor-General"), "Solicitor-General"
    mandatedTerms.Add LCase("Attorney-General"), "Attorney-General"
End Sub

' ════════════════════════════════════════════════════════════
'  PUBLIC: Add a mandated term at runtime
'  The term must contain a hyphen (e.g. "Director-General").
'  If the term already exists it is silently ignored.
' ════════════════════════════════════════════════════════════
Public Sub AddMandatedTerm(term As String)
    If mandatedTerms Is Nothing Then
        InitDefaultTerms
    End If

    Dim lcKey As String
    lcKey = LCase(term)

    If Not mandatedTerms.Exists(lcKey) Then
        mandatedTerms.Add lcKey, term
    End If
End Sub

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunMandatedLegalTermForms()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Mandated Legal Term Forms"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_MandatedLegalTermForms(doc)

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
           vbInformation, "Mandated Legal Term Forms"
End Sub
