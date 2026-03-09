Attribute VB_Name = "Rule12_LicenceLicense"
' ============================================================
' Rule12_LicenceLicense.bas
' Proofreading rule: checks correct UK usage of licence (noun)
' vs license (verb). Also handles compounds and derivatives.
'
' UK convention:
'   licence = noun ("a licence", "the licence holder")
'   license = verb ("to license", "shall license")
'   licensed, licensing = always -s- (verb derivatives)
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "licence_license"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Search a range for licence/license issues
' ════════════════════════════════════════════════════════════
Private Sub SearchForLicenceIssues(searchRange As Range, _
                                    doc As Document, _
                                    ByRef issues As Collection)
    Dim searchTerms As Variant
    Dim t As Long

    ' Search for the base forms; skip derivatives that are always correct
    searchTerms = Array("licence", "license", "sub-licence", "sub-license", _
                        "re-licence", "re-license")

    For t = LBound(searchTerms) To UBound(searchTerms)
        SearchSingleTerm CStr(searchTerms(t)), searchRange, doc, issues
    Next t
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Search for a single term and analyse context
' ════════════════════════════════════════════════════════════
Private Sub SearchSingleTerm(ByVal term As String, _
                              searchRange As Range, _
                              doc As Document, _
                              ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
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
        If Not PleadingsEngine.IsInPageRange(rng) Then
            GoTo ContinueSearch
        End If

        ' Determine if the found word uses -s- or -c-
        usesS = (InStr(1, LCase(rng.Text), "license") > 0)

        ' Skip "licensed" and "licensing" — always correct with -s-
        Dim foundLower As String
        foundLower = LCase(Trim(rng.Text))
        If foundLower = "licensed" Or foundLower = "licensing" Then
            GoTo ContinueSearch
        End If

        ' ── Get surrounding context ──────────────────────────
        contextBefore = GetContextBefore(rng, doc, 50)
        contextAfter = GetContextAfter(rng, doc, 50)

        ' Extract the last word before the match
        wordBefore = GetLastWord(contextBefore)

        ' Extract the first word after the match
        wordAfter = GetFirstWord(contextAfter)

        ' ── Determine noun or verb context ───────────────────
        baseIsVerb = IsVerbIndicator(wordBefore)
        baseIsNoun = IsNounIndicator(wordBefore) Or IsNounFollower(wordAfter)

        ' ── Decide if there is an issue ──────────────────────
        issueText = ""
        suggestion = ""

        If usesS And baseIsNoun And Not baseIsVerb Then
            ' "license" used in noun context — should be "licence"
            issueText = "'" & rng.Text & "' appears in a noun context; " & _
                        "UK convention uses 'licence' for the noun"
            suggestion = ReplaceSWithC(rng.Text)
        ElseIf Not usesS And baseIsVerb And Not baseIsNoun Then
            ' "licence" used in verb context — should be "license"
            issueText = "'" & rng.Text & "' appears in a verb context; " & _
                        "UK convention uses 'license' for the verb"
            suggestion = ReplaceCWithS(rng.Text)
        ElseIf (usesS And Not baseIsVerb And Not baseIsNoun) Or _
               (Not usesS And Not baseIsVerb And Not baseIsNoun) Then
            ' Context ambiguous
            issueText = "'" & rng.Text & "' — unable to determine noun/verb context; " & _
                        "review context to ensure correct UK spelling"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present — ambiguous
            issueText = "'" & rng.Text & "' — conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf Not usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present — ambiguous
            issueText = "'" & rng.Text & "' — conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        End If

        ' Only create issue if we found something to flag
        If Len(issueText) > 0 Then
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       issueText, _
                       suggestion, _
                       rng.Start, _
                       rng.End, _
                       "possible_error"
            issues.Add issue
        End If

ContinueSearch:
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Get text before the match range (up to N chars)
' ════════════════════════════════════════════════════════════
Private Function GetContextBefore(rng As Range, doc As Document, _
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
        GetContextBefore = ""
        Exit Function
    End If
    On Error GoTo 0

    GetContextBefore = contextRng.Text
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Get text after the match range (up to N chars)
' ════════════════════════════════════════════════════════════
Private Function GetContextAfter(rng As Range, doc As Document, _
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
        GetContextAfter = ""
        Exit Function
    End If
    On Error GoTo 0

    GetContextAfter = contextRng.Text
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Extract the last word from a context string
' ════════════════════════════════════════════════════════════
Private Function GetLastWord(ByVal text As String) As String
    Dim trimmed As String
    Dim i As Long
    Dim ch As String

    trimmed = Trim(text)
    If Len(trimmed) = 0 Then
        GetLastWord = ""
        Exit Function
    End If

    ' Walk backward from end to find last word boundary
    For i = Len(trimmed) To 1 Step -1
        ch = Mid(trimmed, i, 1)
        If ch = " " Or ch = vbCr Or ch = vbLf Or ch = vbTab Then
            GetLastWord = LCase(Mid(trimmed, i + 1))
            Exit Function
        End If
    Next i

    GetLastWord = LCase(trimmed)
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Extract the first word from a context string
' ════════════════════════════════════════════════════════════
Private Function GetFirstWord(ByVal text As String) As String
    Dim trimmed As String
    Dim spacePos As Long

    trimmed = Trim(text)
    If Len(trimmed) = 0 Then
        GetFirstWord = ""
        Exit Function
    End If

    spacePos = InStr(1, trimmed, " ")
    If spacePos > 0 Then
        GetFirstWord = LCase(Left(trimmed, spacePos - 1))
    Else
        GetFirstWord = LCase(trimmed)
    End If

    ' Strip trailing punctuation
    Dim result As String
    Dim ch As String
    result = GetFirstWord
    Do While Len(result) > 0
        ch = Right(result, 1)
        If ch Like "[A-Za-z]" Then Exit Do
        result = Left(result, Len(result) - 1)
    Loop
    GetFirstWord = result
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a word is a verb indicator
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a word is a noun indicator
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if the word after indicates noun usage
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Replace -s- with -c- in licence/license words
' ════════════════════════════════════════════════════════════
Private Function ReplaceSWithC(ByVal word As String) As String
    ReplaceSWithC = Replace(word, "license", "licence", , , vbTextCompare)
    ReplaceSWithC = Replace(ReplaceSWithC, "License", "Licence", , , vbBinaryCompare)
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Replace -c- with -s- in licence/license words
' ════════════════════════════════════════════════════════════
Private Function ReplaceCWithS(ByVal word As String) As String
    ReplaceCWithS = Replace(word, "licence", "license", , , vbTextCompare)
    ReplaceCWithS = Replace(ReplaceCWithS, "Licence", "License", , , vbBinaryCompare)
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunLicenceLicense()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Licence License"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_LicenceLicense(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Licence License"
End Sub
