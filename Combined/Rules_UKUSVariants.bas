Attribute VB_Name = "Rules_UKUSVariants"
' ============================================================
' Rules_UKUSVariants.bas
' Combined proofreading rules for UK/US English variants.
'
' Rule 12 — Licence/License:
'   Checks correct UK usage of licence (noun) vs license (verb).
'   Also handles compounds and derivatives.
'   UK convention:
'     licence = noun ("a licence", "the licence holder")
'     license = verb ("to license", "shall license")
'     licensed, licensing = always -s- (verb derivatives)
'
' Rule 13 — Colour Formatting:
'   Detects non-standard font colours in the document body.
'   Identifies the dominant text colour and flags any runs
'   using a different colour (excluding hyperlinks and
'   heading-styled paragraphs).
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
'   - Microsoft Scripting Runtime (Scripting.Dictionary)
' ============================================================
Option Explicit

Private Const RULE_NAME_LICENCE As String = "licence_license"
Private Const RULE_NAME_COLOUR As String = "colour_formatting"

' ================================================================
' ================================================================
'  RULE 12 — LICENCE / LICENSE
' ================================================================
' ================================================================

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT — Licence/License
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
            issue.Init RULE_NAME_LICENCE, _
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

' ================================================================
' ================================================================
'  RULE 13 — COLOUR FORMATTING
' ================================================================
' ================================================================

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT — Colour Formatting
' ════════════════════════════════════════════════════════════
Public Function Check_ColourFormatting(doc As Document) As Collection
    Dim issues As New Collection
    Dim colourCounts As Object ' Scripting.Dictionary
    Dim para As Paragraph
    Dim rn As Range
    Dim runColor As Long
    Dim dominantColour As Long
    Dim maxCount As Long
    Dim runText As String

    ' ── First pass: count colour usage per run ───────────────
    Set colourCounts = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Dim paraRange As Range
        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass1
        End If

        ' Skip paragraphs outside page range
        If Not PleadingsEngine.IsInPageRange(paraRange) Then
            GoTo NextParaPass1
        End If

        ' Iterate runs within the paragraph
        Dim r As Long
        For r = 1 To paraRange.Runs.Count
            Err.Clear
            Set rn = paraRange.Runs(r)
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextRunPass1
            End If

            ' Skip whitespace-only runs
            runText = rn.Text
            If Len(Trim(Replace(Replace(runText, vbCr, ""), vbLf, ""))) = 0 Then
                GoTo NextRunPass1
            End If

            runColor = rn.Font.Color
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextRunPass1
            End If

            If colourCounts.Exists(runColor) Then
                colourCounts(runColor) = colourCounts(runColor) + 1
            Else
                colourCounts.Add runColor, 1
            End If

NextRunPass1:
        Next r

NextParaPass1:
    Next para
    On Error GoTo 0

    ' ── Determine dominant colour ────────────────────────────
    If colourCounts.Count = 0 Then
        Set Check_ColourFormatting = issues
        Exit Function
    End If

    dominantColour = 0
    maxCount = 0
    Dim colourKey As Variant
    For Each colourKey In colourCounts.keys
        If colourCounts(colourKey) > maxCount Then
            maxCount = colourCounts(colourKey)
            dominantColour = CLng(colourKey)
        End If
    Next colourKey

    ' ── Second pass: flag non-dominant, non-automatic colours ─
    Const WD_COLOR_AUTOMATIC As Long = -16777216

    ' Tracking for grouping consecutive same-colour runs
    Dim groupStartPos As Long
    Dim groupEndPos As Long
    Dim groupColour As Long
    Dim groupActive As Boolean
    Dim groupParaRange As Range

    groupActive = False

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaPass2
        End If

        ' Skip paragraphs outside page range
        If Not PleadingsEngine.IsInPageRange(paraRange) Then
            ' Flush any active group before skipping
            If groupActive Then
                FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                groupActive = False
            End If
            GoTo NextParaPass2
        End If

        ' Skip heading-styled paragraphs (may have intentional colour)
        Dim styleName As String
        styleName = ""
        styleName = para.Style.NameLocal
        If Err.Number <> 0 Then
            Err.Clear
            styleName = ""
        End If
        If LCase(Left(styleName, 7)) = "heading" Then
            If groupActive Then
                FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                groupActive = False
            End If
            GoTo NextParaPass2
        End If

        ' Iterate runs
        For r = 1 To paraRange.Runs.Count
            Err.Clear
            Set rn = paraRange.Runs(r)
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextRunPass2
            End If

            ' Skip whitespace-only runs
            runText = rn.Text
            If Len(Trim(Replace(Replace(runText, vbCr, ""), vbLf, ""))) = 0 Then
                GoTo NextRunPass2
            End If

            runColor = rn.Font.Color
            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextRunPass2
            End If

            ' Skip if colour matches dominant or is automatic
            If runColor = dominantColour Or runColor = WD_COLOR_AUTOMATIC Then
                ' Flush any active group
                If groupActive Then
                    FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                    groupActive = False
                End If
                GoTo NextRunPass2
            End If

            ' Skip hyperlinks
            If IsRunInsideHyperlink(rn, doc) Then
                If groupActive Then
                    FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                    groupActive = False
                End If
                GoTo NextRunPass2
            End If

            ' ── This run has a non-standard colour ───────────
            If groupActive And runColor = groupColour And _
               rn.Start = groupEndPos Then
                ' Extend existing group
                groupEndPos = rn.End
            Else
                ' Flush previous group if any
                If groupActive Then
                    FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
                End If
                ' Start new group
                groupStartPos = rn.Start
                groupEndPos = rn.End
                groupColour = runColor
                groupActive = True
            End If

NextRunPass2:
        Next r

NextParaPass2:
    Next para

    ' Flush final group
    If groupActive Then
        FlushColourGroup doc, issues, groupStartPos, groupEndPos, groupColour
    End If
    On Error GoTo 0

    Set Check_ColourFormatting = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Flush a grouped colour issue
' ════════════════════════════════════════════════════════════
Private Sub FlushColourGroup(doc As Document, _
                              ByRef issues As Collection, _
                              ByVal startPos As Long, _
                              ByVal endPos As Long, _
                              ByVal fontColor As Long)
    Dim issue As PleadingsIssue
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

    locStr = PleadingsEngine.GetLocationString(rng, doc)
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

    Set issue = New PleadingsIssue
    issue.Init RULE_NAME_COLOUR, _
               locStr, _
               "Non-standard font colour " & hexStr & " detected: '" & _
               previewText & "'", _
               "Change font colour to match document default", _
               startPos, _
               endPos, _
               "possible_error"
    issues.Add issue
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Convert a Long colour value to hex string
' ════════════════════════════════════════════════════════════
Private Function ColourToHex(ByVal colorVal As Long) As String
    Dim R As Long
    Dim G As Long
    Dim B As Long

    ' Word stores colours as BGR in Long format
    R = colorVal Mod 256
    G = (colorVal \ 256) Mod 256
    B = (colorVal \ 65536) Mod 256

    ColourToHex = "#" & Right("0" & Hex(R), 2) & _
                        Right("0" & Hex(G), 2) & _
                        Right("0" & Hex(B), 2)
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a run is inside a hyperlink
' ════════════════════════════════════════════════════════════
Private Function IsRunInsideHyperlink(rn As Range, doc As Document) As Boolean
    Dim hl As Hyperlink

    On Error Resume Next
    For Each hl In doc.Hyperlinks
        Err.Clear
        If hl.Range.Start <= rn.Start And hl.Range.End >= rn.End Then
            IsRunInsideHyperlink = True
            Exit Function
        End If
        If Err.Number <> 0 Then
            Err.Clear
        End If
    Next hl
    On Error GoTo 0

    IsRunInsideHyperlink = False
End Function
