Attribute VB_Name = "Rules_Punctuation"
' ============================================================
' Rules_Punctuation.bas
' Combined proofreading rules for punctuation checking:
'   - Slash style (Rule14): checks forward slash spacing
'     consistency and flags unexpected backslashes.
'   - Bracket integrity (Rule16): checks for mismatched,
'     unmatched, and improperly nested brackets: (), [], {}.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_SLASH As String = "slash_style"
Private Const RULE_NAME_BRACKET As String = "bracket_integrity"
Private Const RULE_NAME_DASH As String = "hyphens"
Private Const RULE_NAME_TRIPLICATE As String = "triplicate_punctuation"

' ?==============================================================?
' ?  SLASH STYLE (Rule14)                                       ?
' ?==============================================================?

' ============================================================
'  MAIN ENTRY POINT: Slash Style
' ============================================================
Public Function Check_SlashStyle(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Forward slashes: determine dominant style ------------
    Dim tightCount As Long
    Dim spacedCount As Long

    tightCount = CountTightSlashes(doc)
    spacedCount = CountSpacedSlashes(doc)

    ' Determine dominant style
    Dim dominantStyle As String
    If tightCount >= spacedCount Then
        dominantStyle = "tight"
    Else
        dominantStyle = "spaced"
    End If

    ' Flag minority style forward slashes
    If dominantStyle = "tight" And spacedCount > 0 Then
        FlagSpacedSlashes doc, issues
    ElseIf dominantStyle = "spaced" And tightCount > 0 Then
        FlagTightSlashes doc, issues
    End If

    ' -- Backslashes ------------------------------------------
    FlagBackslashes doc, issues

    Set Check_SlashStyle = issues
End Function

' ============================================================
'  PRIVATE: Count tight slashes using wildcard search
'  Excludes conventional tight pairs (and/or, his/her, etc.)
'  so they don't bias the dominant-style determination.
' ============================================================
Private Function CountTightSlashes(doc As Document) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean

    cnt = 0
    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = "[! ]/[! ]"
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        ' Skip URLs and dates
        If Not IsURLContext(rng, doc) And Not IsDateSlash(rng) Then
            ' Skip conventional tight pairs (and/or, his/her, etc.)
            If Not IsConventionalTightSlash(rng, doc) Then
                cnt = cnt + 1
            End If
        End If

        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0

    CountTightSlashes = cnt
End Function

' ============================================================
'  PRIVATE: Count spaced slashes using literal search
' ============================================================
Private Function CountSpacedSlashes(doc As Document) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean

    cnt = 0
    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = " / "
        .MatchWildcards = False
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        ' Skip URLs
        If Not IsURLContext(rng, doc) Then
            cnt = cnt + 1
        End If

        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0

    CountSpacedSlashes = cnt
End Function

' ============================================================
'  PRIVATE: Flag spaced slashes (minority when tight is dominant)
' ============================================================
Private Sub FlagSpacedSlashes(doc As Document, ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = " / "
        .MatchWildcards = False
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        If Not EngineIsInPageRange(rng) Then GoTo ContinueSpaced
        If IsURLContext(rng, doc) Then GoTo ContinueSpaced

        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set finding = CreateIssueDict(RULE_NAME_SLASH, locStr, "Spaced slash '" & rng.Text & "' differs from dominant tight style", "Remove spaces around slash for consistency", rng.Start, rng.End, "possible_error")
        issues.Add finding

ContinueSpaced:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Flag tight slashes (minority when spaced is dominant)
' ============================================================
Private Sub FlagTightSlashes(doc As Document, ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = "[! ]/[! ]"
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos2 As Long
    lastPos2 = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos2 Then Exit Do   ' stall guard
        lastPos2 = rng.Start

        If Not EngineIsInPageRange(rng) Then GoTo ContinueTight
        If IsURLContext(rng, doc) Then GoTo ContinueTight
        If IsDateSlash(rng) Then GoTo ContinueTight
        If IsConventionalTightSlash(rng, doc) Then GoTo ContinueTight

        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set finding = CreateIssueDict(RULE_NAME_SLASH, locStr, "Tight slash '" & rng.Text & "' differs from dominant spaced style", "Add spaces around slash for consistency", rng.Start, rng.End, "possible_error")
        issues.Add finding

ContinueTight:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Flag unexpected backslashes
' ============================================================
Private Sub FlagBackslashes(doc As Document, ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String
    Dim context As String
    Dim fontName As String

    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = "\"
        .MatchWildcards = False
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos3 As Long
    lastPos3 = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos3 Then Exit Do   ' stall guard
        lastPos3 = rng.Start

        If Not EngineIsInPageRange(rng) Then GoTo ContinueBackslash

        ' Get surrounding context for skip checks
        Dim contextStart As Long
        Dim contextEnd As Long
        Dim contextRng As Range

        contextStart = rng.Start - 5
        If contextStart < 0 Then contextStart = 0
        contextEnd = rng.End + 10
        If contextEnd > doc.Content.End Then contextEnd = doc.Content.End

        Set contextRng = doc.Range(contextStart, contextEnd)
        If Err.Number <> 0 Then
            Err.Clear
            context = ""
        Else
            context = LCase(contextRng.Text)
        End If

        ' Skip file paths: drive letter pattern like C:\ or UNC \\server
        If IsDriveLetterPath(context) Or IsUNCPath(context) Then
            GoTo ContinueBackslash
        End If

        ' Skip code-font runs (Courier, Consolas)
        fontName = ""
        fontName = rng.Font.Name
        If Err.Number <> 0 Then
            Err.Clear
            fontName = ""
        End If
        If IsCodeFontName(fontName) Then
            GoTo ContinueBackslash
        End If

        ' Skip URLs
        If InStr(1, context, "://") > 0 Then
            GoTo ContinueBackslash
        End If

        ' Flag the backslash
        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set finding = CreateIssueDict(RULE_NAME_SLASH, locStr, "Unexpected backslash -- did you mean forward slash?", "Replace '\' with '/'", rng.Start, rng.End, "possible_error")
        issues.Add finding

ContinueBackslash:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Check if a tight slash is a conventional pair
'  (and/or, his/her, etc.) that should always be tight
'  regardless of the document's dominant slash style.
' ============================================================
Private Function IsConventionalTightSlash(rng As Range, doc As Document) As Boolean
    IsConventionalTightSlash = False
    On Error Resume Next

    ' Expand range to capture surrounding word context
    Dim ctxStart As Long, ctxEnd As Long
    ctxStart = rng.Start - 12
    If ctxStart < 0 Then ctxStart = 0
    ctxEnd = rng.End + 12
    If ctxEnd > doc.Content.End Then ctxEnd = doc.Content.End

    Dim ctxRng As Range
    Set ctxRng = doc.Range(ctxStart, ctxEnd)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function

    Dim ctxText As String
    ctxText = LCase$(ctxRng.Text)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0

    ' Known conventionally-tight slash pairs
    Dim tightPairs As Variant
    tightPairs = Array("and/or", "either/or", "his/her", "he/she", _
                       "s/he", "w/o", "n/a", "c/o", "a/c", "y/n", "yes/no", _
                       "input/output", "read/write", "true/false", "on/off", _
                       "open/close", "start/end", "pass/fail", "client/server")
    Dim tp As Variant
    For Each tp In tightPairs
        If InStr(1, ctxText, CStr(tp), vbTextCompare) > 0 Then
            IsConventionalTightSlash = True
            Exit Function
        End If
    Next tp

    ' General word/word alternative detection:
    ' If the match text is letter(s)/letter(s) and both sides are short
    ' English words (2-12 chars), treat as a functional alternative pair.
    Dim matchText As String
    matchText = LCase$(rng.Text)
    Dim slashPos As Long
    slashPos = InStr(1, matchText, "/")
    If slashPos > 1 And slashPos < Len(matchText) Then
        Dim lWord As String, rWord As String
        lWord = Left$(matchText, slashPos - 1)
        rWord = Mid$(matchText, slashPos + 1)
        ' Both sides are purely alphabetic and short
        If Len(lWord) >= 2 And Len(lWord) <= 12 And _
           Len(rWord) >= 2 And Len(rWord) <= 12 Then
            If IsAlphaOnly(lWord) And IsAlphaOnly(rWord) Then
                IsConventionalTightSlash = True
                Exit Function
            End If
        End If
    End If
End Function

' Helper: check if a string is purely alphabetic (a-z)
Private Function IsAlphaOnly(ByVal s As String) As Boolean
    Dim ci As Long
    For ci = 1 To Len(s)
        Dim cc As String
        cc = Mid$(s, ci, 1)
        If Not ((cc >= "a" And cc <= "z") Or (cc >= "A" And cc <= "Z")) Then
            IsAlphaOnly = False
            Exit Function
        End If
    Next ci
    IsAlphaOnly = True
End Function

' ============================================================
'  PRIVATE: Check if context suggests a URL
' ============================================================
Private Function IsURLContext(rng As Range, doc As Document) As Boolean
    Dim contextStart As Long
    Dim contextEnd As Long
    Dim contextRng As Range
    Dim context As String

    On Error Resume Next
    contextStart = rng.Start - 30
    If contextStart < 0 Then contextStart = 0
    contextEnd = rng.End + 30
    If contextEnd > doc.Content.End Then contextEnd = doc.Content.End

    Set contextRng = doc.Range(contextStart, contextEnd)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsURLContext = False
        Exit Function
    End If

    context = LCase(contextRng.Text)
    On Error GoTo 0

    IsURLContext = (InStr(1, context, "://") > 0) Or _
                   (InStr(1, context, "http") > 0) Or _
                   (InStr(1, context, "www") > 0)
End Function

' ============================================================
'  PRIVATE: Check if slash is part of a date (digits only)
' ============================================================
Private Function IsDateSlash(rng As Range) As Boolean
    Dim matchText As String
    Dim i As Long
    Dim ch As String
    Dim hasSlash As Boolean

    matchText = rng.Text
    If Len(matchText) < 3 Then
        IsDateSlash = False
        Exit Function
    End If

    ' Check that all non-slash characters are digits
    hasSlash = False
    For i = 1 To Len(matchText)
        ch = Mid(matchText, i, 1)
        If ch = "/" Then
            hasSlash = True
        ElseIf Not (ch >= "0" And ch <= "9") Then
            IsDateSlash = False
            Exit Function
        End If
    Next i

    IsDateSlash = hasSlash
End Function

' ============================================================
'  PRIVATE: Check for drive letter path pattern (e.g. C:\)
' ============================================================
Private Function IsDriveLetterPath(ByVal context As String) As Boolean
    Dim i As Long
    Dim ch As String

    ' Look for pattern: letter followed by :\
    For i = 1 To Len(context) - 2
        ch = Mid(context, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Then
            If Mid(context, i + 1, 2) = ":\" Then
                IsDriveLetterPath = True
                Exit Function
            End If
        End If
    Next i

    IsDriveLetterPath = False
End Function

' ============================================================
'  PRIVATE: Check for UNC path pattern (\\server)
' ============================================================
Private Function IsUNCPath(ByVal context As String) As Boolean
    IsUNCPath = (InStr(1, context, "\\") > 0)
End Function

' ?==============================================================?
' ?  BRACKET INTEGRITY (Rule16)                                 ?
' ?==============================================================?

' ============================================================
'  MAIN ENTRY POINT: Bracket Integrity
' ============================================================
Public Function Check_BracketIntegrity(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim paraStart As Long

    ' Counters per bracket type (reset per paragraph)
    Dim parenOpen As Long, parenClose As Long
    Dim sqOpen As Long, sqClose As Long
    Dim curlyOpen As Long, curlyClose As Long

    ' Position of first unmatched bracket (for issue location)
    Dim firstParenPos As Long, firstSqPos As Long, firstCurlyPos As Long

    Dim b() As Byte, bMax As Long
    Dim i As Long, code As Long, pos As Long

    For Each para In doc.Paragraphs
        On Error Resume Next
        paraText = para.Range.Text
        paraStart = para.Range.Start
        If Err.Number <> 0 Then
            Err.Clear: On Error GoTo 0
            GoTo NxtPara
        End If
        On Error GoTo 0

        If LenB(paraText) = 0 Then GoTo NxtPara

        ' Compute list prefix length for position correction
        Dim bktListPrefixLen As Long
        bktListPrefixLen = GetDashListPrefixLen(para, paraText)

        ' Reset counters
        parenOpen = 0: parenClose = 0
        sqOpen = 0: sqClose = 0
        curlyOpen = 0: curlyClose = 0
        firstParenPos = -1: firstSqPos = -1: firstCurlyPos = -1

        b = paraText
        bMax = UBound(b) - 1

        For i = 0 To bMax Step 2
            code = b(i) Or (CLng(b(i + 1)) * 256&)
            pos = paraStart + (i \ 2) - bktListPrefixLen

            Select Case code
                Case 40   ' (
                    parenOpen = parenOpen + 1
                    If firstParenPos < 0 Then firstParenPos = pos
                Case 41   ' )
                    parenClose = parenClose + 1
                    If firstParenPos < 0 Then firstParenPos = pos
                Case 91   ' [
                    sqOpen = sqOpen + 1
                    If firstSqPos < 0 Then firstSqPos = pos
                Case 93   ' ]
                    sqClose = sqClose + 1
                    If firstSqPos < 0 Then firstSqPos = pos
                Case 123  ' {
                    curlyOpen = curlyOpen + 1
                    If firstCurlyPos < 0 Then firstCurlyPos = pos
                Case 125  ' }
                    curlyClose = curlyClose + 1
                    If firstCurlyPos < 0 Then firstCurlyPos = pos
            End Select
        Next i

        ' Report once per bracket type if counts don't match
        If parenOpen <> parenClose Then
            CreateBracketIssue doc, issues, firstParenPos, "()", _
                "Unbalanced parentheses: " & parenOpen & " opened, " & _
                parenClose & " closed"
        End If
        If sqOpen <> sqClose Then
            CreateBracketIssue doc, issues, firstSqPos, "[]", _
                "Unbalanced square brackets: " & sqOpen & " opened, " & _
                sqClose & " closed"
        End If
        If curlyOpen <> curlyClose Then
            CreateBracketIssue doc, issues, firstCurlyPos, "{}", _
                "Unbalanced curly braces: " & curlyOpen & " opened, " & _
                curlyClose & " closed"
        End If

        ' -- Stack-based nesting check (only when counts balance) --
        If parenOpen = parenClose And sqOpen = sqClose _
           And curlyOpen = curlyClose _
           And (parenOpen + sqOpen + curlyOpen) > 0 Then
            Dim stk() As Long, stkTop As Long
            stkTop = 0
            ReDim stk(1 To parenOpen + sqOpen + curlyOpen)
            Dim nestBad As Boolean, nestPos As Long
            nestBad = False
            For i = 0 To bMax Step 2
                code = b(i) Or (CLng(b(i + 1)) * 256&)
                Select Case code
                    Case 40, 91, 123  ' open bracket
                        stkTop = stkTop + 1
                        If stkTop > UBound(stk) Then ReDim Preserve stk(1 To stkTop + 4)
                        stk(stkTop) = code
                    Case 41, 93, 125  ' close bracket
                        If stkTop = 0 Then
                            nestBad = True
                            nestPos = paraStart + (i \ 2) - bktListPrefixLen
                            Exit For
                        End If
                        If Not CodesMatch(stk(stkTop), code) Then
                            nestBad = True
                            nestPos = paraStart + (i \ 2) - bktListPrefixLen
                            Exit For
                        End If
                        stkTop = stkTop - 1
                End Select
            Next i
            If nestBad Then
                CreateBracketIssue doc, issues, nestPos, "()", _
                    "Improperly nested brackets (e.g. overlapping pairs)"
            End If
        End If

NxtPara:
    Next para

    Set Check_BracketIntegrity = issues
End Function

' ------------------------------------------------------------
'  Code-point bracket matching (no string comparison)
' ------------------------------------------------------------
Private Function CodesMatch(ByVal openCode As Long, _
        ByVal closeCode As Long) As Boolean
    Select Case openCode
        Case 40:  CodesMatch = (closeCode = 41)   ' ( -> )
        Case 91:  CodesMatch = (closeCode = 93)   ' [ -> ]
        Case 123: CodesMatch = (closeCode = 125)  ' { -> }
        Case Else: CodesMatch = False
    End Select
End Function

' ============================================================
'  PRIVATE: Create a bracket integrity finding
' ============================================================
Private Sub CreateBracketIssue(doc As Document, _
                                ByRef issues As Collection, _
                                ByVal pos As Long, _
                                ByVal bracketChar As String, _
                                ByVal issueText As String)
    Dim finding As Object
    Dim locStr As String
    Dim rng As Range

    On Error Resume Next
    Set rng = doc.Range(pos, pos + 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    ' Skip if outside page range
    If Not EngineIsInPageRange(rng) Then
        On Error GoTo 0
        Exit Sub
    End If

    locStr = EngineGetLocationString(rng, doc)
    If Err.Number <> 0 Then
        locStr = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0

    ' Determine suggestion based on bracket type
    Dim suggestion As String
    Select Case bracketChar
        Case "()", "(", ")"
            suggestion = "Add or correct matching parenthesis"
        Case "[]", "[", "]"
            suggestion = "Add or correct matching square bracket"
        Case "{}", "{", "}"
            suggestion = "Add or correct matching curly brace"
        Case Else
            suggestion = "Review bracket pairing"
    End Select

    Set finding = CreateIssueDict(RULE_NAME_BRACKET, locStr, issueText, suggestion, pos, pos + 1, "error")
    issues.Add finding
End Sub

' ?==============================================================?
' ?  SHARED PRIVATE HELPERS                                     ?
' ?==============================================================?

' ============================================================
'  PRIVATE: Check if a font name is a code font (Courier, Consolas)
'  Shared by FlagBackslashes and IsCodeFont
' ============================================================
Private Function IsCodeFontName(ByVal fontName As String) As Boolean
    IsCodeFontName = (LCase(fontName) Like "*courier*") Or _
                     (LCase(fontName) Like "*consolas*")
End Function


' Calculate the length of auto-generated list numbering text
Private Function GetDashListPrefixLen(para As Paragraph, ByVal paraText As String) As Long
    GetDashListPrefixLen = 0
    On Error Resume Next
    Dim lStr As String
    lStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    If Len(lStr) = 0 Then On Error GoTo 0: Exit Function
    If Len(paraText) > Len(lStr) Then
        If Left$(paraText, Len(lStr)) = lStr Then
            GetDashListPrefixLen = Len(lStr)
            If Mid$(paraText, GetDashListPrefixLen + 1, 1) = vbTab Then
                GetDashListPrefixLen = GetDashListPrefixLen + 1
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
' ============================================================
'  TRIPLICATE PUNCTUATION
'  Flags three or more consecutive identical punctuation marks:
'    (((  )))  [[[  ]]]  """  '''  ,,,
'  Deliberately excludes "..." (ellipsis).
' ============================================================
Public Function Check_TriplicatePunctuation(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim finding As Object
    Dim locStr As String
    Dim i As Long
    Dim ch As String
    Dim runLen As Long

    ' Characters to check (NOT including "." to avoid flagging ellipsis)
    Dim targets As String
    targets = "()[]""',"

    On Error Resume Next
    For Each para In doc.Paragraphs
        paraText = para.Range.Text
        If Err.Number <> 0 Then paraText = "": Err.Clear

        If Len(paraText) < 3 Then GoTo NextTriPara

        ' Check page range
        If Not EngineIsInPageRange(para.Range) Then GoTo NextTriPara

        i = 1
        Do While i <= Len(paraText) - 2
            ch = Mid$(paraText, i, 1)
            If InStr(targets, ch) > 0 Then
                ' Count consecutive identical chars
                runLen = 1
                Do While i + runLen <= Len(paraText) And Mid$(paraText, i + runLen, 1) = ch
                    runLen = runLen + 1
                Loop
                If runLen >= 3 Then
                    locStr = EngineGetLocationString(para.Range, doc)
                    Dim matched As String
                    matched = String$(runLen, ch)
                    Set finding = CreateIssueDict( _
                        RULE_NAME_TRIPLICATE, locStr, _
                        "Triplicate punctuation: '" & matched & "'", _
                        "Remove repeated punctuation", _
                        para.Range.Start + i - 1, _
                        para.Range.Start + i - 1 + runLen, _
                        "error", False, "", matched)
                    issues.Add finding
                    i = i + runLen
                Else
                    i = i + runLen
                End If
            Else
                i = i + 1
            End If
        Loop
NextTriPara:
    Next para
    On Error GoTo 0

    Set Check_TriplicatePunctuation = issues
End Function

Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False, _
                                 Optional ByVal replacementText_ As String = "", _
                                 Optional ByVal matchedText_ As String = "", _
                                 Optional ByVal anchorKind_ As String = "exact_text", _
                                 Optional ByVal confidenceLabel_ As String = "high", _
                                 Optional ByVal sourceParagraphIndex_ As Long = 0) As Object
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
    If autoFixSafe_ Then d("ReplacementText") = replacementText_
    d("MatchedText") = matchedText_
    d("AnchorKind") = anchorKind_
    d("ConfidenceLabel") = confidenceLabel_
    d("SourceParagraphIndex") = sourceParagraphIndex_
    Set CreateIssueDict = d
End Function


' ================================================================
' ================================================================
'  DASH USAGE (en-dash / em-dash / hyphen)
'  Context-dependent checks:
'   1. Hyphen in number ranges -> should be en-dash
'   2. Double-hyphen "--" -> should be em-dash
'   3. En-dash between words (compound) -> should be hyphen
'   4. Spaced en-dash -> should probably be em-dash
' ================================================================
' ================================================================

Public Function Check_DashUsage(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim paraRange As Range
    Dim finding As Object
    Dim locStr As String

    Dim reHyphenRange As Object
    Set reHyphenRange = CreateObject("VBScript.RegExp")
    reHyphenRange.Global = True
    ' Matches digit(s) - hyphen - digit(s) as a number range
    reHyphenRange.Pattern = "(\d)-(\d)"

    Dim reDoubleHyphen As Object
    Set reDoubleHyphen = CreateObject("VBScript.RegExp")
    reDoubleHyphen.Global = True
    reDoubleHyphen.Pattern = "--"

    Dim enDash As String
    enDash = ChrW(8211)
    Dim emDash As String
    emDash = ChrW(8212)

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDash

        If Not EngineIsInPageRange(paraRange) Then GoTo NextParaDash

        paraText = paraRange.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDash
        ' Strip para mark
        If Len(paraText) > 0 Then
            If Right$(paraText, 1) = vbCr Or Right$(paraText, 1) = Chr(13) Then
                paraText = Left$(paraText, Len(paraText) - 1)
            End If
        End If
        If Len(paraText) < 2 Then GoTo NextParaDash

        ' Calculate auto-number prefix offset
        Dim dashListPrefixLen As Long
        dashListPrefixLen = GetDashListPrefixLen(para, paraText)

        ' --- Check 1: Hyphen in number ranges (digit-digit) ---
        Dim mHR As Object
        Set mHR = reHyphenRange.Execute(paraText)
        Dim hm As Object
        For Each hm In mHR
            Dim hrStart As Long
            hrStart = paraRange.Start + hm.FirstIndex - dashListPrefixLen
            ' The hyphen is at offset +length_of_first_digit
            ' In pattern (\d)-(\d), hyphen is at FirstIndex + 1
            Dim hyphenPos As Long
            hyphenPos = hrStart + 1
            Dim hrEnd As Long
            hrEnd = hyphenPos + 1  ' just the hyphen

            Err.Clear
            Dim hrRng As Range
            Set hrRng = doc.Range(hyphenPos, hrEnd)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            Else
                locStr = EngineGetLocationString(hrRng, doc)
            End If

            Set finding = CreateIssueDict(RULE_NAME_DASH, locStr, _
                "Hyphen used in number range. Use an en-dash (" & enDash & ") for ranges.", _
                "Replace hyphen with en-dash", hyphenPos, hrEnd, "error", True, enDash)
            issues.Add finding
        Next hm

        ' --- Check 2: Double-hyphen "--" should be em-dash ---
        Dim mDH As Object
        Set mDH = reDoubleHyphen.Execute(paraText)
        Dim dhm As Object
        For Each dhm In mDH
            Dim dhStart As Long
            dhStart = paraRange.Start + dhm.FirstIndex - dashListPrefixLen
            Dim dhEnd As Long
            dhEnd = dhStart + 2

            Err.Clear
            Dim dhRng As Range
            Set dhRng = doc.Range(dhStart, dhEnd)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            Else
                locStr = EngineGetLocationString(dhRng, doc)
            End If

            Set finding = CreateIssueDict(RULE_NAME_DASH, locStr, _
                "Double-hyphen found. Use an em-dash (" & emDash & ") instead.", _
                "Replace with em-dash", dhStart, dhEnd, "error", True, emDash)
            issues.Add finding
        Next dhm

        ' --- Check 3: En-dash between letters (compound word) ---
        ' Pattern: letter + en-dash + letter (no spaces) = should be hyphen
        Dim enPos As Long
        enPos = InStr(1, paraText, enDash)
        Do While enPos > 0
            If enPos > 1 And enPos < Len(paraText) Then
                Dim chBefore As String
                Dim chAfter As String
                chBefore = Mid$(paraText, enPos - 1, 1)
                chAfter = Mid$(paraText, enPos + 1, 1)

                Dim beforeIsLetter As Boolean
                Dim afterIsLetter As Boolean
                beforeIsLetter = (chBefore >= "A" And chBefore <= "Z") Or _
                                 (chBefore >= "a" And chBefore <= "z")
                afterIsLetter = (chAfter >= "A" And chAfter <= "Z") Or _
                                (chAfter >= "a" And chAfter <= "z")

                If beforeIsLetter And afterIsLetter Then
                    Dim enStart As Long
                    enStart = paraRange.Start + enPos - 1 - dashListPrefixLen
                    Dim enEnd As Long
                    enEnd = enStart + 1

                    Err.Clear
                    Dim enRng As Range
                    Set enRng = doc.Range(enStart, enEnd)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    Else
                        locStr = EngineGetLocationString(enRng, doc)
                    End If

                    Set finding = CreateIssueDict(RULE_NAME_DASH, locStr, _
                        "En-dash (" & enDash & ") used between words. Use a hyphen (-) for compound words.", _
                        "Replace en-dash with hyphen", enStart, enEnd, "error", True, "-")
                    issues.Add finding
                End If

                ' Check 4: Spaced en-dash (" – ") -> should be em-dash (" — ")
                ' Exception: spaced en-dash between numbers is correct for ranges
                Dim beforeIsSpace As Boolean
                Dim afterIsSpace As Boolean
                beforeIsSpace = (chBefore = " ")
                afterIsSpace = (chAfter = " ")

                If beforeIsSpace And afterIsSpace Then
                    ' Check if this is a number range (digit before space and digit after space)
                    Dim isNumberRange As Boolean
                    isNumberRange = False
                    If enPos > 2 And enPos + 1 < Len(paraText) Then
                        Dim charBeforeSpace As String
                        Dim charAfterSpace As String
                        charBeforeSpace = Mid$(paraText, enPos - 2, 1)
                        charAfterSpace = Mid$(paraText, enPos + 2, 1)
                        If (charBeforeSpace >= "0" And charBeforeSpace <= "9") And _
                           (charAfterSpace >= "0" And charAfterSpace <= "9") Then
                            isNumberRange = True
                        End If
                    End If
                    If isNumberRange Then GoTo NextEnDashPos
                    Dim snStart As Long
                    snStart = paraRange.Start + enPos - 1 - dashListPrefixLen
                    Dim snEnd As Long
                    snEnd = snStart + 1

                    Err.Clear
                    Dim snRng As Range
                    Set snRng = doc.Range(snStart, snEnd)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    Else
                        locStr = EngineGetLocationString(snRng, doc)
                    End If

                    Set finding = CreateIssueDict(RULE_NAME_DASH, locStr, _
                        "Spaced en-dash (" & enDash & ") found. Consider using an em-dash (" & emDash & ") for parenthetical interruptions.", _
                        emDash, snStart, snEnd, "warning", False)
                    issues.Add finding
                End If
            End If

NextEnDashPos:
            enPos = InStr(enPos + 1, paraText, enDash)
        Loop

NextParaDash:
    Next para
    On Error GoTo 0

    Set Check_DashUsage = issues
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
