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
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_SLASH As String = "slash_style"
Private Const RULE_NAME_BRACKET As String = "bracket_integrity"

' ╔══════════════════════════════════════════════════════════════╗
' ║  SLASH STYLE (Rule14)                                       ║
' ╚══════════════════════════════════════════════════════════════╝

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT: Slash Style
' ════════════════════════════════════════════════════════════
Public Function Check_SlashStyle(doc As Document) As Collection
    Dim issues As New Collection

    ' ── Forward slashes: determine dominant style ────────────
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

    ' ── Backslashes ──────────────────────────────────────────
    FlagBackslashes doc, issues

    Set Check_SlashStyle = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Count tight slashes using wildcard search
' ════════════════════════════════════════════════════════════
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

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

        ' Skip URLs and dates
        If Not IsURLContext(rng, doc) And Not IsDateSlash(rng) Then
            cnt = cnt + 1
        End If

        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0

    CountTightSlashes = cnt
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Count spaced slashes using literal search
' ════════════════════════════════════════════════════════════
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

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Flag spaced slashes (minority when tight is dominant)
' ════════════════════════════════════════════════════════════
Private Sub FlagSpacedSlashes(doc As Document, ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
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

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

        If Not PleadingsEngine.IsInPageRange(rng) Then GoTo ContinueSpaced
        If IsURLContext(rng, doc) Then GoTo ContinueSpaced

        locStr = PleadingsEngine.GetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set issue = New PleadingsIssue
        issue.Init RULE_NAME_SLASH, _
                   locStr, _
                   "Spaced slash '" & rng.Text & "' differs from dominant tight style", _
                   "Remove spaces around slash for consistency", _
                   rng.Start, _
                   rng.End, _
                   "possible_error"
        issues.Add issue

ContinueSpaced:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Flag tight slashes (minority when spaced is dominant)
' ════════════════════════════════════════════════════════════
Private Sub FlagTightSlashes(doc As Document, ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
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

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

        If Not PleadingsEngine.IsInPageRange(rng) Then GoTo ContinueTight
        If IsURLContext(rng, doc) Then GoTo ContinueTight
        If IsDateSlash(rng) Then GoTo ContinueTight

        locStr = PleadingsEngine.GetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set issue = New PleadingsIssue
        issue.Init RULE_NAME_SLASH, _
                   locStr, _
                   "Tight slash '" & rng.Text & "' differs from dominant spaced style", _
                   "Add spaces around slash for consistency", _
                   rng.Start, _
                   rng.End, _
                   "possible_error"
        issues.Add issue

ContinueTight:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Flag unexpected backslashes
' ════════════════════════════════════════════════════════════
Private Sub FlagBackslashes(doc As Document, ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
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

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

        If Not PleadingsEngine.IsInPageRange(rng) Then GoTo ContinueBackslash

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
        locStr = PleadingsEngine.GetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set issue = New PleadingsIssue
        issue.Init RULE_NAME_SLASH, _
                   locStr, _
                   "Unexpected backslash — did you mean forward slash?", _
                   "Replace '\' with '/'", _
                   rng.Start, _
                   rng.End, _
                   "possible_error"
        issues.Add issue

ContinueBackslash:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if context suggests a URL
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if slash is part of a date (digits only)
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check for drive letter path pattern (e.g. C:\)
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check for UNC path pattern (\\server)
' ════════════════════════════════════════════════════════════
Private Function IsUNCPath(ByVal context As String) As Boolean
    IsUNCPath = (InStr(1, context, "\\") > 0)
End Function

' ╔══════════════════════════════════════════════════════════════╗
' ║  BRACKET INTEGRITY (Rule16)                                 ║
' ╚══════════════════════════════════════════════════════════════╝

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT: Bracket Integrity
' ════════════════════════════════════════════════════════════
Public Function Check_BracketIntegrity(doc As Document) As Collection
    Dim issues As New Collection
    Dim docText As String
    Dim textLen As Long
    Dim i As Long
    Dim ch As String

    ' ── Stack using parallel arrays ──────────────────────────
    Dim stackChars() As String
    Dim stackPositions() As Long
    Dim stackTop As Long

    ReDim stackChars(0 To 1000)
    ReDim stackPositions(0 To 1000)
    stackTop = -1 ' empty stack

    ' ── Get full document text ───────────────────────────────
    On Error Resume Next
    docText = doc.Content.Text
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set Check_BracketIntegrity = issues
        Exit Function
    End If
    On Error GoTo 0

    textLen = Len(docText)
    If textLen = 0 Then
        Set Check_BracketIntegrity = issues
        Exit Function
    End If

    ' ── Iterate character by character ───────────────────────
    For i = 1 To textLen
        ch = Mid(docText, i, 1)

        ' Only process bracket characters
        If ch = "(" Or ch = "[" Or ch = "{" Or _
           ch = ")" Or ch = "]" Or ch = "}" Then

            ' Skip brackets in code-font runs
            If IsCodeFont(doc, i - 1) Then GoTo NextChar

            If ch = "(" Or ch = "[" Or ch = "{" Then
                ' ── Push opening bracket onto stack ──────────
                stackTop = stackTop + 1

                ' Grow arrays if needed
                If stackTop > UBound(stackChars) Then
                    ReDim Preserve stackChars(0 To stackTop + 500)
                    ReDim Preserve stackPositions(0 To stackTop + 500)
                End If

                stackChars(stackTop) = ch
                stackPositions(stackTop) = i - 1 ' 0-based doc position

            Else
                ' ── Closing bracket: pop and check match ─────
                If stackTop < 0 Then
                    ' Empty stack: unmatched closing bracket
                    CreateBracketIssue doc, issues, i - 1, ch, _
                        "Unmatched closing bracket '" & ch & "' with no corresponding opener"
                Else
                    Dim openChar As String
                    Dim openPos As Long
                    openChar = stackChars(stackTop)
                    openPos = stackPositions(stackTop)
                    stackTop = stackTop - 1

                    ' Check if bracket types match
                    If Not BracketsMatch(openChar, ch) Then
                        ' Mismatched bracket pair
                        CreateBracketIssue doc, issues, openPos, openChar, _
                            "Mismatched bracket: opened with '" & openChar & _
                            "' but closed with '" & ch & "'"
                        CreateBracketIssue doc, issues, i - 1, ch, _
                            "Mismatched bracket: closing '" & ch & _
                            "' does not match opener '" & openChar & "'"
                    End If
                End If
            End If
        End If

NextChar:
    Next i

    ' ── Any remaining on stack are unmatched openers ─────────
    Dim s As Long
    For s = 0 To stackTop
        CreateBracketIssue doc, issues, stackPositions(s), stackChars(s), _
            "Unmatched opening bracket '" & stackChars(s) & "' with no corresponding closer"
    Next s

    Set Check_BracketIntegrity = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if brackets match
' ════════════════════════════════════════════════════════════
Private Function BracketsMatch(ByVal openCh As String, _
                                ByVal closeCh As String) As Boolean
    Select Case openCh
        Case "("
            BracketsMatch = (closeCh = ")")
        Case "["
            BracketsMatch = (closeCh = "]")
        Case "{"
            BracketsMatch = (closeCh = "}")
        Case Else
            BracketsMatch = False
    End Select
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if position is in a code font
' ════════════════════════════════════════════════════════════
Private Function IsCodeFont(doc As Document, ByVal pos As Long) As Boolean
    Dim rng As Range
    Dim fontName As String

    On Error Resume Next
    Set rng = doc.Range(pos, pos + 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsCodeFont = False
        Exit Function
    End If

    fontName = ""
    fontName = rng.Font.Name
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsCodeFont = False
        Exit Function
    End If
    On Error GoTo 0

    IsCodeFont = IsCodeFontName(fontName)
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Create a bracket integrity issue
' ════════════════════════════════════════════════════════════
Private Sub CreateBracketIssue(doc As Document, _
                                ByRef issues As Collection, _
                                ByVal pos As Long, _
                                ByVal bracketChar As String, _
                                ByVal issueText As String)
    Dim issue As PleadingsIssue
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
    If Not PleadingsEngine.IsInPageRange(rng) Then
        On Error GoTo 0
        Exit Sub
    End If

    locStr = PleadingsEngine.GetLocationString(rng, doc)
    If Err.Number <> 0 Then
        locStr = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0

    ' Determine suggestion based on bracket type
    Dim suggestion As String
    If bracketChar = "(" Or bracketChar = ")" Then
        suggestion = "Add or correct matching parenthesis"
    ElseIf bracketChar = "[" Or bracketChar = "]" Then
        suggestion = "Add or correct matching square bracket"
    ElseIf bracketChar = "{" Or bracketChar = "}" Then
        suggestion = "Add or correct matching curly brace"
    Else
        suggestion = "Review bracket pairing"
    End If

    Set issue = New PleadingsIssue
    issue.Init RULE_NAME_BRACKET, _
               locStr, _
               issueText, _
               suggestion, _
               pos, _
               pos + 1, _
               "error"
    issues.Add issue
End Sub

' ╔══════════════════════════════════════════════════════════════╗
' ║  SHARED PRIVATE HELPERS                                     ║
' ╚══════════════════════════════════════════════════════════════╝

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if a font name is a code font (Courier, Consolas)
'  Shared by FlagBackslashes and IsCodeFont
' ════════════════════════════════════════════════════════════
Private Function IsCodeFontName(ByVal fontName As String) As Boolean
    IsCodeFontName = (LCase(fontName) Like "*courier*") Or _
                     (LCase(fontName) Like "*consolas*")
End Function
