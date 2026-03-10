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

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

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

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

        If Not EngineIsInPageRange(rng) Then GoTo ContinueTight
        If IsURLContext(rng, doc) Then GoTo ContinueTight
        If IsDateSlash(rng) Then GoTo ContinueTight

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

    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do

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

        Set finding = CreateIssueDict(RULE_NAME_SLASH, locStr, "Unexpected backslash " & Chr(8212) & " did you mean forward slash?", "Replace '\' with '/'", rng.Start, rng.End, "possible_error")
        issues.Add finding

ContinueBackslash:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

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
    Dim docText As String

    ' -- Get full document text ---------------------------------
    On Error Resume Next
    docText = doc.Content.Text
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        Set Check_BracketIntegrity = issues
        Exit Function
    End If
    On Error GoTo 0

    If LenB(docText) = 0 Then
        Set Check_BracketIntegrity = issues
        Exit Function
    End If

    ' -- Byte-array scan (avoids 600K Mid$ allocs) -------------
    Dim b() As Byte
    b = docText
    Dim bMax As Long
    bMax = UBound(b) - 1

    ' Stack: parallel Long arrays (code-point + doc position)
    Dim stackCodes() As Long
    Dim stackPos() As Long
    Dim stackTop As Long
    stackTop = -1
    ReDim stackCodes(0 To 1000)
    ReDim stackPos(0 To 1000)

    ' Reusable range for font checks (created once, moved
    ' via SetRange -- avoids per-bracket doc.Range creation)
    Dim fontRng As Range
    Set fontRng = doc.Range(0, 1)

    Dim i As Long, code As Long, pos As Long
    Dim isOpen As Boolean, isClose As Boolean
    Dim fontName As String
    Dim openCode As Long, openPos As Long

    For i = 0 To bMax Step 2
        code = b(i) Or (CLng(b(i + 1)) * 256&)

        ' Fast integer check: ( ) [ ] { }
        isOpen = (code = 40 Or code = 91 Or code = 123)
        isClose = (code = 41 Or code = 93 Or code = 125)

        If isOpen Or isClose Then
            pos = i \ 2  ' document position (0-based)

            ' Code-font gate (reusable range, no alloc)
            On Error Resume Next
            Err.Clear
            fontRng.SetRange pos, pos + 1
            If Err.Number = 0 Then
                fontName = fontRng.Font.Name
                If Err.Number <> 0 Then fontName = "": Err.Clear
                If IsCodeFontName(fontName) Then
                    On Error GoTo 0
                    GoTo NxtBracket
                End If
            Else
                Err.Clear
            End If
            On Error GoTo 0

            If isOpen Then
                stackTop = stackTop + 1
                If stackTop > UBound(stackCodes) Then
                    ReDim Preserve stackCodes(0 To stackTop + 500)
                    ReDim Preserve stackPos(0 To stackTop + 500)
                End If
                stackCodes(stackTop) = code
                stackPos(stackTop) = pos
            Else
                ' Closing bracket
                If stackTop < 0 Then
                    CreateBracketIssue doc, issues, pos, ChrW$(code), _
                        "Unmatched closing bracket '" & ChrW$(code) & _
                        "' with no corresponding opener"
                Else
                    openCode = stackCodes(stackTop)
                    openPos = stackPos(stackTop)
                    stackTop = stackTop - 1

                    If Not CodesMatch(openCode, code) Then
                        CreateBracketIssue doc, issues, openPos, _
                            ChrW$(openCode), _
                            "Mismatched bracket: opened with '" & _
                            ChrW$(openCode) & "' but closed with '" & _
                            ChrW$(code) & "'"
                        CreateBracketIssue doc, issues, pos, _
                            ChrW$(code), _
                            "Mismatched bracket: closing '" & _
                            ChrW$(code) & "' does not match opener '" & _
                            ChrW$(openCode) & "'"
                    End If
                End If
            End If
        End If

NxtBracket:
    Next i

    ' -- Unmatched openers remaining on stack -------------------
    Dim s As Long
    For s = 0 To stackTop
        CreateBracketIssue doc, issues, stackPos(s), _
            ChrW$(stackCodes(s)), _
            "Unmatched opening bracket '" & ChrW$(stackCodes(s)) & _
            "' with no corresponding closer"
    Next s

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
    If bracketChar = "(" Or bracketChar = ")" Then
        suggestion = "Add or correct matching parenthesis"
    ElseIf bracketChar = "[" Or bracketChar = "]" Then
        suggestion = "Add or correct matching square bracket"
    ElseIf bracketChar = "{" Or bracketChar = "}" Then
        suggestion = "Add or correct matching curly brace"
    Else
        suggestion = "Review bracket pairing"
    End If

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
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function
