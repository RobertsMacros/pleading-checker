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
'   - TextAnchoring.bas (IsInPageRange, IsPastPageFilter,
'     AddIssue, SafeRange, SafeLocationString,
'     CreateRegex, IsLetterChar, FindAll, IterateParagraphs)
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
    Dim results As Collection
    Set results = TextAnchoring.FindAll(doc, "[! ]/[! ]", False, False, True)

    Dim cnt As Long
    cnt = 0

    Dim item As Variant
    For Each item In results
        Dim rng As Range
        Set rng = TextAnchoring.SafeRange(doc, item(0), item(1))
        If Not rng Is Nothing Then
            ' Skip URLs and dates
            If Not IsURLContext(rng, doc) And Not IsDateSlash(rng) Then
                ' Skip conventional tight pairs (and/or, his/her, etc.)
                If Not IsConventionalTightSlash(rng, doc) Then
                    cnt = cnt + 1
                End If
            End If
        End If
    Next item

    CountTightSlashes = cnt
End Function

' ============================================================
'  PRIVATE: Count spaced slashes using literal search
' ============================================================
Private Function CountSpacedSlashes(doc As Document) As Long
    Dim results As Collection
    Set results = TextAnchoring.FindAll(doc, " / ", False, False, False)

    Dim cnt As Long
    cnt = 0

    Dim item As Variant
    For Each item In results
        Dim rng As Range
        Set rng = TextAnchoring.SafeRange(doc, item(0), item(1))
        If Not rng Is Nothing Then
            ' Skip URLs
            If Not IsURLContext(rng, doc) Then
                cnt = cnt + 1
            End If
        End If
    Next item

    CountSpacedSlashes = cnt
End Function

' ============================================================
'  PRIVATE: Flag spaced slashes (minority when tight is dominant)
' ============================================================
Private Sub FlagSpacedSlashes(doc As Document, ByRef issues As Collection)
    Dim results As Collection
    Set results = TextAnchoring.FindAll(doc, " / ", False, False, False)

    Dim item As Variant
    Dim sPos As Long, ePos As Long, mText As String
    For Each item In results
        sPos = item(0)
        ePos = item(1)
        mText = item(2)

        Dim rng As Range
        Set rng = TextAnchoring.SafeRange(doc, sPos, ePos)
        If rng Is Nothing Then GoTo ContinueSpaced
        If IsURLContext(rng, doc) Then GoTo ContinueSpaced

        TextAnchoring.AddIssue issues, RULE_NAME_SLASH, doc, rng, _
            "Spaced slash '" & mText & "' differs from dominant tight style", _
            "Remove spaces around slash for consistency", sPos, ePos, "possible_error"

ContinueSpaced:
    Next item
End Sub

' ============================================================
'  PRIVATE: Flag tight slashes (minority when spaced is dominant)
' ============================================================
Private Sub FlagTightSlashes(doc As Document, ByRef issues As Collection)
    Dim results As Collection
    Set results = TextAnchoring.FindAll(doc, "[! ]/[! ]", False, False, True)

    Dim item As Variant
    Dim sPos As Long, ePos As Long, mText As String
    For Each item In results
        sPos = item(0)
        ePos = item(1)
        mText = item(2)

        Dim rng As Range
        Set rng = TextAnchoring.SafeRange(doc, sPos, ePos)
        If rng Is Nothing Then GoTo ContinueTight
        If IsURLContext(rng, doc) Then GoTo ContinueTight
        If IsDateSlash(rng) Then GoTo ContinueTight
        If IsConventionalTightSlash(rng, doc) Then GoTo ContinueTight

        TextAnchoring.AddIssue issues, RULE_NAME_SLASH, doc, rng, _
            "Tight slash '" & mText & "' differs from dominant spaced style", _
            "Add spaces around slash for consistency", sPos, ePos, "possible_error"

ContinueTight:
    Next item
End Sub

' ============================================================
'  PRIVATE: Flag unexpected backslashes
' ============================================================
Private Sub FlagBackslashes(doc As Document, ByRef issues As Collection)
    Dim results As Collection
    Set results = TextAnchoring.FindAll(doc, "\", False, False, False)

    Dim item As Variant
    Dim sPos As Long, ePos As Long
    Dim context As String
    Dim fontName As String

    On Error Resume Next
    For Each item In results
        sPos = item(0)
        ePos = item(1)

        Dim rng As Range
        Set rng = TextAnchoring.SafeRange(doc, sPos, ePos)
        If rng Is Nothing Then GoTo ContinueBackslash

        ' Get surrounding context for skip checks
        Dim contextStart As Long
        Dim contextEnd As Long
        contextStart = sPos - 5
        If contextStart < 0 Then contextStart = 0
        contextEnd = ePos + 10
        If contextEnd > doc.Content.End Then contextEnd = doc.Content.End

        Dim contextRng As Range
        Set contextRng = TextAnchoring.SafeRange(doc, contextStart, contextEnd)
        If contextRng Is Nothing Then
            context = ""
        Else
            context = LCase(contextRng.Text)
        End If

        ' Skip file paths: drive letter pattern like C:\ or UNC \\server
        If IsDriveLetterPath(context) Or IsUNCPath(context) Then GoTo ContinueBackslash

        ' Skip code-font runs (Courier, Consolas)
        fontName = ""
        Err.Clear
        fontName = rng.Font.Name
        If Err.Number <> 0 Then Err.Clear: fontName = ""
        If IsCodeFontName(fontName) Then GoTo ContinueBackslash

        ' Skip URLs
        If InStr(1, context, "://") > 0 Then GoTo ContinueBackslash

        ' Flag the backslash
        TextAnchoring.AddIssue issues, RULE_NAME_SLASH, doc, rng, _
            "Unexpected backslash -- did you mean forward slash?", _
            "Replace '\' with '/'", sPos, ePos, "possible_error"

ContinueBackslash:
    Next item
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

    ' -- Cheap global pre-check: count all brackets in the document.
    '    If all three types balance globally, skip the expensive
    '    per-paragraph traversal entirely.
    Dim gPO As Long, gPC As Long
    Dim gSO As Long, gSC As Long
    Dim gCO As Long, gCC As Long
    Dim gBytes() As Byte, gLen As Long, gIdx As Long, gCode As Long

    On Error Resume Next
    Dim fullText As String
    fullText = doc.Content.Text
    If Err.Number <> 0 Then
        Err.Clear
        fullText = ""
    End If
    On Error GoTo 0

    If Len(fullText) > 0 Then
        gBytes = fullText
        gLen = UBound(gBytes) - 1
        For gIdx = 0 To gLen Step 2
            gCode = gBytes(gIdx) Or (CLng(gBytes(gIdx + 1)) * 256&)
            Select Case gCode
                Case 40: gPO = gPO + 1
                Case 41: gPC = gPC + 1
                Case 91: gSO = gSO + 1
                Case 93: gSC = gSC + 1
                Case 123: gCO = gCO + 1
                Case 125: gCC = gCC + 1
            End Select
        Next gIdx

        ' If all brackets balance globally, no per-paragraph issues possible
        If gPO = gPC And gSO = gSC And gCO = gCC Then
            Set Check_BracketIntegrity = issues
            Exit Function
        End If
    End If

    Set issues = TextAnchoring.IterateParagraphs(doc, "Rules_Punctuation", "ProcessParagraph_BracketIntegrity")

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
    Dim rng As Range
    Set rng = TextAnchoring.SafeRange(doc, pos, pos + 1)
    If rng Is Nothing Then Exit Sub

    ' Skip if outside page range
    If Not TextAnchoring.IsInPageRange(rng) Then Exit Sub

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

    TextAnchoring.AddIssue issues, RULE_NAME_BRACKET, doc, rng, issueText, suggestion, pos, pos + 1
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

' ============================================================
'  TRIPLICATE PUNCTUATION
'  Flags three or more consecutive identical punctuation marks:
'    (((  )))  [[[  ]]]  """  '''  ,,,
'  Deliberately excludes "..." (ellipsis).
' ============================================================
Public Function Check_TriplicatePunctuation(doc As Document) As Collection
    Set Check_TriplicatePunctuation = TextAnchoring.IterateParagraphs(doc, "Rules_Punctuation", "ProcessParagraph_TriplicatePunctuation")
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
    Set Check_DashUsage = TextAnchoring.IterateParagraphs(doc, "Rules_Punctuation", "ProcessParagraph_DashUsage")
End Function

' ============================================================
'  ProcessParagraph_TriplicatePunctuation
'  Per-paragraph handler extracted from Check_TriplicatePunctuation.
'  Scans paraText for runs of 3+ identical punctuation chars
'  (excluding "." to avoid flagging ellipsis).
' ============================================================
Public Sub ProcessParagraph_TriplicatePunctuation(doc As Document, paraRange As Range, _
        paraText As String, paraStart As Long, listPrefixLen As Long, _
        ByRef issues As Collection)
    Dim i As Long
    Dim ch As String
    Dim runLen As Long

    ' Characters to check (NOT including "." to avoid flagging ellipsis)
    Dim targets As String
    targets = "()[]""',"

    If Len(paraText) < 3 Then Exit Sub

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
                Dim matched As String
                matched = String$(runLen, ch)
                Dim sPos As Long, ePos As Long
                sPos = paraStart + i - 1 - listPrefixLen
                ePos = sPos + runLen
                Dim rng As Range
                Set rng = TextAnchoring.SafeRange(doc, sPos, ePos)
                TextAnchoring.AddIssue issues, RULE_NAME_TRIPLICATE, doc, rng, _
                    "Triplicate punctuation: '" & matched & "'", "Remove repeated punctuation", _
                    sPos, ePos, "error", False, "", matched
            End If
            i = i + runLen
        Else
            i = i + 1
        End If
    Loop
End Sub

' ============================================================
'  ProcessParagraph_DashUsage
'  Per-paragraph handler extracted from Check_DashUsage.
'  Applies all dash checks (hyphen in number range, double
'  hyphen, en-dash between words, spaced en-dash).
' ============================================================
Public Sub ProcessParagraph_DashUsage(doc As Document, paraRange As Range, _
        paraText As String, paraStart As Long, listPrefixLen As Long, _
        ByRef issues As Collection)

    Dim reHyphenRange As Object
    Set reHyphenRange = TextAnchoring.CreateRegex("(\d)-(\d)")

    Dim reDoubleHyphen As Object
    Set reDoubleHyphen = TextAnchoring.CreateRegex("--")

    Dim enDash As String
    enDash = ChrW(8211)
    Dim emDash As String
    emDash = ChrW(8212)

    ' Strip para mark
    If Len(paraText) > 0 Then
        If Right$(paraText, 1) = vbCr Or Right$(paraText, 1) = Chr(13) Then
            paraText = Left$(paraText, Len(paraText) - 1)
        End If
    End If
    If Len(paraText) < 2 Then Exit Sub

    ' --- Check 1: Hyphen in number ranges (digit-digit) ---
    Dim mHR As Object
    Set mHR = reHyphenRange.Execute(paraText)
    Dim hm As Object
    For Each hm In mHR
        Dim hrStart As Long
        hrStart = paraStart + hm.FirstIndex - listPrefixLen
        ' The hyphen is at offset +1 (after first digit in pattern)
        Dim hyphenPos As Long
        hyphenPos = hrStart + 1
        Dim hrEnd As Long
        hrEnd = hyphenPos + 1  ' just the hyphen

        Dim hrRng As Range
        Set hrRng = TextAnchoring.SafeRange(doc, hyphenPos, hrEnd)
        TextAnchoring.AddIssue issues, RULE_NAME_DASH, doc, hrRng, _
            "Hyphen used in number range. Use an en-dash (" & enDash & ") for ranges.", _
            "Replace hyphen with en-dash", hyphenPos, hrEnd, "error", True, enDash
    Next hm

    ' --- Check 2: Double-hyphen "--" should be em-dash ---
    Dim mDH As Object
    Set mDH = reDoubleHyphen.Execute(paraText)
    Dim dhm As Object
    For Each dhm In mDH
        Dim dhStart As Long
        dhStart = paraStart + dhm.FirstIndex - listPrefixLen
        Dim dhEnd As Long
        dhEnd = dhStart + 2

        Dim dhRng As Range
        Set dhRng = TextAnchoring.SafeRange(doc, dhStart, dhEnd)
        TextAnchoring.AddIssue issues, RULE_NAME_DASH, doc, dhRng, _
            "Double-hyphen found. Use an em-dash (" & emDash & ") instead.", _
            "Replace with em-dash", dhStart, dhEnd, "error", True, emDash
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
            beforeIsLetter = TextAnchoring.IsLetterChar(chBefore)
            afterIsLetter = TextAnchoring.IsLetterChar(chAfter)

            If beforeIsLetter And afterIsLetter Then
                Dim enStart As Long
                enStart = paraStart + enPos - 1 - listPrefixLen
                Dim enEnd As Long
                enEnd = enStart + 1

                Dim enRng As Range
                Set enRng = TextAnchoring.SafeRange(doc, enStart, enEnd)
                TextAnchoring.AddIssue issues, RULE_NAME_DASH, doc, enRng, _
                    "En-dash (" & enDash & ") used between words. Use a hyphen (-) for compound words.", _
                    "Replace en-dash with hyphen", enStart, enEnd, "error", True, "-"
            End If

            ' Check 4: Spaced en-dash (" - ") -> should be em-dash (" -- ")
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
                If Not isNumberRange Then
                    Dim snStart As Long
                    snStart = paraStart + enPos - 1 - listPrefixLen
                    Dim snEnd As Long
                    snEnd = snStart + 1

                    Dim snRng As Range
                    Set snRng = TextAnchoring.SafeRange(doc, snStart, snEnd)
                    TextAnchoring.AddIssue issues, RULE_NAME_DASH, doc, snRng, _
                        "Spaced en-dash (" & enDash & ") found. Consider using an em-dash (" & emDash & ") for parenthetical interruptions.", _
                        emDash, snStart, snEnd, "warning", False
                End If
            End If
        End If

        enPos = InStr(enPos + 1, paraText, enDash)
    Loop
End Sub

' ============================================================
'  ProcessParagraph_BracketIntegrity
'  Per-paragraph handler extracted from Check_BracketIntegrity.
'  Performs byte-array bracket counting and stack-based nesting
'  check for a single paragraph.  The global pre-check (counting
'  all brackets in doc.Content.Text) is NOT done here -- it
'  belongs in the Check_BracketIntegrity function.
' ============================================================
Public Sub ProcessParagraph_BracketIntegrity(doc As Document, paraRange As Range, _
        paraText As String, paraStart As Long, listPrefixLen As Long, _
        ByRef issues As Collection)
    If LenB(paraText) = 0 Then Exit Sub

    ' Counters per bracket type
    Dim parenOpen As Long, parenClose As Long
    Dim sqOpen As Long, sqClose As Long
    Dim curlyOpen As Long, curlyClose As Long

    ' Position of first unmatched bracket (for issue location)
    Dim firstParenPos As Long, firstSqPos As Long, firstCurlyPos As Long

    parenOpen = 0: parenClose = 0
    sqOpen = 0: sqClose = 0
    curlyOpen = 0: curlyClose = 0
    firstParenPos = -1: firstSqPos = -1: firstCurlyPos = -1

    Dim b() As Byte, bMax As Long
    Dim i As Long, code As Long, pos As Long

    b = paraText
    bMax = UBound(b) - 1

    For i = 0 To bMax Step 2
        code = b(i) Or (CLng(b(i + 1)) * 256&)
        pos = paraStart + (i \ 2) - listPrefixLen

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
        CreateBracketIssue doc, issues, firstParenPos, "()", "Unbalanced parentheses: " & parenOpen & " opened, " & parenClose & " closed"
    End If
    If sqOpen <> sqClose Then
        CreateBracketIssue doc, issues, firstSqPos, "[]", "Unbalanced square brackets: " & sqOpen & " opened, " & sqClose & " closed"
    End If
    If curlyOpen <> curlyClose Then
        CreateBracketIssue doc, issues, firstCurlyPos, "{}", "Unbalanced curly braces: " & curlyOpen & " opened, " & curlyClose & " closed"
    End If

    ' -- Stack-based nesting check (only when counts balance) --
    If parenOpen = parenClose And sqOpen = sqClose And curlyOpen = curlyClose And (parenOpen + sqOpen + curlyOpen) > 0 Then
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
                        nestPos = paraStart + (i \ 2) - listPrefixLen
                        Exit For
                    End If
                    If Not CodesMatch(stk(stkTop), code) Then
                        nestBad = True
                        nestPos = paraStart + (i \ 2) - listPrefixLen
                        Exit For
                    End If
                    stkTop = stkTop - 1
            End Select
        Next i
        If nestBad Then
            CreateBracketIssue doc, issues, nestPos, "()", "Improperly nested brackets (e.g. overlapping pairs)"
        End If
    End If
End Sub
