Attribute VB_Name = "TextAnchoring"
' ============================================================
' TextAnchoring.bas
' Centralised text-anchoring utilities used by every rule module.
'
' Provides:
'   - Page-range filtering  (delegates to PleadingsEngine)
'   - Location strings      (delegates to PleadingsEngine)
'   - Issue-dict factory    (single canonical implementation)
'   - Preference getters    (delegates to PleadingsEngine)
'   - Profiling wrappers    (delegates to PleadingsEngine)
'   - Text helpers          (punctuation, list prefix, etc.)
'
' All functions are Public so rule modules can call them
' directly (e.g.  TextAnchoring.IsInPageRange).
' ============================================================
Option Explicit

' ============================================================
'  PAGE-RANGE FILTERING
' ============================================================
Public Function IsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    IsInPageRange = PleadingsEngine.IsInPageRange(rng)
    If Err.Number <> 0 Then
        Debug.Print "TextAnchoring.IsInPageRange: fallback (Err " & Err.Number & ")"
        IsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function IsPastPageFilter(ByVal startPos As Long) As Boolean
    On Error Resume Next
    IsPastPageFilter = PleadingsEngine.IsPastPageFilter(startPos)
    If Err.Number <> 0 Then
        IsPastPageFilter = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function IsInPageRangeByPos(ByVal startPos As Long, _
                                    ByVal endPos As Long) As Boolean
    On Error Resume Next
    IsInPageRangeByPos = PleadingsEngine.IsInPageRangeByPos(startPos, endPos)
    If Err.Number <> 0 Then
        IsInPageRangeByPos = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ============================================================
'  LOCATION STRING
' ============================================================
Public Function GetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    GetLocationString = PleadingsEngine.GetLocationString(rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "TextAnchoring.GetLocationString: fallback (Err " & Err.Number & ")"
        GetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ============================================================
'  ISSUE-DICT FACTORY  (single canonical implementation)
' ============================================================
Public Function CreateIssueDict(ByVal ruleName_ As String, _
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

' ============================================================
'  PREFERENCE GETTERS  (delegate to PleadingsEngine)
' ============================================================
Public Function GetSpellingMode() As String
    On Error Resume Next
    GetSpellingMode = PleadingsEngine.GetSpellingMode()
    If Err.Number <> 0 Then
        GetSpellingMode = "UK"
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function GetSpaceStylePref() As String
    On Error Resume Next
    GetSpaceStylePref = PleadingsEngine.GetSpaceStylePref()
    If Err.Number <> 0 Then
        GetSpaceStylePref = "ONE"
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function GetDateFormatPref() As String
    On Error Resume Next
    GetDateFormatPref = PleadingsEngine.GetDateFormatPref()
    If Err.Number <> 0 Then
        GetDateFormatPref = "UK"
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function IsWhitelistedTerm(ByVal term As String) As Boolean
    On Error Resume Next
    IsWhitelistedTerm = PleadingsEngine.IsWhitelistedTerm(term)
    If Err.Number <> 0 Then
        IsWhitelistedTerm = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Sub SetWhitelist(dict As Object)
    On Error Resume Next
    PleadingsEngine.SetWhitelist dict
    If Err.Number <> 0 Then
        Debug.Print "TextAnchoring.SetWhitelist: fallback (Err " & Err.Number & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub SetPageRange(ByVal startPg As Long, ByVal endPg As Long)
    On Error Resume Next
    PleadingsEngine.SetPageRange startPg, endPg
    If Err.Number <> 0 Then
        Debug.Print "TextAnchoring.SetPageRange: fallback (Err " & Err.Number & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ============================================================
'  PROFILING WRAPPERS
' ============================================================
Public Sub PerfTimerStart(ByVal label As String)
    On Error Resume Next
    PleadingsEngine.PerfTimerStart label
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Public Sub PerfTimerEnd(ByVal label As String)
    On Error Resume Next
    PleadingsEngine.PerfTimerEnd label
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Public Sub PerfCount(ByVal label As String, Optional ByVal increment As Long = 1)
    On Error Resume Next
    PleadingsEngine.PerfCount label, increment
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

' ============================================================
'  SHARED HELPERS  (eliminate boilerplate in rule modules)
' ============================================================

' Get location string with built-in error handling.
' Returns "unknown location" on any failure.
Public Function SafeLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    SafeLocationString = PleadingsEngine.GetLocationString(rng, doc)
    If Err.Number <> 0 Then SafeLocationString = "unknown location": Err.Clear
    On Error GoTo 0
End Function

' Create a Range with error handling. Returns Nothing on failure.
Public Function SafeRange(doc As Document, ByVal startPos As Long, ByVal endPos As Long) As Range
    On Error Resume Next
    Set SafeRange = doc.Range(startPos, endPos)
    If Err.Number <> 0 Then Set SafeRange = Nothing: Err.Clear
    On Error GoTo 0
End Function

' Check if a character is whitespace (space, tab, NBSP, CR, LF, vertical tab).
Public Function IsWhitespaceChar(ByVal ch As String) As Boolean
    IsWhitespaceChar = (ch = " " Or ch = vbTab Or ch = ChrW(160) Or ch = vbCr Or ch = vbLf Or ch = Chr(11))
End Function

' Factory for VBScript.RegExp objects.
Public Function CreateRegex(ByVal pattern As String, Optional ByVal isGlobal As Boolean = True, Optional ByVal ignoreCase As Boolean = False) As Object
    Set CreateRegex = CreateObject("VBScript.RegExp")
    CreateRegex.Global = isGlobal
    CreateRegex.IgnoreCase = ignoreCase
    CreateRegex.pattern = pattern
End Function

' One-call issue creation: fetches location string, builds issue dict, adds to collection.
' Pass rng = Nothing and supply locStr via the overload if you already have a location string.
Public Sub AddIssue(ByRef issues As Collection, ByVal ruleName As String, doc As Document, rng As Object, ByVal msg As String, ByVal suggestion As String, ByVal startPos As Long, ByVal endPos As Long, Optional ByVal severity As String = "error", Optional ByVal autoFixSafe As Boolean = False, Optional ByVal replacementText As String = "", Optional ByVal matchedText As String = "", Optional ByVal anchorKind As String = "exact_text", Optional ByVal confidence As String = "high")
    Dim locStr As String
    If rng Is Nothing Then
        locStr = "unknown location"
    Else
        locStr = SafeLocationString(rng, doc)
    End If
    Dim d As Object
    Set d = CreateIssueDict(ruleName, locStr, msg, suggestion, startPos, endPos, severity, autoFixSafe, replacementText, matchedText, anchorKind, confidence)
    issues.Add d
End Sub

' Generic paragraph iterator.  Calls the named ProcessParagraph_ sub
' via Application.Run with the standard signature:
'   (doc, paraRange, paraText, paraStart, listPrefixLen, issues)
' Returns the populated issues collection.
Public Function IterateParagraphs(doc As Document, ByVal moduleName As String, ByVal procName As String) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim paraStart As Long
    Dim listPrefixLen As Long

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextPara_IP
        If IsPastPageFilter(paraRange.Start) Then Exit For
        If Not IsInPageRange(paraRange) Then GoTo NextPara_IP

        paraText = StripParaMarkChar(paraRange.Text)
        If Err.Number <> 0 Then Err.Clear: GoTo NextPara_IP
        If Len(paraText) < 2 Then GoTo NextPara_IP

        paraStart = paraRange.Start
        listPrefixLen = GetListPrefixLen(para, paraText)

        Err.Clear
        Application.Run moduleName & "." & procName, doc, paraRange, paraText, paraStart, listPrefixLen, issues
        If Err.Number <> 0 Then Err.Clear
NextPara_IP:
    Next para
    On Error GoTo 0
    Set IterateParagraphs = issues
End Function

' Generic Find loop with stall guard.
' Returns a Collection of 3-element arrays: Array(startPos, endPos, matchText).
' Only matches inside the page range are returned.
Public Function FindAll(doc As Document, ByVal searchText As String, Optional ByVal wholeWord As Boolean = True, Optional ByVal matchCase As Boolean = True, Optional ByVal useWildcards As Boolean = False, Optional searchRange As Range = Nothing) As Collection
    Dim results As New Collection
    Dim rng As Range
    If searchRange Is Nothing Then
        Set rng = doc.Content.Duplicate
    Else
        Set rng = searchRange.Duplicate
    End If

    With rng.Find
        .ClearFormatting
        .Text = searchText
        .MatchWholeWord = wholeWord
        .MatchCase = matchCase
        .MatchWildcards = useWildcards
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    On Error Resume Next
    Do
        Err.Clear
        Dim didFind As Boolean
        didFind = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: Exit Do
        If Not didFind Then Exit Do
        If rng.Start <= lastPos Then Exit Do
        lastPos = rng.Start
        If IsInPageRange(rng) Then
            results.Add Array(rng.Start, rng.End, rng.Text)
        End If
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: Exit Do
    Loop
    On Error GoTo 0
    Set FindAll = results
End Function

' Generic footnote iterator.  Calls the named handler via Application.Run
' with signature: (doc, fn As Footnote, noteText As String, issues As Collection)
' Handles error recovery and page-range filtering on fn.Reference.
Public Function ForEachFootnote(doc As Document, ByVal moduleName As String, ByVal procName As String) As Collection
    Dim issues As New Collection
    Dim i As Long
    Dim fn As Footnote
    Dim noteText As String

    For i = 1 To doc.Footnotes.Count
        On Error Resume Next
        Set fn = doc.Footnotes(i)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFN_FEF
        On Error GoTo 0

        On Error Resume Next
        If Not IsInPageRange(fn.Reference) Then On Error GoTo 0: GoTo NextFN_FEF
        On Error GoTo 0

        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFN_FEF
        On Error GoTo 0

        On Error Resume Next
        Application.Run moduleName & "." & procName, doc, fn, noteText, issues
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
NextFN_FEF:
    Next i
    Set ForEachFootnote = issues
End Function

' ============================================================
'  TEXT HELPERS
' ============================================================

' Strip leading/trailing punctuation from a word token
Public Function StripPunctuation(ByVal word As String) As String
    Dim ch As String
    Dim startPos As Long
    Dim endPos As Long

    word = Trim(word)
    If Len(word) = 0 Then
        StripPunctuation = ""
        Exit Function
    End If

    startPos = 1
    Do While startPos <= Len(word)
        ch = Mid$(word, startPos, 1)
        If IsPunctuation(ch) Then
            startPos = startPos + 1
        Else
            Exit Do
        End If
    Loop

    endPos = Len(word)
    Do While endPos >= startPos
        ch = Mid$(word, endPos, 1)
        If IsPunctuation(ch) Then
            endPos = endPos - 1
        Else
            Exit Do
        End If
    Loop

    If startPos > endPos Then
        StripPunctuation = ""
    Else
        StripPunctuation = Mid$(word, startPos, endPos - startPos + 1)
    End If
End Function

' Check if a single character is punctuation
Public Function IsPunctuation(ByVal ch As String) As Boolean
    Dim PUNCT_CHARS As String
    PUNCT_CHARS = ".,;:!?""'()[]{}/-" & ChrW$(8220) & ChrW$(8221) & _
                  ChrW$(8216) & ChrW$(8217) & ChrW$(8212) & ChrW$(8211)
    IsPunctuation = (InStr(1, PUNCT_CHARS, ch) > 0)
End Function

' Check if a single character is a letter (A-Z, a-z)
Public Function IsLetterChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsLetterChar = False
        Exit Function
    End If
    Dim c As Long
    c = AscW(ch)
    IsLetterChar = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

' ============================================================
'  ARRAY MERGE HELPERS
' ============================================================

' Merge 2 Variant arrays into one flat Variant array
Public Function MergeArrays2(a1 As Variant, a2 As Variant) As Variant
    Dim total As Long
    total = UBound(a1) - LBound(a1) + 1 _
          + UBound(a2) - LBound(a2) + 1
    Dim out() As Variant
    ReDim out(0 To total - 1)
    Dim idx As Long
    idx = 0
    Dim v As Variant
    For Each v In a1: out(idx) = v: idx = idx + 1: Next v
    For Each v In a2: out(idx) = v: idx = idx + 1: Next v
    MergeArrays2 = out
End Function

' Merge 3 Variant arrays into one flat Variant array
Public Function MergeArrays3(a1 As Variant, a2 As Variant, a3 As Variant) As Variant
    Dim total As Long
    total = UBound(a1) - LBound(a1) + 1 _
          + UBound(a2) - LBound(a2) + 1 _
          + UBound(a3) - LBound(a3) + 1
    Dim out() As Variant
    ReDim out(0 To total - 1)
    Dim idx As Long
    idx = 0
    Dim v As Variant
    For Each v In a1: out(idx) = v: idx = idx + 1: Next v
    For Each v In a2: out(idx) = v: idx = idx + 1: Next v
    For Each v In a3: out(idx) = v: idx = idx + 1: Next v
    MergeArrays3 = out
End Function

' Strip trailing paragraph mark (vbCr / Chr(13)) from range text
Public Function StripParaMarkChar(ByVal txt As String) As String
    If Len(txt) > 0 Then
        If Right$(txt, 1) = vbCr Or Right$(txt, 1) = Chr$(13) Then
            StripParaMarkChar = Left$(txt, Len(txt) - 1)
        Else
            StripParaMarkChar = txt
        End If
    Else
        StripParaMarkChar = txt
    End If
End Function

' Calculate the length of auto-generated list numbering text
' that appears in Range.Text but doesn't map to document positions.
Public Function GetListPrefixLen(para As Paragraph, ByVal paraText As String) As Long
    GetListPrefixLen = 0
    On Error Resume Next
    Dim lStr As String
    lStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0
    If Len(lStr) = 0 Then Exit Function
    If Len(paraText) > Len(lStr) Then
        If Left$(paraText, Len(lStr)) = lStr Then
            GetListPrefixLen = Len(lStr)
        End If
    End If
End Function
