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
