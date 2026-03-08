Attribute VB_Name = "Rule15_ListPunctuation"
' ============================================================
' Rule15_ListPunctuation.bas
' Proofreading rule: checks punctuation consistency at the end
' of list items (semicolon, full stop, comma, colon, or none).
'
' Also checks:
'   - Last item should end with a full stop when dominant is
'     semicolon
'   - Penultimate item should include "and" or "or" before
'     its terminal punctuation
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "list_punctuation"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_ListPunctuation(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim totalParas As Long

    ' ── Collect all paragraphs into arrays for easier processing ─
    totalParas = doc.Paragraphs.Count
    If totalParas = 0 Then
        Set Check_ListPunctuation = issues
        Exit Function
    End If

    ' Arrays to hold paragraph info
    Dim paraStarts() As Long
    Dim paraEnds() As Long
    Dim paraTexts() As String
    Dim paraIsList() As Boolean
    Dim paraListID() As Long

    ReDim paraStarts(1 To totalParas)
    ReDim paraEnds(1 To totalParas)
    ReDim paraTexts(1 To totalParas)
    ReDim paraIsList(1 To totalParas)
    ReDim paraListID(1 To totalParas)

    paraIdx = 0
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        paraIdx = paraIdx + 1

        Dim paraRange As Range
        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            paraStarts(paraIdx) = 0
            paraEnds(paraIdx) = 0
            paraTexts(paraIdx) = ""
            paraIsList(paraIdx) = False
            paraListID(paraIdx) = 0
            GoTo NextParaCollect
        End If

        paraStarts(paraIdx) = paraRange.Start
        paraEnds(paraIdx) = paraRange.End
        paraTexts(paraIdx) = paraRange.Text

        ' Check if paragraph is a list item
        Dim listType As Long
        listType = 0
        listType = paraRange.ListFormat.ListType
        If Err.Number <> 0 Then
            Err.Clear
            listType = 0
        End If

        paraIsList(paraIdx) = (listType <> 0) ' 0 = wdListNoNumbering

        ' Get a list identifier for grouping
        Dim listID As Long
        listID = 0
        If paraIsList(paraIdx) Then
            listID = paraRange.ListFormat.List.ListParagraphs.Count
            If Err.Number <> 0 Then
                Err.Clear
                ' Fallback: use list level + approximate position
                listID = paraRange.ListFormat.ListLevelNumber + 1
                If Err.Number <> 0 Then
                    Err.Clear
                    listID = 1
                End If
            End If
        End If
        paraListID(paraIdx) = listID

NextParaCollect:
    Next para
    On Error GoTo 0

    ' ── Group consecutive list paragraphs into lists ─────────
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim inGroup As Boolean

    inGroup = False
    Dim p As Long

    For p = 1 To totalParas
        If paraIsList(p) Then
            If Not inGroup Then
                groupStart = p
                inGroup = True
            End If
            groupEnd = p
        Else
            If inGroup Then
                ' Process the list group
                ProcessListGroup doc, issues, paraStarts, paraEnds, paraTexts, _
                                 groupStart, groupEnd
                inGroup = False
            End If
        End If
    Next p

    ' Process final group if document ends with a list
    If inGroup Then
        ProcessListGroup doc, issues, paraStarts, paraEnds, paraTexts, _
                         groupStart, groupEnd
    End If

    Set Check_ListPunctuation = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Process a single list group for punctuation issues
' ════════════════════════════════════════════════════════════
Private Sub ProcessListGroup(doc As Document, _
                              ByRef issues As Collection, _
                              ByRef paraStarts() As Long, _
                              ByRef paraEnds() As Long, _
                              ByRef paraTexts() As String, _
                              ByVal groupStart As Long, _
                              ByVal groupEnd As Long)
    Dim itemCount As Long
    Dim i As Long
    Dim endings() As String
    Dim endingCounts As Object ' Dictionary
    Dim dominantEnding As String
    Dim maxCount As Long

    itemCount = groupEnd - groupStart + 1
    If itemCount < 2 Then Exit Sub ' Single-item list, nothing to check

    ' ── Classify the ending of each list item ────────────────
    ReDim endings(groupStart To groupEnd)

    For i = groupStart To groupEnd
        endings(i) = ClassifyEnding(paraTexts(i))
    Next i

    ' ── Count endings to find dominant ───────────────────────
    Set endingCounts = CreateObject("Scripting.Dictionary")

    For i = groupStart To groupEnd
        If endingCounts.Exists(endings(i)) Then
            endingCounts(endings(i)) = endingCounts(endings(i)) + 1
        Else
            endingCounts.Add endings(i), 1
        End If
    Next i

    dominantEnding = ""
    maxCount = 0
    Dim key As Variant
    For Each key In endingCounts.keys
        If endingCounts(key) > maxCount Then
            maxCount = endingCounts(key)
            dominantEnding = CStr(key)
        End If
    Next key

    ' ── Flag items that deviate from dominant ending ──────────
    For i = groupStart To groupEnd
        If endings(i) <> dominantEnding Then
            ' Skip the last item if dominant is semicolon (special rule below)
            If dominantEnding = "semicolon" And i = groupEnd Then
                GoTo ContinueItem
            End If

            Dim rng As Range
            Dim locStr As String
            Dim issue As PleadingsIssue

            On Error Resume Next
            Set rng = doc.Range(paraStarts(i), paraEnds(i))
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo ContinueItem
            End If

            If Not PleadingsEngine.IsInPageRange(rng) Then
                On Error GoTo 0
                GoTo ContinueItem
            End If

            locStr = PleadingsEngine.GetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       "List item ending '" & endings(i) & "' differs from " & _
                       "dominant ending '" & dominantEnding & "'", _
                       "Change ending punctuation to match list style (" & _
                       dominantEnding & ")", _
                       paraStarts(i), _
                       paraEnds(i), _
                       "possible_error"
            issues.Add issue
        End If

ContinueItem:
    Next i

    ' ── Special: if dominant is semicolon, last item should end with full stop ─
    If dominantEnding = "semicolon" Then
        If endings(groupEnd) <> "full_stop" Then
            On Error Resume Next
            Set rng = doc.Range(paraStarts(groupEnd), paraEnds(groupEnd))
            If Err.Number = 0 Then
                If PleadingsEngine.IsInPageRange(rng) Then
                    locStr = PleadingsEngine.GetLocationString(rng, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If

                    Set issue = New PleadingsIssue
                    issue.Init RULE_NAME, _
                               locStr, _
                               "Last list item should end with a full stop, not '" & _
                               endings(groupEnd) & "'", _
                               "End the final list item with a full stop", _
                               paraStarts(groupEnd), _
                               paraEnds(groupEnd), _
                               "possible_error"
                    issues.Add issue
                End If
            End If
            On Error GoTo 0
        End If

        ' ── Check penultimate item for "and" or "or" ─────────
        If itemCount >= 2 Then
            Dim penIdx As Long
            penIdx = groupEnd - 1
            Dim penText As String
            penText = LCase(Trim(StripTrailingCr(paraTexts(penIdx))))

            Dim hasConjunction As Boolean
            hasConjunction = False

            ' Check if text ends with "and;" or "or;" or similar
            If Right(penText, 4) = "and;" Or Right(penText, 3) = "or;" Or _
               Right(penText, 4) = "and," Or Right(penText, 3) = "or," Or _
               Right(penText, 3) = "and" Or Right(penText, 2) = "or" Then
                hasConjunction = True
            End If

            ' Also check for "and" / "or" as last word before punctuation
            Dim lastWords As String
            lastWords = GetLastNChars(penText, 10)
            If InStr(1, lastWords, " and") > 0 Or InStr(1, lastWords, " or") > 0 Then
                hasConjunction = True
            End If

            If Not hasConjunction Then
                On Error Resume Next
                Set rng = doc.Range(paraStarts(penIdx), paraEnds(penIdx))
                If Err.Number = 0 Then
                    If PleadingsEngine.IsInPageRange(rng) Then
                        locStr = PleadingsEngine.GetLocationString(rng, doc)
                        If Err.Number <> 0 Then
                            locStr = "unknown location"
                            Err.Clear
                        End If

                        Set issue = New PleadingsIssue
                        issue.Init RULE_NAME, _
                                   locStr, _
                                   "Penultimate list item should include 'and' or 'or' " & _
                                   "before terminal punctuation", _
                                   "Add 'and' or 'or' before the semicolon", _
                                   paraStarts(penIdx), _
                                   paraEnds(penIdx), _
                                   "possible_error"
                        issues.Add issue
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    End If
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Classify the ending punctuation of a list item
' ════════════════════════════════════════════════════════════
Private Function ClassifyEnding(ByVal text As String) As String
    Dim trimmed As String
    Dim lastChar As String

    trimmed = StripTrailingCr(text)
    trimmed = Trim(trimmed)

    If Len(trimmed) = 0 Then
        ClassifyEnding = "none"
        Exit Function
    End If

    lastChar = Right(trimmed, 1)

    Select Case lastChar
        Case ";"
            ClassifyEnding = "semicolon"
        Case "."
            ClassifyEnding = "full_stop"
        Case ","
            ClassifyEnding = "comma"
        Case ":"
            ClassifyEnding = "colon"
        Case Else
            ClassifyEnding = "none"
    End Select
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Strip trailing carriage return / line feed
' ════════════════════════════════════════════════════════════
Private Function StripTrailingCr(ByVal text As String) As String
    Dim result As String
    result = text

    Do While Len(result) > 0
        Dim lastCh As String
        lastCh = Right(result, 1)
        If lastCh = vbCr Or lastCh = vbLf Or lastCh = Chr(13) Or lastCh = Chr(10) Then
            result = Left(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop

    StripTrailingCr = result
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Get last N characters of a string
' ════════════════════════════════════════════════════════════
Private Function GetLastNChars(ByVal text As String, ByVal n As Long) As String
    If Len(text) <= n Then
        GetLastNChars = text
    Else
        GetLastNChars = Right(text, n)
    End If
End Function
