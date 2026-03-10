Attribute VB_Name = "Rule13_colour_formatting"
' ============================================================
' Rule13_colour-formatting.bas
' Proofreading rule: detects non-standard font colours in the
' document body. Identifies the dominant text colour and flags
' any runs using a different colour (excluding hyperlinks and
' heading-styled paragraphs).
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "colour_formatting"

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_ColourFormatting(doc As Document) As Collection
    Dim issues As New Collection
    Dim colourCounts As Object ' Scripting.Dictionary
    Dim para As Paragraph
    Dim rn As Range
    Dim runColor As Long
    Dim dominantColour As Long
    Dim maxCount As Long
    Dim runText As String

    ' -- First pass: count colour usage per run ---------------
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
        If Not EngineIsInPageRange(paraRange) Then
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

    ' -- Determine dominant colour ----------------------------
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

    ' -- Second pass: flag non-dominant, non-automatic colours -
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
        If Not EngineIsInPageRange(paraRange) Then
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

            ' -- This run has a non-standard colour -----------
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

' ============================================================
'  PRIVATE: Flush a grouped colour issue
' ============================================================
Private Sub FlushColourGroup(doc As Document, _
                              ByRef issues As Collection, _
                              ByVal startPos As Long, _
                              ByVal endPos As Long, _
                              ByVal fontColor As Long)
    Dim issue As Object
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

    locStr = EngineGetLocationString(rng, doc)
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

    Set issue = CreateIssueDict(RULE_NAME, locStr, "Non-standard font colour " & hexStr & " detected: '" & previewText & "'", "Change font colour to match document default", startPos, endPos, "possible_error")
    issues.Add issue
End Sub

' ============================================================
'  PRIVATE: Convert a Long colour value to hex string
' ============================================================
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

' ============================================================
'  PRIVATE: Check if a run is inside a hyperlink
' ============================================================
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

' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based issue (no class dependency)
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
'  Late-bound wrapper: EngineIsInPageRange
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  Late-bound wrapper: EngineGetLocationString
' ----------------------------------------------------------------

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
