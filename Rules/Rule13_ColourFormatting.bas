Attribute VB_Name = "Rule13_ColourFormatting"
' ============================================================
' Rule13_ColourFormatting.bas
' Proofreading rule: detects non-standard font colours in the
' document body. Identifies the dominant text colour and flags
' any runs using a different colour (excluding hyperlinks and
' heading-styled paragraphs).
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
'   - Microsoft Scripting Runtime (Scripting.Dictionary)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "colour_formatting"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
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
    issue.Init RULE_NAME, _
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

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunColourFormatting()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Colour Formatting"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_ColourFormatting(doc)

    ' ── Highlight issues in document ─────────────────────────
    Dim iss As PleadingsIssue
    Dim rng As Range
    Dim i As Long
    For i = 1 To issues.Count
        Set iss = issues(i)
        If iss.RangeStart >= 0 And iss.RangeEnd > iss.RangeStart Then
            On Error Resume Next
            Set rng = doc.Range(iss.RangeStart, iss.RangeEnd)
            rng.HighlightColorIndex = wdYellow
            doc.Comments.Add Range:=rng, _
                Text:="[" & iss.RuleName & "] " & iss.Issue & _
                      " " & Chr(8212) & " Suggestion: " & iss.Suggestion
            On Error GoTo 0
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Colour Formatting"
End Sub
