Attribute VB_Name = "Rule03_SequentialNumbering"
' ============================================================
' Rule03_SequentialNumbering.bas
' Proofreading rule: verifies that numbered lists maintain
' correct sequential ordering within each list context.
'
' Checks both:
'   1. Word-native list formatting (ListFormat objects)
'   2. Manually typed numbering (e.g. "1.", "2.", "3." at
'      paragraph start without Word list formatting)
'
' Detects: skipped items, backwards numbering, duplicates.
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
'   - Microsoft Scripting Runtime (Scripting.Dictionary)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "sequential_numbering"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_SequentialNumbering(doc As Document) As Collection
    Dim issues As New Collection

    ' ── Check Word-native numbered lists ──────────────────
    CheckNativeListNumbering doc, issues

    ' ── Check manually typed numbering ────────────────────
    CheckManualNumbering doc, issues

    Set Check_SequentialNumbering = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check Word-native list numbering
'  Uses a Scripting.Dictionary keyed by list identifier to
'  track expected next values per list and level.
'
'  Each top-level key maps to a Dictionary of levels, where
'  each level stores the expected next value.
' ════════════════════════════════════════════════════════════
Private Sub CheckNativeListNumbering(doc As Document, _
                                      ByRef issues As Collection)
    Dim listContexts As New Scripting.Dictionary  ' listKey -> Dictionary(level -> expectedNext)
    Dim para As Paragraph
    Dim paraRange As Range
    Dim listType As Long
    Dim listKey As String
    Dim listLevel As Long
    Dim listValue As Long
    Dim expectedNext As Long
    Dim levelDict As Scripting.Dictionary
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim issueText As String
    Dim suggestion As String
    Dim prevLevel As Long

    ' Track the previous level per list to detect level changes
    Dim prevLevelDict As New Scripting.Dictionary  ' listKey -> prevLevel

    On Error Resume Next

    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextPara
        End If

        ' ── Skip non-list paragraphs ─────────────────────
        listType = paraRange.ListFormat.listType
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextPara
        End If

        ' wdListNoNumbering = 0; skip these
        If listType = 0 Then GoTo NextPara

        ' Only check numbered lists (wdListSimpleNumbering=1,
        ' wdListOutlineNumbering=4, wdListMixedNumbering=5)
        ' Skip bullet lists (wdListBullet=2, wdListPictureBullet=6)
        If listType = 2 Or listType = 6 Then GoTo NextPara

        ' ── Skip if outside configured page range ────────
        If Not PleadingsEngine.IsInPageRange(paraRange) Then
            GoTo NextPara
        End If

        ' ── Determine list key (unique identifier) ───────
        ' Try to use the List object's ListID first; fall back
        ' to a synthetic key built from type + position.
        listKey = ""
        Err.Clear
        Dim lstObj As Object
        Set lstObj = paraRange.ListFormat.List
        If Err.Number = 0 And Not lstObj Is Nothing Then
            listKey = "List_" & CStr(ObjPtr(lstObj))
        End If
        If Err.Number <> 0 Or Len(listKey) = 0 Then
            Err.Clear
            ' Synthetic key: use list type and an approximation
            listKey = "Synth_" & CStr(listType) & "_" & CStr(paraRange.ListFormat.ListLevelNumber)
        End If
        Err.Clear

        ' ── Get current list value and level ─────────────
        listValue = paraRange.ListFormat.listValue
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextPara
        End If

        listLevel = paraRange.ListFormat.ListLevelNumber
        If Err.Number <> 0 Then
            Err.Clear
            listLevel = 1
        End If

        ' ── Initialise tracking for this list if new ─────
        If Not listContexts.Exists(listKey) Then
            Dim newLevelDict As New Scripting.Dictionary
            listContexts.Add listKey, newLevelDict
            prevLevelDict.Add listKey, 0
        End If

        Set levelDict = listContexts(listKey)
        prevLevel = prevLevelDict(listKey)

        ' ── Handle level changes ─────────────────────────
        ' When we go to a deeper level, that level starts fresh.
        ' When we return to a shallower level, reset all deeper levels.
        If listLevel <> prevLevel And prevLevel > 0 Then
            If listLevel < prevLevel Then
                ' Returning to shallower level: reset deeper levels
                Dim resetLevel As Variant
                Dim keysToRemove As New Collection
                For Each resetLevel In levelDict.keys
                    If CLng(resetLevel) > listLevel Then
                        keysToRemove.Add resetLevel
                    End If
                Next resetLevel
                Dim removeIdx As Long
                For removeIdx = 1 To keysToRemove.Count
                    levelDict.Remove keysToRemove(removeIdx)
                Next removeIdx
                Set keysToRemove = Nothing
            End If
        End If

        ' ── Check expected value at this level ───────────
        If Not levelDict.Exists(listLevel) Then
            ' First item at this level in this list; record starting value
            levelDict.Add listLevel, listValue + 1
        Else
            expectedNext = levelDict(listLevel)

            If listValue = expectedNext Then
                ' Correct sequence; update expected next
                levelDict(listLevel) = listValue + 1

            ElseIf listValue = expectedNext - 1 Then
                ' Duplicate number
                Err.Clear
                locStr = PleadingsEngine.GetLocationString(paraRange, doc)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                End If

                issueText = "Duplicate number " & listValue & " at level " & listLevel
                suggestion = "Expected " & expectedNext & "; remove or renumber the duplicate"

                Set issue = New PleadingsIssue
                issue.Init RULE_NAME, locStr, issueText, suggestion, _
                           paraRange.Start, paraRange.End, "error"
                issues.Add issue
                ' Do not advance expectedNext for duplicates

            ElseIf listValue > expectedNext Then
                ' Skipped item(s)
                Err.Clear
                locStr = PleadingsEngine.GetLocationString(paraRange, doc)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                End If

                issueText = "Expected " & expectedNext & " but found " & listValue & _
                            " -- possible skipped item(s)"
                suggestion = "Check whether items " & expectedNext & " through " & _
                             (listValue - 1) & " are missing"

                Set issue = New PleadingsIssue
                issue.Init RULE_NAME, locStr, issueText, suggestion, _
                           paraRange.Start, paraRange.End, "error"
                issues.Add issue

                ' Update expected to continue from current
                levelDict(listLevel) = listValue + 1

            ElseIf listValue < expectedNext - 1 Then
                ' Numbering went backwards
                Err.Clear
                locStr = PleadingsEngine.GetLocationString(paraRange, doc)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                End If

                issueText = "Expected " & expectedNext & " but found " & listValue & _
                            " -- numbering went backwards"
                suggestion = "Renumber this item to " & expectedNext & " or check list continuity"

                Set issue = New PleadingsIssue
                issue.Init RULE_NAME, locStr, issueText, suggestion, _
                           paraRange.Start, paraRange.End, "error"
                issues.Add issue

                ' Update expected to continue from current
                levelDict(listLevel) = listValue + 1
            Else
                ' Normal sequence
                levelDict(listLevel) = listValue + 1
            End If
        End If

        ' Record previous level for this list
        prevLevelDict(listKey) = listLevel

NextPara:
    Next para
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check manually typed numbering
'  Detects paragraphs that start with a number pattern
'  (e.g. "1.", "2.", "12.3") but have no Word list formatting.
'  Tracks these separately and checks for sequence breaks.
' ════════════════════════════════════════════════════════════
Private Sub CheckManualNumbering(doc As Document, _
                                  ByRef issues As Collection)
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim listType As Long
    Dim manualNum As Long
    Dim expectedNext As Long
    Dim tracking As Boolean
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim issueText As String
    Dim suggestion As String

    expectedNext = 0
    tracking = False

    On Error Resume Next

    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextManualPara
        End If

        paraText = Trim(paraRange.Text)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextManualPara
        End If

        ' ── Only process non-list paragraphs ─────────────
        listType = paraRange.ListFormat.listType
        If Err.Number <> 0 Then
            Err.Clear
            listType = 0
        End If

        ' If this paragraph has Word list formatting, skip it
        ' and break any manual tracking chain
        If listType <> 0 Then
            tracking = False
            expectedNext = 0
            GoTo NextManualPara
        End If

        ' ── Check if paragraph starts with a number pattern ─
        ' Patterns: "N." or "N)" where N is one or more digits
        manualNum = ExtractLeadingNumber(paraText)

        If manualNum < 0 Then
            ' No number pattern found; break tracking chain
            ' but only if the paragraph has substantial text
            ' (skip blank lines to allow gaps between items)
            If Len(paraText) > 1 Then
                tracking = False
                expectedNext = 0
            End If
            GoTo NextManualPara
        End If

        ' ── Skip if outside configured page range ────────
        If Not PleadingsEngine.IsInPageRange(paraRange) Then
            GoTo NextManualPara
        End If

        ' ── Start or continue tracking ───────────────────
        If Not tracking Then
            ' First manually numbered paragraph in a sequence
            tracking = True
            expectedNext = manualNum + 1
            GoTo NextManualPara
        End If

        ' ── Check sequence ───────────────────────────────
        If manualNum = expectedNext Then
            ' Correct sequence
            expectedNext = manualNum + 1

        ElseIf manualNum > expectedNext Then
            ' Skipped item(s)
            Err.Clear
            locStr = PleadingsEngine.GetLocationString(paraRange, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If

            issueText = "Manual numbering: expected " & expectedNext & _
                        " but found " & manualNum & " -- possible skipped item(s)"
            suggestion = "Check whether items " & expectedNext & " through " & _
                         (manualNum - 1) & " are missing"

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, locStr, issueText, suggestion, _
                       paraRange.Start, paraRange.End, "error"
            issues.Add issue

            expectedNext = manualNum + 1

        ElseIf manualNum < expectedNext And manualNum = expectedNext - 1 Then
            ' Duplicate
            Err.Clear
            locStr = PleadingsEngine.GetLocationString(paraRange, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If

            issueText = "Manual numbering: duplicate number " & manualNum
            suggestion = "Remove or renumber the duplicate item"

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, locStr, issueText, suggestion, _
                       paraRange.Start, paraRange.End, "error"
            issues.Add issue

        ElseIf manualNum < expectedNext - 1 Then
            ' Backwards
            Err.Clear
            locStr = PleadingsEngine.GetLocationString(paraRange, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If

            issueText = "Manual numbering: expected " & expectedNext & _
                        " but found " & manualNum & " -- numbering went backwards"
            suggestion = "Renumber this item to " & expectedNext & " or check sequence"

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, locStr, issueText, suggestion, _
                       paraRange.Start, paraRange.End, "error"
            issues.Add issue

            expectedNext = manualNum + 1
        Else
            ' Normal (covers any other case)
            expectedNext = manualNum + 1
        End If

NextManualPara:
    Next para
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Extract leading number from paragraph text
'  Returns the number if the text starts with a pattern like
'  "1.", "12.", "3)", "42)"; returns -1 if no match.
'  Uses the VBA Like operator for pattern matching.
' ════════════════════════════════════════════════════════════
Private Function ExtractLeadingNumber(ByVal txt As String) As Long
    Dim trimmed As String
    Dim numStr As String
    Dim i As Long
    Dim ch As String

    trimmed = Trim(txt)
    ExtractLeadingNumber = -1

    If Len(trimmed) = 0 Then Exit Function

    ' Check first character is a digit
    If Not (trimmed Like "#*") Then Exit Function

    ' Extract consecutive digits from the start
    numStr = ""
    For i = 1 To Len(trimmed)
        ch = Mid(trimmed, i, 1)
        If ch >= "0" And ch <= "9" Then
            numStr = numStr & ch
        Else
            Exit For
        End If
    Next i

    ' Must have at least one digit
    If Len(numStr) = 0 Then Exit Function

    ' The character after the digits must be "." or ")"
    ' to qualify as a numbering pattern
    If i <= Len(trimmed) Then
        ch = Mid(trimmed, i, 1)
        If ch = "." Or ch = ")" Then
            On Error Resume Next
            ExtractLeadingNumber = CLng(numStr)
            If Err.Number <> 0 Then
                ExtractLeadingNumber = -1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunSequentialNumbering()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Sequential Numbering"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_SequentialNumbering(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Sequential Numbering"
End Sub
