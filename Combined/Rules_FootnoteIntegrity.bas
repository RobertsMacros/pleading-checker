Attribute VB_Name = "Rules_FootnoteIntegrity"
' ============================================================
' Rules_FootnoteIntegrity.bas
' Proofreading rule: checks footnote and endnote integrity.
'
' Checks performed:
'   1. Sequential numbering — no gaps in index sequence
'   2. Placement after punctuation — reference marks should
'      follow punctuation, not letters or spaces
'   3. Empty footnotes — footnotes with no content
'   4. Duplicate content — two footnotes with identical text
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "footnote_integrity"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_FootnoteIntegrity(doc As Document) As Collection
    Dim issues As New Collection

    ' ── Check footnotes ─────────────────────────────────────
    If doc.Footnotes.Count > 0 Then
        CheckNoteSequence doc, doc.Footnotes, "Footnote", issues
        CheckNotePlacement doc, doc.Footnotes, "Footnote", issues
        CheckEmptyNotes doc, doc.Footnotes, "Footnote", issues
        CheckDuplicateNotes doc, doc.Footnotes, "Footnote", issues
    End If

    ' ── Check endnotes ──────────────────────────────────────
    If doc.Endnotes.Count > 0 Then
        CheckEndnoteSequence doc, doc.Endnotes, "Endnote", issues
        CheckEndnotePlacement doc, doc.Endnotes, "Endnote", issues
        CheckEmptyEndnotes doc, doc.Endnotes, "Endnote", issues
        CheckDuplicateEndnotes doc, doc.Endnotes, "Endnote", issues
    End If

    Set Check_FootnoteIntegrity = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check sequential numbering for footnotes
' ════════════════════════════════════════════════════════════
Private Sub CheckNoteSequence(doc As Document, _
                               notes As Footnotes, _
                               noteType As String, _
                               ByRef issues As Collection)
    Dim i As Long
    Dim expectedIdx As Long
    Dim fn As Footnote
    Dim issue As PleadingsIssue
    Dim locStr As String

    expectedIdx = 1

    For i = 1 To notes.Count
        Set fn = notes(i)

        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(fn.Reference) Then
            expectedIdx = expectedIdx + 1
            On Error GoTo 0
            GoTo NextFootnote
        End If
        On Error GoTo 0

        If fn.Index <> expectedIdx Then
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       noteType & " numbering gap: expected " & expectedIdx & ", found " & fn.Index, _
                       "Renumber " & LCase(noteType) & "s sequentially", _
                       fn.Reference.Start, _
                       fn.Reference.End, _
                       "error"
            issues.Add issue
        End If

        expectedIdx = expectedIdx + 1

NextFootnote:
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check sequential numbering for endnotes
' ════════════════════════════════════════════════════════════
Private Sub CheckEndnoteSequence(doc As Document, _
                                  notes As Endnotes, _
                                  noteType As String, _
                                  ByRef issues As Collection)
    Dim i As Long
    Dim expectedIdx As Long
    Dim en As Endnote
    Dim issue As PleadingsIssue
    Dim locStr As String

    expectedIdx = 1

    For i = 1 To notes.Count
        Set en = notes(i)

        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(en.Reference) Then
            expectedIdx = expectedIdx + 1
            On Error GoTo 0
            GoTo NextEndnoteSeq
        End If
        On Error GoTo 0

        If en.Index <> expectedIdx Then
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(en.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       noteType & " numbering gap: expected " & expectedIdx & ", found " & en.Index, _
                       "Renumber " & LCase(noteType) & "s sequentially", _
                       en.Reference.Start, _
                       en.Reference.End, _
                       "error"
            issues.Add issue
        End If

        expectedIdx = expectedIdx + 1

NextEndnoteSeq:
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check placement after punctuation for footnotes
' ════════════════════════════════════════════════════════════
Private Sub CheckNotePlacement(doc As Document, _
                                notes As Footnotes, _
                                noteType As String, _
                                ByRef issues As Collection)
    Dim i As Long
    Dim fn As Footnote
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim charBefore As String
    Dim refStart As Long

    For i = 1 To notes.Count
        Set fn = notes(i)

        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFnPlace
        End If
        On Error GoTo 0

        refStart = fn.Reference.Start

        ' Check character before the reference mark
        If refStart > 0 Then
            On Error Resume Next
            charBefore = doc.Range(refStart - 1, refStart).Text
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo NextFnPlace
            End If
            On Error GoTo 0

            If Not IsPunctuation(charBefore) Then
                On Error Resume Next
                locStr = PleadingsEngine.GetLocationString(fn.Reference, doc)
                If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
                On Error GoTo 0

                Set issue = New PleadingsIssue
                issue.Init RULE_NAME, _
                           locStr, _
                           noteType & " " & fn.Index & " reference not placed after punctuation", _
                           "Place " & LCase(noteType) & " reference after punctuation mark", _
                           fn.Reference.Start, _
                           fn.Reference.End, _
                           "error"
                issues.Add issue
            End If
        End If

NextFnPlace:
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check placement after punctuation for endnotes
' ════════════════════════════════════════════════════════════
Private Sub CheckEndnotePlacement(doc As Document, _
                                   notes As Endnotes, _
                                   noteType As String, _
                                   ByRef issues As Collection)
    Dim i As Long
    Dim en As Endnote
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim charBefore As String
    Dim refStart As Long

    For i = 1 To notes.Count
        Set en = notes(i)

        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(en.Reference) Then
            On Error GoTo 0
            GoTo NextEnPlace
        End If
        On Error GoTo 0

        refStart = en.Reference.Start

        If refStart > 0 Then
            On Error Resume Next
            charBefore = doc.Range(refStart - 1, refStart).Text
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo NextEnPlace
            End If
            On Error GoTo 0

            If Not IsPunctuation(charBefore) Then
                On Error Resume Next
                locStr = PleadingsEngine.GetLocationString(en.Reference, doc)
                If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
                On Error GoTo 0

                Set issue = New PleadingsIssue
                issue.Init RULE_NAME, _
                           locStr, _
                           noteType & " " & en.Index & " reference not placed after punctuation", _
                           "Place " & LCase(noteType) & " reference after punctuation mark", _
                           en.Reference.Start, _
                           en.Reference.End, _
                           "error"
                issues.Add issue
            End If
        End If

NextEnPlace:
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check for empty footnotes
' ════════════════════════════════════════════════════════════
Private Sub CheckEmptyNotes(doc As Document, _
                             notes As Footnotes, _
                             noteType As String, _
                             ByRef issues As Collection)
    Dim i As Long
    Dim fn As Footnote
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim noteText As String

    For i = 1 To notes.Count
        Set fn = notes(i)

        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFnEmpty
        End If
        On Error GoTo 0

        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFnEmpty
        On Error GoTo 0

        noteText = Trim(Replace(noteText, vbCr, ""))
        noteText = Trim(Replace(noteText, vbLf, ""))

        If Len(noteText) = 0 Then
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       noteType & " " & fn.Index & " has empty content", _
                       "Add content or remove the empty " & LCase(noteType), _
                       fn.Reference.Start, _
                       fn.Reference.End, _
                       "error"
            issues.Add issue
        End If

NextFnEmpty:
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check for empty endnotes
' ════════════════════════════════════════════════════════════
Private Sub CheckEmptyEndnotes(doc As Document, _
                                notes As Endnotes, _
                                noteType As String, _
                                ByRef issues As Collection)
    Dim i As Long
    Dim en As Endnote
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim noteText As String

    For i = 1 To notes.Count
        Set en = notes(i)

        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(en.Reference) Then
            On Error GoTo 0
            GoTo NextEnEmpty
        End If
        On Error GoTo 0

        On Error Resume Next
        noteText = en.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextEnEmpty
        On Error GoTo 0

        noteText = Trim(Replace(noteText, vbCr, ""))
        noteText = Trim(Replace(noteText, vbLf, ""))

        If Len(noteText) = 0 Then
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(en.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       noteType & " " & en.Index & " has empty content", _
                       "Add content or remove the empty " & LCase(noteType), _
                       en.Reference.Start, _
                       en.Reference.End, _
                       "error"
            issues.Add issue
        End If

NextEnEmpty:
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check for duplicate footnote content
' ════════════════════════════════════════════════════════════
Private Sub CheckDuplicateNotes(doc As Document, _
                                 notes As Footnotes, _
                                 noteType As String, _
                                 ByRef issues As Collection)
    Dim contentDict As Object
    Set contentDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim fn As Footnote
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim noteText As String
    Dim cleanText As String

    For i = 1 To notes.Count
        Set fn = notes(i)

        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFnDup
        End If
        On Error GoTo 0

        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFnDup
        On Error GoTo 0

        cleanText = Trim(Replace(noteText, vbCr, ""))
        cleanText = Trim(Replace(cleanText, vbLf, ""))

        ' Skip empty notes (already flagged separately)
        If Len(cleanText) = 0 Then GoTo NextFnDup

        If contentDict.Exists(cleanText) Then
            ' This is a duplicate
            Dim firstIdx As Long
            firstIdx = CLng(contentDict(cleanText))

            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       noteType & " " & fn.Index & " has identical content to " & LCase(noteType) & " " & firstIdx, _
                       "Remove duplicate or differentiate content", _
                       fn.Reference.Start, _
                       fn.Reference.End, _
                       "possible_error"
            issues.Add issue
        Else
            contentDict.Add cleanText, fn.Index
        End If

NextFnDup:
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check for duplicate endnote content
' ════════════════════════════════════════════════════════════
Private Sub CheckDuplicateEndnotes(doc As Document, _
                                    notes As Endnotes, _
                                    noteType As String, _
                                    ByRef issues As Collection)
    Dim contentDict As Object
    Set contentDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim en As Endnote
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim noteText As String
    Dim cleanText As String

    For i = 1 To notes.Count
        Set en = notes(i)

        On Error Resume Next
        If Not PleadingsEngine.IsInPageRange(en.Reference) Then
            On Error GoTo 0
            GoTo NextEnDup
        End If
        On Error GoTo 0

        On Error Resume Next
        noteText = en.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextEnDup
        On Error GoTo 0

        cleanText = Trim(Replace(noteText, vbCr, ""))
        cleanText = Trim(Replace(cleanText, vbLf, ""))

        If Len(cleanText) = 0 Then GoTo NextEnDup

        If contentDict.Exists(cleanText) Then
            Dim firstEnIdx As Long
            firstEnIdx = CLng(contentDict(cleanText))

            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(en.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       noteType & " " & en.Index & " has identical content to " & LCase(noteType) & " " & firstEnIdx, _
                       "Remove duplicate or differentiate content", _
                       en.Reference.Start, _
                       en.Reference.End, _
                       "possible_error"
            issues.Add issue
        Else
            contentDict.Add cleanText, en.Index
        End If

NextEnDup:
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if character is punctuation
' ════════════════════════════════════════════════════════════
Private Function IsPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case ".", ",", ";", ":", """", "'", ")", _
             ChrW(8221), ChrW(8217), ChrW(8220), ChrW(8216), _
             "!", "?"
            IsPunctuation = True
        Case Else
            IsPunctuation = False
    End Select
End Function
