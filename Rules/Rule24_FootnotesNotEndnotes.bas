Attribute VB_Name = "Rule24_FootnotesNotEndnotes"
' ============================================================
' Rule24_FootnotesNotEndnotes.bas
' Proofreading rule: flags documents that use endnotes instead
' of (or in addition to) footnotes.
'
' Hart's Rules require footnotes rather than endnotes.
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "footnotes_not_endnotes"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_FootnotesNotEndnotes(doc As Document) As Collection
    Dim issues As New Collection
    Dim issue As PleadingsIssue

    On Error Resume Next

    Dim endCount As Long
    Dim fnCount As Long
    endCount = doc.Endnotes.Count
    fnCount = doc.Footnotes.Count

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set Check_FootnotesNotEndnotes = issues
        Exit Function
    End If
    On Error GoTo 0

    If endCount > 0 And fnCount = 0 Then
        ' Document uses only endnotes
        Set issue = New PleadingsIssue
        issue.Init RULE_NAME, _
                   "document level", _
                   "Document uses endnotes instead of footnotes.", _
                   "Use footnotes rather than endnotes.", _
                   0, _
                   0, _
                   "error", _
                   False
        issues.Add issue

    ElseIf endCount > 0 And fnCount > 0 Then
        ' Document uses both
        Set issue = New PleadingsIssue
        issue.Init RULE_NAME, _
                   "document level", _
                   "Document uses both footnotes and endnotes.", _
                   "Use footnotes rather than endnotes.", _
                   0, _
                   0, _
                   "error", _
                   False
        issues.Add issue
    End If

    ' If only footnotes exist (endCount = 0): no issue

    Set Check_FootnotesNotEndnotes = issues
End Function
