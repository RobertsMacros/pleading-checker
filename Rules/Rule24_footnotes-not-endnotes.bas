Attribute VB_Name = "Rule24_footnotes_not_endnotes"
' ============================================================
' Rule24_footnotes-not-endnotes.bas
' Proofreading rule: flags documents that use endnotes instead
' of (or in addition to) footnotes.
'
' Hart's Rules require footnotes rather than endnotes.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "footnotes_not_endnotes"

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_FootnotesNotEndnotes(doc As Document) As Collection
    Dim issues As New Collection
    Dim issue As Object

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
        Set issue = CreateIssueDict(RULE_NAME, "document level", "Document uses endnotes instead of footnotes.", "Use footnotes rather than endnotes.", 0, 0, "error", False)
        issues.Add issue

    ElseIf endCount > 0 And fnCount > 0 Then
        ' Document uses both
        Set issue = CreateIssueDict(RULE_NAME, "document level", "Document uses both footnotes and endnotes.", "Use footnotes rather than endnotes.", 0, 0, "error", False)
        issues.Add issue
    End If

    ' If only footnotes exist (endCount = 0): no issue

    Set Check_FootnotesNotEndnotes = issues
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
