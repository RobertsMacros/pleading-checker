Attribute VB_Name = "Rule18_page_range"
' ============================================================
' Rule18_page-range.bas
' Utility rule: configures page range restriction for other rules.
' The form calls SetRange() to define the page window, and
' Check_PageRange() pushes those values into PleadingsEngine.
'
' Dependencies:
'   - PleadingsEngine.bas (SetPageRange)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "page_range"

' -- Module-level page range state ---------------------------
Private mStartPage As Long   ' 0 = no restriction
Private mEndPage   As Long   ' 0 = no restriction

' ============================================================
'  PUBLIC: SetRange
'  Called by the form to configure the page window before
'  rules are executed. Pass 0, 0 to clear the restriction.
' ============================================================
Public Sub SetRange(s As Long, e As Long)
    mStartPage = s
    mEndPage = e
End Sub

' ============================================================
'  MAIN ENTRY POINT
'  Pushes the configured page range into PleadingsEngine
'  so that IsInPageRange() respects the restriction.
'  Returns an empty Collection (this rule produces no issues).
' ============================================================
Public Function Check_PageRange(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' Push the stored page range into the engine
    EngineSetPageRange mStartPage, mEndPage

    On Error GoTo 0

    Set Check_PageRange = issues
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.SetPageRange
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.SetPageRange
' ----------------------------------------------------------------
Private Sub EngineSetPageRange(ByVal startPg As Long, ByVal endPg As Long)
    On Error Resume Next
    Application.Run "PleadingsEngine.SetPageRange", startPg, endPg
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub
