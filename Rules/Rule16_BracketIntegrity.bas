Attribute VB_Name = "Rule16_BracketIntegrity"
' ============================================================
' Rule16_BracketIntegrity.bas
' Proofreading rule: checks for mismatched, unmatched, and
' improperly nested brackets: (), [], {}.
'
' Uses a stack-based approach with parallel arrays (since VBA
' Collections cannot hold UDTs). Skips brackets in code-font
' runs (Courier, Consolas).
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "bracket_integrity"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_BracketIntegrity(doc As Document) As Collection
    Dim issues As New Collection
    Dim docText As String
    Dim textLen As Long
    Dim i As Long
    Dim ch As String

    ' ── Stack using parallel arrays ──────────────────────────
    Dim stackChars() As String
    Dim stackPositions() As Long
    Dim stackTop As Long

    ReDim stackChars(0 To 1000)
    ReDim stackPositions(0 To 1000)
    stackTop = -1 ' empty stack

    ' ── Get full document text ───────────────────────────────
    On Error Resume Next
    docText = doc.Content.Text
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set Check_BracketIntegrity = issues
        Exit Function
    End If
    On Error GoTo 0

    textLen = Len(docText)
    If textLen = 0 Then
        Set Check_BracketIntegrity = issues
        Exit Function
    End If

    ' ── Iterate character by character ───────────────────────
    For i = 1 To textLen
        ch = Mid(docText, i, 1)

        ' Only process bracket characters
        If ch = "(" Or ch = "[" Or ch = "{" Or _
           ch = ")" Or ch = "]" Or ch = "}" Then

            ' Skip brackets in code-font runs
            If IsCodeFont(doc, i - 1) Then GoTo NextChar

            If ch = "(" Or ch = "[" Or ch = "{" Then
                ' ── Push opening bracket onto stack ──────────
                stackTop = stackTop + 1

                ' Grow arrays if needed
                If stackTop > UBound(stackChars) Then
                    ReDim Preserve stackChars(0 To stackTop + 500)
                    ReDim Preserve stackPositions(0 To stackTop + 500)
                End If

                stackChars(stackTop) = ch
                stackPositions(stackTop) = i - 1 ' 0-based doc position

            Else
                ' ── Closing bracket: pop and check match ─────
                If stackTop < 0 Then
                    ' Empty stack: unmatched closing bracket
                    CreateBracketIssue doc, issues, i - 1, ch, _
                        "Unmatched closing bracket '" & ch & "' with no corresponding opener"
                Else
                    Dim openChar As String
                    Dim openPos As Long
                    openChar = stackChars(stackTop)
                    openPos = stackPositions(stackTop)
                    stackTop = stackTop - 1

                    ' Check if bracket types match
                    If Not BracketsMatch(openChar, ch) Then
                        ' Mismatched bracket pair
                        CreateBracketIssue doc, issues, openPos, openChar, _
                            "Mismatched bracket: opened with '" & openChar & _
                            "' but closed with '" & ch & "'"
                        CreateBracketIssue doc, issues, i - 1, ch, _
                            "Mismatched bracket: closing '" & ch & _
                            "' does not match opener '" & openChar & "'"
                    End If
                End If
            End If
        End If

NextChar:
    Next i

    ' ── Any remaining on stack are unmatched openers ─────────
    Dim s As Long
    For s = 0 To stackTop
        CreateBracketIssue doc, issues, stackPositions(s), stackChars(s), _
            "Unmatched opening bracket '" & stackChars(s) & "' with no corresponding closer"
    Next s

    Set Check_BracketIntegrity = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if brackets match
' ════════════════════════════════════════════════════════════
Private Function BracketsMatch(ByVal openCh As String, _
                                ByVal closeCh As String) As Boolean
    Select Case openCh
        Case "("
            BracketsMatch = (closeCh = ")")
        Case "["
            BracketsMatch = (closeCh = "]")
        Case "{"
            BracketsMatch = (closeCh = "}")
        Case Else
            BracketsMatch = False
    End Select
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if position is in a code font
' ════════════════════════════════════════════════════════════
Private Function IsCodeFont(doc As Document, ByVal pos As Long) As Boolean
    Dim rng As Range
    Dim fontName As String

    On Error Resume Next
    Set rng = doc.Range(pos, pos + 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsCodeFont = False
        Exit Function
    End If

    fontName = ""
    fontName = rng.Font.Name
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsCodeFont = False
        Exit Function
    End If
    On Error GoTo 0

    IsCodeFont = (LCase(fontName) Like "*courier*") Or _
                 (LCase(fontName) Like "*consolas*")
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Create a bracket integrity issue
' ════════════════════════════════════════════════════════════
Private Sub CreateBracketIssue(doc As Document, _
                                ByRef issues As Collection, _
                                ByVal pos As Long, _
                                ByVal bracketChar As String, _
                                ByVal issueText As String)
    Dim issue As PleadingsIssue
    Dim locStr As String
    Dim rng As Range

    On Error Resume Next
    Set rng = doc.Range(pos, pos + 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    ' Skip if outside page range
    If Not PleadingsEngine.IsInPageRange(rng) Then
        On Error GoTo 0
        Exit Sub
    End If

    locStr = PleadingsEngine.GetLocationString(rng, doc)
    If Err.Number <> 0 Then
        locStr = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0

    ' Determine suggestion based on bracket type
    Dim suggestion As String
    If bracketChar = "(" Or bracketChar = ")" Then
        suggestion = "Add or correct matching parenthesis"
    ElseIf bracketChar = "[" Or bracketChar = "]" Then
        suggestion = "Add or correct matching square bracket"
    ElseIf bracketChar = "{" Or bracketChar = "}" Then
        suggestion = "Add or correct matching curly brace"
    Else
        suggestion = "Review bracket pairing"
    End If

    Set issue = New PleadingsIssue
    issue.Init RULE_NAME, _
               locStr, _
               issueText, _
               suggestion, _
               pos, _
               pos + 1, _
               "error"
    issues.Add issue
End Sub
