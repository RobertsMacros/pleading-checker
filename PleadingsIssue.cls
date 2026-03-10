VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PleadingsIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' PleadingsIssue.cls
' Structured result object for the Pleadings Checker rule engine.
' Each instance represents a single issue found by a rule.
'
' Installation:
'   1. Open the VBA Editor (Alt+F11)
'   2. File > Import File > select PleadingsIssue.cls
' ============================================================
Option Explicit

' -- Properties ------------------------------------------------
Private pRuleName   As String   ' e.g. "british_spelling"
Private pLocation   As String   ' e.g. "page 14 paragraph 2"
Private pIssue      As String   ' e.g. "US spelling detected: 'color'"
Private pSuggestion As String   ' e.g. "colour"
Private pSeverity   As String   ' "error" or "possible_error"
Private pRangeStart As Long     ' character position for highlighting
Private pRangeEnd   As Long     ' character position end
Private pAutoFixSafe As Boolean ' whether auto-fix can be safely applied

' -- RuleName --------------------------------------------------
Public Property Get RuleName() As String
    RuleName = pRuleName
End Property

Public Property Let RuleName(ByVal value As String)
    pRuleName = value
End Property

' -- Location --------------------------------------------------
Public Property Get Location() As String
    Location = pLocation
End Property

Public Property Let Location(ByVal value As String)
    pLocation = value
End Property

' -- Issue -----------------------------------------------------
Public Property Get Issue() As String
    Issue = pIssue
End Property

Public Property Let Issue(ByVal value As String)
    pIssue = value
End Property

' -- Suggestion ------------------------------------------------
Public Property Get Suggestion() As String
    Suggestion = pSuggestion
End Property

Public Property Let Suggestion(ByVal value As String)
    pSuggestion = value
End Property

' -- Severity --------------------------------------------------
Public Property Get Severity() As String
    If pSeverity = "" Then
        Severity = "error"
    Else
        Severity = pSeverity
    End If
End Property

Public Property Let Severity(ByVal value As String)
    pSeverity = value
End Property

' -- RangeStart ------------------------------------------------
Public Property Get RangeStart() As Long
    RangeStart = pRangeStart
End Property

Public Property Let RangeStart(ByVal value As Long)
    pRangeStart = value
End Property

' -- RangeEnd --------------------------------------------------
Public Property Get RangeEnd() As Long
    RangeEnd = pRangeEnd
End Property

Public Property Let RangeEnd(ByVal value As Long)
    pRangeEnd = value
End Property

' -- AutoFixSafe -------------------------------------------------
Public Property Get AutoFixSafe() As Boolean
    AutoFixSafe = pAutoFixSafe
End Property

Public Property Let AutoFixSafe(ByVal value As Boolean)
    pAutoFixSafe = value
End Property

' -- Helper: initialise all properties at once -----------------
Public Sub Init(ByVal ruleName_ As String, _
                ByVal location_ As String, _
                ByVal issue_ As String, _
                ByVal suggestion_ As String, _
                ByVal rangeStart_ As Long, _
                ByVal rangeEnd_ As Long, _
                Optional ByVal severity_ As String = "error", _
                Optional ByVal autoFixSafe_ As Boolean = False)
    pRuleName = ruleName_
    pLocation = location_
    pIssue = issue_
    pSuggestion = suggestion_
    pRangeStart = rangeStart_
    pRangeEnd = rangeEnd_
    pSeverity = severity_
    pAutoFixSafe = autoFixSafe_
End Sub

' -- Helper: format as JSON object string ----------------------
Public Function ToJSON() As String
    Dim s As String
    s = "    {" & vbCrLf
    s = s & "      ""rule"": """ & EscJSON(pRuleName) & """," & vbCrLf
    s = s & "      ""location"": """ & EscJSON(pLocation) & """," & vbCrLf
    s = s & "      ""severity"": """ & EscJSON(Severity) & """," & vbCrLf
    s = s & "      ""issue"": """ & EscJSON(pIssue) & """," & vbCrLf
    s = s & "      ""suggestion"": """ & EscJSON(pSuggestion) & """," & vbCrLf
    s = s & "      ""auto_fix_safe"": " & IIf(pAutoFixSafe, "true", "false") & vbCrLf
    s = s & "    }"
    ToJSON = s
End Function

' -- Private: escape for JSON ----------------------------------
Private Function EscJSON(ByVal txt As String) As String
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, vbCr, "\r")
    txt = Replace(txt, vbLf, "\n")
    txt = Replace(txt, vbTab, "\t")
    EscJSON = txt
End Function
