Attribute VB_Name = "Rules_Terms"
' ============================================================
' Rules_Terms.bas
' Term-related rules:
'   Rule05 - Custom term whitelist (populates shared whitelist)
'
' Previous rules (defined terms, phrase consistency) have been
' retired as part of the MVP pruning pass.
' ============================================================
Option Explicit

Private Const RULE05_NAME As String = "custom_term_whitelist"

' ============================================================
'  RULE 05: CUSTOM TERM WHITELIST
' ============================================================
Public Function Check_CustomTermWhitelist(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' -- Define default whitelist terms ----------------------
    Dim terms As Variant
    Dim batch1 As Variant, batch2 As Variant
    batch1 = Array( _
        "co-counsel", "anti-suit injunction", "pre-action", _
        "re-examination", "cross-examination", "counter-claim", _
        "sub-contract", "non-disclosure", "inter-partes", _
        "ex-parte", "bona fide")
    batch2 = Array( _
        "prima facie", "pro rata", "ad hoc", "de facto", _
        "de jure", "inter alia", "mutatis mutandis", _
        "pari passu", "ultra vires", "vis-a-vis")
    terms = MergeArrays2(batch1, batch2)

    ' -- Build the dictionary -------------------------------
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim t As Variant
    For Each t In terms
        Dim lcTerm As String
        lcTerm = LCase(CStr(t))
        If Not dict.Exists(lcTerm) Then
            dict.Add lcTerm, True
        End If
    Next t

    ' -- Store in the engine for other rules to query -------
    EngineSetWhitelist dict

    On Error GoTo 0

    ' This rule returns no issues -- it is purely a setup rule
    Set Check_CustomTermWhitelist = issues
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.SetWhitelist
' ----------------------------------------------------------------
Private Sub EngineSetWhitelist(dict As Object)
    On Error Resume Next
    Application.Run "PleadingsEngine.SetWhitelist", dict
    If Err.Number <> 0 Then
        Debug.Print "EngineSetWhitelist: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
'  Merge 2 Variant arrays into one flat Variant array
' ----------------------------------------------------------------
Private Function MergeArrays2(a1 As Variant, a2 As Variant) As Variant
    Dim total As Long
    total = UBound(a1) - LBound(a1) + 1 _
          + UBound(a2) - LBound(a2) + 1
    Dim out() As Variant
    ReDim out(0 To total - 1)
    Dim idx As Long
    idx = 0
    Dim v As Variant
    For Each v In a1: out(idx) = v: idx = idx + 1: Next v
    For Each v In a2: out(idx) = v: idx = idx + 1: Next v
    MergeArrays2 = out
End Function
