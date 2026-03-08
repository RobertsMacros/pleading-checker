Attribute VB_Name = "Rule05_CustomTermWhitelist"
' ============================================================
' Rule05_CustomTermWhitelist.bas
' Utility rule that populates the PleadingsEngine whitelist
' with standard legal/Latin terms. Does not find issues itself.
' Other rules can call PleadingsEngine.IsWhitelistedTerm()
' to skip flagging these accepted terms.
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "custom_term_whitelist"

' ════════════════════════════════════════════════════════════
'  MAIN RULE FUNCTION
' ════════════════════════════════════════════════════════════
Public Function Check_CustomTermWhitelist(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' ── Define default whitelist terms ──────────────────────
    Dim terms As Variant
    terms = Array( _
        "co-counsel", _
        "anti-suit injunction", _
        "pre-action", _
        "re-examination", _
        "cross-examination", _
        "counter-claim", _
        "sub-contract", _
        "non-disclosure", _
        "inter-partes", _
        "ex-parte", _
        "bona fide", _
        "prima facie", _
        "pro rata", _
        "ad hoc", _
        "de facto", _
        "de jure", _
        "inter alia", _
        "mutatis mutandis", _
        "pari passu", _
        "ultra vires", _
        "vis-a-vis" _
    )

    ' ── Build the dictionary ───────────────────────────────
    Dim dict As New Scripting.Dictionary
    Dim t As Variant
    For Each t In terms
        Dim lcTerm As String
        lcTerm = LCase(CStr(t))
        If Not dict.Exists(lcTerm) Then
            dict.Add lcTerm, True
        End If
    Next t

    ' ── Store in the engine for other rules to query ───────
    PleadingsEngine.SetWhitelist dict

    On Error GoTo 0

    ' This rule returns no issues — it is purely a setup rule
    Set Check_CustomTermWhitelist = issues
End Function
