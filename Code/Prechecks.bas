Attribute VB_Name = "Prechecks"
' ============================================================
' Prechecks.bas
' Fast document-level pre-checks that can short-circuit
' expensive per-paragraph or per-note rule scans.
'
' Each precheck sets a module-level flag that the corresponding
' rule handler reads to decide whether to skip its work.
' Only call prechecks for rules that are actually enabled.
'
' Public API:
'   RunPrechecks(doc, config)   -- runs all relevant prechecks
'   SkipBracketIntegrity        -- True if brackets balance globally
'   SkipSpelling                -- True if no search terms in doc text
'   CachedDocText               -- full doc text (reused across prechecks)
' ============================================================
Option Explicit

' -- Cached document text (read once, reused across prechecks) --
Private mDocText As String
Private mDocTextLC As String     ' lower-case version for spelling
Private mHasDocText As Boolean

' -- Result flags -----------------------------------------------
Private mSkipBracketIntegrity As Boolean
Private mSkipSpelling As Boolean

' ============================================================
'  PUBLIC: RunPrechecks
'  Runs all applicable prechecks for enabled rules.
'  Call this once before the paragraph and footnote loops.
' ============================================================
Public Sub RunPrechecks(doc As Document, config As Object)
    ' Reset state
    mSkipBracketIntegrity = False
    mSkipSpelling = False
    mHasDocText = False
    mDocText = ""
    mDocTextLC = ""

    ' Cache document text (shared across prechecks)
    On Error Resume Next
    mDocText = doc.Content.Text
    If Err.Number <> 0 Then
        Err.Clear
        mDocText = ""
    End If
    On Error GoTo 0
    mHasDocText = (Len(mDocText) > 0)

    ' -- Bracket integrity precheck ----------------------------
    If PleadingsEngine.IsRuleEnabled(config, "punctuation") Then
        PrecheckBracketIntegrity
    End If

    ' -- Spelling precheck -------------------------------------
    If PleadingsEngine.IsRuleEnabled(config, "spellchecker") Then
        PrecheckSpelling
    End If
End Sub

' ============================================================
'  PUBLIC: Property accessors for result flags
' ============================================================

Public Property Get SkipBracketIntegrity() As Boolean
    SkipBracketIntegrity = mSkipBracketIntegrity
End Property

Public Property Get SkipSpelling() As Boolean
    SkipSpelling = mSkipSpelling
End Property

' ============================================================
'  PUBLIC: Cleanup -- free cached text after all rules complete
' ============================================================
Public Sub ClearPrechecks()
    mDocText = ""
    mDocTextLC = ""
    mHasDocText = False
    mSkipBracketIntegrity = False
    mSkipSpelling = False
End Sub

' ============================================================
'  PRIVATE: Bracket integrity precheck
'  Counts all (), [], {} in the full document text.
'  If all three types balance, per-paragraph scan is skipped.
' ============================================================
Private Sub PrecheckBracketIntegrity()
    If Not mHasDocText Then Exit Sub

    Dim gPO As Long, gPC As Long
    Dim gSO As Long, gSC As Long
    Dim gCO As Long, gCC As Long
    Dim gBytes() As Byte, gLen As Long, gIdx As Long, gCode As Long

    gBytes = mDocText
    gLen = UBound(gBytes) - 1
    For gIdx = 0 To gLen Step 2
        gCode = gBytes(gIdx) Or (CLng(gBytes(gIdx + 1)) * 256&)
        Select Case gCode
            Case 40: gPO = gPO + 1
            Case 41: gPC = gPC + 1
            Case 91: gSO = gSO + 1
            Case 93: gSC = gSC + 1
            Case 123: gCO = gCO + 1
            Case 125: gCC = gCC + 1
        End Select
    Next gIdx

    If gPO = gPC And gSO = gSC And gCO = gCC Then
        mSkipBracketIntegrity = True
    End If
End Sub

' ============================================================
'  PRIVATE: Spelling precheck
'  Checks if ANY of the ~133 spelling search terms exist in
'  the document text.  Uses representative root forms to
'  minimise lookups.  If none are found, the entire spelling
'  rule can be skipped.
' ============================================================
Private Sub PrecheckSpelling()
    If Not mHasDocText Then Exit Sub

    ' Build lower-case doc text once (reused for all InStr calls)
    If Len(mDocTextLC) = 0 Then mDocTextLC = LCase$(mDocText)

    ' Representative root forms that cover spelling-pair categories.
    ' If none of these exist, the full tokenisation can be skipped.
    ' We check roots rather than all 133 pairs for speed.
    Dim roots As Variant
    roots = Array( _
        "color", "colour", "favor", "favour", "honor", "honour", _
        "labor", "labour", "neighbor", "neighbour", "behavior", "behaviour", _
        "center", "centre", "fiber", "fibre", "theater", "theatre", _
        "defense", "defence", "offense", "offence", _
        "analog", "catalogue", "dialog", "dialogue", _
        "organize", "organise", "realize", "realise", "recognize", "recognise", _
        "authorize", "authorise", "utilize", "utilise", _
        "organization", "organisation", "authorization", "authorisation", _
        "gray", "grey", "skeptic", "sceptic", _
        "traveled", "travelled", "canceled", "cancelled", _
        "fulfill", "fulfil", "enrollment", "enrolment", _
        "acknowledgment", "acknowledgement", _
        "aging", "ageing", "artifact", "artefact")

    Dim i As Long
    For i = LBound(roots) To UBound(roots)
        If InStr(1, mDocTextLC, CStr(roots(i)), vbBinaryCompare) > 0 Then
            ' At least one term found -- full scan needed
            Exit Sub
        End If
    Next i

    ' No representative terms found -- skip spelling rule
    mSkipSpelling = True
End Sub
