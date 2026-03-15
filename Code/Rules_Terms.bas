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

' -- Module-level user-managed whitelist dictionary --
Private userWhitelist As Object

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

    ' -- Build the dictionary from defaults + user terms ----
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

    ' Merge in user-managed terms
    If Not userWhitelist Is Nothing Then
        Dim uk As Variant
        For Each uk In userWhitelist.keys
            If Not dict.Exists(CStr(uk)) Then
                dict.Add CStr(uk), True
            End If
        Next uk
    End If

    ' -- Store in the engine for other rules to query -------
    TextAnchoring.SetWhitelist dict

    On Error GoTo 0

    ' This rule returns no issues -- it is purely a setup rule
    Set Check_CustomTermWhitelist = issues
End Function

' ============================================================
'  WHITELIST MANAGEMENT (UI parity with brand rules)
' ============================================================

' Return the user-managed whitelist dictionary for UI display
Public Function GetWhitelistTerms() As Object
    EnsureUserWhitelist
    Set GetWhitelistTerms = userWhitelist
End Function

' Add a term to the user-managed whitelist
Public Sub AddWhitelistTerm(ByVal term As String)
    EnsureUserWhitelist
    Dim lc As String
    lc = LCase(Trim(term))
    If Len(lc) > 0 And Not userWhitelist.Exists(lc) Then
        userWhitelist.Add lc, True
    End If
End Sub

' Remove a term from the user-managed whitelist
Public Sub RemoveWhitelistTerm(ByVal term As String)
    EnsureUserWhitelist
    Dim lc As String
    lc = LCase(Trim(term))
    If userWhitelist.Exists(lc) Then
        userWhitelist.Remove lc
    End If
End Sub

' Save user whitelist to a text file (one term per line)
Public Function SaveWhitelistTerms(ByVal filePath As String) As Boolean
    SaveWhitelistTerms = False
    EnsureUserWhitelist

    On Error GoTo SaveFail
    Dim fNum As Long
    fNum = FreeFile
    Open filePath For Output As #fNum
    Dim k As Variant
    For Each k In userWhitelist.keys
        Print #fNum, CStr(k)
    Next k
    Close #fNum
    SaveWhitelistTerms = True
    Exit Function

SaveFail:
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
End Function

' Load user whitelist from a text file (one term per line)
Public Function LoadWhitelistTerms(ByVal filePath As String) As Boolean
    LoadWhitelistTerms = False

    If Dir(filePath) = "" Then Exit Function

    On Error GoTo LoadFail
    Dim fNum As Long
    fNum = FreeFile
    Open filePath For Input As #fNum

    Set userWhitelist = CreateObject("Scripting.Dictionary")
    Dim line As String
    Do While Not EOF(fNum)
        Line Input #fNum, line
        line = LCase(Trim(line))
        If Len(line) > 0 And Not userWhitelist.Exists(line) Then
            userWhitelist.Add line, True
        End If
    Loop
    Close #fNum
    LoadWhitelistTerms = True
    Exit Function

LoadFail:
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
End Function

' Default persistence path
Public Function GetDefaultWhitelistPath() As String
    Dim sep As String
    sep = Application.PathSeparator
    #If Mac Then
        GetDefaultWhitelistPath = Environ("HOME") & sep & "Library" & sep & _
                                   "Application Support" & sep & "PleadingsChecker" & sep & "whitelist.txt"
    #Else
        GetDefaultWhitelistPath = Environ("APPDATA") & sep & "PleadingsChecker" & sep & "whitelist.txt"
    #End If
End Function

' Ensure the user whitelist dictionary exists
Private Sub EnsureUserWhitelist()
    If userWhitelist Is Nothing Then
        Set userWhitelist = CreateObject("Scripting.Dictionary")
    End If
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
