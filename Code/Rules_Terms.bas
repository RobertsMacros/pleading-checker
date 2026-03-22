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
    terms = TextAnchoring.MergeArrays2(batch1, batch2)

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

    ' Ensure the parent directory exists before writing.
    EnsureParentDir filePath

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

' Ensure the parent directory of a file path exists.
' Works for both Windows (backslash) and Mac (forward slash) paths.
' Uses MkDir iteratively for each missing ancestor.  No external
' dependencies (FileSystemObject, shell, etc.).
Private Sub EnsureParentDir(ByVal filePath As String)
    On Error Resume Next
    Dim sep As String
    sep = Application.PathSeparator
    ' Extract parent directory
    Dim lastSep As Long
    lastSep = InStrRev(filePath, sep)
    If lastSep <= 0 Then Exit Sub
    Dim parentDir As String
    parentDir = Left$(filePath, lastSep - 1)
    ' If directory already exists, nothing to do.
    If Dir(parentDir, vbDirectory) <> "" Then Exit Sub
    ' Walk the path from root and create each missing segment.
    Dim parts() As String
    parts = Split(parentDir, sep)
    Dim built As String
    built = parts(0) ' drive letter or first segment
    Dim i As Long
    For i = 1 To UBound(parts)
        built = built & sep & parts(i)
        If Dir(built, vbDirectory) = "" Then
            MkDir built
        End If
    Next i
    Err.Clear
    On Error GoTo 0
End Sub


