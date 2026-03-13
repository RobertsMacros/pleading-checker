Attribute VB_Name = "modDebugLog"
' ============================================================
' modDebugLog.bas
' Lightweight, removable debugging layer for the Pleadings
' Checker Word VBA project.
'
' USAGE:
'   - Set DEBUG_MODE = True to enable logging
'   - Set DEBUG_MODE = False to no-op all calls (near zero overhead)
'   - All trace/log output goes to the Immediate Window (Ctrl+G)
'     AND a rolling in-memory buffer
'   - Call DebugLogFlushToImmediate to replay the buffer
'   - Call DebugLogGetText() to copy the full log as a string
'   - To remove: delete this module, then remove the small
'     TraceEnter/TraceStep/etc. calls from PleadingsEngine.bas
'
' Dependencies: None (Word VBA only, late-bound, Option Explicit)
' ============================================================
Option Explicit

' ============================================================
'  A. GLOBAL TOGGLE
' ============================================================
Public Const DEBUG_MODE As Boolean = True

' ============================================================
'  B. ROLLING IN-MEMORY LOG BUFFER
' ============================================================
Private Const LOG_CAP As Long = 2000          ' max entries kept
Private logBuf()      As String               ' circular buffer
Private logHead       As Long                 ' next write slot
Private logCount      As Long                 ' entries written
Private logSeq        As Long                 ' sequence counter
Private logInited     As Boolean              ' lazy init flag
Private callDepth     As Long                 ' indent depth

Private Sub EnsureLogInit()
    If logInited Then Exit Sub
    ReDim logBuf(0 To LOG_CAP - 1)
    logHead = 0
    logCount = 0
    logSeq = 0
    callDepth = 0
    logInited = True
End Sub

' ============================================================
'  CORE: Write one line to buffer + Immediate Window
' ============================================================
Private Sub LogLine(ByVal msg As String)
    EnsureLogInit
    logSeq = logSeq + 1
    Dim ts As String
    ts = Format(Timer, "00000.00")
    Dim prefix As String
    prefix = "[" & Format(logSeq, "00000") & " T" & ts & "] "
    ' Indent by call depth
    If callDepth > 0 Then
        prefix = prefix & String(callDepth * 2, " ")
    End If
    Dim fullLine As String
    fullLine = prefix & msg
    ' Write to Immediate Window
    Debug.Print fullLine
    ' Write to circular buffer
    logBuf(logHead) = fullLine
    logHead = (logHead + 1) Mod LOG_CAP
    If logCount < LOG_CAP Then logCount = logCount + 1
End Sub

' ============================================================
'  C. TRACE HELPERS
' ============================================================
Public Sub TraceEnter(ByVal procName As String)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    LogLine ">> ENTER " & procName
    callDepth = callDepth + 1
    On Error GoTo 0
End Sub

Public Sub TraceStep(ByVal procName As String, ByVal stepName As String)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    LogLine "-- " & procName & ": " & stepName
    On Error GoTo 0
End Sub

Public Sub TraceExit(ByVal procName As String, Optional ByVal summary As String = "")
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    If callDepth > 0 Then callDepth = callDepth - 1
    If Len(summary) > 0 Then
        LogLine "<< EXIT  " & procName & " (" & summary & ")"
    Else
        LogLine "<< EXIT  " & procName
    End If
    On Error GoTo 0
End Sub

Public Sub TraceFail(ByVal procName As String, ByVal reason As String)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    LogLine "!! FAIL  " & procName & ": " & reason
    On Error GoTo 0
End Sub

' ============================================================
'  D. GENERAL LOGGING HELPERS
' ============================================================
Public Sub DebugLog(ByVal message As String)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    LogLine message
    On Error GoTo 0
End Sub

Public Sub DebugLogKV(ByVal keyName As String, ByVal keyValue As String)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    LogLine keyName & " = " & keyValue
    On Error GoTo 0
End Sub

Public Sub DebugLogError(ByVal procName As String, _
                         ByVal stepName As String, _
                         ByVal errNumber As Long, _
                         ByVal errDescription As String)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    LogLine "!! ERROR " & procName & " @ " & stepName & _
            " -- Err " & errNumber & ": " & errDescription
    On Error GoTo 0
End Sub

' ============================================================
'  D2. WORD OBJECT DIAGNOSTICS
' ============================================================

' --- Range diagnostics ---
Public Sub DebugLogRange(ByVal labelText As String, ByVal rng As Range)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    If rng Is Nothing Then
        LogLine "RANGE [" & labelText & "]: Nothing"
        On Error GoTo 0
        Exit Sub
    End If

    Dim info As String
    info = "RANGE [" & labelText & "]:"

    ' Start / End / Length
    Dim rStart As Long, rEnd As Long
    rStart = rng.Start: If Err.Number <> 0 Then rStart = -1: Err.Clear
    rEnd = rng.End:     If Err.Number <> 0 Then rEnd = -1: Err.Clear
    info = info & " start=" & rStart & " end=" & rEnd & " len=" & (rEnd - rStart)

    ' Collapsed?
    If rStart = rEnd Then info = info & " COLLAPSED"

    ' Story type
    Dim storyType As Long
    storyType = rng.StoryType: If Err.Number <> 0 Then storyType = -1: Err.Clear
    info = info & " story=" & storyType
    If storyType = 1 Then
        info = info & "(Main)"
    ElseIf storyType = 2 Then
        info = info & "(Footnotes)"
    ElseIf storyType = 3 Then
        info = info & "(Endnotes)"
    ElseIf storyType = 6 Then
        info = info & "(TextFrame)"
    End If

    ' Text preview (first 60 chars)
    Dim preview As String
    preview = ""
    preview = rng.Text: If Err.Number <> 0 Then preview = "<err>": Err.Clear
    If Len(preview) > 60 Then preview = Left$(preview, 60) & "..."
    preview = Replace(Replace(Replace(preview, vbCr, "\r"), vbLf, "\n"), vbTab, "\t")
    info = info & " text=""" & preview & """"

    ' In table?
    Dim inTable As Boolean
    inTable = False
    If Not rng.Tables Is Nothing Then
        If rng.Tables.Count > 0 Then inTable = True
    End If
    If Err.Number <> 0 Then Err.Clear
    If inTable Then info = info & " IN_TABLE"

    ' Fields
    Dim fieldCnt As Long
    fieldCnt = 0
    fieldCnt = rng.Fields.Count: If Err.Number <> 0 Then fieldCnt = -1: Err.Clear
    If fieldCnt > 0 Then info = info & " fields=" & fieldCnt

    ' Content controls
    Dim ccCnt As Long
    ccCnt = 0
    ccCnt = rng.ContentControls.Count: If Err.Number <> 0 Then ccCnt = -1: Err.Clear
    If ccCnt > 0 Then info = info & " contentControls=" & ccCnt

    ' Document protection
    Dim protType As Long
    protType = -1
    protType = rng.Document.ProtectionType: If Err.Number <> 0 Then protType = -1: Err.Clear
    If protType <> -1 Then info = info & " docProtection=" & protType

    LogLine info
    On Error GoTo 0
End Sub

' --- Document diagnostics ---
Public Sub DebugLogDoc(ByVal labelText As String, ByVal doc As Document)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    If doc Is Nothing Then
        LogLine "DOC [" & labelText & "]: Nothing"
        On Error GoTo 0
        Exit Sub
    End If

    Dim info As String
    info = "DOC [" & labelText & "]:"

    ' Name
    Dim docName As String
    docName = doc.Name: If Err.Number <> 0 Then docName = "<err>": Err.Clear
    info = info & " name=""" & docName & """"

    ' Path
    Dim docPath As String
    docPath = doc.Path: If Err.Number <> 0 Then docPath = "<err>": Err.Clear
    If Len(docPath) > 0 Then info = info & " path=""" & docPath & """"

    ' Protection
    Dim protType As Long
    protType = -1
    protType = doc.ProtectionType: If Err.Number <> 0 Then protType = -1: Err.Clear
    info = info & " protection=" & protType
    If protType = -1 Then
        info = info & "(None)"
    ElseIf protType = 0 Then
        info = info & "(AllowOnlyRevisions)"
    ElseIf protType = 1 Then
        info = info & "(AllowOnlyComments)"
    ElseIf protType = 2 Then
        info = info & "(AllowOnlyFormFields)"
    ElseIf protType = 3 Then
        info = info & "(NoProtection)"
    End If

    ' Track revisions
    Dim trackRev As Boolean
    trackRev = doc.TrackRevisions: If Err.Number <> 0 Then Err.Clear
    info = info & " trackRevisions=" & trackRev

    ' Show revisions
    Dim showRev As Long
    showRev = -1
    showRev = doc.ActiveWindow.View.RevisionsFilter.Markup
    If Err.Number <> 0 Then Err.Clear

    ' Comments count
    Dim cmtCnt As Long
    cmtCnt = 0
    cmtCnt = doc.Comments.Count: If Err.Number <> 0 Then cmtCnt = -1: Err.Clear
    info = info & " comments=" & cmtCnt

    ' Revisions count
    Dim revCnt As Long
    revCnt = 0
    revCnt = doc.Revisions.Count: If Err.Number <> 0 Then revCnt = -1: Err.Clear
    info = info & " revisions=" & revCnt

    LogLine info
    On Error GoTo 0
End Sub

' --- Revision diagnostics ---
Public Sub DebugLogRevision(ByVal labelText As String, ByVal rev As Revision)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    If rev Is Nothing Then
        LogLine "REVISION [" & labelText & "]: Nothing"
        On Error GoTo 0
        Exit Sub
    End If

    Dim info As String
    info = "REVISION [" & labelText & "]:"

    ' Type
    Dim revType As Long
    revType = rev.Type: If Err.Number <> 0 Then revType = -1: Err.Clear
    info = info & " type=" & revType
    If revType = 1 Then
        info = info & "(Insert)"
    ElseIf revType = 2 Then
        info = info & "(Delete)"
    ElseIf revType = 6 Then
        info = info & "(PropertyChange)"
    End If

    ' Range preview
    Dim rStart As Long, rEnd As Long
    rStart = rev.Range.Start: If Err.Number <> 0 Then rStart = -1: Err.Clear
    rEnd = rev.Range.End:     If Err.Number <> 0 Then rEnd = -1: Err.Clear
    info = info & " start=" & rStart & " end=" & rEnd

    Dim preview As String
    preview = ""
    preview = rev.Range.Text: If Err.Number <> 0 Then preview = "<err>": Err.Clear
    If Len(preview) > 40 Then preview = Left$(preview, 40) & "..."
    preview = Replace(Replace(Replace(preview, vbCr, "\r"), vbLf, "\n"), vbTab, "\t")
    info = info & " text=""" & preview & """"

    ' Author
    Dim revAuthor As String
    revAuthor = ""
    revAuthor = rev.Author: If Err.Number <> 0 Then revAuthor = "<err>": Err.Clear
    If Len(revAuthor) > 0 Then info = info & " author=""" & revAuthor & """"

    LogLine info
    On Error GoTo 0
End Sub

' --- Comment diagnostics ---
Public Sub DebugLogComment(ByVal labelText As String, ByVal cmt As Comment)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    If cmt Is Nothing Then
        LogLine "COMMENT [" & labelText & "]: Nothing"
        On Error GoTo 0
        Exit Sub
    End If

    Dim info As String
    info = "COMMENT [" & labelText & "]:"

    ' Author / initials
    Dim cmtAuthor As String
    cmtAuthor = cmt.Author: If Err.Number <> 0 Then cmtAuthor = "<err>": Err.Clear
    info = info & " author=""" & cmtAuthor & """"

    Dim cmtInitials As String
    cmtInitials = cmt.Initial: If Err.Number <> 0 Then cmtInitials = "<err>": Err.Clear
    info = info & " initials=""" & cmtInitials & """"

    ' Comment text preview
    Dim cmtText As String
    cmtText = ""
    cmtText = cmt.Range.Text: If Err.Number <> 0 Then cmtText = "<err>": Err.Clear
    If Len(cmtText) > 60 Then cmtText = Left$(cmtText, 60) & "..."
    cmtText = Replace(Replace(cmtText, vbCr, "\r"), vbLf, "\n")
    info = info & " text=""" & cmtText & """"

    ' Scope (anchor) preview
    Dim scopeText As String
    scopeText = ""
    scopeText = cmt.Scope.Text: If Err.Number <> 0 Then scopeText = "<err>": Err.Clear
    If Len(scopeText) > 40 Then scopeText = Left$(scopeText, 40) & "..."
    scopeText = Replace(Replace(scopeText, vbCr, "\r"), vbLf, "\n")
    info = info & " scope=""" & scopeText & """"

    ' Scope range
    Dim scStart As Long, scEnd As Long
    scStart = cmt.Scope.Start: If Err.Number <> 0 Then scStart = -1: Err.Clear
    scEnd = cmt.Scope.End:     If Err.Number <> 0 Then scEnd = -1: Err.Clear
    info = info & " scopeStart=" & scStart & " scopeEnd=" & scEnd

    LogLine info
    On Error GoTo 0
End Sub

' ============================================================
'  E. FLUSH / OUTPUT HELPERS
' ============================================================
Public Sub DebugLogClear()
    EnsureLogInit
    ReDim logBuf(0 To LOG_CAP - 1)
    logHead = 0
    logCount = 0
    logSeq = 0
    callDepth = 0
End Sub

Public Sub DebugLogFlushToImmediate()
    If Not logInited Then Exit Sub
    Dim idx As Long, startIdx As Long
    If logCount < LOG_CAP Then
        startIdx = 0
    Else
        startIdx = logHead  ' oldest entry
    End If
    Debug.Print "=== DEBUG LOG REPLAY (" & logCount & " entries) ==="
    Dim lineIdx As Long
    For lineIdx = 0 To logCount - 1
        idx = (startIdx + lineIdx) Mod LOG_CAP
        Debug.Print logBuf(idx)
    Next lineIdx
    Debug.Print "=== END DEBUG LOG ==="
End Sub

Public Function DebugLogGetText() As String
    If Not logInited Then
        DebugLogGetText = ""
        Exit Function
    End If
    Dim result As String
    Dim idx As Long, startIdx As Long
    If logCount < LOG_CAP Then
        startIdx = 0
    Else
        startIdx = logHead
    End If
    result = "=== DEBUG LOG (" & logCount & " entries) ===" & vbCrLf
    Dim lineIdx As Long
    For lineIdx = 0 To logCount - 1
        idx = (startIdx + lineIdx) Mod LOG_CAP
        result = result & logBuf(idx) & vbCrLf
    Next lineIdx
    result = result & "=== END DEBUG LOG ==="
    DebugLogGetText = result
End Function

Public Function DebugLogSaveToTextFile(ByVal filePath As String) As Boolean
    DebugLogSaveToTextFile = False
    If Not logInited Then Exit Function
    If logCount = 0 Then Exit Function

    Dim fileNum As Integer
    fileNum = FreeFile
    On Error Resume Next
    Open filePath For Output As #fileNum
    If Err.Number <> 0 Then
        Debug.Print "DebugLogSaveToTextFile: cannot open " & filePath & _
                    " (Err " & Err.Number & ": " & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    On Error GoTo SaveErr
    Dim idx As Long, startIdx As Long
    If logCount < LOG_CAP Then
        startIdx = 0
    Else
        startIdx = logHead
    End If
    Print #fileNum, "=== DEBUG LOG (" & logCount & " entries) ==="
    Dim lineIdx As Long
    For lineIdx = 0 To logCount - 1
        idx = (startIdx + lineIdx) Mod LOG_CAP
        Print #fileNum, logBuf(idx)
    Next lineIdx
    Print #fileNum, "=== END DEBUG LOG ==="
    Close #fileNum
    DebugLogSaveToTextFile = True
    Exit Function

SaveErr:
    On Error Resume Next
    Close #fileNum
    Debug.Print "DebugLogSaveToTextFile: write error " & Err.Number & ": " & Err.Description
    On Error GoTo 0
End Function

Public Sub DebugLogFlushToDocument(Optional ByVal doc As Document = Nothing)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    Dim targetDoc As Document
    If doc Is Nothing Then
        Set targetDoc = Documents.Add
    Else
        Set targetDoc = doc
    End If
    If Err.Number <> 0 Then
        Debug.Print "DebugLogFlushToDocument: cannot create/use document"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    targetDoc.Content.Text = DebugLogGetText()
    On Error GoTo 0
End Sub

' ============================================================
'  F. WRAPPER HELPERS FOR RISKY OPERATIONS
' ============================================================

' --- Try to set range text (tracked or untracked) ---
Public Function TrySetRangeText(ByVal rng As Range, _
                                ByVal newText As String, _
                                ByVal procName As String, _
                                ByVal stepName As String) As Boolean
    TrySetRangeText = False
    If Not DEBUG_MODE Then
        On Error Resume Next
        rng.Text = newText
        TrySetRangeText = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error Resume Next
    DebugLogRange procName & "." & stepName & " BEFORE", rng
    DebugLog procName & "." & stepName & ": setting text to """ & _
             Left$(Replace(Replace(newText, vbCr, "\r"), vbLf, "\n"), 60) & """"

    Err.Clear
    rng.Text = newText

    If Err.Number <> 0 Then
        DebugLogError procName, stepName & " rng.Text=", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    DebugLogRange procName & "." & stepName & " AFTER", rng
    TrySetRangeText = True
    On Error GoTo 0
End Function

' --- Try to set formatted text (copy from source range) ---
Public Function TrySetFormattedText(ByVal rng As Range, _
                                    ByVal srcRange As Range, _
                                    ByVal procName As String, _
                                    ByVal stepName As String) As Boolean
    TrySetFormattedText = False
    If Not DEBUG_MODE Then
        On Error Resume Next
        srcRange.Copy
        rng.Paste
        TrySetFormattedText = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error Resume Next
    DebugLogRange procName & "." & stepName & " target BEFORE", rng
    DebugLogRange procName & "." & stepName & " source", srcRange
    DebugLog procName & "." & stepName & ": copying formatted text"

    Err.Clear
    srcRange.Copy
    If Err.Number <> 0 Then
        DebugLogError procName, stepName & " srcRange.Copy", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    rng.Paste
    If Err.Number <> 0 Then
        DebugLogError procName, stepName & " rng.Paste", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    DebugLogRange procName & "." & stepName & " target AFTER", rng
    TrySetFormattedText = True
    On Error GoTo 0
End Function

' --- Try to delete a range ---
Public Function TryDeleteRange(ByVal rng As Range, _
                               ByVal procName As String, _
                               ByVal stepName As String) As Boolean
    TryDeleteRange = False
    If Not DEBUG_MODE Then
        On Error Resume Next
        rng.Delete
        TryDeleteRange = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error Resume Next
    DebugLogRange procName & "." & stepName & " DELETE target", rng

    Err.Clear
    rng.Delete
    If Err.Number <> 0 Then
        DebugLogError procName, stepName & " rng.Delete", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    DebugLog procName & "." & stepName & ": delete OK"
    TryDeleteRange = True
    On Error GoTo 0
End Function

' --- Try to add a comment ---
Public Function TryAddComment(ByVal doc As Document, _
                              ByVal anchorRange As Range, _
                              ByVal commentText As String, _
                              ByRef newComment As Comment, _
                              ByVal procName As String, _
                              ByVal stepName As String) As Boolean
    TryAddComment = False
    Set newComment = Nothing
    If Not DEBUG_MODE Then
        On Error Resume Next
        Set newComment = doc.Comments.Add(Range:=anchorRange, Text:=commentText)
        TryAddComment = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error Resume Next
    DebugLogRange procName & "." & stepName & " comment anchor", anchorRange
    DebugLog procName & "." & stepName & ": adding comment, text=""" & _
             Left$(commentText, 80) & """"

    Err.Clear
    Set newComment = doc.Comments.Add(Range:=anchorRange, Text:=commentText)
    If Err.Number <> 0 Then
        DebugLogError procName, stepName & " Comments.Add", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    DebugLogComment procName & "." & stepName & " added", newComment
    TryAddComment = True
    On Error GoTo 0
End Function

' --- Try to unprotect a document ---
Public Function TryUnprotectDocument(ByVal doc As Document, _
                                     ByVal procName As String, _
                                     ByVal stepName As String) As Boolean
    TryUnprotectDocument = False
    If Not DEBUG_MODE Then
        On Error Resume Next
        If doc.ProtectionType <> -1 Then doc.Unprotect
        TryUnprotectDocument = (doc.ProtectionType = -1)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error Resume Next
    Dim protBefore As Long
    protBefore = doc.ProtectionType
    If Err.Number <> 0 Then protBefore = -99: Err.Clear

    DebugLog procName & "." & stepName & ": unprotecting doc, protBefore=" & protBefore

    If protBefore = -1 Then
        DebugLog procName & "." & stepName & ": already unprotected"
        TryUnprotectDocument = True
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    doc.Unprotect
    If Err.Number <> 0 Then
        DebugLogError procName, stepName & " doc.Unprotect", Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Dim protAfter As Long
    protAfter = doc.ProtectionType
    If Err.Number <> 0 Then protAfter = -99: Err.Clear
    DebugLog procName & "." & stepName & ": protAfter=" & protAfter

    If protAfter <> -1 Then
        TraceFail procName, stepName & ": unprotect did not take effect, protAfter=" & protAfter
    Else
        TryUnprotectDocument = True
    End If
    On Error GoTo 0
End Function

' --- Try to protect a document ---
Public Function TryProtectDocument(ByVal doc As Document, _
                                   ByVal protType As Long, _
                                   ByVal procName As String, _
                                   ByVal stepName As String) As Boolean
    TryProtectDocument = False
    On Error Resume Next
    If DEBUG_MODE Then
        DebugLog procName & "." & stepName & ": protecting doc, targetType=" & protType
    End If

    Err.Clear
    doc.Protect Type:=protType
    If Err.Number <> 0 Then
        If DEBUG_MODE Then
            DebugLogError procName, stepName & " doc.Protect", Err.Number, Err.Description
        End If
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    TryProtectDocument = True
    On Error GoTo 0
End Function

' ============================================================
'  G. FILE-SYSTEM HELPERS (no FSO dependency)
' ============================================================

' --- Recursively ensure a folder path exists ---
' Returns True if the folder exists (or was created), False on failure.
Public Function EnsureDirectoryExists(ByVal folderPath As String) As Boolean
    EnsureDirectoryExists = False
    If Len(folderPath) = 0 Then Exit Function

    ' Strip trailing separator
    Dim sep As String
    sep = Application.PathSeparator
    If Right$(folderPath, 1) = sep Then
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If
    If Len(folderPath) = 0 Then Exit Function

    ' Already exists?
    On Error Resume Next
    Dim testDir As String
    testDir = Dir(folderPath, vbDirectory)
    If Err.Number <> 0 Then testDir = "": Err.Clear
    On Error GoTo 0
    If Len(testDir) > 0 Then
        EnsureDirectoryExists = True
        Exit Function
    End If

    ' Walk path components, creating as needed
    Dim parts() As String
    parts = Split(folderPath, sep)
    If UBound(parts) < 0 Then Exit Function

    Dim built As String
    Dim i As Long

    #If Mac Then
        ' Unix paths start with /  so parts(0) = ""
        If Left$(folderPath, 1) = sep Then
            built = sep & parts(1)
            i = 2
        Else
            built = parts(0)
            i = 1
        End If
    #Else
        built = parts(0)   ' drive letter e.g. "C:"
        i = 1
    #End If

    For i = i To UBound(parts)
        built = built & sep & parts(i)
        On Error Resume Next
        testDir = ""
        testDir = Dir(built, vbDirectory)
        If Err.Number <> 0 Then testDir = "": Err.Clear
        If Len(testDir) = 0 Then
            MkDir built
            If Err.Number <> 0 Then
                If DEBUG_MODE Then
                    Debug.Print "EnsureDirectoryExists: MkDir failed for """ & built & _
                                """ (Err " & Err.Number & ": " & Err.Description & ")"
                End If
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next i

    EnsureDirectoryExists = True
End Function

' --- Extract parent directory from a file path ---
Public Function GetParentDirectory(ByVal filePath As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    Dim lastSep As Long
    lastSep = InStrRev(filePath, sep)
    If lastSep > 0 Then
        GetParentDirectory = Left$(filePath, lastSep - 1)
    Else
        GetParentDirectory = ""
    End If
End Function
