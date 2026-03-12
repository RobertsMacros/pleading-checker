Attribute VB_Name = "PleadingsLauncher"
' ============================================================
' PleadingsLauncher.bas
' Lightweight launcher for the Pleadings Checker.
' Uses MsgBox/InputBox only -- no UserForm required.
'
' Dependencies:
'   - PleadingsEngine.bas
' ============================================================
Option Explicit

' ============================================================
'  MAIN LAUNCHER (called by PleadingsEngine.PleadingsChecker)
' ============================================================
Public Sub LaunchChecker()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Pleadings Checker"
        Exit Sub
    End If

    ' -- Choose action --
    Dim choice As Long
    choice = MsgBox("Pleadings Checker" & vbCrLf & vbCrLf & _
                    "Document: " & ActiveDocument.Name & vbCrLf & vbCrLf & _
                    "All imported rule modules will run." & vbCrLf & _
                    "Click Yes to run checks, No for options, Cancel to exit.", _
                    vbYesNoCancel + vbInformation, "Pleadings Checker")

    Select Case choice
        Case vbYes
            RunChecks
        Case vbNo
            ShowOptions
        Case vbCancel
            Exit Sub
    End Select
End Sub

' ============================================================
'  RUN CHECKS
' ============================================================
Private Sub RunChecks()
    Dim cfg As Object
    Set cfg = Application.Run("PleadingsEngine.InitRuleConfig")

    ' -- Page range prompt --
    Dim pgInput As String
    pgInput = InputBox("Page range (e.g. 1-10, or leave blank for all pages):", _
                        "Pleadings Checker - Page Range", "")
    If Len(Trim(pgInput)) > 0 Then
        ParsePageRange pgInput
    Else
        Application.Run "PleadingsEngine.SetPageRange", 0, 0
    End If

    ' -- Spelling mode prompt --
    Dim spMode As Long
    spMode = MsgBox("Enforce UK spelling?" & vbCrLf & vbCrLf & _
                    "Yes = UK spelling (default)" & vbCrLf & _
                    "No = US spelling", _
                    vbYesNo + vbQuestion, "Spelling Mode")
    If spMode = vbNo Then
        Application.Run "PleadingsEngine.SetSpellingMode", "US"
    Else
        Application.Run "PleadingsEngine.SetSpellingMode", "UK"
    End If

    ' -- Run --
    Application.StatusBar = "Pleadings Checker: running checks..."
    DoEvents

    Dim issues As Collection
    Set issues = Application.Run("PleadingsEngine.RunAllPleadingsRules", ActiveDocument, cfg)

    Application.StatusBar = ""

    ' -- Show results --
    Dim errCount As Long
    errCount = Application.Run("PleadingsEngine.GetRuleErrorCount")

    If issues.Count = 0 Then
        If errCount > 0 Then
            MsgBox "No issues found, but " & errCount & " rule(s) failed to run." & vbCrLf & vbCrLf & _
                   "Check Immediate window (Ctrl+G) for details.", _
                   vbExclamation, "Pleadings Checker"
        Else
            MsgBox "No issues found -- document looks clean.", _
                   vbInformation, "Pleadings Checker"
        End If
        Exit Sub
    End If

    Dim applyChoice As Long
    applyChoice = MsgBox(issues.Count & " issue(s) found." & vbCrLf & vbCrLf & _
                         "Apply to document?" & vbCrLf & _
                         "Yes = Apply as tracked changes" & vbCrLf & _
                         "No = Highlight + comments only" & vbCrLf & _
                         "Cancel = View results only", _
                         vbYesNoCancel + vbInformation, _
                         "Pleadings Checker -- " & _
                         issues.Count & " Issue(s)")

    Select Case applyChoice
        Case vbYes
            Application.Run "PleadingsEngine.ApplySuggestionsAsTrackedChanges", ActiveDocument, issues, True
            MsgBox issues.Count & " issue(s) applied as tracked changes.", _
                   vbInformation, "Pleadings Checker"
        Case vbNo
            Application.Run "PleadingsEngine.ApplyHighlights", ActiveDocument, issues, True
            MsgBox issues.Count & " issue(s) highlighted with comments.", _
                   vbInformation, "Pleadings Checker"
        Case vbCancel
            ' Just show summary, already displayed above
    End Select

    ' -- Offer to export report --
    Dim exportChoice As Long
    exportChoice = MsgBox("Export JSON report?", vbYesNo + vbQuestion, "Pleadings Checker")
    If exportChoice = vbYes Then
        ExportReport issues
    End If
End Sub

' ============================================================
'  OPTIONS MENU
' ============================================================
Private Sub ShowOptions()
    Dim optChoice As Long
    optChoice = MsgBox("Options:" & vbCrLf & vbCrLf & _
                       "Yes = Manage brand name rules" & vbCrLf & _
                       "No = Run checks (go back)", _
                       vbYesNo + vbQuestion, "Pleadings Checker - Options")

    If optChoice = vbYes Then
        ManageBrands
    Else
        RunChecks
    End If
End Sub

' ============================================================
'  BRAND NAME MANAGEMENT
' ============================================================
Private Sub ManageBrands()
    Dim action As String
    action = InputBox("Brand name management:" & vbCrLf & vbCrLf & _
                      "Type ADD to add a brand rule" & vbCrLf & _
                      "Type LOAD to load from file" & vbCrLf & _
                      "Type SAVE to save to file" & vbCrLf & _
                      "Or leave blank to go back.", _
                      "Pleadings Checker - Brands", "")

    Select Case UCase(Trim(action))
        Case "ADD"
            Dim correct As String
            correct = InputBox("Enter the correct brand form:", "Add Brand Rule", "")
            If Len(Trim(correct)) = 0 Then Exit Sub

            Dim incorrect As String
            incorrect = InputBox("Enter incorrect variants (comma-separated):", _
                                  "Add Brand Rule", "")
            If Len(Trim(incorrect)) = 0 Then Exit Sub

            On Error Resume Next
            Application.Run "Rules_Brands.AddBrandRule", correct, incorrect
            If Err.Number <> 0 Then
                MsgBox "Rules_Brands module not imported.", vbExclamation, "Pleadings Checker"
                Err.Clear
            Else
                MsgBox "Brand rule added: " & correct, vbInformation, "Pleadings Checker"
            End If
            On Error GoTo 0

        Case "LOAD"
            Dim loadPath As String
            #If Mac Then
                loadPath = Environ("HOME") & "/Library/Application Support/PleadingsChecker/brand_rules.txt"
            #Else
                loadPath = Environ("APPDATA") & "\PleadingsChecker\brand_rules.txt"
            #End If
            On Error Resume Next
            Application.Run "Rules_Brands.LoadBrandRules", loadPath
            If Err.Number <> 0 Then
                MsgBox "Rules_Brands module not imported or file not found.", _
                       vbExclamation, "Pleadings Checker"
                Err.Clear
            Else
                MsgBox "Brand rules loaded.", vbInformation, "Pleadings Checker"
            End If
            On Error GoTo 0

        Case "SAVE"
            Dim savePath As String
            #If Mac Then
                savePath = Environ("HOME") & "/Library/Application Support/PleadingsChecker/brand_rules.txt"
                On Error Resume Next
                MkDir Environ("HOME") & "/Library/Application Support/PleadingsChecker"
            #Else
                savePath = Environ("APPDATA") & "\PleadingsChecker\brand_rules.txt"
                On Error Resume Next
                MkDir Environ("APPDATA") & "\PleadingsChecker"
            #End If
            Err.Clear
            Application.Run "Rules_Brands.SaveBrandRules", savePath
            If Err.Number <> 0 Then
                MsgBox "Rules_Brands module not imported.", vbExclamation, "Pleadings Checker"
                Err.Clear
            Else
                MsgBox "Brand rules saved to:" & vbCrLf & savePath, _
                       vbInformation, "Pleadings Checker"
            End If
            On Error GoTo 0

        Case Else
            ' Go back
    End Select
End Sub

' ============================================================
'  EXPORT REPORT
' ============================================================
Private Sub ExportReport(issues As Collection)
    Dim reportPath As String
    Dim sep As String
    sep = Application.PathSeparator

    If ActiveDocument.Path <> "" Then
        ' Strip any extension from the document name
        Dim baseName As String
        baseName = ActiveDocument.Name
        Dim dotPos As Long
        dotPos = InStrRev(baseName, ".")
        If dotPos > 1 Then baseName = Left$(baseName, dotPos - 1)
        reportPath = ActiveDocument.Path & sep & baseName & "_pleadings_report.json"
    Else
        ' No saved path: use temp directory (cross-platform)
        #If Mac Then
            reportPath = Environ("TMPDIR")
            If Len(reportPath) = 0 Then reportPath = "/tmp"
            reportPath = reportPath & sep & "pleadings_report.json"
        #Else
            reportPath = Environ("TEMP") & sep & "pleadings_report.json"
        #End If
    End If

    Dim summary As String
    summary = Application.Run("PleadingsEngine.GenerateReport", issues, reportPath, ActiveDocument)

    MsgBox "Report saved to:" & vbCrLf & reportPath, _
           vbInformation, "Pleadings Checker"
End Sub

' ============================================================
'  PARSE PAGE RANGE INPUT
'  Supports: "5", "1-10", "1:10", "1" & ChrW(8211) & "10",
'  "1-3, 7-9", "1-3, 5, 8-12" (comma-separated segments).
'  Sets the overall min-max envelope.
' ============================================================
Private Sub ParsePageRange(ByVal pageInput As String)
    pageInput = Trim(pageInput)
    If Len(pageInput) = 0 Then
        Application.Run "PleadingsEngine.SetPageRange", 0, 0
        Exit Sub
    End If

    ' Normalise separators: en-dash and colon to hyphen
    pageInput = Replace(pageInput, ChrW(8211), "-")  ' en-dash
    pageInput = Replace(pageInput, ":", "-")

    ' Split on comma for multi-segment support
    Dim segments() As String
    segments = Split(pageInput, ",")

    Dim globalMin As Long, globalMax As Long
    globalMin = 2147483647  ' Long max
    globalMax = 0

    Dim s As Long
    For s = 0 To UBound(segments)
        Dim seg As String
        seg = Trim(segments(s))
        If Len(seg) = 0 Then GoTo NextSeg

        Dim dashPos As Long
        dashPos = InStr(1, seg, "-")
        If dashPos > 1 Then
            Dim lPart As String, rPart As String
            lPart = Trim(Left$(seg, dashPos - 1))
            rPart = Trim(Mid$(seg, dashPos + 1))
            If IsNumeric(lPart) And IsNumeric(rPart) Then
                Dim lo As Long, hi As Long
                lo = CLng(lPart)
                hi = CLng(rPart)
                If lo < globalMin Then globalMin = lo
                If hi > globalMax Then globalMax = hi
            End If
        ElseIf IsNumeric(seg) Then
            Dim pg As Long
            pg = CLng(seg)
            If pg < globalMin Then globalMin = pg
            If pg > globalMax Then globalMax = pg
        End If
NextSeg:
    Next s

    If globalMax > 0 And globalMin <= globalMax Then
        Application.Run "PleadingsEngine.SetPageRange", globalMin, globalMax
    Else
        ' Invalid input -- use all pages
        Application.Run "PleadingsEngine.SetPageRange", 0, 0
    End If
End Sub
